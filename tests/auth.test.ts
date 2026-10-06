import {
    buildScreenContext,
    exchangeToken,
    interceptOAuthCards,
    MAX_CONTEXT_JSON_LENGTH,
    DirectLineActivity,
    DirectLineConnection
} from "../AgentChatWidget/auth";

/** Minimal stand-in for a Direct Line activity stream. */
function fakeConnection(): { connection: DirectLineConnection; emit: (a: DirectLineActivity) => void; unsubscribed: () => boolean } {
    let listener: ((a: DirectLineActivity) => void) | undefined;
    let closed = false;
    return {
        connection: {
            activity$: {
                subscribe(observer) {
                    listener = observer.next;
                    return { unsubscribe: () => { closed = true; listener = undefined; } };
                }
            }
        },
        emit: (a) => listener?.(a),
        unsubscribed: () => closed
    };
}

function oauthActivity(id: string, connectionName = "AgentSSO"): DirectLineActivity {
    return {
        id,
        type: "message",
        attachments: [{ contentType: "application/vnd.microsoft.card.oauth", content: { connectionName } }]
    };
}

function fakeMsal(opts: { silentFails?: boolean } = {}) {
    return {
        initialize: jest.fn().mockResolvedValue(undefined),
        getAllAccounts: jest.fn().mockReturnValue([{ username: "user@example.com" }]),
        acquireTokenSilent: opts.silentFails
            ? jest.fn().mockRejectedValue(new Error("interaction_required"))
            : jest.fn().mockResolvedValue({ accessToken: "silent-token" }),
        acquireTokenPopup: jest.fn().mockResolvedValue({ accessToken: "popup-token" })
    };
}

const flush = () => new Promise((r) => setTimeout(r, 0));

describe("buildScreenContext - ContextJson is untrusted", () => {
    let warn: jest.SpyInstance;
    beforeEach(() => { warn = jest.spyOn(console, "warn").mockImplementation(() => undefined); });
    afterEach(() => warn.mockRestore());

    it("cannot overwrite the user or record the control was given", () => {
        const result = buildScreenContext({
            userId: "real@example.com",
            recordId: "123",
            contextJson: JSON.stringify({ userId: "attacker@example.com", recordId: "999", site: "north" })
        });

        expect(result.userId).toBe("real@example.com");
        expect(result.recordId).toBe("123");
        expect(result.site).toBe("north");
        expect(warn).toHaveBeenCalled();
    });

    it("cannot set a reserved field that the control left empty", () => {
        const result = buildScreenContext({ contextJson: JSON.stringify({ userId: "attacker@example.com" }) });
        expect(result.userId).toBeUndefined();
    });

    it("ignores arrays and primitives", () => {
        expect(buildScreenContext({ contextJson: "[1,2,3]" })).toEqual({});
        expect(buildScreenContext({ contextJson: "42" })).toEqual({});
        expect(buildScreenContext({ contextJson: "null" })).toEqual({});
    });

    it("drops prototype-polluting keys", () => {
        const result = buildScreenContext({ contextJson: '{"__proto__":{"polluted":true},"ok":1}' });
        expect(result.ok).toBe(1);
        expect(({} as Record<string, unknown>).polluted).toBeUndefined();
    });

    it("ignores oversized payloads", () => {
        const big = JSON.stringify({ note: "x".repeat(MAX_CONTEXT_JSON_LENGTH) });
        expect(buildScreenContext({ userId: "u", contextJson: big })).toEqual({ userId: "u" });
    });
});

describe("exchangeToken", () => {
    const base = { activityId: "a1", connectionName: "AgentSSO", conversationId: "c1", directLineToken: "dl", scope: "api://botid-x/.default" };
    let warn: jest.SpyInstance;
    beforeEach(() => { warn = jest.spyOn(console, "warn").mockImplementation(() => undefined); });
    afterEach(() => { warn.mockRestore(); (global as { fetch?: unknown }).fetch = undefined; });

    it("returns true when Direct Line accepts the exchange", async () => {
        const fetchMock = jest.fn().mockResolvedValue({ ok: true, status: 200 });
        (global as { fetch?: unknown }).fetch = fetchMock;

        await expect(exchangeToken({ ...base, msalInstance: fakeMsal() })).resolves.toBe(true);

        const body = JSON.parse(fetchMock.mock.calls[0][1].body);
        expect(body.name).toBe("signin/tokenExchange");
        expect(body.value).toEqual({ id: "a1", connectionName: "AgentSSO", token: "silent-token" });
    });

    it("returns false on an HTTP error rather than treating it as success", async () => {
        (global as { fetch?: unknown }).fetch = jest.fn().mockResolvedValue({ ok: false, status: 401 });

        await expect(exchangeToken({ ...base, msalInstance: fakeMsal() })).resolves.toBe(false);
        expect(warn).toHaveBeenCalled();
    });

    it("falls back to a popup only when silent acquisition fails", async () => {
        (global as { fetch?: unknown }).fetch = jest.fn().mockResolvedValue({ ok: true, status: 200 });
        const msal = fakeMsal({ silentFails: true });

        await expect(exchangeToken({ ...base, msalInstance: msal })).resolves.toBe(true);
        expect(msal.acquireTokenPopup).toHaveBeenCalledTimes(1);
    });

    it("returns false when the network call throws", async () => {
        (global as { fetch?: unknown }).fetch = jest.fn().mockRejectedValue(new Error("offline"));
        await expect(exchangeToken({ ...base, msalInstance: fakeMsal() })).resolves.toBe(false);
    });
});

describe("interceptOAuthCards", () => {
    let warn: jest.SpyInstance;
    beforeEach(() => {
        warn = jest.spyOn(console, "warn").mockImplementation(() => undefined);
        (global as { fetch?: unknown }).fetch = jest.fn().mockResolvedValue({ ok: true, status: 200 });
    });
    afterEach(() => { warn.mockRestore(); (global as { fetch?: unknown }).fetch = undefined; });

    it("exchanges a replayed card only once", async () => {
        const fake = fakeConnection();
        const msal = fakeMsal();
        interceptOAuthCards({ connection: fake.connection, conversationId: "c", directLineToken: "dl", msalInstance: msal, agentId: "x" });

        fake.emit(oauthActivity("a1"));
        fake.emit(oauthActivity("a1"));
        await flush();

        expect(msal.acquireTokenSilent).toHaveBeenCalledTimes(1);
    });

    it("ignores connections outside the allow-list", async () => {
        const fake = fakeConnection();
        const msal = fakeMsal();
        interceptOAuthCards({
            connection: fake.connection, conversationId: "c", directLineToken: "dl", msalInstance: msal, agentId: "x",
            allowedConnectionNames: ["AgentSSO"]
        });

        fake.emit(oauthActivity("a1", "SomethingElse"));
        await flush();

        expect(msal.acquireTokenSilent).not.toHaveBeenCalled();
    });

    it("ignores messages that are not OAuth cards", async () => {
        const fake = fakeConnection();
        const msal = fakeMsal();
        interceptOAuthCards({ connection: fake.connection, conversationId: "c", directLineToken: "dl", msalInstance: msal, agentId: "x" });

        fake.emit({ id: "m1", type: "message", attachments: [{ contentType: "application/vnd.microsoft.card.adaptive" }] });
        fake.emit({ id: "t1", type: "typing" });
        await flush();

        expect(msal.acquireTokenSilent).not.toHaveBeenCalled();
    });

    it("stops listening once cleaned up", () => {
        const fake = fakeConnection();
        const cleanup = interceptOAuthCards({ connection: fake.connection, conversationId: "c", directLineToken: "dl", msalInstance: fakeMsal(), agentId: "x" });
        cleanup();
        expect(fake.unsubscribed()).toBe(true);
    });
});
