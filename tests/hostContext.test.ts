import { resolveHostContext, buildScreenContext } from "../AgentChatWidget/auth";

const FORM_ID = "{3F2504E0-4F89-11D3-9A0C-0305E82C3301}";
const formHost = {
    mode: { contextInfo: { entityId: FORM_ID, entityTypeName: "incident" } },
    userSettings: { userId: "{11111111-2222-3333-4444-555555555555}", userName: "Alex Example" }
};

describe("resolveHostContext", () => {
    it("uses the form record and signed-in user when properties are empty", () => {
        const r = resolveHostContext({}, formHost);
        expect(r.recordId).toBe("3f2504e0-4f89-11d3-9a0c-0305e82c3301");
        expect(r.recordTable).toBe("incident");
        expect(r.recordSource).toBe("form");
        expect(r.userName).toBe("Alex Example");
        expect(r.userId).toBe("{11111111-2222-3333-4444-555555555555}");
    });

    it("lets explicit properties win over the host", () => {
        const r = resolveHostContext(
            { recordId: "abc", recordTable: "account", userId: "u1", userName: "Sam" },
            formHost
        );
        expect(r).toMatchObject({ recordId: "abc", recordTable: "account", userId: "u1", userName: "Sam", recordSource: "property" });
    });

    it("treats whitespace properties as empty", () => {
        const r = resolveHostContext({ recordId: "  ", userName: " " }, formHost);
        expect(r.recordSource).toBe("form");
        expect(r.userName).toBe("Alex Example");
    });

    it("sends no record for a new, unsaved form", () => {
        const r = resolveHostContext({}, { mode: { contextInfo: { entityId: "", entityTypeName: "incident" } } });
        expect(r.recordId).toBeNull();
        expect(r.recordTable).toBeNull();
        expect(r.recordSource).toBeNull();
    });

    it("ignores contextInfo that is not a GUID and a table name", () => {
        expect(resolveHostContext({}, { mode: { contextInfo: { entityId: "not-a-guid", entityTypeName: "incident" } } }).recordId).toBeNull();
        expect(resolveHostContext({}, { mode: { contextInfo: { entityId: FORM_ID, entityTypeName: 42 } } }).recordId).toBeNull();
        expect(resolveHostContext({}, { mode: undefined }).recordId).toBeNull();
    });

    it("regression: takes the table from the form when the id comes from the form", () => {
        const warn = jest.spyOn(console, "warn").mockImplementation(() => undefined);
        const r = resolveHostContext({ recordTable: "msdyn_workorder" }, formHost);
        expect(r.recordId).toBe("3f2504e0-4f89-11d3-9a0c-0305e82c3301");
        expect(r.recordTable).toBe("incident");
        expect(warn).toHaveBeenCalled();
        warn.mockRestore();
    });

    it("produces screen context that ContextJson still cannot override", () => {
        const host = resolveHostContext({}, formHost);
        const ctx = buildScreenContext({ ...host, contextJson: JSON.stringify({ recordId: "evil", stage: "Triage" }) });
        expect(ctx.recordId).toBe("3f2504e0-4f89-11d3-9a0c-0305e82c3301");
        expect(ctx.stage).toBe("Triage");
    });
});
