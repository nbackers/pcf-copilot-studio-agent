/**
 * Authentication for an embedded Copilot Studio agent.
 *
 * Two modes:
 *
 *   Development - a Direct Line secret is exchanged for a token in the browser.
 *                 Simple, and unsuitable for anything real: the secret is readable
 *                 by anyone who can open the app.
 *
 *   SSO         - the Copilot Studio token endpoint issues the Direct Line token, and
 *                 the signed-in user's identity is passed to the agent through an
 *                 OAuth card exchange. The agent knows who it is talking to.
 *
 * The SSO path is the poorly documented one, and interceptOAuthCards() below is the
 * specific piece most samples omit.
 */

const WEBCHAT_CDN = "https://cdn.botframework.com/botframework-webchat/latest/webchat.js";
const MSAL_CDN = "https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js";
const DIRECTLINE_HOST = "https://directline.botframework.com";

/** Minimal shape of the WebChat global we rely on. */
export interface WebChatGlobal {
    createDirectLine(options: { token: string }): DirectLineConnection;
    renderWebChat(props: unknown, element: HTMLElement): void;
    createStore(
        initialState: unknown,
        middleware: (context: StoreContext) => (next: DispatchFn) => DispatchFn
    ): unknown;
}

export interface StoreContext {
    dispatch: (action: StoreAction) => void;
}

export type DispatchFn = (action: StoreAction) => unknown;

export interface StoreAction {
    type: string;
    payload?: Record<string, unknown>;
}

export interface DirectLineActivity {
    id?: string;
    type: string;
    attachments?: { contentType: string; content?: { connectionName?: string } }[];
}

export interface DirectLineConnection {
    activity$: {
        subscribe(observer: { next: (activity: DirectLineActivity) => void }): { unsubscribe(): void };
    };
}

interface MsalAccount { username?: string }
interface MsalResult { accessToken: string }

interface MsalInstance {
    initialize(): Promise<void>;
    getAllAccounts(): MsalAccount[];
    acquireTokenSilent(request: { scopes: string[]; account?: MsalAccount }): Promise<MsalResult>;
    acquireTokenPopup(request: { scopes: string[] }): Promise<MsalResult>;
}

interface MsalGlobal {
    PublicClientApplication: new (config: unknown) => MsalInstance;
}

declare global {
    interface Window {
        WebChat?: WebChatGlobal;
        msal?: MsalGlobal;
    }
}

/** Load a script once, resolving when it is ready. */
export function loadScript(src: string): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        if (document.querySelector(`script[src="${src}"]`)) {
            resolve();
            return;
        }
        const el = document.createElement("script");
        el.src = src;
        el.onload = () => resolve();
        el.onerror = () => reject(new Error(`Failed to load ${src}`));
        document.head.appendChild(el);
    });
}

/** Resolve the WebChat global after its script has loaded. */
function requireWebChat(): WebChatGlobal {
    if (!window.WebChat) {
        throw new Error("WebChat did not load. Check that the CDN domain is allowed in the manifest.");
    }
    return window.WebChat;
}

/**
 * Development connection from a Direct Line secret.
 *
 * The secret is exchanged for a short-lived token rather than handed to WebChat
 * directly, which limits what leaks if the page is inspected. It does not make the
 * secret safe - it is still present in the app.
 */
export async function connectWithSecret(secret: string): Promise<DirectLineConnection> {
    const response = await fetch(`${DIRECTLINE_HOST}/v3/directline/tokens/generate`, {
        method: "POST",
        headers: {
            Authorization: `Bearer ${secret}`,
            "Content-Type": "application/json"
        }
    });

    if (!response.ok) {
        throw new Error(`Direct Line token exchange failed: ${response.status}`);
    }

    const { token } = (await response.json()) as { token: string };

    await loadScript(WEBCHAT_CDN);
    return requireWebChat().createDirectLine({ token });
}

export interface SsoOptions {
    tokenEndpointUrl: string;
    clientId: string;
    agentId: string;
    allowedConnectionNames?: string[];
}

export interface SsoConnection {
    connection: DirectLineConnection;
    cleanup: () => void;
}

/**
 * SSO connection.
 *
 * The token endpoint is published by Copilot Studio under Settings > Advanced >
 * Metadata. It returns both a Direct Line token and a conversation id; the conversation
 * id is required to post the token exchange back.
 */
export async function connectWithSso(options: SsoOptions): Promise<SsoConnection> {
    const { tokenEndpointUrl, clientId, agentId, allowedConnectionNames } = options;

    if (!tokenEndpointUrl) throw new Error("TokenEndpointUrl is required for SSO mode");
    if (!clientId) throw new Error("ClientId is required for SSO mode");
    if (!agentId) throw new Error("AgentId is required for SSO mode");

    await Promise.all([loadScript(MSAL_CDN), loadScript(WEBCHAT_CDN)]);

    if (!window.msal) {
        throw new Error("MSAL did not load. Check that login.microsoftonline.com is allowed in the manifest.");
    }

    const msalInstance = new window.msal.PublicClientApplication({
        auth: {
            clientId,
            authority: "https://login.microsoftonline.com/common",
            redirectUri: window.location.origin
        },
        // sessionStorage rather than localStorage: the token should not outlive the tab.
        cache: { cacheLocation: "sessionStorage" }
    });

    await msalInstance.initialize();

    const response = await fetch(tokenEndpointUrl);
    if (!response.ok) {
        throw new Error(`Token endpoint failed: ${response.status}`);
    }

    const { token, conversationId } = (await response.json()) as {
        token: string;
        conversationId: string;
    };

    const connection = requireWebChat().createDirectLine({ token });

    const cleanup = interceptOAuthCards({
        connection,
        conversationId,
        directLineToken: token,
        msalInstance,
        agentId,
        allowedConnectionNames
    });

    return { connection, cleanup };
}

interface InterceptOptions {
    connection: DirectLineConnection;
    conversationId: string;
    directLineToken: string;
    msalInstance: MsalInstance;
    agentId: string;
    /**
     * Optional allow-list of OAuth connection names. When set, cards for any other
     * connection are ignored rather than silently exchanged.
     */
    allowedConnectionNames?: string[];
}

/**
 * Watch for OAuth cards and complete the exchange silently.
 *
 * This is the piece that makes SSO feel like SSO.
 *
 * When the agent needs the user's identity it sends an OAuth card, which by default
 * renders a "Sign in" button - so a user already signed in to the host app is asked to
 * sign in again. That is exactly what people are trying to avoid when they ask for SSO.
 *
 * Instead: intercept the card, acquire a token for the agent's scope, and post a
 * signin/tokenExchange invoke back into the conversation. The card resolves itself and
 * the user sees nothing.
 *
 * Silent acquisition is attempted first and falls back to a popup, which is only reached
 * when consent has genuinely not been given.
 */
export function interceptOAuthCards(options: InterceptOptions): () => void {
    const { connection, conversationId, directLineToken, msalInstance, agentId, allowedConnectionNames } = options;
    const scope = `api://botid-${agentId}/.default`;

    // Direct Line can replay activities on reconnect. Exchange each card once only,
    // or a replay triggers repeated silent token acquisition.
    const handled = new Set<string>();

    const subscription = connection.activity$.subscribe({
        next: (activity: DirectLineActivity) => {
            if (activity.type !== "message" || !activity.attachments) return;

            for (const attachment of activity.attachments) {
                if (attachment.contentType !== "application/vnd.microsoft.card.oauth") continue;

                const connectionName = attachment.content?.connectionName;
                if (!connectionName) continue;

                if (allowedConnectionNames && !allowedConnectionNames.includes(connectionName)) {
                    console.warn(`Ignoring OAuth card for unexpected connection '${connectionName}'.`);
                    continue;
                }

                const activityId = activity.id ?? "";
                const key = `${activityId}|${connectionName}`;
                if (handled.has(key)) continue;
                handled.add(key);

                void exchangeToken({
                    activityId,
                    connectionName,
                    conversationId,
                    directLineToken,
                    msalInstance,
                    scope
                });
            }
        }
    });

    return () => subscription.unsubscribe();
}

export interface ExchangeOptions {
    activityId: string;
    connectionName: string;
    conversationId: string;
    directLineToken: string;
    msalInstance: MsalInstance;
    scope: string;
}

/**
 * Exchange a user token for the agent's OAuth card.
 *
 * Returns true only when Direct Line accepted the exchange. Any failure leaves the card
 * visible so the user can sign in manually, rather than hitting a dead end.
 */
export async function exchangeToken(options: ExchangeOptions): Promise<boolean> {
    const { activityId, connectionName, conversationId, directLineToken, msalInstance, scope } = options;

    try {
        const account = msalInstance.getAllAccounts()[0];

        let result: MsalResult;
        try {
            result = await msalInstance.acquireTokenSilent({ scopes: [scope], account });
        } catch {
            // Only reached when consent has not been given.
            result = await msalInstance.acquireTokenPopup({ scopes: [scope] });
        }

        const response = await fetch(`${DIRECTLINE_HOST}/v3/directline/conversations/${conversationId}/activities`, {
            method: "POST",
            headers: {
                Authorization: `Bearer ${directLineToken}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                type: "invoke",
                name: "signin/tokenExchange",
                value: { id: activityId, connectionName, token: result.accessToken },
                from: { id: "user" }
            })
        });

        // fetch resolves on HTTP errors. Without this check a 401 or 500 looks like a
        // completed exchange and the user is left with a card that never resolves.
        if (!response.ok) {
            throw new Error(`Direct Line rejected the token exchange: ${response.status}`);
        }

        return true;
    } catch (error) {
        // Leave the card visible so the user can sign in manually rather than hitting a
        // dead end.
        console.warn("Silent token exchange failed, falling back to the card.", error);
        return false;
    }
}

export interface ScreenContextInput {
    userId?: string | null;
    userName?: string | null;
    recordId?: string | null;
    recordTable?: string | null;
    contextJson?: string | null;
}

/** The parts of the PCF context that resolveHostContext reads. Kept structural for testing. */
export interface HostContextSource {
    mode?: unknown;
    userSettings?: { userId?: string; userName?: string } | null;
}

const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function normaliseGuid(value: unknown): string | null {
    if (typeof value !== "string") return null;
    const trimmed = value.trim().replace(/^\{|\}$/g, "").toLowerCase();
    return GUID_PATTERN.test(trimmed) ? trimmed : null;
}

function nonEmpty(value: string | null | undefined): string | null {
    return typeof value === "string" && value.trim() !== "" ? value.trim() : null;
}

/**
 * Fill user and record from the host when the properties are left empty.
 *
 * Explicit properties always win, so a canvas app keeps full control. On a model-driven
 * form the record's own id cannot be bound to a property (the primary key is not a
 * bindable column), so the form record is read from `context.mode.contextInfo`, and the
 * user from `context.userSettings`. `contextInfo` is not part of the typed PCF API, so it
 * is read defensively: anything that is not a GUID and a table name is ignored.
 */
export function resolveHostContext(
    props: Pick<ScreenContextInput, "userId" | "userName" | "recordId" | "recordTable">,
    host: HostContextSource
): Required<Pick<ScreenContextInput, "userId" | "userName" | "recordId" | "recordTable">> & {
    recordSource: "property" | "form" | null;
} {
    let recordId = nonEmpty(props.recordId);
    let recordTable = nonEmpty(props.recordTable);
    let recordSource: "property" | "form" | null = recordId ? "property" : null;

    if (!recordId) {
        const info = (host.mode as { contextInfo?: { entityId?: unknown; entityTypeName?: unknown } } | undefined)
            ?.contextInfo;
        const formId = normaliseGuid(info?.entityId);
        const formTable = typeof info?.entityTypeName === "string" ? nonEmpty(info.entityTypeName) : null;
        // A new, unsaved record has no id yet. Send nothing rather than a half-populated record.
        if (formId && formTable) {
            // The id and table must describe the same record, so both come from the form.
            if (recordTable && recordTable.toLowerCase() !== formTable.toLowerCase()) {
                console.warn(
                    `RecordTable '${recordTable}' was ignored: RecordId is empty, so the form's own record (${formTable}) is used.`
                );
            }
            recordId = formId;
            recordTable = formTable;
            recordSource = "form";
        }
    }

    return {
        userId: nonEmpty(props.userId) ?? nonEmpty(host.userSettings?.userId) ?? null,
        userName: nonEmpty(props.userName) ?? nonEmpty(host.userSettings?.userName) ?? null,
        recordId,
        recordTable: recordId ? recordTable : nonEmpty(props.recordTable),
        recordSource
    };
}

/** Fields set from trusted control properties. ContextJson may never overwrite them. */
const RESERVED_CONTEXT_KEYS = new Set(["userId", "userName", "recordId", "recordTable"]);

/** Keys that would let a JSON payload tamper with object prototypes. */
const UNSAFE_KEYS = new Set(["__proto__", "constructor", "prototype"]);

/** Generous for screen context, small enough to stop a bound field flooding the agent. */
export const MAX_CONTEXT_JSON_LENGTH = 4000;

/**
 * Build the screen context sent to the agent when the conversation starts.
 *
 * This is the reason to embed an agent rather than link to one. The agent starts knowing
 * who the user is and what they are looking at, so the user does not have to describe
 * their own screen.
 *
 * ContextJson is treated as untrusted: it is often bound to app data, so it must not be
 * able to impersonate another user or point the agent at a different record.
 */
export function buildScreenContext(props: ScreenContextInput): Record<string, unknown> {
    const context: Record<string, unknown> = {
        userId: props.userId ?? null,
        userName: props.userName ?? null,
        recordId: props.recordId ?? null,
        recordTable: props.recordTable ?? null
    };

    if (props.contextJson) {
        if (props.contextJson.length > MAX_CONTEXT_JSON_LENGTH) {
            console.warn(`ContextJson exceeds ${MAX_CONTEXT_JSON_LENGTH} characters and was ignored.`);
        } else {
            let parsed: unknown;
            try {
                parsed = JSON.parse(props.contextJson);
            } catch {
                console.warn("ContextJson is not valid JSON and was ignored.");
            }

            if (parsed !== undefined) {
                if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) {
                    console.warn("ContextJson must be a JSON object and was ignored.");
                } else {
                    for (const [key, value] of Object.entries(parsed as Record<string, unknown>)) {
                        if (RESERVED_CONTEXT_KEYS.has(key)) {
                            console.warn(`ContextJson cannot override '${key}'; the value was ignored.`);
                            continue;
                        }
                        if (UNSAFE_KEYS.has(key)) continue;
                        context[key] = value;
                    }
                }
            }
        }
    }

    // Drop empties so the agent is not handed a payload of nulls.
    return Object.fromEntries(
        Object.entries(context).filter(([, v]) => v !== null && v !== "")
    );
}
