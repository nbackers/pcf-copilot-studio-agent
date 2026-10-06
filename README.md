<div align="center">

# PCF Copilot Studio Agent

**Embed a Copilot Studio agent inside a Power App**

[![Maturity](https://img.shields.io/badge/maturity-implemented-success?style=flat-square)](#status)
[![Build](https://img.shields.io/badge/build-passing-success?style=flat-square)](.github/workflows/build.yml)
[![Tests](https://img.shields.io/badge/tests-31_passing-success?style=flat-square)](tests/)
[![TypeScript](https://img.shields.io/badge/TypeScript-3178C6?style=flat-square&logo=typescript&logoColor=white)](AgentChatWidget/)
[![PCF](https://img.shields.io/badge/PCF-0F6CBD?style=flat-square)](#)
[![Sample code](https://img.shields.io/badge/sample_code-not_production_ready-orange?style=flat-square)](#disclaimer)
[![Licence](https://img.shields.io/badge/licence-MIT-blue?style=flat-square)](LICENSE)

</div>

SSO with silent token exchange, screen context, and an offline demo mode that needs no network.
The control builds and its auth logic is unit tested against mocks; the live path against a real
agent is **not yet verified** - see [Status](#status).

> **Which agents?** The control connects over Direct Line, so it should work with any published
> Copilot Studio agent that exposes that channel, regardless of harness. That is reasoning, not a
> tested result.

---

## The problem

"How do I put my agent inside my app?" is one of the most-asked Power Platform questions and has no
good published answer.

The default is to send users somewhere else - Teams, or a separate chat surface. That breaks the
task they were doing, and it strips the agent of everything it could have known. The user ends up
describing their own screen to an assistant that is running inside the same tenant as the record
they're looking at.

Three things make this harder than it looks:

**Auth is the hard part, and samples skip it.** Most examples use a Direct Line secret, which is
fine on a laptop and unacceptable in production - anyone who can open the app can read it. The SSO
path exists but the piece that makes it work silently is undocumented, so users who are already
signed in get asked to sign in again. Which is precisely what they wanted SSO to avoid.

**The agent has no context.** Without passing it, the agent doesn't know who the user is, what
record is open, or what screen they're on.

**Live agents make fragile demos.** Conference wifi drops. A model takes eleven seconds on the one
question the story was built around, or answers differently than in rehearsal. None of that reflects
the product, but it's what the audience remembers.

## What this solves

| Problem | How this repo solves it |
|---|---|
| Users bounced out of the app | Agent embedded in the app, in a pop-out panel |
| Direct Line secret in production | Both auth paths, with the trade-off stated plainly |
| SSO still prompts for sign-in | OAuth card interception and silent token exchange |
| Agent doesn't know the context | User, record and table passed on conversation start |
| Card buttons silently do nothing | Documented `__isBotFrameworkCardAction` requirement |
| Demos fail on the network | Offline scripted scenarios, no live connection needed |

---

## What's in this repo

**This one builds, lints and tests.**

```powershell
npm install
npm run build    # ESLint + TypeScript + webpack -> out/controls/AgentChatWidget
npm test         # 31 unit tests
npm start        # PCF test harness in a browser
```

| Included | Not included |
|---|---|
| Complete PCF control (TypeScript), builds clean | A packaged `.zip` solution |
| Both auth paths, with the OAuth card interception | A deployed, live-tested agent connection |
| Offline demo mode with five scenarios | Adaptive Card rendering in demo mode (turns render as text) |
| 31 unit tests, including the auth paths against mocked MSAL and Direct Line | An end-to-end test against a real agent |
| CI running lint, test, build and a check that every manifest resource exists | |

**Verification status:** the build, lint and test pipeline runs in CI on every push. The tests cover
the security-sensitive logic against mocks: token exchange success and HTTP failure, silent-to-popup
fallback, replayed OAuth cards, connection allow-listing, and `ContextJson` being unable to overwrite
the user or record. What mocks cannot prove is that a real agent accepts the exchange, so connect it
to your own agent before relying on it.

---

## Auth: two paths, one honest recommendation

### Development - Direct Line secret

```
DirectLineSecret = <your secret>
```

Works immediately. The secret is exchanged for a short-lived token rather than handed to WebChat
directly, which limits what leaks - but **the secret is still in the app**. Anyone who can open it
can read it. Development only.

### Production - SSO

```
TokenEndpointUrl = <Copilot Studio > Settings > Advanced > Metadata>
ClientId         = <Entra app registration client id>
AgentId          = <Copilot Studio Entra agent id>
```

No secret in the app, and the agent knows who it is talking to.

### The part that is actually undocumented

When the agent needs the user's identity it sends an **OAuth card**, which by default renders a
"Sign in" button. A user already signed in to the host app is asked to sign in *again* - the exact
thing SSO was meant to prevent.

The fix is to intercept the card and complete the exchange yourself:

1. Watch `activity$` for attachments of type `application/vnd.microsoft.card.oauth`
2. Acquire a token for `api://botid-{agentId}/.default` - silently, falling back to a popup
3. `POST` a `signin/tokenExchange` invoke back to the conversation

The card resolves itself and the user sees nothing. Implementation in
[`auth.ts`](AgentChatWidget/auth.ts).

Silent acquisition is tried first; the popup is only reached when consent genuinely hasn't been
given. If the exchange fails, the card is left visible so the user can sign in manually rather than
hitting a dead end.

---

## Screen context

This is the reason to embed rather than link.

```
UserId      = User().Email
UserName    = User().FullName
RecordId    = ThisItem.ID
RecordTable = "your_table"
ContextJson = "{""shift"":""afternoon"",""site"":""north""}"
```

Sent to the agent when the conversation starts, so it opens already knowing who the user is and what
they're looking at. The user never has to describe their own screen.

**On a model-driven form, leave these empty.** The control reads the form's record and the signed-in
user from the host, so a case or work order form needs no bindings at all. See
[docs/model-driven.md](docs/model-driven.md).

---

## Demo mode

```
DemoMode     = true
DemoScenario = "incident"
```

Plays a scripted conversation with no network dependency. Five generic scenarios ship with
it - start of shift, floor walk, incident coordination, returning from leave, and team check-in - in [`demo.ts`](AgentChatWidget/demo.ts). Replace them with your own.

**On honesty:** a scripted mode is fine as long as it's labelled. The point of most demos is the
workflow, not proof that a model can form a sentence. Run the live agent for the parts where the
model's own behaviour is what you're demonstrating.

---

## Card buttons that do nothing

A common and baffling failure. WebChat's renderer branches on the shape of `data`:

```
typeof data === 'string'         -> imBack (sends the string as a message)
data.__isBotFrameworkCardAction  -> performCardAction(data.cardAction)
otherwise                        -> postBack (activity has a value but NO text)
```

The Teams `{ msteams: { type: "messageBack" } }` convention is a **Teams-client** thing. WebChat
doesn't read it, so those buttons fall into `postBack` and produce a **textless activity** - nothing
for a trigger phrase to match, and no visible user message. Hence "clicking does nothing".

Shape that works on both channels:

```js
data: {
  __isBotFrameworkCardAction: true,
  cardAction: { type: 'messageBack', text: 'Confirm incident report', displayText: 'Review and confirm' },
  msteams:    { type: 'messageBack', text: 'Confirm incident report', displayText: 'Review and confirm' }
}
```

`text` must equal a trigger phrase. `displayText` is what appears in the transcript.

---

## Contents

| Path | Purpose |
|---|---|
| `AgentChatWidget/index.ts` | The control - lifecycle, panel UI, connection and demo wiring |
| `AgentChatWidget/auth.ts` | Both auth paths and the OAuth card interception |
| `AgentChatWidget/demo.ts` | Offline demo scenarios |
| `AgentChatWidget/css/` | Styles |
| `tests/` | Unit tests |
| `docs/setup.md` | Build, deploy and configure |
| `docs/model-driven.md` | Using the control on a model-driven form, with record context from the host |

---

## Status

The control **builds, lints and passes 31 tests**, verified in CI on every push. The auth logic is
tested against mocked MSAL and Direct Line, including the failure paths.

What has **not** been done here:

- Connecting to a live agent and confirming the end-to-end token exchange. Mocks prove the control
  behaves correctly given a response; they cannot prove a real agent sends that response.
- Verifying against agents on each Copilot Studio harness. Direct Line is a publishing channel
  rather than an authoring concern, so it should not matter, but that is reasoning rather than a
  result. If you try it, please open an issue and say what happened.
- Placing the control on a live model-driven form. The form-record fallback reads
  `context.mode.contextInfo`, which is not in the typed PCF API; it is unit tested against simulated
  host objects only.

### Known limitations

- **Screen context is sent once, at conversation start.** If the user moves to a different record
  while the panel is open, the agent keeps the original context until the conversation restarts.
- **Direct Line secret mode is not blocked at runtime.** It is documented as development only, but
  nothing stops a maker shipping it. Treat a populated `DirectLineSecret` property in a production
  app as a defect.
- **`AllowedConnectionNames` is optional.** Left blank, any OAuth card the agent sends is exchanged
  silently. Set it to the agent's SSO connection name in production.

If the platform ships a first-party control that does this properly, use it.

---

## Disclaimer

This is **sample code**, published as a reusable reference pattern.

- Provided **as is**, without warranty of any kind, express or implied. See [LICENSE](LICENSE).
- **Not production ready.** Treat it as a starting point, not a finished solution. Review, test and
  harden it against your own requirements before any real use.
- **Not an official Microsoft product** and not affiliated with, endorsed by, or supported by
  Microsoft. Product names are trademarks of their respective owners.
- **No support commitment.** Issues and pull requests are welcome, but nothing here carries an SLA.
- Some behaviours documented here rely on **undocumented or preview platform features** that can
  change without notice. Verify against current documentation before depending on them.
- You are responsible for security, privacy, licensing and regulatory compliance in your own
  environment.

---

## Licence

MIT - see [LICENSE](LICENSE).
