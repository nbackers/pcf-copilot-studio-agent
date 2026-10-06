# Using the control on a model-driven form

The same control works on a model-driven form, where it is arguably more useful: the agent opens
already knowing which case, work order or account the user is looking at.

## What happens automatically

Leave `UserId`, `UserName`, `RecordId` and `RecordTable` **empty** on the form and the control fills
them from the host:

| Context | Source on a model-driven form |
|---|---|
| `recordId` | The form's record, from `context.mode.contextInfo.entityId`, normalised to a lowercase GUID |
| `recordTable` | The form's table, from `context.mode.contextInfo.entityTypeName` |
| `userId`, `userName` | The signed-in user, from `context.userSettings` |

Anything you set on a property wins, with one rule: the record id and table always come from the
same place. If `RecordId` is empty, the form's own table is used and a `RecordTable` value is
ignored with a console warning. On a new, unsaved record there is no id yet, so no record is sent
rather than a half-populated one.

`contextInfo` is not part of the typed PCF API. The control reads it defensively and ignores
anything that is not a GUID and a table name. This fallback is unit tested against simulated host
objects; it has not been verified on every form type (quick create, main form dialog, custom page).

## Add it to a form

1. Import the solution containing the control.
2. Open the table's main form in the form designer.
3. Add a single-line text column to the form for the control to sit on. The manifest needs one bound
   text property (**Bound Field**); the control does not read or write its value. A dedicated,
   hidden-label column such as `Agent panel` keeps it out of the way.
4. Select that column, then **Components** > **Add component** > the agent chat control.
5. Set the connection properties (`TokenEndpointUrl` or `DirectLineSecret`, `ClientId`, `AgentId`,
   `AllowedConnectionNames`) as static values. Leave the context properties empty.
6. Give the section enough height: a full-width, one-column section of at least 500 pixels works.
7. Save and publish.

## Pass extra context

Static JSON can be set on `ContextJson` in the form designer, for example the form's purpose:

```json
{"surface":"case-form","team":"tier-2"}
```

To send live field values, bind `ContextJson` to a text column that a business rule, plug-in or
calculated column keeps populated. `ContextJson` is treated as untrusted: it cannot override the
user or record, it must be a JSON object, and it is capped at 4,000 characters.

## What the agent receives

On a case form, with nothing set:

```json
{
  "userId": "{11111111-2222-3333-4444-555555555555}",
  "userName": "Alex Example",
  "recordId": "3f2504e0-4f89-11d3-9a0c-0305e82c3301",
  "recordTable": "incident",
  "surface": "case-form"
}
```

In the agent, read these from the `startConversation` event and use `recordId` and `recordTable` to
look up the record with the Dataverse connector. Do not trust the agent to have read access just
because the user can see the form; scope the agent's connection and rely on Dataverse security roles.

## Canvas apps and custom pages

Nothing changes. Set the properties explicitly as in [setup.md](setup.md#pass-screen-context); the
host fallback only applies when they are empty.
