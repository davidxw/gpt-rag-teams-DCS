# GitHub Copilot instructions — gpt-rag-teams

## What this repo is

A Microsoft Teams bot that acts as a **thin chat front-end** for the
[GPT-RAG Solution Accelerator](https://github.com/davidxw/GPT-RAG-DCS)
orchestrator. It does not do retrieval, prompting, content filtering, or
conversation persistence on its own — those all live in the orchestrator.

## Target orchestrator version

This bot targets **GPT-RAG orchestrator v1.0.1**.

The HTTP contract — exposed by the `orc` Azure Function in
`gpt-rag-orchestrator` v1.0.1 — is:

```
POST <ORCHESTRATOR_ENDPOINT>            (e.g. https://<func>.azurewebsites.net/api/orc)
Headers:
  Content-Type: application/json
  x-functions-key: <key>                (authLevel = function)

Request body:
{
  "conversation_id":       string?,     // empty/omitted on the first turn
  "question":              string,
  "client_principal_id":   string,
  "client_principal_name": string,
  "client_group_names":    string       // comma-separated group object ids
}

Response body:
{
  "conversation_id": string,
  "answer":          string,
  "data_points":     unknown,           // array of source snippets
  "thoughts":        string             // chain-of-thought / debug trace
}
```

**Do not** invent fields, change casing, or assume streaming — v1.0.1 is a
single non-streaming JSON request/response. If you need to add features
that depend on a newer orchestrator contract, gate them behind a config
flag and document the minimum orchestrator version required.

## Architecture & key files

* [`src/index.ts`](../src/index.ts) — thin bootstrap that calls
  `await app.start(config.port)` on the Teams SDK `App`. The SDK hosts
  its own HTTP server and exposes `POST /api/messages` automatically.
  There is no longer a custom `/healthz` route.
* [`src/teamsBot.ts`](../src/teamsBot.ts) — Teams SDK `App` from
  `@microsoft/teams.apps` with a single `app.on('message', …)` handler.
  Persists the orchestrator-issued `conversation_id` per Teams
  conversation in an in-process `Map<string,string>` keyed by
  `activity.conversation.id` so multi-turn context is preserved
  server-side in CosmosDB. Replies are sent as
  `new MessageActivity(answer).addAiGenerated()` so Teams renders the
  [AI generated](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bot-messages-ai-generated-content)
  label.
* [`src/orchestratorClient.ts`](../src/orchestratorClient.ts) — typed
  HTTP client. Resolves the function key in this order:
  1. `ORCHESTRATOR_FUNCTION_KEY` env var (dev / static).
  2. ARM `listKeys` on `Microsoft.Web/sites/{func}/functions/orc` using
     `ChainedTokenCredential(ManagedIdentityCredential, AzureCliCredential)`.
     The key is cached on the instance after first lookup.
* [`src/config.ts`](../src/config.ts) — env loader. All env access goes
  through here; do not call `process.env` directly elsewhere. Also
  bridges the legacy `MicrosoftAppId`/`MicrosoftAppPassword`/
  `MicrosoftAppTenantId` names to the Teams SDK's expected
  `CLIENT_ID`/`CLIENT_SECRET`/`TENANT_ID` so existing App Service
  settings keep working.
* [`appPackage/manifest.json`](../appPackage/manifest.json) — Teams app
  manifest, schema **v1.19**. Toolkit placeholders `${{TEAMS_APP_ID}}`
  and `${{BOT_ID}}` are filled in by `teamsApp/create` /
  `botFramework/create` during provision.

## Tech stack & conventions

* **Node 20**, TypeScript strict, `commonjs` modules, ES2022 target.
  Output to `./dist`. Source under `./src`. Do not add ESM-only deps.
* Teams SDK (the renamed Teams AI library v2): `@microsoft/teams.apps`
  for the `App` class, `@microsoft/teams.api` for `MessageActivity` and
  related activity types, `@microsoft/teams.common` for logging. **Do
  not** reintroduce `@microsoft/teams-ai` or `botbuilder` — those are
  the deprecated v1 stack.
* The Teams SDK `App` hosts its own HTTP server. Don't add restify /
  Express / a `CloudAdapter` wrapper; just call `await app.start(port)`.
* Azure auth: `@azure/identity` ^4, always via `ChainedTokenCredential`
  so the same code path works locally (`az login`) and in Azure (MI).
* Use the global `fetch` (Node 20 built-in). Do not add `node-fetch` /
  `axios` / `got`.

## What this bot must NOT do

* Do not call Azure OpenAI, AI Search, CosmosDB, Storage, or any
  knowledge source directly. Everything goes through the orchestrator.
* Do not implement its own prompt engineering, content filtering, RAI,
  or citation rendering logic. The orchestrator owns these.
* Do not embed secrets, tenant ids, or function keys in source. Local
  dev uses `.env` (gitignored); Azure uses App Service settings backed
  by Key Vault references.
* Do not add infrastructure (Bicep / Terraform / azd) here. Hosting is
  provisioned by the parent accelerator's `azure.yaml` `teamsBot`
  service entry. This repo only owns the Teams app package + bot code.

## Deferred features (do not add unless asked)

The current implementation covers chat round-trip + the Teams
[AI generated](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bot-messages-ai-generated-content)
label. The following are planned and tracked in
`../GPT-RAG-DCS/.research/teams-integration-review.html`. Don't preempt
them — they each have design considerations:

* Streaming responses via the Teams SDK `stream` helper (requires
  orchestrator changes; not available in v1.0.1).
* Citations rendered via `MessageActivity.addCitation(...)` from
  `data_points`.
* 👍/👎 feedback loop.
* SSO + on-behalf-of for per-user document trimming (uses `access_token`
  field on the orchestrator request — also not in v1.0.1).
* Proactive notifications (subscription updates, ingestion completions).
* Restoring a `/healthz` endpoint (the Teams SDK App doesn't expose one
  by default; would need a plugin or sidecar).

## When making changes

1. Type-check before declaring done: `npx tsc --noEmit`.
2. Keep `src/orchestratorClient.ts` the *only* place that knows the
   orchestrator wire format. Export typed interfaces from it; consumers
   import the types.
3. If you change the request or response shape, also update:
   - The contract block at the top of this file.
   - The `OrchestratorRequest` / `OrchestratorResponse` interfaces.
   - The README's Architecture section.
   - The minimum-orchestrator-version note above.
4. Manifest changes must keep `manifestVersion: "1.19"` and preserve the
   `${{TEAMS_APP_ID}}` / `${{BOT_ID}}` Toolkit placeholders.
5. For any new env var: add it to `src/config.ts`, `.env.example`,
   `env/.env.dev`, `env/.env.prod`, and the README env table — in that
   order, in a single change.
