# gpt-rag-teams

A Microsoft Teams bot front-end for the [GPT-RAG Solution Accelerator](https://github.com/davidxw/GPT-RAG-DCS).

The bot is intentionally thin: it forwards each user turn to the GPT-RAG
**orchestrator** Function App (`POST /api/orc`) and renders the answer back
into Teams. All retrieval, prompting, content filtering, and conversation
persistence are handled by the orchestrator, exactly as for the existing web
front-end.

## Architecture

```
Teams client ──► Bot Service ──► gpt-rag-teams (this repo)
                                          │
                                          ▼
                              POST /api/orc  (x-functions-key)
                                          │
                                          ▼
                     gpt-rag-orchestrator (Azure Functions)
                                          │
                                          ▼
                       Azure OpenAI · AI Search · CosmosDB
```

* `src/index.ts` — bootstraps the Teams SDK [`App`](https://learn.microsoft.com/microsoftteams/platform/teams-sdk/essentials/on-activity/overview)
  from `@microsoft/teams.apps`, which hosts its own HTTP server and exposes
  `POST /api/messages` automatically.
* `src/teamsBot.ts` — single `app.on('message', …)` handler that calls the
  orchestrator, persists the orchestrator-issued `conversation_id` per Teams
  conversation (in-memory `Map`) so multi-turn context continues in CosmosDB,
  and replies with a `MessageActivity(...).addAiGenerated()` so Teams renders
  the [AI generated](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bot-messages-ai-generated-content)
  label on every answer.
* `src/orchestratorClient.ts` — typed HTTP client. In Azure it uses the
  bot's user-assigned managed identity to call ARM `listKeys` and fetch
  the orchestrator's function key (mirrors the `gpt-rag-frontend` pattern).
* `appPackage/manifest.json` — Teams app manifest (schema v1.19).

## Prerequisites

* Node.js **20** and npm.
* A deployed GPT-RAG orchestrator Function App (or run locally on
  `http://localhost:7071`).
* For local Teams testing: the
  [Microsoft 365 Agents Toolkit](https://aka.ms/teams-toolkit) VS Code
  extension (or the Bot Framework Emulator).
* For sideloading: two icons in `appPackage/`:
  * `color.png` — 192×192
  * `outline.png` — 32×32 (transparent background, white outline)

## Local development

```pwsh
npm install
copy .env.example .env       # then fill in the values
npm run dev                  # ts-node-dev with reload
```

Required `.env` values for a minimal local run:

| Variable | Notes |
| --- | --- |
| `ORCHESTRATOR_ENDPOINT` | e.g. `http://localhost:7071/api/orc` |
| `ORCHESTRATOR_FUNCTION_KEY` | Empty if running orchestrator locally with `authLevel=anonymous` for dev. |
| `CLIENT_ID` / `CLIENT_SECRET` / `TENANT_ID` | Teams SDK bot identity. The legacy `MicrosoftAppId` / `MicrosoftAppPassword` / `MicrosoftAppTenantId` names are also accepted (bridged in `src/config.ts`). |

The Teams SDK auto-detects the auth method from these env vars (client
secret, user-assigned MI, or federated identity); see the
[App Authentication reference](https://learn.microsoft.com/microsoftteams/platform/teams-sdk/essentials/app-authentication).

Test in the **Bot Framework Emulator** by pointing it at
`http://localhost:3978/api/messages` and the AAD app credentials above.

## Build & package the Teams app

```pwsh
npm run build                 # tsc -> ./dist
npm run package:teams         # zips appPackage/ into appPackage/build/appPackage.dev.zip
```

Sideload the resulting zip into Teams via **Apps → Manage your apps → Upload an app**.

## Production deployment

This repo deliberately does **not** carry its own Bicep/azd infrastructure.
The hosting App Service (or Container App), user-assigned managed identity,
and Bot Service registration are provisioned by the parent
[GPT-RAG accelerator](https://github.com/davidxw/GPT-RAG-DCS) — add a
`teamsBot` service entry to its `azure.yaml` and run `azd deploy teamsBot`.

The bot expects these App Service application settings (typically Key Vault
references):

* `CLIENT_ID=<MI client id>` *(or the legacy `MicrosoftAppId`)*
* `TENANT_ID=<tenant id>` *(or the legacy `MicrosoftAppTenantId`)*
* `CLIENT_SECRET=` left **empty** when using user-assigned MI
* `ORCHESTRATOR_ENDPOINT=https://<orchestrator-funcapp>.azurewebsites.net/api/orc`
* `AZURE_SUBSCRIPTION_ID`, `AZURE_RESOURCE_GROUP_NAME`, `AZURE_ORCHESTRATOR_FUNC_NAME`
  — used by managed identity to fetch the orchestrator function key at runtime.

The bot's MI must have **Contributor** (or a custom role with
`Microsoft.Web/sites/functions/listKeys/action`) on the orchestrator Function App.

> **Note on `/healthz`:** the previous restify-based server exposed a custom
> `GET /healthz` probe. The Teams SDK App hosts its own HTTP server and
> doesn't currently expose that route. Configure the App Service health
> check at TCP level, or remove the explicit health-check requirement.

## Roadmap

The current implementation covers the chat round-trip plus the Teams
[AI generated](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bot-messages-ai-generated-content)
label on every reply. Planned follow-ups (tracked in
`../GPT-RAG-DCS/.research/teams-integration-review.html`):

* Streaming responses via the Teams SDK `stream` helper (requires an
  orchestrator that supports incremental output).
* Citations rendered via `MessageActivity.addCitation(...)` from
  orchestrator `data_points`.
* 👍/👎 feedback loop wired into the orchestrator's eval pipeline.
* SSO + on-behalf-of for per-user document trimming.
* Proactive notifications (subscription updates, ingestion completions).
