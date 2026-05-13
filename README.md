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

* `src/index.ts` — restify server, Bot Framework `CloudAdapter`.
* `src/teamsBot.ts` — Teams AI `Application`; one `message` handler that
  calls the orchestrator and persists `orchestratorConversationId` per
  Teams conversation so multi-turn context is preserved in CosmosDB.
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
| `MicrosoftAppId` / `MicrosoftAppPassword` | From your Bot Service / AAD app registration. |

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

* `MicrosoftAppType=UserAssignedMSI`
* `MicrosoftAppId=<MI client id>`
* `MicrosoftAppTenantId=<tenant id>`
* `ORCHESTRATOR_ENDPOINT=https://<orchestrator-funcapp>.azurewebsites.net/api/orc`
* `AZURE_SUBSCRIPTION_ID`, `AZURE_RESOURCE_GROUP_NAME`, `AZURE_ORCHESTRATOR_FUNC_NAME`
  — used by managed identity to fetch the orchestrator function key at runtime.

The bot's MI must have **Contributor** (or a custom role with
`Microsoft.Web/sites/functions/listKeys/action`) on the orchestrator Function App.

## Roadmap

The current implementation focuses on the chat round-trip only. Planned
follow-ups (tracked in `../GPT-RAG-DCS/.research/teams-integration-review.html`):

* Streaming responses via `streamingResponse`.
* Citations rendered as adaptive-card source chips.
* AI-generated content label + 👍/👎 feedback loop.
* SSO + on-behalf-of for per-user document trimming.
* Proactive notifications (subscription updates, ingestion completions).
