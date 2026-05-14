# Overview of the Basic AI Chatbot template

This app template is built on top of [Microsoft Teams SDK](https://aka.ms/teams-ai-library-v2).
It showcases an agent app that responds to user questions like ChatGPT. This enables your users to talk with the AI agent in Teams.

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 20, 22.
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) latest version or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli).

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
1. In file *env/.env.playground.user*, fill in the URL of the backend orchestrator service in ORCHESTRATOR_ENDPOINT (e.g. http://localhost:7071/api/orc)
1. Press F5 to start debugging which launches your app in Microsoft 365 Agents Playground using a web browser. Select `Debug in Microsoft 365 Agents Playground`.
1. You can send any message to get a response from the agent.

To manually start the application and the playground:

Terminal -> Run Task -> Start application (Microsoft 365 Agents Playground)

Then, when started:

Terminal -> Run Task -> Start Microsoft 365 Agents Playground

**Congratulations**! You are running an application that can now interact with users in Microsoft 365 Agents Playground:

![Basic AI Chatbot](https://github.com/user-attachments/assets/984af126-222b-4c98-9578-0744790b103a)

## What's included in the template

| Folder       | Contents                                            |
| - | - |
| `.vscode`    | VSCode files for debugging                          |
| `appPackage` | Templates for the application manifest        |
| `env`        | Environment files                                   |
| `infra`      | Templates for provisioning Azure resources          |
| `src`        | The source code for the application                 |


The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`m365agents.yml`|This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`m365agents.local.yml`|This overrides `m365agents.yml` with actions that enable local execution and debugging.|
|`m365agents.playground.yml`| This overrides `m365agents.yml` with actions that enable local execution and debugging in Microsoft 365 Agents Playground.|

## Configuring the orchestrator connection

The bot talks to the GPT-RAG orchestrator (the `orc` Azure Function) and to the
`feedback` function in the same Function App. Both endpoints require an
`x-functions-key` header (unless they're set to `authLevel: anonymous`).

### Required variables

| Variable | Purpose |
| - | - |
| `ORCHESTRATOR_ENDPOINT` | Full URL to the `orc` function, **including** `/api/orc` (e.g. `https://<funcapp>.azurewebsites.net/api/orc`). The feedback URL is derived by replacing the trailing `/orc` with `/feedback`. |
| `ORCHESTRATOR_FUNCTION_KEY` | Function or host key for `orc`. Optional if you use the managed-identity fallback below. |
| `FEEDBACK_FUNCTION_KEY` | Function or host key for `feedback`. Optional if you use the managed-identity fallback below. |

### Optional — managed-identity / ARM `listKeys` fallback

If the static keys above are not set, the bot will look up the keys at runtime
via ARM `listKeys` using `ChainedTokenCredential(ManagedIdentity, AzureCli)`.
That requires the bot's identity to have permission to call
`Microsoft.Web/sites/functions/listKeys/action` on the Function App, plus:

| Variable | Purpose |
| - | - |
| `AZURE_SUBSCRIPTION_ID` | Subscription containing the Function App. |
| `AZURE_RESOURCE_GROUP_NAME` | Resource group of the Function App. |
| `AZURE_ORCHESTRATOR_FUNC_NAME` | Function App name (shared between `orc` and `feedback`). |

### Local development (Playground / Teams Toolkit)

There are two layers loaded in this order:

1. `env-cmd` reads `.localConfigs` (for Local) or `.localConfigs.playground`
   (for Playground) — these files are **regenerated on every Toolkit Provision**
   from `env/.env.local` / `env/.env.playground` via the
   `file/createOrUpdateEnvironmentFile` action in
   [`m365agents.local.yml`](m365agents.local.yml) /
   [`m365agents.playground.yml`](m365agents.playground.yml). Editing
   `.localConfigs*` directly is fine for a quick test, but the change will be
   wiped on the next Provision.
2. `dotenv` then loads `.env` **without overwriting** anything already set.

**Recommended:** put secrets in [`.env`](.env) (gitignored, loaded by `dotenv`,
shared across both Local and Playground profiles):

```env
ORCHESTRATOR_FUNCTION_KEY=<key from Function App → orc → Function Keys>
FEEDBACK_FUNCTION_KEY=<key from Function App → feedback → Function Keys>
```

Set the endpoint per profile in `env/.env.local` / `env/.env.playground` so it
follows the Toolkit environment, and re-Provision after changes. Variables that
must reach the Node process via `env-cmd` also need to be plumbed through the
`envs:` block in the matching `m365agents.*.yml` — today that block forwards
`ORCHESTRATOR_ENDPOINT` and `STORAGE_ACCOUNT`. Add new variables there if you
want them per-profile rather than in `.env`.

### Azure (production)

The bot runs on App Service with a system-assigned managed identity. Configure
these as **App Settings** (or Key Vault references) on the bot's App Service:

- `ORCHESTRATOR_ENDPOINT` — the production `orc` URL.
- Either `ORCHESTRATOR_FUNCTION_KEY` + `FEEDBACK_FUNCTION_KEY` (simplest), **or**
  `AZURE_SUBSCRIPTION_ID` + `AZURE_RESOURCE_GROUP_NAME` +
  `AZURE_ORCHESTRATOR_FUNC_NAME` and grant the bot's managed identity the
  `Microsoft.Web/sites/functions/listKeys/action` permission on the Function App
  (the built-in **Website Contributor** role includes it; a narrower custom role
  is preferred).
- `STORAGE_ACCOUNT` — name of the storage account holding source documents.
  The bot reads blobs via its managed identity, so grant **Storage Blob Data
  Reader** on that account.

Static keys are easier to set up; managed-identity lookup avoids storing secrets
and lets the Function App rotate keys without redeploying the bot. The key
lookup is cached in memory for the lifetime of the process.

## Extend the template

To extend the Basic AI Chatbot template with more AI capabilities, explore [Microsoft Teams SDK documentation](https://aka.ms/m365-agents-toolkit/teams-agent-extend-ai).

## Additional information and references

- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)
