import * as dotenv from "dotenv";

dotenv.config();

function optional(name: string): string {
  return process.env[name] ?? "";
}

function required(name: string): string {
  const v = process.env[name];
  if (!v) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return v;
}

// --- Teams SDK env-var bridge ---------------------------------------------
// The new Teams SDK (`@microsoft/teams.apps`, replacing the deprecated
// `@microsoft/teams-ai` v1 / "Teams AI library") reads bot identity from
// CLIENT_ID / CLIENT_SECRET / TENANT_ID env vars. Existing App Service
// settings provisioned by the parent GPT-RAG accelerator still use the
// legacy MicrosoftApp* names, so we forward them here to keep both worlds
// working without re-provisioning.
//
// See: https://learn.microsoft.com/microsoftteams/platform/teams-sdk/essentials/app-authentication
function bridge(target: string, fallback: string): void {
  if (!process.env[target] && process.env[fallback]) {
    process.env[target] = process.env[fallback];
  }
}
bridge("CLIENT_ID", "MicrosoftAppId");
bridge("CLIENT_SECRET", "MicrosoftAppPassword");
bridge("TENANT_ID", "MicrosoftAppTenantId");

export const config = {
  port: parseInt(process.env.PORT ?? "3978", 10),

  // --- Bot identity (Teams SDK reads these directly from env) -------------
  // Local dev:  CLIENT_ID + CLIENT_SECRET + TENANT_ID (client-secret auth).
  // Production: CLIENT_ID + TENANT_ID with the Web App's user-assigned
  //             managed identity attached (passwordless / UMI auth).
  clientId: optional("CLIENT_ID"),
  clientSecret: optional("CLIENT_SECRET"),
  tenantId: optional("TENANT_ID"),

  // --- Orchestrator ---
  // Full URL of the orchestrator HTTP function, e.g.
  //   https://<funcapp>.azurewebsites.net/api/orc
  // For local dev against the orchestrator running on Functions Core Tools:
  //   http://localhost:7071/api/orc
  orchestratorEndpoint: required("ORCHESTRATOR_ENDPOINT"),

  // Optional: a static function key for the orchestrator (dev/test).
  // If omitted in Azure, the bot will use its managed identity to call
  // ARM `listKeys` and fetch the orchestrator function key dynamically
  // (mirrors the pattern used by gpt-rag-frontend).
  orchestratorFunctionKey: optional("ORCHESTRATOR_FUNCTION_KEY"),

  // --- Required only when using MI to fetch the function key ---
  azureSubscriptionId: optional("AZURE_SUBSCRIPTION_ID"),
  azureResourceGroup: optional("AZURE_RESOURCE_GROUP_NAME"),
  azureOrchestratorFuncName: optional("AZURE_ORCHESTRATOR_FUNC_NAME"),

  // --- Citations / document viewing -------------------------------------
  // Storage account name that holds the source documents. The bot reads
  // blobs directly using its managed identity (DefaultAzureCredential) and
  // exposes them via GET /api/documents. Citation URLs point back to the
  // bot, never to storage. Leave blank to disable click-through.
  //
  // Required role on the storage account: `Storage Blob Data Reader` for
  // the bot's user-assigned managed identity.
  storageAccount: optional("STORAGE_ACCOUNT"),

  // Container that holds the documents. Defaults to `documents` to match
  // the GPT-RAG accelerator convention.
  storageContainer: process.env.STORAGE_CONTAINER || "documents",

  // Public host (no scheme) used to build absolute citation URLs back to
  // this bot's GET /api/documents endpoint. Auto-populated from common
  // hosts: BOT_DOMAIN (Toolkit local dev tunnel) → WEBSITE_HOSTNAME
  // (Azure App Service). Override with BOT_BASE_URL if needed.
  botBaseUrl:
    process.env.BOT_BASE_URL ||
    (process.env.BOT_DOMAIN ? `https://${process.env.BOT_DOMAIN}` : "") ||
    (process.env.WEBSITE_HOSTNAME ? `https://${process.env.WEBSITE_HOSTNAME}` : ""),

  // Number of characters of the orchestrator's raw answer to print to
  // the console for debugging citation/markdown issues. Set to 0 to
  // disable. Default 500.
  rawAnswerLogChars: parseInt(process.env.RAW_ANSWER_LOG_CHARS ?? "500", 10),
};
