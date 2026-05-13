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

export const config = {
  port: parseInt(process.env.PORT ?? "3978", 10),

  // --- Bot identity ---
  // Local dev:    MicrosoftAppType=MultiTenant + MicrosoftAppId + MicrosoftAppPassword
  // Production:   MicrosoftAppType=UserAssignedMSI + MicrosoftAppId=<MI client id> + MicrosoftAppTenantId
  botType: optional("MicrosoftAppType") || "MultiTenant",
  botId: optional("MicrosoftAppId"),
  botPassword: optional("MicrosoftAppPassword"),
  botTenantId: optional("MicrosoftAppTenantId"),

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
};
