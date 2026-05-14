import { DefaultAzureCredential } from "@azure/identity";
import { config } from "./config";

export interface OrchestratorRequest {
  /** Existing conversation id, or omit/empty for the first turn. */
  conversation_id?: string;
  question: string;
  client_principal_id?: string;
  client_principal_name?: string;
  /** Comma-separated list of group object ids (string), as the orchestrator expects. */
  client_group_names?: string;
}

export interface OrchestratorDataPoint {
  /** Source filename, e.g. "ResidentialTenancyAct-2010-042.pdf". */
  fileName?: string;
  /** Document title surfaced in the orchestrator answer markers. */
  title?: string;
  /** Page number the chunk came from (1-based). */
  page?: number;
  /** Chunk text used to ground the LLM answer. */
  content?: string;
  /** Optional pre-built link to the source. */
  url?: string;
  /** Tolerate any other fields the orchestrator might emit. */
  [key: string]: unknown;
}

export interface OrchestratorResponse {
  conversation_id: string;
  answer: string;
  data_points?: OrchestratorDataPoint[] | unknown;
  thoughts?: string;
}

/**
 * Thin HTTP client around the GPT-RAG orchestrator's `orc` HTTP-trigger
 * function. The orchestrator contract is defined by
 * `gpt-rag-orchestrator/orc/__init__.py` and `orchestrator.py`:
 *
 *   POST <ORCHESTRATOR_ENDPOINT>
 *   x-functions-key: <key>
 *   {
 *     "conversation_id": string?,
 *     "question": string,
 *     "client_principal_id": string,
 *     "client_principal_name": string,
 *     "client_group_names": string
 *   }
 *
 * Response:
 *   { "conversation_id": string, "answer": string,
 *     "data_points": ..., "thoughts": ... }
 */
export class OrchestratorClient {
  private cachedKey: string | null = null;
  // DefaultAzureCredential picks the right identity in each environment:
  //   • Azure App Service: managed identity. When a user-assigned MI is
  //     attached, set the App Setting `AZURE_CLIENT_ID` to that UMI's
  //     clientId so DAC unambiguously selects it.
  //   • Local dev: `az login` (AzureCliCredential) or VS Code sign-in.
  private credential = new DefaultAzureCredential();

  async ask(req: OrchestratorRequest): Promise<OrchestratorResponse> {
    const key = await this.getFunctionKey();

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
    };
    if (key) {
      headers["x-functions-key"] = key;
    }

    const res = await fetch(config.orchestratorEndpoint, {
      method: "POST",
      headers,
      body: JSON.stringify(req),
    });

    if (!res.ok) {
      const text = await res.text().catch(() => "<no body>");
      console.error(
        `[orchestratorClient] ✗ orchestrator HTTP ${res.status} | endpoint=${config.orchestratorEndpoint} | ` +
          `keyHeaderSent=${key ? "yes" : "NO"} | body=${text.slice(0, 500)}`
      );
      throw new Error(
        `Orchestrator returned HTTP ${res.status}: ${text.slice(0, 500)}`
      );
    }

    return (await res.json()) as OrchestratorResponse;
  }

  private async getFunctionKey(): Promise<string> {
    if (config.orchestratorFunctionKey) {
      console.log(
        "[orchestratorClient] using static ORCHESTRATOR_FUNCTION_KEY env var"
      );
      return config.orchestratorFunctionKey;
    }
    if (this.cachedKey) {
      return this.cachedKey;
    }

    if (
      !config.azureSubscriptionId ||
      !config.azureResourceGroup ||
      !config.azureOrchestratorFuncName
    ) {
      // No static key and no MI lookup info — assume the endpoint is
      // anonymous (e.g. private endpoint with network ACLs only).
      console.warn(
        "[orchestratorClient] ⚠ no ORCHESTRATOR_FUNCTION_KEY and missing one of " +
          "AZURE_SUBSCRIPTION_ID / AZURE_RESOURCE_GROUP_NAME / AZURE_ORCHESTRATOR_FUNC_NAME — " +
          "sending request with NO x-functions-key header (will 401 if function auth is required)"
      );
      return "";
    }

    console.log(
      `[orchestratorClient] fetching function key via MI listKeys | sub=${config.azureSubscriptionId} | ` +
        `rg=${config.azureResourceGroup} | func=${config.azureOrchestratorFuncName}`
    );

    const token = await this.credential.getToken(
      "https://management.azure.com/.default"
    );
    if (!token) {
      throw new Error("Could not acquire ARM token for function key lookup");
    }

    const url =
      `https://management.azure.com/subscriptions/${config.azureSubscriptionId}` +
      `/resourceGroups/${config.azureResourceGroup}` +
      `/providers/Microsoft.Web/sites/${config.azureOrchestratorFuncName}` +
      `/functions/orc/listKeys?api-version=2022-03-01`;

    const res = await fetch(url, {
      method: "POST",
      headers: { Authorization: `Bearer ${token.token}` },
    });

    if (!res.ok) {
      const text = await res.text().catch(() => "<no body>");
      console.error(
        `[orchestratorClient] ✗ ARM listKeys ${res.status} for orc | url=${url} | body=${text.slice(0, 500)} | ` +
          `(403 usually means the bot's managed identity lacks 'Microsoft.Web/sites/functions/listKeys/action' on the function app)`
      );
      throw new Error(
        `Function listKeys failed (${res.status}): ${text.slice(0, 500)}`
      );
    }

    const body = (await res.json()) as { default?: string };
    if (!body.default) {
      throw new Error("listKeys response did not include 'default' key");
    }

    console.log(
      "[orchestratorClient] ✓ fetched orc function key via MI listKeys (cached for process lifetime)"
    );
    this.cachedKey = body.default;
    return this.cachedKey;
  }
}
