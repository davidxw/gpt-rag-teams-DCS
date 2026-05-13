import {
  AzureCliCredential,
  ChainedTokenCredential,
  ManagedIdentityCredential,
} from "@azure/identity";
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
  private credential = new ChainedTokenCredential(
    new ManagedIdentityCredential(),
    new AzureCliCredential()
  );

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
      throw new Error(
        `Orchestrator returned HTTP ${res.status}: ${text.slice(0, 500)}`
      );
    }

    return (await res.json()) as OrchestratorResponse;
  }

  private async getFunctionKey(): Promise<string> {
    if (config.orchestratorFunctionKey) {
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
      return "";
    }

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
      throw new Error(
        `Function listKeys failed (${res.status}): ${text.slice(0, 500)}`
      );
    }

    const body = (await res.json()) as { default?: string };
    if (!body.default) {
      throw new Error("listKeys response did not include 'default' key");
    }

    this.cachedKey = body.default;
    return this.cachedKey;
  }
}
