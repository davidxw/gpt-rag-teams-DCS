import {
  AzureCliCredential,
  ChainedTokenCredential,
  ManagedIdentityCredential,
} from "@azure/identity";
import { config } from "./config";

export interface FeedbackPayload {
  conversation_id: string;
  question: string;
  answer: string;
  rating: string;
  comment: string;
  message_index: number;
  client_principal_id: string;
  client_principal_name: string;
  /**
   * Origin of the feedback (e.g. "teams", "frontend"). Lets the backend
   * distinguish feedback coming from this Teams bot vs. other clients.
   */
  source: string;
  /**
   * Array of group object ids. Sent for parity with the feedback
   * contract; this bot does not currently resolve the user's groups, so
   * the value is typically an empty array.
   */
  client_group_names: string[];
}

/**
 * Thin HTTP client around the GPT-RAG `feedback` HTTP-trigger function.
 *
 * The feedback function lives in the same Function App as `orc`, so the
 * URL is derived from `ORCHESTRATOR_ENDPOINT` by replacing the trailing
 * `/orc` segment with `/feedback` (mirroring the Python frontend's
 * `base_url = ORCHESTRATOR_ENDPOINT.rsplit('/', 1)[0]; url = base_url + '/feedback'`).
 *
 * Function key resolution mirrors `OrchestratorClient`:
 *   1. `FEEDBACK_FUNCTION_KEY` env var (dev / static).
 *   2. ARM `listKeys` on `Microsoft.Web/sites/{func}/functions/feedback`
 *      using `ChainedTokenCredential(ManagedIdentity, AzureCli)`. Cached
 *      after the first lookup.
 */
export class FeedbackClient {
  private cachedKey: string | null = null;
  private credential = new ChainedTokenCredential(
    new ManagedIdentityCredential(),
    new AzureCliCredential()
  );

  /**
   * POSTs the feedback payload. Returns true on HTTP 2xx, false otherwise.
   * Never throws — callers want errors to be invisible to the user.
   */
  async send(payload: FeedbackPayload): Promise<boolean> {
    try {
      const url = this.buildUrl();
      const key = await this.getFunctionKey();

      const headers: Record<string, string> = {
        "Content-Type": "application/json",
      };
      if (key) {
        headers["x-functions-key"] = key;
      }

      const res = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify(payload),
        // Match the Python client's 15s timeout via AbortSignal.
        signal: AbortSignal.timeout(15_000),
      });

      if (!res.ok) {
        const text = await res.text().catch(() => "<no body>");
        console.error(
          `[feedbackClient] ✗ feedback HTTP ${res.status}: ${text.slice(0, 500)}`
        );
        return false;
      }

      return true;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error(`[feedbackClient] ✗ feedback POST failed: ${msg}`);
      return false;
    }
  }

  /** Replace trailing `/orc` (case-insensitive) with `/feedback`. */
  private buildUrl(): string {
    return config.orchestratorEndpoint.replace(/\/orc\/?$/i, "/feedback");
  }

  private async getFunctionKey(): Promise<string> {
    if (config.feedbackFunctionKey) {
      return config.feedbackFunctionKey;
    }
    if (this.cachedKey) {
      return this.cachedKey;
    }

    if (
      !config.azureSubscriptionId ||
      !config.azureResourceGroup ||
      !config.azureOrchestratorFuncName
    ) {
      // No static key and no MI lookup info — assume anonymous endpoint.
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
      `/functions/feedback/listKeys?api-version=2022-03-01`;

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
