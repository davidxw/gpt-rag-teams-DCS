import { App } from "@microsoft/teams.apps";
import { MessageActivity } from "@microsoft/teams.api";
import { OrchestratorClient } from "./orchestratorClient";

/**
 * Per-Teams-conversation map of the orchestrator-issued `conversation_id`.
 *
 * The orchestrator (CosmosDB-backed) is the source of truth for the actual
 * conversation history; we only need to remember the id so subsequent
 * turns continue the same server-side conversation.
 *
 * NOTE: This is in-process memory only — same volatility as the previous
 * `MemoryStorage`-backed implementation. Conversations restart after a
 * pod/app restart. Move to a durable store (Azure Tables / Cosmos / Redis)
 * if longer continuity is required.
 */
const orchestratorConvByTeamsConv = new Map<string, string>();
const orchestrator = new OrchestratorClient();

export const app = new App();

app.on("message", async ({ activity, send }) => {
  const text = (activity.text ?? "").trim();
  if (!text) {
    return;
  }

  const teamsConvId = activity.conversation?.id ?? "";
  const userId = activity.from?.aadObjectId ?? activity.from?.id ?? "";
  const userName = activity.from?.name ?? "";

  // Show the typing indicator while the orchestrator runs.
  // `send` accepts any IActivity-shaped payload; typing has no payload.
  await send({ type: "typing" } as any);

  try {
    const result = await orchestrator.ask({
      conversation_id: orchestratorConvByTeamsConv.get(teamsConvId),
      question: text,
      client_principal_id: userId,
      client_principal_name: userName,
      client_group_names: "",
    });

    console.log("[teamsBot] Orchestrator result:", JSON.stringify(result));

    // Persist the orchestrator-issued conversation id so subsequent turns
    // continue the same conversation in CosmosDB.
    if (teamsConvId && result.conversation_id) {
      orchestratorConvByTeamsConv.set(teamsConvId, result.conversation_id);
    }

    const answerText = result.answer || "_(no answer returned)_";
    console.log("[teamsBot] Sending answer:", answerText.slice(0, 200));

    // Send a plain string first to verify the round-trip works,
    // then we can re-introduce MessageActivity + addAiGenerated().
    await send(answerText);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[teamsBot] Orchestrator call failed:", msg);
    await send(
      "Sorry — I couldn't reach the GPT-RAG orchestrator. Please try again in a moment."
    );
  }
});

app.event("error", ({ error }) => {
  console.error("[teamsBot] Unhandled error:", error);
});

