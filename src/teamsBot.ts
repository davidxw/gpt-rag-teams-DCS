import { App } from "@microsoft/teams.apps";
import { MessageActivity } from "@microsoft/teams.api";
import { OrchestratorClient } from "./orchestratorClient";
import {
  buildCitationAppearance,
  coerceDataPoints,
  rewriteAnswerWithCitations,
} from "./citations";

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
  const startedAt = Date.now();
  const text = (activity.text ?? "").trim();
  const teamsConvId = activity.conversation?.id ?? "";
  const userId = activity.from?.aadObjectId ?? activity.from?.id ?? "";
  const userName = activity.from?.name ?? "";
  const existingOrchConvId = orchestratorConvByTeamsConv.get(teamsConvId);

  console.log(
    `[teamsBot] ← message received | user='${userName}' (${userId}) | ` +
      `teamsConv=${teamsConvId} | orchestratorConv=${existingOrchConvId ?? "(new)"} | ` +
      `chars=${text.length}`
  );
  console.log(`[teamsBot]   text: ${text.slice(0, 200)}${text.length > 200 ? "…" : ""}`);

  if (!text) {
    console.log("[teamsBot] empty message body, ignoring.");
    return;
  }

  // Show the typing indicator while the orchestrator runs.
  // `send` accepts any IActivity-shaped payload; typing has no payload.
  await send({ type: "typing" } as any);

  try {
    console.log(`[teamsBot] → orchestrator.ask (conv=${existingOrchConvId ?? "(new)"})`);
    const orchStart = Date.now();
    const result = await orchestrator.ask({
      conversation_id: existingOrchConvId,
      question: text,
      client_principal_id: userId,
      client_principal_name: userName,
      client_group_names: "",
    });
    console.log(
      `[teamsBot] ← orchestrator.ask (${Date.now() - orchStart}ms) | ` +
        `conv=${result.conversation_id} | answer_chars=${(result.answer ?? "").length} | ` +
        `data_points=${Array.isArray(result.data_points) ? result.data_points.length : 0}`
    );

    // Persist the orchestrator-issued conversation id so subsequent turns
    // continue the same conversation in CosmosDB.
    if (teamsConvId && result.conversation_id) {
      orchestratorConvByTeamsConv.set(teamsConvId, result.conversation_id);
    }

    const answerText = result.answer || "_(no answer returned)_";
    console.log(
      `[teamsBot]   raw answer (first 500 chars): ${answerText.slice(0, 500).replace(/\n/g, "\\n")}`
    );

    // Convert orchestrator `[file][PageN][title]` markers into Teams-native
    // citations: rewrite the inline markers to `[N]` and attach a
    // `CitationAppearance` per unique reference. The appearance carries a
    // GET URL (Option A, via DOCUMENT_URL_TEMPLATE) and a modal Adaptive
    // Card with the source excerpt (Option B, dev preview).
    const { text: rewritten, citations } =
      rewriteAnswerWithCitations(answerText);
    const dataPoints = coerceDataPoints(result.data_points);

    console.log(
      `[teamsBot] → sending reply | citations=${citations.length} | chars=${rewritten.length}`
    );
    if (citations.length === 0) {
      console.warn(
        `[teamsBot]   ⚠ no citations parsed — orchestrator markers may not match the [file][PageN] regex; check raw answer above.`
      );
    }

    // NOTE: Do NOT set `textFormat: 'markdown'` here. Teams' default
    // renderer already interprets markdown (bold, lists, headers) AND
    // overlays the `[N]` citation chips. Explicitly forcing
    // `textFormat: 'markdown'` switches Teams to a markdown-only path
    // that strips citation interactivity (the brackets render as plain
    // text). See:
    // https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bot-messages-ai-generated-content
    const message = new MessageActivity(
      rewritten || "_(no answer returned)_"
    ).addAiGenerated();

    for (const citation of citations) {
      message.addCitation(
        citation.position,
        buildCitationAppearance(citation, dataPoints)
      );
    }

    await send(message);
    console.log(
      `[teamsBot] ✓ reply sent | total=${Date.now() - startedAt}ms | teamsConv=${teamsConvId}`
    );
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error(
      `[teamsBot] ✗ orchestrator call failed (${Date.now() - startedAt}ms):`,
      msg
    );
    await send(
      "Sorry — I couldn't reach the GPT-RAG orchestrator. Please try again in a moment."
    );
  }
});

app.event("error", ({ error }) => {
  console.error("[teamsBot] Unhandled error:", error);
});

