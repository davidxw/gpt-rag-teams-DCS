import { App } from "@microsoft/teams.apps";
import { MessageActivity, ActivityLike } from "@microsoft/teams.api";
import { OrchestratorClient } from "./orchestratorClient";
import { FeedbackClient } from "./feedbackClient";
import { config } from "./config";
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

/**
 * 0-based counter of bot replies per Teams conversation. Used to populate
 * `message_index` on outbound feedback payloads.
 */
const turnIndexByTeamsConv = new Map<string, number>();

/**
 * Per-bot-message metadata captured at send time. Keyed by the
 * `SentActivity.id` returned from `send(...)` so the feedback handler can
 * recover the orchestrator conversation id, the original question, the
 * answer, and the turn index when the user clicks 👍 / 👎 on a specific
 * reply.
 *
 * Same in-process volatility caveat as `orchestratorConvByTeamsConv`.
 */
type SentTurnMeta = {
  orchestratorConvId: string;
  question: string;
  answer: string;
  messageIndex: number;
  clientPrincipalId: string;
  clientPrincipalName: string;
};
const turnMetaByBotMessageId = new Map<string, SentTurnMeta>();

/**
 * Most recent bot message id per Teams conversation. Used as a fallback
 * when an inbound feedback invoke arrives with no `replyToId` (e.g. in
 * Microsoft 365 Agents Playground, which doesn't always populate it).
 */
const latestBotMessageIdByTeamsConv = new Map<string, string>();

const orchestrator = new OrchestratorClient();
const feedbackClient = new FeedbackClient();

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

  // Show the typing indicator while the orchestrator runs. Teams expires
  // a typing indicator after ~10s, so we resend it every 4s until the
  // reply is on its way. `send` accepts any IActivity-shaped payload;
  // typing has no payload.
  await send({ type: "typing" } as any);
  const typingInterval = setInterval(() => {
    send({ type: "typing" } as any).catch(() => {
      /* swallow transient channel errors — typing is best-effort */
    });
  }, 4000);

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
    if (config.rawAnswerLogChars > 0) {
      console.log(
        `[teamsBot]   raw answer (first ${config.rawAnswerLogChars} chars): ` +
          answerText.slice(0, config.rawAnswerLogChars).replace(/\n/g, "\\n")
      );
    }

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
    ).addAiGenerated()
    .addFeedback();

    for (const citation of citations) {
      message.addCitation(
        citation.position,
        buildCitationAppearance(citation, dataPoints)
      );
    }

    const sent = await send(message);

    // Cache per-turn metadata so the feedback handler can build a full
    // payload (conversation_id, question, answer, message_index, …) from
    // just the `replyToId` on the inbound feedback invoke.
    if (sent?.id) {
      const messageIndex = (turnIndexByTeamsConv.get(teamsConvId) ?? -1) + 1;
      turnIndexByTeamsConv.set(teamsConvId, messageIndex);
      turnMetaByBotMessageId.set(sent.id, {
        orchestratorConvId: result.conversation_id,
        question: text,
        answer: answerText,
        messageIndex,
        clientPrincipalId: userId,
        clientPrincipalName: userName,
      });
      if (teamsConvId) {
        latestBotMessageIdByTeamsConv.set(teamsConvId, sent.id);
      }
    } else {
      console.warn(
        `[teamsBot]   ⚠ send() returned no id — feedback for this reply will not be linkable to the orchestrator turn.`
      );
    }

    console.log(
      `[teamsBot] ✓ reply sent | total=${Date.now() - startedAt}ms | teamsConv=${teamsConvId} | botMsgId=${sent?.id ?? "(none)"}`
    );
  } catch (err) {
    // Native `fetch` wraps the underlying network error in a generic
    // `TypeError: fetch failed`, putting the real reason on `.cause`
    // (and sometimes nested `.cause.cause` for DNS / TLS issues).
    // Surface as much detail as we can to the log.
    const parts: string[] = [];
    if (err instanceof Error) {
      parts.push(`${err.name}: ${err.message}`);
      let cause: unknown = (err as { cause?: unknown }).cause;
      while (cause) {
        if (cause instanceof Error) {
          const code = (cause as { code?: string }).code;
          parts.push(
            `caused by ${cause.name}${code ? ` [${code}]` : ""}: ${cause.message}`
          );
          cause = (cause as { cause?: unknown }).cause;
        } else {
          parts.push(`caused by ${String(cause)}`);
          cause = undefined;
        }
      }
    } else {
      parts.push(String(err));
    }
    console.error(
      `[teamsBot] ✗ orchestrator call failed (${Date.now() - startedAt}ms) | ` +
        `endpoint=${config.orchestratorEndpoint} | conv=${existingOrchConvId ?? "(new)"}: ` +
        parts.join(" | ")
    );
    if (err instanceof Error && err.stack) {
      console.error(`[teamsBot]   stack: ${err.stack}`);
    }
    await send(
      "Sorry — I couldn't reach the GPT-RAG orchestrator. Please try again in a moment."
    );
  } finally {
    clearInterval(typingInterval);
  }
});

app.on("message.submit.feedback", async ({ activity, log }) => {
  const { reaction, feedback: feedbackJson } = activity.value.actionValue;
  const teamsConvId = activity.conversation?.id ?? "";

  // Diagnostic: log the keys present on the inbound activity so we can see
  // what the channel actually populated (replyToId, relatesTo, etc.).
  log.info(
    `Feedback invoke received | id=${activity.id} | replyToId=${activity.replyToId ?? "(null)"} | ` +
      `teamsConv=${teamsConvId} | relatesToId=${(activity as any).relatesTo?.activityId ?? "(none)"}`
  );

  // Resolve the bot's outgoing message id this feedback belongs to.
  // Real Teams populates `replyToId`; Microsoft 365 Agents Playground
  // (and some other harnesses) do not — fall back to the most recent bot
  // reply in the same Teams conversation.
  let botMessageId: string | undefined =
    activity.replyToId ??
    (activity as any).relatesTo?.activityId ??
    undefined;
  let resolution: "replyToId" | "latest-fallback" = "replyToId";
  if (!botMessageId && teamsConvId) {
    botMessageId = latestBotMessageIdByTeamsConv.get(teamsConvId);
    resolution = "latest-fallback";
  }

  if (!botMessageId) {
    log.warn(
      `Could not resolve bot message id for feedback (no replyToId and no recent reply for teamsConv=${teamsConvId}). Dropping feedback.`
    );
    return;
  }

  const meta = turnMetaByBotMessageId.get(botMessageId);
  if (!meta) {
    log.warn(
      `No cached turn metadata for botMessageId=${botMessageId} ` +
        `(resolution=${resolution}; likely sent before the most recent app restart). Dropping feedback.`
    );
    return;
  }

  // The `feedback` field may arrive as:
  //   • an already-parsed object (real Teams / Bot Framework SDK), e.g.
  //     `{ feedbackText: "Nice!" }`
  //   • a JSON string envelope, e.g. `'{"feedbackText":"Nice!"}'`
  //   • a raw plain string (Microsoft 365 Agents Playground sends just
  //     the typed text)
  // Normalize all three into a plain string for the feedback backend,
  // which strictly requires `comment` to be a string.
  let comment = "";
  if (feedbackJson != null) {
    if (typeof feedbackJson === "object") {
      const obj = feedbackJson as { feedbackText?: unknown };
      comment = typeof obj.feedbackText === "string" ? obj.feedbackText : "";
    } else if (typeof feedbackJson === "string" && feedbackJson.length > 0) {
      try {
        const parsed = JSON.parse(feedbackJson);
        if (parsed && typeof parsed === "object") {
          comment =
            typeof parsed.feedbackText === "string" ? parsed.feedbackText : "";
        } else if (typeof parsed === "string") {
          // JSON-encoded plain string, e.g. '"hello"'.
          comment = parsed;
        } else {
          comment = feedbackJson;
        }
      } catch {
        // Not JSON — assume it's the raw comment text (Playground behavior).
        comment = feedbackJson;
      }
    }
  }

  // Teams sends `like`/`dislike`; the feedback backend expects `up`/`down`.
  const rating =
    reaction === "like" ? "up" : reaction === "dislike" ? "down" : reaction;

  const payload = {
    conversation_id: meta.orchestratorConvId,
    question: meta.question,
    answer: meta.answer,
    rating,
    comment,
    message_index: meta.messageIndex,
    client_principal_id: meta.clientPrincipalId,
    client_principal_name: meta.clientPrincipalName,
    source: "teams",
    // Bot doesn't currently resolve the user's groups; send empty array
    // for parity with the feedback contract.
    client_group_names: [] as string[],
  };

  log.info(
    `Resolved feedback | teamsConv=${teamsConvId} | botMsgId=${botMessageId} (via ${resolution}) | ` +
      `orchestratorConv=${meta.orchestratorConvId} | messageIndex=${meta.messageIndex} | ` +
      `rating=${rating} (raw=${reaction}) | commentChars=${comment.length}`
  );

  // Fire-and-forget POST to the `feedback` function. The Teams invoke
  // response must complete promptly and any failure must be invisible to
  // the user — FeedbackClient.send() never throws and logs internally.
  void feedbackClient.send(payload).then((ok) => {
    if (ok) {
      log.info(
        `Feedback POSTed OK | botMsgId=${botMessageId} | orchestratorConv=${meta.orchestratorConvId}`
      );
    }
  });
});

app.event("error", ({ error }) => {
  console.error("[teamsBot] Unhandled error:", error);
});

