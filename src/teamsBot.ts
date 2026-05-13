import { MemoryStorage, TurnContext } from "botbuilder";
import { Application, TurnState } from "@microsoft/teams-ai";
import { OrchestratorClient } from "./orchestratorClient";

interface ConversationData {
  /** UUID returned by the orchestrator on the first turn; reused thereafter. */
  orchestratorConversationId?: string;
}

type AppTurnState = TurnState<ConversationData>;

const storage = new MemoryStorage();
const orchestrator = new OrchestratorClient();

export const app = new Application<AppTurnState>({
  storage,
});

app.activity("message", async (context: TurnContext, state: AppTurnState) => {
  const text = (context.activity.text ?? "").trim();
  if (!text) {
    return;
  }

  const userId =
    context.activity.from?.aadObjectId ?? context.activity.from?.id ?? "";
  const userName = context.activity.from?.name ?? "";

  // Show the typing indicator while the orchestrator runs.
  await context.sendActivities([{ type: "typing" }]);

  try {
    const result = await orchestrator.ask({
      conversation_id: state.conversation.orchestratorConversationId,
      question: text,
      client_principal_id: userId,
      client_principal_name: userName,
      client_group_names: "",
    });

    // Persist the orchestrator-issued conversation id so subsequent turns
    // continue the same conversation in CosmosDB.
    state.conversation.orchestratorConversationId = result.conversation_id;

    await context.sendActivity(result.answer || "_(no answer returned)_");
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error("[teamsBot] Orchestrator call failed:", msg);
    await context.sendActivity(
      "Sorry — I couldn't reach the GPT-RAG orchestrator. Please try again in a moment."
    );
  }
});

app.error(async (context, error) => {
  console.error("[teamsBot] Unhandled error:", error);
  await context.sendActivity("Something went wrong. Please try again.");
});
