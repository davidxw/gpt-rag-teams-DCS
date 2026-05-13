import * as restify from "restify";
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  ConfigurationServiceClientCredentialFactory,
} from "botbuilder";

import { app } from "./teamsBot";
import { config } from "./config";

const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppType: config.botType,
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppTenantId: config.botTenantId,
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error("[adapter] Turn error:", error);
  await context.sendActivity("The bot encountered an error.");
};

const server = restify.createServer({ name: "gpt-rag-teams" });
server.use(restify.plugins.bodyParser());

// Liveness/readiness probe (used by App Service health checks).
server.get("/healthz", (_req, res, next) => {
  res.send(200, { status: "ok" });
  return next();
});

// Bot Framework messaging endpoint. This is what the Azure Bot Service
// registration's `messagingEndpoint` must point at, e.g.
//   https://<your-app-service>.azurewebsites.net/api/messages
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await app.run(context);
  });
});

server.listen(config.port, () => {
  console.log(`gpt-rag-teams listening on http://localhost:${config.port}`);
  console.log(`Messaging endpoint: POST /api/messages`);
});
