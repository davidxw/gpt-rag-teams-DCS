// Loading config first ensures the env-var bridge in ./config runs before
// the Teams SDK App tries to read CLIENT_ID / CLIENT_SECRET / TENANT_ID.
import { config } from "./config";
import { app } from "./teamsBot";

async function main(): Promise<void> {
  // The Teams SDK App from `@microsoft/teams.apps` hosts its own HTTP
  // server and exposes the bot messaging endpoint at POST /api/messages
  // automatically. This is what the Azure Bot Service registration's
  // `messagingEndpoint` must point at, e.g.
  //   https://<your-app-service>.azurewebsites.net/api/messages
  await app.start(config.port);
  console.log(
    `gpt-rag-teams listening on http://localhost:${config.port} (POST /api/messages)`
  );
}

main().catch((err) => {
  console.error("[index] Fatal startup error:", err);
  process.exit(1);
});

