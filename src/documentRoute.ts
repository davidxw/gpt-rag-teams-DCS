import type { Request, Response } from "express";
import {
  BlobServiceClient,
  type BlobDownloadResponseParsed,
} from "@azure/storage-blob";
import {
  AzureCliCredential,
  ChainedTokenCredential,
  ManagedIdentityCredential,
  type TokenCredential,
} from "@azure/identity";
import type { App } from "@microsoft/teams.apps";
import { config } from "./config";

/**
 * Registers `GET /api/documents?name=<blobPath>` on the Teams SDK App's
 * underlying Express adapter. The route streams the blob bytes from the
 * configured storage account using DefaultAzureCredential — no keys, no
 * SAS. The citation URLs in chat point at this endpoint so storage stays
 * private.
 *
 * Required env: `STORAGE_ACCOUNT`
 * Required RBAC: `Storage Blob Data Reader` for the bot's identity on
 * either the storage account or the target container.
 *
 * SECURITY: this endpoint is currently unauthenticated. Anyone with the
 * URL can fetch any blob in the configured container. Layer App Service
 * Easy Auth or HMAC-signed URLs on top before exposing publicly.
 */
export function registerDocumentRoute(app: App): void {
  if (!config.storageAccount) {
    console.log(
      "[documentRoute] STORAGE_ACCOUNT not set; skipping /api/documents route."
    );
    return;
  }

  const credential: TokenCredential = new ChainedTokenCredential(
    new ManagedIdentityCredential(
      config.clientId ? { clientId: config.clientId } : undefined
    ),
    new AzureCliCredential()
  );

  const serviceClient = new BlobServiceClient(
    `https://${config.storageAccount}.blob.core.windows.net`,
    credential
  );
  const containerClient = serviceClient.getContainerClient(
    config.storageContainer
  );

  // The Teams SDK App exposes the underlying adapter at `app.server.adapter`.
  // The default adapter is `ExpressAdapter`, which proxies through Express's
  // `get`/`post`/etc. methods. We type-narrow with a runtime check.
  const adapter = app.server.adapter as unknown as {
    get?: (path: string, handler: (req: Request, res: Response) => void) => void;
  };
  if (typeof adapter.get !== "function") {
    console.error(
      "[documentRoute] HTTP adapter does not expose Express `.get()`; " +
        "cannot register document route."
    );
    return;
  }

  adapter.get("/api/documents", async (req: Request, res: Response) => {
    const rawName = (req.query.name ?? "") as string;
    const name = rawName.trim();

    // Path-traversal & shape guards.
    if (
      !name ||
      name.length > 1024 ||
      name.includes("..") ||
      name.startsWith("/") ||
      name.startsWith("\\")
    ) {
      res.status(400).type("text/plain").send("Invalid blob name.");
      return;
    }

    try {
      const blobClient = containerClient.getBlobClient(name);
      const exists = await blobClient.exists();
      if (!exists) {
        res.status(404).type("text/plain").send("Document not found.");
        return;
      }

      const range = req.headers.range;
      const download: BlobDownloadResponseParsed = range
        ? await blobClient.download(0, undefined, {
            rangeGetContentMD5: false,
            // Pass the raw Range header through so the SDK honors it.
            // (storage-blob v12 accepts a custom range via the second-arg
            // overload; for simplicity we just full-download here and
            // delegate range support to a future enhancement.)
          })
        : await blobClient.download();

      res.status(download._response.status === 206 ? 206 : 200);
      if (download.contentType) {
        res.setHeader("Content-Type", download.contentType);
      }
      if (download.contentLength != null) {
        res.setHeader("Content-Length", String(download.contentLength));
      }
      // Inline so PDFs render in-browser; Teams citation modal opens externally.
      res.setHeader(
        "Content-Disposition",
        `inline; filename="${encodeURIComponent(basename(name))}"`
      );
      res.setHeader("Cache-Control", "private, max-age=300");

      const stream = download.readableStreamBody;
      if (!stream) {
        res.status(502).type("text/plain").send("Empty response from storage.");
        return;
      }
      stream.on("error", (err) => {
        console.error("[documentRoute] Stream error:", err);
        if (!res.headersSent) {
          res.status(502).type("text/plain").send("Failed to stream blob.");
        } else {
          res.destroy(err);
        }
      });
      stream.pipe(res);
    } catch (err: unknown) {
      const e = err as { statusCode?: number; code?: string; message?: string };
      const status =
        Number(e?.statusCode) === 403
          ? 403
          : Number(e?.statusCode) === 404
            ? 404
            : 500;
      console.error(`[documentRoute] Error fetching '${name}':`, err);
      if (!res.headersSent) {
        // Surface the underlying cause to the caller — this endpoint is only
        // reachable internally today (see SECURITY note above); when Easy
        // Auth / HMAC are added, consider trimming to a generic message in
        // production.
        const detail = [e?.code, e?.message].filter(Boolean).join(": ");
        res
          .status(status)
          .type("text/plain")
          .send(`Failed to fetch document (${status})${detail ? `: ${detail}` : ""}`);
      }
    }
  });

  console.log(
    `[documentRoute] GET /api/documents → ${config.storageAccount}/${config.storageContainer}`
  );
}

function basename(p: string): string {
  const i = Math.max(p.lastIndexOf("/"), p.lastIndexOf("\\"));
  return i >= 0 ? p.slice(i + 1) : p;
}
