import express, { Request, Response } from "express";
import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth.js";
import axios from "axios";
import crypto from "crypto";
import { v4 as uuidv4 } from "uuid";
import dotenv from "dotenv";
import { OfficeParser } from "officeparser";
import { createRequire } from "module";
import { pathToFileURL } from "url";
import fs from "fs";

dotenv.config();
const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;
app.post("/webhook", (req: Request, res: Response) => {
  if (req.query && req.query.validationToken) {
    console.log("Validation request received. Token:", req.query.validationToken);
    res.setHeader("Content-Type", "text/plain");
    return res.status(200).send(req.query.validationToken);
  }
  const { value } = req.body;
  if (value && Array.isArray(value)) {
    for (const notification of value) {
      if (notification.clientState !== process.env.CLIENT_STATE) {
        console.warn("Invalid clientState. Ignoring notification.");
        continue;
      }
      console.log(`Processing change for resource: ${notification.resource}`);
      handleSharePointChange(notification);
    }
  }
  res.status(202).send();
});
function generateJWT() {
  const header = { alg: "HS256", typ: "JWT" };
  const eat = Math.floor(Date.now() / 1000) + 60;
  const payload = { eat: eat };
  const secret = "ee133968565b34a46070aaec193b18760f0d516b258cdd888486f492576e1954699f70169bf28af4c39bc24559ca3e90fc9c98383d7d24ecfbbefa98e0926662";
  const base64urlEncode = (obj: any) => {
    return Buffer.from(JSON.stringify(obj))
      .toString("base64")
      .replace(/=/g, "")
      .replace(/\+/g, "-")
      .replace(/\//g, "_");
  };
  const encodedHeader = base64urlEncode(header);
  const encodedPayload = base64urlEncode(payload);
  const unsignedToken = `${encodedHeader}.${encodedPayload}`;
  const signature = crypto
    .createHmac("sha256", secret)
    .update(unsignedToken)
    .digest("base64")
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");

  return `${unsignedToken}.${signature}`;
}
const processedItems = new Set<string>();

// --- Processed Items Persistence ---
// Stores item IDs of documents we've already sent to the Embedding API,
// keyed by item.id with their lastModifiedDateTime. This prevents duplicate
// embeddings when SharePoint replays the same item across delta calls or
// when the server restarts mid-window.
const PROCESSED_ITEMS_FILE = "processed_items.json";

function loadProcessedItems(): Map<string, string> {
  try {
    if (fs.existsSync(PROCESSED_ITEMS_FILE)) {
      const data = JSON.parse(fs.readFileSync(PROCESSED_ITEMS_FILE, "utf-8"));
      const map = new Map<string, string>(Object.entries(data));
      console.log(`📌 Loaded ${map.size} previously processed item(s).`);
      return map;
    }
  } catch {
    console.warn("⚠️ Could not load processed items file, starting fresh.");
  }
  return new Map<string, string>();
}

function saveProcessedItems(map: Map<string, string>) {
  try {
    const obj = Object.fromEntries(map);
    fs.writeFileSync(PROCESSED_ITEMS_FILE, JSON.stringify(obj, null, 2));
  } catch (err) {
    console.error("❌ Failed to save processed items:", err);
  }
}

const persistedProcessedItems: Map<string, string> = loadProcessedItems();

// Server startup timestamp — only files created AFTER this time will be processed.
// This prevents pre-existing files from being re-synced on every server restart.
const SERVER_START_TIME = new Date();
console.log(`🕒 Server start time: ${SERVER_START_TIME.toISOString()} — only files uploaded after this will be synced.`);

// --- Delta Link Persistence ---
// Stores the last deltaLink so we only get NEW changes after a restart,
// instead of replaying the entire folder history (which includes deleted items).
const DELTA_LINK_FILE = "delta_link.json";

function loadDeltaLink(): string | null {
  try {
    if (fs.existsSync(DELTA_LINK_FILE)) {
      const data = JSON.parse(fs.readFileSync(DELTA_LINK_FILE, "utf-8"));
      console.log("📌 Loaded saved deltaLink — will only fetch new changes.");
      return data.deltaLink || null;
    }
  } catch {
    console.warn("⚠️ Could not load delta link file, starting fresh.");
  }
  return null;
}

function saveDeltaLink(deltaLink: string) {
  try {
    fs.writeFileSync(DELTA_LINK_FILE, JSON.stringify({ deltaLink, savedAt: new Date().toISOString() }));
    console.log("💾 Saved deltaLink for next sync.");
  } catch (err) {
    console.error("❌ Failed to save deltaLink:", err);
  }
}

let cachedDeltaLink: string | null = loadDeltaLink();

// Prevents concurrent sync runs (webhook + polling race condition)
let isSyncing = false;

async function handleSharePointChange(notification: any) {
  if (isSyncing) {
    console.log("⏳ Sync already in progress, skipping this trigger.");
    return;
  }
  isSyncing = true;
  try {
    await _handleSharePointChange(notification);
  } finally {
    isSyncing = false;
  }
}

async function _handleSharePointChange(notification: any) {
  const token = await getAccessToken();
  if (!token) return;
  const client = Client.init({
    authProvider: (done) => {
      done(null, token);
    },
  });
  try {
    console.log("Fetching specific changes from SharePoint...");
    const driveId = process.env.SHAREPOINT_DRIVE_ID;
    const folderId = process.env.SHAREPOINT_FOLDER_ID;

    // Use the saved deltaLink if available, otherwise start a fresh delta query
    let url: string;
    if (cachedDeltaLink) {
      url = cachedDeltaLink;
      console.log("🔄 Using saved deltaLink to fetch only NEW changes...");
    } else {
      url = `/drives/${driveId}/items/${folderId}/delta`;
      console.log("🆕 No deltaLink saved — fetching full delta (first run)...");
    }

    // Collect all changes across paginated responses
    let allChanges: any[] = [];
    let nextLink: string | null = url;
    let newDeltaLink: string | null = null;

    while (nextLink) {
      const response = cachedDeltaLink && nextLink === url
        ? await axios.get(nextLink, { headers: { Authorization: `Bearer ${token}` } }).then(r => r.data)
        : await client.api(nextLink).get();

      if (response.value) {
        allChanges = allChanges.concat(response.value);
      }

      // Follow pagination
      if (response["@odata.nextLink"]) {
        nextLink = response["@odata.nextLink"];
        cachedDeltaLink = null; // Reset so next iteration uses client.api
      } else {
        newDeltaLink = response["@odata.deltaLink"] || null;
        nextLink = null;
      }
    }

    // Save the new deltaLink so next time we only get changes after this point
    if (newDeltaLink) {
      cachedDeltaLink = newDeltaLink;
      saveDeltaLink(newDeltaLink);
    }

    if (allChanges.length > 0) {
      console.log(`📊 Delta returned ${allChanges.length} item(s). Filtering...`);
      for (const item of allChanges) {

        // ✅ SKIP DELETED ITEMS — the delta API returns items with a "deleted"
        // facet when they've been removed from the folder.
        if (item.deleted) {
          console.log(`🗑️ Skipping deleted item: "${item.name || item.id}"`);
          continue;
        }

        const createdDate = new Date(item.createdDateTime);
        if (createdDate > SERVER_START_TIME && item.file) {
          if (processedItems.has(item.id)) {
            continue;
          }
          // Skip if we've already processed this exact version (same lastModifiedDateTime)
          const lastModified = item.lastModifiedDateTime || "";
          const previouslyProcessed = persistedProcessedItems.get(item.id);
          if (previouslyProcessed && previouslyProcessed >= lastModified) {
            console.log(`⏭️  Skipping already-processed file: "${item.name}" (modified: ${lastModified})`);
            processedItems.add(item.id);
            continue;
          }
          processedItems.add(item.id);
          console.log(`✨ NEW DOCUMENT DETECTED: "${item.name}"`);
          try {
            const fileId = uuidv4();
            const fileUrl = item["@microsoft.graph.downloadUrl"] || item.webUrl;
            console.log(`🆔 Generated File ID: ${fileId}`);
            console.log(`🔗 Using SharePoint URL: ${fileUrl}`);

            // Extract document text using officeparser
            let parsedText = `Document traced from SharePoint: ${item.name}`;
            try {
              let buffer: Buffer | null = null;
              const downloadUrl = item["@microsoft.graph.downloadUrl"];

              if (downloadUrl) {
                console.log(`📥 Downloading document via downloadUrl for: ${item.name}`);
                const response = await axios.get(downloadUrl, { responseType: 'arraybuffer' });
                buffer = Buffer.from(response.data);
              } else {
                console.log(`📥 Downloading document content directly via Graph API for: ${item.name}`);
                const contentUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}/content`;
                const response = await axios.get(contentUrl, {
                  headers: { Authorization: `Bearer ${token}` },
                  responseType: 'arraybuffer'
                });
                buffer = Buffer.from(response.data);
              }

              if (buffer) {
                console.log(`📄 Parsing document content...`);

                // Workaround for pdfjs-dist ESM loader issue on Windows
                // Converts the absolute C:/ path to a proper file:// URL
                const require = createRequire(import.meta.url);
                const workerPath = require.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs');
                const pdfWorkerSrc = pathToFileURL(workerPath).href;

                const startTime = Date.now();
                // Using officeparser to extract text from the file buffer
                const ast = await OfficeParser.parseOffice(buffer, {
                  ocr: true,
                  extractAttachments: true,
                  outputErrorToConsole: true,
                  pdfWorkerSrc: pdfWorkerSrc,
                  ocrConfig: {
                    langPath: process.cwd()
                  }
                });
                const extractedText = ast.toText();
                const durationMs = Date.now() - startTime;

                const totalSec = Math.floor(durationMs / 1000);
                const mins = Math.floor(totalSec / 60);
                const secs = totalSec % 60;
                const timeString = mins > 0 ? `${mins}m ${secs}s` : `${(durationMs / 1000).toFixed(1)}s`;

                if (extractedText && extractedText.trim().length > 0) {
                  parsedText = extractedText;
                  console.log(`✅ Document parsed successfully in ${timeString}. Extracted ${parsedText.length} characters.`);
                } else {
                  console.warn(`⚠️ No text extracted from document. Using default string.`);
                }
              }
            } catch (parseErr: any) {
              console.error(`❌ Failed to parse document ${item.name}:`, parseErr.response?.data || parseErr.message);
            }

            console.log(`\n================ PARSED TEXT FOR ${item.name} ================`);
            console.log(parsedText);
            console.log(`====================================================================\n`);

            // Save the extracted text to a local file so the user can easily inspect it
            try {
              const logFileName = `parsed_${item.name}.txt`;
              fs.writeFileSync(logFileName, parsedText);
              console.log(`💾 Saved entire parsed text to ${logFileName} for easy inspection.`);
            } catch (fsErr) {
              console.error("Could not save parsed text to file:", fsErr);
            }

            // 1. Prepare the payload for the Embedding API
            const embeddingPayload = {
              metadata: {
                category: "integration",
                sub_type: item.parentReference.name || "General",
                fileUrl: fileUrl,
                fileId: fileId,
                file_name: item.name,
                fileType: item.file.mimeType || "application/pdf"
              },
              content: parsedText,
              collectionName: "6943a72296416fb578171fca",
              asyncProcessing: true,
              chunkSize: 2000000,
              overlap: 50000,
              returnEmbedding: false
            };
            // 2. Trigger the Embedding API
            console.log(`🚀 Triggering Embedding API for: ${item.name}...`);
            const jwt = generateJWT();
            const apiResponse = await axios.post("https://api-dev-ai.vithiit.com/generate-embedding", embeddingPayload, {
              headers: {
                "x-api-key": "5ba266fb7bee9a60",
                "Authorization": `Bearer ${jwt}`,
                "Content-Type": "application/json"
              }
            });
            console.log(`✅ Embedding API Success for ${item.name}:\n`, JSON.stringify(apiResponse.data, null, 2));

            // Persist this item so we don't re-process it across server restarts
            persistedProcessedItems.set(item.id, item.lastModifiedDateTime || new Date().toISOString());
            saveProcessedItems(persistedProcessedItems);

            // If the response contains chunks, they will now be fully printed in the console
            // so you can confirm if the image content was extracted correctly.
          } catch (apiError: any) {
            console.error(`❌ Failed to process ${item.name}:`, apiError.response?.data || apiError.message);
          }
        }
      }
    } else {
      console.log("✅ No new changes detected.");
    }
  } catch (error: any) {
    console.error("Error tracing changes:", error.body || error);
  }
}
export function startServer() {
  app.listen(PORT, () => {
    console.log(`Webhook server listening on port ${PORT}`);
    const POLLING_INTERVAL_MS = 2 * 60 * 1000;
    console.log(`⏱️ Automatic polling enabled. Checking SharePoint every ${POLLING_INTERVAL_MS / 1000} seconds...`);

    setInterval(async () => {
      console.log("🔄 [Polling] Checking for new SharePoint documents...");
      await handleSharePointChange({ resource: "automatic-polling" });
    }, POLLING_INTERVAL_MS);
  });
}
