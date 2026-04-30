import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth.js";
import dotenv from "dotenv";
import fs from "fs";
dotenv.config();

interface FolderInfo {
  name: string;
  id: string;
  webUrl: string;
  itemCount: number;
  children: FolderInfo[];
}

/**
 * Recursively walk a folder and return its full subtree (1 level deep by default).
 */
async function walkFolder(
  client: Client,
  driveId: string,
  folderId: string,
  folderName: string,
  webUrl: string,
  depth: number,
  maxDepth: number
): Promise<FolderInfo> {
  const node: FolderInfo = {
    name: folderName,
    id: folderId,
    webUrl,
    itemCount: 0,
    children: [],
  };

  if (depth >= maxDepth) return node;

  try {
    const childrenRes = await client
      .api(`/drives/${driveId}/items/${folderId}/children`)
      .select("id,name,webUrl,folder,file")
      .top(200)
      .get();

    if (childrenRes.value && Array.isArray(childrenRes.value)) {
      for (const child of childrenRes.value) {
        if (child.folder) {
          const sub = await walkFolder(
            client,
            driveId,
            child.id,
            child.name,
            child.webUrl,
            depth + 1,
            maxDepth
          );
          node.children.push(sub);
        } else if (child.file) {
          node.itemCount++;
        }
      }
    }
  } catch (err: any) {
    console.warn(`⚠️ Could not list children of "${folderName}":`, err.body || err.message);
  }

  return node;
}

function printTree(node: FolderInfo, indent = "") {
  const fileLabel = node.itemCount > 0 ? ` (${node.itemCount} file${node.itemCount === 1 ? "" : "s"})` : "";
  console.log(`${indent}📁 ${node.name}${fileLabel}`);
  console.log(`${indent}   id: ${node.id}`);
  for (const child of node.children) {
    printTree(child, indent + "   ");
  }
}

function flattenForEnv(node: FolderInfo, parentPath: string, output: { path: string; id: string }[]) {
  const myPath = parentPath ? `${parentPath}/${node.name}` : node.name;
  output.push({ path: myPath, id: node.id });
  for (const child of node.children) {
    flattenForEnv(child, myPath, output);
  }
}

async function resolveFolderIds() {
  const token = await getAccessToken();
  if (!token) {
    console.error("Could not get access token.");
    return;
  }
  const client = Client.init({
    authProvider: (done) => {
      done(null, token);
    },
  });

  const hostname = "vithiit.sharepoint.com";
  const sitePath = "sites/CSOD-VITHI";
  const MAX_DEPTH = 2; // 1 = only top-level folders, 2 = + immediate subfolders

  try {
    console.log(`🔍 Resolving IDs for Site: ${sitePath} ...\n`);

    const site = await client.api(`/sites/${hostname}:/${sitePath}`).get();
    const siteId = site.id;
    console.log(`✅ Site ID: ${siteId}`);

    const drive = await client.api(`/sites/${siteId}/drive`).get();
    const driveId = drive.id;
    console.log(`✅ Drive ID: ${driveId}\n`);

    console.log(`📂 Walking folder tree (max depth: ${MAX_DEPTH}) under "Shared Documents"...\n`);

    const rootChildren = await client
      .api(`/drives/${driveId}/root/children`)
      .select("id,name,webUrl,folder,file")
      .top(200)
      .get();

    const topLevelFolders: FolderInfo[] = [];
    if (rootChildren.value && Array.isArray(rootChildren.value)) {
      for (const item of rootChildren.value) {
        if (item.folder) {
          const sub = await walkFolder(
            client,
            driveId,
            item.id,
            item.name,
            item.webUrl,
            1,
            MAX_DEPTH
          );
          topLevelFolders.push(sub);
        }
      }
    }

    console.log("================================================================");
    console.log("📋 FOLDER TREE");
    console.log("================================================================\n");
    for (const folder of topLevelFolders) {
      printTree(folder);
      console.log("");
    }

    console.log("================================================================");
    console.log("📌 FLAT MAPPING (path → id)");
    console.log("================================================================\n");
    const flat: { path: string; id: string }[] = [];
    for (const folder of topLevelFolders) {
      flattenForEnv(folder, "", flat);
    }
    for (const f of flat) {
      console.log(`${f.path}\n  → ${f.id}\n`);
    }

    // Save tree to JSON file for easy review and reference
    const outputFile = "sharepoint_folders.json";
    fs.writeFileSync(
      outputFile,
      JSON.stringify(
        {
          siteId,
          driveId,
          sitePath,
          generatedAt: new Date().toISOString(),
          tree: topLevelFolders,
          flat,
        },
        null,
        2
      )
    );
    console.log(`💾 Saved full tree to ${outputFile}\n`);

    console.log("================================================================");
    console.log("⚙️  SUGGESTED .env CONFIG");
    console.log("================================================================");
    console.log(`SHAREPOINT_SITE_ID=${siteId}`);
    console.log(`SHAREPOINT_DRIVE_ID=${driveId}`);
    console.log(`# Pick the folders you want to sync (top-level recommended for clean categories):`);
    console.log(`# SHAREPOINT_FOLDER_IDS=<id1>,<id2>,<id3>`);
    console.log("");
    console.log("📝 Next: share this list with Vimal Sir to confirm which folders to include.");
  } catch (error: any) {
    console.error("Error resolving folder IDs:", error.body || error);
  }
}

resolveFolderIds();
