import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth.js";
import dotenv from "dotenv";
dotenv.config();
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
  const folderName = "Integration Consultant";
  try {
    console.log(`Resolving IDs for Site: ${sitePath} ...`);
    // 1. Get Site ID
    const site = await client.api(`/sites/${hostname}:/${sitePath}`).get();
    const siteId = site.id;
    console.log(`✅ Site ID: ${siteId}`);
    // 2. Find the Folder ID
    console.log(`Searching for folder: "${folderName}" ...`);
    const drive = await client.api(`/sites/${siteId}/drive`).get();
    const driveId = drive.id;
    // Search for the folder inside the drive
    const search = await client.api(`/drives/${driveId}/root/children`).filter(`name eq '${folderName}'`).get();
    if (search.value && search.value.length > 0) {
      const folderId = search.value[0].id;
      console.log(`✅ Folder ID: ${folderId}`);
      console.log("\n--- Update your .env file with these values ---");
      console.log(`SHAREPOINT_SITE_ID=${siteId}`);
      console.log(`SHAREPOINT_FOLDER_ID=${folderId}`);
      console.log(`SHAREPOINT_DRIVE_ID=${driveId}`);
    } else {
      console.warn(`❌ Could not find a folder named '${folderName}' in the root.`);
      console.log("Check the folder name exactly as it appears in SharePoint.");
    }
  } catch (error: any) {
    console.error("Error resolving folder IDs:", error.body || error);
  }
}
resolveFolderIds();
