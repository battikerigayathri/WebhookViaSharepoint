import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "./auth.js";
import dotenv from "dotenv";
dotenv.config();
async function createSubscription() {
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
  const subscription = {
    changeType: "updated",
    notificationUrl: process.env.NOTIFICATION_URL,
    resource: `drives/${process.env.SHAREPOINT_DRIVE_ID}/root`,
    expirationDateTime: new Date(Date.now() + 86400000 * 2).toISOString(),
    clientState: process.env.CLIENT_STATE,
  };
  try {
    const res = await client.api("/subscriptions").post(subscription);
    console.log("Subscription created successfully:", res);
  } catch (error: any) {
    console.error("Error creating subscription:", error.body || error);
  }
}
// Run if called directly
if (import.meta.url.endsWith("subscription.js") || import.meta.url.endsWith("subscription.ts")) {
  createSubscription();
}
export { createSubscription };
