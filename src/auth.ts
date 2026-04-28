import * as msal from "@azure/msal-node";
import dotenv from "dotenv";
dotenv.config();
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};
const cca = new msal.ConfidentialClientApplication(msalConfig);
export async function getAccessToken(): Promise<string | null> {
  const clientCredentialRequest = {
    scopes: ["https://graph.microsoft.com/.default"],
  };
  try {
    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);
    return response?.accessToken || null;
  } catch (error) {
    console.error("Error acquiring token:", error);
    return null;
  }
}
