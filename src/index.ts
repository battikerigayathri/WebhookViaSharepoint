import { startServer } from "./server.js";
import dotenv from "dotenv";
dotenv.config();
console.log("Starting SharePoint Webhook Service...");
try {
  startServer();
} catch (error) {
  console.error("Failed to start server:", error);
}
