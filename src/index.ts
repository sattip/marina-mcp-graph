#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { GraphClient } from "./graph-client.js";
import { sendEmail, listEmails, readEmail, replyEmail } from "./tools/email.js";
import { listEvents, createEvent } from "./tools/calendar.js";

// Load config from environment
function loadConfig() {
  const required = (key: string): string => {
    const val = process.env[key];
    if (!val) throw new Error(`Missing required env var: ${key}`);
    return val;
  };

  return {
    tenantId: required("GRAPH_TENANT_ID"),
    clientId: required("GRAPH_CLIENT_ID"),
    clientSecret: required("GRAPH_CLIENT_SECRET"),
    defaultFromEmail: process.env.DEFAULT_FROM_EMAIL || "marinabot@obstechnologies.com",
  };
}

const config = loadConfig();
const graph = new GraphClient(config);

const server = new McpServer({
  name: "marina-mcp-graph",
  version: "1.0.0",
});

// === Email Tools ===

server.tool(
  "send_email",
  "Send an email via Microsoft Graph API",
  {
    to: z.union([z.string(), z.array(z.string())]).describe("Recipient email(s)"),
    subject: z.string().describe("Email subject"),
    body: z.string().describe("Email body (HTML supported)"),
    from: z.string().optional().describe("Sender email (defaults to marinabot@obstechnologies.com)"),
    cc: z.union([z.string(), z.array(z.string())]).optional().describe("CC recipient(s)"),
    bcc: z.union([z.string(), z.array(z.string())]).optional().describe("BCC recipient(s)"),
    contentType: z.enum(["Text", "HTML"]).optional().describe("Body content type (default: HTML)"),
  },
  async (params) => {
    const result = await sendEmail(graph, params);
    return { content: [{ type: "text" as const, text: result }] };
  }
);

server.tool(
  "list_emails",
  "List emails from a mailbox",
  {
    mailbox: z.string().describe("Email address of the mailbox"),
    folder: z.string().optional().describe("Folder name (default: inbox)"),
    top: z.number().optional().describe("Number of emails to return (default: 10)"),
    filter: z.string().optional().describe("OData filter expression"),
  },
  async (params) => {
    const result = await listEmails(graph, params);
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "read_email",
  "Read a specific email by ID",
  {
    mailbox: z.string().describe("Email address of the mailbox"),
    messageId: z.string().describe("The email message ID"),
  },
  async (params) => {
    const result = await readEmail(graph, params);
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "reply_email",
  "Reply to an email",
  {
    mailbox: z.string().describe("Email address of the mailbox"),
    messageId: z.string().describe("The email message ID to reply to"),
    body: z.string().describe("Reply body (HTML supported)"),
  },
  async (params) => {
    const result = await replyEmail(graph, params);
    return { content: [{ type: "text" as const, text: result }] };
  }
);

// === Calendar Tools ===

server.tool(
  "list_events",
  "List calendar events in a time range",
  {
    mailbox: z.string().describe("Email address of the mailbox"),
    start: z.string().describe("Start datetime (ISO 8601)"),
    end: z.string().describe("End datetime (ISO 8601)"),
    top: z.number().optional().describe("Max events to return (default: 20)"),
  },
  async (params) => {
    const result = await listEvents(graph, params);
    return { content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "create_event",
  "Create a calendar event",
  {
    mailbox: z.string().describe("Email address of the mailbox"),
    subject: z.string().describe("Event subject"),
    start: z.string().describe("Start datetime (ISO 8601)"),
    end: z.string().describe("End datetime (ISO 8601)"),
    attendees: z.array(z.string()).optional().describe("Attendee email addresses"),
    body: z.string().optional().describe("Event body/description (HTML)"),
    location: z.string().optional().describe("Event location"),
  },
  async (params) => {
    const result = await createEvent(graph, params);
    return { content: [{ type: "text" as const, text: result }] };
  }
);

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("marina-mcp-graph server running on stdio");
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
