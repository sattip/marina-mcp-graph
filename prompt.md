Build a Microsoft Graph API MCP server in Node.js (TypeScript).

## What it does
An MCP (Model Context Protocol) server that exposes Microsoft Graph API operations as tools. The server holds the OAuth credentials internally — the AI agent calling the tools never sees the tokens.

## Tools to implement

### Email
- `send_email(from, to, subject, body, cc?, bcc?)` — send email via Graph API
- `list_emails(mailbox, folder?, top?, filter?)` — list emails from a mailbox
- `read_email(mailbox, message_id)` — read a specific email
- `reply_email(mailbox, message_id, body)` — reply to an email

### Calendar (future)
- `list_events(mailbox, start, end)` — list calendar events
- `create_event(mailbox, subject, start, end, attendees?)` — create event

## Architecture
- Use `@modelcontextprotocol/sdk` for MCP server
- Use `@azure/msal-node` for OAuth client_credentials flow
- Config via `.env` file (NOT hardcoded):
  - `GRAPH_TENANT_ID`
  - `GRAPH_CLIENT_ID`  
  - `GRAPH_CLIENT_SECRET`
  - `DEFAULT_FROM_EMAIL=marinabot@obstechnologies.com`
- Token caching (refresh before expiry)
- Proper error handling with meaningful messages
- Transport: stdio (for OpenClaw integration)

## File structure
```
marina-mcp-graph/
├── src/
│   ├── index.ts          # MCP server entry point
│   ├── graph-client.ts   # Graph API client with auth
│   ├── tools/
│   │   ├── email.ts      # Email tools
│   │   └── calendar.ts   # Calendar tools (stub for now)
│   └── types.ts          # TypeScript types
├── .env.example          # Example config (NO real secrets)
├── package.json
├── tsconfig.json
├── README.md             # Setup instructions
└── .gitignore
```

## Important
- .env must be in .gitignore
- .env.example has placeholder values only
- README explains: how to set up, how to connect to OpenClaw, security model
- Use ES modules
- Node >= 18
- Add a "build" and "start" script

Build the complete project. After building, create an initial git commit.
