# marina-mcp-graph

MCP server for Microsoft Graph API. Exposes email and calendar operations as tools for AI agents — **secrets never leave the server**.

## Why?

When an AI agent (like OpenClaw/Marina) needs to send emails, the traditional approach leaks API tokens into the agent's context (prompts, logs, compaction summaries). This MCP server keeps credentials isolated:

```
[AI Agent]  →  MCP call: send_email(to, subject, body)  →  [This Server]
                                                              ↓
                                                         Graph API (with token)
                                                              ↓
                                                         Email sent ✅
```

The agent never sees `GRAPH_CLIENT_SECRET`. It just calls `send_email`.

## Tools

### Email
| Tool | Description |
|------|-------------|
| `send_email` | Send email (to, subject, body, cc, bcc) |
| `list_emails` | List emails from a mailbox |
| `read_email` | Read a specific email by ID |
| `reply_email` | Reply to an email |

### Calendar
| Tool | Description |
|------|-------------|
| `list_events` | List calendar events in a time range |
| `create_event` | Create a calendar event |

## Setup

### 1. Install

```bash
git clone https://github.com/sattip/marina-mcp-graph.git
cd marina-mcp-graph
npm install
npm run build
```

### 2. Configure

```bash
cp .env.example .env
# Edit .env with your Azure App Registration credentials
chmod 600 .env
```

Required Azure App Registration permissions (Application type, admin consented):
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `Calendars.Read`
- `Calendars.ReadWrite`

### 3. Test

```bash
npm start
# Server starts on stdio — use with an MCP client
```

### 4. Connect to OpenClaw

Add to your OpenClaw MCP config:

```json
{
  "mcpServers": {
    "graph": {
      "command": "node",
      "args": ["dist/index.js"],
      "cwd": "/path/to/marina-mcp-graph",
      "env": {
        "GRAPH_TENANT_ID": "your-tenant-id",
        "GRAPH_CLIENT_ID": "your-client-id",
        "GRAPH_CLIENT_SECRET": "your-client-secret",
        "DEFAULT_FROM_EMAIL": "marinabot@obstechnologies.com"
      }
    }
  }
}
```

Or if running on a **separate machine** (recommended for security):

```json
{
  "mcpServers": {
    "graph": {
      "url": "http://100.x.x.x:3100/mcp"
    }
  }
}
```

## Security Model

| Threat | Protection |
|--------|-----------|
| Token in AI context/logs | ✅ Token never enters agent context |
| Prompt injection asking for secrets | ✅ Agent doesn't have secrets to leak |
| Context compaction including tokens | ✅ Only tool names/results in context |
| Unauthorized API calls | ⚠️ Agent can call any exposed tool (restrict via MCP policies) |
| Server compromise | ⚠️ Mitigate by running on isolated host (Pi, VPS, container) |

## License

MIT
