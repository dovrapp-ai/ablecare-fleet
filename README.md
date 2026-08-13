# AbleCare Fleet

## Private ChatGPT Pro RingCentral reader

The server exposes a read-only MCP endpoint at `http://127.0.0.1:5173/mcp` for use through OpenAI Secure MCP Tunnel.

The connector intentionally exposes only two tools:

- `search`: find up to 20 recent RingCentral SMS messages by phone number, words, direction, or `recent`/`latest`.
- `fetch`: retrieve the full text and metadata for one message returned by `search`.

There is no MCP tool for sending, deleting, importing, or editing texts. The `/mcp` endpoint rejects non-loopback traffic, so deploying this code does not publicly expose RingCentral messages.

### Local setup

1. Configure the existing RingCentral environment variables: `RC_CLIENT_ID`, `RC_CLIENT_SECRET`, and `RC_JWT`.
2. Run `npm install`.
3. Run `npm start`.
4. Confirm the MCP server locally with MCP Inspector at `http://127.0.0.1:5173/mcp`.
5. Create an OpenAI Secure MCP Tunnel in the Platform dashboard, run `tunnel-client` on this machine, and point it to the local MCP URL.
6. In ChatGPT web, enable Developer mode under **Settings > Apps > Advanced Settings**, create a private app using the tunnel endpoint, and scan tools.

Keep `tunnel-client` running whenever ChatGPT needs to look up RingCentral texts.
