# AbleCare Fleet Dashboard Deployment

This app must be hosted as a Node web service for Outlook, RingCentral, Square, and reminder routes to work from any link. Static hosting can show the dashboard, but it cannot run `/api/*`.

## Render Setup

1. Connect the GitHub repository to Render as a Blueprint.
2. Use `render.yaml` from the repository root.
3. Add the secret environment values in the Render dashboard.
4. Deploy the service.
5. Use the Render service URL as the daily dashboard link.

## Required Secrets

Do not commit these values to GitHub.

- `SQUARE_ACCESS_TOKEN`
- `RC_CLIENT_ID`
- `RC_CLIENT_SECRET`
- `RC_JWT`
- `RC_FROM_NUMBER`
- `RC_WEBHOOK_URL`
- `MS_GRAPH_TENANT_ID`
- `MS_GRAPH_CLIENT_ID`
- `MS_GRAPH_CLIENT_SECRET`

## Required Non-Secret Values

- `FIREBASE_DATABASE_URL`
- `SQUARE_ENVIRONMENT`
- `RC_SERVER_URL`
- `OUTLOOK_SENDER_EMAIL`
- `REMINDER_AUTOSEND_ENABLED`
- `REMINDER_SEND_HOUR`
- `REMINDER_SEND_MINUTE`
- `REMINDER_LOOKAHEAD_DAYS`

Keep `REMINDER_AUTOSEND_ENABLED=false` until production reminder previews have been checked.
