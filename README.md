# POD Teams Bot — Setup Guide

Personal Teams bot that gives you interactive buttons/menus to trigger n8n workflows.

## Architecture

```
You (Teams Chat)
    ↓ click button
Teams Bot (this app, on Coolify)
    ↓ POST to webhook
n8n Workflow (on your Hetzner instance)
    ↓ returns result
Teams Bot → Adaptive Card with result → You
```

---

## Step-by-Step Setup

### Step 1: Azure Bot Registration

1. Go to https://portal.azure.com
2. Search "Azure Bot" → Create
3. Settings:
   - Bot handle: `pod-teams-bot`
   - Subscription: your subscription
   - Resource group: create `pod-teams-bot-rg`
   - Pricing: **F0 (Free)**
   - Type of App: **Single Tenant**
   - Creation type: **Create new Microsoft App ID**
4. After creation, go to the resource:
   - **Configuration** → copy the **Microsoft App ID**
   - Click **"Manage Password"** → **New client secret** → copy the value
   - Leave **Messaging endpoint** blank for now
5. Go to **Channels** → Add **Microsoft Teams** → Save

### Step 2: Deploy to Coolify

1. Push this repo to a Git repository (GitHub/GitLab)
2. In Coolify, create a new service from the Git repo
3. Set environment variables:
   ```
   MICROSOFT_APP_ID=<from step 1>
   MICROSOFT_APP_PASSWORD=<from step 1>
   N8N_WEBHOOK_BASE_URL=https://n8n.productondemand.co
   PORT=3978
   ```
4. Set the exposed port to 3978
5. Assign a domain, e.g., `pod-bot.productondemand.co`
6. Deploy — make sure HTTPS is enabled

### Step 3: Set Messaging Endpoint in Azure

Go back to Azure Portal → your bot → Configuration:
- Set Messaging endpoint to: `https://pod-bot.productondemand.co/api/messages`
- Save

### Step 4: Create Teams App Package

1. Edit `teams-manifest/manifest.json`:
   - Replace both `{{MICROSOFT_APP_ID}}` with your actual App ID
2. Add icon files to `teams-manifest/`:
   - `color.png` (192x192)
   - `outline.png` (32x32)
3. Zip the 3 files together:
   ```bash
   cd teams-manifest
   zip pod-bot.zip manifest.json color.png outline.png
   ```

### Step 5: Sideload in Teams

1. Open Microsoft Teams
2. Click "Apps" in the sidebar
3. Click "Manage your apps" (bottom left)
4. Click "Upload an app" → "Upload a custom app"
5. Select your `pod-bot.zip` file
6. Click "Add" — the bot will appear in your chat list

### Step 6: Set Up n8n Webhook Workflows

For each workflow you want accessible from the bot, create a matching
n8n workflow with a **Webhook** trigger node:

1. In n8n, create a new workflow
2. Add a **Webhook** node as the trigger
3. Set the HTTP Method to POST
4. Set the path to match your config (e.g., `/teams-daily-report`)
5. The full URL will be: `https://n8n.productondemand.co/webhook/teams-daily-report`
6. Add your workflow logic after the webhook node
7. End with a **Respond to Webhook** node that returns JSON
8. Activate the workflow

---

## Configuring Workflows

### Option A: Edit index.js directly

Modify the `CONFIG.workflows` array in `index.js`.

### Option B: Use environment variable

Set the `WORKFLOWS` env var to a JSON array:

```json
[
  {
    "id": "daily-report",
    "name": "📊 Daily Report",
    "description": "Get today's summary",
    "webhookPath": "/webhook/teams-daily-report",
    "category": "Reports"
  },
  {
    "id": "create-task",
    "name": "✅ Create Task",
    "description": "Create a task in Zoho",
    "webhookPath": "/webhook/teams-task",
    "category": "Tasks",
    "inputs": ["project", "task_title", "assignee"]
  }
]
```

### Workflow properties:
- `id` — unique identifier
- `name` — display name (supports emoji)
- `description` — shown in menu
- `webhookPath` — path appended to N8N_WEBHOOK_BASE_URL
- `category` — groups buttons in the menu
- `inputs` (optional) — array of input field names; bot will prompt for each

---

## Testing

1. Open Teams, find "POD Bot" in your chats
2. Type "menu" or "hi"
3. You should see the interactive menu with buttons
4. Click a workflow button
5. If it has inputs, fill them in step by step
6. The bot calls n8n and shows the result

---

## Troubleshooting

- **Bot doesn't respond**: Check Coolify logs, verify messaging endpoint URL
- **401 errors**: App ID or password mismatch
- **n8n webhooks fail**: Make sure workflows are active and webhook paths match
- **Can't sideload**: Your Teams admin may need to enable custom app uploads
