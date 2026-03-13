const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  CardFactory,
  TurnContext,
} = require('botbuilder');
const restify = require('restify');
const fetch = require('node-fetch');
const fs = require('fs');
const path = require('path');

// ============================================================
// CONFIG
// ============================================================
const n8nBaseUrl = process.env.N8N_WEBHOOK_BASE_URL || 'https://n8n.productondemand.co';
const port = process.env.PORT || 3978;
const NOTIFY_SECRET = process.env.NOTIFY_SECRET || ''; // Optional shared secret for /api/notify

// File to persist conversation references across restarts
const CONV_REF_FILE = path.join(__dirname, 'conversation-references.json');

let workflows = [];
try {
  workflows = JSON.parse(process.env.WORKFLOWS || '[]');
} catch (e) {
  workflows = [];
}

// ============================================================
// DEFAULT WORKFLOWS
// ============================================================
if (workflows.length === 0) {
  workflows = [
    {
      id: 'daily-report',
      name: '📊 Daily Report',
      description: "Get today's summary across all clients",
      webhookPath: '/webhook/teams-daily-report',
      category: 'Reports',
    },
    {
      id: 'crm-update',
      name: '📇 CRM Quick Update',
      description: 'Log a quick CRM update for a client',
      webhookPath: '/webhook/teams-crm-update',
      category: 'CRM',
      inputs: ['client_name', 'update_note'],
    },
    {
      id: 'send-followup',
      name: '📧 Send Follow-up',
      description: 'Trigger follow-up email workflow',
      webhookPath: '/webhook/teams-followup',
      category: 'Email',
    },
    {
      id: 'pipeline-status',
      name: '💰 Pipeline Status',
      description: 'Check current deal pipeline',
      webhookPath: '/webhook/teams-pipeline',
      category: 'Reports',
    },
    {
      id: 'create-task',
      name: '✅ Create Task',
      description: 'Create a new task in Notion',
      webhookPath: '/webhook/teams-task',
      category: 'Tasks',
      inputs: ['task_title', 'assignee', 'due_date'],
    },
  ];
}

// ============================================================
// CONVERSATION REFERENCE STORAGE (for proactive messaging)
// ============================================================
const conversationReferences = {};

function loadConversationReferences() {
  try {
    if (fs.existsSync(CONV_REF_FILE)) {
      const data = JSON.parse(fs.readFileSync(CONV_REF_FILE, 'utf8'));
      Object.assign(conversationReferences, data);
      console.log(`   📋 Loaded ${Object.keys(data).length} conversation reference(s)`);
    }
  } catch (e) {
    console.warn('   ⚠️  Could not load conversation references:', e.message);
  }
}

function saveConversationReferences() {
  try {
    fs.writeFileSync(CONV_REF_FILE, JSON.stringify(conversationReferences, null, 2));
  } catch (e) {
    console.warn('   ⚠️  Could not save conversation references:', e.message);
  }
}

function addConversationReference(activity) {
  const ref = TurnContext.getConversationReference(activity);
  const key = ref.user?.aadObjectId || ref.user?.id || 'default';
  conversationReferences[key] = ref;
  // Also store under a friendly key so n8n can just use "vivek" or "default"
  conversationReferences['default'] = ref;
  saveConversationReferences();
  return key;
}

// ============================================================
// BOT FRAMEWORK AUTH + ADAPTER
// ============================================================
const botFrameworkAuth = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(botFrameworkAuth);

adapter.onTurnError = async (context, error) => {
  console.error(`\n❌ [onTurnError] ${error.message}`);
  console.error(error.stack);
  await context.sendActivity('⚠️ Something went wrong. Please try again.');
};

// ============================================================
// ADAPTIVE CARD BUILDER
// ============================================================
function buildMainMenuCard() {
  const actions = workflows.map((w) => ({
    type: 'Action.Submit',
    title: w.name,
    data: { action: 'run_workflow', workflowId: w.id },
  }));

  return CardFactory.adaptiveCard({
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      {
        type: 'TextBlock',
        text: '🚀 POD Command Center',
        weight: 'Bolder',
        size: 'Large',
      },
      {
        type: 'TextBlock',
        text: 'Pick a workflow to run, or type a message to chat with Ollama.',
        wrap: true,
        spacing: 'Small',
      },
    ],
    actions: actions.slice(0, 6),
  });
}

// ============================================================
// MESSAGE HANDLER
// ============================================================
async function handleMessage(context) {
  // Always update conversation reference on every message
  const refKey = addConversationReference(context.activity);

  const text = (context.activity.text || '').trim().toLowerCase();
  const value = context.activity.value; // From Adaptive Card button clicks

  // Handle Adaptive Card submissions
  if (value && value.action === 'run_workflow') {
    const wf = workflows.find((w) => w.id === value.workflowId);
    if (wf) {
      await context.sendActivity(`⏳ Running **${wf.name}**...`);
      try {
        const res = await fetch(`${n8nBaseUrl}${wf.webhookPath}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            source: 'teams-bot',
            user: context.activity.from?.name || 'Unknown',
            ...(value.inputs || {}),
          }),
        });
        const data = await res.json();
        await context.sendActivity(
          data.message || data.result || `✅ **${wf.name}** completed.`
        );
      } catch (err) {
        await context.sendActivity(`❌ Error running ${wf.name}: ${err.message}`);
      }
      return;
    }
  }

  // Handle text commands
  if (text === 'menu' || text === 'help' || text === 'start') {
    await context.sendActivity({ attachments: [buildMainMenuCard()] });
    return;
  }

  if (text === 'status') {
    const refs = Object.keys(conversationReferences).filter(k => k !== 'default');
    await context.sendActivity(
      `🤖 **POD Bot Status**\n\n` +
      `- Workflows: ${workflows.length}\n` +
      `- Proactive messaging: ✅ Ready\n` +
      `- Conversation refs stored: ${refs.length}\n` +
      `- Your ref key: ${refKey}`
    );
    return;
  }

  // Default: forward to Bot Hub for Ollama processing
  try {
    await context.sendActivity('🤔 Thinking...');
    const res = await fetch(`${n8nBaseUrl}/webhook/teams-bot-hub`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'ollama_query',
        prompt: context.activity.text,
        reply_to_teams: false, // Bot will handle the reply directly
        model: 'llama3.1:8b',
        source: 'teams-bot',
        user: context.activity.from?.name || 'Unknown',
      }),
    });
    const data = await res.json();
    if (data.success && data.ollamaResponse) {
      await context.sendActivity(data.ollamaResponse);
    } else if (data.error) {
      await context.sendActivity(`⚠️ ${data.error}`);
    } else {
      await context.sendActivity('🤷 No response from Ollama. Try again.');
    }
  } catch (err) {
    await context.sendActivity(`❌ Error: ${err.message}`);
  }
}

// ============================================================
// SERVER
// ============================================================
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

// --- Main bot messaging endpoint (Teams → Bot) ---
server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, async (context) => {
    if (context.activity.type === 'message') {
      await handleMessage(context);
    } else if (context.activity.type === 'conversationUpdate') {
      // Store ref on conversation update too
      if (context.activity.membersAdded) {
        const isBotAdded = context.activity.membersAdded.some(
          (m) => m.id === context.activity.recipient.id
        );
        if (!isBotAdded) {
          // User was added (or started convo) — store their reference
          addConversationReference(context.activity);
          await context.sendActivity(
            "👋 Welcome to the **POD Command Center**! I'm your personal workflow bot.\n\n" +
            "Type **menu** for workflows, or just type a question to chat with Ollama.\n\n" +
            "✅ Proactive messaging is enabled — I can send you reminders!"
          );
          await context.sendActivity({ attachments: [buildMainMenuCard()] });
        }
      }
    }
  });
});

// --- Proactive notify endpoint (n8n → Bot → Teams 1:1 chat) ---
server.post('/api/notify', async (req, res) => {
  try {
    const body = req.body || {};

    // Optional: validate shared secret
    if (NOTIFY_SECRET && body.secret !== NOTIFY_SECRET) {
      res.send(401, { error: 'Invalid or missing secret' });
      return;
    }

    const message = body.message || body.text || '';
    if (!message) {
      res.send(400, { error: 'Missing required field: message (or text)' });
      return;
    }

    // Determine which user to notify
    const userKey = body.user_key || 'default';
    const conversationRef = conversationReferences[userKey];

    if (!conversationRef) {
      res.send(404, {
        error: `No conversation reference found for key: "${userKey}". ` +
               `The user must message the bot at least once first. ` +
               `Available keys: ${Object.keys(conversationReferences).join(', ')}`,
      });
      return;
    }

    // Send proactive message
    await adapter.continueConversationAsync(
      process.env.MicrosoftAppId,
      conversationRef,
      async (turnContext) => {
        // Support both plain text and HTML
        if (body.content_type === 'html') {
          await turnContext.sendActivity({
            type: 'message',
            textFormat: 'xml',
            text: message,
          });
        } else if (body.card) {
          // Support sending Adaptive Cards
          const card = CardFactory.adaptiveCard(
            typeof body.card === 'string' ? JSON.parse(body.card) : body.card
          );
          await turnContext.sendActivity({ attachments: [card] });
        } else {
          await turnContext.sendActivity(message);
        }
      }
    );

    res.send(200, {
      success: true,
      delivered_to: userKey,
      timestamp: new Date().toISOString(),
    });
  } catch (err) {
    console.error('❌ /api/notify error:', err.message);
    res.send(500, { error: err.message });
  }
});

// --- Health check ---
server.get('/api/health', async (req, res) => {
  res.send(200, {
    status: 'ok',
    bot: 'POD Teams Bot',
    workflows: workflows.length,
    conversationRefs: Object.keys(conversationReferences).filter(k => k !== 'default').length,
    proactiveReady: Object.keys(conversationReferences).length > 0,
    uptime: process.uptime(),
  });
});

// ============================================================
// STARTUP
// ============================================================
loadConversationReferences();

server.listen(port, () => {
  console.log(`\n🤖 POD Teams Bot running on port ${port}`);
  console.log(`   Endpoint: http://localhost:${port}/api/messages`);
  console.log(`   Notify:   http://localhost:${port}/api/notify`);
  console.log(`   Health:   http://localhost:${port}/api/health`);
  console.log(`   App Type: ${process.env.MicrosoftAppType || 'not set'}`);
  console.log(`   Tenant:   ${process.env.MicrosoftAppTenantId || 'not set'}`);
  console.log(`   App ID:   ${process.env.MicrosoftAppId || 'not set'}`);
  console.log(`   Notify Secret: ${NOTIFY_SECRET ? '✅ Set' : '⚠️  Not set (open)'}`);
  console.log(`   Workflows loaded: ${workflows.length}`);
  workflows.forEach((w) => console.log(`     - ${w.name} → ${w.webhookPath}`));
  console.log(`   Proactive messaging: ${Object.keys(conversationReferences).length > 0 ? '✅ Ready' : '⏳ Waiting for first message'}`);
});
