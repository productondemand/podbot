const { BotFrameworkAdapter, CardFactory, MessageFactory } = require('botbuilder');
const restify = require('restify');
const fetch = require('node-fetch');

// ============================================================
// CONFIG — supports both naming conventions for App ID/Password
// ============================================================
const CONFIG = {
  // Support both MICROSOFT_APP_ID and MicrosoftAppId (Bot Framework standard)
  microsoftAppId:
    process.env.MicrosoftAppId ||
    process.env.MICROSOFT_APP_ID ||
    '',
  microsoftAppPassword:
    process.env.MicrosoftAppPassword ||
    process.env.MICROSOFT_APP_PASSWORD ||
    '',
  microsoftAppType:
    process.env.MicrosoftAppType || 'MultiTenant',
  n8nBaseUrl:
    process.env.N8N_WEBHOOK_BASE_URL || 'https://n8n.productondemand.co',
  hubPath:
    process.env.N8N_HUB_PATH || '/webhook/teams-bot-hub',
  ollamaModel:
    process.env.OLLAMA_MODEL || 'llama3.1:8b',
  port: process.env.PORT || 3978,
  workflows: JSON.parse(process.env.WORKFLOWS || '[]'),
};

// ============================================================
// DEFAULT WORKFLOWS
// ============================================================
if (CONFIG.workflows.length === 0) {
  CONFIG.workflows = [
    {
      id: 'task-reminder',
      name: '📋 Task Reminders',
      description: 'Check overdue and upcoming tasks',
      webhookPath: '/webhook/teams-bot-hub',
      hubAction: 'reminder',
      category: 'Tasks',
    },
    {
      id: 'daily-report',
      name: '📊 Daily Report',
      description: "Get today's summary across all clients",
      webhookPath: '/webhook/teams-daily-report',
      category: 'Reports',
    },
    {
      id: 'ask-ollama',
      name: '🤖 Ask AI',
      description: 'Ask a question to the AI assistant',
      webhookPath: '/webhook/teams-bot-hub',
      hubAction: 'ollama_query',
      category: 'AI',
      inputs: ['question'],
    },
  ];
}

// ============================================================
// BOT ADAPTER
// ============================================================
const adapter = new BotFrameworkAdapter({
  appId: CONFIG.microsoftAppId,
  appPassword: CONFIG.microsoftAppPassword,
});

adapter.onTurnError = async (context, error) => {
  console.error(`[Bot Error] ${error.message}\n${error.stack}`);
  await context.sendActivity('⚠️ Something went wrong. Please try again.');
};

console.log(`[Config] AppId: ${CONFIG.microsoftAppId ? CONFIG.microsoftAppId.substring(0, 8) + '...' : 'NOT SET'}`);
console.log(`[Config] AppType: ${CONFIG.microsoftAppType}`);
console.log(`[Config] Hub: ${CONFIG.n8nBaseUrl}${CONFIG.hubPath}`);

// ============================================================
// STATE
// ============================================================
const userState = {};

function getUserState(userId) {
  if (!userState[userId]) {
    userState[userId] = {
      pendingWorkflow: null,
      pendingInputs: {},
      inputStep: 0,
      conversationHistory: [], // for future multi-turn Ollama conversations
    };
  }
  return userState[userId];
}

// ============================================================
// HUB CALLER — all actions go through the Teams Bot Hub
// ============================================================

async function callHub(payload) {
  const url = `${CONFIG.n8nBaseUrl}${CONFIG.hubPath}`;
  console.log(`[Hub] Calling ${url} with action: ${payload.action}`);

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      timeout: 120000, // 2 min for Ollama responses
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Hub returned ${response.status}: ${text}`);
    }

    return await response.json();
  } catch (err) {
    console.error(`[Hub Error] ${err.message}`);
    return { success: false, error: err.message };
  }
}

async function callN8nWorkflow(workflow, inputs = {}) {
  // If workflow routes through the hub, use callHub
  if (workflow.hubAction) {
    const payload = {
      action: workflow.hubAction,
      reply_to_teams: false, // bot handles reply itself
      ...inputs,
    };
    if (workflow.hubAction === 'ollama_query') {
      payload.prompt = inputs.question || inputs.prompt || '';
      payload.model = CONFIG.ollamaModel;
    }
    return await callHub(payload);
  }

  // Direct n8n webhook call for non-hub workflows
  const url = `${CONFIG.n8nBaseUrl}${workflow.webhookPath}`;
  console.log(`[n8n] Calling: ${url}`);

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        source: 'teams-bot',
        workflowId: workflow.id,
        timestamp: new Date().toISOString(),
        ...inputs,
      }),
    });

    if (!response.ok) {
      throw new Error(`n8n returned ${response.status}: ${response.statusText}`);
    }

    return await response.json();
  } catch (err) {
    console.error(`[n8n Error] ${err.message}`);
    return { error: true, message: `Failed to run workflow: ${err.message}` };
  }
}

// ============================================================
// CARD BUILDERS
// ============================================================

function buildMainMenuCard() {
  const categories = {};
  CONFIG.workflows.forEach((wf) => {
    const cat = wf.category || 'General';
    if (!categories[cat]) categories[cat] = [];
    categories[cat].push(wf);
  });

  const body = [
    {
      type: 'TextBlock',
      text: '🤖 POD Command Center',
      weight: 'Bolder',
      size: 'Large',
    },
    {
      type: 'TextBlock',
      text: 'Select a workflow — or just type any question to ask the AI.',
      wrap: true,
      spacing: 'Small',
      isSubtle: true,
    },
  ];

  const actions = [];

  Object.entries(categories).forEach(([category, workflows]) => {
    body.push({
      type: 'TextBlock',
      text: category,
      weight: 'Bolder',
      size: 'Medium',
      spacing: 'Medium',
      separator: true,
    });

    workflows.forEach((wf) => {
      actions.push({
        type: 'Action.Submit',
        title: wf.name,
        data: { action: 'run_workflow', workflowId: wf.id },
      });
    });
  });

  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.4',
    body,
    actions,
  });
}

function buildInputCard(workflow, inputField, stepIndex, totalSteps) {
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      {
        type: 'TextBlock',
        text: `${workflow.name} — Step ${stepIndex + 1} of ${totalSteps}`,
        weight: 'Bolder',
        size: 'Medium',
      },
      {
        type: 'TextBlock',
        text: `Enter **${inputField.replace(/_/g, ' ')}**:`,
        wrap: true,
      },
      {
        type: 'Input.Text',
        id: 'inputValue',
        placeholder: `Type ${inputField.replace(/_/g, ' ')} here...`,
        isMultiline: inputField === 'question' || inputField === 'context' || inputField === 'update_note',
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: 'Submit',
        data: { action: 'submit_input', workflowId: workflow.id, inputField },
      },
      {
        type: 'Action.Submit',
        title: '❌ Cancel',
        data: { action: 'cancel' },
      },
    ],
  });
}

function buildResultCard(workflowName, result) {
  // Extract the best text from the result
  let resultText = '';

  if (result.ollamaResponse) {
    resultText = result.ollamaResponse;
  } else if (result.message) {
    resultText = result.message;
  } else if (typeof result === 'string') {
    resultText = result;
  } else if (result.error) {
    resultText = `❌ Error: ${result.error || result.reason || 'Unknown error'}`;
  } else {
    resultText = JSON.stringify(result, null, 2);
  }

  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      {
        type: 'TextBlock',
        text: `✅ ${workflowName}`,
        weight: 'Bolder',
        size: 'Medium',
        color: result.error ? 'Attention' : 'Good',
      },
      {
        type: 'TextBlock',
        text: resultText.substring(0, 3000),
        wrap: true,
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: '🏠 Back to Menu',
        data: { action: 'menu' },
      },
    ],
  });
}

function buildOllamaCard(question, answer, model) {
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      {
        type: 'TextBlock',
        text: '🤖 AI Response',
        weight: 'Bolder',
        size: 'Medium',
        color: 'Accent',
      },
      {
        type: 'TextBlock',
        text: `**You asked:** ${question.substring(0, 200)}`,
        wrap: true,
        isSubtle: true,
        size: 'Small',
      },
      {
        type: 'TextBlock',
        text: answer.substring(0, 3000),
        wrap: true,
        spacing: 'Medium',
      },
      {
        type: 'TextBlock',
        text: `_Model: ${model || 'llama3.1:8b'}_`,
        wrap: true,
        isSubtle: true,
        size: 'Small',
        spacing: 'Small',
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: '🏠 Menu',
        data: { action: 'menu' },
      },
    ],
  });
}

// ============================================================
// MESSAGE HANDLER
// ============================================================

async function handleMessage(context) {
  const userId = context.activity.from.id;
  const state = getUserState(userId);
  const rawText = (context.activity.text || '').trim();
  const text = rawText.toLowerCase();
  const cardData = context.activity.value;

  // --- Handle Adaptive Card button clicks ---
  if (cardData) {
    const { action, workflowId } = cardData;

    if (action === 'menu' || action === 'cancel') {
      state.pendingWorkflow = null;
      state.pendingInputs = {};
      state.inputStep = 0;
      await context.sendActivity({ attachments: [buildMainMenuCard()] });
      return;
    }

    if (action === 'run_workflow') {
      const workflow = CONFIG.workflows.find((w) => w.id === workflowId);
      if (!workflow) {
        await context.sendActivity('⚠️ Workflow not found.');
        return;
      }

      if (workflow.inputs && workflow.inputs.length > 0) {
        state.pendingWorkflow = workflow;
        state.pendingInputs = {};
        state.inputStep = 0;
        await context.sendActivity({
          attachments: [buildInputCard(workflow, workflow.inputs[0], 0, workflow.inputs.length)],
        });
        return;
      }

      await context.sendActivity(`⏳ Running **${workflow.name}**...`);
      const result = await callN8nWorkflow(workflow);
      await context.sendActivity({ attachments: [buildResultCard(workflow.name, result)] });
      return;
    }

    if (action === 'submit_input') {
      const workflow = state.pendingWorkflow;
      if (!workflow) {
        await context.sendActivity({ attachments: [buildMainMenuCard()] });
        return;
      }

      state.pendingInputs[cardData.inputField] = cardData.inputValue;
      state.inputStep++;

      if (state.inputStep < workflow.inputs.length) {
        const nextInput = workflow.inputs[state.inputStep];
        await context.sendActivity({
          attachments: [buildInputCard(workflow, nextInput, state.inputStep, workflow.inputs.length)],
        });
        return;
      }

      await context.sendActivity(`⏳ Running **${workflow.name}**...`);
      const result = await callN8nWorkflow(workflow, state.pendingInputs);
      state.pendingWorkflow = null;
      state.pendingInputs = {};
      state.inputStep = 0;
      await context.sendActivity({ attachments: [buildResultCard(workflow.name, result)] });
      return;
    }
  }

  // --- Handle text commands ---
  if (text === 'menu' || text === 'help' || text === 'start') {
    await context.sendActivity({ attachments: [buildMainMenuCard()] });
    return;
  }

  if (text === 'hi' || text === 'hello') {
    await context.sendActivity("👋 Hey Vivek! Type a question to ask the AI, or type **menu** to see workflows.");
    return;
  }

  // --- Default: treat any free-text as an Ollama question ---
  await context.sendActivity(`⏳ Thinking...`);

  const hubResult = await callHub({
    action: 'ollama_query',
    prompt: rawText,
    model: CONFIG.ollamaModel,
    reply_to_teams: false,
    system: 'You are a helpful AI assistant for Vivek at ProductOnDemand, a product consulting firm. Answer concisely and clearly.',
  });

  if (hubResult.ollamaResponse) {
    await context.sendActivity({
      attachments: [buildOllamaCard(rawText, hubResult.ollamaResponse, hubResult.model_used)],
    });
  } else {
    await context.sendActivity(
      `⚠️ Couldn't get a response: ${hubResult.error || 'Hub did not return an Ollama answer. Make sure the Teams Bot Hub workflow is active and returning ollamaResponse.'}`
    );
  }
}

// ============================================================
// SERVER
// ============================================================

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, async (context) => {
    if (context.activity.type === 'message') {
      await handleMessage(context);
    } else if (context.activity.type === 'conversationUpdate') {
      if (
        context.activity.membersAdded &&
        context.activity.membersAdded.some((m) => m.id !== context.activity.recipient.id)
      ) {
        await context.sendActivity("👋 Welcome to the **POD Command Center**! Type a question to ask the AI, or type **menu** to see workflows.");
        await context.sendActivity({ attachments: [buildMainMenuCard()] });
      }
    }
  });
});

server.listen(CONFIG.port, () => {
  console.log(`\n🤖 POD Teams Bot running on port ${CONFIG.port}`);
  console.log(`   Endpoint: http://localhost:${CONFIG.port}/api/messages`);
  console.log(`   Hub: ${CONFIG.n8nBaseUrl}${CONFIG.hubPath}`);
  console.log(`   Ollama model: ${CONFIG.ollamaModel}`);
  console.log(`   Workflows loaded: ${CONFIG.workflows.length}`);
});
