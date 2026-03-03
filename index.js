const { BotFrameworkAdapter, CardFactory, ActionTypes, MessageFactory } = require('botbuilder');
const restify = require('restify');
const fetch = require('node-fetch');

// ============================================================
// CONFIG — set these via environment variables
// ============================================================
const CONFIG = {
  microsoftAppId: process.env.MICROSOFT_APP_ID || '',
  microsoftAppPassword: process.env.MICROSOFT_APP_PASSWORD || '',
  n8nBaseUrl: process.env.N8N_WEBHOOK_BASE_URL || 'https://n8n.productondemand.co',
  port: process.env.PORT || 3978,
  // Add your n8n workflows here. Each needs a webhook trigger in n8n.
  // The webhookPath is appended to n8nBaseUrl
  workflows: JSON.parse(process.env.WORKFLOWS || '[]'),
};

// ============================================================
// DEFAULT WORKFLOWS — override with WORKFLOWS env var
// These are examples; replace with your actual workflows
// ============================================================
if (CONFIG.workflows.length === 0) {
  CONFIG.workflows = [
    {
      id: 'daily-report',
      name: '📊 Daily Report',
      description: 'Get today\'s summary across all clients',
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
      inputs: ['recipient', 'context'],
    },
    {
      id: 'check-pipeline',
      name: '💰 Pipeline Status',
      description: 'Check deal pipeline across clients',
      webhookPath: '/webhook/teams-pipeline',
      category: 'Reports',
    },
    {
      id: 'task-create',
      name: '✅ Create Task',
      description: 'Create a task in Zoho Projects',
      webhookPath: '/webhook/teams-task',
      category: 'Tasks',
      inputs: ['project', 'task_title', 'assignee'],
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
  console.error(`[Bot Error] ${error.message}`);
  await context.sendActivity('⚠️ Something went wrong. Please try again.');
};

// ============================================================
// STATE — track conversation context per user
// ============================================================
const userState = {};

function getUserState(userId) {
  if (!userState[userId]) {
    userState[userId] = { pendingWorkflow: null, pendingInputs: {}, inputStep: 0 };
  }
  return userState[userId];
}

// ============================================================
// CARD BUILDERS
// ============================================================

function buildMainMenuCard() {
  // Group workflows by category
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
      text: 'Select a workflow to run:',
      wrap: true,
      spacing: 'Small',
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
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: 'Submit',
        data: {
          action: 'submit_input',
          workflowId: workflow.id,
          inputField,
        },
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
  const resultText =
    typeof result === 'string' ? result : JSON.stringify(result, null, 2);

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
        color: 'Good',
      },
      {
        type: 'TextBlock',
        text: resultText.substring(0, 2000),
        wrap: true,
        fontType: 'Default',
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

// ============================================================
// N8N WEBHOOK CALLER
// ============================================================

async function callN8nWorkflow(workflow, inputs = {}) {
  const url = `${CONFIG.n8nBaseUrl}${workflow.webhookPath}`;
  console.log(`[n8n] Calling: ${url}`, inputs);

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

    const data = await response.json();
    return data;
  } catch (err) {
    console.error(`[n8n Error] ${err.message}`);
    return { error: true, message: `Failed to run workflow: ${err.message}` };
  }
}

// ============================================================
// MESSAGE HANDLER
// ============================================================

async function handleMessage(context) {
  const userId = context.activity.from.id;
  const state = getUserState(userId);
  const text = (context.activity.text || '').trim().toLowerCase();
  const cardData = context.activity.value; // from Adaptive Card submissions

  // --- Handle Adaptive Card button clicks ---
  if (cardData) {
    const { action, workflowId, inputField, inputValue } = cardData;

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

      // If workflow needs inputs, start collecting them
      if (workflow.inputs && workflow.inputs.length > 0) {
        state.pendingWorkflow = workflow;
        state.pendingInputs = {};
        state.inputStep = 0;
        const firstInput = workflow.inputs[0];
        await context.sendActivity({
          attachments: [
            buildInputCard(workflow, firstInput, 0, workflow.inputs.length),
          ],
        });
        return;
      }

      // No inputs needed — run immediately
      await context.sendActivity(`⏳ Running **${workflow.name}**...`);
      const result = await callN8nWorkflow(workflow);
      await context.sendActivity({
        attachments: [buildResultCard(workflow.name, result)],
      });
      return;
    }

    if (action === 'submit_input') {
      const workflow = state.pendingWorkflow;
      if (!workflow) {
        await context.sendActivity({ attachments: [buildMainMenuCard()] });
        return;
      }

      // Store the input
      state.pendingInputs[cardData.inputField] = cardData.inputValue;
      state.inputStep++;

      // Check if we need more inputs
      if (state.inputStep < workflow.inputs.length) {
        const nextInput = workflow.inputs[state.inputStep];
        await context.sendActivity({
          attachments: [
            buildInputCard(
              workflow,
              nextInput,
              state.inputStep,
              workflow.inputs.length
            ),
          ],
        });
        return;
      }

      // All inputs collected — run the workflow
      await context.sendActivity(`⏳ Running **${workflow.name}**...`);
      const result = await callN8nWorkflow(workflow, state.pendingInputs);
      state.pendingWorkflow = null;
      state.pendingInputs = {};
      state.inputStep = 0;
      await context.sendActivity({
        attachments: [buildResultCard(workflow.name, result)],
      });
      return;
    }
  }

  // --- Handle text commands ---
  if (text === 'menu' || text === 'help' || text === 'hi' || text === 'hello' || text === 'start') {
    await context.sendActivity({ attachments: [buildMainMenuCard()] });
    return;
  }

  // Default: show menu
  await context.sendActivity(
    "👋 Hey Vivek! Type **menu** or tap a button below to get started."
  );
  await context.sendActivity({ attachments: [buildMainMenuCard()] });
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
      // Welcome message when bot is installed
      if (
        context.activity.membersAdded &&
        context.activity.membersAdded.some(
          (m) => m.id !== context.activity.recipient.id
        )
      ) {
        await context.sendActivity(
          "👋 Welcome to the **POD Command Center**! I'm your personal workflow bot."
        );
        await context.sendActivity({ attachments: [buildMainMenuCard()] });
      }
    }
  });
});

server.listen(CONFIG.port, () => {
  console.log(`\n🤖 POD Teams Bot running on port ${CONFIG.port}`);
  console.log(`   Endpoint: http://localhost:${CONFIG.port}/api/messages`);
  console.log(`   Workflows loaded: ${CONFIG.workflows.length}`);
  CONFIG.workflows.forEach((w) => console.log(`     - ${w.name} → ${w.webhookPath}`));
});
