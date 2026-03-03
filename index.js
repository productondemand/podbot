const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  CardFactory,
} = require('botbuilder');
const restify = require('restify');
const fetch = require('node-fetch');

// ============================================================
// CONFIG
// ============================================================
const n8nBaseUrl = process.env.N8N_WEBHOOK_BASE_URL || 'https://n8n.productondemand.co';
const port = process.env.PORT || 3978;
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
// BOT ADAPTER — CloudAdapter for single-tenant support
// CloudAdapter reads these env vars automatically:
//   MicrosoftAppType=SingleTenant
//   MicrosoftAppId=<your app id>
//   MicrosoftAppPassword=<your client secret>
//   MicrosoftAppTenantId=<your tenant id>
// ============================================================
const botFrameworkAuth = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(botFrameworkAuth);

adapter.onTurnError = async (context, error) => {
  console.error(`[Bot Error] ${error.message}`);
  console.error(error.stack);
  await context.sendActivity('⚠️ Something went wrong. Please try again.');
};

// ============================================================
// STATE
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
  const categories = {};
  workflows.forEach((wf) => {
    const cat = wf.category || 'General';
    if (!categories[cat]) categories[cat] = [];
    categories[cat].push(wf);
  });

  const body = [
    { type: 'TextBlock', text: '🤖 POD Command Center', weight: 'Bolder', size: 'Large' },
    { type: 'TextBlock', text: 'Select a workflow to run:', wrap: true, spacing: 'Small' },
  ];

  const actions = [];

  Object.entries(categories).forEach(([category, wfs]) => {
    body.push({
      type: 'TextBlock',
      text: category,
      weight: 'Bolder',
      size: 'Medium',
      spacing: 'Medium',
      separator: true,
    });
    wfs.forEach((wf) => {
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
      { type: 'TextBlock', text: `${workflow.name} — Step ${stepIndex + 1} of ${totalSteps}`, weight: 'Bolder', size: 'Medium' },
      { type: 'TextBlock', text: `Enter **${inputField.replace(/_/g, ' ')}**:`, wrap: true },
      { type: 'Input.Text', id: 'inputValue', placeholder: `Type ${inputField.replace(/_/g, ' ')} here...` },
    ],
    actions: [
      { type: 'Action.Submit', title: 'Submit', data: { action: 'submit_input', workflowId: workflow.id, inputField } },
      { type: 'Action.Submit', title: '❌ Cancel', data: { action: 'cancel' } },
    ],
  });
}

function buildResultCard(workflowName, result) {
  const resultText = typeof result === 'string' ? result : JSON.stringify(result, null, 2);
  return CardFactory.adaptiveCard({
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: `✅ ${workflowName}`, weight: 'Bolder', size: 'Medium', color: 'Good' },
      { type: 'TextBlock', text: resultText.substring(0, 2000), wrap: true, fontType: 'Default' },
    ],
    actions: [
      { type: 'Action.Submit', title: '🏠 Back to Menu', data: { action: 'menu' } },
    ],
  });
}

// ============================================================
// N8N WEBHOOK CALLER
// ============================================================

async function callN8nWorkflow(workflow, inputs = {}) {
  const url = `${n8nBaseUrl}${workflow.webhookPath}`;
  console.log(`[n8n] Calling: ${url}`, inputs);
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ source: 'teams-bot', workflowId: workflow.id, timestamp: new Date().toISOString(), ...inputs }),
    });
    if (!response.ok) throw new Error(`n8n returned ${response.status}: ${response.statusText}`);
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
  const cardData = context.activity.value;

  if (cardData) {
    const { action, workflowId, inputField } = cardData;

    if (action === 'menu' || action === 'cancel') {
      state.pendingWorkflow = null;
      state.pendingInputs = {};
      state.inputStep = 0;
      await context.sendActivity({ attachments: [buildMainMenuCard()] });
      return;
    }

    if (action === 'run_workflow') {
      const workflow = workflows.find((w) => w.id === workflowId);
      if (!workflow) { await context.sendActivity('⚠️ Workflow not found.'); return; }

      if (workflow.inputs && workflow.inputs.length > 0) {
        state.pendingWorkflow = workflow;
        state.pendingInputs = {};
        state.inputStep = 0;
        await context.sendActivity({ attachments: [buildInputCard(workflow, workflow.inputs[0], 0, workflow.inputs.length)] });
        return;
      }

      await context.sendActivity(`⏳ Running **${workflow.name}**...`);
      const result = await callN8nWorkflow(workflow);
      await context.sendActivity({ attachments: [buildResultCard(workflow.name, result)] });
      return;
    }

    if (action === 'submit_input') {
      const workflow = state.pendingWorkflow;
      if (!workflow) { await context.sendActivity({ attachments: [buildMainMenuCard()] }); return; }

      state.pendingInputs[cardData.inputField] = cardData.inputValue;
      state.inputStep++;

      if (state.inputStep < workflow.inputs.length) {
        const nextInput = workflow.inputs[state.inputStep];
        await context.sendActivity({ attachments: [buildInputCard(workflow, nextInput, state.inputStep, workflow.inputs.length)] });
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

  if (['menu', 'help', 'hi', 'hello', 'start'].includes(text)) {
    await context.sendActivity({ attachments: [buildMainMenuCard()] });
    return;
  }

  await context.sendActivity("👋 Hey Vivek! Type **menu** or tap a button below to get started.");
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
      if (context.activity.membersAdded && context.activity.membersAdded.some((m) => m.id !== context.activity.recipient.id)) {
        await context.sendActivity("👋 Welcome to the **POD Command Center**! I'm your personal workflow bot.");
        await context.sendActivity({ attachments: [buildMainMenuCard()] });
      }
    }
  });
});

server.listen(port, () => {
  console.log(`\n🤖 POD Teams Bot running on port ${port}`);
  console.log(`   Endpoint: http://localhost:${port}/api/messages`);
  console.log(`   App Type: ${process.env.MicrosoftAppType || 'not set'}`);
  console.log(`   Tenant: ${process.env.MicrosoftAppTenantId || 'not set'}`);
  console.log(`   App ID: ${process.env.MicrosoftAppId || 'not set'}`);
  console.log(`   Workflows loaded: ${workflows.length}`);
  workflows.forEach((w) => console.log(`     - ${w.name} → ${w.webhookPath}`));
});
