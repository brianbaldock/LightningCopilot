import * as Copilot from '@microsoft/agents-copilotstudio-client/dist/src/browser.mjs';

// Map the specific exports you actually use (to keep tree shaking).
const {
    CopilotStudioClient,
    CopilotStudioWebChat
} = Copilot;

// Assign to window
window.MicrosoftAgents = { CopilotStudioClient, CopilotStudioWebChat };