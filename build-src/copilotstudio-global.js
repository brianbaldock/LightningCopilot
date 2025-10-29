// Import from the package root; the exports map chooses browser.mjs for browser builds.
import { CopilotStudioClient, CopilotStudioWebChat } from '@microsoft/agents-copilotstudio-client';

window.MicrosoftAgents = {
  CopilotStudioClient,
  CopilotStudioWebChat
};