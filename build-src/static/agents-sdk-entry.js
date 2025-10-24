// Import the pre-minified browser build and expose only required exports.
import {
  CopilotStudioClient,
  CopilotStudioWebChat
} from '@microsoft/agents-copilotstudio-client/dist/src/browser.mjs';

window.MicrosoftAgents = {
  CopilotStudioClient,
  CopilotStudioWebChat
};

// Import full Web Chat; pick out what you need.
import * as FullWebChat from 'botframework-webchat';

const {
    renderWebChat,
    createDirectLine,
    createStore,
    React,
    ReactDOM
} = FullWebChat;

// Re-export for the global window.WebChat
export {
    renderWebChat,
    createDirectLine,
    createStore,
    React,
    ReactDOM
};