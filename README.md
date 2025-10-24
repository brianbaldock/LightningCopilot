# LightningCopilot
LightningCopilot enables seamless integration of Microsoft Copilot Studio agents directly within Salesforce Lightning Web Components (LWC) — with full Entra ID (Azure AD) authentication, MSAL SSO, and Adaptive Cards support.

## Project Layout
- `lightningCopilot/main/default/lwc/lightningCopilotAuth` — Lightning Web Component that hosts the Lightning Copilot experience, handles Entra ID authentication, and negotiates Direct Line chat traffic.
- `lightningCopilot/main/default/lwc/inlineError` — Lightweight helper used to surface inline error messaging inside the copilot shell.
- `build-src` and `static-resources-build` — Source and build artifacts for the Microsoft Copilot Studio web client bundle that is loaded at runtime.

## Renaming the Copilot Experience
The repo ships with the default name **LightningCopilot**. If you need to publish the copilot under a different brand, update the following touchpoints:
- `sfdx-project.json` — Change the package directory path (`lightningCopilot`) if you rename the source folder.
- `lightningCopilot/main/default/lwc/lightningCopilotAuth` — Rename the folder, files, and the exported class `LightningCopilotAuth` to match your new component name. Update the custom DOM events (`lightningcopilotauthsignin`, `lightningcopilotauthsignout`, `lightningcopilotautherror`) if anything in your org listens for them.
- `lightningCopilot/main/default/lwc/lightningCopilotAuth/lightningCopilotAuth.html` — Replace the button label “Sign in to Lightning Copilot” or any other user-facing copy with your preferred name.
- `lightningCopilot/main/default/lwc/lightningCopilotAuth/lightningCopilotAuth.js` — Update any user messaging (for example, error text that references “Lightning Copilot”) and the log prefix `[LightningCopilotAuth]` if you change the component name.
- `lightningCopilot/main/default/lwc/lightningCopilotAuth/lightningCopilotAuth.css` — Adjust the root CSS class `.lightning-copilot-shell` if you choose a different component namespace.

Keep this README aligned with your chosen name so future maintainers know which assets correspond to your copilot.
