import { LightningElement, track, wire, api } from 'lwc';
import { loadScript } from 'lightning/platformResourceLoader';
import msalBrowserResource from '@salesforce/resourceUrl/msalBrowser';
import copilotStudioClient from '@salesforce/resourceUrl/copilotStudioClient';
import adaptiveCardResource from '@salesforce/resourceUrl/adaptiveCard';

import CLIENT_ID from '@salesforce/label/c.MSAL_ClientId';
import TENANT_ID from '@salesforce/label/c.MSAL_TenantId';
import REDIRECT_URI from '@salesforce/label/c.MSAL_RedirectUri';
import SCOPES from '@salesforce/label/c.MSAL_Scopes';
import COPILOT_EMBED_URL from '@salesforce/label/c.COPILOT_EmbedUrl';
import COPILOT_AGENT_URL from '@salesforce/label/c.COPILOT_AgentUrl';

import userId from '@salesforce/user/Id';
import { getRecord } from 'lightning/uiRecordApi';

const USER_EMAIL_FIELD = 'User.Email';
const USER_USERNAME_FIELD = 'User.Username';

const ALLOWED_EMBED_HOSTS = [
    'copilotstudio.microsoft.com',
    'copilotstudio-df.microsoft.com',
    'web.powervirtualagents.microsoft.com',
    'web.powervirtualagents-df.microsoft.com',
    'api.bap.microsoft.com'
]; // Only these hosts permitted

let msalLib;
let msalInstance;

function parseIdsFromAgentUrl(rawUrl) {
    try {
        const u = new URL(rawUrl);
        // Expect: /environments/{envId}/bots/{botId}/...
        const parts = u.pathname.split('/').filter(Boolean);
        const envIdx = parts.indexOf('environments');
        const botIdx = parts.indexOf('bots');
        const environmentId = envIdx >= 0 ? parts[envIdx + 1] : '';
        const botId = botIdx >= 0 ? parts[botIdx + 1] : '';
        return { environmentId, botId };
    } catch {
        return { environmentId: '', botId: '' };
    }
}

/* ===== Helpers ===== */

// MSAL diagnostics helper
class MsalDiag {
    static extract(e) {
        return {
            name: e?.name,
            code: (e?.errorCode || e?.code || '').toLowerCase(),
            subError: (e?.subError || '').toLowerCase(),
            message: e?.errorMessage || e?.message,
            status: e?.status || e?.response?.status
        };
    }
    static isInteraction(diag) {
        return ['interaction_required','login_required','consent_required','no_tokens_found'].includes(diag.code)
            || (diag.subError && ['message_only','basic_action','additional_action'].includes(diag.subError));
    }
}

// Lightning Copilot component responsible for authentication and Direct Line chat orchestration.
export default class LightningCopilotAuth extends LightningElement {
    @track signedIn = false;
    @track accountLabel = '';
    @track error;
    @track iframeLoading = true;
    @track chatConnecting = false;
    @track showLoginButton = false;
    @track statusMessage = 'Checking sign-in...';
    @track isAgentTyping = false;
    @track transcript = [];

    _agentIds = { environmentId: '', botId: '' };

    salesforceUserEmail;
    _initialized = false;
    _ssoAttempted = false;
    _loginInFlight = false;
    @api disableSilentInIframe = false;

    parsedEmbedOrigin;
    sanitizedEmbedUrl = '';
    validEmbed = false;

    _aadToken;
    _aadTokenExpires;

    // Direct Line chat state (Option B)
    conversationId;
    watermark = null;
    dlToken;
    dlDomain = 'https://directline.botframework.com';
    pollAbort;
    _pollTimeoutId;
    dlClient; // DirectLine client from SDK (streaming)
    _dlStarted = false; // idempotency guard to avoid reconnect loops
    _oauthHandled = new Set(); // connectionName dedupe within a session
    _scriptsReady = false;
    _activitySubscription;
    _typingTimeout;
    _authCheckComplete = false;
    adaptiveCardsLib;
    _clientActivityIdSeed = 0;
    _scrollFrame;
    _dlSessionToken = 0;

    get authBusy() {
        if (this.error) return false;
        return !this.showLoginButton && (!this._scriptsReady || !msalInstance || this._loginInFlight || !this._authCheckComplete);
    }

    get composerDisabled() {
        const streamingReady = !!this.dlClient;
        const restReady = !!(this.dlToken && this.conversationId);
        return !(streamingReady || restReady) || this.chatConnecting;
    }

    get inIframe() {
        try {
            return window.self !== window.top;
        } catch {
            return true;
        }
    }

    get scopesArr() {
        return (SCOPES || '').split(/\s+/).filter(Boolean);
    }

    // New helpers to avoid multi-resource scope errors
    get oidcScopes() {
        return ['openid','profile','email','offline_access'].filter(s =>
            this.scopesArr.includes(s)
        );
    }

    separateResourceScopes() {
        // Return a map: { powerPlatform: [...], customApi: [...], others: [...] }
        const resourceScopes = this.scopesArr.filter(s =>
            !this.oidcScopes.includes(s)
        );
        const buckets = {
            powerPlatform: [],
            customApi: [],
            others: []
        };
        resourceScopes.forEach(s => {
            if (s.startsWith('https://api.powerplatform.com/')) {
                buckets.powerPlatform.push(s);
            } else if (s.startsWith('api://08dcb614-2a48-4238-937e-51c7e36d6b0b/')) {
                buckets.customApi.push(s);
            } else {
                buckets.others.push(s);
            }
        });
        return buckets;
    }

    // Choose the resource scopes needed for the embed session.
    get powerPlatformLoginScopes() {
        const { powerPlatform } = this.separateResourceScopes();
        // Prefer the modern Copilot Studio delegated scope; then legacy; then .default; then first PP scope.
        const prefer = [
            /\/CopilotStudio\.Copilots\.Invoke$/i,
            /user_impersonation$/i,
            /\.default$/i
        ];
        let required = null;
        for (const rx of prefer) {
            required = powerPlatform.find(s => rx.test(s));
            if (required) break;
        }
        // If none matched but caller provided at least one PP scope, use the first.
        if (!required && powerPlatform.length) {
            required = powerPlatform[0];
        }
        // Final fallback if label has no PP scope at all.
        if (!required) {
            return ['https://api.powerplatform.com/CopilotStudio.Copilots.Invoke', ...this.oidcScopes];
        }
        return [required, ...this.oidcScopes];
    }

    // Helper to supply Chat.Invoke scopes when a custom API token is required
    get customApiLoginScopes() {
        const { customApi } = this.separateResourceScopes();
        if (!customApi.length) return [];
        // Keep only ONE resource set + OIDC
        return [customApi[0], ...this.oidcScopes];
    }

    logDebug(msg, obj) {
        // eslint-disable-next-line no-console
        console.info('[LightningCopilotAuth]', msg, obj || '');
    }

    async loadAdaptiveCardsLibrary() {
        if (this.adaptiveCardsLib) {
            return;
        }
        this.adaptiveCardsLib = this.resolveAdaptiveCardsLib();
        if (this.adaptiveCardsLib) {
            return;
        }
        try {
            await loadScript(this, adaptiveCardResource);
        } catch (e) {
            this.logDebug('AdaptiveCards script load failed', e);
        }
        this.adaptiveCardsLib = this.resolveAdaptiveCardsLib();
        if (!this.adaptiveCardsLib) {
            this.logDebug('AdaptiveCards library unavailable after load. Ensure the static resource "adaptiveCard" points to adaptivecards.min.js.');
        }
    }

    resolveAdaptiveCardsLib() {
        if (window.AdaptiveCards) {
            return window.AdaptiveCards;
        }
        if (window.WebChat) {
            if (window.WebChat.AdaptiveCards) {
                return window.WebChat.AdaptiveCards;
            }
            if (window.WebChat.dependencies?.AdaptiveCards) {
                return window.WebChat.dependencies.AdaptiveCards;
            }
        }
        if (window.MicrosoftAgents?.AdaptiveCards) {
            return window.MicrosoftAgents.AdaptiveCards;
        }
        return null;
    }

    renderAdaptiveCardsInTranscript() {
        if (!this.adaptiveCardsLib) {
            return;
        }
        this.transcript.forEach(message => {
            (message.attachments || []).forEach(cardAttachment => {
                if (!cardAttachment.isAdaptiveCard) {
                    return;
                }
                const cardId = cardAttachment.id;
                const host = this.template.querySelector(`[data-card-id="${cardId}"]`);
                if (!host || host.childElementCount) {
                    return;
                }
                this.renderAdaptiveCardIntoHost(cardAttachment.content, host);
            });
        });
    }

    renderAdaptiveCardIntoHost(cardPayload, hostElement) {
        if (!this.adaptiveCardsLib || !hostElement) {
            return;
        }
        try {
            const AdaptiveCards = this.adaptiveCardsLib;
            const card = new AdaptiveCards.AdaptiveCard();
            if (AdaptiveCards.HostConfig) {
                card.hostConfig = new AdaptiveCards.HostConfig({
                    fontFamily: 'Segoe UI, Helvetica Neue, Arial, sans-serif'
                });
            }
            card.onExecuteAction = action => {
                try {
                    if (action.type === 'Action.OpenUrl' && action.url) {
                        window.open(action.url, '_blank', 'noopener,noreferrer');
                    } else if (action.type === 'Action.Submit') {
                        this.postActivity({
                            type: 'event',
                            name: 'adaptiveCard/action',
                            value: action.data,
                            from: { id: 'user', name: this.accountLabel || 'You' }
                        }).catch(err => this.setError(err));
                    }
                } catch (e) {
                    this.logDebug('Adaptive card action failed', e);
                }
            };
            let payload = cardPayload || {};
            if (typeof payload === 'string') {
                try {
                    payload = JSON.parse(payload);
                } catch {
                    payload = {};
                }
            }
            card.parse(JSON.parse(JSON.stringify(payload)));
            const renderedCard = card.render();
            hostElement.innerHTML = '';
            hostElement.appendChild(renderedCard);
        } catch (e) {
            this.logDebug('Adaptive card render failed', e);
            hostElement.textContent = 'Unable to render adaptive card.';
        }
    }

    generateClientActivityId() {
        this._clientActivityIdSeed += 1;
        return `lwc-${Date.now()}-${this._clientActivityIdSeed}`;
    }

    pushTranscriptEntry(entry) {
        const normalized = this.decorateMessageEntry({
            id: entry.id || this.generateClientActivityId(),
            key: entry.key || `${entry.from || 'agent'}-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            from: entry.from || 'agent',
            text: entry.text || '',
            attachments: entry.attachments || [],
            actions: entry.actions || [],
            timestamp: entry.timestamp || new Date().toISOString(),
            status: entry.status || 'delivered',
            clientActivityId: entry.clientActivityId || null
        });
        this.transcript = [...this.transcript, normalized];
        this.scheduleScrollToBottom();
        return normalized;
    }

    replaceTranscriptEntry(index, updatedEntry) {
        if (index < 0 || index >= this.transcript.length) {
            return;
        }
        const updated = this.decorateMessageEntry({ ...this.transcript[index], ...updatedEntry });
        const next = [...this.transcript];
        next.splice(index, 1, updated);
        this.transcript = next;
    }

    markTranscriptEntryStatus(clientActivityId, status) {
        if (!clientActivityId) return;
        const index = this.transcript.findIndex(msg => msg.clientActivityId === clientActivityId);
        if (index >= 0) {
            this.replaceTranscriptEntry(index, { status });
        }
    }

    replaceOrAppendFromActivity(activity, messageFactory) {
        const clientActivityId = activity.channelData?.clientActivityID || activity.clientActivityID;
        if (clientActivityId) {
            const index = this.transcript.findIndex(msg => msg.clientActivityId === clientActivityId);
            if (index >= 0) {
                const messageData = messageFactory();
                this.replaceTranscriptEntry(index, {
                    status: 'delivered',
                    attachments: messageData.attachments,
                    text: messageData.text,
                    actions: messageData.actions || []
                });
                return true;
            }
        }
        return false;
    }

    scheduleScrollToBottom() {
        if (this._scrollFrame) {
            cancelAnimationFrame(this._scrollFrame);
        }
        this._scrollFrame = requestAnimationFrame(() => {
            const log = this.template.querySelector('[data-id="log"]');
            if (log) {
                log.scrollTop = log.scrollHeight;
            }
        });
    }

    determineActivityRole(activity) {
        const role = (activity.from?.role || '').toLowerCase();
        if (role === 'user') return 'user';
        return 'agent';
    }

    mapActivityAttachments(activity) {
        const attachments = Array.isArray(activity.attachments) ? activity.attachments : [];
        return attachments
            .filter(att => att && att.contentType)
            .map((att, idx) => {
                const normalizedContent = att.content;
                const contentText = typeof normalizedContent === 'string'
                    ? normalizedContent
                    : normalizedContent ? JSON.stringify(normalizedContent, null, 2) : '';
                return {
                    id: `${activity.id || activity.replyToId || this.generateClientActivityId()}-att-${idx}`,
                    contentType: (att.contentType || '').toLowerCase(),
                    content: contentText,
                    name: att.name || att.contentType || 'Attachment',
                    contentUrl: att.contentUrl,
                    isAdaptiveCard: (att.contentType || '').toLowerCase() === 'application/vnd.microsoft.card.adaptive'
                };
            });
    }

    buildMessageFromActivity(activity) {
        const attachments = this.mapActivityAttachments(activity);
        return this.decorateMessageEntry({
            id: activity.id || this.generateClientActivityId(),
            key: (activity.id || this.generateClientActivityId()) + '-msg',
            from: this.determineActivityRole(activity),
            text: activity.text || activity.speak || '',
            attachments,
            actions: [],
            timestamp: activity.timestamp || new Date().toISOString(),
            status: 'delivered',
            clientActivityId: activity.channelData?.clientActivityID || activity.clientActivityId || null
        });
    }

    decorateMessageEntry(entry) {
        const direction = entry.from === 'user' ? 'user' : 'agent';
        const status = entry.status || 'delivered';
        const statusClass = this.deriveStatusClass(status);
        return {
            ...entry,
            from: direction,
            direction,
            lineClass: `chat-line ${direction}`,
            bubbleClass: `chat-bubble ${direction} ${statusClass}`,
            status,
            statusClass,
            isPending: status === 'pending',
            isFailed: status === 'failed'
        };
    }

    deriveStatusClass(status) {
        switch (status) {
        case 'pending':
            return 'pending';
        case 'failed':
            return 'failed';
        case 'sent':
            return 'sent';
        default:
            return 'delivered';
        }
    }

    @wire(getRecord, {
        recordId: userId,
        fields: [USER_EMAIL_FIELD, USER_USERNAME_FIELD]
    })
    wiredUser({ data, error }) {
        if (data) {
            const email = data.fields.Email?.value;
            const username = data.fields.Username?.value;
            this.salesforceUserEmail = email || username;
            if (msalInstance && !this.signedIn && !this._ssoAttempted) {
                this.trySilentBootstrap();
            }
        } else if (error) {
            // eslint-disable-next-line no-console
            console.info('[LightningCopilotAuth] User email unavailable for silent SSO.', error);
        }
    }

    renderedCallback() {
        if (!this._initialized) {
            this._initialized = true;
            this.statusMessage = 'Checking sign-in...';
            this.showLoginButton = false;
            this.isAgentTyping = false;
            this._authCheckComplete = false;
            this.validateEmbedLabel();
            this.ensureHostObserver();
            Promise.all([
                loadScript(this, msalBrowserResource),
                loadScript(this, copilotStudioClient)
            ])
            .then(async () => {
                msalLib = window.msal;
                if (!msalLib) {
                    throw new Error('MSAL library not available after load.');
                }
                await this.loadAdaptiveCardsLibrary();
                this.initMsal();
                // Direct Line chat starts post sign-in via onSignedIn()
            })
            .catch(e => this.setError(e))
            .finally(() => {
                this._scriptsReady = true;
            });
        }

        this.renderAdaptiveCardsInTranscript();
    }

    validateEmbedLabel() {
        const agentSrc = (this._manualOverride || COPILOT_AGENT_URL || COPILOT_EMBED_URL || '').trim();
        this.logDebug('validateEmbedLabel()', agentSrc);
        if (!agentSrc) {
            this.setError(new Error('COPILOT_AgentUrl/COPILOT_EmbedUrl label empty.'));
            return;
        }
        let url;
        try {
            url = new URL(agentSrc);
        } catch {
            this.setError(new Error('Invalid Lightning Copilot agent URL format.'));
            return;
        }
        if (url.protocol !== 'https:') {
            this.setError(new Error('Embed URL must use HTTPS.'));
            return;
        }
        if (!ALLOWED_EMBED_HOSTS.includes(url.host)) {
            this.setError(new Error('Embed host not in allowlist.'));
            return;
        }
        if (/\/canvas(\/|\?|$)/i.test(url.pathname)) {
            url.pathname = url.pathname.replace(/\/canvas/i, '/WebChat');
        }
        if (!/\/bots\//i.test(url.pathname)) {
            this.setError(new Error('Agent URL must contain /bots/ path.'));
            return;
        }
        url.hash = '';
        this.sanitizedEmbedUrl = url.toString();
        this.parsedEmbedOrigin = url.origin;
        this.validEmbed = true;
        this._agentIds = parseIdsFromAgentUrl(this.sanitizedEmbedUrl);
        this.logDebug('[SDK] ids (json)', JSON.parse(JSON.stringify(this._agentIds))); // force real display
    }

    sanitizeDirectLineDomain(rawDomain) {
        if (!rawDomain) {
            return null;
        }
        try {
            const candidate = rawDomain.trim();
            if (!candidate) {
                return null;
            }
            const ensured = candidate.startsWith('http') ? candidate : `https://${candidate}`;
            const url = new URL(ensured);
            if (url.protocol !== 'https:') {
                return null;
            }
            const host = url.host.toLowerCase();
            if (!host.includes('botframework.')) {
                return null;
            }
            return `${url.protocol}//${host}`;
        } catch {
            return null;
        }
    }

    get effectiveRedirectUri() {
        const raw = (REDIRECT_URI || '').trim();
        try {
            if (raw) {
                const u = new URL(raw, window.location.href);
                if (u.protocol === 'https:' || window.location.protocol !== 'https:') {
                    return u.origin + u.pathname; // strip query/fragment
                }
            }
        } catch {}
        return window.location.origin;
    }

    initMsal() {
        msalInstance = new msalLib.PublicClientApplication({
            auth: {
                clientId: CLIENT_ID,
                authority: `https://login.microsoftonline.com/${TENANT_ID}`,
                redirectUri: this.effectiveRedirectUri,
                navigateToLoginRequestUrl: false
            },
            // Keep tokens in sessionStorage to avoid long-lived persistence if an XSS occurs.
            cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
            system: {
                allowRedirectInIframe: true,
                allowNativeBroker: false,
                loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false, logLevel: (msalLib?.LogLevel?.Error) || 3 }
            }
        });

        msalInstance.addEventCallback(evt => {
            if (evt.eventType === msalLib.EventType.LOGIN_SUCCESS && evt.payload?.account && evt.payload?.authenticationResult) {
                // Redirect flows come through here
                this.onSignedIn(evt.payload.authenticationResult);
            }
        });

        msalInstance.handleRedirectPromise()
            .then(result => {
                this._loginInFlight = false;
                if (result?.account) {
                    this.onSignedIn(result);
                    return;
                }
                // Kick off the auth flow immediately: silent → popup (no redirect)
                this.startAuthFlow();
            })
            .catch(e => {
                this._loginInFlight = false;
                // Even if redirect promise errors, continue with non-redirect flow
                this.startAuthFlow();
                this.setError(e);
            });
    }

    async trySilentBootstrap() {
        if (this._ssoAttempted) return;
        if (this.disableSilentInIframe) return;
        this._ssoAttempted = true;
        if (!msalInstance) return;

        const loginHint = this.salesforceUserEmail;
        if (!loginHint) {
            this.logDebug('Silent bootstrap: no loginHint yet');
            return;
        }

        const scopes = this.powerPlatformLoginScopes;
        try {
            this.logDebug('Silent bootstrap ssoSilent', { loginHint, scopes });
            const resp = await msalInstance.ssoSilent({ loginHint, scopes });
            if (resp?.account) {
                msalInstance.setActiveAccount(resp.account);
                this.onSignedIn(resp); // pass whole result
            }
        } catch (e) {
            const code = (e?.errorCode || e?.code || '').toLowerCase();
            this.logDebug('Silent bootstrap failed', code);
            // Interactive fallback (popup) is handled by startAuthFlow()
        }
    }

    async handleLogin() {
        if (this._loginInFlight) return;
        this.error = undefined;
        if (!msalInstance) {
            this.setError(new Error('Authentication client not ready.')); // prevents null deref
            return;
        }
        this.statusMessage = 'Redirecting to sign-in...';
        this.showLoginButton = false;
        return this.startAuthFlow(true);
    }

    async startAuthFlow(interactive = false) {
        if (this._loginInFlight) return;

        if (!interactive) {
            this.showLoginButton = false;
            this.statusMessage = 'Checking sign-in...';
        }
        this._authCheckComplete = false;

        // If already logged in & we have a token, just proceed
        if (this.signedIn && this._aadToken) {
            this._authCheckComplete = true;
            return;
        }

        const scopes = this.powerPlatformLoginScopes;
        const accounts = msalInstance?.getAllAccounts?.() || [];
        const active = msalInstance?.getActiveAccount?.() || accounts[0] || null;
        const loginHint = this.salesforceUserEmail;

        // 1) No account detected
        if (!active) {
            if (!interactive) {
                this.showLoginButton = true;
                this.statusMessage = 'We couldn\'t sign you in automatically. Please sign in.';
                this._authCheckComplete = true;
                return;
            }
            this._loginInFlight = true;
            try {
                await msalInstance.loginRedirect({ scopes, loginHint });
                return; // resumes via handleRedirectPromise/addEventCallback
            } catch (e) {
                this.setError(e);
            } finally {
                this._loginInFlight = false;
            }
            return;
        }

        // 2) Try silent token for existing account
        try {
            const silentResp = await msalInstance.acquireTokenSilent({ scopes, account: active });
            if (silentResp?.account) {
                this._authCheckComplete = true;
                this.onSignedIn(silentResp);
                return;
            }
        } catch (e) {
            this.logDebug('startAuthFlow acquireTokenSilent failed', MsalDiag.extract(e));
        }

        if (!interactive) {
            this.showLoginButton = true;
            this.statusMessage = 'We couldn\'t sign you in automatically. Please sign in.';
            this._authCheckComplete = true;
            return;
        }

        // 3) Interactive fallback via redirect
        this._loginInFlight = true;
        try {
            await msalInstance.acquireTokenRedirect({ scopes, account: active });
            return; // resumes via handleRedirectPromise/addEventCallback
        } catch (e) {
            this.setError(e);
            this.showLoginButton = true;
            this.statusMessage = 'Sign-in failed. Please try again.';
            this._authCheckComplete = true;
        } finally {
            this._loginInFlight = false;
        }
    }

    // Accept the full AuthenticationResult when available
    onSignedIn(authResultOrAccount) {
        let acct;
        if (authResultOrAccount?.accessToken) {
            // AuthenticationResult
            acct = authResultOrAccount.account;
            this._aadToken = authResultOrAccount.accessToken;
            this._aadTokenExpires = authResultOrAccount.expiresOn; // Date
            this.logDebug('Stored initial AAD token', {
                expiresOn: this._aadTokenExpires,
                scopes: authResultOrAccount.scopes
            });
        } else {
            // Legacy path where only account passed (keep minimal support)
            acct = authResultOrAccount;
        }
        if (!acct) {
            this.logDebug('onSignedIn called without account');
            return;
        }

        this.accountLabel = (acct.username || acct.name || 'Signed in');
        this.signedIn = true;
        this.showLoginButton = false;
        this.statusMessage = '';
        this.isAgentTyping = false;
        this._authCheckComplete = true;
        this.dispatchEvent(new CustomEvent('lightningcopilotauthsignin', {
            detail: { username: acct.username, name: acct.name }
        }));
        this.iframeLoading = true;
        // Start native Direct Line chat (no WebChat dependency)
        this.startDirectLineChat();
        // Keep silent refresh scheduling (optional for later token use)
        if (this._aadTokenExpires instanceof Date) {
            const msUntilExpiry = this._aadTokenExpires.getTime() - Date.now();
            const refreshIn = Math.max(msUntilExpiry - 60000, 30000);
            if (msUntilExpiry > 0) {
                setTimeout(() => this.refreshAadToken().catch(() => {}), refreshIn);
            }
        }
    }

    // Override debugSnapshot to include more info
    @api
    debugSnapshot() {
        return {
            signedIn: this.signedIn,
            accountLabel: this.accountLabel,
            rawLabel: COPILOT_EMBED_URL,
            manualOverride: this._manualOverride || null,
            validEmbed: this.validEmbed,
            sanitizedEmbedUrl: this.sanitizedEmbedUrl,
            iframeLoading: this.iframeLoading,
            hasAadToken: !!this._aadToken,
            aadTokenExpires: this._aadTokenExpires ? this._aadTokenExpires.toISOString() : null,
            error: this.error,
            agentIds: this._agentIds,
            hasMicrosoftAgentsGlobal: !!window.MicrosoftAgents,
            dlStarted: this._dlStarted
        };
    }

    async refreshAadToken() {
        if (!msalInstance?.getActiveAccount()) return;
        const scopes = this.powerPlatformLoginScopes;
        try {
            const resp = await msalInstance.acquireTokenSilent({
                scopes,
                account: msalInstance.getActiveAccount()
            });
            if (resp?.accessToken) {
                this._aadToken = resp.accessToken;
                this._aadTokenExpires = resp.expiresOn;
                this.logDebug('Token silently refreshed', { expiresOn: this._aadTokenExpires });
                // If needed, you could renew Direct Line via createConnection here.
            }
        } catch (e) {
            this.logDebug('Silent token refresh failed', MsalDiag.extract(e));
            // Do not force interactive—user continues until token truly expires
        }
    }

    // Direct Line integration
    async startDirectLineChat() {
        if (this._dlStarted || this.chatConnecting) {
            return;
        }

        const context = this.getDirectLineContext();
        if (!context) {
            return;
        }

        const sessionToken = ++this._dlSessionToken;
        this.chatConnecting = true;
        try {
            const client = this.createCopilotClient(context);
            const connection = await window.MicrosoftAgents.CopilotStudioWebChat.createConnection(
                client,
                { showTyping: true }
            );

            if (sessionToken !== this._dlSessionToken) {
                this.disposeDirectLineClient(connection);
                return;
            }

            if (this.tryAdoptStreamingConnection(connection, sessionToken)) {
                return;
            }

            await this.adoptRestConnection(connection, sessionToken);
        } catch (e) {
            const msg = (e && e.message) || '';
            if (/Failed to fetch/i.test(msg) || /CSP|Content Security Policy|blocked/i.test(msg)) {
                this.setError(new Error('CSP blocked Direct Line. Ask admin to add https://directline.botframework.com (and regional *.botframework.com if needed) to CSP Trusted Sites (connect-src).'));
            } else {
                this.setError(e);
            }
        } finally {
            if (sessionToken === this._dlSessionToken) {
                this.chatConnecting = false;
            }
        }
    }

    async startDirectLineConversation() {
        const res = await fetch(`${this.dlDomain}/v3/directline/conversations`, {
            method: 'POST',
            headers: { Authorization: `Bearer ${this.dlToken}` }
        });
        if (!res.ok) {
            const txt = await res.text().catch(() => '');
            throw new Error(`Failed to start Direct Line conversation (${res.status}) ${txt}`);
        }
        return res.json();
    }

    renderActivity(activity) {
        if (!activity) {
            return;
        }

        if (this.tryHandleOAuthAttachment(activity)) return;
        if (this.tryHandleAgentTyping(activity)) return;
        if (this.tryHandleSigninAction(activity)) return;
        this.tryHandleMessage(activity);
    }

    tryHandleOAuthAttachment(activity) {
        try {
            const oauthAttachment = (activity.attachments || []).find(a =>
                (a?.contentType || '').toLowerCase() === 'application/vnd.microsoft.card.oauth'
            );
            if (!oauthAttachment?.content) {
                return false;
            }

            const key = oauthAttachment.content.connectionName
                || oauthAttachment.content.connection?.name
                || oauthAttachment.content.tokenExchangeResource?.id
                || activity.id;
            if (!key || !this._oauthHandled.has(key)) {
                if (key) {
                    this._oauthHandled.add(key);
                }
                this.handleOAuthCard(activity, oauthAttachment.content).catch(() => {});
            }
            return true;
        } catch {
            return false;
        }
    }

    tryHandleAgentTyping(activity) {
        if (activity.type === 'typing' && (activity.from?.role || '').toLowerCase() !== 'user') {
            this.setAgentTyping(true);
            return true;
        }
        return false;
    }

    tryHandleSigninAction(activity) {
        try {
            const actions = activity.suggestedActions?.actions || [];
            const signin = actions.find(a => (a?.type || '').toLowerCase() === 'signin');
            if (!signin?.value) {
                return false;
            }

            const key = signin.connectionName
                || signin.connection?.name
                || signin.value
                || activity.id;
            if (!key || !this._oauthHandled.has(key)) {
                if (key) {
                    this._oauthHandled.add(key);
                }
                this.handleSigninSuggestedAction(activity, signin).catch(() => {});
            }
            return true;
        } catch {
            return false;
        }
    }

    tryHandleMessage(activity) {
        if (activity.type !== 'message') {
            return false;
        }

        const messageData = this.buildMessageFromActivity(activity);
        if (!this.replaceOrAppendFromActivity(activity, () => messageData)) {
            this.pushTranscriptEntry(messageData);
        }
        if (messageData.from !== 'user') {
            this.setAgentTyping(false);
        }
        return true;
    }

    handleRestart() {
        this.clearChatTranscript();
        this.resetDirectLineState({ preserveTranscript: true });
        this.startDirectLineChat();
    }

    handleTranscriptAction(event) {
        const messageId = event.currentTarget.dataset.messageId;
        const actionId = event.currentTarget.dataset.actionId;
        if (!messageId || !actionId) {
            return;
        }
        const message = this.transcript.find(entry => entry.id === messageId);
        if (!message) {
            return;
        }
        const action = (message.actions || []).find(a => a.id === actionId);
        if (!action) {
            return;
        }
        try {
            if (action.type === 'openUrl' && action.value) {
                window.open(action.value, '_blank', 'noopener,noreferrer');
            }
        } catch (e) {
            this.logDebug('Transcript action failed', e);
        }
    }

    setAgentTyping(active, ttl = 5000) {
        if (active) {
            this.isAgentTyping = true;
            this.clearTypingTimeout();
            this._typingTimeout = setTimeout(() => {
                this.isAgentTyping = false;
                this.clearTypingTimeout();
            }, ttl);
            return;
        }
        this.isAgentTyping = false;
        this.clearTypingTimeout();
    }

    clearTypingTimeout() {
        if (this._typingTimeout) {
            clearTimeout(this._typingTimeout);
            this._typingTimeout = null;
        }
    }

    clearChatTranscript() {
        this.transcript = [];
    }

    disposeDirectLineClient(client) {
        if (!client) {
            return;
        }
        try {
            if (typeof client.dispose === 'function') {
                client.dispose();
            } else if (typeof client.close === 'function') {
                client.close();
            } else if (typeof client.shutdown === 'function') {
                client.shutdown();
            }
        } catch (e) {
            this.logDebug('Direct Line client cleanup failed', e);
        }
    }

    unsubscribeFromActivities() {
        if (this._activitySubscription) {
            try {
                this._activitySubscription.unsubscribe?.();
            } catch {}
            this._activitySubscription = null;
        }
    }

    stopPollingActivities() {
        if (this.pollAbort) {
            try {
                this.pollAbort.abort();
            } catch {}
            this.pollAbort = null;
        }
        if (this._pollTimeoutId) {
            clearTimeout(this._pollTimeoutId);
            this._pollTimeoutId = null;
        }
    }

    markChatSessionReady(sessionToken) {
        if (sessionToken && sessionToken !== this._dlSessionToken) {
            return;
        }
        this._dlStarted = true;
        this._oauthHandled = new Set();
        this.setAgentTyping(false);
        this.iframeLoading = false;
    }

    resetDirectLineState({ preserveTranscript = false } = {}) {
        this._dlSessionToken += 1;
        this._dlStarted = false;
        this.chatConnecting = false;
        this.stopPollingActivities();
        this.unsubscribeFromActivities();
        this.disposeDirectLineClient(this.dlClient);
        this.dlClient = null;
        this.conversationId = null;
        this.dlToken = null;
        this.watermark = null;
        this._oauthHandled = new Set();
        if (!preserveTranscript) {
            this.clearChatTranscript();
        }
        this.setAgentTyping(false);
        this.dlDomain = 'https://directline.botframework.com';
        if (this._scrollFrame) {
            cancelAnimationFrame(this._scrollFrame);
            this._scrollFrame = null;
        }
    }

    getDirectLineContext() {
        const { environmentId, botId } = this._agentIds || {};
        if (!this._aadToken || !environmentId || !botId || !window.MicrosoftAgents) {
            this.logDebug('DL start skipped (prereqs missing)', {
                hasToken: !!this._aadToken,
                environmentId,
                botId,
                hasAgents: !!window.MicrosoftAgents
            });
            return null;
        }
        return { environmentId, botId };
    }

    createCopilotClient({ environmentId, botId }) {
        return new window.MicrosoftAgents.CopilotStudioClient(
            {
                tenantId: TENANT_ID,
                environmentId,
                agentIdentifier: botId,
                appClientId: CLIENT_ID,
                authority: `https://login.microsoftonline.com/${TENANT_ID}`
            },
            this._aadToken
        );
    }

    tryAdoptStreamingConnection(connection, sessionToken) {
        if (!connection || typeof connection.postActivity !== 'function' || !('activity$' in connection)) {
            return false;
        }
        if (sessionToken !== this._dlSessionToken) {
            this.disposeDirectLineClient(connection);
            return true;
        }
        this.dlClient = connection;
        this.unsubscribeFromActivities();
        try {
            this._activitySubscription = this.dlClient.activity$.subscribe(activity => this.renderActivity(activity));
        } catch {
            this.logDebug('Direct Line stream subscription failed');
        }
        this.markChatSessionReady(sessionToken);
        this.logDebug('Direct Line streaming client established');
        return true;
    }

    async adoptRestConnection(connection, sessionToken) {
        if (sessionToken !== this._dlSessionToken) {
            this.disposeDirectLineClient(connection);
            return;
        }
        this.dlToken = connection?.token || connection?.result?.token || null;
        const domain = connection?.domain || connection?.result?.domain;
        const sanitizedDomain = this.sanitizeDirectLineDomain(domain);
        if (sanitizedDomain) {
            this.dlDomain = sanitizedDomain;
        } else if (domain) {
            this.logDebug('Direct Line domain rejected (unsafe)', domain);
        }
        if (!this.dlToken) {
            throw new Error('Direct Line token not returned');
        }

        const json = await this.startDirectLineConversation();
        if (sessionToken !== this._dlSessionToken) {
            this.dlToken = null;
            return;
        }
        this.conversationId = json.conversationId;
        this.watermark = null;
        this.markChatSessionReady(sessionToken);
        this.startPollingActivities();
        this.logDebug('Direct Line REST client established', { domain: this.dlDomain });
    }

    startPollingActivities() {
        this.stopPollingActivities();

        const poll = async () => {
            if (!this._dlStarted) {
                return;
            }

            const controller = new AbortController();
            this.pollAbort = controller;

            try {
                if (!this.conversationId || !this.dlToken) {
                    return;
                }

                const url = new URL(`${this.dlDomain}/v3/directline/conversations/${this.conversationId}/activities`);
                if (this.watermark) {
                    url.searchParams.set('watermark', this.watermark);
                }

                const res = await fetch(url, {
                    headers: { Authorization: `Bearer ${this.dlToken}` },
                    signal: controller.signal
                });
                if (!res.ok) {
                    throw new Error(`DL poll failed ${res.status}`);
                }

                const json = await res.json();
                this.watermark = json.watermark || this.watermark;
                (json.activities || []).forEach(activity => this.renderActivity(activity));
            } catch {
                await new Promise(resolve => setTimeout(resolve, 1200));
            } finally {
                if (this.pollAbort === controller) {
                    this.pollAbort = null;
                }
                if (this._dlStarted) {
                    this._pollTimeoutId = setTimeout(poll, 800);
                }
            }
        };

        poll();
    }

    async handleOAuthCard(activity, oauthContent) {
        // Attempt to satisfy OAuthPrompt by handing the existing MSAL token to the bot via token exchange/events.
        // If it fails, render a visible sign-in button as fallback.
        const connectionName = oauthContent.connectionName || oauthContent.connection?.name;
        const exchangeId = oauthContent.tokenExchangeResource?.id;
        const text = oauthContent.text || 'Sign in is required.';

        // If we have no AAD token, we cannot silently complete; show button.
        if (!this._aadToken || !connectionName) {
            this.renderOAuthFallback(text, oauthContent?.buttons);
            return;
        }

        const payloads = [];
        payloads.push({
            type: 'invoke',
            name: 'signin/tokenExchange',
            value: { id: exchangeId, connectionName, token: this._aadToken },
            from: { id: 'user', name: this.accountLabel || 'You' }
        });
        payloads.push({
            type: 'event',
            name: 'tokens/response',
            value: { connectionName, token: this._aadToken },
            from: { id: 'user', name: this.accountLabel || 'You' }
        });

        let succeeded = false;
        try {
            for (const p of payloads) {
                await this.postActivity(p);
            }
            succeeded = true;
            // eslint-disable-next-line no-console
            console.info('[OAuth][SSO] token handed off via Direct Line');
        } catch (e) {
            // eslint-disable-next-line no-console
            console.debug && console.debug('[OAuth][SSO] handoff failed', e);
        }

        if (!succeeded) {
            this.renderOAuthFallback(text, oauthContent?.buttons);
        }
    }

    async handleSigninSuggestedAction(activity, signin) {
        // Try silent completion by sending tokens/response and tokenExchange first.
        const text = activity.text || 'Sign in is required.';
        if (!this._aadToken) { this.renderSigninFallback(text, signin); return; }

        const connectionName = signin.connectionName || signin.connection?.name;
        const payloads = [];
        if (!connectionName) {
            this.renderSigninFallback(text, signin);
            return;
        }

        payloads.push({ type: 'event', name: 'tokens/response', value: { connectionName, token: this._aadToken }, from: { id: 'user', name: this.accountLabel || 'You' } });
        payloads.push({ type: 'invoke', name: 'signin/tokenExchange', value: { connectionName, token: this._aadToken }, from: { id: 'user', name: this.accountLabel || 'You' } });

        let ok = false;
        try { for (const p of payloads) { await this.postActivity(p); } ok = true; } catch {}
        if (!ok) this.renderSigninFallback(text, signin);
    }

    renderOAuthFallback(text, buttons = []) {
        this.setAgentTyping(false);
        const btn = (buttons || []).find(b => (b?.type || '').toLowerCase() === 'signin');
        if (btn?.value && this.isSafeUrl(btn.value)) {
            this.pushTranscriptEntry({
                from: 'agent',
                text,
                actions: [
                    {
                        id: this.generateClientActivityId(),
                        type: 'openUrl',
                        title: btn.title || 'Sign in',
                        value: btn.value
                    }
                ]
            });
        } else {
            if (btn?.value) {
                this.logDebug('OAuth fallback URL blocked (unsafe scheme)', btn.value);
            }
            this.pushTranscriptEntry({ from: 'agent', text });
        }
    }

    renderSigninFallback(text, signin) {
        this.setAgentTyping(false);
        if (signin?.value && this.isSafeUrl(signin.value)) {
            this.pushTranscriptEntry({
                from: 'agent',
                text,
                actions: [
                    {
                        id: this.generateClientActivityId(),
                        type: 'openUrl',
                        title: signin.title || 'Sign in',
                        value: signin.value
                    }
                ]
            });
        } else {
            if (signin?.value) {
                this.logDebug('Signin fallback URL blocked (unsafe scheme)', signin.value);
            }
            this.pushTranscriptEntry({ from: 'agent', text });
        }
    }

    isSafeUrl(url) {
        try {
            const parsed = new URL(url, window.location.href);
            return parsed.protocol === 'https:';
        } catch {
            return false;
        }
    }

    async postActivity(activity) {
        // Prefer streaming client when available
        if (this.dlClient && typeof this.dlClient.postActivity === 'function') {
            await new Promise((resolve, reject) => {
                const sub = this.dlClient.postActivity(activity).subscribe({
                    next: () => {},
                    error: err => { try { sub.unsubscribe?.(); } catch {} reject(err); },
                    complete: () => { try { sub.unsubscribe?.(); } catch {} resolve(); }
                });
            });
            return;
        }
        // REST fallback
        if (!this.conversationId || !this.dlToken) throw new Error('DL not ready');
        const res = await fetch(`${this.dlDomain}/v3/directline/conversations/${this.conversationId}/activities`, {
            method: 'POST',
            headers: { Authorization: `Bearer ${this.dlToken}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(activity)
        });
        if (!res.ok) throw new Error(`Failed to post activity ${res.status}`);
    }

    onKeydown(e) {
        if (e.key === 'Enter') this.send();
    }

    async send() {
        const input = this.template.querySelector('[data-id="msg"]');
        if (!input) return;
        const text = (input.value || '').trim();
        if (!text) return;
        if (!this._dlStarted) {
            this.logDebug('Send skipped: chat not ready');
            return;
        }

        input.value = '';

        const clientActivityId = this.generateClientActivityId();
        this.pushTranscriptEntry({
            id: clientActivityId,
            key: clientActivityId,
            from: 'user',
            text,
            status: 'pending',
            clientActivityId
        });
        this.setAgentTyping(true, 12000);

        const activityPayload = {
            type: 'message',
            from: { id: 'user', name: this.accountLabel || 'You' },
            text,
            channelData: { clientActivityID: clientActivityId }
        };

        try {
            await this.postActivity(activityPayload);
            this.markTranscriptEntryStatus(clientActivityId, 'sent');
        } catch (err) {
            this.markTranscriptEntryStatus(clientActivityId, 'failed');
            this.setAgentTyping(false);
            this.setError(err);
        }
    }

    disconnectedCallback() {
        if (this._hostObserver) {
            try { this._hostObserver.disconnect(); } catch {}
            this._hostObserver = null;
        }
        this.resetDirectLineState();
    }

    async ensureAadToken() {
        // Returns valid token or null; never interactive.
        if (this._aadToken && this._aadTokenExpires instanceof Date) {
            const remainingMs = this._aadTokenExpires.getTime() - Date.now();
            if (remainingMs > 60000) {
                return this._aadToken;
            }
        }
        try {
            await this.refreshAadToken();
            return this._aadToken || null;
        } catch {
            return this._aadToken || null;
        }
    }

    // Retrieve a Chat.Invoke token if the agent requires downstream custom API calls
    async getCustomApiToken() {
        const account = msalInstance?.getActiveAccount();
        if (!account) return null;
        const scopes = this.customApiLoginScopes;
        if (!scopes.length) return null;
        try {
            const resp = await msalInstance.acquireTokenSilent({ scopes, account });
            return resp?.accessToken || null;
        } catch (e) {
            const diag = MsalDiag.extract(e);
            if (MsalDiag.isInteraction(diag)) {
                try {
                    const popupResp = await msalInstance.acquireTokenPopup({ scopes, account });
                    return popupResp?.accessToken || null;
                } catch (popupErr) {
                    this.logDebug('getCustomApiToken popup failed', MsalDiag.extract(popupErr));
                }
            }
            this.logDebug('getCustomApiToken failed', diag);
            return null;
        }
    }

    async handleLogout() {
        try {
            const acct = msalInstance?.getActiveAccount();
            if (!acct) {
                this.clearSession();
                return;
            }
            if (window.self === window.top) {
                await msalInstance.logoutRedirect({
                    account: acct,
                    postLogoutRedirectUri: this.effectiveRedirectUri
                });
                return;
            }
            await msalInstance.logoutPopup({ account: acct }).catch(() => {});
        } catch (e) {
            this.setError(e);
        }
    }

    clearSession() {
        this.signedIn = false;
        this.accountLabel = '';
        this.dispatchEvent(new CustomEvent('lightningcopilotauthsignout'));
        this.resetDirectLineState();
        this._aadToken = null;
        this._aadTokenExpires = null;
        this.showLoginButton = false;
        this.statusMessage = 'Checking sign-in...';
        this._authCheckComplete = false;
    }

    handleIframeLoad() {
        this.iframeLoading = false;
        this.postIframeIdentity();
    }

    handleIframeError() {
        this.setError(new Error('Failed to load Lightning Copilot embed content.'));
    }

    postIframeIdentity() {
        if (!this.signedIn || !this.parsedEmbedOrigin) return;
        const iframe = this.template.querySelector('iframe.copilot-embed');
        if (!iframe) return;
        try {
            iframe.contentWindow?.postMessage(
                { type: 'copilotIdentity', user: this.accountLabel },
                this.parsedEmbedOrigin
            );
        } catch (e) {
            console.info('[LightningCopilotAuth] postMessage identity failed (non-fatal).', e); // eslint-disable-line no-console
        }
    }

    setError(e) {
        const diag = {
            name: e?.name,
            code: e?.errorCode || e?.code,
            subError: e?.subError,
            status: e?.status || e?.response?.status,
            message: e?.errorMessage || e?.message
        };
        const parts = Object.values(diag).filter(Boolean);
        this.error = parts.join(' | ');
        this.iframeLoading = false; // ensure UI shows error instead of spinner
        this.isAgentTyping = false;
        clearTimeout(this._typingTimeout);
        this._typingTimeout = null;
        console.groupCollapsed('[LightningCopilotAuth] Error detail');
        console.error('Diag:', diag);
        console.error('Raw:', e);
        console.groupEnd();
        this.dispatchEvent(new CustomEvent('lightningcopilotautherror', { detail: { message: this.error, diag } }));
    }

    @api
    setEmbedUrlOverride(raw) {
        // Reject overrides to non-allowlisted hosts
        if (raw) {
            try {
                const u = new URL(raw.trim(), window.location.href);
                if (!ALLOWED_EMBED_HOSTS.includes(u.host)) {
                    this.setError(new Error('Override host not permitted.'));
                    return;
                }
            } catch {
                this.setError(new Error('Invalid override URL.'));
                return;
            }
        }
        this.logDebug('Manual embed override', raw);
        this._manualOverride = raw;
        this.validateEmbedLabel();
    }

    ensureHostObserver() {
        if (this._hostObserver) return;
        this._hostObserver = new MutationObserver(() => {
            if (!this._dlStarted) this.startDirectLineChat();
        });
        this._hostObserver.observe(this.template, { childList: true, subtree: true });
    }
}
