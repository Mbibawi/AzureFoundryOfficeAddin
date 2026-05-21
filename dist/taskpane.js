import { FoundryGateway as fg } from './FoundryGateway.js';
import { WordApp, ExcelApp, PowerPointApp } from './OfficeApp.js';
export function byID(id) {
    return document.getElementById(id);
}
export function el(tag, attr) {
    const e = document.createElement(tag);
    if (!attr)
        return e;
    if (attr.className)
        e.className = attr.className;
    if (attr.innerHTML)
        e.innerHTML = attr.innerHTML;
    if (attr.innerText)
        e.innerText = attr.innerText;
    if (attr.textContent)
        e.textContent = attr.textContent;
    for (const k in attr)
        e.setAttribute(k, attr[k]);
    return e;
}
class UI {
    constructor() {
        this.STORAGE_KEY = 'foundry_addin_config';
        this.fetchModels = (cfg) => fg.fetchModels(cfg);
        this.streamCompletion = (cfg, modelId, messages) => fg.streamCompletion(cfg, modelId, messages);
        this.config = null;
        this.models = [];
        this.chatHistory = [];
        this.streaming = false;
        this.stopFlag = false;
        this.docText = '';
        this.getDocumentText = () => this.officeApp.getDocumentText();
        const host = Office?.context?.host;
        if (host === Office.HostType.Excel)
            this.officeApp = new ExcelApp();
        else if (host === Office.HostType.PowerPoint)
            this.officeApp = new PowerPointApp();
        else
            this.officeApp = new WordApp();
    }
    saveConfig(c) {
        try {
            localStorage.setItem(this.STORAGE_KEY, JSON.stringify(c));
        }
        catch { }
    }
    loadConfig() {
        try {
            const raw = localStorage.getItem(this.STORAGE_KEY);
            if (raw)
                return JSON.parse(raw);
        }
        catch { }
        return null;
    }
    clearConfig() {
        try {
            localStorage.removeItem(this.STORAGE_KEY);
        }
        catch { }
    }
    getModelMeta(id) {
        const l = id.toLowerCase();
        if (l.includes('claude'))
            return { icon: '🟠', colorClass: 'model-anthropic', provider: 'Anthropic / Claude' };
        if (l.includes('gpt') || l.includes('o1') || l.includes('o3') || l.includes('o4'))
            return { icon: '🟢', colorClass: 'model-openai', provider: 'OpenAI / Azure' };
        if (l.includes('llama') || l.includes('meta'))
            return { icon: '🔵', colorClass: 'model-meta', provider: 'Meta / Llama' };
        if (l.includes('mistral') || l.includes('mixtral'))
            return { icon: '🟣', colorClass: 'model-mistral', provider: 'Mistral AI' };
        if (l.includes('phi'))
            return { icon: '🔷', colorClass: 'model-openai', provider: 'Microsoft / Phi' };
        if (l.includes('deepseek'))
            return { icon: '🐋', colorClass: 'model-other', provider: 'DeepSeek' };
        if (l.includes('cohere'))
            return { icon: '🔶', colorClass: 'model-other', provider: 'Cohere' };
        return { icon: '🤖', colorClass: 'model-other', provider: 'Unknown provider' };
    }
    renderHeader(title, badge = '', showClearChat = false) {
        const header = el('div', { className: "header" });
        header.append(el('div', { className: "header-logo" }), el("span", { className: "header-title", inneText: title }));
        if (badge)
            header.appendChild(el("span", { className: "badge", innerText: badge }));
        if (showClearChat) {
            const btn = el("button", { className: "btn btn-ghost ml-auto", id: "btn-clear-chat", title: "Clear conversation", });
            btn.innerHTML = ` <svg width="14" height = "14" viewBox = "0 0 24 24" fill = "none" stroke = "currentColor" stroke - width="2" stroke - linecap="round" stroke - linejoin="round" >
      <path d="M3 6h18" /> <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6" /> <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2" />
      </svg>`;
            header.appendChild(btn);
        }
        return header;
    }
    renderSetupView() {
        const setup = el("div", { className: "setup" });
        const content = el("div", { id: "screen-setup-content" });
        content.append(el("div", { className: "setup-intro" }), el("div", {
            className: "field", innerHTML: `<label>Resource Name</label>
            <input type="text" id="input-resource" placeholder="my-foundry-resource" autocomplete="off" spellcheck="false" />
            <span class="field-hint">The short name of your Foundry resource — e.g. <code>my-foundry-resource</code> from <code>my-foundry-resource.services.ai.azure.com</code></span>`
        }), el("div", {
            className: "field", innerHTML: `<label>Project Name</label>
  <input type="text" id="input-project" placeholder="azure-foundry-claude-api-gateway" autocomplete="off" />
  <span class="field-hint">The project name under your Foundry resource</span>`
        }), el("div", {
            className: "field", innerHTML: `<label>API Key</label>
            <input type="password" id="input-apikey" placeholder="••••••••••••••••" autocomplete="off" />
            <span class="field-hint">Found in Azure Portal → your Foundry resource → <code>Keys and Endpoint</code> → KEY 1</span>`
        }), el("div", { id: "setup-error", className: "alert alert-error hidden" }), el("button", {
            className: "btn btn-primary btn-full", id: "btn-connect",
            innerHTML: `<span id="btn-connect-label">Connect to Foundry</span>
            <span id="btn-connect-spinner" class="spinner hidden"></span>`
        }));
        setup.append(this.renderHeader('Foundry AI', 'SETUP'), content);
        return setup;
    }
    renderMainView() {
        const screen = el('div', { id: 'screen-main' });
        const mainTop = el('div', { className: 'main-top' });
        mainTop.innerHTML = `
      <div class="flex-row">
        <span class="section-label" style="flex-shrink:0;">Model</span>
        <div class="model-selector" style="flex:1;">
          <select id="model-select">
            <option value="">Loading models…</option>
          </select>
          <button class="btn btn-ghost" id="btn-refresh-models" title="Refresh models">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 2v6h-6"/><path d="M3 12a9 9 0 0 1 15-6.7L21 8"/><path d="M3 22v-6h6"/><path d="M21 12a9 9 0 0 1-15 6.7L3 16"/>
            </svg>
          </button>
        </div>
      </div>
      <div id="model-card" class="model-card">
        <div id="model-card-icon" class="model-card-icon model-other">🤖</div>
        <div class="model-card-info">
          <div id="model-card-name" class="model-card-name">No model selected</div>
          <div id="model-card-provider" class="model-card-provider">—</div>
        </div>
      </div>`;
        const chatArea = el('div', { className: 'chat-area', id: 'chat-area' });
        const emptyState = el('div', { className: 'empty-state', id: 'empty-state' });
        emptyState.innerHTML = `
      <div class="empty-state-icon">⬡</div>
      <div class="empty-state-title">Ready to assist</div>
      <div class="empty-state-sub">Select a model above, then ask a question or analyse your document.</div>`;
        const quickPrompts = el('div', { className: 'quick-prompts' });
        this.officeApp.getQuickPrompts().forEach(p => {
            const btn = el('button', { className: 'quick-prompt-btn' });
            btn.setAttribute('data-prompt', p.prompt);
            btn.innerHTML = `<span class="qp-icon">${p.icon}</span> ${p.label}`;
            quickPrompts.appendChild(btn);
        });
        emptyState.appendChild(quickPrompts);
        chatArea.appendChild(emptyState);
        const inputArea = el('div', { className: 'input-area' });
        inputArea.innerHTML = `
      <label class="doc-context-toggle" for="chk-include-doc">
        <input type="checkbox" id="chk-include-doc" checked />
        <span class="doc-context-toggle-label">
          <strong>Include ${this.officeApp.getHostLabel()}</strong>
        </span>
        <span class="doc-word-count" id="doc-word-count"></span>
      </label>
      <div class="input-row">
        <textarea id="user-input" placeholder="Ask anything…" rows="1"></textarea>
        <button class="send-btn" id="btn-send" disabled>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
          </svg>
        </button>
        <button class="stop-btn hidden" id="btn-stop">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor">
            <rect x="4" y="4" width="16" height="16" rx="2"/>
          </svg>
        </button>
      </div>`;
        const footer = el('div', { className: 'settings-footer' });
        footer.innerHTML = `
      <div class="status-dot connected" id="footer-status-dot"></div>
      <span class="conn-name" id="footer-conn-name"></span>
      <button class="btn btn-ghost" id="btn-disconnect">Disconnect</button>`;
        screen.append(this.renderHeader('Foundry AI', '', true), mainTop, chatArea, inputArea, footer);
        return screen;
    }
    showScreen(name) {
        const app = byID('app');
        app.innerHTML = '';
        if (name === 'setup') {
            app.appendChild(this.renderSetupView());
            this.wireSetupEvents();
        }
        else {
            app.appendChild(this.renderMainView());
            this.wireMainEvents();
            if (this.chatHistory.length > 0) {
                this.chatHistory.forEach(msg => {
                    if (msg.role !== 'system') {
                        const text = msg.content.includes('Here is the document content:')
                            ? msg.content.split('---\n\n')[1] || msg.content
                            : msg.content;
                        this.appendMessage(msg.role, text);
                    }
                });
            }
            this.populateModelSelect();
            this.updateWordCount();
            if (this.config) {
                byID('footer-conn-name').textContent = `${this.config.resource}.services.ai.azure.com`;
            }
        }
    }
    populateModelSelect() {
        const sel = byID('model-select');
        sel.innerHTML = '';
        if (this.models.length === 0) {
            sel.innerHTML = '<option value="">No models found</option>';
            return;
        }
        this.models.forEach((m) => {
            const opt = document.createElement('option');
            opt.value = m.id;
            opt.textContent = m.name;
            sel.appendChild(opt);
        });
        this.updateModelCard(sel.value);
    }
    updateModelCard(modelId) {
        const meta = this.getModelMeta(modelId);
        const icon = byID('model-card-icon');
        icon.textContent = meta.icon;
        icon.className = `model-card-icon ${meta.colorClass}`;
        byID('model-card-name').textContent = modelId || 'No model selected';
        byID('model-card-provider').textContent = meta.provider;
    }
    appendMessage(role, text) {
        this.officeApp.hide('empty-state');
        const wrap = el('div', { className: 'message' });
        const label = el('div', { className: `message-role ${role === 'user' ? 'user-role' : ''}`, textContent: role === 'user' ? '▸ You' : '▸ Assistant' });
        const bubble = el('div', { className: `message-bubble ${role === 'user' ? 'user-bubble' : ''}`, textContent: text });
        wrap.appendChild(label);
        wrap.appendChild(bubble);
        byID('chat-area').appendChild(wrap);
        this.scrollChat();
        return bubble;
    }
    scrollChat() {
        const area = byID('chat-area');
        area.scrollTop = area.scrollHeight;
    }
    clearChat() {
        this.chatHistory = [];
        const area = byID('chat-area');
        const kids = Array.from(area.children);
        kids.forEach((k) => {
            if (k.id !== 'empty-state')
                k.remove();
        });
        this.officeApp.show('empty-state');
    }
    updateWordCount() {
        const wc = byID('doc-word-count');
        if (!this.docText) {
            wc.textContent = '';
            return;
        }
        const words = this.docText.trim().split(/\s+/).filter(Boolean).length;
        wc.textContent = `${words.toLocaleString()} words`;
    }
    async sendMessage() {
        if (this.streaming || !this.config)
            return;
        const input = byID('user-input');
        const prompt = input.value.trim();
        if (!prompt)
            return;
        const includeDoc = (byID('chk-include-doc')).checked;
        const contextData = includeDoc ? await this.officeApp.getActiveContext() : null;
        if (this.chatHistory.length === 0) {
            const systemMsg = `You are a specialized AI assistant for Microsoft ${this.officeApp.getHostLabel()}. `
                + `Help the user with their document by answering questions, suggesting edits, and explaining your reasoning clearly.`;
            this.chatHistory.push({ role: 'system', content: systemMsg });
        }
        let userContent = prompt;
        if (contextData) {
            const header = contextData.source === 'selection' ? 'SELECTED TEXT' : 'FULL DOCUMENT CONTENT';
            userContent = `--- ${header} ---\n${contextData.text}\n---\n\nUSER REQUEST: ${prompt}`;
        }
        this.chatHistory.push({ role: 'user', content: userContent });
        this.appendMessage('user', prompt);
        input.value = '';
        this.resizeTextarea(input);
        byID('btn-send').disabled = true;
        const bubble = this.appendMessage('assistant', '');
        await this.handleStreamingResponse(bubble);
    }
    async handleStreamingResponse(bubble) {
        this.streaming = true;
        this.stopFlag = false;
        byID('btn-send').disabled = true;
        byID('btn-stop').classList.remove('hidden');
        let fullText = '';
        try {
            const modelId = byID('model-select').value;
            if (!modelId)
                throw new Error('No model selected. Please choose a model from the dropdown.');
            const gen = this.streamCompletion(this.config, modelId, this.chatHistory);
            for await (const token of gen) {
                if (this.stopFlag)
                    break;
                fullText += token;
                bubble.textContent = fullText;
                this.scrollChat();
            }
            if (fullText) {
                this.chatHistory.push({ role: 'assistant', content: fullText });
            }
            else if (!this.stopFlag) {
                bubble.textContent = '(No response received.)';
            }
        }
        catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            bubble.textContent = `Error: ${msg}`;
        }
        finally {
            this.streaming = false;
            this.stopFlag = false;
            byID('btn-stop').classList.add('hidden');
            const userInput = byID('user-input');
            byID('btn-send').disabled = userInput.value.trim().length === 0;
        }
    }
    showAlert(containerId, msg, type) {
        const container = byID(containerId);
        const div = document.createElement('div');
        div.className = `alert alert-${type}`;
        div.textContent = msg;
        container.appendChild(div);
        setTimeout(() => div.remove(), 6000);
    }
    setSetupError(msg) {
        const el = byID('setup-error');
        el.textContent = msg;
        el.classList.remove('hidden');
    }
    clearSetupError() {
        byID('setup-error').classList.add('hidden');
    }
    async connect(resource, project, apiKey) {
        this.clearSetupError();
        const btnLabel = byID('btn-connect-label');
        const btnSpinner = byID('btn-connect-spinner');
        const btn = byID('btn-connect');
        btnLabel.textContent = 'Connecting…';
        btnSpinner.classList.remove('hidden');
        btn.disabled = true;
        try {
            const cfg = {
                resource,
                project,
                apiKey
            };
            const found = await this.fetchModels(cfg);
            if (found.length === 0) {
                throw new Error('No deployed models found on this resource. Deploy at least one model in Azure AI Foundry first.');
            }
            this.config = cfg;
            this.models = found;
            this.saveConfig(cfg);
            this.enterMain();
        }
        catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            this.setSetupError(msg);
        }
        finally {
            btnLabel.textContent = 'Connect to Foundry';
            btnSpinner.classList.add('hidden');
            btn.disabled = false;
        }
    }
    async enterMain() {
        if (!this.config)
            return;
        this.showScreen('main');
        this.populateModelSelect();
        byID('footer-conn-name').textContent = `${this.config.resource}.services.ai.azure.com`;
        this.docText = await this.getDocumentText();
        this.updateWordCount();
    }
    disconnect() {
        this.config = null;
        this.models = [];
        this.chatHistory = [];
        this.docText = '';
        this.clearConfig();
        this.clearChat();
        this.showScreen('setup');
    }
    resizeTextarea(el) {
        el.style.height = 'auto';
        el.style.height = `${Math.min(el.scrollHeight, 120)}px`;
    }
    wireSetupEvents() {
        const btnConnect = byID('btn-connect');
        if (!btnConnect)
            return;
        btnConnect.addEventListener('click', () => {
            const resource = (byID('input-resource')).value.trim();
            const project = (byID('input-project')).value.trim();
            const apiKey = (byID('input-apikey')).value.trim();
            if (!resource) {
                this.setSetupError('Please enter your Foundry resource name.');
                return;
            }
            if (!project) {
                this.setSetupError('Please enter your Foundry resource name.');
                return;
            }
            if (!apiKey) {
                this.setSetupError('Please enter your API key.');
                return;
            }
            this.connect(resource, project, apiKey);
        });
        const inputs = [
            byID('input-resource'),
            byID('input-apikey')
        ].filter(Boolean);
        inputs.forEach((inp) => {
            inp.addEventListener('keydown', (e) => {
                if (e.key === 'Enter')
                    byID('btn-connect').click();
            });
        });
    }
    wireEvents() {
    }
    wireMainEvents() {
        const modelSelect = byID('model-select');
        if (!modelSelect)
            return;
        modelSelect.addEventListener('change', (e) => {
            this.updateModelCard(e.target.value);
        });
        byID('btn-refresh-models').addEventListener('click', async () => {
            if (!this.config)
                return;
            const btn = byID('btn-refresh-models');
            btn.disabled = true;
            try {
                const found = await this.fetchModels(this.config);
                this.models = found;
                this.populateModelSelect();
            }
            catch (err) {
                const msg = err instanceof Error ? err.message : String(err);
                this.showAlert('chat-area', `Refresh failed: ${msg}`, 'error');
            }
            finally {
                btn.disabled = false;
            }
        });
        byID('btn-send').addEventListener('click', () => this.sendMessage());
        byID('btn-stop').addEventListener('click', () => { this.stopFlag = true; });
        const userInput = byID('user-input');
        userInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                this.sendMessage();
            }
        });
        userInput.addEventListener('input', () => {
            this.resizeTextarea(userInput);
            byID('btn-send').disabled = userInput.value.trim().length === 0;
        });
        byID('btn-clear-chat').addEventListener('click', () => this.clearChat());
        byID('btn-disconnect').addEventListener('click', () => this.disconnect());
        document.querySelectorAll('.quick-prompt-btn').forEach((btn) => {
            btn.addEventListener('click', () => {
                const prompt = btn.getAttribute('data-prompt') ?? '';
                userInput.value = prompt;
                this.resizeTextarea(userInput);
                byID('btn-send').disabled = false;
                userInput.focus();
            });
        });
        byID('chk-include-doc').addEventListener('change', async () => {
            if (this.docText === '') {
                this.docText = await this.getDocumentText();
                this.updateWordCount();
            }
        });
    }
    async init() {
        const saved = this.loadConfig();
        if (saved) {
            try {
                const found = await this.fetchModels(saved);
                if (found.length > 0) {
                    this.config = saved;
                    this.models = found;
                    await this.enterMain();
                    return;
                }
            }
            catch {
            }
        }
        this.showScreen('setup');
    }
}
const ui = new UI();
if (typeof Office !== 'undefined') {
    Office.onReady(() => ui.init());
}
else {
    document.addEventListener('DOMContentLoaded', () => ui.init());
}
//# sourceMappingURL=taskpane.js.map