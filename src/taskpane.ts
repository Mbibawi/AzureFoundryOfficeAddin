import { FoundryGateway as fg } from './FoundryGateway.js';
import { WordApp, ExcelApp, PowerPointApp } from './OfficeApp.js';

export function byID<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}


export function el<T extends HTMLElement>(tag: string, attr?: { [key: string]: string }) {
  const e = document.createElement(tag) as T;
  if (!attr) return e;
  if (attr.className) e.className = attr.className;
  if (attr.innerHTML) e.innerHTML = attr.innerHTML;
  if (attr.innerText) e.innerText = attr.innerText;
  if (attr.textContent) e.textContent = attr.textContent;
  for (const k in attr) e.setAttribute(k, attr[k]);
  return e;
}

class UI {
  private readonly STORAGE_KEY = 'foundry_addin_config';
  private readonly fetchModels = (cfg: Config) => fg.fetchModels(cfg);
  private readonly streamCompletion = (cfg: Config, modelId: string, messages: ChatMessage[]) => fg.streamCompletion(cfg, modelId, messages);
  private officeApp: WordApp | ExcelApp | PowerPointApp;
  private config: Config | null = null;
  private models: FoundryModel[] = [];
  private chatHistory: ChatMessage[] = [];
  private streaming: boolean = false;
  private stopFlag: boolean = false;
  private docText: string = '';
  private getDocumentText = () => this.officeApp.getDocumentText();


  constructor() {
    const host = Office?.context?.host;
    if (host === Office.HostType.Excel) this.officeApp = new ExcelApp();
    else if (host === Office.HostType.PowerPoint) this.officeApp = new PowerPointApp();
    else this.officeApp = new WordApp();
  }

  saveConfig(c: Config): void {
    try { localStorage.setItem(this.STORAGE_KEY, JSON.stringify(c)); } catch { }
  }

  loadConfig(): Config | null {
    try {
      const raw = localStorage.getItem(this.STORAGE_KEY);
      if (raw) return JSON.parse(raw) as Config;
    } catch { }
    return null;
  }

  clearConfig(): void {
    try { localStorage.removeItem(this.STORAGE_KEY); } catch { }
  }

  // ─────────────────────────────────────────────
  //  STATE
  // ─────────────────────────────────────────────



  // ─────────────────────────────────────────────
  //  STORAGE
  // ─────────────────────────────────────────────



  // ─────────────────────────────────────────────
  //  FOUNDRY API
  // ─────────────────────────────────────────────





  // ─────────────────────────────────────────────
  //  DOCUMENT READING (Office.js)
  // ─────────────────────────────────────────────


  // ─────────────────────────────────────────────
  //  MODEL HELPERS
  // ─────────────────────────────────────────────


  getModelMeta(id: string): ModelMeta {
    const l = id.toLowerCase();
    if (l.includes('claude')) return { icon: '🟠', colorClass: 'model-anthropic', provider: 'Anthropic / Claude' };
    if (l.includes('gpt') || l.includes('o1') || l.includes('o3') || l.includes('o4'))
      return { icon: '🟢', colorClass: 'model-openai', provider: 'OpenAI / Azure' };
    if (l.includes('llama') || l.includes('meta'))
      return { icon: '🔵', colorClass: 'model-meta', provider: 'Meta / Llama' };
    if (l.includes('mistral') || l.includes('mixtral'))
      return { icon: '🟣', colorClass: 'model-mistral', provider: 'Mistral AI' };
    if (l.includes('phi')) return { icon: '🔷', colorClass: 'model-openai', provider: 'Microsoft / Phi' };
    if (l.includes('deepseek')) return { icon: '🐋', colorClass: 'model-other', provider: 'DeepSeek' };
    if (l.includes('cohere')) return { icon: '🔶', colorClass: 'model-other', provider: 'Cohere' };
    return { icon: '🤖', colorClass: 'model-other', provider: 'Unknown provider' };
  }

  renderHeader(title: string, badge: string = '', showClearChat: boolean = false): HTMLElement {
    const header = el<HTMLDivElement>('div', { className: "header" });
    header.append(
      el<HTMLDivElement>('div', { className: "header-logo" }),
      el<HTMLSpanElement>("span", { className: "header-title", inneText: title }));

    if (badge) header.appendChild(el<HTMLSpanElement>("span", { className: "badge", innerText: badge }));

    if (showClearChat) {
      const btn = el<HTMLButtonElement>("button", { className: "btn btn-ghost ml-auto", id: "btn-clear-chat", title: "Clear conversation", });
      btn.innerHTML = ` <svg width="14" height = "14" viewBox = "0 0 24 24" fill = "none" stroke = "currentColor" stroke - width="2" stroke - linecap="round" stroke - linejoin="round" >
      <path d="M3 6h18" /> <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6" /> <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2" />
      </svg>`
      header.appendChild(btn);
    }


    return header
  }

  renderSetupView(): HTMLElement {
    const setup = el<HTMLDivElement>("div", { className: "setup" });
    const content = el("div", { id: "screen-setup-content" });
    content.append(
      el("div", { className: "setup-intro" }),
      el("div", {
        className: "field", innerHTML: `<label>Resource Name</label>
            <input type="text" id="input-resource" placeholder="my-foundry-resource" autocomplete="off" spellcheck="false" />
            <span class="field-hint">The short name of your Foundry resource — e.g. <code>my-foundry-resource</code> from <code>my-foundry-resource.services.ai.azure.com</code></span>`}),
      el("div", {
        className: "field", innerHTML: `<label>API Key</label>
            <input type="password" id="input-apikey" placeholder="••••••••••••••••" autocomplete="off" />
            <span class="field-hint">Found in Azure Portal → your Foundry resource → <code>Keys and Endpoint</code> → KEY 1</span>`}),
      el("div", { id: "setup-error", className: "alert alert-error hidden" }),
      el("button", {
        className: "btn btn-primary btn-full", id: "btn-connect",
        innerHTML: `<span id="btn-connect-label">Connect to Foundry</span>
            <span id="btn-connect-spinner" class="spinner hidden"></span>`})

    )
    setup.append(
      this.renderHeader('Foundry AI', 'SETUP'),
      content,

    );


    return setup;
  }

  // FIX 1: renderMainView now returns HTMLElement (consistent with renderSetupView and renderHeader)
  // so that renderHeader()'s HTMLElement return value is appended correctly, not coerced to "[object HTMLDivElement]"
  renderMainView(): HTMLElement {
    const screen = el<HTMLDivElement>('div', { id: 'screen-main' });

    // main-top: model selector + model card
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

    // chat area + quick prompts
    const chatArea = el('div', { className: 'chat-area', id: 'chat-area' });
    const emptyState = el('div', { className: 'empty-state', id: 'empty-state' });
    emptyState.innerHTML = `
      <div class="empty-state-icon">⬡</div>
      <div class="empty-state-title">Ready to assist</div>
      <div class="empty-state-sub">Select a model above, then ask a question or analyse your document.</div>`;
    const quickPrompts = el('div', { className: 'quick-prompts' });
    this.officeApp.getQuickPrompts().forEach(p => {
      const btn = el<HTMLButtonElement>('button', { className: 'quick-prompt-btn' });
      btn.setAttribute('data-prompt', p.prompt);
      btn.innerHTML = `<span class="qp-icon">${p.icon}</span> ${p.label}`;
      quickPrompts.appendChild(btn);
    });
    emptyState.appendChild(quickPrompts);
    chatArea.appendChild(emptyState);

    // input area
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

    // footer
    const footer = el('div', { className: 'settings-footer' });
    footer.innerHTML = `
      <div class="status-dot connected" id="footer-status-dot"></div>
      <span class="conn-name" id="footer-conn-name"></span>
      <button class="btn btn-ghost" id="btn-disconnect">Disconnect</button>`;

    screen.append(this.renderHeader('Foundry AI', '', true), mainTop, chatArea, inputArea, footer);
    return screen;
  }

  showScreen(name: 'setup' | 'main'): void {
    const app = byID<HTMLElement>('app');
    app.innerHTML = '';

    if (name === 'setup') {
      app.appendChild(this.renderSetupView());
      this.wireSetupEvents();
    } else {
      app.appendChild(this.renderMainView());
      this.wireMainEvents();

      // Restore state
      if (this.chatHistory.length > 0) {
        this.chatHistory.forEach(msg => {
          if (msg.role !== 'system') {
            const text = msg.content.includes('Here is the document content:')
              ? msg.content.split('---\n\n')[1] || msg.content
              : msg.content;
            this.appendMessage(msg.role as 'user' | 'assistant', text);
          }
        });
      }

      this.populateModelSelect();
      this.updateWordCount();
      if (this.config) {
        byID<HTMLElement>('footer-conn-name').textContent = `${this.config.resource}.services.ai.azure.com`;
      }
    }
  }


  populateModelSelect(): void {
    const sel = byID<HTMLSelectElement>('model-select');
    sel.innerHTML = '';

    if (this.models.length === 0) {
      sel.innerHTML = '<option value="">No models found</option>';
      return;
    }

    this.models.forEach((m) => {
      const opt = document.createElement('option');
      opt.value = m.id;
      opt.textContent = m.id;
      sel.appendChild(opt);
    });

    this.updateModelCard(sel.value);
  }

  updateModelCard(modelId: string): void {
    const meta = this.getModelMeta(modelId);
    const icon = byID<HTMLElement>('model-card-icon');
    icon.textContent = meta.icon;
    icon.className = `model-card-icon ${meta.colorClass}`;
    byID<HTMLElement>('model-card-name').textContent = modelId || 'No model selected';
    byID<HTMLElement>('model-card-provider').textContent = meta.provider;
  }

  appendMessage(role: 'user' | 'assistant', text: string): HTMLElement {
    this.officeApp.hide('empty-state');

    const wrap = el('div', {className:'message'});
  
    const label = el('div', {className: `message-role ${role === 'user' ? 'user-role' : ''}`, textContent: role === 'user' ? '▸ You' : '▸ Assistant'});

    const bubble = el('div', {className:`message-bubble ${role === 'user' ? 'user-bubble' : ''}`, textContent: text});

    wrap.appendChild(label);
    wrap.appendChild(bubble);
    byID<HTMLElement>('chat-area').appendChild(wrap);
    this.scrollChat();
    return bubble;
  }

  scrollChat(): void {
    const area = byID<HTMLElement>('chat-area');
    area.scrollTop = area.scrollHeight;
  }

  clearChat(): void {
    this.chatHistory = [];
    const area = byID<HTMLElement>('chat-area');
    const kids = Array.from(area.children) as HTMLElement[];
    kids.forEach((k) => {
      if (k.id !== 'empty-state') k.remove();
    });
    this.officeApp.show('empty-state');
  }


  updateWordCount(): void {
    const wc = byID<HTMLElement>('doc-word-count');
    if (!this.docText) { wc.textContent = ''; return; }
    const words = this.docText.trim().split(/\s+/).filter(Boolean).length;
    wc.textContent = `${words.toLocaleString()} words`;
  }



  async sendMessage(): Promise<void> {
    if (this.streaming || !this.config) return;

    const input = byID<HTMLTextAreaElement>('user-input');
    const prompt = input.value.trim();
    if (!prompt) return;

    const includeDoc = (byID<HTMLInputElement>('chk-include-doc')).checked;
    const contextData = includeDoc ? await this.officeApp.getActiveContext() : null;

    // Initialize history with a hidden system prompt for the first turn
    if (this.chatHistory.length === 0) {
      const systemMsg = `You are a specialized AI assistant for Microsoft ${this.officeApp.getHostLabel()}. `
        + `Help the user with their document by answering questions, suggesting edits, and explaining your reasoning clearly.`;
      this.chatHistory.push({ role: 'system', content: systemMsg });
    }

    // Build the user message with context headers
    let userContent = prompt;
    if (contextData) {
      const header = contextData.source === 'selection' ? 'SELECTED TEXT' : 'FULL DOCUMENT CONTENT';
      userContent = `--- ${header} ---\n${contextData.text}\n---\n\nUSER REQUEST: ${prompt}`;
    }

    this.chatHistory.push({ role: 'user', content: userContent });
    this.appendMessage('user', prompt);
    input.value = '';
    this.resizeTextarea(input);
    byID<HTMLButtonElement>('btn-send').disabled = true;

    const bubble = this.appendMessage('assistant', '');
    await this.handleStreamingResponse(bubble);
  }

  // FIX 2: Plain text streaming — structured output + stream:true is incompatible with
  // Claude and most other Foundry-hosted models. The system prompt now instructs the
  // model in plain language; no JSON schema enforcement on the wire.
  async handleStreamingResponse(bubble: HTMLElement): Promise<void> {
    this.streaming = true;
    this.stopFlag = false;
    byID<HTMLButtonElement>('btn-send').disabled = true;
    byID<HTMLButtonElement>('btn-stop').classList.remove('hidden');

    let fullText = '';
    try {
      const modelId = byID<HTMLSelectElement>('model-select').value;
      if (!modelId) throw new Error('No model selected. Please choose a model from the dropdown.');

      const gen = this.streamCompletion(this.config!, modelId, this.chatHistory);
      for await (const token of gen) {
        if (this.stopFlag) break;
        fullText += token;
        bubble.textContent = fullText;
        this.scrollChat();
      }

      if (fullText) {
        this.chatHistory.push({ role: 'assistant', content: fullText });
      } else if (!this.stopFlag) {
        bubble.textContent = '(No response received.)';
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      bubble.textContent = `Error: ${msg}`;
    } finally {
      this.streaming = false;
      this.stopFlag = false;
      byID<HTMLButtonElement>('btn-stop').classList.add('hidden');
      const userInput = byID<HTMLTextAreaElement>('user-input');
      byID<HTMLButtonElement>('btn-send').disabled = userInput.value.trim().length === 0;
    }
  }

  // ─────────────────────────────────────────────
  //  ALERTS
  // ─────────────────────────────────────────────

  showAlert(containerId: string, msg: string, type: 'error' | 'success' | 'info'): void {
    const container = byID<HTMLElement>(containerId);
    const div = document.createElement('div');
    div.className = `alert alert-${type}`;
    div.textContent = msg;
    container.appendChild(div);
    setTimeout(() => div.remove(), 6000);
  }

  setSetupError(msg: string): void {
    const el = byID<HTMLElement>('setup-error');
    el.textContent = msg;
    el.classList.remove('hidden');
  }

  clearSetupError(): void {
    byID<HTMLElement>('setup-error').classList.add('hidden');
  }

  // ─────────────────────────────────────────────
  //  CONNECT
  // ─────────────────────────────────────────────

  async connect(resource: string, apiKey: string): Promise<void> {
    this.clearSetupError();

    const btnLabel = byID<HTMLElement>('btn-connect-label');
    const btnSpinner = byID<HTMLElement>('btn-connect-spinner');
    const btn = byID<HTMLButtonElement>('btn-connect');

    btnLabel.textContent = 'Connecting…';
    btnSpinner.classList.remove('hidden');
    btn.disabled = true;

    try {
      const cfg: Config = { resource: resource.trim(), apiKey: apiKey.trim() };
      const found = await this.fetchModels(cfg);

      if (found.length === 0) {
        throw new Error('No deployed models found on this resource. Deploy at least one model in Azure AI Foundry first.');
      }

      this.config = cfg;
      this.models = found;
      this.saveConfig(cfg);
      this.enterMain();

    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      this.setSetupError(msg);
    } finally {
      btnLabel.textContent = 'Connect to Foundry';
      btnSpinner.classList.add('hidden');
      btn.disabled = false;
    }
  }

  // ─────────────────────────────────────────────
  //  ENTER MAIN
  // ─────────────────────────────────────────────

  async enterMain(): Promise<void> {
    if (!this.config) return;

    this.showScreen('main');
    this.populateModelSelect();

    byID<HTMLElement>('footer-conn-name').textContent = `${this.config.resource}.services.ai.azure.com`;

    // Load document text in background
    this.docText = await this.getDocumentText();
    this.updateWordCount();
  }

  // ─────────────────────────────────────────────
  //  DISCONNECT
  // ─────────────────────────────────────────────

  disconnect(): void {
    this.config = null;
    this.models = [];
    this.chatHistory = [];
    this.docText = '';
    this.clearConfig();
    this.clearChat();
    this.showScreen('setup');
  }

  // ─────────────────────────────────────────────
  //  TEXTAREA AUTO-RESIZE
  // ─────────────────────────────────────────────

  resizeTextarea(el: HTMLTextAreaElement): void {
    el.style.height = 'auto';
    el.style.height = `${Math.min(el.scrollHeight, 120)}px`;
  }

  // ─────────────────────────────────────────────
  //  EVENT WIRING
  // ─────────────────────────────────────────────

  wireSetupEvents(): void {
    const btnConnect = byID<HTMLButtonElement>('btn-connect');
    if (!btnConnect) return;

    btnConnect.addEventListener('click', () => {
      const resource = (byID<HTMLInputElement>('input-resource')).value.trim();
      const apiKey = (byID<HTMLInputElement>('input-apikey')).value.trim();
      if (!resource) { this.setSetupError('Please enter your Foundry resource name.'); return; }
      if (!apiKey) { this.setSetupError('Please enter your API key.'); return; }
      this.connect(resource, apiKey);
    });

    const inputs = [
      byID<HTMLInputElement>('input-resource'),
      byID<HTMLInputElement>('input-apikey')
    ].filter(Boolean);

    inputs.forEach((inp) => {
      inp.addEventListener('keydown', (e: KeyboardEvent) => {
        if (e.key === 'Enter') byID<HTMLButtonElement>('btn-connect').click();
      });
    });
  }

  wireEvents(): void {
    // Deprecated: event wiring now handled by wireSetupEvents and wireMainEvents.
  }

  wireMainEvents(): void {
    const modelSelect = byID<HTMLSelectElement>('model-select');
    if (!modelSelect) return;

    modelSelect.addEventListener('change', (e: Event) => {
      this.updateModelCard((e.target as HTMLSelectElement).value);
    });

    byID<HTMLButtonElement>('btn-refresh-models').addEventListener('click', async () => {
      if (!this.config) return;
      const btn = byID<HTMLButtonElement>('btn-refresh-models');
      btn.disabled = true;
      try {
        const found = await this.fetchModels(this.config);
        this.models = found;
        this.populateModelSelect();
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        this.showAlert('chat-area', `Refresh failed: ${msg}`, 'error');
      } finally {
        btn.disabled = false;
      }
    });

    byID<HTMLButtonElement>('btn-send').addEventListener('click', () => this.sendMessage());
    byID<HTMLButtonElement>('btn-stop').addEventListener('click', () => { this.stopFlag = true; });

    const userInput = byID<HTMLTextAreaElement>('user-input');
    userInput.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        this.sendMessage();
      }
    });
    userInput.addEventListener('input', () => {
      this.resizeTextarea(userInput);
      byID<HTMLButtonElement>('btn-send').disabled = userInput.value.trim().length === 0;
    });

    byID<HTMLButtonElement>('btn-clear-chat').addEventListener('click', () => this.clearChat());
    byID<HTMLButtonElement>('btn-disconnect').addEventListener('click', () => this.disconnect());

    document.querySelectorAll<HTMLButtonElement>('.quick-prompt-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const prompt = btn.getAttribute('data-prompt') ?? '';
        userInput.value = prompt;
        this.resizeTextarea(userInput);
        byID<HTMLButtonElement>('btn-send').disabled = false;
        userInput.focus();
      });
    });

    byID<HTMLInputElement>('chk-include-doc').addEventListener('change', async () => {
      if (this.docText === '') {
        this.docText = await this.getDocumentText();
        this.updateWordCount();
      }
    });
  }

  // ─────────────────────────────────────────────
  //  INIT
  // ─────────────────────────────────────────────

  async init(): Promise<void> {
    // Try to restore saved config
    const saved = this.loadConfig();
    if (saved) {
      try {
        const found = await this.fetchModels(saved);
        if (found.length > 0) {
          this.config = saved;
          this.models = found;
          // FIX 3: use enterMain() instead of showScreen('main') directly so that
          // footer label and docText are both populated on config restore
          await this.enterMain();
          return;
        }
      } catch {
        // saved config invalid — fall through to setup
      }
    }

    this.showScreen('setup');
  }
}


// Boot via Office.js when running inside Word,
// or directly in browser for dev testing
const ui = new UI();
if (typeof Office !== 'undefined') {
  Office.onReady(() => ui.init());
} else {
  document.addEventListener('DOMContentLoaded', () => ui.init());
}