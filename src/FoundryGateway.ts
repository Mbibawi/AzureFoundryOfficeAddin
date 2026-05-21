/// <reference path="../types.d.ts" />
export class FoundryGateway {
  /** Construct base URL for a given resource */
  static baseUrl(resource: string): string {
    return `https://${resource}.services.ai.azure.com/`;
  }

  /** Fetch the list of deployed models for a config */
  static async fetchModels(cfg: Config): Promise<FoundryModel[]> {

    return [
      {
        id: 'claude-sonnet-4-6',
        name: 'Claude Sonnet 4.6',
        object: 'model',
        created: 1751063000000,
        owned_by: 'Anthropic AI'
      },
      {
        id: 'claude-haiku-4-5',
        name: 'Claude Haiku 4.5',
        object: 'model',
        created: 1751063000000,
        owned_by: 'Anthropic AI'
      },
      {
        id: 'claude-opus-4-7',
        name: 'Claude Opus 4.7',
        object: 'model',
        created: 1751063000000,
        owned_by: 'Anthropic AI'
      }]

    //const url = `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${cfg.resourceGroup}/providers/Microsoft.CognitiveServices/accounts/${accountName}/deployments?api-version=2024-06-01-preview`
    const url = `https://${cfg.resource}.services.ai.azure.com/openai/v1/models`;
    const res = await fetch(url, {
      headers: { 'api-key': cfg.apiKey },
    });
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`HTTP ${res.status}: ${body}`);
    }
    const data = (await res.json()) as FoundryResponse;
    return data.data ?? data.value ?? [];
  }

  /** Stream chat completions from a model */
  static async *streamCompletion(
    cfg: Config,
    modelId: string,
    messages: ChatMessage[],
    max: number = 4096,
    claude: boolean = true): AsyncGenerator<string> {
    const baseURL = `https://${cfg.resource}.services.ai.azure.com`;

    const url = claude ? `${baseURL}/anthropic/v1/messages` : `${baseURL}/api/projects/${cfg.project}/openai/v1/chat/completions`

    const body: { messages: ChatMessage[], model: string, stream: boolean, max_completion_tokens?: number, max_tokens?: number, system?: string } = {
      messages,
      model: modelId,
      stream: true,
    };

    const headers: { [key: string]: string } = {
      'Content-Type': 'application/json',
    };

    if (claude) {
      headers['anthropic-version'] = '2023-06-01';
      headers['x-api-key'] = cfg.apiKey;
      //headers['x-ms-model-mesh-model-name'] = modelId;
      body.system = messages.find(m => m.role === 'system')?.content ?? '';
      body.messages = messages.filter(m => m.role !== 'system');
      body.max_tokens = max;
    } else {
      headers['api-key'] = cfg.apiKey;
      body.max_completion_tokens = max;
    }

    const res = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const body = await res.text();
      throw new Error(`HTTP ${res.status}: ${body}`);
    }

    const reader = res.body!.getReader();
    const dec = new TextDecoder();
    let buf = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buf += dec.decode(value, { stream: true });
      const lines = buf.split('\n');

      // Keep the last partial line in the buffer
      buf = lines.pop() ?? '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith('data:')) continue;

        const jsonStr = trimmed.slice(5).trim();
        if (jsonStr === '[DONE]') return;

        try {
          const chunk = JSON.parse(jsonStr);
          const delta = chunk?.choices?.[0]?.delta?.content;
          if (typeof delta === 'string' && delta.length > 0) {
            yield delta;
          }
        } catch {
          // Ignore malformed lines during streaming
        }
      }
    }
  }
}