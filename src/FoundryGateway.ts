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
        id: 'claude-opus-4-6',
        name: 'Claude Opus 4.6',
        object: 'model',
        created: 1751063000000,
        owned_by: 'Anthropic AI'
      },
      {
        id: 'Kimi-K2.6',
        name: 'Kimi K2.6',
        object: 'model',
        created: 1751063000000,
        owned_by: 'Anthropic AI'
      },

    ]

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
    claude: boolean = false
  ): AsyncGenerator<string> {

    const baseURL = `https://${cfg.resource}.services.ai.azure.com`;

    if (modelId.includes('claude')) claude = true;

    yield* claude ? streamClaude() : streamOpenAI();

    async function* streamClaude(): AsyncGenerator<string> {
      const res = await fetch(`${baseURL}/anthropic/v1/messages`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': cfg.apiKey,
          'anthropic-version': '2023-06-01',
        },
        body: JSON.stringify({
          model: modelId,
          max_tokens: max,
          system: messages.find(m => m.role === 'system')?.content ?? '',
          messages: messages.filter(m => m.role !== 'system'),
          stream: true,
        }),
      });

      if (!res.ok) {
        const errText = await res.text();
        throw new Error(`HTTP ${res.status}: ${errText}`);
      }

      const reader = res.body!.getReader();
      const dec = new TextDecoder();
      let buf = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buf += dec.decode(value, { stream: true });
        const lines = buf.split('\n');
        buf = lines.pop() ?? '';

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed.startsWith('data:')) continue;
          try {
            const chunk = JSON.parse(trimmed.slice(5).trim());
            if (chunk?.type === 'message_stop') return;
            if (chunk?.type === 'content_block_delta' &&
              chunk?.delta?.type === 'text_delta') {
              yield chunk.delta.text;
            }
          } catch {
            // Ignore malformed lines
          }
        }
      }
    }

    async function* streamOpenAI(): AsyncGenerator<string> {
      const client = new OpenAI({
        apiKey: cfg.apiKey,
        baseURL: `${baseURL}/openai/v1`,
      });

      const completion = await client.chat.completions.create({
        model: modelId,
        messages: messages,
        stream:true
      });

      for await (const chunk of completion) {
        const text = chunk.choices[0]?.delta?.content;
        if (typeof text === 'string' && text.length > 0) {
          yield text;
        }
      }
    }

  }
}