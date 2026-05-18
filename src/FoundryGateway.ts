/// <reference path="../types.d.ts" />
export class FoundryGateway {
  /** Construct base URL for a given resource */
  static baseUrl(resource: string): string {
    return `https://${resource}.services.ai.azure.com/`;
  }

  /** Fetch the list of deployed models for a config */
  static async fetchModels(cfg: Config): Promise<FoundryModel[]> {
    const url = `${FoundryGateway.baseUrl(cfg.resource)}models`;
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
    messages: ChatMessage[]
  ): AsyncGenerator<string> {
    const url = `${FoundryGateway.baseUrl(cfg.resource)}api/projects/${cfg.project}/models/${encodeURIComponent(
      modelId
    )}/chat/completions?api-version=2024-10-21`;

    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'api-key': cfg.apiKey,
      },
      body: JSON.stringify({
        messages,
        stream: true,
        max_completion_tokens: 4096,
        // response_format removed: json_schema + stream:true is incompatible with
        // Claude and most other Azure AI Foundry-hosted models, and causes HTTP 400s
        // or garbled output. Plain streaming text is handled by the caller.
      }),
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