import OpenAI from "https://esm.sh/openai@4";
export class FoundryGateway {
    static baseUrl(resource) {
        return `https://${resource}.services.ai.azure.com/`;
    }
    static async fetchModels(cfg) {
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
        ];
        const url = `https://${cfg.resource}.services.ai.azure.com/openai/v1/models`;
        const res = await fetch(url, {
            headers: { 'api-key': cfg.apiKey },
        });
        if (!res.ok) {
            const body = await res.text();
            throw new Error(`HTTP ${res.status}: ${body}`);
        }
        const data = (await res.json());
        return data.data ?? data.value ?? [];
    }
    static async *streamCompletion(cfg, modelId, messages, max = 4096, claude = false) {
        const baseURL = `https://${cfg.resource}.services.ai.azure.com`;
        if (modelId.includes('claude'))
            claude = true;
        yield* claude ? streamClaude() : streamOpenAI();
        async function* streamClaude() {
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
            const reader = res.body.getReader();
            const dec = new TextDecoder();
            let buf = '';
            while (true) {
                const { done, value } = await reader.read();
                if (done)
                    break;
                buf += dec.decode(value, { stream: true });
                const lines = buf.split('\n');
                buf = lines.pop() ?? '';
                for (const line of lines) {
                    const trimmed = line.trim();
                    if (!trimmed.startsWith('data:'))
                        continue;
                    try {
                        const chunk = JSON.parse(trimmed.slice(5).trim());
                        if (chunk?.type === 'message_stop')
                            return;
                        if (chunk?.type === 'content_block_delta' &&
                            chunk?.delta?.type === 'text_delta') {
                            yield chunk.delta.text;
                        }
                    }
                    catch {
                    }
                }
            }
        }
        async function* streamOpenAI() {
            const client = new OpenAI({
                apiKey: cfg.apiKey,
                baseURL: `${baseURL}/openai/v1`,
            });
            const completion = await client.chat.completions.create({
                model: modelId,
                messages: messages,
                stream: true
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
//# sourceMappingURL=FoundryGateway.js.map