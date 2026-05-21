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
                id: 'claude-opus-4-7',
                name: 'Claude Opus 4.7',
                object: 'model',
                created: 1751063000000,
                owned_by: 'Anthropic AI'
            }
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
    static async *streamCompletion(cfg, modelId, messages, max = 4096) {
        const url = `https://${cfg.resource}.services.ai.azure.com/api/projects/${cfg.project}/openai/v1/chat/completions`;
        const res = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'api-key': cfg.apiKey,
            },
            body: JSON.stringify({
                messages,
                model: modelId,
                stream: true,
                max_completion_tokens: max,
            }),
        });
        if (!res.ok) {
            const body = await res.text();
            throw new Error(`HTTP ${res.status}: ${body}`);
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
                const jsonStr = trimmed.slice(5).trim();
                if (jsonStr === '[DONE]')
                    return;
                try {
                    const chunk = JSON.parse(jsonStr);
                    const delta = chunk?.choices?.[0]?.delta?.content;
                    if (typeof delta === 'string' && delta.length > 0) {
                        yield delta;
                    }
                }
                catch {
                }
            }
        }
    }
}
//# sourceMappingURL=FoundryGateway.js.map