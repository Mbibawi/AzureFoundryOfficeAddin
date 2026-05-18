export class FoundryGateway {
    static baseUrl(resource) {
        return `https://${resource}.services.ai.azure.com/`;
    }
    static async fetchModels(cfg) {
        const url = `${FoundryGateway.baseUrl(cfg.resource)}models`;
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
    static async *streamCompletion(cfg, modelId, messages) {
        const url = `${FoundryGateway.baseUrl(cfg.resource)}api/projects/${cfg.project}/models/${encodeURIComponent(modelId)}/chat/completions?api-version=2024-10-21`;
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