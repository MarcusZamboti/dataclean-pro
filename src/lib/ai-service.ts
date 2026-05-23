/**
 * AI Excel Copilot Service
 * Connects directly to Gemini or OpenAI APIs via native fetch.
 * Generates highly targeted JavaScript transformation code based on a data sample.
 */

export interface AIServiceResponse {
  explanation: string;
  code: string;
  log: string;
}

export async function transformDataWithAI(
  columns: string[],
  sampleRows: Record<string, unknown>[],
  userPrompt: string,
  apiKey?: string,
  provider: 'gemini' | 'openai' = 'gemini'
): Promise<AIServiceResponse> {
  const systemPrompt = `Você é um Engenheiro de Dados especialista em manipulação de planilhas de Excel.
Você manipula dados contidos em uma estrutura de Array de Objetos JSON JavaScript.
Sua missão é gerar um código JavaScript puro para transformar a planilha conforme a instrução do usuário.

Você DEVE retornar exclusivamente um objeto JSON estruturado no formato abaixo, sem qualquer texto adicional fora dele, sem blocos de markdown e sem explicações externas:

{
  "explanation": "Uma breve explicação amigável em português explicando as alterações que o seu código fará na planilha.",
  "code": "O código JavaScript de uma função anônima que recebe o parâmetro 'data' (a array completa) e retorna a array modificada. Exemplo: (data) => { return data.map(row => { row.nova_coluna = 'exemplo'; return row; }); }",
  "log": "Uma mensagem curta de registro para o histórico de auditoria. Exemplo: 'Coluna nova_coluna adicionada.'"
}

Restrições cruciais do código gerado:
1. O código do campo 'code' DEVE ser JavaScript vanilla puro, compatível com Node.js (V8) e NÃO deve conter imports, chamadas de rede (fetch) ou leituras do sistema de arquivos (fs). Deve ser apenas lógica computacional de arrays/objetos.
2. Não use delimitadores markdown adicionais (\`\`\`) dentro do campo "code", apenas a string do código em uma única linha ou com quebras de linha normais \\n. A função no campo "code" deve ter a sintaxe: (data) => { ... } ou function(data) { ... } e sempre retornar a array 'data' modificada.
3. Se o usuário pedir para criar uma coluna, certifique-se de adicioná-la a cada objeto. Se for realizar cálculos numéricos, lembre-se de tratar valores vazios ou não-numéricos (usando parseFloat(x) || 0) para evitar erros de NaN.
4. Lembre-se de preservar as outras colunas existentes em cada linha ao retornar (se usar .map(), retorne o objeto completo).

Informações sobre a Planilha:
- Colunas disponíveis: ${JSON.stringify(columns)}
- Amostra das primeiras ${sampleRows.length} linhas de dados: ${JSON.stringify(sampleRows, null, 2)}

Instrução do Usuário para Executar nos Dados:
"${userPrompt}"`;

  const finalApiKey = apiKey || (provider === 'gemini' ? process.env.GEMINI_API_KEY : process.env.OPENAI_API_KEY);

  if (!finalApiKey) {
    throw new Error(`Chave de API para o provedor ${provider.toUpperCase()} não foi configurada. Por favor, adicione sua chave de API nas configurações ou no arquivo .env.`);
  }

  try {
    if (provider === 'gemini') {
      // Gemini API call (v1beta) - We use gemini-2.5-flash as default, compatible with standard JSON
      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${finalApiKey}`;
      
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          contents: [{ parts: [{ text: systemPrompt }] }],
          generationConfig: {
            responseMimeType: "application/json",
            temperature: 0.1,
          }
        }),
      });

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Erro na API do Gemini: ${response.status} - ${errText}`);
      }

      const resData = await response.json();
      const content = resData.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!content) {
        throw new Error('Resposta vazia da API do Gemini.');
      }

      const parsed: AIServiceResponse = JSON.parse(content.trim());
      return parsed;
    } else {
      // OpenAI API call
      const url = 'https://api.openai.com/v1/chat/completions';
      
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${finalApiKey}`,
        },
        body: JSON.stringify({
          model: 'gpt-4o-mini',
          response_format: { type: "json_object" },
          temperature: 0.1,
          messages: [
            { role: 'system', content: 'Você é um assistente estruturado que responde apenas em formato JSON válido.' },
            { role: 'user', content: systemPrompt }
          ],
        }),
      });

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Erro na API da OpenAI: ${response.status} - ${errText}`);
      }

      const resData = await response.json();
      const content = resData.choices?.[0]?.message?.content;
      if (!content) {
        throw new Error('Resposta vazia da API da OpenAI.');
      }

      const parsed: AIServiceResponse = JSON.parse(content.trim());
      return parsed;
    }
  } catch (error) {
    console.error('Erro de comunicação com a Inteligência Artificial:', error);
    throw error;
  }
}
