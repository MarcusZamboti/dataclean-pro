import { readFileContent, processFileData, cleanOldProcessedFiles, CleaningConfig } from '@/lib/data-cleaner';
import { NextRequest, NextResponse } from 'next/server';
import { transformDataWithAI } from '@/lib/ai-service';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File | null;
    const configStr = formData.get('config') as string | null;
    const aiInstruction = formData.get('aiInstruction') as string | null;
    const aiProvider = formData.get('aiProvider') as string | null;
    const aiApiKey = formData.get('aiApiKey') as string | null;
    
    let config: CleaningConfig | undefined;
    if (configStr) {
      try {
        config = JSON.parse(configStr) as CleaningConfig;
      } catch (err) {
        console.error('Failed to parse cleaning config:', err);
      }
    }

    if (!file) {
      return NextResponse.json(
        { error: 'Nenhum arquivo enviado. Por favor, selecione um arquivo .csv ou .xlsx.' },
        { status: 400 }
      );
    }

    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.csv') && !fileName.endsWith('.xlsx')) {
      return NextResponse.json(
        { error: 'Formato de arquivo inválido. Apenas arquivos .csv e .xlsx são aceitos.' },
        { status: 400 }
      );
    }

    // Limit file size (50MB)
    if (file.size > 50 * 1024 * 1024) {
      return NextResponse.json(
        { error: 'Arquivo muito grande. O limite máximo é 50MB.' },
        { status: 400 }
      );
    }

    const buffer = Buffer.from(await file.arrayBuffer());

    // Read file content
    const { data, format } = readFileContent(buffer, file.name);

    if (data.length === 0) {
      return NextResponse.json(
        { error: 'O arquivo está vazio ou não contém dados válidos.' },
        { status: 400 }
      );
    }

    // Process the data (with optional AI pre-instruction)
    let processedDataset = data;
    let initialLogs: any[] = [];

    if (aiInstruction && aiInstruction.trim()) {
      try {
        const columns = Object.keys(data[0]);
        const sampleRows = data.slice(0, 3);
        
        const aiResult = await transformDataWithAI(
          columns,
          sampleRows,
          aiInstruction,
          aiApiKey || undefined,
          (aiProvider as 'gemini' | 'openai') || 'gemini'
        );

        // Execute AI-generated code securely
        const fn = new Function('data', `return (${aiResult.code})(data)`);
        const transformed = fn(JSON.parse(JSON.stringify(data)));
        
        if (Array.isArray(transformed) && transformed.length > 0) {
          processedDataset = transformed as Record<string, unknown>[];
          initialLogs.push({
            type: 'info',
            description: `Copilot IA (Pré-limpeza): ${aiResult.log}`,
          });
        }
      } catch (aiErr) {
        console.error('AI Pre-cleaning transformation failed:', aiErr);
        const errMsg = aiErr instanceof Error ? aiErr.message : aiErr;
        return NextResponse.json(
          { error: `A instrução de IA pré-limpeza falhou: ${errMsg}` },
          { status: 400 }
        );
      }
    }

    const result = processFileData(processedDataset, file.name, format, config);
    if (initialLogs.length > 0) {
      result.logs = [...initialLogs, ...result.logs];
    }

    // Clean old files in the background without blocking current request
    setTimeout(() => {
      cleanOldProcessedFiles();
    }, 100);

    return NextResponse.json(result);
  } catch (error) {
    console.error('Error processing file:', error);
    const message = error instanceof Error ? error.message : 'Erro interno ao processar o arquivo.';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
