import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import os from 'os';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { readFileContent, CleaningLog } from '@/lib/data-cleaner';
import { transformDataWithAI } from '@/lib/ai-service';

export async function POST(request: NextRequest) {
  try {
    const { id, prompt, apiKey, provider } = await request.json();

    if (!id || !prompt) {
      return NextResponse.json(
        { error: 'ID do arquivo e prompt do usuário são obrigatórios.' },
        { status: 400 }
      );
    }

    // Strict UUIDv4 validation to prevent Path Traversal
    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
    if (!uuidRegex.test(id)) {
      return NextResponse.json(
        { error: 'Identificador de arquivo inválido.' },
        { status: 400 }
      );
    }

    const processedDir = path.join(os.tmpdir(), 'dataclean-pro-processed');
    const jsonPath = path.join(processedDir, `${id}.json`);

    if (!fs.existsSync(jsonPath)) {
      return NextResponse.json(
        { error: 'Arquivo temporário ou logs não encontrados. Ele pode ter expirado.' },
        { status: 404 }
      );
    }

    // Load file metadata
    const meta = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
    const fileName = meta.fileName;
    const format = meta.format; // 'csv' ou 'xlsx'
    const filePath = path.join(processedDir, `${id}.${format}`);

    if (!fs.existsSync(filePath)) {
      return NextResponse.json(
        { error: 'Arquivo de planilha física não encontrado.' },
        { status: 404 }
      );
    }

    // Read full dataset
    const fileBuffer = fs.readFileSync(filePath);
    const { data: allData } = readFileContent(fileBuffer, fileName);

    if (!allData || allData.length === 0) {
      return NextResponse.json(
        { error: 'A planilha de dados está vazia ou não pôde ser lida.' },
        { status: 400 }
      );
    }

    const columns = Object.keys(allData[0]);
    const sampleRows = allData.slice(0, 3); // Grab first 3 rows as sample

    // Call AI to generate JavaScript transformation code
    const aiResult = await transformDataWithAI(
      columns,
      sampleRows,
      prompt,
      apiKey,
      provider || 'gemini'
    );

    // Execute the transformation code securely
    let transformedData: Record<string, unknown>[] = [];
    try {
      // The AI code represents a function like (data) => { ... }
      const fn = new Function('data', `return (${aiResult.code})(data)`);
      const result = fn(JSON.parse(JSON.stringify(allData))); // deep copy to prevent mutating in-place during trial
      
      if (!Array.isArray(result)) {
        throw new Error('O código gerado pela IA não retornou uma array de dados válida.');
      }
      transformedData = result as Record<string, unknown>[];
    } catch (evalErr) {
      console.error('Failed to execute AI-generated code:', evalErr, '\nCode was:', aiResult.code);
      return NextResponse.json(
        { error: `A IA gerou um código com falha de execução: ${evalErr instanceof Error ? evalErr.message : evalErr}` },
        { status: 500 }
      );
    }

    if (transformedData.length === 0) {
      return NextResponse.json(
        { error: 'A transformação resultou em zero linhas de dados.' },
        { status: 400 }
      );
    }

    // Save transformed files back to disk
    if (format === 'xlsx') {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(transformedData);
      
      // Auto-fit column widths for maximum legibility and professional design
      if (transformedData.length > 0) {
        const cols = Object.keys(transformedData[0]);
        ws['!cols'] = cols.map(col => {
          const maxLength = Math.max(
            col.length,
            ...transformedData.map(row => String(row[col] ?? '').length)
          );
          // Minimum width of 10, maximum of 50 to keep layout organized
          return { wch: Math.max(10, Math.min(50, maxLength + 3)) };
        });
      }

      XLSX.utils.book_append_sheet(wb, ws, 'Dados Limpos');
      const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      fs.writeFileSync(filePath, buf);
    } else {
      const csv = Papa.unparse(transformedData);
      fs.writeFileSync(filePath, csv, 'utf-8');
    }

    // Update metadata json
    const newLog: CleaningLog = {
      type: 'info',
      description: `Copilot IA: ${aiResult.log}`,
    };
    
    const updatedLogs = [...(meta.logs || []), newLog];
    const newColumns = Object.keys(transformedData[0]);
    
    fs.writeFileSync(
      jsonPath,
      JSON.stringify({
        ...meta,
        logs: updatedLogs,
        columns: newColumns,
      }),
      'utf-8'
    );

    // Return the response, including preview (first 10 rows) and full cleaned data
    return NextResponse.json({
      success: true,
      explanation: aiResult.explanation,
      log: aiResult.log,
      columns: newColumns,
      cleanedRowCount: transformedData.length,
      cleanedPreview: transformedData.slice(0, 10),
      cleanedData: transformedData,
      logs: updatedLogs,
    });
  } catch (error) {
    console.error('Error in AI Transform route:', error);
    const message = error instanceof Error ? error.message : 'Erro interno ao processar transformação com IA.';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
