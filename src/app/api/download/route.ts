import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';

export async function GET(request: NextRequest) {
  try {
    const id = request.nextUrl.searchParams.get('id');

    if (!id) {
      return NextResponse.json(
        { error: 'ID do arquivo não especificado.' },
        { status: 400 }
      );
    }

    // Strict UUIDv4 validation to prevent Path Traversal and sanitize inputs
    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
    if (!uuidRegex.test(id)) {
      return NextResponse.json(
        { error: 'Identificador de arquivo inválido.' },
        { status: 400 }
      );
    }

    const processedDir = path.join(process.cwd(), 'processed');

    // Try to find the file (could be csv or xlsx)
    const csvPath = path.join(processedDir, `${id}.csv`);
    const xlsxPath = path.join(processedDir, `${id}.xlsx`);
    const jsonPath = path.join(processedDir, `${id}.json`);

    let filePath: string | null = null;
    let contentType = '';
    let downloadName = 'dados_limpos';

    // Check JSON metadata for original format and filename
    if (fs.existsSync(jsonPath)) {
      const meta = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
      downloadName = meta.fileName
        ? meta.fileName.replace(/\.[^.]+$/, '') + '_limpo.' + (meta.format === 'xlsx' ? 'xlsx' : 'csv')
        : 'dados_limpos';
    }

    if (fs.existsSync(xlsxPath)) {
      filePath = xlsxPath;
      contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    } else if (fs.existsSync(csvPath)) {
      filePath = csvPath;
      contentType = 'text/csv; charset=utf-8';
    }

    if (!filePath || !fs.existsSync(filePath)) {
      return NextResponse.json(
        { error: 'Arquivo processado não encontrado. Ele pode ter expirado.' },
        { status: 404 }
      );
    }

    const fileBuffer = fs.readFileSync(filePath);

    return new NextResponse(fileBuffer, {
      headers: {
        'Content-Type': contentType,
        'Content-Disposition': `attachment; filename="${encodeURIComponent(downloadName)}"`,
      },
    });
  } catch (error) {
    console.error('Error downloading file:', error);
    return NextResponse.json(
      { error: 'Erro ao baixar o arquivo.' },
      { status: 500 }
    );
  }
}
