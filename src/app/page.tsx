'use client';

import React, { useState, useCallback, useRef, useEffect } from 'react';
import {
  Upload,
  Download,
  FileSpreadsheet,
  AlertTriangle,
  CheckCircle2,
  Copy,
  Trash2,
  ArrowRight,
  Loader2,
  FileText,
  Eraser,
  BarChart3,
  ChevronDown,
  ChevronUp,
  Info,
  Columns2,
  Sparkles,
  Send,
  Key,
  Bot,
  Settings,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Separator } from '@/components/ui/separator';
import { Progress } from '@/components/ui/progress';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table';
import { useToast } from '@/hooks/use-toast';
import { Switch } from '@/components/ui/switch';
import { Label } from '@/components/ui/label';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import {
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip as RechartsTooltip,
} from 'recharts';

// ============ TYPES ============

interface CleaningLog {
  type: 'error_fix' | 'duplicate_removed' | 'missing_filled' | 'standardization' | 'column_split' | 'info';
  column?: string;
  row?: number;
  original?: string;
  corrected?: string;
  description: string;
}

interface CleaningResult {
  id: string;
  originalFileName: string;
  originalFormat: 'csv' | 'xlsx';
  rowCount: number;
  columnCount: number;
  cleanedRowCount: number;
  cleanedColumnCount: number;
  columns: string[];
  originalPreview: Record<string, unknown>[];
  cleanedPreview: Record<string, unknown>[];
  logs: CleaningLog[];
  stats: {
    errorsFixed: number;
    duplicatesRemoved: number;
    missingValuesFilled: number;
    columnsStandardized: number;
    columnsSplit: number;
  };
  downloadUrl: string;
  cleanedData: Record<string, unknown>[];
}

// ============ LOG TYPE CONFIG ============

const logTypeConfig: Record<string, { label: string; color: string; icon: React.ReactNode }> = {
  error_fix: {
    label: 'Correção',
    color: 'bg-amber-100 text-amber-800 border-amber-200',
    icon: <AlertTriangle className="h-3.5 w-3.5" />,
  },
  duplicate_removed: {
    label: 'Duplicata',
    color: 'bg-rose-100 text-rose-800 border-rose-200',
    icon: <Copy className="h-3.5 w-3.5" />,
  },
  missing_filled: {
    label: 'Preenchido',
    color: 'bg-sky-100 text-sky-800 border-sky-200',
    icon: <Eraser className="h-3.5 w-3.5" />,
  },
  standardization: {
    label: 'Padronização',
    color: 'bg-violet-100 text-violet-800 border-violet-200',
    icon: <FileText className="h-3.5 w-3.5" />,
  },
  column_split: {
    label: 'Coluna Dividida',
    color: 'bg-orange-100 text-orange-800 border-orange-200',
    icon: <Columns2 className="h-3.5 w-3.5" />,
  },
  info: {
    label: 'Info',
    color: 'bg-emerald-100 text-emerald-800 border-emerald-200',
    icon: <Info className="h-3.5 w-3.5" />,
  },
};

// ============ MAIN COMPONENT ============

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [result, setResult] = useState<CleaningResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [showAllLogs, setShowAllLogs] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  // Configurações de Limpeza Customizável
  const [config, setConfig] = useState({
    standardizeColumns: true,
    splitColumns: true,
    fixErrors: true,
    removeDuplicates: true,
    fillMissing: true,
    missingStrategy: 'smart' as 'smart' | 'none',
  });

  const [aiInstruction, setAiInstruction] = useState('');

  // Copilot de IA de Excel
  const [aiPrompt, setAiPrompt] = useState('');
  const [isAiProcessing, setIsAiProcessing] = useState(false);
  const [aiMessages, setAiMessages] = useState<{ role: 'user' | 'assistant'; content: string; explanation?: string }[]>([]);
  const [aiProvider, setAiProvider] = useState<'gemini' | 'openai'>('gemini');
  const [aiApiKey, setAiApiKey] = useState('');
  const [showSettings, setShowSettings] = useState(false);

  // Carregar configurações de IA salvas
  useEffect(() => {
    if (typeof window !== 'undefined') {
      const savedKey = localStorage.getItem('dataclean_ai_key');
      const savedProvider = localStorage.getItem('dataclean_ai_provider');
      if (savedKey) setAiApiKey(savedKey);
      if (savedProvider) setAiProvider(savedProvider as 'gemini' | 'openai');
    }
  }, []);

  const handleAiTransform = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!aiPrompt.trim() || !result) return;

    const userMsg = aiPrompt;
    setAiPrompt('');
    setAiMessages(prev => [...prev, { role: 'user', content: userMsg }]);
    setIsAiProcessing(true);

    try {
      const response = await fetch('/api/ai-transform', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          id: result.id,
          prompt: userMsg,
          apiKey: aiApiKey || undefined,
          provider: aiProvider,
        }),
      });

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error || 'Erro ao processar instrução de IA');
      }

      const data = await response.json();
      
      // Update result dynamically
      setResult(prev => {
        if (!prev) return null;
        return {
          ...prev,
          columns: data.columns,
          cleanedRowCount: data.cleanedRowCount,
          cleanedPreview: data.cleanedPreview,
          cleanedData: data.cleanedData,
          logs: data.logs,
        };
      });

      setAiMessages(prev => [
        ...prev,
        {
          role: 'assistant',
          content: `Planilha atualizada com sucesso! Alteração realizada: **${data.log}**`,
          explanation: data.explanation,
        },
      ]);

      toast({
        title: 'IA Copilot Executada!',
        description: data.log,
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Erro desconhecido';
      setAiMessages(prev => [
        ...prev,
        {
          role: 'assistant',
          content: `Erro ao processar solicitação: ${message}`,
        },
      ]);
      toast({
        title: 'Falha no Copilot',
        description: message,
        variant: 'destructive',
      });
    } finally {
      setIsAiProcessing(false);
    }
  };

  const handleFileSelect = useCallback((selectedFile: File) => {
    const name = selectedFile.name.toLowerCase();
    if (!name.endsWith('.csv') && !name.endsWith('.xlsx')) {
      toast({
        title: 'Formato inválido',
        description: 'Por favor, selecione um arquivo .csv ou .xlsx',
        variant: 'destructive',
      });
      return;
    }
    setFile(selectedFile);
    setResult(null);
    setError(null);
  }, [toast]);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile) handleFileSelect(droppedFile);
    },
    [handleFileSelect]
  );

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleProcess = async () => {
    if (!file) return;

    setIsProcessing(true);
    setProgress(0);
    setError(null);

    // Simulate progress
    const progressInterval = setInterval(() => {
      setProgress(prev => {
        if (prev >= 90) {
          clearInterval(progressInterval);
          return 90;
        }
        return prev + Math.random() * 15;
      });
    }, 300);

    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('config', JSON.stringify(config));
      
      if (aiInstruction.trim()) {
        formData.append('aiInstruction', aiInstruction);
        formData.append('aiProvider', aiProvider);
        formData.append('aiApiKey', aiApiKey || '');
      }

      const response = await fetch('/api/process', {
        method: 'POST',
        body: formData,
      });

      clearInterval(progressInterval);
      setProgress(100);

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error || 'Erro ao processar o arquivo');
      }

      const data: CleaningResult = await response.json();
      setResult(data);

      toast({
        title: 'Processamento concluído!',
        description: `${data.stats.errorsFixed + data.stats.duplicatesRemoved + data.stats.missingValuesFilled} correções aplicadas`,
      });
    } catch (err) {
      clearInterval(progressInterval);
      const message = err instanceof Error ? err.message : 'Erro desconhecido';
      setError(message);
      toast({
        title: 'Erro no processamento',
        description: message,
        variant: 'destructive',
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleReset = () => {
    setFile(null);
    setResult(null);
    setError(null);
    setProgress(0);
    setAiInstruction('');
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleDownload = () => {
    if (!result || !result.cleanedData || result.cleanedData.length === 0) {
      toast({
        title: 'Erro no download',
        description: 'Não há dados limpos disponíveis para baixar.',
        variant: 'destructive',
      });
      return;
    }

    try {
      const baseName = result.originalFileName.replace(/\.[^.]+$/, '');
      const cleanedName = `${baseName}_limpo`;

      if (result.originalFormat === 'xlsx') {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(result.cleanedData);

        // Auto-fit column widths for maximum legibility and professional design
        const cols = Object.keys(result.cleanedData[0]);
        ws['!cols'] = cols.map(col => {
          const maxLength = Math.max(
            col.length,
            ...result.cleanedData.map((row: any) => String(row[col] ?? '').length)
          );
          // Minimum width of 10, maximum of 50 to keep layout organized
          return { wch: Math.max(10, Math.min(50, maxLength + 3)) };
        });

        XLSX.utils.book_append_sheet(wb, ws, 'Dados Limpos');
        XLSX.writeFile(wb, `${cleanedName}.xlsx`);
      } else {
        const csv = Papa.unparse(result.cleanedData);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', `${cleanedName}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

      toast({
        title: 'Download iniciado!',
        description: 'Seu arquivo limpo foi gerado localmente com sucesso.',
      });
    } catch (err) {
      console.error('Erro ao gerar o arquivo de download local:', err);
      toast({
        title: 'Erro ao baixar',
        description: 'Ocorreu um erro ao estruturar o arquivo de download no seu navegador.',
        variant: 'destructive',
      });
    }
  };

  const truncateValue = (val: unknown): string => {
    if (val === null || val === undefined) return '';
    const str = String(val);
    return str.length > 30 ? str.substring(0, 30) + '...' : str;
  };

  const visibleLogs = showAllLogs ? result?.logs : result?.logs.slice(0, 15);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-slate-100">
      {/* Header */}
      <header className="border-b bg-white/80 backdrop-blur-sm sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <div className="flex items-center gap-3">
            <div className="h-10 w-10 rounded-xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-lg shadow-emerald-200">
              <FileSpreadsheet className="h-5 w-5 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-900">DataClean Pro</h1>
              <p className="text-xs text-slate-500">Sistema Automatizado de Limpeza de Dados</p>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Upload Section */}
        {!result && (
          <div className="space-y-6">
            {/* Upload Area */}
            <Card className="border-0 shadow-lg shadow-slate-200/50">
              <CardContent className="p-0">
                <div
                  onDrop={handleDrop}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  className={`
                    relative p-8 sm:p-12 rounded-xl transition-all duration-300 cursor-pointer
                    ${isDragging
                      ? 'bg-emerald-50 border-2 border-dashed border-emerald-400 scale-[1.01]'
                      : file
                        ? 'bg-slate-50 border-2 border-solid border-emerald-300'
                        : 'bg-gradient-to-br from-slate-50 to-white border-2 border-dashed border-slate-200 hover:border-emerald-300 hover:bg-emerald-50/30'
                    }
                  `}
                  onClick={() => fileInputRef.current?.click()}
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".csv,.xlsx"
                    className="hidden"
                    onChange={e => {
                      const f = e.target.files?.[0];
                      if (f) handleFileSelect(f);
                    }}
                  />

                  <div className="text-center space-y-4">
                    {file ? (
                      <>
                        <div className="mx-auto h-16 w-16 rounded-2xl bg-emerald-100 flex items-center justify-center">
                          <FileSpreadsheet className="h-8 w-8 text-emerald-600" />
                        </div>
                        <div>
                          <p className="text-lg font-semibold text-slate-900">{file.name}</p>
                          <p className="text-sm text-slate-500">
                            {(file.size / 1024).toFixed(1)} KB
                            {file.name.endsWith('.xlsx') ? ' • Excel' : ' • CSV'}
                          </p>
                        </div>
                        <p className="text-sm text-slate-400">Clique para trocar o arquivo</p>
                      </>
                    ) : (
                      <>
                        <div className="mx-auto h-16 w-16 rounded-2xl bg-slate-100 flex items-center justify-center">
                          <Upload className="h-8 w-8 text-slate-400" />
                        </div>
                        <div>
                          <p className="text-lg font-semibold text-slate-700">
                            Arraste e solte seu arquivo aqui
                          </p>
                          <p className="text-sm text-slate-400 mt-1">
                            ou clique para selecionar • Formatos aceitos: .csv e .xlsx
                          </p>
                        </div>
                      </>
                    )}
                  </div>
                </div>
              </CardContent>
            </Card>

            {/* Painel de Configurações de Limpeza */}
            {file && !isProcessing && (
              <Card className="border-0 shadow-lg shadow-slate-200/50 bg-white">
                <CardHeader className="pb-3 border-b">
                  <CardTitle className="text-sm font-bold text-slate-800 flex items-center gap-2">
                    <Settings className="h-4 w-4 text-emerald-600 animate-spin-slow" />
                    Opções Personalizadas de Limpeza
                  </CardTitle>
                  <CardDescription className="text-xs">
                    Escolha quais regras e otimizações deseja aplicar nesta planilha
                  </CardDescription>
                </CardHeader>
                <CardContent className="p-4 space-y-4">
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                    <div className="flex items-center justify-between p-3 rounded-lg border bg-slate-50/50 hover:bg-slate-50 transition-colors">
                      <div className="space-y-0.5 pr-2">
                        <Label htmlFor="std-col" className="text-xs font-semibold text-slate-700 cursor-pointer">Padronizar Colunas</Label>
                        <p className="text-[10px] text-slate-400">Normaliza nomes (remove acentos, espaços e minúsculas)</p>
                      </div>
                      <Switch
                        id="std-col"
                        checked={config.standardizeColumns}
                        onCheckedChange={(checked) => setConfig(prev => ({ ...prev, standardizeColumns: checked }))}
                      />
                    </div>

                    <div className="flex items-center justify-between p-3 rounded-lg border bg-slate-50/50 hover:bg-slate-50 transition-colors">
                      <div className="space-y-0.5 pr-2">
                        <Label htmlFor="split-col" className="text-xs font-semibold text-slate-700 cursor-pointer">Dividir Colunas</Label>
                        <p className="text-[10px] text-slate-400">Detecta e divide dados compostos (separadores /, | ou ;)</p>
                      </div>
                      <Switch
                        id="split-col"
                        checked={config.splitColumns}
                        onCheckedChange={(checked) => setConfig(prev => ({ ...prev, splitColumns: checked }))}
                      />
                    </div>

                    <div className="flex items-center justify-between p-3 rounded-lg border bg-slate-50/50 hover:bg-slate-50 transition-colors">
                      <div className="space-y-0.5 pr-2">
                        <Label htmlFor="fix-err" className="text-xs font-semibold text-slate-700 cursor-pointer">Corrigir Erros e Formatação</Label>
                        <p className="text-[10px] text-slate-400">Remove chars invisíveis, formata datas ISO e moedas BR/US</p>
                      </div>
                      <Switch
                        id="fix-err"
                        checked={config.fixErrors}
                        onCheckedChange={(checked) => setConfig(prev => ({ ...prev, fixErrors: checked }))}
                      />
                    </div>

                    <div className="flex items-center justify-between p-3 rounded-lg border bg-slate-50/50 hover:bg-slate-50 transition-colors">
                      <div className="space-y-0.5 pr-2">
                        <Label htmlFor="rem-dup" className="text-xs font-semibold text-slate-700 cursor-pointer">Remover Duplicatas</Label>
                        <p className="text-[10px] text-slate-400">Elimina linhas redundantes idênticas mantendo a primeira</p>
                      </div>
                      <Switch
                        id="rem-dup"
                        checked={config.removeDuplicates}
                        onCheckedChange={(checked) => setConfig(prev => ({ ...prev, removeDuplicates: checked }))}
                      />
                    </div>

                    <div className="flex items-center justify-between p-3 rounded-lg border bg-slate-50/50 hover:bg-slate-50 transition-colors">
                      <div className="space-y-0.5 pr-2">
                        <Label htmlFor="fill-mis" className="text-xs font-semibold text-slate-700 cursor-pointer">Imputação Inteligente</Label>
                        <p className="text-[10px] text-slate-400">Preenche lacunas (mediana para números e moda para textos)</p>
                      </div>
                      <Switch
                        id="fill-mis"
                        checked={config.fillMissing}
                        onCheckedChange={(checked) => setConfig(prev => ({ ...prev, fillMissing: checked, missingStrategy: checked ? 'smart' : 'none' }))}
                      />
                    </div>
                  </div>

                  {/* Instruções de IA pré-processamento */}
                  <div className="space-y-2 mt-4 pt-4 border-t border-slate-100">
                    <Label htmlFor="ai-instruction" className="text-xs font-bold text-emerald-700 flex items-center gap-1.5 cursor-pointer">
                      <Sparkles className="h-3.5 w-3.5 animate-pulse" />
                      Instruções Especiais de IA antes de processar (Opcional)
                    </Label>
                    <textarea
                      id="ai-instruction"
                      value={aiInstruction}
                      onChange={(e) => setAiInstruction(e.target.value)}
                      placeholder="Ex: 'Remova a coluna CPF', 'Crie uma coluna lucro = preco - custo', 'Mantenha apenas registros de SP'..."
                      className="w-full text-xs rounded-lg border border-slate-200 p-2.5 min-h-[65px] focus:outline-none focus:ring-1 focus:ring-emerald-500 bg-white"
                      disabled={isProcessing}
                    />
                    <p className="text-[10px] text-slate-400 leading-tight">
                      Se preenchido, a IA selecionada nas configurações executará esta instrução na planilha antes do motor de limpeza clássico tratar os dados. (Requer chave de IA salva).
                    </p>
                  </div>
                </CardContent>
              </Card>
            )}

            {/* Process Button */}
            {file && (
              <div className="flex justify-center gap-4">
                <Button
                  onClick={handleProcess}
                  disabled={isProcessing}
                  size="lg"
                  className="px-8 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 shadow-lg shadow-emerald-200 text-white font-semibold"
                >
                  {isProcessing ? (
                    <>
                      <Loader2 className="mr-2 h-5 w-5 animate-spin" />
                      Processando...
                    </>
                  ) : (
                    <>
                      <Eraser className="mr-2 h-5 w-5" />
                      Processar e Limpar Dados
                    </>
                  )}
                </Button>
                <Button
                  onClick={handleReset}
                  variant="outline"
                  size="lg"
                  disabled={isProcessing}
                  className="px-6"
                >
                  <Trash2 className="mr-2 h-4 w-4" />
                  Limpar
                </Button>
              </div>
            )}

            {/* Progress Bar */}
            {isProcessing && (
              <Card className="border-0 shadow-lg shadow-slate-200/50">
                <CardContent className="p-6">
                  <div className="space-y-3">
                    <div className="flex justify-between text-sm">
                      <span className="text-slate-600 font-medium">Processando dados...</span>
                      <span className="text-emerald-600 font-semibold">{Math.round(progress)}%</span>
                    </div>
                    <Progress value={progress} className="h-2" />
                    <div className="flex flex-wrap gap-4 text-xs text-slate-400">
                      <span className="flex items-center gap-1">
                        <div className="h-1.5 w-1.5 rounded-full bg-emerald-400 animate-pulse" />
                        Padronizando colunas
                      </span>
                      <span className="flex items-center gap-1">
                        <div className="h-1.5 w-1.5 rounded-full bg-orange-400 animate-pulse" />
                        Dividindo colunas compostas
                      </span>
                      <span className="flex items-center gap-1">
                        <div className="h-1.5 w-1.5 rounded-full bg-amber-400 animate-pulse" />
                        Corrigindo erros
                      </span>
                      <span className="flex items-center gap-1">
                        <div className="h-1.5 w-1.5 rounded-full bg-rose-400 animate-pulse" />
                        Removendo duplicatas
                      </span>
                      <span className="flex items-center gap-1">
                        <div className="h-1.5 w-1.5 rounded-full bg-sky-400 animate-pulse" />
                        Preenchendo vazios
                      </span>
                    </div>
                  </div>
                </CardContent>
              </Card>
            )}

            {/* Error Message */}
            {error && (
              <Card className="border-rose-200 bg-rose-50">
                <CardContent className="p-6">
                  <div className="flex items-start gap-3">
                    <AlertTriangle className="h-5 w-5 text-rose-500 mt-0.5 shrink-0" />
                    <div>
                      <p className="font-semibold text-rose-800">Erro no processamento</p>
                      <p className="text-sm text-rose-600 mt-1">{error}</p>
                    </div>
                  </div>
                </CardContent>
              </Card>
            )}

            {/* Features Info */}
            {!file && !isProcessing && (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 mt-8">
                {[
                  {
                    icon: <Columns2 className="h-5 w-5" />,
                    title: 'Separação de Colunas',
                    desc: 'Detecta dados compostos separados por /, | ou ; e divide em duas colunas',
                    color: 'text-orange-600 bg-orange-50',
                  },
                  {
                    icon: <AlertTriangle className="h-5 w-5" />,
                    title: 'Detecção de Erros',
                    desc: 'Identifica e corrige formatação incorreta, valores inválidos e caracteres especiais',
                    color: 'text-amber-600 bg-amber-50',
                  },
                  {
                    icon: <Copy className="h-5 w-5" />,
                    title: 'Remoção de Duplicatas',
                    desc: 'Elimina linhas duplicadas mantendo a primeira ocorrência',
                    color: 'text-rose-600 bg-rose-50',
                  },
                  {
                    icon: <Eraser className="h-5 w-5" />,
                    title: 'Preenchimento Inteligente',
                    desc: 'Preenche campos vazios com mediana, valor mais frequente ou interpolação',
                    color: 'text-sky-600 bg-sky-50',
                  },
                  {
                    icon: <FileText className="h-5 w-5" />,
                    title: 'Padronização',
                    desc: 'Normaliza nomes de colunas, tipos de dados e formatos',
                    color: 'text-violet-600 bg-violet-50',
                  },
                ].map((feature, i) => (
                  <Card key={i} className="border-0 shadow-sm hover:shadow-md transition-shadow">
                    <CardContent className="p-5">
                      <div className={`h-10 w-10 rounded-lg ${feature.color} flex items-center justify-center mb-3`}>
                        {feature.icon}
                      </div>
                      <h3 className="font-semibold text-slate-900 text-sm">{feature.title}</h3>
                      <p className="text-xs text-slate-500 mt-1 leading-relaxed">{feature.desc}</p>
                    </CardContent>
                  </Card>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Results Section */}
        {result && (
          <div className="space-y-6">
            {/* Stats Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-4">
              <Card className="border-0 shadow-sm bg-gradient-to-br from-amber-50 to-white">
                <CardContent className="p-4">
                  <div className="flex items-center gap-3">
                    <div className="h-10 w-10 rounded-lg bg-amber-100 flex items-center justify-center">
                      <AlertTriangle className="h-5 w-5 text-amber-600" />
                    </div>
                    <div>
                      <p className="text-2xl font-bold text-amber-700">{result.stats.errorsFixed}</p>
                      <p className="text-xs text-amber-600">Erros Corrigidos</p>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card className="border-0 shadow-sm bg-gradient-to-br from-rose-50 to-white">
                <CardContent className="p-4">
                  <div className="flex items-center gap-3">
                    <div className="h-10 w-10 rounded-lg bg-rose-100 flex items-center justify-center">
                      <Copy className="h-5 w-5 text-rose-600" />
                    </div>
                    <div>
                      <p className="text-2xl font-bold text-rose-700">{result.stats.duplicatesRemoved}</p>
                      <p className="text-xs text-rose-600">Duplicatas Removidas</p>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card className="border-0 shadow-sm bg-gradient-to-br from-sky-50 to-white">
                <CardContent className="p-4">
                  <div className="flex items-center gap-3">
                    <div className="h-10 w-10 rounded-lg bg-sky-100 flex items-center justify-center">
                      <Eraser className="h-5 w-5 text-sky-600" />
                    </div>
                    <div>
                      <p className="text-2xl font-bold text-sky-700">{result.stats.missingValuesFilled}</p>
                      <p className="text-xs text-sky-600">Campos Preenchidos</p>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card className="border-0 shadow-sm bg-gradient-to-br from-violet-50 to-white">
                <CardContent className="p-4">
                  <div className="flex items-center gap-3">
                    <div className="h-10 w-10 rounded-lg bg-violet-100 flex items-center justify-center">
                      <FileText className="h-5 w-5 text-violet-600" />
                    </div>
                    <div>
                      <p className="text-2xl font-bold text-violet-700">{result.stats.columnsStandardized}</p>
                      <p className="text-xs text-violet-600">Colunas Padronizadas</p>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card className="border-0 shadow-sm bg-gradient-to-br from-orange-50 to-white">
                <CardContent className="p-4">
                  <div className="flex items-center gap-3">
                    <div className="h-10 w-10 rounded-lg bg-orange-100 flex items-center justify-center">
                      <Columns2 className="h-5 w-5 text-orange-600" />
                    </div>
                    <div>
                      <p className="text-2xl font-bold text-orange-700">{result.stats.columnsSplit}</p>
                      <p className="text-xs text-orange-600">Colunas Divididas</p>
                    </div>
                  </div>
                </CardContent>
              </Card>
            </div>

            {/* Summary Row */}
            <Card className="border-0 shadow-sm">
              <CardContent className="p-5">
                <div className="flex flex-wrap items-center justify-between gap-4">
                  <div className="flex items-center gap-4">
                    <div className="h-12 w-12 rounded-xl bg-emerald-100 flex items-center justify-center">
                      <CheckCircle2 className="h-6 w-6 text-emerald-600" />
                    </div>
                    <div>
                      <p className="font-bold text-slate-900">
                        {result.originalFileName}
                        <ArrowRight className="inline h-4 w-4 mx-2 text-emerald-500" />
                        Dados limpos
                      </p>
                      <p className="text-sm text-slate-500">
                        {result.rowCount} linhas → {result.cleanedRowCount} linhas •{' '}
                        {result.columnCount} colunas → {result.cleanedColumnCount} colunas •{' '}
                        Formato: {result.originalFormat.toUpperCase()}
                      </p>
                    </div>
                  </div>
                  <div className="flex gap-3">
                    <Button
                      onClick={handleDownload}
                      className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 shadow-md shadow-emerald-200 text-white"
                    >
                      <Download className="mr-2 h-4 w-4" />
                      Baixar Arquivo Limpo
                    </Button>
                    <Button onClick={handleReset} variant="outline">
                      <Upload className="mr-2 h-4 w-4" />
                      Novo Arquivo
                    </Button>
                  </div>
                </div>
              </CardContent>
            </Card>

            {/* Data Preview & Logs */}
            <Tabs defaultValue="preview" className="space-y-4">
              <TabsList className="bg-white shadow-sm border">
                <TabsTrigger value="preview" className="gap-2">
                  <FileSpreadsheet className="h-4 w-4" />
                  Visualização dos Dados
                </TabsTrigger>
                <TabsTrigger value="insights" className="gap-2">
                  <BarChart3 className="h-4 w-4" />
                  Saúde e Gráficos
                </TabsTrigger>
                <TabsTrigger value="copilot" className="gap-2 text-emerald-700 font-semibold">
                  <Sparkles className="h-4 w-4 text-emerald-600 animate-pulse" />
                  Copilot IA Excel
                </TabsTrigger>
                <TabsTrigger value="logs" className="gap-2">
                  <FileText className="h-4 w-4" />
                  Log de Correções
                  <Badge variant="secondary" className="ml-1 text-xs">
                    {result.logs.length}
                  </Badge>
                </TabsTrigger>
              </TabsList>

              {/* Data Preview Tab */}
              <TabsContent value="preview">
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                  {/* Original Data */}
                  <Card className="border-0 shadow-sm">
                    <CardHeader className="pb-3">
                      <CardTitle className="text-base flex items-center gap-2">
                        <div className="h-3 w-3 rounded-full bg-slate-400" />
                        Dados Originais
                        <Badge variant="outline" className="text-xs">
                          {result.rowCount} linhas
                        </Badge>
                      </CardTitle>
                      <CardDescription className="text-xs">
                        Primeiras 10 linhas do arquivo original
                      </CardDescription>
                    </CardHeader>
                    <CardContent className="p-0">
                      <ScrollArea className="h-[400px]">
                        <div className="overflow-x-auto">
                          <Table>
                            <TableHeader>
                              <TableRow className="bg-slate-50">
                                <TableHead className="w-12 text-xs">#</TableHead>
                                {result.columns.map(col => (
                                  <TableHead key={col} className="text-xs whitespace-nowrap min-w-[100px] max-w-[200px] truncate text-ellipsis">
                                    {col}
                                  </TableHead>
                                ))}
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {result.originalPreview.map((row, i) => (
                                <TableRow key={i}>
                                  <TableCell className="text-xs text-slate-400 font-mono">
                                    {i + 1}
                                  </TableCell>
                                  {result.columns.map(col => (
                                    <TableCell key={col} className="text-xs whitespace-nowrap max-w-[200px] truncate text-ellipsis">
                                      {truncateValue(row[col])}
                                    </TableCell>
                                  ))}
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </div>
                      </ScrollArea>
                    </CardContent>
                  </Card>

                  {/* Cleaned Data */}
                  <Card className="border-0 shadow-sm">
                    <CardHeader className="pb-3">
                      <CardTitle className="text-base flex items-center gap-2">
                        <div className="h-3 w-3 rounded-full bg-emerald-500" />
                        Dados Limpos
                        <Badge variant="outline" className="text-xs text-emerald-600 border-emerald-200 bg-emerald-50">
                          {result.cleanedRowCount} linhas
                        </Badge>
                      </CardTitle>
                      <CardDescription className="text-xs">
                        Primeiras 10 linhas após processamento
                      </CardDescription>
                    </CardHeader>
                    <CardContent className="p-0">
                      <ScrollArea className="h-[400px]">
                        <div className="overflow-x-auto">
                          <Table>
                            <TableHeader>
                              <TableRow className="bg-emerald-50">
                                <TableHead className="w-12 text-xs">#</TableHead>
                                {result.columns.map(col => (
                                  <TableHead key={col} className="text-xs whitespace-nowrap min-w-[100px] max-w-[200px] truncate text-ellipsis">
                                    {col}
                                  </TableHead>
                                ))}
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {result.cleanedPreview.map((row, i) => (
                                <TableRow key={i}>
                                  <TableCell className="text-xs text-slate-400 font-mono">
                                    {i + 1}
                                  </TableCell>
                                  {result.columns.map(col => (
                                    <TableCell key={col} className="text-xs whitespace-nowrap max-w-[200px] truncate text-ellipsis">
                                      {truncateValue(row[col])}
                                    </TableCell>
                                  ))}
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </div>
                      </ScrollArea>
                    </CardContent>
                  </Card>
                </div>
              </TabsContent>

              {/* Logs Tab */}
              <TabsContent value="logs">
                <Card className="border-0 shadow-sm">
                  <CardHeader className="pb-3">
                    <div className="flex items-center justify-between">
                      <div>
                        <CardTitle className="text-base">Log de Correções Aplicadas</CardTitle>
                        <CardDescription className="text-xs mt-1">
                          Registro detalhado de todas as modificações realizadas nos dados
                        </CardDescription>
                      </div>
                      <div className="flex gap-2">
                        {Object.entries(logTypeConfig).filter(([key]) => key !== 'info').map(([key, config]) => {
                          const count = result.logs.filter(l => l.type === key).length;
                          if (count === 0) return null;
                          return (
                            <Badge key={key} variant="outline" className="text-xs gap-1">
                              {config.icon}
                              {config.label}: {count}
                            </Badge>
                          );
                        })}
                      </div>
                    </div>
                  </CardHeader>
                  <CardContent className="p-0">
                    <ScrollArea className="h-[500px]">
                      <div className="divide-y">
                        {visibleLogs?.map((log, i) => {
                          const config = logTypeConfig[log.type] || logTypeConfig.info;
                          return (
                            <div key={i} className="px-6 py-3 hover:bg-slate-50 transition-colors">
                              <div className="flex items-start gap-3">
                                <Badge
                                  variant="outline"
                                  className={`text-[10px] shrink-0 gap-1 ${config.color}`}
                                >
                                  {config.icon}
                                  {config.label}
                                </Badge>
                                <div className="flex-1 min-w-0">
                                  <p className="text-sm text-slate-700">{log.description}</p>
                                  {log.original && log.corrected && (
                                    <div className="flex items-center gap-2 mt-1">
                                      <code className="text-xs bg-slate-100 px-1.5 py-0.5 rounded text-slate-500 max-w-[200px] truncate">
                                        {log.original}
                                      </code>
                                      <ArrowRight className="h-3 w-3 text-emerald-500 shrink-0" />
                                      <code className="text-xs bg-emerald-50 px-1.5 py-0.5 rounded text-emerald-700 max-w-[200px] truncate">
                                        {log.corrected}
                                      </code>
                                    </div>
                                  )}
                                </div>
                                {log.row && (
                                  <span className="text-xs text-slate-400 shrink-0">
                                    Linha {log.row}
                                  </span>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                      {result.logs.length > 15 && (
                        <div className="p-4 text-center border-t">
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={() => setShowAllLogs(!showAllLogs)}
                            className="text-xs"
                          >
                            {showAllLogs ? (
                              <>
                                <ChevronUp className="h-3 w-3 mr-1" />
                                Mostrar menos
                              </>
                            ) : (
                              <>
                                <ChevronDown className="h-3 w-3 mr-1" />
                                Mostrar todas as {result.logs.length} correções
                              </>
                            )}
                          </Button>
                        </div>
                      )}
                    </ScrollArea>
                  </CardContent>
                </Card>
              </TabsContent>

              {/* Insights Gráficos Tab */}
              <TabsContent value="insights">
                {(() => {
                  const barData = result ? [
                    { name: 'Erros Fixados', quantidade: result.stats.errorsFixed, fill: '#f59e0b' },
                    { name: 'Duplicatas', quantidade: result.stats.duplicatesRemoved, fill: '#f43f5e' },
                    { name: 'Preenchidos', quantidade: result.stats.missingValuesFilled, fill: '#0ea5e9' },
                    { name: 'Padronizados', quantidade: result.stats.columnsStandardized, fill: '#8b5cf6' },
                    { name: 'Divididos', quantidade: result.stats.columnsSplit, fill: '#f97316' },
                  ].filter(d => d.quantidade > 0) : [];

                  const totalCells = result ? result.rowCount * result.columnCount : 0;
                  const totalAlterations = result ? (
                    result.stats.errorsFixed +
                    result.stats.missingValuesFilled +
                    result.stats.columnsStandardized +
                    result.stats.columnsSplit
                  ) : 0;
                  const healthyCells = Math.max(0, totalCells - totalAlterations);

                  const pieData = [
                    { name: 'Originais Saudáveis', valor: healthyCells, color: '#10b981' },
                    { name: 'Corrigidos/Limpos', valor: totalAlterations, color: '#f59e0b' },
                  ];

                  return (
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                      {/* Card Rosca - Saúde dos Dados */}
                      <Card className="border-0 shadow-md bg-white">
                        <CardHeader className="pb-3">
                          <CardTitle className="text-base flex items-center gap-2">
                            <BarChart3 className="h-4 w-4 text-emerald-600" />
                            Saúde Geral da Planilha
                          </CardTitle>
                          <CardDescription className="text-xs">
                            Proporção de células íntegras originais vs corrigidas
                          </CardDescription>
                        </CardHeader>
                        <CardContent className="h-[280px] flex items-center justify-center">
                          {totalCells === 0 ? (
                            <div className="text-center text-slate-400 text-sm">
                              Sem dados suficientes para gerar análise.
                            </div>
                          ) : (
                            <div className="w-full h-full flex flex-col sm:flex-row items-center justify-center gap-6">
                              <div className="w-[180px] h-[180px] shrink-0">
                                <ResponsiveContainer width="100%" height="100%">
                                  <PieChart>
                                    <Pie
                                      data={pieData}
                                      cx="50%"
                                      cy="50%"
                                      innerRadius={50}
                                      outerRadius={70}
                                      paddingAngle={4}
                                      dataKey="valor"
                                    >
                                      {pieData.map((entry, index) => (
                                        <Cell key={`cell-${index}`} fill={entry.color} />
                                      ))}
                                    </Pie>
                                    <RechartsTooltip formatter={(value) => `${value} células`} />
                                  </PieChart>
                                </ResponsiveContainer>
                              </div>
                              <div className="space-y-2">
                                <div className="flex items-center gap-2">
                                  <div className="h-3 w-3 rounded-full bg-emerald-500" />
                                  <span className="text-xs text-slate-600 font-semibold">Íntegros: {healthyCells} ({((healthyCells / Math.max(1, totalCells)) * 100).toFixed(1)}%)</span>
                                </div>
                                <div className="flex items-center gap-2">
                                  <div className="h-3 w-3 rounded-full bg-amber-500" />
                                  <span className="text-xs text-slate-600 font-semibold">Corrigidos: {totalAlterations} ({((totalAlterations / Math.max(1, totalCells)) * 100).toFixed(1)}%)</span>
                                </div>
                                <div className="text-[10px] text-slate-400 pt-2 border-t mt-1">
                                  Total: {totalCells} células ({result.rowCount} linhas x {result.columnCount} colunas).
                                </div>
                              </div>
                            </div>
                          )}
                        </CardContent>
                      </Card>

                      {/* Card Barras - Distribuição de Correções */}
                      <Card className="border-0 shadow-md bg-white">
                        <CardHeader className="pb-3">
                          <CardTitle className="text-base flex items-center gap-2">
                            <BarChart3 className="h-4 w-4 text-emerald-600" />
                            Correções por Categoria
                          </CardTitle>
                          <CardDescription className="text-xs">
                            Total de anomalias limpas e padronizações efetuadas
                          </CardDescription>
                        </CardHeader>
                        <CardContent className="h-[280px] pb-4 flex items-center justify-center">
                          {barData.length === 0 ? (
                            <div className="text-center text-slate-400 text-sm">
                              Nenhuma correção foi necessária nos dados analisados! 🌟
                            </div>
                          ) : (
                            <ResponsiveContainer width="100%" height="100%">
                              <BarChart data={barData} layout="vertical" margin={{ top: 10, right: 30, left: 10, bottom: 5 }}>
                                <XAxis type="number" stroke="#94a3b8" fontSize={10} />
                                <YAxis dataKey="name" type="category" stroke="#94a3b8" fontSize={10} width={90} />
                                <RechartsTooltip formatter={(value) => [`${value} alterações`, 'Quantidade']} />
                                <Bar dataKey="quantidade" radius={[0, 4, 4, 0]}>
                                  {barData.map((entry, index) => (
                                    <Cell key={`cell-${index}`} fill={entry.fill} />
                                  ))}
                                </Bar>
                              </BarChart>
                            </ResponsiveContainer>
                          )}
                        </CardContent>
                      </Card>
                    </div>
                  );
                })()}
              </TabsContent>

              {/* Copilot IA Tab */}
              <TabsContent value="copilot">
                <div className="flex flex-col md:flex-row gap-6">
                  {/* Chat Area */}
                  <Card className="flex-1 border-0 shadow-md flex flex-col h-[520px] bg-white">
                    <CardHeader className="pb-3 border-b flex flex-row items-center justify-between">
                      <div className="flex items-center gap-2">
                        <div className="h-2.5 w-2.5 rounded-full bg-emerald-500 animate-pulse" />
                        <div>
                          <CardTitle className="text-base flex items-center gap-1.5">
                            <Sparkles className="h-4 w-4 text-emerald-600" />
                            Copilot IA Excel
                          </CardTitle>
                          <CardDescription className="text-xs">
                            Crie colunas, faça cálculos matemáticos ou filtre registros usando português
                          </CardDescription>
                        </div>
                      </div>
                      <Button
                        variant="ghost"
                        size="sm"
                        onClick={() => setShowSettings(!showSettings)}
                        className="text-xs md:hidden"
                      >
                        <Settings className="h-4 w-4 text-slate-500" />
                      </Button>
                    </CardHeader>
                    
                    <CardContent className="flex-1 p-0 flex flex-col justify-between overflow-hidden">
                      {/* Messages Area */}
                      <ScrollArea className="flex-1 p-4">
                        {aiMessages.length === 0 ? (
                          <div className="h-full flex flex-col items-center justify-center text-center p-6 space-y-5 mt-6">
                            <div className="h-12 w-12 rounded-2xl bg-emerald-50 flex items-center justify-center text-emerald-600">
                              <Sparkles className="h-6 w-6" />
                            </div>
                            <div className="max-w-[420px] space-y-2">
                              <h3 className="font-semibold text-slate-800 text-sm">O que deseja fazer na planilha hoje?</h3>
                              <p className="text-xs text-slate-400">
                                Como seu copiloto inteligente, posso manipular a sua planilha em tempo real. Veja algumas ideias de comandos rápidos:
                              </p>
                            </div>
                            {/* Quick Prompts */}
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 max-w-[480px] w-full pt-1">
                              {[
                                'Crie uma coluna "imposto" calculando 12% da coluna preço',
                                'Filtre as linhas onde o status seja "pendente"',
                                'Junte as colunas "nome" e "sobrenome" na coluna "nome_completo"',
                                'Ordene a tabela inteira por data de criação',
                              ].map((pText, i) => (
                                <button
                                  key={i}
                                  type="button"
                                  onClick={() => setAiPrompt(pText)}
                                  className="text-left text-xs p-2.5 rounded-lg border bg-slate-50 hover:bg-emerald-50/50 hover:border-emerald-200 transition-colors text-slate-600 hover:text-emerald-700 font-medium"
                                >
                                  {pText}
                                </button>
                              ))}
                            </div>
                          </div>
                        ) : (
                          <div className="space-y-4">
                            {aiMessages.map((msg, i) => (
                              <div
                                key={i}
                                className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'}`}
                              >
                                <div
                                  className={`max-w-[85%] rounded-2xl p-3.5 shadow-sm text-sm ${
                                    msg.role === 'user'
                                      ? 'bg-slate-100 text-slate-800 rounded-br-none'
                                      : 'bg-emerald-50 border border-emerald-100 text-slate-700 rounded-bl-none'
                                  }`}
                                >
                                  {msg.role === 'assistant' && (
                                    <div className="flex items-center gap-1.5 mb-1.5 pb-1 border-b border-emerald-100 text-emerald-800 font-bold text-xs">
                                      <Bot className="h-3.5 w-3.5 text-emerald-600 animate-bounce" />
                                      Copilot IA
                                    </div>
                                  )}
                                  <p className="leading-relaxed whitespace-pre-wrap">{msg.content}</p>
                                  {msg.explanation && (
                                    <div className="mt-2.5 pt-2.5 border-t border-emerald-200/50 text-[11px] text-slate-500 italic bg-emerald-100/30 p-2 rounded">
                                      <strong>Como foi feito:</strong> {msg.explanation}
                                    </div>
                                  )}
                                </div>
                              </div>
                            ))}
                            {isAiProcessing && (
                              <div className="flex justify-start">
                                <div className="max-w-[85%] rounded-2xl p-3.5 bg-slate-50 border border-slate-100 text-slate-500 text-xs flex items-center gap-2 rounded-bl-none">
                                  <Loader2 className="h-4 w-4 animate-spin text-emerald-600" />
                                  Processando planilha e aplicando transformações na IA...
                                </div>
                              </div>
                            )}
                          </div>
                        )}
                      </ScrollArea>

                      {/* Input Form */}
                      <form onSubmit={handleAiTransform} className="p-3 border-t bg-slate-50 flex gap-2">
                        <input
                          type="text"
                          value={aiPrompt}
                          onChange={(e) => setAiPrompt(e.target.value)}
                          placeholder="Digite aqui o que fazer na planilha (ex: 'calcule preco * 1.1'...)"
                          className="flex-1 text-sm bg-white rounded-lg border border-slate-200 px-3.5 py-2.5 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-shadow"
                          disabled={isAiProcessing}
                        />
                        <Button
                          type="submit"
                          disabled={isAiProcessing || !aiPrompt.trim()}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white shrink-0 shadow-md shadow-emerald-100"
                        >
                          <Send className="h-4 w-4" />
                        </Button>
                      </form>
                    </CardContent>
                  </Card>

                  {/* Settings Panel */}
                  <Card className={`w-full md:w-[280px] shrink-0 border-0 shadow-md bg-white ${showSettings ? 'block' : 'hidden md:block'}`}>
                    <CardHeader className="pb-3 border-b">
                      <CardTitle className="text-sm flex items-center gap-2">
                        <Key className="h-4 w-4 text-emerald-600" />
                        Configurações da IA
                      </CardTitle>
                      <CardDescription className="text-[11px]">
                        Conecte sua API Key para habilitar o copiloto de IA
                      </CardDescription>
                    </CardHeader>
                    <CardContent className="p-4 space-y-4">
                      {/* Provider Select */}
                      <div className="space-y-1.5">
                        <Label htmlFor="ai-provider" className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Provedor de Modelos</Label>
                        <select
                          id="ai-provider"
                          value={aiProvider}
                          onChange={(e) => {
                            const p = e.target.value as 'gemini' | 'openai';
                            setAiProvider(p);
                            if (typeof window !== 'undefined') localStorage.setItem('dataclean_ai_provider', p);
                          }}
                          className="w-full text-xs rounded border border-slate-200 p-2 bg-white focus:outline-none focus:ring-1 focus:ring-emerald-500"
                        >
                          <option value="gemini">Google Gemini (Recomendado)</option>
                          <option value="openai">OpenAI (GPT-4o-mini)</option>
                        </select>
                      </div>

                      {/* API Key Input */}
                      <div className="space-y-1.5">
                        <Label htmlFor="ai-key" className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Chave de API (Key)</Label>
                        <input
                          id="ai-key"
                          type="password"
                          value={aiApiKey}
                          onChange={(e) => setAiApiKey(e.target.value)}
                          placeholder={aiProvider === 'gemini' ? 'AIzaSy...' : 'sk-proj-...'}
                          className="w-full text-xs rounded border border-slate-200 p-2 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                        />
                        <p className="text-[10px] text-slate-400 leading-tight">
                          {aiApiKey ? '✓ Chave salva localmente' : 'Se preferir, insira GEMINI_API_KEY ou OPENAI_API_KEY nas variáveis de ambiente do backend (.env).'}
                        </p>
                      </div>

                      {/* Save Key Button */}
                      <Button
                        onClick={() => {
                          if (typeof window !== 'undefined') {
                            localStorage.setItem('dataclean_ai_key', aiApiKey);
                            localStorage.setItem('dataclean_ai_provider', aiProvider);
                          }
                          toast({
                            title: 'IA Configurada!',
                            description: 'Suas credenciais foram salvas localmente no navegador com segurança.',
                          });
                        }}
                        className="w-full text-xs py-1.5 bg-slate-800 hover:bg-slate-900 text-white shadow-sm"
                        size="sm"
                      >
                        Salvar Ajustes
                      </Button>
                      
                      <Separator />
                      
                      {/* Safety Info */}
                      <div className="space-y-2 text-[10px] text-slate-400 leading-relaxed bg-slate-50 p-2 rounded border">
                        <p className="font-semibold text-slate-600 flex items-center gap-1">
                          <Bot className="h-3 w-3 text-emerald-600" />
                          Segurança e Performance
                        </p>
                        <p>A IA cria o script de transformação baseado em apenas 3 linhas de exemplo (amostra). Os dados completos são transformados localmente no seu computador, protegendo a sua privacidade.</p>
                      </div>
                    </CardContent>
                  </Card>
                </div>
              </TabsContent>
            </Tabs>

            {/* Column Summary */}
            <Card className="border-0 shadow-sm">
              <CardHeader className="pb-3">
                <CardTitle className="text-base">Resumo das Colunas</CardTitle>
                <CardDescription className="text-xs">
                  Tipos detectados e estatísticas por coluna
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                  {result.columns.map(col => {
                    const colLogs = result.logs.filter(l => l.column === col);
                    const emptyCount = result.originalPreview.filter(
                      row => row[col] === null || row[col] === undefined || row[col] === ''
                    ).length;
                    return (
                      <div key={col} className="rounded-lg border bg-white p-3 space-y-2">
                        <div className="flex items-center justify-between">
                          <code className="text-sm font-semibold text-slate-800">{col}</code>
                          <Badge variant="outline" className="text-[10px]">
                            {colLogs.length} alterações
                          </Badge>
                        </div>
                        {emptyCount > 0 && (
                          <p className="text-xs text-slate-400">
                            {emptyCount} valores vazios na preview
                          </p>
                        )}
                        {colLogs.length > 0 && (
                          <div className="flex flex-wrap gap-1">
                            {Object.entries(logTypeConfig)
                              .filter(([key]) => key !== 'info')
                              .map(([key, config]) => {
                                const count = colLogs.filter(l => l.type === key).length;
                                if (count === 0) return null;
                                return (
                                  <span key={key} className={`text-[10px] px-1.5 py-0.5 rounded ${config.color}`}>
                                    {config.label}: {count}
                                  </span>
                                );
                              })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </CardContent>
            </Card>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t bg-white mt-auto">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <p className="text-center text-xs text-slate-400">
            DataClean Pro • Sistema Automatizado de Limpeza de Dados • CSV & XLSX
          </p>
        </div>
      </footer>
    </div>
  );
}
