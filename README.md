# DataClean Pro

Sistema automatizado de limpeza, padronização e processamento de dados em arquivos CSV e XLSX.

## Funcionalidades

### Pipeline de Limpeza Automática
- **Padronização de Colunas** — Normaliza nomes de colunas (remove acentos, espaços, caracteres especiais)
- **Correção de Erros** — Identifica e corrige formatação incorreta, valores inválidos e caracteres especiais
- **Remoção de Duplicatas** — Elimina linhas duplicadas mantendo a primeira ocorrência
- **Preenchimento Inteligente** — Preenche campos vazios com mediana, valor mais frequente ou interpolação
- **Divisão de Colunas Compostas** — Detecta dados compostos separados por `/`, `|` ou `;` e divide em duas colunas

### Detecções Automáticas
- Colunas de data por nome e formato
- Colunas numéricas com formatação brasileira (1.234,56) e americana (1,234.56)
- Colunas compostas com delimitadores consistentes
- Caracteres invisíveis e especiais (zero-width, BOM, non-breaking spaces)

## Tecnologias

- **Next.js 16** — Framework React com App Router
- **TypeScript** — Tipagem estática
- **Tailwind CSS 4** — Estilização utilitária
- **shadcn/ui** — Componentes de UI acessíveis
- **PapaParse** — Parseamento de CSV
- **SheetJS (xlsx)** — Leitura de arquivos XLSX

## Primeiros Passos

### Pré-requisitos
- Node.js 18+ ou Bun

### Instalação

```bash
git clone https://github.com/SEU_USUARIO/dataclean-pro.git
cd dataclean-pro
npm install
```

### Executando

```bash
# Modo desenvolvimento
npm run dev

# Build de produção
npm run build

# Iniciar em produção
npm run start
```

O aplicativo estará disponível em `http://localhost:3000`.

## Como Usar

1. **Upload do Arquivo** — Arraste e solte ou selecione um arquivo CSV/XLSX
2. **Processar** — Clique em "Processar e Limpar Dados"
3. **Revisar** — Veja a prévia dos dados limpos e o log de todas as correções
4. **Baixar** — Baixe o arquivo processado no formato original

## Estrutura do Projeto

```
dataclean-pro/
├── src/
│   ├── app/
│   │   ├── page.tsx              # Página principal com UI completa
│   │   ├── layout.tsx            # Layout raiz com fontes e metadata
│   │   ├── globals.css           # Estilos globais e tema
│   │   └── api/
│   │       ├── process/route.ts  # API de processamento de arquivos
│   │       └── download/route.ts # API de download dos resultados
│   ├── lib/
│   │   ├── data-cleaner.ts       # Motor principal de limpeza de dados
│   │   ├── db.ts                 # Configuração do banco
│   │   └── utils.ts              # Utilitários gerais
│   ├── components/ui/            # Componentes shadcn/ui
│   └── hooks/                    # Custom React hooks
├── public/                       # Arquivos estáticos
├── prisma/                       # Schema do banco
├── package.json
├── tsconfig.json
├── next.config.ts
└── tailwind.config.ts
```

## Pipeline de Processamento

```
Arquivo CSV/XLSX
       ↓
[1] Padronização de Colunas
       ↓
[2] Divisão de Colunas Compostas
       ↓
[3] Detecção e Correção de Erros
       ↓
[4] Remoção de Duplicatas
       ↓
[5] Preenchimento de Campos Vazios
       ↓
Exportação (XLSX / CSV)
```

## Licença

Este projeto é de uso pessoal. Todos os direitos reservados.
