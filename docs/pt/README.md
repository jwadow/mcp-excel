<div align="center">

# 📊 Excel MCP Server

**Análise rápida e eficiente de planilhas através de operações atômicas, construído especificamente para agentes de IA**

[🇬🇧 English](../../README.md) • [🇷🇺 Русский](../ru/README.md) • [🇨🇳 中文](../zh/README.md) • [🇪🇸 Español](../es/README.md) • [🇯🇵 日本語](../ja/README.md) • 🇧🇷 Português

Feito com ❤️ por [@Jwadow](https://github.com/jwadow)

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)
[![Sponsor](https://img.shields.io/badge/💖_Apoiar-Desenvolvimento-ff69b4)](#-apoie-o-projeto)

**Analise planilhas Excel com seu agente de IA através de operações atômicas — sem despejar dados no contexto**

*Funciona com OpenCode, Claude Code, Codex app, Cursor, Cline, Roo Code, Kilo Code e outros agentes de IA compatíveis com MCP*

[Por que isso existe](#-por-que-isso-existe) • [O que seu agente pode fazer](#-o-que-seu-agente-pode-fazer) • [Instalação e configuração](#%EF%B8%8F-instalação-e-configuração) • [Ferramentas disponíveis](#%EF%B8%8F-ferramentas-disponíveis) • [💖 Doar](#-apoie-o-projeto)

</div>

---

## 🤨 Por que isso existe

**O problema:** A maioria das ferramentas Excel para IA despeja dados brutos da planilha no contexto do agente. Isso inunda a janela de contexto, deixa tudo lento, e a IA ainda pode calcular errado ou se confundir em grandes conjuntos de dados.

**Este projeto:** Pense em SQL para Excel. Seu agente de IA compõe operações atômicas (`filter_and_count`, `aggregate`, `group_by`) e obtém resultados precisos — não milhares de linhas.

O agente analisa dados **sem vê-los**. Os resultados chegam como números, fórmulas e insights.

> *"Isso é como trabalhar com um banco de dados através de SQL, não arrastando tudo para a memória."*
> — Agente de IA após analisar uma planilha em produção

### 🔌 O que é MCP?

[Model Context Protocol](https://modelcontextprotocol.io) é um padrão aberto que permite aos agentes de IA usar ferramentas externas.

Este projeto é uma dessas ferramentas. Quando você conecta este servidor ao seu agente de IA (OpenCode, Claude Code, Codex app, Cursor, Cline, Roo Code, Kilo Code, etc.), seu agente ganha um monte de novos comandos para trabalhar com arquivos Excel — filtragem, contagem, agregação, análise.

**A vantagem principal:** Sua IA não carrega milhares de linhas de planilha na memória. Em vez disso, faz perguntas específicas e obtém respostas precisas. Mais rápido, mais preciso, sem estouro de contexto.

---

## 💬 O que os agentes de IA dizem

Feedback real de agentes de IA que usaram este servidor MCP em produção:

> *"Analisei 34.211 linhas sem carregar dados no contexto. Cada operação retorna apenas o resultado — contagem, soma, média. O contexto permanece limpo. As operações executam em 25-45ms independentemente do tamanho do arquivo."*

> *"Isso é SQL para Excel. Consultas, filtros, agregação — sem despejar dados no contexto. Ferramenta sólida para tarefas analíticas."*

> *"O sistema de filtros lida bem com lógica complexa. Grupos AND/OR aninhados, 12 operadores, condições ilimitadas. Construí uma classificação multicategoria sem escrever código."*

> *"As operações em lote são eficientes. Uma chamada `filter_and_count_batch` em vez de múltiplas solicitações separadas. O arquivo carrega uma vez, todos os filtros são aplicados, os resultados chegam juntos."*

*Sim, agentes agora escrevem avaliações. Estas são reflexões reais de agentes de IA analisando dados de planilhas do mundo real. Bem-vindo a 2026.*

---

## 🚀 O que seu agente pode fazer

Uma vez conectado, seu agente de IA ganha um monte de ferramentas especializadas para analisar dados tabulares. O agente recebe apenas consultas precisas e resultados confiáveis.

### 📊 Exploração de dados
- **Inspecionar arquivos** - estrutura, planilhas, colunas, tipos de dados (detecta automaticamente cabeçalhos bagunçados)
- **Perfilar colunas** - estatísticas, contagens de nulos, valores principais, qualidade de dados em uma chamada
- **Buscar dados** - pesquisar em múltiplas planilhas, localizar colunas em qualquer lugar

### 🔍 Filtragem e consultas
- **12 operadores de filtro** - `==`, `!=`, `>`, `<`, `>=`, `<=`, `in`, `not_in`, `contains`, `startswith`, `endswith`, `regex`
- **Lógica complexa** - grupos AND/OR aninhados, operador NOT, condições ilimitadas
- **Operações em lote** - classificar dados em múltiplas categorias em uma solicitação (6x mais rápido)
- **Análise de sobreposição** - diagramas de Venn, contagens de interseção, operações de conjuntos

### 📈 Agregação e análise
- **8 funções de agregação** - sum, mean, median, min, max, std, var, count
- **Agrupar por** - tabelas dinâmicas com múltiplas colunas de agrupamento
- **Análise estatística** - correlações (Pearson/Spearman/Kendall), detecção de outliers (IQR/Z-score)
- **Séries temporais** - crescimento período a período, médias móveis, totais acumulados

### 🏆 Operações avançadas
- **Classificação** - top-N, bottom-N, classificação por percentil (com suporte a agrupamento)
- **Colunas calculadas** - expressões aritméticas entre colunas
- **Validação de dados** - encontrar duplicatas, valores nulos, verificações de qualidade de dados
- **Comparação de planilhas** - diferenças entre versões, encontrar mudanças

### ⚡ Recursos de desempenho
- **Operações atômicas** - resultados em 20-50ms, independentemente do tamanho do arquivo
- **Cache inteligente** - arquivo carregado uma vez, reutilizado para todas as operações
- **Linhas de amostra** - visualizar dados filtrados sem recuperação completa
- **Proteção de contexto** - limites inteligentes previnem estouro do contexto de IA

### 📋 Integração com Excel
- **Geração de fórmulas** - cada resultado inclui fórmula Excel para atualizações dinâmicas
- **Saída TSV** - copiar-colar resultados diretamente no Excel
- **Suporte legado** - funciona com arquivos .xls antigos (Excel 97-2003)
- **Multi-planilha** - analisar múltiplas planilhas em um arquivo

**Exemplos de consultas que seu agente agora pode lidar:**
- *"Mostre os 10 melhores clientes por receita"*
- *"Encontre todos os pedidos do Q4 onde o valor > R$1000"*
- *"Calcule o crescimento mês a mês para cada categoria de produto"*
- *"Quais clientes são VIP e ativos? (análise de sobreposição)"*
- *"Encontre duplicatas na coluna Email"*

## ⚙️ Instalação e configuração

### Pré-requisitos

**Python 3.10 ou superior** — [Baixar aqui](https://www.python.org/downloads/)

### Passo 1: Clonar repositório

```bash
git clone https://github.com/jwadow/mcp-excel.git
cd mcp-excel
```

*Não tem Git? Clique em "Code" → "Download ZIP" no topo desta página do repositório, extraia e abra o terminal nessa pasta.*

### Passo 2: Escolher método de instalação

<details>
<summary><b>🎯 Opção A: Poetry (Recomendado)</b></summary>

Poetry é um gerenciador de dependências moderno do Python (substitui pip+venv+requirements.txt).
[Instale-o](https://python-poetry.org/docs/#installation): `pip install poetry` ou `pipx install poetry`

**Instalar dependências:**
```bash
poetry install
```

**Configurar seu agente de IA:**

Adicione isso às suas configurações MCP (config JSON):
```json
{
  "mcpServers": {
    "excel": {
      "command": "poetry",
      "args": ["run", "python", "-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

**Importante:** Substitua `C:/path/to/mcp-excel` pelo caminho real para o repositório clonado.

</details>

<details>
<summary><b>📦 Opção B: pip com ambiente virtual</b></summary>

**Instalar dependências:**
```bash
# Windows
python -m venv venv
venv\Scripts\activate
pip install -e .

# Linux/Mac
python -m venv venv
source venv/bin/activate
pip install -e .
```

**Encontrar caminho do Python no venv:**
```bash
# Windows
where python

# Linux/Mac
which python
```

**Configurar seu agente de IA:**

Adicione isso às suas configurações MCP (config JSON):
```json
{
  "mcpServers": {
    "excel": {
      "command": "C:/path/to/mcp-excel/venv/Scripts/python.exe",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

**Importante:**
- Substitua `C:/path/to/mcp-excel/venv/Scripts/python.exe` pelo caminho real do comando `where python`
- No Linux/Mac use o caminho de `which python` (ex: `/path/to/mcp-excel/venv/bin/python`)

</details>

<details>
<summary><b>🐍 Opção C: Python do sistema (Não recomendado)</b></summary>

**Instalar dependências globalmente:**
```bash
pip install "mcp>=1.1.0" "pandas>=2.2.0" "pydantic>=2.10.0" "xlrd>=2.0.1" "openpyxl>=3.1.0" "psutil>=6.1.0" "python-dateutil>=2.9.0"
```

**Configurar seu agente de IA:**
```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

⚠️ **Aviso:** Isso polui seu ambiente Python global. Use Poetry ou venv em vez disso.

</details>

### Passo 3: Verificar instalação

Reinicie seu agente de IA e teste:
```
"Analise o arquivo Excel em C:/Users/SeuNome/Documents/test.xlsx"
```

Se funcionar - pronto! Se não, verifique:
- O caminho para o repositório está correto em `cwd`
- O caminho do Python está correto em `command` (para método pip)
- Todas as dependências estão instaladas

### Agentes de IA suportados

Funciona com qualquer agente de IA compatível com MCP.

⚠️ **Importante:** Este é um servidor MCP. Ele executa automaticamente quando seu agente de IA precisa. Não execute manualmente no terminal.

## 💡 Uso

Após a configuração, reinicie seu agente de IA e peça para analisar arquivos Excel:

```
"Analise o arquivo Excel em C:/Users/SeuNome/Documents/sales.xls"
"Mostre os 10 melhores clientes por receita de sales.xlsx"
"Encontre duplicatas na coluna 'Email' em contacts.xlsx"
"Calcule o crescimento mês a mês de revenue.xls"
```

## 🛠️ Ferramentas disponíveis

<details>
<summary><b>📋 Referência completa de ferramentas (25 ferramentas) - Clique para expandir</b></summary>

### 📊 Inspeção de arquivos (5 ferramentas)

#### `inspect_file`
Obter visão geral da estrutura do arquivo - planilhas, dimensões, formato.
**Usar para:** Exploração inicial do arquivo, descoberta de planilhas, validação de formato
**Retorna:** Lista de planilhas, contagens de linhas/colunas, metadados do arquivo

#### `get_sheet_info`
Análise detalhada da planilha com detecção automática de cabeçalhos.
**Usar para:** Entender estrutura de dados, tipos de colunas, visualização de amostras
**Retorna:** Nomes/tipos de colunas, contagem de linhas, dados de amostra (3 linhas), info de detecção de cabeçalhos

#### `get_column_names`
Enumeração rápida de colunas sem carregar dados completos.
**Usar para:** Validação de esquema, construção de filtros, verificação de disponibilidade de colunas
**Retorna:** Lista de nomes de colunas, contagem de colunas

#### `get_data_profile`
Perfilamento abrangente de colunas - tipos, estatísticas, nulos, valores principais.
**Usar para:** Exploração inicial de dados, avaliação de qualidade, análise de distribuição
**Retorna:** Por coluna: tipo, % nulos, contagem única, estatísticas (numérico), top N valores
**Eficincia:** Substitui 10+ chamadas separadas (get_column_stats + get_value_counts + find_nulls)

#### `find_column`
Localizar coluna em múltiplas planilhas.
**Usar para:** Navegação multi-planilha, descoberta de dados, análise entre planilhas
**Retorna:** Lista de planilhas com localizações de colunas, índices, contagens de linhas (sem distinção de maiúsculas)

---

### 📥 Recuperação de dados (3 ferramentas)

#### `get_unique_values`
Extrair valores únicos de uma coluna.
**Usar para:** Exploração de dados, construção de filtros, descoberta de valores distintos, verificações de qualidade de dados
**Retorna:** Lista de valores únicos, contagem, flag de truncamento (se exceder o limite)
**Limite padrão:** 100 valores

#### `get_value_counts`
Análise de frequência - top N valores mais comuns.
**Usar para:** Análise de distribuição, identificar categorias dominantes, detecção de desequilíbrio de dados
**Retorna:** Dicionário valor → contagem, contagem total, saída TSV
**Padrão:** Top 10 valores

#### `filter_and_get_rows`
Recuperar linhas filtradas com paginação.
**Usar para:** Extração de dados, inspeção de amostras, análise detalhada, exportação
**Retorna:** Linhas filtradas (lista de dicionários), contagem total, saída TSV
**Paginação:** Suporte a limit/offset

---

### 🔍 Filtragem e contagem (3 ferramentas)

#### `filter_and_count`
Contar linhas que correspondem a condições com 14 operadores.
**Operadores:** `==`, `!=`, `>`, `<`, `>=`, `<=`, `in`, `not_in`, `contains`, `startswith`, `endswith`, `regex`, `is_null`, `is_not_null`
**Lógica:** Grupos AND/OR aninhados, operador NOT, condições ilimitadas
**Usar para:** Classificação, segmentação, validação de dados, contagem de categorias
**Retorna:** Contagem + fórmula Excel (COUNTIFS), linhas de amostra opcionais

#### `filter_and_count_batch`
Classificar dados em múltiplas categorias em uma chamada (6x mais rápido).
**Usar para:** Classificação multicategoria, segmentação de mercado, controle de qualidade
**Retorna:** Contagem + fórmula por categoria, tabela TSV para Excel
**Eficiência:** Carrega arquivo uma vez, aplica todos os filtros, retorna todos os resultados

#### `analyze_overlap`
Análise de diagrama de Venn - interseções, uniões, zonas exclusivas.
**Usar para:** Análise de sobreposição, oportunidades de venda cruzada, verificações de consistência de dados
**Retorna:** Contagens de conjuntos, interseções aos pares (A ∩ B), união, dados de Venn (2-3 conjuntos)
**Exemplos:** Clientes VIP E ativos, sobreposições de categorias de produtos, pedidos concluídos SEM data de conclusão

---

### 📈 Agregação e análise (2 ferramentas)

#### `aggregate`
Realizar agregação com filtros opcionais (8 operações).
**Operações:** `sum`, `mean`, `median`, `min`, `max`, `std`, `var`, `count`
**Usar para:** Totais, médias, valores mín/máx, resumos estatísticos, agregações condicionais, cálculos de KPI
**Retorna:** Valor agregado + fórmula Excel (SUMIF, AVERAGEIF, etc.)
**Especial:** Autoconversão de números armazenados como texto para numérico

#### `group_by`
Tabela dinâmica com agrupamento de múltiplas colunas.
**Usar para:** Análise de categorias, agrupamento hierárquico, vendas por região/produto
**Retorna:** Dados agrupados com valores agregados, saída TSV
**Suporta:** Múltiplas colunas de agrupamento, todas as 8 operações de agregação

---

### 📊 Estatísticas (3 ferramentas)

#### `get_column_stats`
Resumo estatístico - contagem, média, mediana, desvio padrão, quartis.
**Usar para:** Análise de distribuição, perfilamento de dados, preparação para detecção de outliers
**Retorna:** Estatísticas completas (min, max, mean, median, std, Q1, Q3), contagem de nulos, saída TSV

#### `correlate`
Matriz de correlação entre 2+ colunas.
**Métodos:** Pearson (linear), Spearman (baseado em classificação), Kendall (baseado em classificação)
**Usar para:** Análise de relacionamentos, dependência de variáveis, seleção de características
**Retorna:** Matriz de correlação (-1 a 1), saída TSV

#### `detect_outliers`
Detecção de anomalias usando método IQR ou Z-score.
**Métodos:** IQR (robusto), Z-score (assume distribuição normal)
**Usar para:** Detecção de fraude, erros de sensores, qualidade de dados, identificação de valores incomuns
**Retorna:** Linhas outliers com índices, contagem, método/limiar usado

---

### ✅ Validação de dados (2 ferramentas)

#### `find_duplicates`
Detectar linhas duplicadas por colunas especificadas.
**Usar para:** Qualidade de dados, planejamento de deduplicação, verificações de integridade
**Retorna:** Todas as linhas duplicadas (incluindo primeira ocorrência), contagem, índices
**Nota:** Usa `duplicated(keep=False)` para marcar todas as duplicatas

#### `find_nulls`
Encontrar valores nulos/vazios com estatísticas detalhadas.
**Usar para:** Verificações de completude, análise de valores ausentes, limpeza de dados
**Retorna:** Por coluna: contagem de nulos, porcentagem, índices (primeiros 100)
**Nota:** Marcadores de posição (".", "-") NÃO são nulos - use operadores `==` ou `in`

---

### 🔄 Operações multi-planilha (2 ferramentas)

#### `search_across_sheets`
Buscar valor em todas as planilhas.
**Usar para:** Busca entre planilhas, rastreamento de valores, localização de dados
**Retorna:** Lista de planilhas com contagens de correspondências, correspondências totais
**Suporta:** Valores numéricos e de string

#### `compare_sheets`
Diferença entre duas planilhas usando coluna chave.
**Usar para:** Comparação de versões, detecção de mudanças, reconciliação, trilhas de auditoria
**Retorna:** Linhas com diferenças, status (only_in_sheet1/sheet2/different_values), comparação lado a lado

---

### 📅 Séries temporais (3 ferramentas)

#### `calculate_period_change`
Análise de crescimento período a período.
**Períodos:** month, quarter, year
**Usar para:** Análise de tendências, rastreamento de crescimento, comparação sazonal, análise ano a ano
**Retorna:** Períodos com valores, mudanças absolutas/percentuais, fórmula Excel

#### `calculate_running_total`
Soma acumulada com agrupamento opcional.
**Usar para:** Análise acumulativa, rastreamento de progresso, cálculos de saldo, fluxo de caixa
**Retorna:** Linhas com totais acumulados, fórmula Excel (SUM($B$2:B2))
**Suporta:** Agrupamento (total acumulado reinicia por grupo)

#### `calculate_moving_average`
Suavização com tamanho de janela especificado.
**Usar para:** Detecção de tendências, redução de ruído, identificação de padrões
**Retorna:** Linhas com médias móveis, fórmula Excel (AVERAGE(B1:B7))
**Exemplos:** Média móvel de 7 dias, suavização de preço de ações de 30 dias

---

### 🏆 Operações avançadas (2 ferramentas)

#### `rank_rows`
Classificar por valor de coluna com filtragem top-N.
**Direções:** desc (maior primeiro), asc (menor primeiro)
**Usar para:** Tabelas de classificação, análise top/bottom, classificação por percentil
**Retorna:** Linhas classificadas com números de classificação, fórmula Excel (RANK)
**Suporta:** Filtragem top-N, classificação dentro de grupos

#### `calculate_expression`
Expressões aritméticas entre colunas.
**Operações:** `+`, `-`, `*`, `/`, parênteses
**Usar para:** Métricas derivadas, cálculos financeiros, análise de proporções, cálculos de KPI
**Retorna:** Valores calculados, fórmula Excel (ex: =A2*B2)
**Exemplos:** Receita = Preço * Quantidade, Margem = (Receita - Custo) / Receita

</details>

## 🗺️ Roteiro

### 📁 Suporte a formatos de arquivo

**Atualmente suportado:**
- ✅ **XLS** - Excel 97-2003 (somente leitura)
- ✅ **XLSX** - Excel 2007+ (somente leitura)

**Planejado:**
- 🔜 **XLSM** - Excel com suporte a macros
- 🔜 **CSV** - Valores separados por vírgula
- 🔜 **TSV** - Valores separados por tabulação
- 🔜 **ODS** - Planilha OpenDocument
- 🔜 **Parquet** - Formato de armazenamento colunar

### 🚀 Recursos

- **Operações de escrita** - Modificar arquivos de planilhas (criar colunas calculadas, atualizar valores)
- **Modo de transporte SSE** - Eventos enviados pelo servidor para acesso remoto
- **Geração avançada de fórmulas** - Fórmulas Excel mais complexas com funções aninhadas
- **Exportação de dados** - Exportar resultados filtrados/agregados para novos arquivos

---

## 📜 Licença

Este projeto está licenciado sob a **GNU Affero General Public License v3.0 (AGPL-3.0)**.

Isso significa:
- ✅ Você pode usar, modificar e distribuir este software
- ✅ Você pode usá-lo para fins comerciais
- ⚠️ **Você deve divulgar o código-fonte** ao distribuir o software
- ⚠️ **Uso em rede é distribuição** — se você executar uma versão modificada em um servidor e permitir que outros interajam com ela, você deve disponibilizar o código-fonte
- ⚠️ As modificações devem ser lançadas sob a mesma licença

Consulte o arquivo [LICENSE](../../LICENSE) para o texto completo da licença.

### Por que AGPL-3.0?

AGPL-3.0 garante que melhorias neste software beneficiem toda a comunidade. Se você modificar este servidor e implantá-lo como serviço, deve compartilhar suas melhorias com seus usuários.

---

## 💖 Apoie o projeto

<div align="center">

<img src="https://raw.githubusercontent.com/Tarikul-Islam-Anik/Animated-Fluent-Emojis/master/Emojis/Smilies/Smiling%20Face%20with%20Hearts.png" alt="Love" width="80" />

**Se este projeto economizou seu tempo ou dinheiro, considere apoiá-lo!**

Cada contribuição ajuda a manter o projeto vivo e crescendo

<br>

### 🤑 Doar

[**☕ Doação única**](https://app.lava.top/jwadow?tabId=donate) • [**💎 Apoio mensal**](https://app.lava.top/jwadow?tabId=subscriptions)

<br>

### 🪙 Ou envie cripto

| Moeda | Rede | Endereço |
|:--------:|:-------:|:--------|
| **USDT** | TRC20 | `TSVtgRc9pkC1UgcbVeijBHjFmpkYHDRu26` |
| **BTC** | Bitcoin | `12GZqxqpcBsqJ4Vf1YreLqwoMGvzBPgJq6` |
| **ETH** | Ethereum | `0xc86eab3bba3bbaf4eb5b5fff8586f1460f1fd395` |
| **SOL** | Solana | `9amykF7KibZmdaw66a1oqYJyi75fRqgdsqnG66AK3jvh` |
| **TON** | TON | `UQBVh8T1H3GI7gd7b-_PPNnxHYYxptrcCVf3qQk5v41h3QTM` |

</div>

---

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor, certifique-se de que:

1. Todas as dependências sejam compatíveis com AGPL
2. O código siga o estilo existente
3. Testes sejam incluídos para novos recursos
4. A documentação esteja atualizada

Para problemas, bugs ou contribuições, por favor abra uma issue no GitHub.

---

## 💬 Precisa de ajuda?

Tem perguntas? Encontrou um bug? Tem uma ideia de recurso? Estamos aqui para ajudar!

**👉 [Abrir uma Issue no GitHub](https://github.com/jwadow/mcp-excel/issues/new)**

Seja você está preso na instalação, encontrou algo quebrado ou apenas quer sugerir uma melhoria — GitHub Issues é o lugar. Não se preocupe se você é novo no GitHub, apenas clique no link acima e descreva sua situação. Vamos resolver juntos.

---

<div align="center">

**[⬆ Voltar ao topo](#-excel-mcp-server)**

</div>
