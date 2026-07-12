# Planilha de Tesouraria — Sistema Completo para Igreja

Planilha Excel profissional para gestão financeira paroquial, com controle de caixa/banco, conciliação, relatórios, alertas e prestação de contas.

## Arquivo principal

**`Tesouraria_Igreja.xlsx`** — abra no Microsoft Excel (recomendado: versão 2016+).

## Estrutura (18 abas visíveis)

| Aba | Função |
|-----|--------|
| **Dashboard** | KPIs, alertas, conciliação, gráficos e filtros |
| **Cadastro Dizimistas** | Nome, telefone, endereço, aniversário |
| **Pesquisa Dizimista** | Busca rápida por nome |
| **Dízimos Mensais** | Lançamentos com conta (Caixa/Banco) |
| **Doações** | Registro de doações |
| **Arrecadação das Missas** | Domingos do ano pré-preenchidos |
| **Gastos** | Despesas por categoria com alertas de limite |
| **Contas Caixa e Banco** | Saldos separados por conta |
| **Resumo Financeiro** | Consolidado mensal com metas e % |
| **Conciliação Bancária** | Conferência planilha × extrato |
| **Relatório Anual** | Prestação de contas automática |
| **Relatório Despesas** | Gastos por categoria e mês |
| **Comparativo Anual** | 2024 × 2025 × 2026 |
| **Plano de Contas** | Referência de receitas e despesas |
| **Fluxo de Caixa** | Saldo atualizado automaticamente |
| **Impressão A4** | Relatórios prontos para imprimir |
| **Configurações** | Ano, igreja, saldos, metas, limites, listas |
| **Novo Exercício** | Instruções + macro VBA |

## Recursos implementados

### 1. Saldo Inicial (Caixa + Banco)
Na aba **Configurações**, informe saldos separados de Caixa e Banco. O total alimenta o Dashboard e o Fluxo de Caixa.

### 2. Conciliação Bancária
Aba dedicada com colunas Valor Planilha, Valor Extrato, Diferença e Conferido (✅/❌). KPIs no Dashboard.

### 3. Relatório Anual
Gerado automaticamente na aba **Relatório Anual** — receitas, despesas por categoria e saldo final.

### 4. Alertas Inteligentes
- 🔴 Gastos > entradas do mês
- 🔴 Saldo negativo
- 🟢 Meta mensal atingida
- 🟡 Categoria acima do limite (formatação condicional em Gastos)

### 5. Impressão A4
Aba **Impressão A4** com prestação mensal, anual, resumo e demonstrativo — cabeçalho com nome da igreja, período, data e assinaturas.

### 6. Novo Exercício
- **Com macro:** importe `vba/NovoExercicio.bas` (Alt+F11) e associe a um botão
- **Sem macro:** `python gerar_planilha.py --novo-exercicio`

### 7. Recursos adicionais
- 📅 Validação de datas (calendário nativo do Excel)
- 🔎 Pesquisa por dizimista
- 👥 Cadastro completo de dizimistas
- 💳 Controle Caixa/Banco separado
- 📈 Comparativo entre anos
- 🎯 Meta mensal com % de cumprimento
- 📑 Relatório por categoria de despesas
- 💾 Dica de backup via OneDrive/Google Drive
- 🔒 Fórmulas protegidas, entradas liberadas
- 📊 Dashboard com filtro de mês

## Primeiros passos

1. **Configurações** → Nome da igreja, ano, saldos Caixa/Banco, metas e limites
2. **Cadastro Dizimistas** → Cadastre os dizimistas
3. Lance dados em Dízimos, Doações, Missas e Gastos (selecione Caixa ou Banco)
4. Acompanhe pelo **Dashboard** e imprima pela aba **Impressão A4**

## Regenerar

```bash
pip install -r requirements.txt
python gerar_planilha.py
python gerar_planilha.py --novo-exercicio   # avança para o próximo ano
```

## Requisitos

- Microsoft Excel 2016+ (fórmulas em português)
- Excel 365 recomendado para Pesquisa Dizimista (função FILTRAR)
- Macro VBA opcional para botão Novo Exercício (salvar como `.xlsm`)
