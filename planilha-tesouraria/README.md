# Planilha de Tesouraria — Igreja

Planilha Excel profissional para controle financeiro de tesouraria paroquial.

## Arquivo principal

**`Tesouraria_Igreja.xlsx`** — abra este arquivo no Microsoft Excel ou Google Sheets.

## Estrutura (9 abas)

| # | Aba | Descrição |
|---|-----|-----------|
| 1 | **Dashboard** | Página inicial com KPIs e 6 gráficos automáticos |
| 2 | **Dízimos Mensais** | Registro de dízimos com mês calculado automaticamente |
| 3 | **Doações** | Registro de doações com tipos padronizados |
| 4 | **Arrecadação das Missas** | 52 domingos de 2026 pré-preenchidos |
| 5 | **Gastos** | Despesas por categoria com listas suspensas |
| 6 | **Resumo Financeiro** | Consolidado mensal automático (Jan–Dez) |
| 7 | **Plano de Contas** | Referência de receitas e despesas |
| 8 | **Fluxo de Caixa** | Saldo atualizado a cada lançamento |
| 9 | **Configurações** | Ano, saldo inicial e listas de validação |

## Layout

- **Cores:** azul-marinho, dourado e branco
- **Células amarelas:** campos de entrada (editáveis)
- **Células azul-claro:** fórmulas (protegidas)
- **Filtros** em todas as tabelas de lançamentos

## Como usar

1. Abra a aba **Configurações** e ajuste o **Ano** e o **Saldo Inicial**
2. Registre lançamentos nas abas de Dízimos, Doações, Missas e Gastos
3. O **Resumo Financeiro** e o **Dashboard** atualizam automaticamente
4. No **Fluxo de Caixa**, ordene pela coluna **Data** para ver cronologicamente

## Regenerar a planilha

```bash
pip install -r requirements.txt
python gerar_planilha.py
```

## Requisitos

- Microsoft Excel 2016+ (fórmulas em português: SOMA, SOMASE, SOMASES, SE, etc.)
- Compatível com LibreOffice Calc (pode exigir ajuste de nomes de funções)
