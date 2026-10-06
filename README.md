# Dashboard de Vendas Xbox Game Pass

Dashboard em Excel que transforma a base de assinaturas do Xbox Game Pass (2024) em indicadores e gráficos.

## Arquivos

| Arquivo | Descrição |
|---|---|
| `dashboard_xbox.xlsx` | Planilha final com o dashboard |
| `base.xlsx` | Base de dados original (modelo do desafio) |
| `build_dashboard.py` | Script que gera o dashboard a partir da base |

## Dados utilizados

Aba **Bases** (tabela `Tabela1`): 295 assinantes, de 01/01/2024 a 16/12/2024.

| Coluna | Significado |
|---|---|
| ID do Assinante, Nome | Identificação do assinante |
| Plano | Core, Standard ou Ultimate |
| Data de Início | Data de início da assinatura |
| Renovação Automática | Sim ou Não |
| Preço da Assinatura / Tipo de Assinatura | Preço do plano e periodicidade (Mensal, Trimestral, Anual) |
| Passe de Temporada EA Play (+ Preço) | Passe adicional EA Play |
| Passe de Temporada Minecraft (+ Preço) | Passe adicional Minecraft |
| Valor do Cupom | Desconto de cupom |
| Valor Total | Valor total (plano + passes - cupom) |

Observação: o `base.xlsx` original está em inglês. O script traduz cabeçalhos e valores (Yes/No, Monthly/Quarterly/Annual) ao gerar o dashboard.

## Estrutura do Excel

- **Assets**: paleta de cores, logos e ícones do modelo.
- **Bases**: dados brutos.
- **Cálculos**: tabelas de apoio com fórmulas (`COUNTIF`, `SUMIF`, `COUNTIFS`, `SUMIFS`) ligadas à aba Bases. Se a base mudar, o dashboard atualiza.
- **Dashboard**: cartões de indicadores e gráficos.

## O que o dashboard mostra

Indicadores (valores conferidos com pandas):

- Assinantes: 295
- Receita total: US$ 7.633
- Ticket médio: US$ 25,87
- Total em cupons: US$ 2.122
- Renovação automática: 50,2%

Gráficos: receita por plano, assinantes por tipo de assinatura, renovação automática, novas assinaturas por mês, receita dos passes adicionais e receita por tipo de assinatura.

## Como reproduzir

1. Instale as dependências: `pip install openpyxl pillow`.
2. Coloque `base.xlsx` na mesma pasta do script.
3. Execute `python build_dashboard.py`. Será gerado `dashboard_xbox.xlsx`.
4. Abra no Excel. As fórmulas são calculadas ao abrir.

Para montar manualmente no Excel, crie as tabelas da aba Cálculos com as mesmas fórmulas e insira os gráficos a partir delas.

## Limitações

- Não há segmentações de dados (slicers), que o openpyxl não cria. Elas podem ser adicionadas no Excel sobre `Tabela1`.
- O mês de dezembro tem dados só até 16/12, então aparece menor.
