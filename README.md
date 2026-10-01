# 🚀 Desafio Técnico de Power BI & Analytics Engineering com Python

> **Candidato:** José Willian  
> **Objetivo:** Demonstração prática de conhecimentos avançados em Power Query (M), Modelagem Dimensional (Star Schema), Fórmulas DAX, Engenharia de Dados com Python e Visualização Executiva.

---

## 🌟 Visão Geral da Solução: O Diferencial dos 150%

Além de desenvolver o relatório oficial em **Power BI** (`Teste_PowerBI_JoseWillian.pbix`), este projeto foi além: foi criado um **Pipeline de Analytics em Python no VS Code** (`pipeline_analytics.py`) e um **Dashboard Web Interativo de Alto Nível** (`index.html`) que pode ser acessado de qualquer navegador e celular.

### 🌐 Link do Dashboard Online (Deploy)
* Acesse a versão web interativa com fichas técnicas DAX completas:  
  👉 **`https://jwillianteo.github.io/teste-powerbi-python/`**
* Ou abra localmente com duplo clique no arquivo [`index.html`](index.html).

---

## 💡 Destaque Sênior: A Análise Crítica dos Dados (A Armadilha do Item 1)

Durante a exploração dos dados com Python e Power Query, identificamos uma **situação crítica de negócio**:

1. Em `dCliente.xlsx`, existem **18 registros** onde a coluna `Cliente` está nula (`null`), porém com `ID Cliente` e `País` válidos.
2. Ao cruzar com a tabela fato `fVendas.xlsx`, constatamos que esses 18 clientes são responsáveis por **364 transações de venda (52,7% dos pedidos)** e **R$ 247.686,32 (96,0% de todo o faturamento da empresa!)**.
3. **Se eliminarmos cegamente as linhas com nulos** (como uma leitura ingênua do *Item 1* sugere), o faturamento por país na análise de metas desaba para apenas 4% do total real (Alemanha passa de R$ 146 mil para R$ 3,6 mil!).
4. **Solução de Engenharia Adotada:** 
   - No Power Query, criamos o parâmetro `pRemoverBrancos`: se `true`, aplica o descarte estrito; se `false`, preserva a integridade referencial substituindo nulos por `"Não informado"`.
   - No Dashboard Web, criamos um **seletor de cenários em tempo real** para que o avaliador possa testar e constatar o impacto das duas decisões!

---

## 📋 Matriz dos 9 Requisitos + Dica de Ouro

| # | Requisito | Solução no Power BI (M / DAX) | Solução em Python (Pandas) |
|---|---|---|---|
| **1** | **Importação & Limpeza** | Leitura via Excel.Workbook, tipagem estrita e tratamento de nulos parametrizado. | `pd.read_excel()` e diagnóstico de integridade com `.fillna('Não informado')`. |
| **2** | **Merge de Tabelas** | `Table.NestedJoin` (Left Outer Join) para trazer Cliente, País, Produto e Subcategoria para `fVendas`. | `df_vendas.merge(df_cliente, how='left').merge(df_produto, how='left')`. |
| **3** | **Coluna Calculada** | `DATEDIFF(fVendas[Data Venda], TODAY(), DAY)`. | `(pd.Timestamp.now().normalize() - df['Data Venda']).dt.days`. |
| **4** | **Cópia fVendas_Filtros** | Referência de consulta filtrando `DateTime.LocalNow() - 30 dias`. *(Explicado por que retorna vazio com dados de 2019)*. | Filtro booleano `df['Data Venda'] >= (hoje - 30 dias)` e versão relativa à data máxima. |
| **5** | **Agrupamento fVendas_Agg** | Criação de `Mês/Ano` via `Date.StartOfMonth` e agrupamento por Cliente, Produto e Mês com soma de vendas. | `df.groupby(['ID Cliente', 'ID Produto', 'MesAno'])['Valor Venda'].sum()`. |
| **6** | **Relacionamentos** | Star Schema 1:N com tabelas dimensão; criação da tabela ponte `dPais` para ligar `fMetas` e `fVendas`. | Modelo relacional e merge unificado com chaves primárias. |
| **7** | **KPIs em DAX** | `[Vendas]`, `[Total de Metas]` e `[% Atingimento] = DIVIDE([Lucro Bruto], [Total de Metas])`. | Funções de agregação e vetorização protegida contra divisão por zero. |
| **8** | **Gráficos e Visuais** | Gráfico de barras ordenado por produto e gráfico de linhas mensais estilizado com marcadores. | Chart.js interativo com barras horizontais e gráfico de linha com área sombreada. |
| **9** | **Performance MoM** | `CALCULATE([Vendas], DATEADD(dCalendario[Date], -1, MONTH))` e cálculo percentual. | `df.groupby('MesAno')['Valor Venda'].sum().pct_change() * 100`. |
| **⭐** | **Dica de Ouro (dCalendario)** | Tabela gerada via `CALENDAR(DATE(2017,1,1), DATE(2019,12,31))` com ordenação por `AnoMês`. | `pd.date_range()` cobrindo o período histórico completo. |

---

## 🛠️ Como Executar o Projeto Localmente

### 1. No Power BI:
1. Abra o arquivo [`Teste_PowerBI_JoseWillian.pbix`](Teste_PowerBI_JoseWillian.pbix).
2. Verifique o modelo de dados na guia **Modelo** e os visuais no painel principal.

### 2. O Pipeline Python:
```bash
# Executa a limpeza, valida as regras e atualiza os arquivos do dashboard
python pipeline_analytics.py
```

### 3. O Dashboard Web Interativo:
* Basta dar dois cliques no arquivo [`index.html`](index.html) para abrir no navegador, ou executar um servidor local:
```bash
python -m http.server 8000
```
Acesse: `http://localhost:8000`

---

## 🚀 Publicação no GitHub Pages
1. Repositório oficial: `https://github.com/jwillianteo/teste-powerbi-python`
2. No GitHub: acesse **Settings** > **Pages** > em *Branch*, selecione **main** e pasta **/ (root)** > clique em **Save**.
3. O link do dashboard estará no ar em:  
   👉 **`https://jwillianteo.github.io/teste-powerbi-python/`**
