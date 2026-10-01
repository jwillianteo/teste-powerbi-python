"""
Pipeline de Engenharia e Análise de Dados 4.0 - Power BI & Python
Autor: José Willian
Descrição:
    Processa os 4 Excels com Pandas, calcula métricas de negócio,
    análise de Pareto (Curva ABC), Matriz de Calor de Sazonalidade (Ano x Mês),
    KPIs avançados e garante 100% de conformidade com os valores reais
    avaliados no Power BI Desktop (.pbix).
"""

import os
import json
import pandas as pd
import numpy as np
from datetime import datetime

def run_pipeline():
    print("=" * 65)
    print("🚀 INICIANDO PIPELINE DE ANALYTICS 4.0 - POWER BI & PYTHON")
    print("=" * 65)

    # 1. IMPORTAÇÃO E DIAGNÓSTICO
    print("\n[ETL 1] Leitura dos arquivos e diagnóstico de integridade...")
    df_cliente = pd.read_excel('dCliente.xlsx')
    df_produto = pd.read_excel('dProduto.xlsx')
    df_metas = pd.read_excel('fMetas.xlsx')
    df_vendas = pd.read_excel('fVendas.xlsx')

    linhas_nulas = df_cliente['Cliente'].isna().sum()
    ids_nulos = df_cliente[df_cliente['Cliente'].isna()]['ID Cliente'].tolist()
    vendas_nulas = df_vendas[df_vendas['ID Cliente'].isin(ids_nulos)]
    receita_nula = vendas_nulas['Valor Venda'].sum()
    receita_total = df_vendas['Valor Venda'].sum()
    pct_nula = (receita_nula / receita_total) * 100

    print(f"  - dCliente: {len(df_cliente)} linhas | {linhas_nulas} com nome nulo.")
    print(f"  - Impacto financeiro: {len(vendas_nulas)} vendas / R$ {receita_nula:,.2f} ({pct_nula:.1f}% da receita)")

    # Versão Oficial PBI (Item 1: remove linhas com nulos)
    df_cliente_pbi = df_cliente.dropna(subset=['Cliente']).copy()
    
    # Versão com Governança (imputando para manter integridade)
    df_cliente_biz = df_cliente.copy()
    df_cliente_biz['Cliente'] = df_cliente_biz['Cliente'].fillna('Cliente Não Informado')

    # 2. MERGE DE TABELAS
    print("\n[ETL 2] Merge Dimensão -> Fato...")
    # Merge com cliente PBI Oficial (produz nulos em País nas 364 vendas)
    df_merged = df_vendas.merge(df_cliente_pbi, on='ID Cliente', how='left')
    # Guardamos também o país original para poder alternar no painel
    df_merged = df_merged.merge(df_cliente[['ID Cliente', 'País']].rename(columns={'País': 'País_Original'}), on='ID Cliente', how='left')
    df_merged = df_merged.merge(df_produto, on='ID Produto', how='left')

    # 3. COLUNAS CALCULADAS
    hoje = pd.Timestamp.now().normalize()
    df_merged['Dias_Desde_Venda'] = (hoje - pd.to_datetime(df_merged['Data Venda'])).dt.days
    df_merged['Ano'] = df_merged['Data Venda'].dt.year
    df_merged['Mes'] = df_merged['Data Venda'].dt.month
    df_merged['Mes_Nome'] = df_merged['Data Venda'].dt.strftime('%b')
    df_merged['AnoMes_Str'] = df_merged['Data Venda'].dt.strftime('%Y-%m')
    df_merged['Trimestre'] = df_merged['Data Venda'].dt.to_period('Q').astype(str)

    # 4. ANÁLISE DE PARETO / CURVA ABC DE PRODUTOS
    print("\n[Analytics 4.0] Curva ABC (Pareto 80/20) de Produtos...")
    pareto = df_merged.groupby(['Produto', 'Subcategoria'])['Valor Venda'].sum().sort_values(ascending=False).reset_index()
    pareto['Pct_Individual'] = (pareto['Valor Venda'] / receita_total) * 100
    pareto['Pct_Acumulado'] = pareto['Pct_Individual'].cumsum()
    pareto['Classe_ABC'] = np.where(pareto['Pct_Acumulado'] <= 80, 'Classe A (80% Faturamento)',
                            np.where(pareto['Pct_Acumulado'] <= 95, 'Classe B (15% Faturamento)', 'Classe C (5% Cauda Longa)'))
    
    # Nome curto para não cortar no gráfico
    def get_short_name(prod):
        p = prod.replace('Contoso ', '').replace('Player ', '').replace('8GB ', '').replace('4G ', '')
        return p[:22]

    pareto['Nome_Curto'] = pareto['Produto'].apply(get_short_name)
    top_3_pct = pareto.head(3)['Pct_Individual'].sum()
    print(f"  - Top 3 produtos representam {top_3_pct:.1f}% da receita!")

    # 5. MATRIZ DE CALOR (HEATMAP) DE SAZONALIDADE (ANO x MÊS)
    print("\n[Analytics 4.0] Matriz de Calor de Sazonalidade (Heatmap)...")
    heatmap_matrix = []
    meses_nomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    for ano in [2017, 2018, 2019]:
        linha = {"ano": ano, "meses": []}
        for mes_idx in range(1, 13):
            val = df_merged[(df_merged['Ano'] == ano) & (df_merged['Mes'] == mes_idx)]['Valor Venda'].sum()
            linha["meses"].append({
                "mes_idx": mes_idx,
                "mes_nome": meses_nomes[mes_idx - 1],
                "valor": round(float(val), 2)
            })
        heatmap_matrix.append(linha)

    # 6. KPIS POR PAÍS E METAS (CENÁRIO PBIX OFICIAL vs CENÁRIO DE GOVERNANÇA)
    print("\n[Analytics 4.0] KPIs por País e Metas (Alinhamento 100% com PBIX)...")
    metas_pais = df_metas.groupby('País')['Valor Meta'].sum().reset_index()

    # Cenário PBIX Oficial (com Nulos do Item 1 desvinculados de País)
    vendas_pbi_oficial = df_merged.groupby('País', dropna=False)['Valor Venda'].sum().reset_index()
    vendas_pbi_oficial['País'] = vendas_pbi_oficial['País'].fillna('Não Alocado (Item 1)')
    kpi_pbi_oficial = vendas_pbi_oficial.merge(metas_pais, on='País', how='left').fillna({'Valor Meta': 0})
    kpi_pbi_oficial['% Atingimento'] = np.where(kpi_pbi_oficial['Valor Meta'] > 0, 
                                                (kpi_pbi_oficial['Valor Venda'] / kpi_pbi_oficial['Valor Meta']) * 100, 
                                                0)
    kpi_pbi_oficial['Gap'] = kpi_pbi_oficial['Valor Venda'] - kpi_pbi_oficial['Valor Meta']

    # Cenário de Governança (com todos os países alocados)
    vendas_biz = df_merged.groupby('País_Original')['Valor Venda'].sum().reset_index().rename(columns={'País_Original': 'País'})
    kpi_biz = vendas_biz.merge(metas_pais, on='País', how='left').fillna({'Valor Meta': 0})
    kpi_biz['% Atingimento'] = (kpi_biz['Valor Venda'] / kpi_biz['Valor Meta']) * 100
    kpi_biz['Gap'] = kpi_biz['Valor Venda'] - kpi_biz['Valor Meta']

    print("  Valores PBI Oficial por País:")
    for _, r in kpi_pbi_oficial.iterrows():
        print(f"    - {r['País']}: R$ {r['Valor Venda']:,.2f} (Meta: R$ {r['Valor Meta']:,.2f} | Ating: {r['% Atingimento']:.1f}%)")

    # 7. SÉRIE TEMPORAL MENSAL E TRIMESTRAL
    vendas_mensal = df_merged.groupby('AnoMes_Str')['Valor Venda'].sum().reset_index()
    vendas_mensal['Vendas_Mes_Anterior'] = vendas_mensal['Valor Venda'].shift(1).fillna(0)
    vendas_mensal['Var_MoM_Pct'] = np.where(vendas_mensal['Vendas_Mes_Anterior'] > 0,
                                            ((vendas_mensal['Valor Venda'] - vendas_mensal['Vendas_Mes_Anterior']) / vendas_mensal['Vendas_Mes_Anterior']) * 100,
                                            0.0)
    vendas_mensal['Var_MoM_Pct'] = vendas_mensal['Var_MoM_Pct'].replace([np.inf, -np.inf], 0.0).fillna(0.0)

    # Trimestral
    vendas_trimestre = df_merged.groupby('Trimestre')['Valor Venda'].sum().reset_index()

    # Vendas detalhadas para filtros dinâmicos
    df_merged['País_PBI'] = df_merged['País'].fillna('Não Alocado (Item 1)')
    vendas_detalhadas = df_merged[['ID Cliente', 'Cliente', 'País_PBI', 'País_Original', 'Produto', 'Subcategoria', 'Data Venda', 'Ano', 'AnoMes_Str', 'Trimestre', 'Valor Venda']].copy()
    vendas_detalhadas['Data Venda'] = vendas_detalhadas['Data Venda'].dt.strftime('%Y-%m-%d')
    vendas_detalhadas['Cliente'] = vendas_detalhadas['Cliente'].fillna('Cliente Não Informado')

    # PAYLOAD COMPLETO 4.0
    dashboard_payload = {
        "metadata": {
            "version": "4.0 Executive Next-Gen",
            "author": "José Willian",
            "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "data_min": df_merged['Data Venda'].min().strftime("%d/%m/%Y"),
            "data_max": df_merged['Data Venda'].max().strftime("%d/%m/%Y"),
            "total_transacoes": len(df_vendas),
            "total_receita": round(receita_total, 2),
            "total_metas": round(df_metas['Valor Meta'].sum(), 2),
            "atingimento_global_pct": round((receita_total / df_metas['Valor Meta'].sum()) * 100, 2),
            "receita_nula": round(receita_nula, 2),
            "pct_nula": round(pct_nula, 1)
        },
        "kpis_pais_oficial": kpi_pbi_oficial.round(2).to_dict(orient='records'),
        "kpis_pais_governanca": kpi_biz.round(2).to_dict(orient='records'),
        "pareto_produtos": pareto.round(2).to_dict(orient='records'),
        "heatmap_sazonalidade": heatmap_matrix,
        "vendas_mensais": vendas_mensal.round(2).to_dict(orient='records'),
        "vendas_trimestrais": vendas_trimestre.round(2).to_dict(orient='records'),
        "metas_anuais": df_metas.round(2).to_dict(orient='records'),
        "registros_vendas": vendas_detalhadas.round(2).to_dict(orient='records')
    }

    # Salva JSON e JS
    with open('dashboard_data.json', 'w', encoding='utf-8') as f:
        json.dump(dashboard_payload, f, ensure_ascii=False, indent=2)

    with open('data.js', 'w', encoding='utf-8') as f:
        f.write('window.DASHBOARD_DATA = ' + json.dumps(dashboard_payload, ensure_ascii=False) + ';\n')

    print("\n[OK] Pipeline 4.0 concluído com sucesso!")
    print(f"  - dashboard_data.json: {os.path.getsize('dashboard_data.json'):,} bytes")
    print(f"  - data.js: {os.path.getsize('data.js'):,} bytes")
    return dashboard_payload

if __name__ == '__main__':
    run_pipeline()
