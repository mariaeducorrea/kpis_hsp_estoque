import pandas as pd 
import openpyxl 

dados = pd.read_csv("C:/Users/Administrador/Desktop/projetos/kpis_hsp_estoque_powerbi/dados/dados.csv")

dados.rename(columns={
    'id_saida_estoque' : 'id_saida',
    'cod_estoque' : 'codigo_estoque',
    'descr_estoque' : 'descricao_estoque',
    'cod_centro_custo' : 'codigo_centro_custo',
    'descr_cc' : 'descricao_centro_custo',
    'cod_produto' : 'codigo_produto',
    'nm_produto' : 'nomde_produto',
    'qtd' : 'quantidade',
    'data_saida' : 'data_saida'
}, inplace=True)

dados['quantidade'] = dados['quantidade'].astype(int)
dados['data_saida'] = pd.to_datetime(dados['data_saida']).dt.date

historico_saidas = dados.copy()

saida_data = historico_saidas.groupby(dados['data_saida']).size().reset_index(name='quantidade')

caminho = "C:/Users/Administrador/Desktop/projetos/kpis_hsp_estoque_powerbi/dados/dados.xlsx"

with pd.ExcelWriter(caminho) as file:
    historico_saidas.to_excel(file, sheet_name='historico saidas', index=False),
    saida_data.to_excel(file, sheet_name='Saidas por data', index=False)

