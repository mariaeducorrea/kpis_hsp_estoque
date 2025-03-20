from conexao import connection
import csv

def extracao():
    conn = connection()
    cur = conn.cursor()

    query = """"
    select 
        se.id_saida_estoque, 
        se.cod_estoque, 
        e.descr_estoque,
        se.cod_centro_custo, 
        cc.descr_cc,
        se.cod_produto, 
        p.nm_produto,
        se.qtd, 
        se.data_saida
    from sigh.saidas_estoques se
    inner join sigh.estoques e 
        on se.cod_estoque = e.id_estoque 
    inner join sigh.centros_custos cc 
        on se.cod_centro_custo = cc.id_centro_custo 
    inner join sigh.produtos p 
        on se.cod_produto = p.id_produto 
    where se.data_saida between '2023-01-01' and '2023-12-01'
    order by data_saida asc;
    """
    
    cur.execute(query)

    rows = cur.fetchall()

    csv_file = "C:/Users/Administrador/Desktop/projetos/kpis_hsp_estoque_powerbi/dados"

    with open(csv_file, mode='w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['id_saida_estoque','cod_estoque','descr_estoque','cod_centro_custo','descr_cc','cod_produto','nm_produto','qtd','data_saida'])
        writer.writerows(rows)

    cur.close()
    conn.close()

extracao()
    