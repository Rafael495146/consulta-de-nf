import pandas as pd

status_opcoes = [
    "Operação Concluída",
    "Em Processamento na Logística",
    "Pendente de Aprovação",
    "Não Entregue / Ocorrência",
]
clientes = [
    "Supermercado Alvorada Ltda",
    "Comercial São José",
    "Distribuidora Central",
    "Mercado Bom Preço",
    "Atacado Boa Vista",
]
cnpjs = [
    "12.345.678/0001-90",
    "98.765.432/0001-12",
    "45.678.901/0001-34",
    "11.222.333/0001-44",
    "55.444.333/0001-22",
]

# Gera filiais como CDV01, CDV02 até CDV30
filiais = [f"CDV{i:02d}" for i in range(1, 31)]


def gerar_dados_aba(inicio_nf):
  dados = []
  for i in range(150):
    num_nf = str(inicio_nf + (i % 50))
    chave = f"352610{str(inicio_nf + i).zfill(12)}50010000{str(i+1).zfill(6)}"
    cnpj = cnpjs[i % len(cnpjs)]
    cliente = clientes[i % len(clientes)]
    filial = filiais[i % len(filiais)]
    valor = round(150.0 + (i * 18.5) % 4500, 2)
    status = status_opcoes[i % len(status_opcoes)]
    emissao = f"2026-09-{(i % 28) + 1:02d}"
    atualizacao = f"2026-10-{(i % 6) + 1:02d}"

    dados.append({
        "Filial": filial,
        "Nome_Cliente": cliente,
        "CNPJ": cnpj,
        "Numero_NF": num_nf,
        "Chave de acesso": chave,
        "Valor da NF": valor,
        "Status": status,
        "Data da emissão": emissao,
        "Data da atualização": atualizacao,
    })
  return pd.DataFrame(dados)


df_bom = gerar_dados_aba(10001)
df_ruim = gerar_dados_aba(20001)
df_recusa = gerar_dados_aba(30001)
df_reentrega = gerar_dados_aba(40001)

with pd.ExcelWriter("base_slips_nfs.xlsx", engine="openpyxl") as writer:
  df_bom.to_excel(writer, sheet_name="Produto Bom", index=False)
  df_ruim.to_excel(writer, sheet_name="Produto Ruim", index=False)
  df_recusa.to_excel(writer, sheet_name="Recusa", index=False)
  df_reentrega.to_excel(writer, sheet_name="Reentrega", index=False)

print("Base Excel multi-CD gerada com sucesso!")