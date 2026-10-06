import pandas as pd

status_opcoes = ["Concluído", "Em processamento", "Pendente", "Não entregue"]
cnpjs = [
    "12.345.678/0001-90",
    "98.765.432/0001-12",
    "45.678.901/0001-34",
    "11.222.333/0001-44",
]


def gerar_dados_aba(inicio_nf):
  dados = []
  for i in range(25):
    num_nf = str(inicio_nf + i)
    chave = f"352610{str(inicio_nf + i).zfill(12)}50010000{str(i+1).zfill(6)}"
    cnpj = cnpjs[i % len(cnpjs)]
    valor = round(150.0 + (i * 35.5), 2)
    status = status_opcoes[i % len(status_opcoes)]
    emissao = f"2026-09-{(i % 28) + 1:02d}"
    atualizacao = f"2026-10-{(i % 6) + 1:02d}"

    dados.append({
        "CNPJ": cnpj,
        "Numero_NF": num_nf,
        "Chave de acesso": chave,
        "Valor da NF": valor,
        "Status": status,
        "Data da emissão": emissao,
        "Data da atualização": atualizacao,
    })
  return pd.DataFrame(dados)


# Gera as abas
df_bom = gerar_dados_aba(10001)
df_ruim = gerar_dados_aba(20001)
df_recusa = gerar_dados_aba(30001)
df_reentrega = gerar_dados_aba(40001)

# Salva o Excel oficial
with pd.ExcelWriter("base_slips_nfs.xlsx", engine="openpyxl") as writer:
  df_bom.to_excel(writer, sheet_name="Produto Bom", index=False)
  df_ruim.to_excel(writer, sheet_name="Produto Ruim", index=False)
  df_recusa.to_excel(writer, sheet_name="Recusa", index=False)
  df_reentrega.to_excel(writer, sheet_name="Reentrega", index=False)

print("Planilha gerada com sucesso!")