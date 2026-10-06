from datetime import datetime
import pandas as pd
import streamlit as st

# Configuração da página com tema corporativo
st.set_page_config(
    page_title="PepsiCo - Consulta e Gestão de NFs",
    page_icon="🔵",
    layout="wide",
)

# Estilização CSS personalizada
st.markdown(
    """
    <style>
        .stApp {
            background-color: #0d1117;
            color: #c9d1d9;
        }
        .main-header {
            font-size: 38px;
            font-weight: 800;
            color: #ffffff;
            margin-bottom: 0px;
            line-height: 1.2;
        }
        .sub-header {
            font-size: 20px;
            color: #8b949e;
            margin-top: 5px;
        }
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #161b22;
            color: #8b949e;
            text-align: center;
            padding: 8px;
            font-size: 12px;
            border-top: 1px solid #30363d;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# Cabeçalho com Logotipo e títulos
col_logo, col_titulo = st.columns([1.8, 5.5])

with col_logo:
  try:
    st.image("logo.png", width=280)
  except:
    st.write("🔵")

with col_titulo:
  st.markdown("<br>", unsafe_allow_html=True)
  st.markdown(
      '<p class="main-header">PepsiCo - Consulta e Gestão de NFs</p>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<p class="sub-header">Sistema Interno de Acompanhamento de Slips e Notas'
      " Fiscais</p>",
      unsafe_allow_html=True,
  )

st.markdown("---")

ARQUIVO_EXCEL = "base_slips_nfs.xlsx"


# Função para carregar os dados de todas as abas
@st.cache_data(ttl=2)
def carregar_dados():
  try:
    return pd.read_excel(ARQUIVO_EXCEL, sheet_name=None)
  except FileNotFoundError:
    return None


abas = carregar_dados()

if abas is None:
  st.error(
      f"❌ O ficheiro '{ARQUIVO_EXCEL}' não foi encontrado na mesma pasta do"
      " script!"
  )
else:
  lista_abas = list(abas.keys())

  # Menu lateral para alternar entre Consultar e Atualizar
  menu = st.sidebar.selectbox(
      "Menu de Navegação", ["🔍 Consultar Status", "✏️ Atualizar Status"]
  )

  if menu == "🔍 Consultar Status":
    st.subheader("Consultar Nota Fiscal")
    termo_busca = st.text_input(
        "Digite a Chave de Acesso ou o Número da NF:",
        placeholder="Ex: 10457 ou 352610...",
    )

    if st.button("Buscar", type="primary"):
      if not termo_busca:
        st.warning("Por favor, digite um valor para pesquisar.")
      else:
        encontrado = False
        resultado_df = None
        aba_encontrada = ""

        # Procurar em todas as abas e linhas
        for nome_aba, df in abas.items():
          df_str = df.astype(str)
          for col in df_str.columns:
            matches = df_str[
                df_str[col].str.contains(str(termo_busca), na=False)
            ]
            if not matches.empty:
              encontrado = True
              resultado_df = matches.iloc[0]
              aba_encontrada = nome_aba
              break
          if encontrado:
            break

        if encontrado:
          st.success(f"Encontrado na aba: **{aba_encontrada.upper()}**")

          # Mapeamento flexível para encontrar colunas independentemente de maiúsculas/acentos
          colunas_disponiveis = {
              c.lower().strip(): c for c in resultado_df.index
          }

          def pegar_valor(possiveis_nomes):
            for nome in possiveis_nomes:
              if nome.lower() in colunas_disponiveis:
                val = resultado_df[colunas_disponiveis[nome.lower()]]
                return "N/D" if pd.isna(val) else val
            return "N/D"

          cnpj = pegar_valor(["CNPJ"])
          num_nf = pegar_valor(["Numero_NF", "Número da NF", "NF", "Nota"])
          valor = pegar_valor(["Valor da NF", "Valor", "R$"])
          chave = pegar_valor(["Chave de acesso", "Chave"])
          emissao = pegar_valor(["Data da emissão", "Emissão"])
          atualizacao = pegar_valor(["Data da atualização", "Última Atualiz."])
          status_atual = str(
              pegar_valor(["Status", "Situação"])
          ).strip().lower()

          c1, c2 = st.columns(2)
          with c1:
            st.text(f"CNPJ: {cnpj}")
            st.text(f"Número da NF: {num_nf}")
            st.text(f"Valor: R$ {valor}")
          with c2:
            st.text(f"Chave:\n{chave}")
            st.text(f"Emissão: {emissao}")
            st.text(f"Última Atualiz.: {atualizacao}")

          st.markdown("### Status Atual:")
          if "conclu" in status_atual:
            st.markdown("🟢 **Concluído**")
          elif "process" in status_atual or "análise" in status_atual:
            st.markdown("🟡 **Em processamento / Em Análise**")
          else:
            st.markdown(f"🔴 **{status_atual.capitalize()}**")
        else:
          st.error(
              "Nenhuma nota fiscal ou chave encontrada com o termo informado."
          )

  elif menu == "✏️ Atualizar Status":
    st.subheader("Atualizar Status de Nota Fiscal")
    aba_selecionada = st.selectbox("Selecione a Aba:", lista_abas)
    df_aba = abas[aba_selecionada]

    # Identifica automaticamente a coluna de número da NF na aba selecionada
    col_nf_candidatas = [
        c
        for c in df_aba.columns
        if "nf" in c.lower() or "nota" in c.lower() or "numero" in c.lower()
    ]

    if col_nf_candidatas:
      col_nf = col_nf_candidatas[0]
      nf_selecionada = st.selectbox(
          "Selecione o Número da NF:", df_aba[col_nf].astype(str)
      )
      novo_status = st.selectbox(
          "Novo Status:",
          ["Pendente", "Em processamento", "Concluído", "Rejeitado"],
      )

      if st.button("Salvar Atualização", type="primary"):
        idx = df_aba[df_aba[col_nf].astype(str) == nf_selecionada].index
        if not idx.empty:
          # Identifica a coluna de status e de data de atualização
          col_status_cand = [c for c in df_aba.columns if "status" in c.lower()]
          col_data_cand = [
              c for c in df_aba.columns if "atualiza" in c.lower()
          ]

          if col_status_cand:
            df_aba.loc[idx, col_status_cand[0]] = novo_status
          if col_data_cand:
            df_aba.loc[idx, col_data_cand[0]] = datetime.now().strftime(
                "%Y-%m-%d"
            )

          with pd.ExcelWriter(ARQUIVO_EXCEL, engine="openpyxl") as writer:
            for nome, df_original in abas.items():
              if nome == aba_selecionada:
                df_aba.to_excel(writer, sheet_name=nome, index=False)
              else:
                df_original.to_excel(writer, sheet_name=nome, index=False)

          st.success(
              "Status atualizado com sucesso e data de alteração registada!"
          )
          st.rerun()
    else:
      st.error(
          "Não foi encontrada nenhuma coluna de número de NF na aba selecionada."
      )

st.markdown(
    '<div class="footer">© PepsiCo Brasil - Internal Tool</div>',
    unsafe_allow_html=True,
)