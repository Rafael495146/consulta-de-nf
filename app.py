from datetime import datetime
import pandas as pd
import streamlit as st

# Configuração da página com tema corporativo
st.set_page_config(
    page_title="PepsiCo - Consulta de NFs",
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
  # Tela de Consulta Direta (Sem menus laterais de edição)
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
          matches = df_str[df_str[col].str.contains(str(termo_busca), na=False)]
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
        colunas_disponiveis = {c.lower().strip(): c for c in resultado_df.index}

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

st.markdown(
    '<div class="footer">© PepsiCo Brasil - Internal Tool</div>',
    unsafe_allow_html=True,
)