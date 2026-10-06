from datetime import datetime
import pandas as pd
import streamlit as st

# Configuração da página com tema corporativo
st.set_page_config(
    page_title="PepsiCo - Gestão e Consulta de NFs",
    page_icon="🔵",
    layout="wide",
)

# Estilização CSS personalizada para dar vida e cor à interface
st.markdown(
    """
    <style>
        .stApp {
            background-color: #0d1117;
            color: #c9d1d9;
        }
        .main-header {
            font-size: 36px;
            font-weight: 800;
            color: #ffffff;
            margin-bottom: 0px;
            line-height: 1.2;
        }
        .sub-header {
            font-size: 18px;
            color: #8b949e;
            margin-top: 5px;
        }
        .metric-card {
            background-color: #161b22;
            border: 1px solid #30363d;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
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
    st.image("logo.png", width=260)
  except:
    st.write("🔵")

with col_titulo:
  st.markdown("<br>", unsafe_allow_html=True)
  st.markdown(
      '<p class="main-header">PepsiCo - Painel de NFs e Slips</p>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<p class="sub-header">Acompanhamento Operacional em Tempo Real</p>',
      unsafe_allow_html=True,
  )

st.markdown("---")

ARQUIVO_EXCEL = "base_slips_nfs.xlsx"


# Função para carregar os dados
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
      " projeto!"
  )
else:
  # Menu de navegação por botões interativos no topo
  st.markdown("### 🧭 Navegação Rápida por Categoria")
  b1, b2, b3, b4, b5 = st.columns(5)

  # Estado da navegação guardado na sessão do Streamlit
  if "nav_mode" not in st.session_state:
    st.session_state.nav_mode = "🔍 Consultar Nota"

  with b1:
    if st.button("🔍 Consultar Nota", use_container_width=True):
      st.session_state.nav_mode = "🔍 Consultar Nota"
  with b2:
    if st.button("🟢 Produto Bom", use_container_width=True):
      st.session_state.nav_mode = "Produto Bom"
  with b3:
    if st.button("🟡 Produto Ruim", use_container_width=True):
      st.session_state.nav_mode = "Produto Ruim"
  with b4:
    if st.button("🔴 Recusa", use_container_width=True):
      st.session_state.nav_mode = "Recusa"
  with b5:
    if st.button("🔵 Reentrega", use_container_width=True):
      st.session_state.nav_mode = "Reentrega"

  st.markdown("---")

  # MODO 1: Consulta Direta de Chave/NF
  if st.session_state.nav_mode == "🔍 Consultar Nota":
    st.subheader("Consultar Nota Fiscal por Chave ou Número")
    termo_busca = st.text_input(
        "Digite a Chave de Acesso ou o Número da NF:",
        placeholder="Ex: 10457 ou 352610...",
    )

    if st.button("Pesquisar Sistema", type="primary"):
      if not termo_busca:
        st.warning("Por favor, digite um valor para pesquisar.")
      else:
        encontrado = False
        resultado_df = None
        aba_encontrada = ""

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
          st.success(
              f"✨ Registo localizado na aba/categoria: **{aba_encontrada.upper()}**"
          )

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

          c1, c2 = st.columns(2)
          with c1:
            st.info(f"**CNPJ:** {cnpj}")
            st.info(f"**Número da NF:** {num_nf}")
            st.info(f"**Valor:** R$ {valor}")
          with c2:
            st.info(f"**Chave de Acesso:**\n{chave}")
            st.info(f"**Emissão:** {emissao}")
            st.info(f"**Última Atualiz.:** {atualizacao}")
        else:
          st.error("Nenhum registo encontrado com o termo informado.")

  # MODOS 2 a 5: Visualização Direta das Abas (Produto Bom, Ruim, Recusa, Reentrega)
  else:
    categoria_ativa = st.session_state.nav_mode
    st.subheader(f"📋 Registos da Categoria: {categoria_ativa}")

    # Tenta encontrar a aba correspondente no Excel de forma flexível
    aba_correspondente = None
    for nome_aba in abas.keys():
      if categoria_ativa.lower() in nome_aba.lower():
        aba_correspondente = nome_aba
        break

    if aba_correspondente:
      df_cat = abas[aba_correspondente]
      st.metric(
          label="Total de Registos nesta Categoria", value=len(df_cat)
      )
      st.dataframe(df_cat, use_container_width=True)
    else:
      st.warning(
          f"Ainda não existe uma aba no Excel com o nome equivalente a"
          f" '{categoria_ativa}'. Certifique-se de que o Excel possui uma aba"
          f" com esse nome para visualizar os dados aqui automaticamente!"
      )

st.markdown(
    '<div class="footer">© PepsiCo Brasil - Internal Tool | Sistema'
    " Operacional Móvel</div>",
    unsafe_allow_html=True,
)