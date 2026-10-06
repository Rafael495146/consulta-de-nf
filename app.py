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
  # Estado da navegação guardado na sessão do Streamlit
  if "nav_mode" not in st.session_state:
    st.session_state.nav_mode = "Home"

  # ==========================================
  # TELA INICIAL: MENU DE ESCOLHA DE CATEGORIA
  # ==========================================
  if st.session_state.nav_mode == "Home":
    st.markdown(
        "<h2 style='text-align: center; color: white;'>Seja Bem-Vindo! 👋</h2>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align: center; color: #8b949e; font-size: 18px;'>Por"
        " favor, selecione a categoria que deseja consultar:</p>",
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    col_vazia1, b1, b2, col_vazia2 = st.columns([1.5, 3, 3, 1.5])
    with b1:
      if st.button("🟢 Produto Bom", use_container_width=True, type="primary"):
        st.session_state.nav_mode = "Produto Bom"
        st.rerun()
    with b2:
      if st.button("🟡 Produto Ruim", use_container_width=True, type="primary"):
        st.session_state.nav_mode = "Produto Ruim"
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    col_vazia3, b3, b4, col_vazia4 = st.columns([1.5, 3, 3, 1.5])
    with b3:
      if st.button("🔴 Recusa", use_container_width=True, type="primary"):
        st.session_state.nav_mode = "Recusa"
        st.rerun()
    with b4:
      if st.button("🔵 Reentrega", use_container_width=True, type="primary"):
        st.session_state.nav_mode = "Reentrega"
        st.rerun()

  # ==========================================
  # TELA DE CONSULTA ESPECÍFICA DA CATEGORIA
  # ==========================================
  else:
    categoria_ativa = st.session_state.nav_mode

    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(f"🔍 Consultar em: {categoria_ativa}")

    # Encontra a aba correspondente no Excel
    aba_correspondente = None
    for nome_aba in abas.keys():
      if categoria_ativa.lower() in nome_aba.lower():
        aba_correspondente = nome_aba
        break

    if aba_correspondente:
      df_cat = abas[aba_correspondente]

      termo_busca = st.text_input(
          f"Digite o Número da NF ou Chave de Acesso ({categoria_ativa}):",
          placeholder="Ex: 10457 ou 352610...",
      )

      if st.button("Pesquisar", type="primary"):
        if not termo_busca:
          st.warning("Por favor, digite um valor para pesquisar.")
        else:
          df_str = df_cat.astype(str)
          encontrado = False
          resultado_row = None

          for col in df_str.columns:
            matches = df_str[
                df_str[col].str.contains(str(termo_busca), na=False)
            ]
            if not matches.empty:
              encontrado = True
              resultado_row = matches.iloc[0]
              break

          if encontrado:
            st.success(f"✨ Registo encontrado na categoria **{categoria_ativa}**!")

            colunas_disponiveis = {
                c.lower().strip(): c for c in resultado_row.index
            }

            def pegar_valor(possiveis_nomes):
              for nome in possiveis_nomes:
                if nome.lower() in colunas_disponiveis:
                  val = resultado_row[colunas_disponiveis[nome.lower()]]
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
              st.info(f"**CNPJ:** {cnpj}")
              st.info(f"**Número da NF:** {num_nf}")
              st.info(f"**Valor:** R$ {valor}")
            with c2:
              st.info(f"**Chave de Acesso:**\n{chave}")
              st.info(f"**Emissão:** {emissao}")
              st.info(f"**Última Atualiz.:** {atualizacao}")

            st.markdown("### Status Atual:")
            if "conclu" in status_atual:
              st.markdown("🟢 **Concluído**")
            elif "process" in status_atual or "análise" in status_atual:
              st.markdown("🟡 **Em processamento / Em Análise**")
            else:
              st.markdown(f"🔴 **{status_atual.capitalize()}**")
          else:
            st.error(
                f"Nenhum registo correspondente encontrado em '{categoria_ativa}'"
                " com o termo informado."
            )
    else:
      st.warning(
          f"Ainda não existe uma aba no Excel com o nome equivalente a"
          f" '{categoria_ativa}'. Certifique-se de que o ficheiro Excel"
          f" possui essa aba configurada!"
      )

st.markdown(
    '<div class="footer">© PepsiCo Brasil - Internal Tool | Sistema'
    " Operacional Móvel</div>",
    unsafe_allow_html=True,
)