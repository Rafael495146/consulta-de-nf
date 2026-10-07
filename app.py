from datetime import datetime
import pandas as pd
import streamlit as st

# Configuração da página com tema corporativo escuro
st.set_page_config(
    page_title="PepsiCo - Painel Operacional",
    page_icon="🔵",
    layout="wide",
)

# Estilização CSS avançada estilo Dashdark X
st.markdown(
    """
    <style>
        .stApp {
            background-color: #0b0f19;
            color: #c9d1d9;
        }
        .main-header {
            font-size: 34px;
            font-weight: 800;
            color: #ffffff;
            margin-bottom: 0px;
            line-height: 1.2;
        }
        .sub-header {
            font-size: 16px;
            color: #8b949e;
            margin-top: 5px;
        }
        .welcome-card {
            background-color: #131b2e;
            border: 1px solid #1f293d;
            padding: 40px;
            border-radius: 16px;
            text-align: center;
            box-shadow: 0 10px 25px rgba(0,0,0,0.5);
            max-width: 650px;
            margin: 40px auto;
        }
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #0b0f19;
            color: #8b949e;
            text-align: center;
            padding: 8px;
            font-size: 12px;
            border-top: 1px solid #1f293d;
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
      '<p class="sub-header">Acompanhamento Operacional e Analítico em Tempo'
      " Real</p>",
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
    st.session_state.nav_mode = "Welcome"

  # ==========================================
  # TELA 1: SEJA BEM-VINDO
  # ==========================================
  if st.session_state.nav_mode == "Welcome":
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="welcome-card">
            <h1 style="color: white; font-size: 32px; margin-bottom: 10px;">Seja Bem-Vindo! 👋</h1>
            <p style="color: #8b949e; font-size: 16px; margin-bottom: 30px;">
                Sistema interno de consulta, gestão e análise operacional de Notas Fiscais e Slips da equipa PepsiCo.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col_btn, col2 = st.columns([2, 2, 2])
    with col_btn:
      if st.button(
          "🚀 Avançar para o Sistema", use_container_width=True, type="primary"
      ):
        st.session_state.nav_mode = "Home"
        st.rerun()

  # ==========================================
  # TELA 2: MENU PRINCIPAL (Categorias + Gráficos)
  # ==========================================
  elif st.session_state.nav_mode == "Home":
    if st.button("⬅️ Voltar à Tela Inicial"):
      st.session_state.nav_mode = "Welcome"
      st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        "<h3 style='text-align: center; color: white;'>Consultas por"
        " Categoria</h3>",
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    # Botões das 4 categorias
    c1, c2, c3, c4 = st.columns(4)
    with c1:
      if st.button("🟢 Produto Bom", use_container_width=True):
        st.session_state.nav_mode = "Produto Bom"
        st.rerun()
    with c2:
      if st.button("🔴 Produto Ruim", use_container_width=True):
        st.session_state.nav_mode = "Produto Ruim"
        st.rerun()
    with c3:
      if st.button("🟡 Recusa", use_container_width=True):
        st.session_state.nav_mode = "Recusa"
        st.rerun()
    with c4:
      if st.button("🔵 Reentrega", use_container_width=True):
        st.session_state.nav_mode = "Reentrega"
        st.rerun()

    st.markdown("<br><hr style='border-color: #1f293d;'><br>", unsafe_allow_html=True)
    st.markdown(
        "<h3 style='text-align: center; color: white;'>Painéis Analíticos e"
        " Gráficos</h3>",
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    # Botões dos Gráficos
    g1, g2 = st.columns(2)
    with g1:
      if st.button(
          "📊 Gráfico: Quantidade de NFs por Categoria", use_container_width=True
      ):
        st.session_state.nav_mode = "Grafico_Qtd"
        st.rerun()
    with g2:
      if st.button(
          "💰 Gráfico: Valor Financeiro (P. Bom & P. Ruim)",
          use_container_width=True,
      ):
        st.session_state.nav_mode = "Grafico_Valor"
        st.rerun()

  # ==========================================
  # TELA 3: GRÁFICO DE QUANTIDADE DE NFs
  # ==========================================
  elif st.session_state.nav_mode == "Grafico_Qtd":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        "📊 Quantidade de Notas Fiscais Recebidas por Categoria"
    )
    st.markdown(
        "<p style='color: #8b949e;'>Volume total de registos processados em cada"
        " aba da base de dados.</p>",
        unsafe_allow_html=True,
    )

    contagem_dados = {}
    categorias_alvo = ["Produto Bom", "Produto Ruim", "Recusa", "Reentrega"]
    for cat in categorias_alvo:
      for nome_aba, df in abas.items():
        if cat.lower() in nome_aba.lower():
          contagem_dados[cat] = len(df)

    if contagem_dados:
      df_grafico = pd.DataFrame(
          list(contagem_dados.items()), columns=["Categoria", "Quantidade"]
      ).set_index("Categoria")
      st.bar_chart(df_grafico, use_container_width=True)
    else:
      st.warning("Não foram encontradas abas correspondentes para o gráfico.")

  # ==========================================
  # TELA 4: GRÁFICO DE VALOR (P. Bom & P. Ruim)
  # ==========================================
  elif st.session_state.nav_mode == "Grafico_Valor":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader("💰 Montante Financeiro - Produto Bom & Produto Ruim")
    st.markdown(
        "<p style='color: #8b949e;'>Soma total dos valores financeiros"
        " associados às categorias de Produto Bom e Produto Ruim.</p>",
        unsafe_allow_html=True,
    )

    valores_dados = {}
    for cat in ["Produto Bom", "Produto Ruim"]:
      for nome_aba, df in abas.items():
        if cat.lower() in nome_aba.lower():
          col_val = None
          for c in df.columns:
            if "valor" in c.lower() or "r$" in c.lower():
              col_val = c
              break
          if col_val:
            soma_val = pd.to_numeric(df[col_val], errors="coerce").sum()
            valores_dados[cat] = round(soma_val, 2)

    if valores_dados:
      df_valores = pd.DataFrame(
          list(valores_dados.items()), columns=["Categoria", "Valor Total (R$)"]
      ).set_index("Categoria")
      st.bar_chart(df_valores, use_container_width=True)
    else:
      st.warning("Não foi possível calcular os valores financeiros das abas.")

  # ==========================================
  # TELA 5: CONSULTA ESPECÍFICA DA CATEGORIA
  # ==========================================
  else:
    categoria_ativa = st.session_state.nav_mode

    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(f"🔍 Consultar em: {categoria_ativa}")

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
            elif "pendente" in status_atual:
              st.markdown("🟠 **Pendente**")
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