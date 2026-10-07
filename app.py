from datetime import datetime
import matplotlib.pyplot as plt
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
  # TELA 2: MENU PRINCIPAL
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
        " Gráficos Estáticos</h3>",
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
      if st.button(
          "📊 Gráfico Fixo: Quantidade de NFs por Categoria",
          use_container_width=True,
      ):
        st.session_state.nav_mode = "Grafico_Qtd"
        st.rerun()
    with g2:
      if st.button(
          "💰 Gráfico Fixo: Valor Financeiro (P. Bom & P. Ruim)",
          use_container_width=True,
      ):
        st.session_state.nav_mode = "Grafico_Valor"
        st.rerun()

  # ==========================================
  # TELA 3: GRÁFICO ESTÁTICO DE QUANTIDADE + FILTRO
  # ==========================================
  elif st.session_state.nav_mode == "Grafico_Qtd":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        "📊 Quantidade de Notas Fiscais por Categoria (Gráfico Fixo)"
    )

    # Filtro de Período (Mês/Ano) recolhido de todas as abas
    todas_linhas = []
    for nome_aba, df in abas.items():
      df_temp = df.copy()
      df_temp["Categoria_Aba"] = nome_aba
      todas_linhas.append(df_temp)

    df_geral = pd.concat(todas_linhas, ignore_index=True)

    # Extrai Ano-Mês da coluna de emissão se existir
    col_emissao = None
    for c in df_geral.columns:
      if "emiss" in c.lower():
        col_emissao = c
        break

    if col_emissao:
      df_geral["AnoMes"] = (
          pd.to_datetime(df_geral[col_emissao], errors="coerce")
          .dt.strftime("%Y-%m")
          .fillna("Geral")
      )
      meses_disponiveis = ["Todos"] + sorted(
          [m for m in df_geral["AnoMes"].unique() if m != "Geral"]
      )
      mes_selecionado = st.selectbox(
          "📅 Filtrar por Mês/Ano (Emissão):", meses_disponiveis
      )

      if mes_selecionado != "Todos":
        df_geral = df_geral[df_geral["AnoMes"] == mes_selecionado]

    # Calcula quantidades filtradas
    contagem_dados = {}
    categorias_alvo = ["Produto Bom", "Produto Ruim", "Recusa", "Reentrega"]
    for cat in categorias_alvo:
      qtd = len(
          df_geral[
              df_geral["Categoria_Aba"].str.lower().str.contains(cat.lower())
          ]
      )
      contagem_dados[cat] = qtd

    # Desenha o gráfico estático com Matplotlib (não mexe ao passar o rato)
    fig, ax = plt.subplots(figsize=(8, 4.5))
    fig.patch.set_facecolor("#131b2e")
    ax.set_facecolor("#131b2e")

    categorias = list(contagem_dados.keys())
    quantidades = list(contagem_dados.values())
    carr_cores = ["#238636", "#da3633", "#9e6a03", "#1f6feb"]

    bars = ax.bar(
        categorias, quantidades, color=carr_cores, width=0.5, edgecolor="none"
    )

    ax.tick_params(colors="#c9d1d9", labelsize=11)
    ax.spines["bottom"].set_color("#30363d")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#30363d")
    ax.grid(axis="y", linestyle="--", alpha=0.2, color="#8b949e")

    for bar in bars:
      height = bar.get_height()
      ax.annotate(
          f"{height}",
          xy=(bar.get_x() + bar.get_width() / 2, height),
          xytext=(0, 3),
          textcoords="offset points",
          ha="center",
          va="bottom",
          color="white",
          fontweight="bold",
      )

    st.pyplot(fig)

  # ==========================================
  # TELA 4: GRÁFICO ESTÁTICO DE VALOR + FILTRO
  # ==========================================
  elif st.session_state.nav_mode == "Grafico_Valor":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        "💰 Montante Financeiro - Produto Bom & Produto Ruim (Gráfico Fixo)"
    )

    todas_linhas_val = []
    for nome_aba, df in abas.items():
      if "bom" in nome_aba.lower() or "ruim" in nome_aba.lower():
        df_temp = df.copy()
        df_temp["Categoria_Aba"] = nome_aba
        todas_linhas_val.append(df_temp)

    if todas_linhas_val:
      df_val_geral = pd.concat(todas_linhas_val, ignore_index=True)

      col_emissao = None
      for c in df_val_geral.columns:
        if "emiss" in c.lower():
          col_emissao = c
          break

      if col_emissao:
        df_val_geral["AnoMes"] = (
            pd.to_datetime(df_val_geral[col_emissao], errors="coerce")
            .dt.strftime("%Y-%m")
            .fillna("Geral")
        )
        meses_disponiveis = ["Todos"] + sorted(
            [m for m in df_val_geral["AnoMes"].unique() if m != "Geral"]
        )
        mes_selecionado = st.selectbox(
            "📅 Filtrar por Mês/Ano (Emissão):", meses_disponiveis
        )

        if mes_selecionado != "Todos":
          df_val_geral = df_val_geral[df_val_geral["AnoMes"] == mes_selecionado]

      valores_dados = {}
      for cat in ["Produto Bom", "Produto Ruim"]:
        df_subset = df_val_geral[
            df_val_geral["Categoria_Aba"].str.lower().str.contains(cat.lower())
        ]
        col_val = None
        for c in df_subset.columns:
          if "valor" in c.lower() or "r$" in c.lower():
            col_val = c
            break
        if col_val:
          soma_val = pd.to_numeric(df_subset[col_val], errors="coerce").sum()
          valores_dados[cat] = round(soma_val, 2)
        else:
          valores_dados[cat] = 0.0

      # Desenha gráfico estático de valores
      fig, ax = plt.subplots(figsize=(7, 4.5))
      fig.patch.set_facecolor("#131b2e")
      ax.set_facecolor("#131b2e")

      categorias = list(valores_dados.keys())
      valores = list(valores_dados.values())
      carr_cores = ["#238636", "#da3633"]

      bars = ax.bar(
          categorias,
          valores,
          color=carr_cores,
          width=0.4,
          edgecolor="none",
      )

      ax.tick_params(colors="#c9d1d9", labelsize=11)
      ax.spines["bottom"].set_color("#30363d")
      ax.spines["top"].set_visible(False)
      ax.spines["right"].set_visible(False)
      ax.spines["left"].set_color("#30363d")
      ax.grid(axis="y", linestyle="--", alpha=0.2, color="#8b949e")

      for bar in bars:
        height = bar.get_height()
        ax.annotate(
            f"R$ {height:,.2f}",
            xy=(bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 3),
            textcoords="offset points",
            ha="center",
            va="bottom",
            color="white",
            fontweight="bold",
        )

      st.pyplot(fig)
    else:
      st.warning("Não há dados financeiros suficientes para exibir o gráfico.")

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