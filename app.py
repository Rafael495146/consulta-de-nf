from datetime import datetime
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="PepsiCo - Painel Operacional",
    page_icon="🔵",
    layout="wide",
)

# Estilização CSS avançada estilo Dashdark X com cartões modernos
st.markdown(
    """
    <style>
        .stApp {
            background-color: #0b0f19;
            color: #c9d1d9;
        }
        .main-header {
            font-size: 32px;
            font-weight: 800;
            color: #ffffff;
            margin-bottom: 0px;
            line-height: 1.2;
        }
        .sub-header {
            font-size: 15px;
            color: #8b949e;
            margin-top: 5px;
        }
        .welcome-card {
            background: linear-gradient(135deg, #131b2e 0%, #0d1527 100%);
            border: 1px solid #1f293d;
            padding: 40px;
            border-radius: 16px;
            text-align: center;
            box-shadow: 0 10px 30px rgba(0,0,0,0.6);
            max-width: 600px;
            margin: 40px auto;
        }
        .section-title {
            font-size: 20px;
            font-weight: 700;
            color: #ffffff;
            margin-bottom: 15px;
            letter-spacing: 0.5px;
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

col_logo, col_titulo = st.columns([1.8, 5.5])

with col_logo:
  try:
    st.image("logo.png", width=240)
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
            <h1 style="color: white; font-size: 28px; margin-bottom: 12px;">Seja Bem-Vindo! 👋</h1>
            <p style="color: #8b949e; font-size: 15px; margin-bottom: 30px; line-height: 1.5;">
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
  # TELA 2: MENU PRINCIPAL (NOVO DESIGN MODERNO)
  # ==========================================
  elif st.session_state.nav_mode == "Home":
    col_voltar, col_vazio = st.columns([1, 6])
    with col_voltar:
      if st.button("⬅️ Início"):
        st.session_state.nav_mode = "Welcome"
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # Bloco 1: Consultas por Categoria
    st.markdown(
        '<p class="section-title">📂 Consultas e Gestão por Categoria</p>',
        unsafe_allow_html=True,
    )
    c1, c2, c3, c4 = st.columns(4)

    with c1:
      st.markdown(
          "<div style='background-color: #131b2e; border: 1px solid #1f293d;"
          " padding: 15px; border-radius: 12px; text-align: center;'>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<p style='color: #238636; font-weight: bold; margin-bottom: 8px;'>🟢"
          " Produto Bom</p>",
          unsafe_allow_html=True,
      )
      if st.button("Consultar Bom", use_container_width=True, key="btn_bom"):
        st.session_state.nav_mode = "Produto Bom"
        st.rerun()
      st.markdown("</div>", unsafe_allow_html=True)

    with c2:
      st.markdown(
          "<div style='background-color: #131b2e; border: 1px solid #1f293d;"
          " padding: 15px; border-radius: 12px; text-align: center;'>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<p style='color: #da3633; font-weight: bold; margin-bottom: 8px;'>🔴"
          " Produto Ruim</p>",
          unsafe_allow_html=True,
      )
      if st.button(
          "Consultar Ruim", use_container_width=True, key="btn_ruim"
      ):
        st.session_state.nav_mode = "Produto Ruim"
        st.rerun()
      st.markdown("</div>", unsafe_allow_html=True)

    with c3:
      st.markdown(
          "<div style='background-color: #131b2e; border: 1px solid #1f293d;"
          " padding: 15px; border-radius: 12px; text-align: center;'>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<p style='color: #d29922; font-weight: bold; margin-bottom: 8px;'>🟡"
          " Recusa</p>",
          unsafe_allow_html=True,
      )
      if st.button(
          "Consultar Recusa", use_container_width=True, key="btn_recusa"
      ):
        st.session_state.nav_mode = "Recusa"
        st.rerun()
      st.markdown("</div>", unsafe_allow_html=True)

    with c4:
      st.markdown(
          "<div style='background-color: #131b2e; border: 1px solid #1f293d;"
          " padding: 15px; border-radius: 12px; text-align: center;'>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<p style='color: #58a6ff; font-weight: bold; margin-bottom: 8px;'>🔵"
          " Reentrega</p>",
          unsafe_allow_html=True,
      )
      if st.button(
          "Consultar Reentrega", use_container_width=True, key="btn_reentrega"
      ):
        st.session_state.nav_mode = "Reentrega"
        st.rerun()
      st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)

    # Bloco 2: Painéis Analíticos
    st.markdown(
        '<p class="section-title">📊 Painéis Analíticos e Gráficos Estáticos</p>',
        unsafe_allow_html=True,
    )
    g1, g2 = st.columns(2)

    with g1:
      st.markdown(
          "<div style='background-color: #131b2e; border: 1px solid #1f293d;"
          " padding: 20px; border-radius: 12px;'>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<p style='color: #ffffff; font-weight: 600;'>Volume de Notas"
          " Fiscais</p>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<p style='color: #8b949e; font-size: 13px;'>Gráfico estatístico com"
          " a quantidade de registos por categoria e filtro temporal.</p>",
          unsafe_allow_html=True,
      )
      if st.button("📈 Ver Gráfico de Quantidade", use_container_width=True):
        st.session_state.nav_mode = "Grafico_Qtd"
        st.rerun()
      st.markdown("</div>", unsafe_allow_html=True)

    with g2:
      st.markdown(
          "<div style='background-color: #131b2e; border: 1px solid #1f293d;"
          " padding: 20px; border-radius: 12px;'>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<p style='color: #ffffff; font-weight: 600;'>Montante Financeiro"
          " (R$)</p>",
          unsafe_allow_html=True,
      )
      st.markdown(
          "<p style='color: #8b949e; font-size: 13px;'>Análise de valores"
          " envolvidos em Produto Bom e Produto Ruim com filtros.</p>",
          unsafe_allow_html=True,
      )
      if st.button("💰 Ver Gráfico Financeiro", use_container_width=True):
        st.session_state.nav_mode = "Grafico_Valor"
        st.rerun()
      st.markdown("</div>", unsafe_allow_html=True)

  # ==========================================
  # TELA 3: GRÁFICO ESTÁTICO DE QUANTIDADE
  # ==========================================
  elif st.session_state.nav_mode == "Grafico_Qtd":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        "📊 Quantidade de Notas Fiscais por Categoria (Gráfico Fixo)"
    )

    todas_linhas = []
    for nome_aba, df in abas.items():
      df_temp = df.copy()
      df_temp["Categoria_Aba"] = nome_aba
      todas_linhas.append(df_temp)

    df_geral = pd.concat(todas_linhas, ignore_index=True)
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

    contagem_dados = {}
    for cat in ["Produto Bom", "Produto Ruim", "Recusa", "Reentrega"]:
      qtd = len(
          df_geral[
              df_geral["Categoria_Aba"].str.lower().str.contains(cat.lower())
          ]
      )
      contagem_dados[cat] = qtd

    fig, ax = plt.subplots(figsize=(8, 4.5))
    fig.patch.set_facecolor("#131b2e")
    ax.set_facecolor("#131b2e")

    bars = ax.bar(
        list(contagem_dados.keys()),
        list(contagem_dados.values()),
        color=["#238636", "#da3633", "#9e6a03", "#1f6feb"],
        width=0.5,
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
  # TELA 4: GRÁFICO ESTÁTICO DE VALOR
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
          valores_dados[cat] = round(
              pd.to_numeric(df_subset[col_val], errors="coerce").sum(), 2
          )
        else:
          valores_dados[cat] = 0.0

      fig, ax = plt.subplots(figsize=(7, 4.5))
      fig.patch.set_facecolor("#131b2e")
      ax.set_facecolor("#131b2e")

      bars = ax.bar(
          list(valores_dados.keys()),
          list(valores_dados.values()),
          color=["#238636", "#da3633"],
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
          matches_all = pd.DataFrame()

          for col in df_str.columns:
            sub = df_str[df_str[col].str.contains(str(termo_busca), na=False)]
            if not sub.empty:
              matches_all = pd.concat([matches_all, sub]).drop_duplicates()

          if not matches_all.empty:
            st.success(
                f"✨ Encontrado(s) {len(matches_all)} registo(s) na categoria"
                f" **{categoria_ativa}**!"
            )

            for idx, row in matches_all.iterrows():
              colunas_disponiveis = {c.lower().strip(): c for c in row.index}

              def pegar_val(nomes):
                for n in nomes:
                  if n.lower() in colunas_disponiveis:
                    val = row[colunas_disponiveis[n.lower()]]
                    return "N/D" if pd.isna(val) else val
                return "N/D"

              cliente = pegar_val(["Nome_Cliente", "Cliente"])
              cnpj = pegar_val(["CNPJ"])
              num_nf = pegar_val(["Numero_NF", "Número da NF", "NF", "Nota"])
              valor = pegar_val(["Valor da NF", "Valor", "R$"])
              chave = pegar_val(["Chave de acesso", "Chave"])
              emissao = pegar_val(["Data da emissão", "Emissão"])
              atualizacao = pegar_val(["Data da atualização", "Última Atualiz."])
              status_atual = str(
                  pegar_val(["Status", "Situação"])
              ).strip().lower()

              with st.container():
                st.markdown(
                    f"--- \n **🏢 Cliente:** `{cliente}` | **CNPJ:** `{cnpj}`"
                )
                c1, c2 = st.columns(2)
                with c1:
                  st.info(f"**Número da NF:** {num_nf}")
                  st.info(f"**Valor:** R$ {valor}")
                  st.info(f"**Emissão:** {emissao}")
                with c2:
                  st.info(f"**Chave de Acesso:**\n{chave}")
                  st.info(f"**Última Atualiz.:** {atualizacao}")

                st.markdown("**Status Operacional:**")
                if "conclu" in status_atual:
                  st.markdown("🟢 **Operação Concluída**")
                elif "process" in status_atual:
                  st.markdown("🟡 **Em Processamento na Logística**")
                elif "pendente" in status_atual:
                  st.markdown("🟠 **Pendente de Aprovação**")
                elif "não entregue" in status_atual or "nao entregue" in status_atual:
                  st.markdown("🔴 **Não Entregue / Ocorrência**")
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