from datetime import datetime
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="PepsiCo - Painel Multi-CD Operacional",
    page_icon="🔵",
    layout="wide",
)

# Estilização CSS com o visual glassmorphism azulado
st.markdown(
    """
    <style>
        .stApp {
            background: linear-gradient(135deg, #0d1b33 0%, #112240 50%, #1a365d 100%);
            color: #c9d1d9;
        }
        .main-header {
            font-size: 30px;
            font-weight: 800;
            color: #ffffff;
            margin-bottom: 0px;
            line-height: 1.2;
        }
        .sub-header {
            font-size: 14px;
            color: #8b949e;
            margin-top: 5px;
        }
        .welcome-card {
            background: rgba(23, 42, 69, 0.85);
            backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.12);
            padding: 40px;
            border-radius: 20px;
            text-align: center;
            box-shadow: 0 20px 40px rgba(0,0,0,0.5);
            max-width: 600px;
            margin: 40px auto;
        }
        .section-title {
            font-size: 20px;
            font-weight: 700;
            color: #ffffff;
            margin-bottom: 20px;
            letter-spacing: 0.5px;
        }
        .card-bom {
            background: linear-gradient(145deg, rgba(23, 42, 69, 0.9) 0%, rgba(15, 30, 50, 0.95) 100%);
            backdrop-filter: blur(14px);
            border: 2px solid #238636;
            padding: 24px;
            border-radius: 18px;
            text-align: center;
            box-shadow: 0 10px 30px rgba(35, 134, 54, 0.25);
            margin-bottom: 12px;
        }
        .card-ruim {
            background: linear-gradient(145deg, rgba(23, 42, 69, 0.9) 0%, rgba(15, 30, 50, 0.95) 100%);
            backdrop-filter: blur(14px);
            border: 2px solid #da3633;
            padding: 24px;
            border-radius: 18px;
            text-align: center;
            box-shadow: 0 10px 30px rgba(218, 54, 51, 0.25);
            margin-bottom: 12px;
        }
        .card-recusa {
            background: linear-gradient(145deg, rgba(23, 42, 69, 0.9) 0%, rgba(15, 30, 50, 0.95) 100%);
            backdrop-filter: blur(14px);
            border: 2px solid #d29922;
            padding: 24px;
            border-radius: 18px;
            text-align: center;
            box-shadow: 0 10px 30px rgba(210, 153, 34, 0.25);
            margin-bottom: 12px;
        }
        .card-reentrega {
            background: linear-gradient(145deg, rgba(23, 42, 69, 0.9) 0%, rgba(15, 30, 50, 0.95) 100%);
            backdrop-filter: blur(14px);
            border: 2px solid #1f6feb;
            padding: 24px;
            border-radius: 18px;
            text-align: center;
            box-shadow: 0 10px 30px rgba(31, 111, 235, 0.25);
            margin-bottom: 12px;
        }
        .footer {
            position: fixed;
            left: 0;
            bottom: 0;
            width: 100%;
            background-color: #0d1b33;
            color: #8b949e;
            text-align: center;
            padding: 8px;
            font-size: 12px;
            border-top: 1px solid rgba(255,255,255,0.1);
        }
    </style>
""",
    unsafe_allow_html=True,
)

ARQUIVO_EXCEL = "base_slips_nfs.xlsx"


@st.cache_data(ttl=1)
def carregar_dados():
  try:
    return pd.read_excel(ARQUIVO_EXCEL, sheet_name=None, dtype=str)
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

  if st.session_state.nav_mode == "Welcome":
    col_logo, col_titulo = st.columns([1.8, 5.5])
    with col_logo:
      try:
        st.image("logo.png", width=240)
      except:
        st.write("🔵")
    with col_titulo:
      st.markdown("<br>", unsafe_allow_html=True)
      st.markdown(
          '<p class="main-header">PepsiCo - Painel Multi-CD NFs e Slips</p>',
          unsafe_allow_html=True,
      )
      st.markdown(
          '<p class="sub-header">Gestão Operacional Integrada para Múltiplas'
          " Filiais (CDv)</p>",
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="welcome-card">
            <h1 style="color: white; font-size: 28px; margin-bottom: 12px;">Seja Bem-Vindo ao Sistema Multi-CD! 👋</h1>
            <p style="color: #8b949e; font-size: 15px; margin-bottom: 30px; line-height: 1.5;">
                Plataforma oficial de consulta e controlo de Notas Fiscais e Slips para as filiais CDV01 a CDV30 da PepsiCo.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col_btn, col2 = st.columns([2, 2, 2])
    with col_btn:
      if st.button(
          "🚀 Entrar no Painel", use_container_width=True, type="primary"
      ):
        st.session_state.nav_mode = "Home"
        st.rerun()

  elif st.session_state.nav_mode == "Home":
    with st.sidebar:
      st.markdown("### 🏢 Seletor de Filial (CD)")
      if st.button("🔄 Atualizar Base do Excel", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

      todas_filiais = []
      for _, df in abas.items():
        if "Filial" in df.columns:
          todas_filiais.extend(df["Filial"].dropna().astype(str).unique())
      lista_filiais = sorted(list(set(todas_filiais)))

      if not lista_filiais:
        lista_filiais = [f"CDV{i:02d}" for i in range(1, 31)]

      filial_selecionada = st.selectbox(
          "Selecione o CD / Filial:", ["Todas as Filiais"] + lista_filiais
      )
      st.markdown("---")
      if st.button("⬅️ Voltar à Tela Inicial", use_container_width=True):
        st.session_state.nav_mode = "Welcome"
        st.rerun()

    st.session_state.filial_ativa = filial_selecionada

    col_logo, col_titulo = st.columns([1.8, 5.5])
    with col_logo:
      try:
        st.image("logo.png", width=220)
      except:
        st.write("🔵")
    with col_titulo:
      st.markdown("<br>", unsafe_allow_html=True)
      st.markdown(
          '<p class="main-header">PepsiCo - Painel Multi-CD</p>',
          unsafe_allow_html=True,
      )
      st.markdown(
          f'<p class="sub-header">Filial Ativa: <b>{filial_selecionada}</b> |'
          " Controlo Operacional em Tempo Real</p>",
          unsafe_allow_html=True,
      )

    st.markdown("---")
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        '<p class="section-title">📂 Consultas e Gestão por Categoria</p>',
        unsafe_allow_html=True,
    )
    c1, c2, c3, c4 = st.columns(4)

    with c1:
      st.markdown(
          """
            <div class="card-bom">
                <div style="font-size: 26px; margin-bottom: 6px;">📦🟢</div>
                <div style="color: #238636; font-size: 17px; font-weight: 700; margin-bottom: 6px;">Produto Bom</div>
                <div style="color: #8b949e; font-size: 12px; min-height: 38px; line-height: 1.3;">Notas faturadas sem ocorrências operacionais.</div>
            </div>
            """,
          unsafe_allow_html=True,
      )
      if st.button("Consultar Bom", use_container_width=True, key="btn_bom"):
        st.session_state.nav_mode = "Produto Bom"
        st.rerun()

    with c2:
      st.markdown(
          """
            <div class="card-ruim">
                <div style="font-size: 26px; margin-bottom: 6px;">📦❌</div>
                <div style="color: #da3633; font-size: 17px; font-weight: 700; margin-bottom: 6px;">Ruim</div>
                <div style="color: #8b949e; font-size: 12px; min-height: 38px; line-height: 1.3;">Notas fiscais com avarias ou devoluções.</div>
            </div>
            """,
          unsafe_allow_html=True,
      )
      if st.button(
          "Consultar Ruim", use_container_width=True, key="btn_ruim"
      ):
        st.session_state.nav_mode = "Produto Ruim"
        st.rerun()

    with c3:
      st.markdown(
          """
            <div class="card-recusa">
                <div style="font-size: 26px; margin-bottom: 6px;">📦⚠️</div>
                <div style="color: #d29922; font-size: 17px; font-weight: 700; margin-bottom: 6px;">Recusa</div>
                <div style="color: #8b949e; font-size: 12px; min-height: 38px; line-height: 1.3;">Acompanhamento de recusas de entrega.</div>
            </div>
            """,
          unsafe_allow_html=True,
      )
      if st.button(
          "Consultar Recusa", use_container_width=True, key="btn_recusa"
      ):
        st.session_state.nav_mode = "Recusa"
        st.rerun()

    with c4:
      st.markdown(
          """
            <div class="card-reentrega">
                <div style="font-size: 26px; margin-bottom: 6px;">🚚📦</div>
                <div style="color: #58a6ff; font-size: 17px; font-weight: 700; margin-bottom: 6px;">Reentrega</div>
                <div style="color: #8b949e; font-size: 12px; min-height: 38px; line-height: 1.3;">Monitorização de notas para novas rotas.</div>
            </div>
            """,
          unsafe_allow_html=True,
      )
      if st.button(
          "Consultar Reentrega", use_container_width=True, key="btn_reentrega"
      ):
        st.session_state.nav_mode = "Reentrega"
        st.rerun()

    st.markdown(
        "<br><hr style='border-color: rgba(255,255,255,0.1);'><br>",
        unsafe_allow_html=True,
    )

    st.markdown(
        '<p class="section-title">📊 Painéis Analíticos e Gráficos Estáticos</p>',
        unsafe_allow_html=True,
    )
    g1, g2 = st.columns(2)

    with g1:
      st.markdown(
          """
            <div style="background: linear-gradient(145deg, rgba(23, 42, 69, 0.9) 0%, rgba(15, 30, 50, 0.95) 100%); backdrop-filter: blur(14px); border: 1px solid rgba(255,255,255,0.1); padding: 22px; border-radius: 18px; text-align: center; margin-bottom: 12px;">
                <div style="font-size: 26px; margin-bottom: 6px;">📈</div>
                <div style="color: #ffffff; font-size: 17px; font-weight: 700; margin-bottom: 6px;">Volume de Notas Fiscais</div>
                <div style="color: #8b949e; font-size: 12px; min-height: 38px; line-height: 1.3;">Gráfico estatístico filtrado por filial e período.</div>
            </div>
            """,
          unsafe_allow_html=True,
      )
      if st.button("Ver Gráfico de Quantidade", use_container_width=True):
        st.session_state.nav_mode = "Grafico_Qtd"
        st.rerun()

    with g2:
      st.markdown(
          """
            <div style="background: linear-gradient(145deg, rgba(23, 42, 69, 0.9) 0%, rgba(15, 30, 50, 0.95) 100%); backdrop-filter: blur(14px); border: 1px solid rgba(255,255,255,0.1); padding: 22px; border-radius: 18px; text-align: center; margin-bottom: 12px;">
                <div style="font-size: 26px; margin-bottom: 6px;">💰</div>
                <div style="color: #ffffff; font-size: 17px; font-weight: 700; margin-bottom: 6px;">Montante Financeiro (R$)</div>
                <div style="color: #8b949e; font-size: 12px; min-height: 38px; line-height: 1.3;">Análise de valores filtrada por filial e período.</div>
            </div>
            """,
          unsafe_allow_html=True,
      )
      if st.button("Ver Gráfico Financeiro", use_container_width=True):
        st.session_state.nav_mode = "Grafico_Valor"
        st.rerun()

  elif st.session_state.nav_mode == "Grafico_Qtd":
    filial_ativa = st.session_state.get("filial_ativa", "Todas as Filiais")
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        f"📊 Quantidade de Notas Fiscais por Categoria (Filial:"
        f" {filial_ativa})"
    )

    todas_linhas = []
    for nome_aba, df in abas.items():
      df_temp = df.copy()
      df_temp["Categoria_Aba"] = nome_aba
      todas_linhas.append(df_temp)

    df_geral = pd.concat(todas_linhas, ignore_index=True)
    if filial_ativa != "Todas as Filiais" and "Filial" in df_geral.columns:
      df_geral = df_geral[df_geral["Filial"].astype(str) == filial_ativa]

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
    fig.patch.set_facecolor("#112240")
    ax.set_facecolor("#112240")

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

  elif st.session_state.nav_mode == "Grafico_Valor":
    filial_ativa = st.session_state.get("filial_ativa", "Todas as Filiais")
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        f"💰 Montante Financeiro - Produto Bom & Ruim (Filial: {filial_ativa})"
    )

    todas_linhas_val = []
    for nome_aba, df in abas.items():
      if "bom" in nome_aba.lower() or "ruim" in nome_aba.lower():
        df_temp = df.copy()
        df_temp["Categoria_Aba"] = nome_aba
        todas_linhas_val.append(df_temp)

    if todas_linhas_val:
      df_val_geral = pd.concat(todas_linhas_val, ignore_index=True)
      if filial_ativa != "Todas as Filiais" and "Filial" in df_val_geral.columns:
        df_val_geral = df_val_geral[
            df_val_geral["Filial"].astype(str) == filial_ativa
        ]

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
      fig.patch.set_facecolor("#112240")
      ax.set_facecolor("#112240")

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

  else:
    categoria_ativa = st.session_state.nav_mode
    filial_ativa = st.session_state.get("filial_ativa", "Todas as Filiais")

    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        f"🔍 Consultar em: {categoria_ativa} | Filial: {filial_ativa}"
    )

    aba_correspondente = None
    for nome_aba in abas.keys():
      if categoria_ativa.lower() in nome_aba.lower():
        aba_correspondente = nome_aba
        break

    if aba_correspondente:
      df_cat = abas[aba_correspondente]
      if filial_ativa != "Todas as Filiais" and "Filial" in df_cat.columns:
        df_cat = df_cat[df_cat["Filial"].astype(str) == filial_ativa]

      termo_busca = st.text_input(
          f"Digite o Número da NF ou Chave de Acesso:",
          placeholder="Ex: 10001 ou 352610...",
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
                f"✨ Encontrado(s) {len(matches_all)} registo(s) na filial"
                f" **{filial_ativa}**!"
            )

            for idx, row in matches_all.iterrows():
              colunas_disponiveis = {c.lower().strip(): c for c in row.index}

              def pegar_val(nomes):
                for n in nomes:
                  if n.lower() in colunas_disponiveis:
                    val = row[colunas_disponiveis[n.lower()]]
                    return "N/D" if pd.isna(val) else val
                return "N/D"

              filial_reg = pegar_val(["Filial"])
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
                    f"--- \n 🏢 **Filial:** `{filial_reg}` | **Cliente:**"
                    f" `{cliente}` | **CNPJ:** `{cnpj}`"
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
                f"Nenhum registo encontrado na filial '{filial_ativa}' com o"
                " termo informado."
            )
    else:
      st.warning(f"A aba '{categoria_ativa}' não foi encontrada na planilha.")

st.markdown(
    '<div class="footer">© PepsiCo Brasil - Sistema Multi-CD | Operacional'
    " de NFs e Slips</div>",
    unsafe_allow_html=True,
)