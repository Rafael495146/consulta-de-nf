from datetime import datetime
import io
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import shutil
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="PepsiCo - Painel Multi-CDV Operacional",
    page_icon="🔵",
    layout="wide",
)

# Estilização CSS avançada, com seletor universal para todos os botões de Voltar
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
        
        /* Pinta de vermelho qualquer botão de voltar */
        [data-testid="stSidebar"] button:has(div p:contains("Voltar")),
        .main button:has(div p:contains("Voltar")),
        div.stButton > button:has(p:contains("Voltar")) {
            background: linear-gradient(135deg, #da3633 0%, #f85149 100%) !important;
            color: #ffffff !important;
            border: 1px solid #ff7b72 !important;
            font-weight: bold !important;
        }
        [data-testid="stSidebar"] button:has(div p:contains("Voltar")):hover,
        .main button:has(div p:contains("Voltar")):hover,
        div.stButton > button:has(p:contains("Voltar")):hover {
            background: linear-gradient(135deg, #b62324 0%, #da3633 100%) !important;
            border: 1px solid #ffa198 !important;
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

# Caminhos dos diretórios e ficheiros no OneDrive
DIRETORIO_ATUAL = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EXCEL = os.path.join(DIRETORIO_ATUAL, "base_slips_nfs.xlsx")
PASTA_HISTORICOS = os.path.join(DIRETORIO_ATUAL, "historicos_mensais")
ARQUIVO_DESTINATARIOS = os.path.join(DIRETORIO_ATUAL, "destinatarios.xlsx")
ARQUIVO_LOG_ENVIOS = os.path.join(DIRETORIO_ATUAL, "historico_envios.xlsx")


@st.cache_data(ttl=1)
def carregar_dados_infinitos():
  dados_consolidados = {}

  if os.path.exists(ARQUIVO_EXCEL):
    try:
      arquivo_temp = os.path.join(DIRETORIO_ATUAL, "temp_leitura_base.xlsx")
      shutil.copy2(ARQUIVO_EXCEL, arquivo_temp)
      df_principal = pd.read_excel(arquivo_temp, sheet_name=None, dtype=str)

      for aba, df in df_principal.items():
        dados_consolidados[aba] = [df]
    except Exception as e:
      pass

  if os.path.exists(PASTA_HISTORICOS):
    for ficheiro in os.listdir(PASTA_HISTORICOS):
      if ficheiro.endswith(".xlsx") and not ficheiro.startswith("~$"):
        caminho_hist = os.path.join(PASTA_HISTORICOS, ficheiro)
        try:
          df_hist = pd.read_excel(caminho_hist, sheet_name=None, dtype=str)
          for aba, df in df_hist.items():
            if aba not in dados_consolidados:
              dados_consolidados[aba] = []
            dados_consolidados[aba].append(df)
        except Exception as e:
          pass

  abas_finais = {}
  if dados_consolidados:
    for aba, lista_dfs in dados_consolidados.items():
      df_concatenado = pd.concat(lista_dfs, ignore_index=True)
      df_concatenado.columns = [
          str(c).strip() for c in df_concatenado.columns
      ]
      df_concatenado = df_concatenado.loc[
          :, ~df_concatenado.columns.duplicated()
      ]
      abas_finais[aba] = df_concatenado
    return abas_finais

  return None


def carregar_destinatarios():
  if os.path.exists(ARQUIVO_DESTINATARIOS):
    try:
      return pd.read_excel(ARQUIVO_DESTINATARIOS, dtype=str)
    except:
      pass
  df_padrao = pd.DataFrame(columns=["Nome", "Cargo", "Filial", "E-mail"])
  df_padrao.to_excel(ARQUIVO_DESTINATARIOS, index=False)
  return df_padrao


def carregar_log_envios():
  if os.path.exists(ARQUIVO_LOG_ENVIOS):
    try:
      return pd.read_excel(ARQUIVO_LOG_ENVIOS, dtype=str)
    except:
      pass
  return pd.DataFrame(
      columns=["Data/Hora", "Filial", "Categoria", "Período", "Destinatário"]
  )


def registar_envio_log(filial, categoria, periodo, destinatario):
  df_log = carregar_log_envios()
  novo_reg = pd.DataFrame(
      [[datetime.now().strftime("%Y-%m-%d %H:%M:%S"), filial, categoria, periodo, destinatario]],
      columns=["Data/Hora", "Filial", "Categoria", "Período", "Destinatário"],
  )
  df_log = pd.concat([df_log, novo_reg], ignore_index=True)
  df_log.to_excel(ARQUIVO_LOG_ENVIOS, index=False)


abas = carregar_dados_infinitos()

if abas is None:
  st.error(
      f"❌ O ficheiro 'base_slips_nfs.xlsx' não foi encontrado na pasta:"
      f" {ARQUIVO_EXCEL}"
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
          '<p class="main-header">PepsiCo - Painel Multi-CDV NFs e Slips</p>',
          unsafe_allow_html=True,
      )
      st.markdown(
          '<p class="sub-header">Gestão Operacional Integrada para Múltiplas'
          " Filiais (CDV) - Base Infinita</p>",
          unsafe_allow_html=True,
      )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="welcome-card">
            <h1 style="color: white; font-size: 28px; margin-bottom: 12px;">Seja Bem-Vindo ao Sistema Multi-CDV! 👋</h1>
            <p style="color: #8b949e; font-size: 15px; margin-bottom: 30px; line-height: 1.5;">
                Plataforma oficial de consulta e controlo de Notas Fiscais e Slips para as filiais da PepsiCo.
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
      st.markdown("### 🏢 Seletor de Filial (CDV)")
      if st.button(
          "🔄 Atualizar Base do Excel",
          use_container_width=True,
          type="primary",
      ):
        st.cache_data.clear()
        st.rerun()

      todas_filiais = []
      for _, df in abas.items():
        if "Filial" in df.columns:
          todas_filiais.extend(df["Filial"].dropna().astype(str).unique())
      lista_filiais = sorted(list(set(todas_filiais)))

      if not lista_filiais:
        lista_filiais = [f"CDV{i:02d}" for i in range(1, 35)]

      filial_selecionada = st.selectbox(
          "Selecione o CDV / Filial:", ["Todas as Filiais"] + lista_filiais
      )

      st.markdown("---")
      if st.button(
          "📥 Exportar Dados (Excel)", use_container_width=True, type="secondary"
      ):
        st.session_state.nav_mode = "Exportar_Dados"
        st.rerun()

      if st.button(
          "📧 Enviar Relatório por E-mail",
          use_container_width=True,
          type="secondary",
      ):
        st.session_state.nav_mode = "Enviar_Email"
        st.rerun()

      if st.button(
          "📊 Painel Comparativo de CDVs",
          use_container_width=True,
          type="secondary",
      ):
        st.session_state.nav_mode = "Comparativo_CDVs"
        st.rerun()

      if st.button(
          "⚖️ Dashboard Consolidado (4 Categorias)",
          use_container_width=True,
          type="secondary",
      ):
        st.session_state.nav_mode = "Dashboard_Geral"
        st.rerun()

      if st.button(
          "🗂️ Histórico de Envios (Auditoria)",
          use_container_width=True,
          type="secondary",
      ):
        st.session_state.nav_mode = "Historico_Auditoria"
        st.rerun()

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
          '<p class="main-header">PepsiCo - Painel Multi-CDV</p>',
          unsafe_allow_html=True,
      )
      st.markdown(
          f'<p class="sub-header">Filial Ativa: <b>{filial_selecionada}</b> |'
          " Controle Operacional em Tempo Real</p>",
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

  elif st.session_state.nav_mode == "Dashboard_Geral":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        "⚖️ Dashboard Consolidado por Filial (Produto Bom, Ruim, Recusa e"
        " Reentrega)"
    )
    st.markdown(
        "Análise visual interativa em primeiro plano e tabela detalhada em"
        " baixo."
    )

    lista_dfs_cat = []
    for nome_aba, df in abas.items():
      df_temp = df.copy()
      cat_nome = "Outros"
      if "bom" in nome_aba.lower():
        cat_nome = "Produto Bom"
      elif "ruim" in nome_aba.lower():
        cat_nome = "Produto Ruim"
      elif "recusa" in nome_aba.lower():
        cat_nome = "Recusa"
      elif "reentrega" in nome_aba.lower():
        cat_nome = "Reentrega"

      if cat_nome != "Outros" and "Filial" in df_temp.columns:
        df_temp["Categoria_Operacional"] = cat_nome
        lista_dfs_cat.append(df_temp[["Filial", "Categoria_Operacional"]])

    if lista_dfs_cat:
      df_geral_cat = pd.concat(lista_dfs_cat, ignore_index=True)
      df_geral_cat["Filial"] = df_geral_cat["Filial"].astype(str).str.strip()

      tabela_pivot = (
          df_geral_cat.pivot_table(
              index="Filial",
              columns="Categoria_Operacional",
              values="Filial",
              aggfunc="count",
              fill_value=0,
          )
          .reset_index()
      )

      for col_essencial in [
          "Produto Bom",
          "Produto Ruim",
          "Recusa",
          "Reentrega",
      ]:
        if col_essencial not in tabela_pivot.columns:
          tabela_pivot[col_essencial] = 0

      tabela_pivot["Total_Geral"] = (
          tabela_pivot["Produto Bom"]
          + tabela_pivot["Produto Ruim"]
          + tabela_pivot["Recusa"]
          + tabela_pivot["Reentrega"]
      )
      tabela_pivot = tabela_pivot.sort_values(by="Total_Geral", ascending=False)

      # --- FILTRO / LUPA DE SELEÇÃO PARA O GRÁFICO ---
      st.markdown("---")
      filtro_grafico = st.selectbox(
          "🔍 Filtrar Categoria no Gráfico Lateral:",
          [
              "Todas as Categorias (Consolidado)",
              "Produto Bom",
              "Produto Ruim",
              "Recusa",
              "Reentrega",
          ],
      )

      st.markdown("### 📊 Gráfico Comparativo Lateral por Filial")
      col_v1, col_g, col_v2 = st.columns([0.1, 5.8, 0.1])
      with col_g:
        fig, ax = plt.subplots(
            figsize=(10, max(4, len(tabela_pivot) * 0.35))
        )
        fig.patch.set_facecolor("#112240")
        ax.set_facecolor("#112240")

        top_piv = tabela_pivot.iloc[::-1]
        y_pos = range(len(top_piv))

        if filtro_grafico == "Todas as Categorias (Consolidado)":
          largura = 0.2
          ax.barh(
              [i - 1.5 * largura for i in y_pos],
              top_piv["Produto Bom"],
              height=largura,
              color="#238636",
              label="Prod. Bom",
          )
          ax.barh(
              [i - 0.5 * largura for i in y_pos],
              top_piv["Produto Ruim"],
              height=largura,
              color="#da3633",
              label="Prod. Ruim",
          )
          ax.barh(
              [i + 0.5 * largura for i in y_pos],
              top_piv["Recusa"],
              height=largura,
              color="#d29922",
              label="Recusa",
          )
          ax.barh(
              [i + 1.5 * largura for i in y_pos],
              top_piv["Reentrega"],
              height=largura,
              color="#1f6feb",
              label="Reentrega",
          )
        else:
          cor_map = {
              "Produto Bom": "#238636",
              "Produto Ruim": "#da3633",
              "Recusa": "#d29922",
              "Reentrega": "#1f6feb",
          }
          ax.barh(
              y_pos,
              top_piv[filtro_grafico],
              height=0.5,
              color=cor_map.get(filtro_grafico, "#1f6feb"),
              label=filtro_grafico,
          )

        ax.set_yticks(list(y_pos))
        ax.set_yticklabels(
            top_piv["Filial"].astype(str), color="#c9d1d9", fontsize=9
        )
        ax.tick_params(colors="#c9d1d9", labelsize=9)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#30363d")
        ax.spines["bottom"].set_color("#30363d")
        ax.grid(axis="x", linestyle="--", alpha=0.2, color="#8b949e")
        ax.legend(
            facecolor="#112240",
            edgecolor="none",
            labelcolor="white",
            fontsize=8,
            loc="lower right",
        )

        st.pyplot(fig)

      st.markdown("---")
      st.markdown("### 📋 Tabela Completa Ordenada (Mais para Menos Lançamentos)")
      st.dataframe(tabela_pivot, use_container_width=True)
    else:
      st.warning(
          "Não foram encontradas bases válidas com a coluna 'Filial' para"
          " gerar o dashboard consolidado."
      )

  elif st.session_state.nav_mode == "Comparativo_CDVs":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        "📊 Painel Comparativo Avançado entre Filiais (CDVs)"
    )
    st.markdown(
        "Compare o desempenho operacional e volume de ocorrências entre"
        " diferentes filiais da PepsiCo lado a lado."
    )

    cat_comp = st.selectbox(
        "Selecione a Categoria para Comparação:", list(abas.keys())
    )
    df_comp = abas[cat_comp].copy()

    if "Filial" in df_comp.columns:
      resumo_filial = (
          df_comp.groupby("Filial")
          .size()
          .reset_index(name="Quantidade_Notas")
          .sort_values(by="Quantidade_Notas", ascending=False)
      )

      col_val_comp = None
      for c in df_comp.columns:
        if "valor" in c.lower() or "r$" in c.lower():
          col_val_comp = c
          break

      if col_val_comp:
        df_comp["Valor_Num"] = pd.to_numeric(
            df_comp[col_val_comp], errors="coerce"
        ).fillna(0)
        resumo_valor = (
            df_comp.groupby("Filial")["Valor_Num"]
            .sum()
            .reset_index(name="Montante_Total")
        )
        resumo_filial = pd.merge(resumo_filial, resumo_valor, on="Filial")

      st.markdown("### 🏆 Ranking de Ocorrências por Filial")
      st.dataframe(resumo_filial, use_container_width=True)

      col_vazia1, col_grafico_comp, col_vazia2 = st.columns([1, 3, 1])
      with col_grafico_comp:
        fig, ax = plt.subplots(figsize=(6, 3.2))
        fig.patch.set_facecolor("#112240")
        ax.set_facecolor("#112240")

        top_filiais = resumo_filial.head(10)
        bars = ax.bar(
            top_filiais["Filial"].astype(str),
            top_filiais["Quantidade_Notas"],
            color="#1f6feb",
            width=0.45,
            edgecolor="none",
        )
        ax.tick_params(colors="#c9d1d9", labelsize=9)
        ax.spines["bottom"].set_color("#30363d")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#30363d")
        ax.grid(axis="y", linestyle="--", alpha=0.2, color="#8b949e")
        plt.xticks(rotation=30)

        st.pyplot(fig)
    else:
      st.warning(
          "A coluna 'Filial' não foi encontrada nesta base para realizar a"
          " comparação."
      )

  elif st.session_state.nav_mode == "Historico_Auditoria":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        "🗂️ Módulo de Histórico de Envios e Logs (Auditoria)"
    )
    st.markdown(
        "Registo automático de todos os relatórios enviados pelo sistema para"
        " rastreabilidade corporativa."
    )

    df_logs = carregar_log_envios()
    if not df_logs.empty:
      st.dataframe(df_logs, use_container_width=True)

      output_log = io.BytesIO()
      with pd.ExcelWriter(output_log, engine="openpyxl") as writer:
        df_logs.to_excel(writer, index=False, sheet_name="Logs_Envios")
      log_data = output_log.getvalue()

      st.download_button(
          label="📥 Baixar Histórico de Auditoria em Excel",
          data=log_data,
          file_name="Auditoria_Envios_PepsiCo.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )
    else:
      st.info(
          "Ainda não existem registos de envio de e-mail efetuados nesta"
          " sessão."
      )

  elif st.session_state.nav_mode == "Exportar_Dados":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader(
        "📥 Central de Exportação de Relatórios em Excel (.xlsx)"
    )
    st.markdown(
        "Filtre os dados por categoria, filial ou período e descarregue o"
        " relatório personalizado instantaneamente."
    )

    col_e1, col_e2 = st.columns(2)

    with col_e1:
      categoria_escolhida = st.selectbox(
          "Selecione a Categoria/Aba:", list(abas.keys())
      )

    df_export = abas[categoria_escolhida].copy()

    lista_filiais_exp = ["Todas as Filiais"]
    if "Filial" in df_export.columns:
      lista_filiais_exp += sorted(
          list(df_export["Filial"].dropna().astype(str).unique())
      )

    with col_e2:
      filial_escolhida_exp = st.selectbox(
          "Filtrar por Filial (CDV):", lista_filiais_exp
      )

    if filial_escolhida_exp != "Todas as Filiais" and "Filial" in df_export.columns:
      df_export = df_export[df_export["Filial"].astype(str) == filial_escolhida_exp]

    col_emissao_exp = None
    for c in df_export.columns:
      if "emiss" in c.lower():
        col_emissao_exp = c
        break

    if col_emissao_exp:
      df_export["AnoMes_Filtro"] = (
          pd.to_datetime(df_export[col_emissao_exp], errors="coerce")
          .dt.strftime("%Y-%m")
          .fillna("Outros/Geral")
      )
      meses_exp = ["Todos os Meses"] + sorted(
          [m for m in df_export["AnoMes_Filtro"].unique() if m != "Outros/Geral"]
      )
      mes_escolhido_exp = st.selectbox("Filtrar por Mês/Ano:", meses_exp)

      if mes_escolhido_exp != "Todos os Meses":
        df_export = df_export[df_export["AnoMes_Filtro"] == mes_escolhido_exp]
      df_export = df_export.drop(columns=["AnoMes_Filtro"])

    busca_livre = st.text_input(
        "🔎 Filtrar por Cliente / Supermercado / CNPJ (Opcional):",
        placeholder="Deixe em branco para extrair tudo...",
    )
    if busca_livre:
      df_str_exp = df_export.astype(str)
      mask = df_str_exp.apply(
          lambda col: col.str.contains(busca_livre, case=False, na=False)
      ).any(axis=1)
      df_export = df_export[mask]

    st.markdown(f"**Registos encontrados para exportação:** `{len(df_export)}`")

    if not df_export.empty:
      with st.expander("👁️ Pré-visualizar dados antes de baixar"):
        st.dataframe(df_export.head(100), use_container_width=True)

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Relatorio_PepsiCo")
      excel_data = output.getvalue()

      nome_arquivo = f"Relatorio_{categoria_escolhida.replace(' ', '_')}_{filial_escolhida_exp}.xlsx"

      st.download_button(
          label="📥 Clique aqui para baixar o ficheiro Excel (.xlsx)",
          data=excel_data,
          file_name=nome_arquivo,
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )
    else:
      st.warning(
          "Nenhum registo encontrado com os filtros selecionados para exportar."
      )

  elif st.session_state.nav_mode == "Enviar_Email":
    if st.button("⬅️ Voltar ao Menu Principal"):
      st.session_state.nav_mode = "Home"
      st.rerun()

    st.markdown("---")
    st.subheader("📧 Envio Automatizado de Relatório por E-mail com Anexo")
    st.markdown(
        "Selecione um destinatário da agenda ou digite manualmente, defina os"
        " filtros e envie o relatório."
    )

    df_dest = carregar_destinatarios()

    with st.expander("📋 Gerir Agenda de Destinatários (Opcional)", expanded=False):
      st.markdown(
          "Adicione ou remova contactos frequentes. Eles ficarão salvos na"
          " agenda."
      )
      col_cad1, col_cad2, col_cad3, col_cad4 = st.columns(4)
      with col_cad1:
        novo_nome = st.text_input("Nome do Gestor:")
      with col_cad2:
        novo_cargo = st.text_input("Cargo:")
      with col_cad3:
        nova_filial = st.text_input("Filial/CDV:")
      with col_cad4:
        novo_email = st.text_input("E-mail corporativo:")

      if st.button("➕ Adicionar à Agenda"):
        if novo_nome and novo_email:
          novo_reg = pd.DataFrame(
              [[novo_nome, novo_cargo, nova_filial, novo_email]],
              columns=["Nome", "Cargo", "Filial", "E-mail"],
          )
          df_dest = pd.concat([df_dest, novo_reg], ignore_index=True)
          df_dest.to_excel(ARQUIVO_DESTINATARIOS, index=False)
          st.success(
              f"✨ Contacto **{novo_nome}** adicionado à agenda com sucesso!"
          )
          st.rerun()
        else:
          st.warning("⚠️ Preencha pelo menos o Nome e o E-mail.")

      if not df_dest.empty:
        st.markdown("**Contactos Atuais na Agenda:**")
        st.dataframe(df_dest, use_container_width=True)

        email_para_remover = st.selectbox(
            "Selecione um e-mail para remover da agenda (opcional):",
            ["Nenhum"] + list(df_dest["E-mail"].dropna().unique()),
        )
        if email_para_remover != "Nenhum":
          if st.button("🗑️ Remover Contacto Selecionado"):
            df_dest = df_dest[df_dest["E-mail"] != email_para_remover]
            df_dest.to_excel(ARQUIVO_DESTINATARIOS, index=False)
            st.success("🗑️ Contacto removido com sucesso!")
            st.rerun()

    st.markdown("---")

    modo_envio_email = st.radio(
        "Como deseja definir o destinatário?",
        [
            "Selecionar da Agenda de Destinatários",
            "Digitar E-mail Manualmente",
        ],
        horizontal=True,
    )

    if modo_envio_email == "Selecionar da Agenda de Destinatários":
      if not df_dest.empty and "E-mail" in df_dest.columns:
        opcoes_agenda = (
            df_dest["Nome"].fillna("Sem Nome")
            + " ("
            + df_dest["E-mail"]
            + ")"
        ).tolist()
        escolha_agenda = st.selectbox(
            "Escolher Contacto da Agenda:", ["Selecione..."] + opcoes_agenda
        )
        if escolha_agenda != "Selecione...":
          email_destino = escolha_agenda.split("(")[-1].replace(")", "").strip()
        else:
          email_destino = ""
      else:
        st.warning(
            "A agenda está vazia. Cadastre contactos acima ou mude para digitação"
            " manual."
        )
        email_destino = ""
    else:
      email_destino = st.text_input(
          "E-mail do Destinatário (Manual):", placeholder="exemplo@outlook.com"
      )

    col_m1, col_m2, col_m3 = st.columns(3)

    todas_filiais_disp = []
    for _, df in abas.items():
      if "Filial" in df.columns:
        todas_filiais_disp.extend(df["Filial"].dropna().astype(str).unique())
    lista_cdv_opcoes = ["Todas as Filiais"] + sorted(
        list(set(todas_filiais_disp))
    )

    with col_m1:
      cdv_escolhido = st.selectbox("Escolher CDV:", lista_cdv_opcoes)

    meses_disponiveis_geral = set()
    for _, df in abas.items():
      for c in df.columns:
        if "emiss" in c.lower():
          anos_meses = (
              pd.to_datetime(df[c], errors="coerce")
              .dt.strftime("%Y-%m")
              .dropna()
              .unique()
          )
          meses_disponiveis_geral.update(anos_meses)
    lista_meses_opcoes = ["Todos os Meses"] + sorted(
        list(meses_disponiveis_geral)
    )

    with col_m2:
      periodo_escolhido = st.selectbox("Escolher Mês/Ano:", lista_meses_opcoes)

    lista_bases_opcoes = [
        "Selecione a categoria...",
        "Todas as Bases",
        "Produto Bom",
        "Produto Ruim",
        "Recusa",
        "Reentrega",
    ]
    with col_m3:
      base_escolhida = st.selectbox(
          "Escolher Base (Obrigatório):", lista_bases_opcoes
      )

    dfs_para_anexo = []
    if base_escolhida != "Selecione a categoria...":
      if base_escolhida == "Todas as Bases":
        bases_alvo = list(abas.keys())
      else:
        bases_alvo = [
            b
            for b in abas.keys()
            if base_escolhida.lower() in b.lower()
        ]

      for b in bases_alvo:
        df_temp = abas[b].copy()
        df_temp["Base_Origem"] = b

        if cdv_escolhido != "Todas as Filiais" and "Filial" in df_temp.columns:
          df_temp = df_temp[df_temp["Filial"].astype(str) == cdv_escolhido]

        if periodo_escolhido != "Todos os Meses":
          col_emiss = None
          for c in df_temp.columns:
            if "emiss" in c.lower():
              col_emiss = c
              break
          if col_emiss:
            df_temp["Periodo_Filtro"] = (
                pd.to_datetime(df_temp[col_emiss].astype(str), errors="coerce")
                .dt.strftime("%Y-%m")
            )
            df_temp = df_temp[
                df_temp["Periodo_Filtro"] == periodo_escolhido
            ]
            df_temp = df_temp.drop(columns=["Periodo_Filtro"])

        dfs_para_anexo.append(df_temp)

    df_final_email = (
        pd.concat(dfs_para_anexo, ignore_index=True)
        if dfs_para_anexo
        else pd.DataFrame()
    )
    total_regs_email = len(df_final_email)

    cat_texto = (
        base_escolhida
        if base_escolhida != "Selecione a categoria..."
        else "..."
    )

    mensagem_travada = f"""Olá,

Segue em anexo o relatório operacional da categoria '{cat_texto}' para a filial '{cdv_escolhido}'.
Total de registos: {total_regs_email}.

Atenciosamente,
PepsiCo do Brasil LTDA"""

    st.text_area(
        "Mensagem do E-mail (Modelo Padrão Imutável):",
        value=mensagem_travada,
        height=160,
        disabled=True,
    )

    excel_anexo_bytes = None
    if total_regs_email > 0:
      output_email = io.BytesIO()
      with pd.ExcelWriter(output_email, engine="openpyxl") as writer:
        df_final_email.to_excel(
            writer, index=False, sheet_name="Relatorio_Envio"
        )
      excel_anexo_bytes = output_email.getvalue()

      nome_anexo_excel = f"Relatorio_PepsiCo_{base_escolhida.replace(' ', '_')}.xlsx"
      st.download_button(
          label="📎 Pré-visualizar e Baixar Ficheiro em Anexo",
          data=excel_anexo_bytes,
          file_name=nome_anexo_excel,
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 Disparar E-mail com Anexo", type="primary"):
      if not email_destino:
        st.warning(
            "⚠️ Por favor, selecione ou digite o e-mail do destinatário."
        )
      elif base_escolhida == "Selecione a categoria...":
        st.error(
            "🚨 É obrigatório escolher uma base válida para prosseguir com o"
            " envio!"
        )
      elif total_regs_email == 0:
        st.warning(
            "⚠️ Nenhum registo encontrado com os filtros selecionados para"
            " enviar no anexo."
        )
      else:
        try:
          remetente_email = "rafaelsilva3365@outlook.com"
          senha_email = "SUA_SENHA_DE_APLICATIVO_AQUI"

          msg = MIMEMultipart()
          msg["From"] = remetente_email
          msg["To"] = email_destino
          msg["Subject"] = (
              f"PepsiCo - Relatório Operacional ({base_escolhida} - CDV:"
              f" {cdv_escolhido})"
          )

          msg.attach(MIMEText(mensagem_travada, "plain"))

          part = MIMEBase("application", "octet-stream")
          part.set_payload(excel_anexo_bytes)
          encoders.encode_base64(part)
          part.add_header(
              "Content-Disposition",
              f"attachment; filename={nome_anexo_excel}",
          )
          msg.attach(part)

          servidor_smtp = smtplib.SMTP("smtp-mail.outlook.com", 587)
          servidor_smtp.starttls()
          servidor_smtp.login(remetente_email, senha_email)
          servidor_smtp.sendmail(
              remetente_email, email_destino, msg.as_string()
          )
          servidor_smtp.quit()

          registar_envio_log(
              cdv_escolhido, base_escolhida, periodo_escolhido, email_destino
          )

          st.success(
              f"✨ E-mail enviado com sucesso para **{email_destino}** com o"
              " anexo Excel e registado no histórico de auditoria!"
          )
        except Exception as e:
          st.error(
              f"❌ Erro ao enviar o e-mail via SMTP do Outlook: {str(e)}"
          )

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

    col_vazia1, col_grafico, col_vazia2 = st.columns([1, 3, 1])
    with col_grafico:
      fig, ax = plt.subplots(figsize=(6, 3.8))
      fig.patch.set_facecolor("#112240")
      ax.set_facecolor("#112240")

      bars = ax.bar(
          list(contagem_dados.keys()),
          list(contagem_dados.values()),
          color=["#238636", "#da3633", "#9e6a03", "#1f6feb"],
          width=0.45,
          edgecolor="none",
      )
      ax.tick_params(colors="#c9d1d9", labelsize=10)
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
            fontsize=9,
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

      col_vazia1, col_grafico, col_vazia2 = st.columns([1, 3, 1])
      with col_grafico:
        fig, ax = plt.subplots(figsize=(5.5, 3.8))
        fig.patch.set_facecolor("#112240")
        ax.set_facecolor("#112240")

        bars = ax.bar(
            list(valores_dados.keys()),
            list(valores_dados.values()),
            color=["#238636", "#da3633"],
            width=0.35,
            edgecolor="none",
        )
        ax.tick_params(colors="#c9d1d9", labelsize=10)
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
              fontsize=9,
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
    '<div class="footer">© PepsiCo Brasil - Sistema Multi-CDV | Controle'
    " Operacional em Tempo Real</div>",
    unsafe_allow_html=True,
)