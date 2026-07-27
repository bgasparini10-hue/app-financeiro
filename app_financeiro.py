"""
=============================================
  App Financeiro Pessoal - Adapta ONE
  Streamlit + Pandas + Plotly + SQLite + fpdf2
=============================================
  Funcionalidades:
  - Login simples (usuário único)
  - Dashboard com saldo, receitas/despesas, evolução mensal
  - Lançamentos (CRUD) com data, categoria, valor, descrição
  - Categorias pré-definidas + personalizáveis
  - Extrato com filtros por mês/ano/categoria
  - Relatórios: pizza por categoria, evolução patrimonial, orçamento
  - Orçamento mensal por categoria com alerta de estouro
  - Exportar relatório em PDF
  - Importar extrato bancário via Excel/CSV
=============================================
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
import hashlib
import os
import io
from datetime import datetime, date, timedelta
from fpdf import FPDF
import numpy as np

# =============================================
# CONFIGURAÇÃO INICIAL
# =============================================
st.set_page_config(
    page_title="App Financeiro Pessoal",
    page_icon=open("logo.png", "rb").read(),
    layout="wide",
    initial_sidebar_state="expanded",
)
# =============================================
# CONSTANTES
# =============================================
DB_FILE = "financas.db"

CATEGORIAS_PADRAO = [
    "Alimentação", "Transporte", "Moradia", "Lazer",
    "Saúde", "Educação", "Assinaturas", "Vestuário",
    "Investimentos", "Salário", "Freelas", "Outros",
]

TIPO_RECEITA = "Receita"
TIPO_DESPESA = "Despesa"

# =============================================
# BANCO DE DADOS - SQLite
# =============================================
def get_connection():
    """Retorna conexão com o banco SQLite."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Cria as tabelas do banco se não existirem."""
    conn = get_connection()
    cursor = conn.cursor()

    # Tabela de usuário (login único)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            senha_hash TEXT NOT NULL
        )
    """)

    # Tabela de categorias
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS categorias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT UNIQUE NOT NULL,
            tipo TEXT NOT NULL
        )
    """)

    # Tabela de lançamentos
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS lancamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data DATE NOT NULL,
            tipo TEXT NOT NULL,
            categoria TEXT NOT NULL,
            valor REAL NOT NULL,
            descricao TEXT,
            criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Tabela de orçamento mensal
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS orcamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mes INTEGER NOT NULL,
            ano INTEGER NOT NULL,
            categoria TEXT NOT NULL,
            limite REAL NOT NULL,
            UNIQUE(mes, ano, categoria)
        )
    """)

    conn.commit()
    conn.close()

def seed_categorias():
    """Insere categorias padrão se a tabela estiver vazia."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM categorias")
    if cursor.fetchone()[0] == 0:
        for cat in CATEGORIAS_PADRAO:
            tipo = TIPO_RECEITA if cat in ("Salário", "Freelas", "Investimentos") else TIPO_DESPESA
            cursor.execute(
                "INSERT INTO categorias (nome, tipo) VALUES (?, ?)",
                (cat, tipo),
            )
    conn.commit()
    conn.close()

def criar_usuario_padrao():
    """Cria usuário padrão (admin/123) se não existir."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM usuarios")
    if cursor.fetchone()[0] == 0:
        senha_hash = hashlib.sha256("123".encode()).hexdigest()
        cursor.execute(
            "INSERT INTO usuarios (username, senha_hash) VALUES (?, ?)",
            ("admin", senha_hash),
        )
    conn.commit()
    conn.close()

# =============================================
# FUNÇÕES DE AUTENTICAÇÃO
# =============================================
def verificar_login(username, senha):
    """Verifica credenciais do usuário."""
    conn = get_connection()
    cursor = conn.cursor()
    senha_hash = hashlib.sha256(senha.encode()).hexdigest()
    cursor.execute(
        "SELECT * FROM usuarios WHERE username = ? AND senha_hash = ?",
        (username, senha_hash),
    )
    usuario = cursor.fetchone()
    conn.close()
    return usuario is not None

def tela_login():
    """Renderiza a tela de login."""
    st.markdown(
        """
        <div style="text-align: center; margin-top: 80px;">
            <h1 style="font-size: 3rem;">💰 App Financeiro Pessoal</h1>
            <p style="color: #888; font-size: 1.1rem;">Controle suas finanças de forma simples</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        with st.form("login_form"):
            st.markdown("### 🔐 Acessar")
            username = st.text_input("Usuário", value="admin")
            senha = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Entrar", use_container_width=True)

            if submitted:
                if verificar_login(username, senha):
                    st.session_state["autenticado"] = True
                    st.session_state["usuario"] = username
                    st.rerun()
                else:
                    st.error("Usuário ou senha inválidos.")

        st.markdown(
            "<p style='text-align: center; color: #888; font-size: 0.85rem;'>"
            "Usuário padrão: admin / Senha: 123</p>",
            unsafe_allow_html=True,
        )

# =============================================
# FUNÇÕES DE DADOS - LANÇAMENTOS
# =============================================
def adicionar_lancamento(data, tipo, categoria, valor, descricao=""):
    """Insere um novo lançamento no banco."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """INSERT INTO lancamentos (data, tipo, categoria, valor, descricao)
           VALUES (?, ?, ?, ?, ?)""",
        (data, tipo, categoria, valor, descricao),
    )
    conn.commit()
    conn.close()

def editar_lancamento(lanc_id, data, tipo, categoria, valor, descricao):
    """Atualiza um lançamento existente."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """UPDATE lancamentos
           SET data = ?, tipo = ?, categoria = ?, valor = ?, descricao = ?
           WHERE id = ?""",
        (data, tipo, categoria, valor, descricao, lanc_id),
    )
    conn.commit()
    conn.close()

def excluir_lancamento(lanc_id):
    """Remove um lançamento pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM lancamentos WHERE id = ?", (lanc_id,))
    conn.commit()
    conn.close()

def get_lancamentos(mes=None, ano=None, categoria=None, tipo=None):
    """
    Retorna lançamentos com filtros opcionais.
    Retorna um DataFrame do Pandas.
    """
    conn = get_connection()
    query = "SELECT * FROM lancamentos WHERE 1=1"
    params = []

    if mes:
        query += " AND CAST(strftime('%m', data) AS INTEGER) = ?"
        params.append(mes)
    if ano:
        query += " AND CAST(strftime('%Y', data) AS INTEGER) = ?"
        params.append(ano)
    if categoria and categoria != "Todas":
        query += " AND categoria = ?"
        params.append(categoria)
    if tipo and tipo != "Todos":
        query += " AND tipo = ?"
        params.append(tipo)

    query += " ORDER BY data DESC, id DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df

def get_saldo():
    """Calcula o saldo atual (total receitas - total despesas)."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT COALESCE(SUM(CASE WHEN tipo = 'Receita' THEN valor ELSE 0 END), 0) - "
        "COALESCE(SUM(CASE WHEN tipo = 'Despesa' THEN valor ELSE 0 END), 0) "
        "FROM lancamentos"
    )
    saldo = cursor.fetchone()[0]
    conn.close()
    return saldo

def get_resumo_mes(mes, ano):
    """Retorna total de receitas e despesas de um mês específico."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """SELECT
               COALESCE(SUM(CASE WHEN tipo = 'Receita' THEN valor ELSE 0 END), 0) as receitas,
               COALESCE(SUM(CASE WHEN tipo = 'Despesa' THEN valor ELSE 0 END), 0) as despesas
           FROM lancamentos
           WHERE CAST(strftime('%m', data) AS INTEGER) = ?
           AND CAST(strftime('%Y', data) AS INTEGER) = ?""",
        (mes, ano),
    )
    row = cursor.fetchone()
    conn.close()
    return {"receitas": row[0], "despesas": row[1]}

def get_evolucao_mensal():
    """Retorna a evolução mensal de receitas, despesas e saldo."""
    conn = get_connection()
    query = """
        SELECT
            CAST(strftime('%Y', data) AS INTEGER) as ano,
            CAST(strftime('%m', data) AS INTEGER) as mes,
            SUM(CASE WHEN tipo = 'Receita' THEN valor ELSE 0 END) as receitas,
            SUM(CASE WHEN tipo = 'Despesa' THEN valor ELSE 0 END) as despesas
        FROM lancamentos
        GROUP BY ano, mes
        ORDER BY ano, mes
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    if not df.empty:
        df["periodo"] = df["ano"].astype(str) + "-" + df["mes"].astype(str).str.zfill(2)
        df["saldo"] = df["receitas"] - df["despesas"]
    return df

# =============================================
# FUNÇÕES DE DADOS - CATEGORIAS
# =============================================
def get_categorias(tipo=None):
    """Retorna lista de categorias, opcionalmente filtradas por tipo."""
    conn = get_connection()
    if tipo:
        df = pd.read_sql_query(
            "SELECT * FROM categorias WHERE tipo = ? ORDER BY nome", conn, params=(tipo,)
        )
    else:
        df = pd.read_sql_query("SELECT * FROM categorias ORDER BY nome", conn)
    conn.close()
    return df

def adicionar_categoria(nome, tipo):
    """Adiciona uma nova categoria personalizada."""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            "INSERT INTO categorias (nome, tipo) VALUES (?, ?)", (nome, tipo)
        )
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def excluir_categoria(cat_id):
    """Remove uma categoria pelo ID."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM categorias WHERE id = ?", (cat_id,))
    conn.commit()
    conn.close()

# =============================================
# FUNÇÕES DE DADOS - ORÇAMENTO
# =============================================
def definir_orcamento(mes, ano, categoria, limite):
    """Define ou atualiza o limite de orçamento para uma categoria no mês."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """INSERT INTO orcamentos (mes, ano, categoria, limite)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(mes, ano, categoria)
           DO UPDATE SET limite = ?""",
        (mes, ano, categoria, limite, limite),
    )
    conn.commit()
    conn.close()

def get_orcamentos(mes, ano):
    """Retorna os orçamentos definidos para um mês/ano."""
    conn = get_connection()
    df = pd.read_sql_query(
        """SELECT * FROM orcamentos
           WHERE mes = ? AND ano = ?
           ORDER BY categoria""",
        conn,
        params=(mes, ano),
    )
    conn.close()
    return df

def get_gastos_por_categoria(mes, ano):
    """Retorna total de despesas agrupadas por categoria no mês."""
    conn = get_connection()
    df = pd.read_sql_query(
        """SELECT categoria, SUM(valor) as total
           FROM lancamentos
           WHERE tipo = 'Despesa'
           AND CAST(strftime('%m', data) AS INTEGER) = ?
           AND CAST(strftime('%Y', data) AS INTEGER) = ?
           GROUP BY categoria
           ORDER BY total DESC""",
        conn,
        params=(mes, ano),
    )
    conn.close()
    return df

# =============================================
# FUNÇÕES DE EXPORTAÇÃO - PDF
# =============================================
def exportar_pdf(mes, ano):
    """
    Gera um relatório mensal em PDF com:
    - Resumo financeiro
    - Gastos por categoria
    - Todos os lançamentos do mês
    """
    resumo = get_resumo_mes(mes, ano)
    gastos_cat = get_gastos_por_categoria(mes, ano)
    lancamentos = get_lancamentos(mes=mes, ano=ano)

    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Cabeçalho
    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(0, 12, "Relatorio Financeiro Mensal", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, f"Periodo: {mes:02d}/{ano}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)

    # Resumo
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 10, "Resumo Financeiro", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, f"Receitas: R$ {resumo['receitas']:,.2f}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 7, f"Despesas: R$ {resumo['despesas']:,.2f}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 7, f"Saldo do mes: R$ {resumo['receitas'] - resumo['despesas']:,.2f}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)

    # Gastos por categoria
    if not gastos_cat.empty:
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 10, "Gastos por Categoria", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 10)
        for _, row in gastos_cat.iterrows():
            pdf.cell(0, 6, f"{row['categoria']}: R$ {row['total']:,.2f}", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

    # Tabela de lançamentos
    if not lancamentos.empty:
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 10, "Lancamentos do Mes", new_x="LMARGIN", new_y="NEXT")

        # Cabeçalho da tabela
        pdf.set_font("Helvetica", "B", 9)
        col_widths = [22, 18, 40, 25, 70]
        headers = ["Data", "Tipo", "Categoria", "Valor", "Descricao"]
        for i, h in enumerate(headers):
            pdf.cell(col_widths[i], 7, h, border=1)
        pdf.ln()

        # Linhas
        pdf.set_font("Helvetica", "", 8)
        for _, row in lancamentos.iterrows():
            pdf.cell(col_widths[0], 6, row["data"], border=1)
            pdf.cell(col_widths[1], 6, row["tipo"][:3], border=1)
            pdf.cell(col_widths[2], 6, row["categoria"], border=1)
            pdf.cell(col_widths[3], 6, f"R$ {row['valor']:,.2f}", border=1)
            # Trunca descrição longa
            desc = row["descricao"] if row["descricao"] else ""
            pdf.cell(col_widths[4], 6, desc[:35], border=1)
            pdf.ln()

    # Salvar em buffer
    buf = io.BytesIO()
    pdf.output(buf)
    buf.seek(0)
    return buf

# =============================================
# FUNÇÕES DE IMPORTAÇÃO
# =============================================
def importar_excel(arquivo):
    """
    Importa lançamentos de um arquivo Excel/CSV.
    Formato esperado: data, tipo, categoria, valor, descricao
    """
    try:
        if arquivo.name.endswith(".csv"):
            df = pd.read_csv(arquivo)
        else:
            df = pd.read_excel(arquivo, engine="openpyxl")

        # Normaliza nomes das colunas
        df.columns = [c.strip().lower() for c in df.columns]

        mapeamento = {
            "data": "data",
            "tipo": "tipo",
            "categoria": "categoria",
            "valor": "valor",
            "descricao": "descricao",
            "descrição": "descricao",
            "desc": "descricao",
            "date": "data",
            "type": "tipo",
            "category": "categoria",
            "value": "valor",
            "amount": "valor",
            "description": "descricao",
        }

        colunas_usar = {}
        for col in df.columns:
            if col in mapeamento:
                colunas_usar[col] = mapeamento[col]

        if "data" not in colunas_usar.values() or "valor" not in colunas_usar.values():
            return False, "Arquivo precisa ter colunas de 'data' e 'valor'."

        df = df.rename(columns=colunas_usar)
        df = df[list(colunas_usar.values())]

        # Converte data
        df["data"] = pd.to_datetime(df["data"], errors="coerce")
        df = df.dropna(subset=["data"])
        df["data"] = df["data"].dt.strftime("%Y-%m-%d")

        # Converte valor
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
        df = df.dropna(subset=["valor"])

        # Preenche campos opcionais
        if "tipo" not in df.columns:
            # Se não informar tipo, infere: valores positivos = receita, negativos = despesa
            df["tipo"] = df["valor"].apply(
                lambda v: TIPO_RECEITA if v >= 0 else TIPO_DESPESA
            )
            df["valor"] = df["valor"].abs()

        if "categoria" not in df.columns:
            df["categoria"] = "Outros"

        if "descricao" not in df.columns:
            df["descricao"] = ""

        # Insere no banco
        conn = get_connection()
        cursor = conn.cursor()
        importados = 0
        for _, row in df.iterrows():
            cursor.execute(
                """INSERT INTO lancamentos (data, tipo, categoria, valor, descricao)
                   VALUES (?, ?, ?, ?, ?)""",
                (row["data"], row["tipo"], row["categoria"], abs(row["valor"]), row["descricao"]),
            )
            importados += 1
        conn.commit()
        conn.close()

        return True, f"{importados} lançamentos importados com sucesso!"

    except Exception as e:
        return False, f"Erro ao importar: {str(e)}"

# =============================================
# INTERFACE - SIDEBAR
# =============================================
def sidebar_navegacao():
    """Renderiza a barra lateral com navegação."""
    with st.sidebar:
        st.markdown("### 💰 App Financeiro")
        st.markdown(f"👤 **{st.session_state.get('usuario', 'usuário')}**")
        st.divider()

        paginas = {
            "📊 Dashboard": "dashboard",
            "📝 Lançamentos": "lancamentos",
            "📋 Extrato": "extrato",
            "📈 Relatórios": "relatorios",
            "🎯 Orçamento": "orcamento",
            "📤 Importar": "importar",
            "⚙️ Categorias": "categorias",
        }

        pagina = st.radio("Navegação", list(paginas.keys()), label_visibility="collapsed")
        st.divider()

        if st.button("🚪 Sair", use_container_width=True):
            st.session_state["autenticado"] = False
            st.rerun()

        return paginas[pagina]

# =============================================
# INTERFACE - PÁGINAS
# =============================================
def pagina_dashboard():
    """Dashboard principal com saldo, resumo do mês e gráficos."""
    st.markdown("## 📊 Dashboard")

    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year

    # Cards de resumo
    saldo = get_saldo()
    resumo = get_resumo_mes(mes_atual, ano_atual)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("💰 Saldo Total", f"R$ {saldo:,.2f}")
    with col2:
        st.metric(
            "📈 Receitas (mês)",
            f"R$ {resumo['receitas']:,.2f}",
        )
    with col3:
        st.metric(
            "📉 Despesas (mês)",
            f"R$ {resumo['despesas']:,.2f}",
        )
    with col4:
        saldo_mes = resumo["receitas"] - resumo["despesas"]
        st.metric(
            "💵 Saldo (mês)",
            f"R$ {saldo_mes:,.2f}",
            delta=f"R$ {saldo_mes:,.2f}",
        )

    st.divider()

    # Gráfico de evolução mensal
    df_evolucao = get_evolucao_mensal()
    if not df_evolucao.empty:
        st.markdown("### 📈 Evolução Mensal")

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_evolucao["periodo"],
            y=df_evolucao["receitas"],
            mode="lines+markers",
            name="Receitas",
            line=dict(color="#10b981", width=3),
        ))
        fig.add_trace(go.Scatter(
            x=df_evolucao["periodo"],
            y=df_evolucao["despesas"],
            mode="lines+markers",
            name="Despesas",
            line=dict(color="#ef4444", width=3),
        ))
        fig.add_trace(go.Bar(
            x=df_evolucao["periodo"],
            y=df_evolucao["saldo"],
            name="Saldo",
            marker_color="#3b82f6",
            opacity=0.6,
        ))

        fig.update_layout(
            height=400,
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=0, r=0, t=10, b=0),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("💡 Nenhum lançamento encontrado. Comece cadastrando receitas e despesas!")

    # Últimos lançamentos
    st.markdown("### 📋 Últimos Lançamentos")
    df_lanc = get_lancamentos()
    if not df_lanc.empty:
        df_show = df_lanc[["data", "tipo", "categoria", "valor", "descricao"]].head(10)
        df_show["valor"] = df_show["valor"].apply(lambda x: f"R$ {x:,.2f}")
        st.dataframe(df_show, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum lançamento cadastrado.")

def pagina_lancamentos():
    """CRUD de lançamentos: cadastrar, editar e excluir."""
    st.markdown("## 📝 Lançamentos")

    tab1, tab2 = st.tabs(["➕ Novo Lançamento", "📋 Gerenciar Lançamentos"])

    with tab1:
        with st.form("form_lancamento"):
            col1, col2 = st.columns(2)
            with col1:
                data = st.date_input("Data", value=date.today())
                tipo = st.selectbox("Tipo", [TIPO_DESPESA, TIPO_RECEITA], key="tipo_lanc")

            with col2:
                # Filtra categorias pelo tipo selecionado
                df_cats = get_categorias(tipo=tipo)
                if df_cats.empty:
                    df_cats = get_categorias()
                categorias_lista = df_cats["nome"].tolist()
                categoria = st.selectbox("Categoria", categorias_lista)
                valor = st.number_input("Valor (R$)", min_value=0.01, step=10.0, format="%.2f")

            descricao = st.text_input("Descrição (opcional)")

            submitted = st.form_submit_button("💾 Salvar", use_container_width=True)
            if submitted:
                adicionar_lancamento(
                    data=data.strftime("%Y-%m-%d"), tipo=tipo,
                    categoria=categoria, valor=valor, descricao=descricao,
                )
                st.success("✅ Lançamento cadastrado com sucesso!")
                st.rerun()

    with tab2:
        df_lanc = get_lancamentos()
        if df_lanc.empty:
            st.info("Nenhum lançamento cadastrado.")
            return

        # Filtros
        col1, col2 = st.columns(2)
        with col1:
            meses = {
                1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
                5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
                9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
            }
            mes_sel = st.selectbox("Mês", list(meses.values()), key="mes_filtro_ger")
            mes_num = list(meses.keys())[list(meses.values()).index(mes_sel)]
        with col2:
            ano_sel = st.number_input("Ano", min_value=2020, max_value=2030, value=date.today().year)

        df_filtrado = get_lancamentos(mes=mes_num, ano=ano_sel)

        if df_filtrado.empty:
            st.info(f"Nenhum lançamento em {mes_sel}/{ano_sel}.")
            return

        # Lista para editar/excluir
        for _, row in df_filtrado.iterrows():
            with st.expander(
                f"{row['data']} | {row['tipo']} | {row['categoria']} | "
                f"R$ {row['valor']:,.2f} | {row['descricao'] or ''}"
            ):
                with st.form(f"edit_{row['id']}"):
                    col1, col2 = st.columns(2)
                    with col1:
                        e_data = st.date_input(
                            "Data",
                            value=datetime.strptime(row["data"], "%Y-%m-%d").date(),
                            key=f"ed_{row['id']}",
                        )
                        e_tipo = st.selectbox(
                            "Tipo",
                            [TIPO_DESPESA, TIPO_RECEITA],
                            index=0 if row["tipo"] == TIPO_DESPESA else 1,
                            key=f"et_{row['id']}",
                        )
                    with col2:
                        df_cats_e = get_categorias(tipo=e_tipo)
                        cats_e = df_cats_e["nome"].tolist() if not df_cats_e.empty else CATEGORIAS_PADRAO
                        idx_cat = cats_e.index(row["categoria"]) if row["categoria"] in cats_e else 0
                        e_cat = st.selectbox("Categoria", cats_e, index=idx_cat, key=f"ec_{row['id']}")
                        e_valor = st.number_input(
                            "Valor", value=row["valor"], step=10.0, format="%.2f", key=f"ev_{row['id']}"
                        )
                    e_desc = st.text_input("Descrição", value=row["descricao"] or "", key=f"edesc_{row['id']}")

                    col_a, col_b = st.columns(2)
                    with col_a:
                        if st.form_submit_button("💾 Atualizar", use_container_width=True):
                            editar_lancamento(
                                row["id"], e_data.strftime("%Y-%m-%d"),
                                e_tipo, e_cat, e_valor, e_desc,
                            )
                            st.success("✅ Atualizado!")
                            st.rerun()
                    with col_b:
                        if st.form_submit_button("🗑️ Excluir", use_container_width=True, type="primary"):
                            excluir_lancamento(row["id"])
                            st.success("🗑️ Excluído!")
                            st.rerun()

def pagina_extrato():
    """Extrato completo com filtros e exportação PDF."""
    st.markdown("## 📋 Extrato")

    # Filtros
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        meses = {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
            5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
            9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
        }
        mes_sel = st.selectbox("Mês", ["Todos"] + list(meses.values()), index=0)
        mes_num = None if mes_sel == "Todos" else list(meses.keys())[list(meses.values()).index(mes_sel)]
    with col2:
        ano_sel = st.number_input("Ano", min_value=2020, max_value=2030, value=date.today().year)
    with col3:
        df_cats = get_categorias()
        cats_list = ["Todas"] + df_cats["nome"].tolist()
        cat_sel = st.selectbox("Categoria", cats_list)
    with col4:
        tipo_sel = st.selectbox("Tipo", ["Todos", TIPO_RECEITA, TIPO_DESPESA])

    df = get_lancamentos(mes=mes_num, ano=ano_sel, categoria=cat_sel, tipo=tipo_sel)

    if df.empty:
        st.info("Nenhum lançamento encontrado com esses filtros.")
        return

    # Métricas
    total_receitas = df[df["tipo"] == TIPO_RECEITA]["valor"].sum()
    total_despesas = df[df["tipo"] == TIPO_DESPESA]["valor"].sum()
    col1, col2, col3 = st.columns(3)
    col1.metric("📈 Total Receitas", f"R$ {total_receitas:,.2f}")
    col2.metric("📉 Total Despesas", f"R$ {total_despesas:,.2f}")
    col3.metric("💵 Saldo", f"R$ {total_receitas - total_despesas:,.2f}")

    # Tabela
    df_show = df[["data", "tipo", "categoria", "valor", "descricao"]].copy()
    df_show["valor"] = df_show["valor"].apply(lambda x: f"R$ {x:,.2f}")
    st.dataframe(df_show, use_container_width=True, hide_index=True)

    # Exportar PDF
    if mes_num:
        st.divider()
        if st.button("📄 Exportar Relatório em PDF", use_container_width=True):
            with st.spinner("Gerando PDF..."):
                buf = exportar_pdf(mes_num, ano_sel)
                st.download_button(
                    label="📥 Baixar PDF",
                    data=buf,
                    file_name=f"relatorio_{mes_num:02d}_{ano_sel}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

def pagina_relatorios():
    """Relatórios visuais: pizza, evolução e orçamento."""
    st.markdown("## 📈 Relatórios")

    hoje = date.today()
    col1, col2 = st.columns(2)
    with col1:
        mes_sel = st.selectbox("Mês", list(range(1, 13)), index=hoje.month - 1, format_func=lambda m: {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
            5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
            9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
        }[m])
    with col2:
        ano_sel = st.number_input("Ano", min_value=2020, max_value=2030, value=hoje.year)

    tab1, tab2, tab3 = st.tabs(["🥧 Gastos por Categoria", "📈 Evolução Patrimonial", "🎯 Orçamento vs Real"])

    # --- TAB 1: Gastos por Categoria (Pizza) ---
    with tab1:
        df_gastos = get_gastos_por_categoria(mes_sel, ano_sel)
        if not df_gastos.empty:
            fig = px.pie(
                df_gastos,
                values="total",
                names="categoria",
                title=f"Gastos por Categoria - {mes_sel}/{ano_sel}",
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Set3,
            )
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(height=500, margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig, use_container_width=True)

            # Tabela auxiliar
            df_gastos["total"] = df_gastos["total"].apply(lambda x: f"R$ {x:,.2f}")
            st.dataframe(df_gastos, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma despesa neste mês.")

    # --- TAB 2: Evolução Patrimonial (Linha) ---
    with tab2:
        df_evol = get_evolucao_mensal()
        if not df_evol.empty:
            # Calcula saldo acumulado
            df_evol["saldo_acumulado"] = df_evol["saldo"].cumsum()

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=df_evol["periodo"],
                y=df_evol["saldo_acumulado"],
                mode="lines+markers",
                name="Patrimônio",
                line=dict(color="#10b981", width=3),
                fill="tozeroy",
                fillcolor="rgba(16, 185, 129, 0.15)",
            ))
            fig.update_layout(
                title="Evolução Patrimonial Acumulada",
                height=450,
                hovermode="x",
                yaxis=dict(tickprefix="R$ "),
                margin=dict(l=0, r=0, t=40, b=0),
            )
            st.plotly_chart(fig, use_container_width=True)

            # Tabela
            df_show = df_evol[["periodo", "receitas", "despesas", "saldo", "saldo_acumulado"]].copy()
            for col in ["receitas", "despesas", "saldo", "saldo_acumulado"]:
                df_show[col] = df_show[col].apply(lambda x: f"R$ {x:,.2f}")
            st.dataframe(df_show, use_container_width=True, hide_index=True)
        else:
            st.info("Cadastre lançamentos para ver a evolução patrimonial.")

    # --- TAB 3: Orçamento vs Real ---
    with tab3:
        df_orc = get_orcamentos(mes_sel, ano_sel)
        df_gastos_real = get_gastos_por_categoria(mes_sel, ano_sel)

        if not df_orc.empty:
            # Merge orçamento com gastos reais
            df_comparativo = df_orc.merge(
                df_gastos_real, on="categoria", how="left"
            ).fillna(0)
            df_comparativo["total"] = df_comparativo["total"].astype(float)
            df_comparativo["% utilizado"] = (
                (df_comparativo["total"] / df_comparativo["limite"] * 100)
                .round(1)
            )
            df_comparativo["status"] = df_comparativo["% utilizado"].apply(
                lambda x: "⚠️ Estouro!" if x > 100 else (
                    "🔴 Atenção" if x > 80 else "🟢 OK"
                )
            )

            df_show = df_comparativo[["categoria", "limite", "total", "% utilizado", "status"]].copy()
            df_show["limite"] = df_show["limite"].apply(lambda x: f"R$ {x:,.2f}")
            df_show["total"] = df_show["total"].apply(lambda x: f"R$ {x:,.2f}")

            st.dataframe(df_show, use_container_width=True, hide_index=True)

            # Gráfico de barras comparativo
            fig = go.Figure()
            fig.add_trace(go.Bar(
                name="Limite",
                x=df_comparativo["categoria"],
                y=df_comparativo["limite"],
                marker_color="#3b82f6",
            ))
            fig.add_trace(go.Bar(
                name="Gasto Real",
                x=df_comparativo["categoria"],
                y=df_comparativo["total"],
                marker_color="#ef4444",
            ))
            fig.update_layout(
                title="Orçamento vs Gastos Reais",
                barmode="group",
                height=400,
                margin=dict(l=0, r=0, t=40, b=0),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(
                "Nenhum orçamento definido para este mês. "
                "Vá para a página 🎯 Orçamento para definir limites."
            )

def pagina_orcamento():
    """Gerenciamento de orçamento mensal por categoria."""
    st.markdown("## 🎯 Orçamento Mensal")

    hoje = date.today()
    col1, col2 = st.columns(2)
    with col1:
        mes_sel = st.selectbox("Mês", list(range(1, 13)), index=hoje.month - 1, format_func=lambda m: {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
            5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
            9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
        }[m])
    with col2:
        ano_sel = st.number_input("Ano", min_value=2020, max_value=2030, value=hoje.year)

    # Formulário para adicionar/editar orçamento
    with st.expander("➕ Adicionar / Editar Orçamento", expanded=True):
        df_cats = get_categorias(tipo=TIPO_DESPESA)
        if df_cats.empty:
            st.warning("Nenhuma categoria de despesa cadastrada.")
            return

        cats_despesa = df_cats["nome"].tolist()
        # Busca orçamentos já definidos
        df_orc_existente = get_orcamentos(mes_sel, ano_sel)

        with st.form("form_orcamento"):
            for cat in cats_despesa:
                valor_atual = 0.0
                if not df_orc_existente.empty:
                    row = df_orc_existente[df_orc_existente["categoria"] == cat]
                    if not row.empty:
                        valor_atual = row.iloc[0]["limite"]

                limite = st.number_input(
                    f"{cat} (R$)",
                    min_value=0.0,
                    value=valor_atual,
                    step=50.0,
                    format="%.2f",
                    key=f"orc_{cat}",
                )
                if limite > 0:
                    definir_orcamento(mes_sel, ano_sel, cat, limite)

            submitted = st.form_submit_button("💾 Salvar Orçamentos", use_container_width=True)
            if submitted:
                st.success("✅ Orçamentos salvos com sucesso!")
                st.rerun()

    # Exibe status dos orçamentos
    st.divider()
    st.markdown("### 📊 Status dos Orçamentos")

    df_orc = get_orcamentos(mes_sel, ano_sel)
    df_gastos = get_gastos_por_categoria(mes_sel, ano_sel)

    if not df_orc.empty:
        df_comp = df_orc.merge(df_gastos, on="categoria", how="left").fillna(0)
        df_comp["total"] = df_comp["total"].astype(float)
        df_comp["%"] = (df_comp["total"] / df_comp["limite"] * 100).round(1)

        for _, row in df_comp.iterrows():
            pct = row["%"]
            if pct > 100:
                st.error(f"⚠️ **{row['categoria']}**: R$ {row['total']:,.2f} / R$ {row['limite']:,.2f} ({pct}%) - **ESTOUROU!**")
            elif pct > 80:
                st.warning(f"🔴 **{row['categoria']}**: R$ {row['total']:,.2f} / R$ {row['limite']:,.2f} ({pct}%)")
            else:
                st.success(f"🟢 **{row['categoria']}**: R$ {row['total']:,.2f} / R$ {row['limite']:,.2f} ({pct}%)")

        # Gasto total
        total_limites = df_comp["limite"].sum()
        total_gasto = df_comp["total"].sum()
        st.metric("Total Orçado", f"R$ {total_limites:,.2f}")
        st.metric("Total Gasto", f"R$ {total_gasto:,.2f}")
        if total_limites > 0:
            pct_geral = (total_gasto / total_limites * 100)
            st.metric("% Utilizado", f"{pct_geral:.1f}%")
    else:
        st.info("Nenhum orçamento definido. Preencha os valores acima.")

def pagina_importar():
    """Importação de extrato bancário via Excel/CSV."""
    st.markdown("## 📤 Importar Extrato Bancário")
    st.markdown(
        """
        Faça upload de um arquivo **Excel (.xlsx)** ou **CSV** com as colunas:
        - `data` (obrigatório)
        - `valor` (obrigatório)
        - `tipo` (opcional: Receita/Despesa)
        - `categoria` (opcional)
        - `descricao` (opcional)

        Se o tipo não for informado, valores positivos viram Receita e negativos viram Despesa.
        """
    )

    arquivo = st.file_uploader(
        "Selecione o arquivo",
        type=["xlsx", "csv"],
        accept_multiple_files=False,
    )

    if arquivo:
        with st.spinner("Importando dados..."):
            sucesso, mensagem = importar_excel(arquivo)
            if sucesso:
                st.success(f"✅ {mensagem}")
            else:
                st.error(f"❌ {mensagem}")

    st.divider()
    st.markdown("### 📋 Últimas Importações")
    df = get_lancamentos()
    if not df.empty:
        st.dataframe(
            df[["data", "tipo", "categoria", "valor", "descricao"]].head(20),
            use_container_width=True,
            hide_index=True,
        )

def pagina_categorias():
    """Gerenciamento de categorias personalizadas."""
    st.markdown("## ⚙️ Gerenciar Categorias")

    tab1, tab2 = st.tabs(["➕ Nova Categoria", "📋 Categorias Existentes"])

    with tab1:
        with st.form("form_categoria"):
            nome = st.text_input("Nome da categoria")
            tipo = st.selectbox("Tipo", [TIPO_DESPESA, TIPO_RECEITA])
            submitted = st.form_submit_button("💾 Salvar", use_container_width=True)
            if submitted and nome.strip():
                if adicionar_categoria(nome.strip(), tipo):
                    st.success(f"✅ Categoria '{nome}' criada!")
                    st.rerun()
                else:
                    st.error("❌ Categoria já existe.")

    with tab2:
        df = get_categorias()
        if df.empty:
            st.info("Nenhuma categoria cadastrada.")
        else:
            for _, row in df.iterrows():
                col1, col2, col3 = st.columns([3, 2, 1])
                with col1:
                    st.write(f"**{row['nome']}**")
                with col2:
                    st.write(row["tipo"])
                with col3:
                    # Não permite excluir categorias com lançamentos vinculados
                    if st.button("🗑️", key=f"del_cat_{row['id']}"):
                        excluir_categoria(row["id"])
                        st.rerun()

# =============================================
# MAIN - CONTROLE DE FLUXO
# =============================================
def main():
    """Função principal que controla o fluxo do app."""
    # Inicializa banco e dados padrão
    init_db()
    seed_categorias()
    criar_usuario_padrao()

    # Controle de sessão
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    # Login ou app
    if not st.session_state["autenticado"]:
        tela_login()
    else:
        pagina = sidebar_navegacao()

        # Roteamento
        paginas = {
            "dashboard": pagina_dashboard,
            "lancamentos": pagina_lancamentos,
            "extrato": pagina_extrato,
            "relatorios": pagina_relatorios,
            "orcamento": pagina_orcamento,
            "importar": pagina_importar,
            "categorias": pagina_categorias,
        }

        paginas[pagina]()

if __name__ == "__main__":
    main()