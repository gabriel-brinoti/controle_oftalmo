from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import os
from dotenv import load_dotenv
import psycopg2
from psycopg2.extras import RealDictCursor
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, date, timedelta
from functools import wraps
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image


load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "chave_local_dev_troque_em_producao")
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("FLASK_ENV") == "production"

DATABASE_URL = os.environ.get("DATABASE_URL")
ALERTA_DIAS = 3


def conectar():
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL não configurada. Configure no .env local ou nas variáveis do Render.")
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


def criar_banco():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS admin (
            id SERIAL PRIMARY KEY,
            usuario TEXT NOT NULL UNIQUE,
            senha_hash TEXT NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS categorias (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL UNIQUE,
            descricao TEXT,
            criado_em TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS produtos (
            id SERIAL PRIMARY KEY,
            nome TEXT NOT NULL,
            categoria_id INTEGER NOT NULL REFERENCES categorias(id),
            lote TEXT,
            data_vencimento DATE NOT NULL,
            data_abertura DATE,
            validade_apos_aberto_dias INTEGER,
            quantidade_atual INTEGER NOT NULL DEFAULT 0,
            estoque_padrao INTEGER NOT NULL DEFAULT 0,
            limite_alerta INTEGER NOT NULL DEFAULT 0,
            observacoes TEXT,
            criado_em TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS movimentacoes (
            id SERIAL PRIMARY KEY,
            produto_id INTEGER,
            produto_nome TEXT NOT NULL,
            tipo_movimentacao TEXT NOT NULL,
            quantidade INTEGER NOT NULL,
            estoque_anterior INTEGER NOT NULL,
            estoque_atual INTEGER NOT NULL,
            observacao TEXT,
            data_movimentacao TIMESTAMP NOT NULL DEFAULT NOW()
        )
    """)

    cursor.execute("SELECT id FROM admin WHERE usuario = %s", ("admin",))
    admin = cursor.fetchone()

    if not admin:
        cursor.execute(
            "INSERT INTO admin (usuario, senha_hash) VALUES (%s, %s)",
            ("admin", generate_password_hash("admin123"))
        )

    conn.commit()
    conn.close()


def login_obrigatorio(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not session.get("admin_logado"):
            return redirect(url_for("login"))
        return func(*args, **kwargs)
    return wrapper


def converter_data(valor):
    if not valor:
        return None
    if isinstance(valor, date):
        return valor
    return datetime.strptime(str(valor), "%Y-%m-%d").date()


def formatar_data(valor):
    if not valor:
        return "-"
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y %H:%M")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    return str(valor)


def calcular_status(produto):
    hoje = date.today()
    alertas = []

    data_vencimento = converter_data(produto["data_vencimento"])

    if data_vencimento:
        dias = (data_vencimento - hoje).days

        if dias < 0:
            alertas.append({"tipo": "danger", "prioridade": 2, "texto": "Vencido pela validade original"})
        elif dias <= ALERTA_DIAS:
            alertas.append({"tipo": "warning", "prioridade": 4, "texto": f"Vence em {dias} dia(s)"})

    data_abertura = converter_data(produto["data_abertura"])
    validade_dias = produto["validade_apos_aberto_dias"]
    vencimento_apos_aberto = None

    if data_abertura and validade_dias:
        vencimento_apos_aberto = data_abertura + timedelta(days=int(validade_dias))
        dias_aberto = (vencimento_apos_aberto - hoje).days

        if dias_aberto < 0:
            alertas.append({"tipo": "danger", "prioridade": 1, "texto": "Vencido após abertura"})
        elif dias_aberto <= ALERTA_DIAS:
            alertas.append({"tipo": "warning", "prioridade": 3, "texto": f"Vence após aberto em {dias_aberto} dia(s)"})

    quantidade = produto["quantidade_atual"]
    limite = produto["limite_alerta"]

    if quantidade == 0:
        alertas.append({"tipo": "danger", "prioridade": 5, "texto": "Estoque zerado"})
    elif quantidade <= limite:
        alertas.append({"tipo": "stock", "prioridade": 6, "texto": "Estoque baixo"})

    if not alertas:
        return {"tipo": "success", "texto": "Produto OK", "vencimento_apos_aberto": vencimento_apos_aberto}

    principal = sorted(alertas, key=lambda x: x["prioridade"])[0]
    return {"tipo": principal["tipo"], "texto": principal["texto"], "vencimento_apos_aberto": vencimento_apos_aberto}


@app.route("/")
def index():
    if session.get("admin_logado"):
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip()
        senha = request.form.get("senha", "").strip()

        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM admin WHERE usuario = %s", (usuario,))
        admin = cursor.fetchone()
        conn.close()

        if admin and check_password_hash(admin["senha_hash"], senha):
            session["admin_logado"] = True
            session["usuario"] = admin["usuario"]
            return redirect(url_for("dashboard"))

        flash("Usuário ou senha inválidos.", "erro")

    return render_template("login.html")


@app.route("/logout")
@login_obrigatorio
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_obrigatorio
def dashboard():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT produtos.*, categorias.nome AS categoria_nome
        FROM produtos
        JOIN categorias ON categorias.id = produtos.categoria_id
        ORDER BY produtos.nome
    """)
    produtos = cursor.fetchall()

    cursor.execute("""
        SELECT *
        FROM movimentacoes
        WHERE data_movimentacao >= NOW() - INTERVAL '30 days'
        ORDER BY data_movimentacao DESC
    """)
    movimentacoes_30 = cursor.fetchall()

    cursor.execute("""
        SELECT produto_nome, SUM(quantidade) AS total_saida
        FROM movimentacoes
        WHERE tipo_movimentacao = 'saida'
          AND data_movimentacao >= NOW() - INTERVAL '30 days'
        GROUP BY produto_nome
        ORDER BY total_saida DESC
        LIMIT 5
    """)
    consumo_top = cursor.fetchall()

    cursor.execute("""
        SELECT 
            COALESCE(SUM(CASE WHEN tipo_movimentacao = 'entrada' THEN quantidade ELSE 0 END), 0) AS entradas,
            COALESCE(SUM(CASE WHEN tipo_movimentacao = 'saida' THEN quantidade ELSE 0 END), 0) AS saidas
        FROM movimentacoes
        WHERE data_movimentacao >= NOW() - INTERVAL '30 days'
    """)
    resumo_mov = cursor.fetchone()

    conn.close()

    produtos_status = []
    contadores = {
        "total": len(produtos),
        "vencidos": 0,
        "proximos": 0,
        "vencidos_abertos": 0,
        "estoque_baixo": 0,
        "estoque_zerado": 0,
        "entradas_mes": resumo_mov["entradas"] or 0,
        "saidas_mes": resumo_mov["saidas"] or 0,
        "movimentacoes_mes": len(movimentacoes_30)
    }

    estoque_ok = 0
    estoque_baixo = 0
    estoque_zerado = 0
    alertas_inteligentes = []
    previsoes = []

    consumo_por_nome = {item["produto_nome"]: item["total_saida"] or 0 for item in consumo_top}

    for produto in produtos:
        status = calcular_status(produto)
        produtos_status.append({"produto": produto, "status": status})
        texto = status["texto"].lower()

        if "vencido após abertura" in texto:
            contadores["vencidos_abertos"] += 1
            alertas_inteligentes.append(f"{produto['nome']} está vencido após abertura.")
        elif "vencido" in texto:
            contadores["vencidos"] += 1
            alertas_inteligentes.append(f"{produto['nome']} está vencido pela validade original.")

        if "vence" in texto:
            contadores["proximos"] += 1
            alertas_inteligentes.append(f"{produto['nome']} precisa de atenção: {status['texto']}.")

        if produto["quantidade_atual"] == 0:
            contadores["estoque_zerado"] += 1
            estoque_zerado += 1
            alertas_inteligentes.append(f"{produto['nome']} está com estoque zerado.")
        elif produto["quantidade_atual"] <= produto["limite_alerta"]:
            contadores["estoque_baixo"] += 1
            estoque_baixo += 1
            alertas_inteligentes.append(f"{produto['nome']} está abaixo do limite de estoque.")
        else:
            estoque_ok += 1

        saidas_30 = consumo_por_nome.get(produto["nome"], 0)
        if saidas_30 > 0:
            media_diaria = float(saidas_30) / 30
            dias_estimados = int(produto["quantidade_atual"] / media_diaria) if produto["quantidade_atual"] > 0 else 0
            previsoes.append({
                "produto": produto["nome"],
                "dias": dias_estimados,
                "estoque": produto["quantidade_atual"],
                "media": round(media_diaria, 2)
            })

    grafico_movimentacoes = {
        "labels": ["Entradas", "Saídas"],
        "dados": [int(contadores["entradas_mes"]), int(contadores["saidas_mes"])]
    }
    grafico_estoque = {"labels": ["OK", "Baixo", "Zerado"], "dados": [estoque_ok, estoque_baixo, estoque_zerado]}
    grafico_consumo = {
        "labels": [item["produto_nome"] for item in consumo_top],
        "dados": [int(item["total_saida"] or 0) for item in consumo_top]
    }

    ranking_consumo = [{"produto": item["produto_nome"], "total": int(item["total_saida"] or 0)} for item in consumo_top]
    produtos_criticos = [item for item in produtos_status if item["status"]["tipo"] != "success"][:8]
    previsoes = sorted(previsoes, key=lambda x: x["dias"])[:5]
    alertas_inteligentes = alertas_inteligentes[:8]

    return render_template(
        "dashboard.html",
        contadores=contadores,
        produtos_status=produtos_status,
        produtos_criticos=produtos_criticos,
        ranking_consumo=ranking_consumo,
        previsoes=previsoes,
        alertas_inteligentes=alertas_inteligentes,
        grafico_movimentacoes=grafico_movimentacoes,
        grafico_estoque=grafico_estoque,
        grafico_consumo=grafico_consumo
    )


@app.route("/categorias")
@login_obrigatorio
def categorias():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM categorias ORDER BY nome")
    categorias = cursor.fetchall()
    conn.close()
    return render_template("categorias.html", categorias=categorias)


@app.route("/categorias/nova", methods=["GET", "POST"])
@login_obrigatorio
def nova_categoria():
    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        descricao = request.form.get("descricao", "").strip()

        if not nome:
            flash("O nome da categoria é obrigatório.", "erro")
            return render_template("nova_categoria.html")

        conn = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO categorias (nome, descricao) VALUES (%s, %s)", (nome, descricao))
            conn.commit()
            flash("Categoria cadastrada com sucesso.", "sucesso")
            return redirect(url_for("categorias"))
        except psycopg2.IntegrityError:
            conn.rollback()
            flash("Essa categoria já existe.", "erro")
        finally:
            conn.close()

    return render_template("nova_categoria.html")


@app.route("/categorias/editar/<int:id>", methods=["GET", "POST"])
@login_obrigatorio
def editar_categoria(id):
    conn = conectar()
    cursor = conn.cursor()

    if request.method == "POST":
        nome = request.form.get("nome", "").strip()
        descricao = request.form.get("descricao", "").strip()

        if not nome:
            flash("O nome da categoria é obrigatório.", "erro")
        else:
            try:
                cursor.execute("UPDATE categorias SET nome = %s, descricao = %s WHERE id = %s", (nome, descricao, id))
                conn.commit()
                flash("Categoria atualizada com sucesso.", "sucesso")
                conn.close()
                return redirect(url_for("categorias"))
            except psycopg2.IntegrityError:
                conn.rollback()
                flash("Essa categoria já existe.", "erro")

    cursor.execute("SELECT * FROM categorias WHERE id = %s", (id,))
    categoria = cursor.fetchone()
    conn.close()

    if not categoria:
        flash("Categoria não encontrada.", "erro")
        return redirect(url_for("categorias"))

    return render_template("editar_categoria.html", categoria=categoria)


@app.route("/categorias/excluir/<int:id>", methods=["POST"])
@login_obrigatorio
def excluir_categoria(id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) AS total FROM produtos WHERE categoria_id = %s", (id,))
    total = cursor.fetchone()["total"]

    if total > 0:
        flash("Não é possível excluir uma categoria com produtos vinculados.", "erro")
    else:
        cursor.execute("DELETE FROM categorias WHERE id = %s", (id,))
        conn.commit()
        flash("Categoria excluída com sucesso.", "sucesso")

    conn.close()
    return redirect(url_for("categorias"))


def listar_categorias():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM categorias ORDER BY nome")
    categorias = cursor.fetchall()
    conn.close()
    return categorias


def registrar_movimentacao(cursor, produto_id, produto_nome, tipo_movimentacao, quantidade, estoque_anterior, estoque_atual, observacao):
    cursor.execute("""
        INSERT INTO movimentacoes (
            produto_id, produto_nome, tipo_movimentacao, quantidade,
            estoque_anterior, estoque_atual, observacao
        ) VALUES (%s, %s, %s, %s, %s, %s, %s)
    """, (produto_id, produto_nome, tipo_movimentacao, quantidade, estoque_anterior, estoque_atual, observacao))


@app.route("/produtos")
@login_obrigatorio
def produtos():
    filtro = request.args.get("filtro", "todos")
    busca = request.args.get("busca", "").strip()
    categoria_id = request.args.get("categoria_id", "").strip()

    conn = conectar()
    cursor = conn.cursor()

    query = """
        SELECT produtos.*, categorias.nome AS categoria_nome
        FROM produtos
        JOIN categorias ON categorias.id = produtos.categoria_id
        WHERE 1=1
    """
    params = []

    if busca:
        query += " AND (produtos.nome ILIKE %s OR produtos.lote ILIKE %s)"
        params.extend([f"%{busca}%", f"%{busca}%"])

    if categoria_id:
        query += " AND produtos.categoria_id = %s"
        params.append(categoria_id)

    query += " ORDER BY produtos.nome"

    cursor.execute(query, params)
    produtos_lista = cursor.fetchall()
    conn.close()

    categorias_lista = listar_categorias()
    produtos_status = []

    for produto in produtos_lista:
        status = calcular_status(produto)
        texto = status["texto"].lower()
        mostrar = True

        if filtro == "vencidos":
            mostrar = "vencido" in texto
        elif filtro == "proximos":
            mostrar = "vence" in texto
        elif filtro == "estoque":
            mostrar = "estoque" in texto

        if mostrar:
            produtos_status.append({"produto": produto, "status": status})

    return render_template(
        "produtos.html",
        produtos_status=produtos_status,
        filtro=filtro,
        busca=busca,
        categoria_id=categoria_id,
        categorias=categorias_lista
    )


@app.route("/produtos/baixar_estoque/<int:id>", methods=["POST"])
@login_obrigatorio
def baixar_estoque(id):
    try:
        quantidade = int(request.form.get("quantidade", 0))
    except ValueError:
        quantidade = 0

    observacao = request.form.get("observacao", "").strip()

    if quantidade <= 0:
        flash("Quantidade inválida.", "erro")
        return redirect(request.referrer or url_for("produtos"))

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM produtos WHERE id = %s", (id,))
    produto = cursor.fetchone()

    if not produto:
        conn.close()
        flash("Produto não encontrado.", "erro")
        return redirect(request.referrer or url_for("produtos"))

    estoque_anterior = produto["quantidade_atual"]

    if quantidade > estoque_anterior:
        conn.close()
        flash("Não é possível baixar mais do que o estoque atual.", "erro")
        return redirect(request.referrer or url_for("produtos"))

    estoque_atual = estoque_anterior - quantidade

    cursor.execute("UPDATE produtos SET quantidade_atual = %s WHERE id = %s", (estoque_atual, id))
    registrar_movimentacao(cursor, id, produto["nome"], "saida", quantidade, estoque_anterior, estoque_atual, observacao)

    conn.commit()
    conn.close()

    flash("Estoque baixado e histórico registrado com sucesso.", "sucesso")
    return redirect(request.referrer or url_for("produtos"))


@app.route("/produtos/repor_estoque/<int:id>", methods=["POST"])
@login_obrigatorio
def repor_estoque(id):
    try:
        quantidade = int(request.form.get("quantidade", 0))
    except ValueError:
        quantidade = 0

    observacao = request.form.get("observacao", "").strip()

    if quantidade <= 0:
        flash("Quantidade inválida.", "erro")
        return redirect(request.referrer or url_for("produtos"))

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM produtos WHERE id = %s", (id,))
    produto = cursor.fetchone()

    if not produto:
        conn.close()
        flash("Produto não encontrado.", "erro")
        return redirect(request.referrer or url_for("produtos"))

    estoque_anterior = produto["quantidade_atual"]
    estoque_atual = estoque_anterior + quantidade

    cursor.execute("UPDATE produtos SET quantidade_atual = %s WHERE id = %s", (estoque_atual, id))
    registrar_movimentacao(cursor, id, produto["nome"], "entrada", quantidade, estoque_anterior, estoque_atual, observacao)

    conn.commit()
    conn.close()

    flash("Estoque reposto e histórico registrado com sucesso.", "sucesso")
    return redirect(request.referrer or url_for("produtos"))


def validar_produto(form):
    erros = []
    nome = form.get("nome", "").strip()
    categoria_id = form.get("categoria_id", "").strip()
    data_vencimento = form.get("data_vencimento", "").strip()
    data_abertura = form.get("data_abertura", "").strip()

    quantidade_atual = int(form.get("quantidade_atual") or 0)
    estoque_padrao = int(form.get("estoque_padrao") or 0)
    limite_alerta = int(form.get("limite_alerta") or 0)
    validade_apos_aberto_dias = form.get("validade_apos_aberto_dias", "").strip()

    if not nome:
        erros.append("Nome do produto é obrigatório.")
    if not categoria_id:
        erros.append("Categoria é obrigatória.")
    if not data_vencimento:
        erros.append("Data de vencimento é obrigatória.")
    if quantidade_atual < 0:
        erros.append("Quantidade atual não pode ser negativa.")
    if estoque_padrao < 0:
        erros.append("Estoque padrão não pode ser negativo.")
    if limite_alerta < 0:
        erros.append("Limite de alerta não pode ser negativo.")

    if data_abertura:
        abertura = converter_data(data_abertura)
        if abertura and abertura > date.today():
            erros.append("Data de abertura não pode ser futura.")
        if not validade_apos_aberto_dias:
            erros.append("Informe a validade após aberto em dias.")

    if validade_apos_aberto_dias and int(validade_apos_aberto_dias) <= 0:
        erros.append("Validade após aberto deve ser maior que zero.")

    return erros


@app.route("/produtos/novo", methods=["GET", "POST"])
@login_obrigatorio
def novo_produto():
    categorias_lista = listar_categorias()

    if request.method == "POST":
        erros = validar_produto(request.form)

        if erros:
            for erro in erros:
                flash(erro, "erro")
            return render_template("novo_produto.html", categorias=categorias_lista)

        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO produtos (
                nome, categoria_id, lote, data_vencimento, data_abertura,
                validade_apos_aberto_dias, quantidade_atual, estoque_padrao,
                limite_alerta, observacoes
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            request.form.get("nome").strip(),
            request.form.get("categoria_id"),
            request.form.get("lote", "").strip(),
            request.form.get("data_vencimento"),
            request.form.get("data_abertura") or None,
            request.form.get("validade_apos_aberto_dias") or None,
            int(request.form.get("quantidade_atual") or 0),
            int(request.form.get("estoque_padrao") or 0),
            int(request.form.get("limite_alerta") or 0),
            request.form.get("observacoes", "").strip()
        ))
        conn.commit()
        conn.close()

        flash("Produto cadastrado com sucesso.", "sucesso")
        return redirect(url_for("produtos"))

    return render_template("novo_produto.html", categorias=categorias_lista)


@app.route("/produtos/editar/<int:id>", methods=["GET", "POST"])
@login_obrigatorio
def editar_produto(id):
    categorias_lista = listar_categorias()
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM produtos WHERE id = %s", (id,))
    produto = cursor.fetchone()

    if not produto:
        conn.close()
        flash("Produto não encontrado.", "erro")
        return redirect(url_for("produtos"))

    if request.method == "POST":
        erros = validar_produto(request.form)

        if erros:
            for erro in erros:
                flash(erro, "erro")
            conn.close()
            return render_template("editar_produto.html", produto=produto, categorias=categorias_lista)

        cursor.execute("""
            UPDATE produtos SET
                nome = %s,
                categoria_id = %s,
                lote = %s,
                data_vencimento = %s,
                data_abertura = %s,
                validade_apos_aberto_dias = %s,
                quantidade_atual = %s,
                estoque_padrao = %s,
                limite_alerta = %s,
                observacoes = %s
            WHERE id = %s
        """, (
            request.form.get("nome").strip(),
            request.form.get("categoria_id"),
            request.form.get("lote", "").strip(),
            request.form.get("data_vencimento"),
            request.form.get("data_abertura") or None,
            request.form.get("validade_apos_aberto_dias") or None,
            int(request.form.get("quantidade_atual") or 0),
            int(request.form.get("estoque_padrao") or 0),
            int(request.form.get("limite_alerta") or 0),
            request.form.get("observacoes", "").strip(),
            id
        ))
        conn.commit()
        conn.close()
        flash("Produto atualizado com sucesso.", "sucesso")
        return redirect(url_for("produtos"))

    conn.close()
    return render_template("editar_produto.html", produto=produto, categorias=categorias_lista)


@app.route("/produtos/excluir/<int:id>", methods=["POST"])
@login_obrigatorio
def excluir_produto(id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM produtos WHERE id = %s", (id,))
    conn.commit()
    conn.close()
    flash("Produto excluído com sucesso.", "sucesso")
    return redirect(url_for("produtos"))


@app.route("/historico")
@login_obrigatorio
def historico():
    busca = request.args.get("busca", "").strip()
    tipo = request.args.get("tipo", "").strip()

    conn = conectar()
    cursor = conn.cursor()

    query = "SELECT * FROM movimentacoes WHERE 1=1"
    params = []

    if busca:
        query += " AND produto_nome ILIKE %s"
        params.append(f"%{busca}%")

    if tipo:
        query += " AND tipo_movimentacao = %s"
        params.append(tipo)

    query += " ORDER BY data_movimentacao DESC"

    cursor.execute(query, params)
    movimentacoes = cursor.fetchall()
    conn.close()

    return render_template("historico.html", movimentacoes=movimentacoes, busca=busca, tipo=tipo, produto_nome=None)


@app.route("/historico/produto/<int:id>")
@login_obrigatorio
def historico_produto(id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT nome FROM produtos WHERE id = %s", (id,))
    produto = cursor.fetchone()
    cursor.execute("SELECT * FROM movimentacoes WHERE produto_id = %s ORDER BY data_movimentacao DESC", (id,))
    movimentacoes = cursor.fetchall()
    conn.close()

    produto_nome = produto["nome"] if produto else "Produto removido"
    return render_template("historico.html", movimentacoes=movimentacoes, busca="", tipo="", produto_nome=produto_nome)


@app.route("/relatorios")
@login_obrigatorio
def relatorios():
    return render_template("relatorios.html")


def nome_arquivo_relatorio(tipo_relatorio, formato):
    data = datetime.now().strftime("%Y-%m-%d")
    nome = tipo_relatorio.replace("_", "-")
    return f"{nome}_{data}.{formato}"


def aplicar_estilo_excel(ws, titulo, total_colunas):
    ws.insert_rows(1, 3)
    ws["A1"] = "Controle Oftalmo"
    ws["A2"] = titulo
    ws["A1"].font = Font(bold=True, size=16, color="0F172A")
    ws["A2"].font = Font(bold=True, size=13, color="2563EB")

    header_row = 4
    fill = PatternFill("solid", fgColor="0F172A")
    font = Font(bold=True, color="FFFFFF")
    border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for col in range(1, total_colunas + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for row in ws.iter_rows(min_row=4):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for col in range(1, total_colunas + 1):
        letter = get_column_letter(col)
        maior = 12
        for cell in ws[letter]:
            if cell.value:
                maior = max(maior, len(str(cell.value)) + 2)
        ws.column_dimensions[letter].width = min(maior, 35)


def gerar_excel(titulo, cabecalhos, linhas, nome_arquivo):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"
    ws.append(cabecalhos)

    for linha in linhas:
        ws.append(linha)

    aplicar_estilo_excel(ws, titulo, len(cabecalhos))

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name=nome_arquivo, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def gerar_pdf(titulo, cabecalhos, linhas, nome_arquivo):
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=1 * cm, leftMargin=1 * cm, topMargin=1 * cm, bottomMargin=1 * cm)
    styles = getSampleStyleSheet()
    elementos = []

    logo_path = os.path.join(app.root_path, "static", "logo.png")
    if os.path.exists(logo_path):
        try:
            elementos.append(Image(logo_path, width=2.0 * cm, height=2.0 * cm))
        except Exception:
            pass

    elementos.append(Paragraph("<b>Controle Oftalmo</b>", styles["Title"]))
    elementos.append(Paragraph(titulo, styles["Heading2"]))
    elementos.append(Paragraph(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
    elementos.append(Spacer(1, 0.4 * cm))

    if not linhas:
        elementos.append(Paragraph("Nenhum dado encontrado para este relatório.", styles["Normal"]))
    else:
        dados = [cabecalhos]
        for linha in linhas:
            dados.append([str(valor) if valor is not None else "-" for valor in linha])

        tabela = Table(dados, repeatRows=1)
        tabela.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        elementos.append(tabela)

    doc.build(elementos)
    output.seek(0)

    return send_file(output, as_attachment=True, download_name=nome_arquivo, mimetype="application/pdf")


def buscar_produtos_para_relatorio(busca=""):
    conn = conectar()
    cursor = conn.cursor()
    query = """
        SELECT produtos.*, categorias.nome AS categoria_nome
        FROM produtos
        JOIN categorias ON categorias.id = produtos.categoria_id
        WHERE 1=1
    """
    params = []

    if busca:
        query += " AND (produtos.nome ILIKE %s OR produtos.lote ILIKE %s)"
        params.extend([f"%{busca}%", f"%{busca}%"])

    query += " ORDER BY produtos.nome"
    cursor.execute(query, params)
    produtos_lista = cursor.fetchall()
    conn.close()
    return produtos_lista


def montar_relatorio_produtos(tipo_relatorio, busca=""):
    produtos_lista = buscar_produtos_para_relatorio(busca)

    cabecalhos = ["Produto", "Categoria", "Lote", "Vencimento", "Aberto em", "Vence após aberto", "Estoque", "Estoque padrão", "Limite alerta", "Status"]
    linhas = []

    for produto in produtos_lista:
        status = calcular_status(produto)
        texto = status["texto"].lower()
        incluir = True

        if tipo_relatorio == "vencimentos":
            incluir = "vencido" in texto or "vence" in texto
        elif tipo_relatorio == "estoque":
            incluir = "estoque" in texto

        if incluir:
            linhas.append([
                produto["nome"],
                produto["categoria_nome"],
                produto["lote"] or "-",
                formatar_data(produto["data_vencimento"]),
                formatar_data(produto["data_abertura"]),
                formatar_data(status["vencimento_apos_aberto"]),
                produto["quantidade_atual"],
                produto["estoque_padrao"],
                produto["limite_alerta"],
                status["texto"]
            ])

    if tipo_relatorio == "produtos":
        titulo = "Relatório geral de produtos"
    elif tipo_relatorio == "vencimentos":
        titulo = "Relatório de vencidos e próximos do vencimento"
    else:
        titulo = "Relatório de estoque baixo e zerado"

    return titulo, cabecalhos, linhas


def montar_relatorio_movimentacoes(busca="", tipo_movimentacao=""):
    conn = conectar()
    cursor = conn.cursor()
    query = "SELECT * FROM movimentacoes WHERE 1=1"
    params = []

    if busca:
        query += " AND produto_nome ILIKE %s"
        params.append(f"%{busca}%")

    if tipo_movimentacao:
        query += " AND tipo_movimentacao = %s"
        params.append(tipo_movimentacao)

    query += " ORDER BY data_movimentacao DESC"
    cursor.execute(query, params)
    movimentacoes = cursor.fetchall()
    conn.close()

    cabecalhos = ["Data", "Produto", "Tipo", "Quantidade", "Estoque anterior", "Estoque atual", "Observação"]
    linhas = []

    for mov in movimentacoes:
        linhas.append([
            formatar_data(mov["data_movimentacao"]),
            mov["produto_nome"],
            "Entrada" if mov["tipo_movimentacao"] == "entrada" else "Saída",
            mov["quantidade"],
            mov["estoque_anterior"],
            mov["estoque_atual"],
            mov["observacao"] or "-"
        ])

    return "Relatório de movimentações de estoque", cabecalhos, linhas


@app.route("/relatorios/exportar")
@login_obrigatorio
def exportar_relatorio():
    tipo_relatorio = request.args.get("tipo_relatorio", "produtos")
    formato = request.args.get("formato", "xlsx")
    busca = request.args.get("busca", "").strip()
    tipo_movimentacao = request.args.get("tipo_movimentacao", "").strip()

    if tipo_relatorio == "movimentacoes":
        titulo, cabecalhos, linhas = montar_relatorio_movimentacoes(busca, tipo_movimentacao)
    else:
        titulo, cabecalhos, linhas = montar_relatorio_produtos(tipo_relatorio, busca)

    if formato not in ["xlsx", "pdf"]:
        flash("Formato de relatório inválido.", "erro")
        return redirect(url_for("relatorios"))

    nome_arquivo = nome_arquivo_relatorio(tipo_relatorio, formato)

    if formato == "xlsx":
        return gerar_excel(titulo, cabecalhos, linhas, nome_arquivo)

    return gerar_pdf(titulo, cabecalhos, linhas, nome_arquivo)


def senha_padrao_ativa():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT senha_hash FROM admin WHERE usuario = %s", ("admin",))
    admin = cursor.fetchone()
    conn.close()

    if not admin:
        return False

    return check_password_hash(admin["senha_hash"], "admin123")


@app.route("/configuracoes")
@login_obrigatorio
def configuracoes():
    return render_template("configuracoes.html", senha_padrao=senha_padrao_ativa())


@app.route("/configuracoes/alterar_senha", methods=["POST"])
@login_obrigatorio
def alterar_senha():
    senha_atual = request.form.get("senha_atual", "").strip()
    nova_senha = request.form.get("nova_senha", "").strip()
    confirmar_senha = request.form.get("confirmar_senha", "").strip()

    if len(nova_senha) < 6:
        flash("A nova senha deve ter pelo menos 6 caracteres.", "erro")
        return redirect(url_for("configuracoes"))

    if nova_senha != confirmar_senha:
        flash("A confirmação da nova senha não confere.", "erro")
        return redirect(url_for("configuracoes"))

    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM admin WHERE usuario = %s", (session.get("usuario", "admin"),))
    admin = cursor.fetchone()

    if not admin or not check_password_hash(admin["senha_hash"], senha_atual):
        conn.close()
        flash("Senha atual incorreta.", "erro")
        return redirect(url_for("configuracoes"))

    cursor.execute("UPDATE admin SET senha_hash = %s WHERE id = %s", (generate_password_hash(nova_senha), admin["id"]))
    conn.commit()
    conn.close()

    flash("Senha alterada com sucesso.", "sucesso")
    return redirect(url_for("configuracoes"))


# Inicialização do banco para ambiente local e deploy
criar_banco()

if __name__ == "__main__":
    app.run(debug=True)
