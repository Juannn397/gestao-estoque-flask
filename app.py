
import os
import re
import secrets
import unicodedata
from datetime import datetime
from functools import wraps

import pandas as pd
from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, session, url_for
from supabase import create_client

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ["FLASK_SECRET_KEY"]
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

db = create_client(
    os.environ["SUPABASE_URL"],
    os.environ["SUPABASE_SECRET_KEY"]
)

CATEGORIAS = [
    "Eletrônicos",
    "Moda e acessórios",
    "Casa e decoração"
]

LOJAS = {
    "Loja_Centro": "Loja Centro",
    "Loja_Norte": "Loja Norte",
    "Loja_Sul": "Loja Sul"
}

COLUNAS = [
    "sku", "produto", "quantidade_estoque",
    "preco_custo", "loja", "ultima_atualizacao"
]


# Restringir o painel ao administrador
def protegido(funcao):
    @wraps(funcao)
    def verificar(*args, **kwargs):
        if not session.get("autenticado"):
            return redirect(url_for("login"))
        return funcao(*args, **kwargs)
    return verificar


@app.route("/login", methods=["GET", "POST"])
def login():
    erro = None

    if request.method == "POST":
        senha = request.form.get("senha", "")
        correta = os.environ["ADMIN_PASSWORD"]

        if secrets.compare_digest(senha, correta):
            session.clear()
            session["autenticado"] = True
            return redirect(url_for("inicio"))

        erro = "Senha incorreta."

    return render_template(
        "index.html", login=True, erro=erro
    )


@app.route("/sair")
def sair():
    session.clear()
    return redirect(url_for("login"))


# Transformar textos em formato padronizado
def limpar_texto(valor):
    if pd.isna(valor):
        return ""

    return " ".join(str(valor).strip().split())


# Normalizar o SKU
def limpar_sku(valor):
    return limpar_texto(valor).upper()


# Converter datas misturadas para AAAA-MM-DD
def limpar_data(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return None

    if isinstance(valor, (datetime, pd.Timestamp)):
        return valor.strftime("%Y-%m-%d")

    texto = str(valor).strip()

    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto, formato).strftime("%Y-%m-%d")
        except ValueError:
            pass

    return None


def limpar_quantidade(valor):
    if pd.isna(valor):
        return None, "SEM_QUANTIDADE"

    try:
        numero = float(valor)

        if not numero.is_integer():
            return None, "QUANTIDADE_INVALIDA"

        if numero < 0:
            return None, "QUANTIDADE_NEGATIVA"

        return int(numero), None

    except (ValueError, TypeError):
        return None, "QUANTIDADE_INVALIDA"


def limpar_preco(valor):
    if pd.isna(valor):
        return None

    try:
        numero = float(valor)

        if numero < 0:
            return None

        return round(numero, 2)

    except (TypeError, ValueError):
        return None


# Ler todas as abas e tratar os registros
def preparar_planilha(arquivo, categoria):
    abas = pd.read_excel(
        arquivo,
        sheet_name=None,
        engine="openpyxl"
    )

    if set(abas.keys()) != set(LOJAS.keys()):
        raise ValueError(
            "O Excel precisa conter exatamente as abas "
            "Loja_Centro, Loja_Norte e Loja_Sul."
        )

    registros = {}

    for nome_aba, tabela in abas.items():

        if not set(COLUNAS).issubset(tabela.columns):
            raise ValueError(
                f"Colunas obrigatórias ausentes em {nome_aba}."
            )

        loja = LOJAS[nome_aba]

        for _, linha in tabela.iterrows():
            sku = limpar_sku(linha["sku"])
            produto = limpar_texto(linha["produto"])

            if not sku or not produto:
                continue

            quantidade, problema = limpar_quantidade(
                linha["quantidade_estoque"]
            )

            preco = limpar_preco(linha["preco_custo"])
            data = limpar_data(linha["ultima_atualizacao"])

            avisos = []

            if problema:
                avisos.append(problema)

            if preco is None:
                avisos.append("PRECO_INVALIDO")

            if data is None:
                avisos.append("DATA_AUSENTE")

            registro = {
                "categoria": categoria,
                "loja": loja,
                "sku": sku,
                "produto": produto,
                "quantidade_estoque": quantidade,
                "preco_custo": preco,
                "ultima_atualizacao": data,
                "status_qualidade": ",".join(avisos) if avisos else "OK"
            }

            # A mesma combinação não entra duas vezes
            chave = (categoria, loja, sku, produto.casefold())
            registros[chave] = registro

    return list(registros.values())


# Dashboard: consultar somente o banco Supabase
@app.route("/")
@protegido
def inicio():
    categoria = request.args.get("categoria", "Todas")
    loja = request.args.get("loja", "Todas")

    consulta = db.table("estoque").select("*")

    if categoria in CATEGORIAS:
        consulta = consulta.eq("categoria", categoria)

    if loja in LOJAS.values():
        consulta = consulta.eq("loja", loja)

    # Os arquivos da aula totalizam menos de 1000 registros.
    dados = consulta.range(0, 999).execute().data

    total_unidades = 0
    valor_estoque = 0
    problemas = 0

    por_categoria = {nome: 0 for nome in CATEGORIAS}
    por_loja = {nome: 0 for nome in LOJAS.values()}

    for item in dados:
        quantidade = item["quantidade_estoque"]
        preco = item["preco_custo"]

        if item["status_qualidade"] != "OK":
            problemas += 1

        if quantidade is not None:
            total_unidades += quantidade
            por_categoria[item["categoria"]] += quantidade
            por_loja[item["loja"]] += quantidade

            if preco is not None:
                valor_estoque += quantidade * float(preco)

    return render_template(
        "index.html",
        login=False,
        categorias=CATEGORIAS,
        lojas=list(LOJAS.values()),
        categoria_atual=categoria,
        loja_atual=loja,
        registros=len(dados),
        unidades=total_unidades,
        valor=valor_estoque,
        problemas=problemas,
        por_categoria=por_categoria,
        por_loja=por_loja,
        mensagem=request.args.get("mensagem", "")
    )


# Upload do Excel e gravação no banco
@app.route("/importar", methods=["POST"])
@protegido
def importar():
    categoria = request.form.get("categoria", "")
    arquivo = request.files.get("arquivo")

    if categoria not in CATEGORIAS:
        return "Selecione uma categoria válida.", 400

    if not arquivo or not arquivo.filename.lower().endswith(".xlsx"):
        return "Selecione um arquivo Excel .xlsx.", 400

    try:
        registros = preparar_planilha(arquivo, categoria)

        if not registros:
            return "Nenhum registro válido encontrado.", 400

        db.table("estoque").upsert(
            registros,
            on_conflict="categoria,loja,sku,produto"
        ).execute()

        mensagem = f"{len(registros)} registros importados: {categoria}"

        return redirect(url_for("inicio", mensagem=mensagem))

    except Exception as erro:
        app.logger.exception("Erro na importação")
        return "Falha na importação. Consulte o terminal ou os logs.", 500


if __name__ == "__main__":
    app.run(debug=True)
