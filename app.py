
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
    "Casa e decoração",
    "Beleza e cuidados pessoais"
]

CATEGORIA_PRODUTO = {
    "FONE BLUETOOTH": "Eletrônicos",
    "CARREGADOR PORTATIL": "Eletrônicos",
    "SMARTWATCH": "Eletrônicos",
    "CAIXA DE SOM": "Eletrônicos",

    "CAMISETA BASICA": "Moda e acessórios",
    "CALCA JEANS": "Moda e acessórios",
    "JAQUETA CORTA-VENTO": "Moda e acessórios",
    "TENIS CASUAL": "Moda e acessórios",

    "LUMINARIA DE MESA": "Casa e decoração",
    "JOGO DE PANELAS": "Casa e decoração",
    "ALMOFADA DECORATIVA": "Casa e decoração",
    "TAPETE": "Casa e decoração",

    "KIT SKINCARE": "Beleza e cuidados pessoais",
    "PERFUME": "Beleza e cuidados pessoais",
    "SECADOR DE CABELO": "Beleza e cuidados pessoais"
}


def identificar_categoria(produto):
    nome = unicodedata.normalize("NFKD", produto)
    nome = "".join(
        caractere for caractere in nome
        if not unicodedata.combining(caractere)
    ).upper()

    nome = " ".join(nome.split())

    if nome not in CATEGORIA_PRODUTO:
        raise ValueError(f"Produto sem categoria: {produto}")

    return CATEGORIA_PRODUTO[nome]

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
            categoria_registro = (
                identificar_categoria(produto)
                if categoria == "Automática (arquivo misto)"
                else categoria
            )
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
                "categoria": categoria_registro,
                "loja": loja,
                "sku": sku,
                "produto": produto,
                "quantidade_estoque": quantidade,
                "preco_custo": preco,
                "ultima_atualizacao": data,
                "status_qualidade": ",".join(avisos) if avisos else "OK"
            }

            # A mesma combinação não entra duas vezes
            chave = (categoria_registro, loja, sku, produto.casefold())
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

    if categoria not in CATEGORIAS + ["Automática (arquivo misto)"]:
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

# Opcoes de graficos disponíveis para o gestor
OPCOES_GRAFICOS = {
    "unidades_categoria": "Unidades por categoria",
    "unidades_loja": "Unidades por loja",
    "valor_categoria": "Valor de estoque por categoria",
    "valor_loja": "Valor de estoque por loja",
    "problemas_categoria": "Problemas por categoria",
    "registros_categoria": "Registros por categoria"
}

def interpretar_grafico(nome, valores):
    # Nao tirar conclusoes quando nao houver valores
    total = sum(valores.values())

    if total <= 0:
        return (
            "Nao ha valores suficientes para analisar "
            "este indicador com os filtros selecionados."
        )

    maior = max(valores.values())

    # Encontrar todos os grupos empatados em primeiro
    lideres = [
        grupo for grupo, valor in valores.items()
        if valor == maior
    ]

    percentual = maior / total * 100
    participacao = f"{percentual:.1f}".replace(".", ",")

    if len(lideres) > 1:
        return (
            "Ha empate entre os maiores grupos: "
            + ", ".join(lideres)
            + ". Consulte os valores antes de tomar decisoes."
        )

    lider = lideres[0]

    if nome.startswith("unidades_"):
        return (
            f"{lider} concentra a maior quantidade de "
            f"unidades no filtro aplicado, com "
            f"{participacao}% do total. "
            "Vale investigar se essa distribuicao "
            "esta alinhada com a demanda."
        )

    if nome.startswith("valor_"):
        valor_formatado = (
            f"{maior:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

        return (
            f"{lider} apresenta o maior valor de custo "
            f"do estoque: R$ {valor_formatado}, "
            f"equivalente a {participacao}% do total. "
            "Vale analisar a concentracao de capital "
            "antes de decidir sobre novas compras."
        )

    if nome == "problemas_categoria":
        return (
            f"{lider} possui o maior numero de registros "
            f"com problemas de qualidade: {int(maior)}. "
            "Priorize a verificacao desses dados. "
            "Compare tambem a taxa de problemas entre categorias."
        )

    if nome == "registros_categoria":
        return (
            f"{lider} concentra mais registros de estoque: "
            f"{int(maior)}, ou {participacao}% do total. "
            "Esses registros nao representam "
            "necessariamente produtos diferentes."
        )

    return "Analise nao disponivel."

@app.route("/graficos")
@protegido
def graficos_personalizados():

    categoria = request.args.get("categoria", "Todas")
    loja = request.args.get("loja", "Todas")

    selecionados = [
        nome for nome in request.args.getlist("grafico")
        if nome in OPCOES_GRAFICOS
    ]

    # Na primeira visita, mostrar duas sugestões
    if not request.args:
        selecionados = [
            "unidades_categoria",
            "valor_loja"
        ]

    # Consultar os registros diretamente no Supabase
    consulta = db.table("estoque").select("*")

    if categoria in CATEGORIAS:
        consulta = consulta.eq("categoria", categoria)

    if loja in LOJAS.values():
        consulta = consulta.eq("loja", loja)

    dados = consulta.range(0, 999).execute().data

    # Agrupar os indicadores
    unidades_categoria = {c: 0 for c in CATEGORIAS}
    unidades_loja = {l: 0 for l in LOJAS.values()}

    valor_categoria = {c: 0 for c in CATEGORIAS}
    valor_loja = {l: 0 for l in LOJAS.values()}

    problemas_categoria = {c: 0 for c in CATEGORIAS}
    registros_categoria = {c: 0 for c in CATEGORIAS}

    for item in dados:
        cat = item["categoria"]
        lj = item["loja"]

        registros_categoria[cat] += 1

        if item["status_qualidade"] != "OK":
            problemas_categoria[cat] += 1

        quantidade = item["quantidade_estoque"]
        preco = item["preco_custo"]

        if quantidade is not None:
            unidades_categoria[cat] += quantidade
            unidades_loja[lj] += quantidade

            if preco is not None:
                valor = quantidade * float(preco)
                valor_categoria[cat] += valor
                valor_loja[lj] += valor

    # Dados prontos para o Chart.js
    fontes = {
        "unidades_categoria": unidades_categoria,
        "unidades_loja": unidades_loja,
        "valor_categoria": valor_categoria,
        "valor_loja": valor_loja,
        "problemas_categoria": problemas_categoria,
        "registros_categoria": registros_categoria
    }

    tipos = {
        "unidades_categoria": "doughnut",
        "unidades_loja": "bar",
        "valor_categoria": "bar",
        "valor_loja": "bar",
        "problemas_categoria": "bar",
        "registros_categoria": "doughnut"
    }

    graficos = []

    for nome in selecionados:
        graficos.append({
            "id": nome,
            "titulo": OPCOES_GRAFICOS[nome],
            "tipo": tipos[nome],
            "labels": list(fontes[nome].keys()),
            "valores": list(fontes[nome].values()),
            "moeda": nome.startswith("valor_"),
            "interpretacao": interpretar_grafico(nome, fontes[nome])
        })

    return render_template(
        "graficos.html",
        opcoes=OPCOES_GRAFICOS,
        selecionados=selecionados,
        graficos=graficos,
        categorias=CATEGORIAS,
        lojas=list(LOJAS.values()),
        categoria_atual=categoria,
        loja_atual=loja
    )

if __name__ == "__main__":
    app.run(debug=True)
