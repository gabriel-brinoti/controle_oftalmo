import json
import sys
from pathlib import Path
from types import SimpleNamespace

from jinja2 import Environment, FileSystemLoader, select_autoescape

from validar_migracao_estoque import ESTOQUES, migrar_backup


def nome_tipo_estoque(tipo):
    return ESTOQUES.get(tipo, "Almoxarifado")


def url_for(endpoint, **values):
    params = "&".join(f"{key}={value}" for key, value in values.items() if value not in (None, ""))
    return f"/{endpoint}" + (f"?{params}" if params else "")


def status_para_quantidade(produto):
    quantidade = produto.get("quantidade_atual") or 0
    limite = produto.get("limite_alerta") or 0

    if quantidade <= 0:
        return {"tipo": "danger", "texto": "Estoque zerado"}
    if quantidade <= limite:
        return {"tipo": "warning", "texto": "Estoque baixo"}
    return {"tipo": "ok", "texto": "Produto OK"}


def montar_produtos_status(migrado):
    produtos = {produto["id"]: produto for produto in migrado["produtos"]}
    lotes_por_estoque = {}

    for lote in migrado["lotes"]:
        chave = lote["produto_estoque_id"]
        lotes_por_estoque.setdefault(chave, []).append(lote)

    itens = []
    for produto_estoque in migrado["produto_estoques"]:
        produto = produtos[produto_estoque["produto_id"]]
        lotes = lotes_por_estoque.get(produto_estoque["id"], [])
        quantidade_atual = sum(lote.get("quantidade_atual") or 0 for lote in lotes)
        proximas_validade = sorted(
            lote["data_vencimento"]
            for lote in lotes
            if lote.get("data_vencimento") and (lote.get("quantidade_atual") or 0) > 0
        )
        item_produto = {
            **produto,
            **produto_estoque,
            "produto_id": produto["id"],
            "produto_estoque_id": produto_estoque["id"],
            "categoria_nome": f"Categoria {produto['categoria_id']}",
            "quantidade_atual": quantidade_atual,
            "total_lotes": len(lotes),
            "proximo_vencimento": proximas_validade[0] if proximas_validade else None,
        }
        itens.append({
            "produto": item_produto,
            "status": status_para_quantidade(item_produto),
            "vencido": False,
        })

    return itens


def renderizar(template, contexto):
    env = Environment(
        loader=FileSystemLoader("templates"),
        autoescape=select_autoescape(["html"]),
    )
    env.globals.update({
        "url_for": url_for,
        "request": SimpleNamespace(args={}, endpoint="produtos"),
        "session": {"perfil": "admin", "usuario_nome": "Teste"},
        "get_flashed_messages": lambda with_categories=True: [],
        "estoques_sistema": ESTOQUES,
        "nome_tipo_estoque": nome_tipo_estoque,
    })
    return env.get_template(template).render(**contexto)


def main():
    if len(sys.argv) < 2:
        print("Uso: python testar_migracao_offline.py caminho_do_backup.json")
        return 2

    caminho = Path(sys.argv[1])
    backup = json.loads(caminho.read_text(encoding="utf-8"))
    migrado = migrar_backup(backup)
    produtos_status = montar_produtos_status(migrado)

    html_produtos = renderizar("produtos.html", {
        "produtos_status": produtos_status[:3],
        "produtos_vencidos": [],
        "filtro": "todos",
        "busca": "",
        "categoria_id": "",
        "categorias": backup.get("dados", {}).get("categorias", []),
        "tipo_estoque": None,
        "tipo_estoque_nome": "Todos os Estoques",
        "pagina": 1,
        "total_paginas": 1,
        "total_resultados": len(produtos_status),
    })

    produtos_select = [
        {
            **produto,
            "categoria_nome": f"Categoria {produto['categoria_id']}",
        }
        for produto in migrado["produtos"][:20]
    ]
    html_entrada = renderizar("nova_entrada.html", {"produtos": produtos_select})

    primeiro_estoque = migrado["produto_estoques"][0]
    lotes = [
        {
            **lote,
            "nome": migrado["produtos"][primeiro_estoque["produto_id"] - 1]["nome"],
            "categoria_nome": f"Categoria {migrado['produtos'][primeiro_estoque['produto_id'] - 1]['categoria_id']}",
            "produto_id": primeiro_estoque["produto_id"],
            "tipo_estoque": primeiro_estoque["tipo_estoque"],
        }
        for lote in migrado["lotes"]
        if lote["produto_estoque_id"] == primeiro_estoque["id"]
    ]
    html_lotes = renderizar("lotes_produto.html", {"lotes": lotes, "produto": lotes[0]})

    checks = {
        "produtos_tem_nova_entrada": "Nova entrada" in html_produtos,
        "produtos_tem_ver_lotes": "Ver lotes" in html_produtos,
        "entrada_tem_destino": "Destino" in html_entrada,
        "entrada_tem_lote": "Lote" in html_entrada,
        "lotes_tem_transferir": "Transferir" in html_lotes,
        "lotes_tem_retirar": "Retirar" in html_lotes,
    }

    for nome, ok in checks.items():
        print(f"{nome}: {'OK' if ok else 'FALHOU'}")

    print(f"produtos_status_renderizados: {len(produtos_status[:3])}")
    print(f"produto_estoques_total: {len(migrado['produto_estoques'])}")
    print(f"lotes_total: {len(migrado['lotes'])}")

    return 0 if all(checks.values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())
