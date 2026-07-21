import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path


ESTOQUES = {
    "almoxarifado": "Almoxarifado",
    "farmacia_satelite": "Farmácia Satélite",
    "carrinho_urgencia": "Carrinho de Urgência",
}


def normalizar_texto(valor):
    return re.sub(r"\s+", " ", str(valor or "").strip()).lower()


def chave_produto(produto):
    return (
        normalizar_texto(produto.get("nome")),
        produto.get("categoria_id"),
        normalizar_texto(produto.get("codigo_barras")),
    )


def chave_produto_estoque(produto_novo_id, tipo_estoque):
    return (produto_novo_id, tipo_estoque or "almoxarifado")


def numero_lote(produto):
    lote = str(produto.get("lote") or "").strip()
    return lote or "Sem lote"


def carregar_backup(caminho):
    with open(caminho, "r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


def migrar_backup(dados):
    produtos_antigos = dados.get("dados", {}).get("produtos", [])
    movimentacoes_antigas = dados.get("dados", {}).get("movimentacoes", [])

    produtos_novos = []
    produto_estoques = []
    lotes = []
    mapa_produtos = {}
    mapa_produto_estoque = {}
    mapa_lote_por_produto_antigo = {}
    mapa_produto_novo_por_produto_antigo = {}

    proximo_produto_id = 1
    proximo_produto_estoque_id = 1
    proximo_lote_id = 1

    for produto in produtos_antigos:
        chave = chave_produto(produto)

        if chave not in mapa_produtos:
            mapa_produtos[chave] = proximo_produto_id
            produtos_novos.append({
                "id": proximo_produto_id,
                "nome": produto.get("nome"),
                "categoria_id": produto.get("categoria_id"),
                "codigo_barras": produto.get("codigo_barras") or "",
                "fabricante": produto.get("fabricante") or "",
                "unidade_medida": produto.get("unidade_medida") or "",
                "validade_apos_aberto_dias": produto.get("validade_apos_aberto_dias"),
                "observacoes": produto.get("observacoes") or "",
                "ativo": produto.get("ativo", True),
                "criado_em": produto.get("criado_em"),
                "ids_antigos": [produto.get("id")],
            })
            proximo_produto_id += 1
        else:
            produto_novo_id = mapa_produtos[chave]
            for produto_novo in produtos_novos:
                if produto_novo["id"] == produto_novo_id:
                    produto_novo["ids_antigos"].append(produto.get("id"))
                    break

        produto_novo_id = mapa_produtos[chave]
        tipo_estoque = produto.get("tipo_estoque") or "almoxarifado"
        chave_pe = chave_produto_estoque(produto_novo_id, tipo_estoque)

        if chave_pe not in mapa_produto_estoque:
            mapa_produto_estoque[chave_pe] = proximo_produto_estoque_id
            produto_estoques.append({
                "id": proximo_produto_estoque_id,
                "produto_id": produto_novo_id,
                "tipo_estoque": tipo_estoque,
                "estoque_padrao": produto.get("estoque_padrao") or 0,
                "limite_alerta": produto.get("limite_alerta") or 0,
                "ativo": True,
            })
            proximo_produto_estoque_id += 1

        produto_estoque_id = mapa_produto_estoque[chave_pe]
        lote = {
            "id": proximo_lote_id,
            "produto_estoque_id": produto_estoque_id,
            "produto_id": produto_novo_id,
            "produto_antigo_id": produto.get("id"),
            "nome_produto": produto.get("nome"),
            "tipo_estoque": tipo_estoque,
            "numero_lote": numero_lote(produto),
            "data_vencimento": produto.get("data_vencimento"),
            "data_abertura": produto.get("data_abertura"),
            "quantidade_inicial": produto.get("quantidade_atual") or 0,
            "quantidade_atual": produto.get("quantidade_atual") or 0,
            "data_entrada": produto.get("criado_em"),
            "observacoes": produto.get("observacoes") or "",
            "ativo": produto.get("ativo", True),
        }
        lotes.append(lote)
        mapa_lote_por_produto_antigo[produto.get("id")] = proximo_lote_id
        mapa_produto_novo_por_produto_antigo[produto.get("id")] = produto_novo_id
        proximo_lote_id += 1

    movimentacoes_novas = []
    for mov in movimentacoes_antigas:
        produto_antigo_id = mov.get("produto_id")
        movimentacoes_novas.append({
            **mov,
            "lote_id": mapa_lote_por_produto_antigo.get(produto_antigo_id),
            "produto_id": mapa_produto_novo_por_produto_antigo.get(produto_antigo_id),
            "quantidade_anterior": mov.get("estoque_anterior"),
            "quantidade_posterior": mov.get("estoque_atual"),
            "origem": mov.get("estoque_origem"),
            "destino": mov.get("estoque_destino"),
            "motivo": mov.get("observacao"),
            "criado_em": mov.get("data_movimentacao"),
        })

    return {
        "produtos": produtos_novos,
        "produto_estoques": produto_estoques,
        "lotes": lotes,
        "movimentacoes": movimentacoes_novas,
    }


def somar_por_chave(produtos):
    totais = defaultdict(int)
    for produto in produtos:
        chave = (
            normalizar_texto(produto.get("nome")),
            produto.get("categoria_id"),
            normalizar_texto(produto.get("codigo_barras")),
            produto.get("tipo_estoque") or "almoxarifado",
        )
        totais[chave] += produto.get("quantidade_atual") or 0
    return totais


def somar_lotes_por_chave(migrado):
    produto_por_id = {produto["id"]: produto for produto in migrado["produtos"]}
    totais = defaultdict(int)
    for lote in migrado["lotes"]:
        produto = produto_por_id[lote["produto_id"]]
        chave = (
            normalizar_texto(produto.get("nome")),
            produto.get("categoria_id"),
            normalizar_texto(produto.get("codigo_barras")),
            lote.get("tipo_estoque") or "almoxarifado",
        )
        totais[chave] += lote.get("quantidade_atual") or 0
    return totais


def gerar_relatorio(caminho_backup, backup, migrado):
    produtos_antigos = backup.get("dados", {}).get("produtos", [])
    categorias = backup.get("dados", {}).get("categorias", [])
    movimentacoes = backup.get("dados", {}).get("movimentacoes", [])

    totais_antigos = somar_por_chave(produtos_antigos)
    totais_novos = somar_lotes_por_chave(migrado)
    divergencias = []

    for chave, total_antigo in totais_antigos.items():
        total_novo = totais_novos.get(chave, 0)
        if total_antigo != total_novo:
            divergencias.append({
                "chave": chave,
                "total_antigo": total_antigo,
                "total_novo": total_novo,
            })

    sem_codigo = sum(1 for produto in produtos_antigos if not str(produto.get("codigo_barras") or "").strip())
    sem_lote = sum(1 for produto in produtos_antigos if not str(produto.get("lote") or "").strip())
    sem_validade = sum(1 for produto in produtos_antigos if not produto.get("data_vencimento"))

    repetidos = defaultdict(list)
    for produto in produtos_antigos:
        repetidos[(normalizar_texto(produto.get("nome")), produto.get("tipo_estoque") or "almoxarifado")].append(produto)
    grupos_repetidos = [grupo for grupo in repetidos.values() if len(grupo) > 1]

    linhas = [
        "RELATÓRIO DE PRÉ-MIGRAÇÃO DO ESTOQUE",
        f"Backup analisado: {caminho_backup}",
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        "",
        "Resumo atual",
        f"- Categorias: {len(categorias)}",
        f"- Registros atuais em produtos: {len(produtos_antigos)}",
        f"- Movimentações atuais: {len(movimentacoes)}",
        f"- Produtos sem código de barras: {sem_codigo}",
        f"- Produtos sem lote: {sem_lote}",
        f"- Produtos sem validade: {sem_validade}",
        "",
        "Resumo após migração",
        f"- Produtos principais: {len(migrado['produtos'])}",
        f"- Produto por estoque: {len(migrado['produto_estoques'])}",
        f"- Lotes: {len(migrado['lotes'])}",
        f"- Movimentações vinculadas a lote: {sum(1 for mov in migrado['movimentacoes'] if mov.get('lote_id'))}",
        "",
        "Conferência de quantidade",
        f"- Grupos conferidos: {len(totais_antigos)}",
        f"- Divergências de quantidade: {len(divergencias)}",
        "",
        "Repetidos que viram lotes",
        f"- Grupos com mesmo nome e estoque: {len(grupos_repetidos)}",
        f"- Registros dentro desses grupos: {sum(len(grupo) for grupo in grupos_repetidos)}",
    ]

    if divergencias:
        linhas.append("")
        linhas.append("Divergências encontradas")
        for item in divergencias[:30]:
            linhas.append(f"- {item['chave']}: antigo={item['total_antigo']} novo={item['total_novo']}")

    return "\n".join(linhas), divergencias


def main():
    if len(sys.argv) < 2:
        print("Uso: python validar_migracao_estoque.py caminho_do_backup.json")
        return 2

    caminho_backup = Path(sys.argv[1])
    if not caminho_backup.exists():
        print(f"Backup não encontrado: {caminho_backup}")
        return 1

    backup = carregar_backup(caminho_backup)
    migrado = migrar_backup(backup)
    relatorio, divergencias = gerar_relatorio(caminho_backup, backup, migrado)

    pasta_saida = Path("backups")
    pasta_saida.mkdir(exist_ok=True)
    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_preview = pasta_saida / f"preview_migracao_estoque_{carimbo}.json"
    caminho_relatorio = pasta_saida / f"relatorio_migracao_estoque_{carimbo}.txt"

    with open(caminho_preview, "w", encoding="utf-8") as arquivo:
        json.dump({
            "origem": str(caminho_backup),
            "gerado_em": datetime.now().isoformat(),
            "dados_migrados": migrado,
        }, arquivo, ensure_ascii=False, indent=2)

    caminho_relatorio.write_text(relatorio, encoding="utf-8")

    print(relatorio)
    print("")
    print(f"Prévia gerada: {caminho_preview}")
    print(f"Relatório gerado: {caminho_relatorio}")
    return 1 if divergencias else 0


if __name__ == "__main__":
    raise SystemExit(main())
