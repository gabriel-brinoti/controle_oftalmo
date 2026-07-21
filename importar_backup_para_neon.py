import argparse
import os
from pathlib import Path

import psycopg2
from psycopg2.extras import RealDictCursor

from validar_migracao_estoque import carregar_backup, migrar_backup


TABELAS_IMPORTADAS = [
    "movimentacoes",
    "lotes",
    "produto_estoques",
    "produtos",
    "categorias",
    "usuarios",
    "admin",
    "licencas",
    "configuracoes_alerta",
    "ordens_compra",
]


def conectar(database_url):
    return psycopg2.connect(database_url, cursor_factory=RealDictCursor, connect_timeout=20)


def executar_schema(cursor, caminho_schema):
    cursor.execute(Path(caminho_schema).read_text(encoding="utf-8"))


def tabela_tem_dados(cursor, tabela):
    cursor.execute(f"SELECT COUNT(*) AS total FROM {tabela}")
    return cursor.fetchone()["total"] > 0


def banco_tem_dados(cursor):
    tabelas = ["categorias", "produtos", "produto_estoques", "lotes", "movimentacoes", "usuarios"]
    return {tabela: tabela_tem_dados(cursor, tabela) for tabela in tabelas}


def limpar_tabelas(cursor):
    cursor.execute("""
        TRUNCATE TABLE
            movimentacoes,
            lotes,
            produto_estoques,
            produtos,
            categorias,
            usuarios,
            admin,
            licencas,
            configuracoes_alerta,
            historico_alertas,
            ordens_compra
        RESTART IDENTITY CASCADE
    """)


def inserir_categorias(cursor, categorias):
    for categoria in categorias:
        cursor.execute("""
            INSERT INTO categorias (id, nome, descricao, ativo, criado_em)
            VALUES (%s, %s, %s, %s, COALESCE(%s::timestamp, NOW()))
            ON CONFLICT (id) DO NOTHING
        """, (
            categoria.get("id"),
            categoria.get("nome"),
            categoria.get("descricao"),
            categoria.get("ativo", True),
            categoria.get("criado_em"),
        ))


def inserir_usuarios(cursor, usuarios):
    for usuario in usuarios:
        cursor.execute("""
            INSERT INTO usuarios (id, nome, usuario, senha_hash, perfil, ativo, criado_em)
            VALUES (%s, %s, %s, %s, %s, %s, COALESCE(%s::timestamp, NOW()))
            ON CONFLICT (id) DO NOTHING
        """, (
            usuario.get("id"),
            usuario.get("nome"),
            usuario.get("usuario"),
            usuario.get("senha_hash"),
            usuario.get("perfil", "estoque"),
            usuario.get("ativo", True),
            usuario.get("criado_em"),
        ))


def inserir_admin_de_usuarios(cursor, usuarios):
    for usuario in usuarios:
        if usuario.get("perfil") == "admin":
            cursor.execute("""
                INSERT INTO admin (usuario, senha_hash)
                VALUES (%s, %s)
                ON CONFLICT (usuario) DO NOTHING
            """, (usuario.get("usuario"), usuario.get("senha_hash")))


def inserir_produtos(cursor, produtos):
    for produto in produtos:
        cursor.execute("""
            INSERT INTO produtos (
                id, nome, categoria_id, codigo_barras, fabricante,
                unidade_medida, validade_apos_aberto_dias,
                observacoes, ativo, criado_em
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, COALESCE(%s::timestamp, NOW()))
            ON CONFLICT (id) DO NOTHING
        """, (
            produto.get("id"),
            produto.get("nome"),
            produto.get("categoria_id"),
            produto.get("codigo_barras"),
            produto.get("fabricante"),
            produto.get("unidade_medida"),
            produto.get("validade_apos_aberto_dias"),
            produto.get("observacoes"),
            produto.get("ativo", True),
            produto.get("criado_em"),
        ))


def inserir_produto_estoques(cursor, produto_estoques):
    for item in produto_estoques:
        cursor.execute("""
            INSERT INTO produto_estoques (
                id, produto_id, tipo_estoque, estoque_padrao, limite_alerta, ativo
            ) VALUES (%s, %s, %s, %s, %s, %s)
            ON CONFLICT (id) DO NOTHING
        """, (
            item.get("id"),
            item.get("produto_id"),
            item.get("tipo_estoque"),
            item.get("estoque_padrao"),
            item.get("limite_alerta"),
            item.get("ativo", True),
        ))


def inserir_lotes(cursor, lotes):
    for lote in lotes:
        cursor.execute("""
            INSERT INTO lotes (
                id, produto_estoque_id, numero_lote, data_vencimento,
                data_abertura, quantidade_inicial, quantidade_atual,
                data_entrada, observacoes, ativo
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, COALESCE(%s::timestamp, NOW()), %s, %s)
            ON CONFLICT (id) DO NOTHING
        """, (
            lote.get("id"),
            lote.get("produto_estoque_id"),
            lote.get("numero_lote"),
            lote.get("data_vencimento"),
            lote.get("data_abertura"),
            lote.get("quantidade_inicial"),
            lote.get("quantidade_atual"),
            lote.get("data_entrada"),
            lote.get("observacoes"),
            lote.get("ativo", True),
        ))


def inserir_movimentacoes(cursor, movimentacoes):
    for mov in movimentacoes:
        cursor.execute("""
            INSERT INTO movimentacoes (
                id, lote_id, produto_id, produto_nome, tipo_movimentacao,
                quantidade, estoque_anterior, estoque_atual,
                quantidade_anterior, quantidade_posterior,
                observacao, motivo, tipo_estoque, origem, destino,
                estoque_origem, estoque_destino, usuario_id, usuario_nome,
                data_movimentacao, criado_em
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, COALESCE(%s::timestamp, NOW()), COALESCE(%s::timestamp, NOW()))
            ON CONFLICT (id) DO NOTHING
        """, (
            mov.get("id"),
            mov.get("lote_id"),
            mov.get("produto_id"),
            mov.get("produto_nome"),
            mov.get("tipo_movimentacao"),
            mov.get("quantidade"),
            mov.get("estoque_anterior") or 0,
            mov.get("estoque_atual") or 0,
            mov.get("quantidade_anterior"),
            mov.get("quantidade_posterior"),
            mov.get("observacao"),
            mov.get("motivo"),
            mov.get("tipo_estoque"),
            mov.get("origem"),
            mov.get("destino"),
            mov.get("estoque_origem"),
            mov.get("estoque_destino"),
            mov.get("usuario_id"),
            mov.get("usuario_nome"),
            mov.get("data_movimentacao"),
            mov.get("criado_em") or mov.get("data_movimentacao"),
        ))


def inserir_licencas(cursor, licencas):
    for licenca in licencas:
        cursor.execute("""
            INSERT INTO licencas (
                id, empresa, plano, status, data_vencimento,
                dias_carencia, chave_licenca, observacoes, atualizado_em, criado_em
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, COALESCE(%s::timestamp, NOW()), COALESCE(%s::timestamp, NOW()))
            ON CONFLICT (id) DO NOTHING
        """, (
            licenca.get("id"),
            licenca.get("empresa"),
            licenca.get("plano"),
            licenca.get("status", "ativo"),
            licenca.get("data_vencimento"),
            licenca.get("dias_carencia", 5),
            licenca.get("chave_licenca"),
            licenca.get("observacoes"),
            licenca.get("atualizado_em"),
            licenca.get("criado_em"),
        ))


def inserir_configuracoes_alerta(cursor, configuracoes):
    for config in configuracoes:
        cursor.execute("""
            INSERT INTO configuracoes_alerta (
                id, email_destino, telefone_whatsapp, usar_email, usar_whatsapp,
                alertar_vencimentos, alertar_estoque, alertar_ordem_compra,
                hora_envio, intervalo_minutos, ultimo_envio_whatsapp, criado_em
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, COALESCE(%s::timestamp, NOW()))
            ON CONFLICT (id) DO NOTHING
        """, (
            config.get("id"),
            config.get("email_destino"),
            config.get("telefone_whatsapp"),
            config.get("usar_email", False),
            config.get("usar_whatsapp", False),
            config.get("alertar_vencimentos", True),
            config.get("alertar_estoque", True),
            config.get("alertar_ordem_compra", True),
            config.get("hora_envio", "08:00"),
            config.get("intervalo_minutos", 720),
            config.get("ultimo_envio_whatsapp"),
            config.get("criado_em"),
        ))


def inserir_ordens_compra(cursor, ordens):
    for ordem in ordens:
        cursor.execute("""
            INSERT INTO ordens_compra (
                id, produto_nome, categoria, quantidade_atual,
                quantidade_sugerida, status, observacoes, criado_em
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, COALESCE(%s::timestamp, NOW()))
            ON CONFLICT (id) DO NOTHING
        """, (
            ordem.get("id"),
            ordem.get("produto_nome"),
            ordem.get("categoria"),
            ordem.get("quantidade_atual") or 0,
            ordem.get("quantidade_sugerida") or 0,
            ordem.get("status", "pendente"),
            ordem.get("observacoes"),
            ordem.get("criado_em"),
        ))


def ajustar_sequences(cursor):
    sequences = {
        "categorias": "categorias_id_seq",
        "produtos": "produtos_id_seq",
        "produto_estoques": "produto_estoques_id_seq",
        "lotes": "lotes_id_seq",
        "movimentacoes": "movimentacoes_id_seq",
        "usuarios": "usuarios_id_seq",
        "licencas": "licencas_id_seq",
        "configuracoes_alerta": "configuracoes_alerta_id_seq",
        "ordens_compra": "ordens_compra_id_seq",
    }
    for tabela, sequence in sequences.items():
        cursor.execute(f"SELECT setval(%s, COALESCE((SELECT MAX(id) FROM {tabela}), 1))", (sequence,))


def importar(database_url, backup_path, reset=False, dry_run=False):
    backup = carregar_backup(backup_path)
    dados_originais = backup.get("dados", {})
    migrado = migrar_backup(backup)

    conn = conectar(database_url)
    cursor = conn.cursor()

    try:
        executar_schema(cursor, "neon_schema.sql")
        dados_existentes = banco_tem_dados(cursor)

        if any(dados_existentes.values()) and not reset:
            conn.rollback()
            raise RuntimeError(
                "O banco Neon já tem dados. Use --reset apenas se tiver certeza que é um banco de teste ou recém-criado."
            )

        if reset:
            limpar_tabelas(cursor)

        inserir_categorias(cursor, dados_originais.get("categorias", []))
        inserir_usuarios(cursor, dados_originais.get("usuarios", []))
        inserir_admin_de_usuarios(cursor, dados_originais.get("usuarios", []))
        inserir_produtos(cursor, migrado["produtos"])
        inserir_produto_estoques(cursor, migrado["produto_estoques"])
        inserir_lotes(cursor, migrado["lotes"])
        inserir_movimentacoes(cursor, migrado["movimentacoes"])
        inserir_licencas(cursor, dados_originais.get("licencas", []))
        inserir_configuracoes_alerta(cursor, dados_originais.get("configuracoes_alerta", []))
        inserir_ordens_compra(cursor, dados_originais.get("ordens_compra", []))
        ajustar_sequences(cursor)

        if dry_run:
            conn.rollback()
        else:
            conn.commit()

        return {
            "dry_run": dry_run,
            "categorias": len(dados_originais.get("categorias", [])),
            "usuarios": len(dados_originais.get("usuarios", [])),
            "produtos": len(migrado["produtos"]),
            "produto_estoques": len(migrado["produto_estoques"]),
            "lotes": len(migrado["lotes"]),
            "movimentacoes": len(migrado["movimentacoes"]),
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def main():
    parser = argparse.ArgumentParser(description="Importa backup atual para um banco Neon novo.")
    parser.add_argument("backup", help="Caminho do Backup_Completo_*.json")
    parser.add_argument("--database-url", default=os.environ.get("NEON_DATABASE_URL") or os.environ.get("DATABASE_URL"))
    parser.add_argument("--reset", action="store_true", help="Apaga dados existentes antes de importar.")
    parser.add_argument("--dry-run", action="store_true", help="Testa a importação e desfaz no final.")
    args = parser.parse_args()

    if not args.database_url:
        raise SystemExit("Informe --database-url ou defina NEON_DATABASE_URL.")

    resultado = importar(args.database_url, args.backup, reset=args.reset, dry_run=args.dry_run)
    print("Importação Neon concluída." if not args.dry_run else "Dry-run Neon concluído sem gravar.")
    for chave, valor in resultado.items():
        print(f"{chave}: {valor}")


if __name__ == "__main__":
    main()
