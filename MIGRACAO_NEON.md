# Migracao para Neon

Este pacote prepara um banco Neon novo usando o backup atual do sistema.

## Arquivos

- `neon_schema.sql`: cria as tabelas novas no Neon.
- `validar_migracao_estoque.py`: valida o backup antes de importar.
- `importar_backup_para_neon.py`: cria a estrutura e importa os dados para o Neon.
- `testar_migracao_offline.py`: testa as telas novas sem conectar no banco.

## Passo seguro

1. Crie um projeto/banco novo no Neon.
2. Copie a connection string do Neon.
3. Rode primeiro em modo teste:

```powershell
$env:NEON_DATABASE_URL="cole_a_url_do_neon_aqui"
python importar_backup_para_neon.py "C:\Users\gbryn\Downloads\Backup_Completo_20-07-2026_19h17m.json" --dry-run
```

4. Se o teste passar, rode a importacao real:

```powershell
$env:NEON_DATABASE_URL="cole_a_url_do_neon_aqui"
python importar_backup_para_neon.py "C:\Users\gbryn\Downloads\Backup_Completo_20-07-2026_19h17m.json"
```

## Trava de seguranca

O importador nao importa em banco que ja tenha dados.

Use `--reset` somente se o banco Neon for novo/de teste e voce aceitar apagar o conteudo:

```powershell
python importar_backup_para_neon.py "C:\Users\gbryn\Downloads\Backup_Completo_20-07-2026_19h17m.json" --reset
```

## Variaveis no Render

Para o sistema online usar Neon como banco:

- `DATABASE_URL`: connection string do Neon.
- `SECRET_KEY`: manter a chave atual.
- `FLASK_ENV`: manter como esta.

Para backup automatico por snapshot Neon:

- `NEON_API_KEY`: chave de API criada no painel do Neon.
- `NEON_PROJECT_ID`: id do projeto Neon.
- `NEON_BRANCH_ID`: id da branch, normalmente a branch `production`.
- `NEON_SNAPSHOT_RETENTION_DAYS`: quantidade de dias para manter snapshots, sugestao `30`.

Depois disso, as variaveis antigas do Supabase podem ser removidas do Render:

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `SUPABASE_BACKUP_BUCKET`

## O que sera migrado

- Categorias
- Usuarios
- Produtos principais
- Configuracao por estoque
- Lotes
- Movimentacoes com vinculo ao lote
- Licencas, configuracoes de alerta e ordens de compra quando existirem no backup

## Conferencia esperada para o backup de 20/07/2026

- 441 registros atuais em produtos
- 398 produtos principais
- 418 produtos por estoque
- 441 lotes
- 52 movimentacoes
- 0 divergencias de quantidade
