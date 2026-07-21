# Migração para Neon

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

4. Se o teste passar, rode a importação real:

```powershell
$env:NEON_DATABASE_URL="cole_a_url_do_neon_aqui"
python importar_backup_para_neon.py "C:\Users\gbryn\Downloads\Backup_Completo_20-07-2026_19h17m.json"
```

## Trava de segurança

O importador não importa em banco que já tenha dados.

Use `--reset` somente se o banco Neon for novo/de teste e você aceitar apagar o conteúdo:

```powershell
python importar_backup_para_neon.py "C:\Users\gbryn\Downloads\Backup_Completo_20-07-2026_19h17m.json" --reset
```

## O que será migrado

- Categorias
- Usuários
- Produtos principais
- Configuração por estoque
- Lotes
- Movimentações com vínculo ao lote
- Licenças, configurações de alerta e ordens de compra quando existirem no backup

## Conferência esperada para o backup de 20/07/2026

- 441 registros atuais em produtos
- 398 produtos principais
- 418 produtos por estoque
- 441 lotes
- 52 movimentações
- 0 divergências de quantidade
