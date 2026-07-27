# Backup automatico com Neon

Implementado:
- Backup JSON local
- Snapshot automatico no Neon
- Botao de teste no painel oculto `/painel_master_backup_9182`
- Botao para baixar backups JSON locais

Variaveis necessarias no Render:
- `NEON_API_KEY`
- `NEON_PROJECT_ID`
- `NEON_BRANCH_ID`
- `NEON_SNAPSHOT_RETENTION_DAYS=30`

Observacao:
A chave `NEON_API_KEY` nunca deve ir para o Git.
