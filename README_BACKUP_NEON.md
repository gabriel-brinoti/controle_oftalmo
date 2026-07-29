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

Variaveis opcionais de performance:
- `DB_POOL_MIN=1`
- `DB_POOL_MAX=5`
- `DB_CONNECT_TIMEOUT=10`

Para o plano atual, comece com `DB_POOL_MAX=5`. Se o Render estiver usando mais de
um worker, lembre que cada worker cria seu proprio pool.

Observacao:
A chave `NEON_API_KEY` nunca deve ir para o Git.

