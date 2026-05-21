# Backup automático com Supabase Storage

Implementado:
- Backup JSON local
- Upload automático para Supabase Storage
- Pasta no bucket: automaticos/
- Botão de teste no painel oculto /painel_master_backup_9182

Variáveis necessárias no Render:
- SUPABASE_URL
- SUPABASE_SERVICE_ROLE_KEY
- SUPABASE_BACKUP_BUCKET=backups

Observação:
A chave service_role nunca deve ir para o Git.
