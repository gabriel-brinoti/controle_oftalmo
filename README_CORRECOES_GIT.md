# Correções aplicadas antes do Git

- Corrigida ordem de criação da tabela movimentacoes antes dos ALTER.
- Corrigido INSERT inicial de configuracoes_alerta com intervalo_minutos.
- Corrigido INSERT de transferência de produto para outro estoque.
- Dashboard principal protegido para admin.
- Exclusão de produto protegida por perfil operacional.
- Backup automático com APScheduler protegido contra duplicação no debug.
- Login aceita usuario/login/username e senha/password.
- Templates enviados foram organizados com nomes corretos, incluindo relatorios.html.

## Checklist interno

- movimentacoes_create_before_alter: OK
- config_insert_values_9: OK
- transfer_insert_12_values: OK
- dashboard_admin: OK
- exclude_product_estoque: OK
- scheduler_called: OK
