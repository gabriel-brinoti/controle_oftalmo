# Controle Oftalmo — Supabase/PostgreSQL

Sistema Flask para controle de qualidade, validade, estoque, histórico, relatórios, dashboard e segurança.

## Banco de dados

Esta versão usa PostgreSQL via Supabase.

Não usa mais SQLite/database.db.

## Configuração local

Crie um arquivo `.env` na raiz do projeto:

```env
DATABASE_URL=sua_url_do_supabase
SECRET_KEY=sua_chave_secreta
FLASK_ENV=development
```

## Instalar dependências

```bash
pip install -r requirements.txt
```

## Rodar localmente

```bash
python app.py
```

Acesse:

```text
http://localhost:5000
```

## Login inicial

```text
usuário: admin
senha: admin123
```

Troque a senha na tela Segurança após o primeiro acesso.

## Render

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
gunicorn app:app
```

Variáveis de ambiente no Render:

```text
DATABASE_URL=sua_url_do_supabase
SECRET_KEY=sua_chave_secreta
FLASK_ENV=production
```

## Observações

- O sistema cria as tabelas automaticamente ao iniciar.
- Para backup completo do banco, use o painel do Supabase.
- Para backup operacional, use os relatórios Excel/PDF do sistema.


## Controle de licença mensal

Esta versão possui controle de licença SaaS:

- Plano visível no dashboard
- Licença ativa: uso normal
- Licença vencida dentro de 5 dias: modo consulta
- Após 5 dias vencida: bloqueio total

A tabela `licencas` é criada automaticamente no PostgreSQL.


## Etapa 6.1 - Estoques separados

Inclui:
- Almoxarifado
- Farmácia Satélite
- Campo `tipo_estoque` em produtos
- Filtro de histórico por estoque
- Movimentações: retirada, devolução e reposição
- Alerta de vencimento alterado para 5 dias


## Etapa 6.2 - Relatórios separados por estoque

Inclui:
- Filtro de estoque nos relatórios
- Exportação PDF/Excel por Almoxarifado, Farmácia Satélite ou Geral
- Dashboard com indicadores separados por estoque


## Etapa 6.3 - Ordem de compra automática

Inclui:
- Página `/ordem_compra`
- Sugestão automática de compra
- Separação por Almoxarifado e Farmácia Satélite
- Exportação PDF/Excel
- Card no dashboard de itens aguardando compra


## Etapa 7 - Alertas automáticos

Inclui:
- Configuração de alertas por e-mail e WhatsApp
- Prévia da mensagem
- Envio manual de teste
- Histórico de envios
- Preparação para rotina diária automática

Variáveis opcionais para e-mail:
SMTP_HOST
SMTP_PORT
SMTP_USER
SMTP_PASSWORD
SMTP_FROM

Variáveis opcionais para WhatsApp:
WHATSAPP_WEBHOOK_URL
WHATSAPP_TOKEN


## Etapa 8 - Transferência inteligente entre estoques

Inclui:
- Botão transferir no produto
- Transferência Almoxarifado ↔ Farmácia Satélite
- Histórico com transferência de saída e entrada
- Ordem de compra sugere transferência antes de compra
- Transferências recentes no dashboard


## Redesign base premium

Inclui:
- Sidebar lateral escura
- Header superior moderno
- Área central com melhor espaçamento
- Navegação estilo SaaS
- Compatibilidade com páginas atuais


## Dashboard Premium

Inclui:
- Hero visual moderno
- Cards métricos premium
- Estoques separados em cards
- Produtos críticos em painel organizado
- Ranking de consumo
- Previsões e alertas inteligentes
- Transferências recentes em layout premium


## Produtos Premium

Inclui:
- Produtos em cards operacionais modernos
- Ações rápidas separadas das movimentações
- Indicador visual de estoque
- Filtros premium
- Status visual por produto
- Layout responsivo


## Relatórios Premium

Inclui:
- Painel analítico premium
- Cards métricos
- Produtos críticos organizados
- Resumo operacional
- Ranking de consumo
- Produtos vencendo
- Últimas movimentações em layout premium


## Alertas Premium

Inclui:
- Central premium de notificações
- Cards de status para e-mail, WhatsApp e horário
- Configuração visual dos canais
- Tipos de alerta organizados
- Prévia premium da mensagem
- Histórico de envios com status visual


## Ordem de Compra Premium

Inclui:
- Central logística premium
- Cards inteligentes de compra
- Métricas de reposição
- Separação por estoque
- Sugestão visual de compra/transferência
- Exportação PDF/Excel mantida


## Central Visual de Organização do Estoque / Categorias Premium

Inclui:
- Categorias em cards premium
- Indicadores por categoria
- Total de produtos, estoque, vencimentos e estoque baixo
- Ações premium para editar, visualizar produtos e excluir
- Formulário de categoria modernizado


## Código de Barras Premium

Inclui:
- Campo codigo_barras em produtos
- Cadastro/edição com código de barras
- Busca de produtos por código
- Tela /codigo_barras para bipagem
- Movimentações rápidas por produto bipado


## Histórico Premium

Inclui:
- Central de auditoria operacional
- Filtros premium
- Cards de resumo
- Timeline de movimentações
- Badges por tipo de movimentação
- Destaque para transferências e observações


## Segurança Premium

Inclui:
- Painel premium de segurança
- Alteração de senha modernizada
- Cards de status operacional
- Checklist de boas práticas
- Painel de controles de segurança


## Identidade Visual HO

Inclui:
- Paleta inspirada no Hospital Oftalmológico de Franca
- Verde petróleo, verde água, creme e branco hospitalar
- Sidebar institucional
- Botões, cards, inputs, tabelas e badges com a nova identidade
- Sistema premium com aparência personalizada para a clínica


## Refinamento Premium Final HO

Inclui:
- Correção de sidebar cortando em zoom 100%
- Logo oficial HO aplicado no sistema
- Layout responsivo refinado
- Ajustes para notebook, monitor e telas menores
- Prevenção de cards sobrepostos
- Padronização visual final da identidade HO


## Scanner Operacional — Etapa 1

Inclui:
- Tela /codigo_barras preparada para leitor USB
- Foco automático no campo de bipagem
- Busca automática via Enter do scanner
- Produto encontrado com painel visual
- Movimentações rápidas após bipagem
- Mensagem para código não encontrado


## Scanner Operacional — Etapa 2

Inclui:
- Modo Consulta e Modo Contínuo
- Baixa rápida automática por bipagem
- Quantidade padrão por bip
- Observação padrão por movimentação
- Retorno automático para próxima leitura


## Scanner Operacional — Etapa 3

Inclui:
- Painel operacional ao vivo
- Últimos produtos bipados
- Total de bipagens recentes
- Total de unidades baixadas
- Indicador de estoque baixo
- Status visual da última leitura


## Scanner Operacional — Etapa 4

Inclui:
- Tabela auditoria_scanner
- Registro de sucesso e erro no scanner
- Usuário, IP, navegador e horário
- Painel de rastreabilidade do scanner
- Auditoria operacional visual


## Etapa 5 — Dashboard Hospitalar em Tempo Real

Inclui:
- Painel central em tempo real
- Movimentações recentes
- Últimas bipagens
- Produtos críticos
- Farmácia vs Almoxarifado
- Produtos mais movimentados
- Auto atualização do painel


## Controle de Acesso por Perfil

Foram adicionados 2 perfis:

### Sala Central / Admin
- Login com usuário padrão/admin já existente
- Acesso geral ao sistema
- Dashboard, tempo real, relatórios, categorias, alertas, configurações, ordem de compra e estoque

### Estoque / Almoxarifado / Farmácia
- Login usando usuário: estoque, almoxarifado ou farmacia
- Acesso operacional limitado
- Estoque, Scanner e Histórico
- Bloqueia relatórios, configurações, alertas, categorias e áreas administrativas

Observação:
A senha segue a mesma validação já existente no sistema. Se quiser senha separada por perfil, a próxima etapa é criar tabela de usuários.


## Usuários Reais com Permissões Individuais

Inclui:
- Tabela usuarios no PostgreSQL
- Login por usuário e senha própria
- Perfil Admin / Sala Central
- Perfil Estoque / Farmácia / Almoxarifado
- Tela de gerenciamento de usuários
- Ativar/desativar usuários
- Trocar senha do usuário
- Auditoria usando nome do usuário logado

Usuários iniciais criados automaticamente:
- admin / admin123
- estoque / admin123


## Auditoria Completa por Usuário

Inclui:
- Colunas usuario_id e usuario_nome em movimentacoes
- Registro automático do usuário logado em cada movimentação
- Filtro de histórico por usuário
- Relatório de movimentações com usuário
- Página /auditoria_usuarios para Sala Central
- Dashboard tempo real exibindo usuário nas ações


## Licenciamento Local Mensal

Inclui:
- Tela /licenciamento para Sala Central/Admin
- Chave de licença local
- Data de vencimento mensal
- Status ativo/bloqueado
- Dias de carência
- Renovação manual
- Bloqueio leve: licença vencida entra em modo consulta
- Bloqueio total quando status for bloqueado ou carência expirar


## Alertas automáticos por WhatsApp

Inclui:
- Envio de alertas críticos por WhatsApp
- Compatível com UltraMsg ou webhook genérico
- Botão de envio manual em Configurações de Alertas
- Controle anti-spam por intervalo em minutos
- Histórico de envio em historico_alertas

Configuração recomendada no Render (.env):

ULTRAMSG_INSTANCE_ID=instance000000
ULTRAMSG_TOKEN=seu_token
WHATSAPP_PHONE=5516999999999

Na tela de Alertas:
- Ativar WhatsApp
- Informar telefone com DDI/DDD
- Configurar intervalo mínimo entre envios
