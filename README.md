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
