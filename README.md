# Controle Oftalmo

Sistema web em Flask para controle de qualidade, validade, estoque, histórico de movimentações, relatórios e backup.

## Funcionalidades

- Login administrativo
- Cadastro de produtos
- Cadastro de categorias
- Controle de validade original
- Controle de validade após abertura
- Alertas de vencimento
- Controle de estoque
- Histórico de movimentações
- Relatórios Excel e PDF
- Dashboard inteligente
- Backup e restauração
- Tela de segurança

## Rodar localmente

```bash
pip install -r requirements.txt
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

Troque a senha na tela **Segurança** após o primeiro acesso.

## Deploy inicial no Render

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
gunicorn app:app
```

Ou use o `Procfile`:

```text
web: gunicorn app:app
```

## Variáveis de ambiente recomendadas

```text
SECRET_KEY=uma_chave_segura_aqui
```

## Observação importante

Esta versão ainda usa SQLite. Para uso comercial online, o próximo passo recomendado é migrar para PostgreSQL para evitar perda de dados em hospedagens gratuitas e melhorar a persistência.
