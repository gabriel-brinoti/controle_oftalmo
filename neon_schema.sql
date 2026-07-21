BEGIN;

CREATE TABLE IF NOT EXISTS admin (
    id SERIAL PRIMARY KEY,
    usuario TEXT NOT NULL UNIQUE,
    senha_hash TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS usuarios (
    id SERIAL PRIMARY KEY,
    nome TEXT NOT NULL,
    usuario TEXT UNIQUE NOT NULL,
    senha_hash TEXT NOT NULL,
    perfil TEXT NOT NULL DEFAULT 'estoque',
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS categorias (
    id SERIAL PRIMARY KEY,
    nome TEXT NOT NULL UNIQUE,
    descricao TEXT,
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    criado_em TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS produtos (
    id SERIAL PRIMARY KEY,
    nome TEXT NOT NULL,
    categoria_id INTEGER NOT NULL REFERENCES categorias(id),
    codigo_barras TEXT,
    fabricante TEXT,
    unidade_medida TEXT,
    validade_apos_aberto_dias INTEGER,
    observacoes TEXT,
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    criado_em TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS produto_estoques (
    id SERIAL PRIMARY KEY,
    produto_id INTEGER NOT NULL REFERENCES produtos(id) ON DELETE CASCADE,
    tipo_estoque TEXT NOT NULL,
    estoque_padrao INTEGER NOT NULL DEFAULT 0,
    limite_alerta INTEGER NOT NULL DEFAULT 0,
    ativo BOOLEAN NOT NULL DEFAULT TRUE,
    criado_em TIMESTAMP NOT NULL DEFAULT NOW(),
    UNIQUE (produto_id, tipo_estoque)
);

CREATE TABLE IF NOT EXISTS lotes (
    id SERIAL PRIMARY KEY,
    produto_estoque_id INTEGER NOT NULL REFERENCES produto_estoques(id) ON DELETE CASCADE,
    numero_lote TEXT,
    data_vencimento DATE,
    data_abertura DATE,
    quantidade_inicial INTEGER NOT NULL DEFAULT 0,
    quantidade_atual INTEGER NOT NULL DEFAULT 0,
    data_entrada TIMESTAMP NOT NULL DEFAULT NOW(),
    observacoes TEXT,
    ativo BOOLEAN NOT NULL DEFAULT TRUE
);

CREATE TABLE IF NOT EXISTS movimentacoes (
    id SERIAL PRIMARY KEY,
    lote_id INTEGER REFERENCES lotes(id),
    produto_id INTEGER REFERENCES produtos(id),
    produto_nome TEXT NOT NULL,
    tipo_movimentacao TEXT NOT NULL,
    quantidade INTEGER NOT NULL,
    estoque_anterior INTEGER NOT NULL DEFAULT 0,
    estoque_atual INTEGER NOT NULL DEFAULT 0,
    quantidade_anterior INTEGER,
    quantidade_posterior INTEGER,
    observacao TEXT,
    motivo TEXT,
    tipo_estoque TEXT DEFAULT 'almoxarifado',
    origem TEXT,
    destino TEXT,
    estoque_origem TEXT,
    estoque_destino TEXT,
    usuario_id INTEGER,
    usuario_nome TEXT,
    data_movimentacao TIMESTAMP NOT NULL DEFAULT NOW(),
    criado_em TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS produtos_abertos (
    id SERIAL PRIMARY KEY,
    lote_id INTEGER REFERENCES lotes(id),
    produto_id INTEGER REFERENCES produtos(id),
    produto_nome TEXT NOT NULL,
    tipo_estoque TEXT NOT NULL,
    numero_lote TEXT,
    data_abertura DATE NOT NULL,
    validade_apos_aberto_dias INTEGER NOT NULL,
    vencimento_apos_aberto DATE NOT NULL,
    quantidade INTEGER NOT NULL DEFAULT 1,
    usuario_id INTEGER,
    usuario_nome TEXT,
    observacoes TEXT,
    criado_em TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS licencas (
    id SERIAL PRIMARY KEY,
    empresa TEXT NOT NULL,
    plano TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'ativo',
    data_vencimento DATE NOT NULL,
    dias_carencia INTEGER NOT NULL DEFAULT 5,
    chave_licenca TEXT,
    observacoes TEXT,
    atualizado_em TIMESTAMP DEFAULT NOW(),
    criado_em TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS configuracoes_alerta (
    id SERIAL PRIMARY KEY,
    email_destino TEXT,
    telefone_whatsapp TEXT,
    usar_email BOOLEAN NOT NULL DEFAULT FALSE,
    usar_whatsapp BOOLEAN NOT NULL DEFAULT FALSE,
    alertar_vencimentos BOOLEAN NOT NULL DEFAULT TRUE,
    alertar_estoque BOOLEAN NOT NULL DEFAULT TRUE,
    alertar_ordem_compra BOOLEAN NOT NULL DEFAULT TRUE,
    hora_envio TEXT NOT NULL DEFAULT '08:00',
    intervalo_minutos INTEGER NOT NULL DEFAULT 720,
    ultimo_envio_whatsapp TIMESTAMP,
    criado_em TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS historico_alertas (
    id SERIAL PRIMARY KEY,
    tipo_alerta TEXT NOT NULL,
    canal TEXT NOT NULL,
    destino TEXT,
    conteudo TEXT,
    status TEXT NOT NULL,
    erro TEXT,
    criado_em TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS ordens_compra (
    id SERIAL PRIMARY KEY,
    produto_nome TEXT NOT NULL,
    categoria TEXT,
    quantidade_atual INTEGER NOT NULL DEFAULT 0,
    quantidade_sugerida INTEGER NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'pendente',
    observacoes TEXT,
    criado_em TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_produtos_nome ON produtos (nome);
CREATE INDEX IF NOT EXISTS idx_produto_estoques_produto ON produto_estoques (produto_id);
CREATE INDEX IF NOT EXISTS idx_lotes_produto_estoque ON lotes (produto_estoque_id);
CREATE INDEX IF NOT EXISTS idx_lotes_vencimento ON lotes (data_vencimento);
CREATE INDEX IF NOT EXISTS idx_movimentacoes_lote ON movimentacoes (lote_id);
CREATE INDEX IF NOT EXISTS idx_movimentacoes_produto ON movimentacoes (produto_id);
CREATE INDEX IF NOT EXISTS idx_movimentacoes_data ON movimentacoes (data_movimentacao);

COMMIT;
