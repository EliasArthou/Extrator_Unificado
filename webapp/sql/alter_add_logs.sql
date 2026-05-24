-- Adicionar colunas de log na tabela execucoes
ALTER TABLE execucoes ADD COLUMN IF NOT EXISTS log_texto TEXT DEFAULT '';
ALTER TABLE execucoes ADD COLUMN IF NOT EXISTS erros_detalhados TEXT DEFAULT '[]';

-- O GRANT por tabela ja cobre novas colunas, mas garantindo:
GRANT SELECT, INSERT, UPDATE, DELETE ON execucoes TO webapp_app;
