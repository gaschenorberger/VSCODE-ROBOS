/*POSTGRESQL*/

CREATE TABLE regime_tributario (
    id SERIAL PRIMARY KEY AUTO,
    nome_arquivo VARCHAR(100),
    ultima_atualizacao DATE,
    ultima_verificacao DATE
);

CREATE TABLE dados_abertos (
    idDadosAbertos SERIAL PRIMARY KEY,
    nome_arquivo VARCHAR(100),
    ultima_atualizacao DATE,
    ultima_verificacao DATE
);

CREATE TABLE portal_devedores (
    idPortalDevedores SERIAL PRIMARY KEY,
    nome_arquivo VARCHAR(100),
    ultima_atualizacao DATE,
    ultima_verificacao DATE
);

INSERT INTO regime_tributario (nome_arquivo, ultima_atualizacao, ultima_verificacao) VALUES 
('TESTE.ZIP', '01-01-2001', '01/10/2024');

INSERT INTO dados_abertos (nome_arquivo, ultima_atualizacao, ultima_verificacao) VALUES 
('TESTE.ZIP', '01-01-2001', '01/10/2024');

INSERT INTO portal_devedores (nome_arquivo, ultima_atualizacao, ultima_verificacao) VALUES 
('TESTE.ZIP', '01-01-2001', '01/10/2024');


SELECT 
nome_arquivo as "NOME DO ARQUIVO",
ultima_atualizacao as "ÚLTIMO ARQUIVO TRANSMITIDO",
ultima_verificacao as "ÚLTIMA VERIFICAÇÃO DE ATUALIZAÇÃO"
FROM regime_tributario;

SELECT 
nome_arquivo as "NOME DO ARQUIVO",
ultima_atualizacao as "ÚLTIMO ARQUIVO TRANSMITIDO",
ultima_verificacao as "ÚLTIMA VERIFICAÇÃO DE ATUALIZAÇÃO"
FROM dados_abertos;

SELECT 
nome_arquivo as "NOME DO ARQUIVO",
ultima_atualizacao as "ÚLTIMO ARQUIVO TRANSMITIDO",
ultima_verificacao as "ÚLTIMA VERIFICAÇÃO DE ATUALIZAÇÃO"
FROM portal_devedores;

ALTER TABLE regime_tributario
ALTER COLUMN nome_arquivo TYPE varchar (100); 

DELETE FROM regime_tributario;

TRUNCATE TABLE dados_abertos RESTART IDENTITY;  /*ZERAR ID*/

TRUNCATE TABLE regime_tributario RESTART IDENTITY;  /*ZERAR ID*/

TRUNCATE TABLE portal_devedores RESTART IDENTITY;	/*ZERAR ID*/