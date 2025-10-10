-- Tabela: tbl_05(1)bProdEstoqueOut
-- Registros: 2

CREATE TABLE tbl_05(1)bProdEstoqueOut (
    EstoqueOutID INTEGER AUTOINCREMENT,
    DataSaida DATETIME,
    EstoqueInIDfk INTEGER,
    QtdEmbsOut INTEGER
);

ALTER TABLE tbl_05(1)bProdEstoqueOut ADD PRIMARY KEY (EstoqueOutID);

