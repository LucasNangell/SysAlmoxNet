-- Tabela: tbl_02(10)aProdUnMedida
-- Registros: 8

CREATE TABLE tbl_02(10)aProdUnMedida (
    ProdUnMedidaID INTEGER AUTOINCREMENT,
    ProdUnMedidaDescriç VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(10)aProdUnMedida ADD PRIMARY KEY (ProdUnMedidaID);

