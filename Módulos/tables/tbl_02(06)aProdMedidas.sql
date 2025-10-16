-- Tabela: tbl_02(06)aProdMedidas
-- Registros: 4

CREATE TABLE tbl_02(06)aProdMedidas (
    ProdMedidaID INTEGER AUTOINCREMENT,
    ProdMedida VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(06)aProdMedidas ADD PRIMARY KEY (ProdMedidaID);

