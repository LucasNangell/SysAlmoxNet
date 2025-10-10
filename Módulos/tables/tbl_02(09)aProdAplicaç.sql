-- Tabela: tbl_02(09)aProdAplicaç
-- Registros: 3

CREATE TABLE tbl_02(09)aProdAplicaç (
    ProdAplicaçaoID INTEGER AUTOINCREMENT,
    ProdAplicaçaoDescriç VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(09)aProdAplicaç ADD PRIMARY KEY (ProdAplicaçaoID);

