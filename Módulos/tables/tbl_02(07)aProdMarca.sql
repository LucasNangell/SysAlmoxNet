-- Tabela: tbl_02(07)aProdMarca
-- Registros: 10

CREATE TABLE tbl_02(07)aProdMarca (
    ProdMarcaID INTEGER AUTOINCREMENT,
    ProdMarca VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(07)aProdMarca ADD PRIMARY KEY (ProdMarcaID);

