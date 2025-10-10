-- Tabela: tbl_02(08)aProdCateg
-- Registros: 6

CREATE TABLE tbl_02(08)aProdCateg (
    ProdCategID INTEGER AUTOINCREMENT,
    ProdCateg VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(08)aProdCateg ADD PRIMARY KEY (ProdCategID);

