-- Tabela: tbl_02(04)aProdCor
-- Registros: 5

CREATE TABLE tbl_02(04)aProdCor (
    ProdCorID INTEGER AUTOINCREMENT,
    ProdCor VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(04)aProdCor ADD PRIMARY KEY (ProdCorID);

