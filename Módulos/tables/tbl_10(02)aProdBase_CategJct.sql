-- Tabela: tbl_10(02)aProdBase_CategJct
-- Registros: 10

CREATE TABLE tbl_10(02)aProdBase_CategJct (
    ProdBase_Categ_jct_ID INTEGER AUTOINCREMENT,
    ProdutoFullIDfk INTEGER,
    CategIDfk INTEGER
);

ALTER TABLE tbl_10(02)aProdBase_CategJct ADD PRIMARY KEY (ProdBase_Categ_jct_ID);

