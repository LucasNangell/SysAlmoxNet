-- Tabela: tbl_10(02)bProdBase_MarcaJct
-- Registros: 10

CREATE TABLE tbl_10(02)bProdBase_MarcaJct (
    ProdBase_Marca_jct_ID INTEGER AUTOINCREMENT,
    ProdutoBaseIDfk INTEGER,
    MarcaIDfk INTEGER
);

ALTER TABLE tbl_10(02)bProdBase_MarcaJct ADD PRIMARY KEY (ProdBase_Marca_jct_ID);

