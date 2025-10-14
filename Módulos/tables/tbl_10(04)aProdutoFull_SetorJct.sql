-- Tabela: tbl_10(04)aProdutoFull_SetorJct
-- Registros: 19

CREATE TABLE tbl_10(04)aProdutoFull_SetorJct (
    ProdJctID INTEGER AUTOINCREMENT,
    ProdutoFullIDfk INTEGER,
    SetorIDfk INTEGER
);

ALTER TABLE tbl_10(04)aProdutoFull_SetorJct ADD PRIMARY KEY (ProdJctID);

