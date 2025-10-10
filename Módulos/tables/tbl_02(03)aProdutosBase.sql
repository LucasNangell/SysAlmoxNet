-- Tabela: tbl_02(03)aProdutosBase
-- Registros: 5

CREATE TABLE tbl_02(03)aProdutosBase (
    ProdutoBaseID INTEGER AUTOINCREMENT,
    ProdutoBaseDescriçao VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(03)aProdutosBase ADD PRIMARY KEY (ProdutoBaseID);

