-- Tabela: tbl_05(1)aProdEstoqueIn
-- Registros: 4

CREATE TABLE tbl_05(1)aProdEstoqueIn (
    EstoqueInID INTEGER AUTOINCREMENT,
    DataEntrada DATETIME,
    ProdutoFullIDfk INTEGER,
    ProdMarcaIDfk INTEGER,
    UnDaEmb_UnMedIDfk INTEGER,
    QtdCons_Prod INTEGER,
    Prods_Emb INTEGER,
    Multiplo BIT DEFAULT 0,
    QtdEmbsIn INTEGER,
    Saldo INTEGER,
    Preço_Embalagem INTEGER,
    Lote VARCHAR(255),
    Validade DATETIME
);

ALTER TABLE tbl_05(1)aProdEstoqueIn ADD PRIMARY KEY (EstoqueInID);
CREATE INDEX tbl_02(1)aProdutoBasetbl_05(1)aProdEstoqueIn ON tbl_05(1)aProdEstoqueIn (ProdutoFullIDfk);

