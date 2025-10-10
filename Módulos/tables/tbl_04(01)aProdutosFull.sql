-- Tabela: tbl_04(01)aProdutosFull
-- Registros: 11

CREATE TABLE tbl_04(01)aProdutosFull (
    ProdutoFullID INTEGER AUTOINCREMENT,
    ProdutoBaseIDfk INTEGER,
    Variaçao VARCHAR(255),
    ProdCorIDfk INTEGER,
    ProdMaterialIDfk INTEGER,
    ProdMedidaIDfk INTEGER,
    Complemento VARCHAR(255),
    UnProdutoIDfk INTEGER,
    UnMedConsumoIDfk INTEGER,
    UnPedidoIDfk INTEGER,
    QtdMinEmEstoque INTEGER,
    ProdAplicaçaoIDfk INTEGER,
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_04(01)aProdutosFull ADD PRIMARY KEY (ProdutoFullID);
CREATE INDEX tbl_02(2)aProdCortbl_02(1)aProdutoBase1 ON tbl_04(01)aProdutosFull (ProdCorIDfk);
CREATE INDEX tbl_02(2)bProdMaterialtbl_02(1)aProdutoBase1 ON tbl_04(01)aProdutosFull (ProdMaterialIDfk);
CREATE INDEX tbl_02(2)cProdDimenstbl_02(1)aProdutoBase2 ON tbl_04(01)aProdutosFull (ProdMedidaIDfk);
CREATE INDEX tbl_02(3)bProdAplicaçtbl_02(1)aProdutoBase1 ON tbl_04(01)aProdutosFull (ProdAplicaçaoIDfk);
CREATE INDEX tbl_02(3)cProdUnMedidatbl_02(1)aProdutoBase1 ON tbl_04(01)aProdutosFull (UnProdutoIDfk);
CREATE INDEX tbl_02(3)cProdUnMedidatbl_02(1)aProdutoBase2 ON tbl_04(01)aProdutosFull (UnMedConsumoIDfk);
CREATE INDEX tbl_02(3)cProdUnMedidatbl_02(1)aProdutoBase3 ON tbl_04(01)aProdutosFull (UnPedidoIDfk);

