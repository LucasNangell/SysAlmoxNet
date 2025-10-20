-- Tabela: tbl_04(01)aProdutosFull enxuto
-- Registros: 12

CREATE TABLE tbl_04(01)aProdutosFull enxuto (
    ProdutoFullID INTEGER AUTOINCREMENT,
    ProdutoFull VARCHAR(255),
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

ALTER TABLE tbl_04(01)aProdutosFull enxuto ADD PRIMARY KEY (ProdutoFullID);

