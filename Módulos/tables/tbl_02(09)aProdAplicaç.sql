-- Tabela: tbl_02(09)aProdAplica�
-- Registros: 3

CREATE TABLE tbl_02(09)aProdAplica� (
    ProdAplica�aoID INTEGER AUTOINCREMENT,
    ProdAplica�aoDescri� VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(09)aProdAplica� ADD PRIMARY KEY (ProdAplica�aoID);

