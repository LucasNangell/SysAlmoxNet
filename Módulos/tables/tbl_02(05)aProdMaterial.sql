-- Tabela: tbl_02(05)aProdMaterial
-- Registros: 3

CREATE TABLE tbl_02(05)aProdMaterial (
    ProdMaterialID INTEGER AUTOINCREMENT,
    ProdMaterial VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(05)aProdMaterial ADD PRIMARY KEY (ProdMaterialID);

