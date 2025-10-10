-- Tabela: tbl_02(12)aSetores
-- Registros: 5

CREATE TABLE tbl_02(12)aSetores (
    SetorID INTEGER AUTOINCREMENT,
    SetorDescriçao VARCHAR(255),
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_02(12)aSetores ADD PRIMARY KEY (SetorID);

