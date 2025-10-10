-- Tabela: tbl_00(00)aSysUsers
-- Registros: 5

CREATE TABLE tbl_00(00)aSysUsers (
    UserID INTEGER AUTOINCREMENT,
    UserLoginStR VARCHAR(255),
    UserName VARCHAR(255),
    SetorIDfk INTEGER,
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_00(00)aSysUsers ADD PRIMARY KEY (UserID);
CREATE INDEX UserID ON tbl_00(00)aSysUsers (UserID);
CREATE UNIQUE INDEX UserLogin ON tbl_00(00)aSysUsers (UserLoginStR);

