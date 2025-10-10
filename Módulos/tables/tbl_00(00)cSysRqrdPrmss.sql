-- Tabela: tbl_00(00)cSysRqrdPrmss
-- Registros: 2

CREATE TABLE tbl_00(00)cSysRqrdPrmss (
    CtrlPermissID INTEGER AUTOINCREMENT,
    Form VARCHAR(255),
    Control VARCHAR(255),
    UserLoginLevlsIDfk INTEGER
);

ALTER TABLE tbl_00(00)cSysRqrdPrmss ADD PRIMARY KEY (CtrlPermissID);
CREATE INDEX UserLoginLevlsID ON tbl_00(00)cSysRqrdPrmss (UserLoginLevlsIDfk);

