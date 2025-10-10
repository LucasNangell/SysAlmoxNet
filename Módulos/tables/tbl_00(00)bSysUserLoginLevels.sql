-- Tabela: tbl_00(00)bSysUserLoginLevels
-- Registros: 8

CREATE TABLE tbl_00(00)bSysUserLoginLevels (
    UserLoginLevlsID INTEGER AUTOINCREMENT,
    UserLoginLevelsDscrptID VARCHAR(255) NOT NULL,
    UserLoginLevelDescriç TEXT,
    Inativo BIT DEFAULT 0
);

ALTER TABLE tbl_00(00)bSysUserLoginLevels ADD PRIMARY KEY (UserLoginLevlsID);
CREATE UNIQUE INDEX UserPermissionCoD ON tbl_00(00)bSysUserLoginLevels (UserLoginLevelsDscrptID);
CREATE INDEX UserPermissionID ON tbl_00(00)bSysUserLoginLevels (UserLoginLevlsID);

