-- Tabela: tbl_10(00)cSysUsersPrmissJct
-- Registros: 6

CREATE TABLE tbl_10(00)cSysUsersPrmissJct (
    UserPermissJctID INTEGER AUTOINCREMENT,
    UserIDfk INTEGER,
    UserLoginLevlsIDfk INTEGER
);

ALTER TABLE tbl_10(00)cSysUsersPrmissJct ADD PRIMARY KEY (UserPermissJctID);
CREATE INDEX UserSetupsID ON tbl_10(00)cSysUsersPrmissJct (UserPermissJctID);

