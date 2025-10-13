-- Consulta: qry_10(00)cSysUsersPrmissSetorJct(Edt)
-- Tipo: SELECT
SELECT [tbl_10(00)cSysUsersPrmissJct].UserPermissJctID, [tbl_10(00)cSysUsersPrmissJct].UserIDfk, [qry_00(00)aSysUsers].UserLoginStR, [qry_00(00)aSysUsers].SetorIDfk, [tbl_10(00)cSysUsersPrmissJct].UserLoginLevlsIDfk, [qry_00(00)aSysUsers].UserName, [tbl_00(00)bSysUserLoginLevels].UserLoginLevelDescriç
FROM ([tbl_10(00)cSysUsersPrmissJct] LEFT JOIN [qry_00(00)aSysUsers] ON [tbl_10(00)cSysUsersPrmissJct].UserIDfk = [qry_00(00)aSysUsers].UserID) LEFT JOIN [tbl_00(00)bSysUserLoginLevels] ON [tbl_10(00)cSysUsersPrmissJct].UserLoginLevlsIDfk = [tbl_00(00)bSysUserLoginLevels].UserLoginLevlsID
ORDER BY [qry_00(00)aSysUsers].UserName;

