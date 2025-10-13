-- Consulta: qry_00(00)cSysRqrdPrmss
-- Tipo: SELECT
SELECT [tbl_00(00)cSysRqrdPrmss].CtrlPermissID, [tbl_00(00)cSysRqrdPrmss].Form, [tbl_00(00)cSysRqrdPrmss].Control, [tbl_00(00)cSysRqrdPrmss].UserLoginLevlsIDfk, [tbl_00(00)bSysUserLoginLevels].UserLoginLevelDescriç
FROM [tbl_00(00)bSysUserLoginLevels] RIGHT JOIN [tbl_00(00)cSysRqrdPrmss] ON [tbl_00(00)bSysUserLoginLevels].UserLoginLevlsID = [tbl_00(00)cSysRqrdPrmss].UserLoginLevlsIDfk;

