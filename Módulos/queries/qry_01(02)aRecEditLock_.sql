-- Consulta: qry_01(02)aRecEditLock_
-- Tipo: SELECT
SELECT [tbl_01(02)aRecEditLock].RecLockID, [tbl_01(02)aRecEditLock].LockedRecID, [tbl_01(02)aRecEditLock].RecSource, [tbl_01(02)aRecEditLock].UserLoginID, [tbl_00(00)aSysUsers].UserLoginStR, [tbl_00(00)aSysUsers].UserName, [tbl_01(02)aRecEditLock].EditStartTime
FROM [tbl_00(00)aSysUsers] RIGHT JOIN [tbl_01(02)aRecEditLock] ON [tbl_00(00)aSysUsers].UserID = [tbl_01(02)aRecEditLock].UserLoginID;

