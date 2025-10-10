-- Consulta: qry_00(00)aSysUsers
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12h (finalizando form Estoque).accdb

SELECT [tbl_00(00)aSysUsers].UserID, [tbl_00(00)aSysUsers].UserLoginStR, [tbl_00(00)aSysUsers].UserName, [tbl_00(00)aSysUsers].SetorIDfk, [tbl_02(12)aSetores].SetorDescriçao, [tbl_00(00)aSysUsers].Inativo
FROM [tbl_02(12)aSetores] RIGHT JOIN [tbl_00(00)aSysUsers] ON [tbl_02(12)aSetores].SetorID = [tbl_00(00)aSysUsers].SetorIDfk;

