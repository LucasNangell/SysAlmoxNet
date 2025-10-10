-- Consulta: qry_02(04)aProdCor
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [tbl_02(04)aProdCor].ProdCorID, [tbl_02(04)aProdCor].ProdCor, [tbl_02(04)aProdCor].Inativo
FROM [tbl_02(04)aProdCor]
ORDER BY [tbl_02(04)aProdCor].ProdCor;

