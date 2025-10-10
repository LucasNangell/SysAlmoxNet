-- Consulta: qry_02(10)aProdUnMedida
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [tbl_02(10)aProdUnMedida].ProdUnMedidaID, [tbl_02(10)aProdUnMedida].ProdUnMedidaDescriç, [tbl_02(10)aProdUnMedida].Inativo
FROM [tbl_02(10)aProdUnMedida]
ORDER BY [tbl_02(10)aProdUnMedida].ProdUnMedidaDescriç;

