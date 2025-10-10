-- Consulta: qry_02(06)aProdMedidas
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [tbl_02(06)aProdMedidas].ProdMedidaID, [tbl_02(06)aProdMedidas].ProdMedida, [tbl_02(06)aProdMedidas].Inativo
FROM [tbl_02(06)aProdMedidas]
ORDER BY [tbl_02(06)aProdMedidas].ProdMedida;

