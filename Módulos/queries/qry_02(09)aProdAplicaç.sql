-- Consulta: qry_02(09)aProdAplicaç
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [tbl_02(09)aProdAplicaç].ProdAplicaçaoID, [tbl_02(09)aProdAplicaç].ProdAplicaçaoDescriç, [tbl_02(09)aProdAplicaç].Inativo
FROM [tbl_02(09)aProdAplicaç]
ORDER BY [tbl_02(09)aProdAplicaç].ProdAplicaçaoDescriç;

