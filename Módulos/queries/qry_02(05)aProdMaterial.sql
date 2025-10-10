-- Consulta: qry_02(05)aProdMaterial
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [tbl_02(05)aProdMaterial].ProdMaterialID, [tbl_02(05)aProdMaterial].ProdMaterial, [tbl_02(05)aProdMaterial].Inativo
FROM [tbl_02(05)aProdMaterial]
ORDER BY [tbl_02(05)aProdMaterial].ProdMaterial;

