-- Consulta: qry_02(12)aSetores
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [tbl_02(12)aSetores].SetorID, [tbl_02(12)aSetores].SetorDescriçao, [tbl_02(12)aSetores].Inativo
FROM [tbl_02(12)aSetores]
ORDER BY [tbl_02(12)aSetores].SetorDescriçao;

