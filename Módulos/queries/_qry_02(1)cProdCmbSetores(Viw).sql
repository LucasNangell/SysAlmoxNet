-- Consulta: _qry_02(1)cProdCmbSetores(Viw)
-- Tipo: SELECT
SELECT DISTINCT [qry_02(1)aProdutoBase(Edt)].ProdutoID AS ProdutoID, [qry_02(09)aSetores(Edt)].SetorID AS SetorID, [qry_02(09)aSetores(Edt)].SetorDescri�ao
FROM [qry_02(1)aProdutoBase(Edt)] INNER JOIN [qry_02(09)aSetores(Edt)] ON [qry_02(1)aProdutoBase(Edt)].SetorIDfk=[qry_02(09)aSetores(Edt)].SetorID;

