-- Consulta: qry_05(1)cProdEstoqueSaldoMarcaGrpd(Viw)
-- Tipo: SELECT
SELECT [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_04(01)aProdutosFull(Edt)].ProdutoFull, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk, Sum([QtdEmbsIn]*[Prods_Emb]) AS SaldoPorMarca
FROM ([qry_04(01)aProdutosFull(Edt)] RIGHT JOIN [qry_05(1)aProdEstoqueIn(Viw)] ON [qry_04(01)aProdutosFull(Edt)].ProdutoFullID = [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk) LEFT JOIN [qry_02(07)aProdMarca] ON [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk = [qry_02(07)aProdMarca].ProdMarcaID
GROUP BY [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_04(01)aProdutosFull(Edt)].ProdutoFull, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk;

