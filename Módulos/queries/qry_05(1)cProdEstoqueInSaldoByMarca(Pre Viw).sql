-- Consulta: qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)
-- Tipo: SELECT
SELECT [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarca, Sum([QtdEmbsIn]*[Prods_Emb]) AS SaldoPorMarca
FROM [qry_02(07)aProdMarca] RIGHT JOIN ([qry_02(03)aProdutosBase] RIGHT JOIN ([qry_04(01)aProdutosFull(Edt)] RIGHT JOIN [qry_05(1)aProdEstoqueIn(Viw)] ON [qry_04(01)aProdutosFull(Edt)].ProdutoFullID = [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk) ON [qry_02(03)aProdutosBase].ProdutoBaseID = [qry_04(01)aProdutosFull(Edt)].ProdutoBaseIDfk) ON [qry_02(07)aProdMarca].ProdMarcaID = [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk
GROUP BY [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarca;

