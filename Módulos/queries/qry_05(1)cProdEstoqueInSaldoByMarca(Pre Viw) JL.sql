-- Consulta: qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw) JL
-- Tipo: SELECT
SELECT [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_10(02)bProdBase_MarcaJct].MarcaIDfk, [qry_02(07)aProdMarca].ProdMarca, Sum([QtdEmbsIn]*[Prods_Emb]) AS SaldoPorMarca, Format(Sum([QtdEmbsIn]*[Prods_Emb]),"#,###") AS SaldoPorMarcaStr
FROM [qry_02(07)aProdMarca] RIGHT JOIN ([qry_10(02)bProdBase_MarcaJct] RIGHT JOIN ([qry_02(03)aProdutosBase] RIGHT JOIN ([qry_04(01)aProdutosFull(Edt)] RIGHT JOIN [qry_05(1)aProdEstoqueIn(Viw)] ON [qry_04(01)aProdutosFull(Edt)].ProdutoFullID = [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk) ON [qry_02(03)aProdutosBase].ProdutoBaseID = [qry_04(01)aProdutosFull(Edt)].ProdutoBaseIDfk) ON [qry_10(02)bProdBase_MarcaJct].ProdutoBaseIDfk = [qry_02(03)aProdutosBase].ProdutoBaseID) ON [qry_02(07)aProdMarca].ProdMarcaID = [qry_10(02)bProdBase_MarcaJct].MarcaIDfk
GROUP BY [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_10(02)bProdBase_MarcaJct].MarcaIDfk, [qry_02(07)aProdMarca].ProdMarca;

