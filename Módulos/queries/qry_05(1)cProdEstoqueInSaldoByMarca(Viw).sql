-- Consulta: qry_05(1)cProdEstoqueInSaldoByMarca(Viw)
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12h (finalizando form Estoque).accdb

SELECT [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_04(01)aProdutosFull(Edt)].ProdutoFull, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarca AS ProdMarcaStr, Sum([QtdEmbsIn]*[Prods_Emb]) AS SaldoPorMarca
FROM ([qry_04(01)aProdutosFull(Edt)] RIGHT JOIN [qry_05(1)aProdEstoqueIn(Viw)] ON [qry_04(01)aProdutosFull(Edt)].ProdutoFullID = [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk) LEFT JOIN [qry_02(07)aProdMarca] ON [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk = [qry_02(07)aProdMarca].ProdMarcaID
GROUP BY [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_04(01)aProdutosFull(Edt)].ProdutoFull, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarca;

