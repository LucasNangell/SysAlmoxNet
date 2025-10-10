-- Consulta: qry_05(1)bProdEstoqueOut(Viw)
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12h (finalizando form Estoque).accdb

SELECT [qry_05(1)bProdEstoqueOut].EstoqueOutID, [qry_05(1)bProdEstoqueOut].DataSaida, [qry_05(1)bProdEstoqueOut].EstoqueInIDfk, [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_04(1)aProdutosFull(Pre Edt)].ProdutoFull, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk, [qry_05(1)bProdEstoqueOut].QtdEmbsOut, [qry_05(1)aProdEstoqueIn(Viw)].Prods_Emb, Sum(([QtdEmbsOut]*[Prods_Emb])) AS TotalProdsOut
FROM (([qry_04(1)aProdutosFull(Pre Edt)] RIGHT JOIN [qry_05(1)aProdEstoqueIn(Viw)] ON [qry_04(1)aProdutosFull(Pre Edt)].ProdutoFullID = [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk) RIGHT JOIN [qry_05(1)bProdEstoqueOut] ON [qry_05(1)aProdEstoqueIn(Viw)].EstoqueInID = [qry_05(1)bProdEstoqueOut].EstoqueInIDfk) LEFT JOIN [qry_02(07)aProdMarca] ON [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk = [qry_02(07)aProdMarca].ProdMarcaID
GROUP BY [qry_05(1)bProdEstoqueOut].EstoqueOutID, [qry_05(1)bProdEstoqueOut].DataSaida, [qry_05(1)bProdEstoqueOut].EstoqueInIDfk, [qry_05(1)aProdEstoqueIn(Viw)].ProdutoFullIDfk, [qry_04(1)aProdutosFull(Pre Edt)].ProdutoFull, [qry_05(1)aProdEstoqueIn(Viw)].ProdMarcaIDfk, [qry_05(1)bProdEstoqueOut].QtdEmbsOut, [qry_05(1)aProdEstoqueIn(Viw)].Prods_Emb;

