-- Consulta: qry_05(1)dProdEstoqueInSaldoByMarca(Viw)
-- Tipo: SELECT
SELECT [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].ProdutoFullIDfk, [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].ProdMarcaIDfk, [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].ProdMarca, [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].SaldoPorMarca, Format([SaldoPorMarca],"#,###") AS C1, String(10-Len([C1])," ") AS C2, [C2] & [C1] AS SaldoPorMarcaStr, [tbl_04(01)aProdutosFull].UnPedidoIDfk, [tbl_02(10)aProdUnMedida].ProdUnMedidaDescriç
FROM [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)] LEFT JOIN ([tbl_04(01)aProdutosFull] LEFT JOIN [tbl_02(10)aProdUnMedida] ON [tbl_04(01)aProdutosFull].UnPedidoIDfk = [tbl_02(10)aProdUnMedida].ProdUnMedidaID) ON [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].ProdutoFullIDfk = [tbl_04(01)aProdutosFull].ProdutoFullID;

