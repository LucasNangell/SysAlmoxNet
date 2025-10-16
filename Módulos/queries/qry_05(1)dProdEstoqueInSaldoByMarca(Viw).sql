-- Consulta: qry_05(1)dProdEstoqueInSaldoByMarca(Viw)
-- Tipo: SELECT
SELECT [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].ProdutoFullIDfk, [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].ProdMarcaIDfk, [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].ProdMarca, [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].SaldoPorMarca, [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].SaldoPorMarcaStr, [tbl_04(01)aProdutosFull].UnPedidoIDfk, [tbl_02(10)aProdUnMedida].ProdUnMedidaDescriç
FROM ([tbl_04(01)aProdutosFull] RIGHT JOIN [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)] ON [tbl_04(01)aProdutosFull].ProdutoFullID = [qry_05(1)cProdEstoqueInSaldoByMarca(Pre Viw)].ProdutoFullIDfk) LEFT JOIN [tbl_02(10)aProdUnMedida] ON [tbl_04(01)aProdutosFull].UnPedidoIDfk = [tbl_02(10)aProdUnMedida].ProdUnMedidaID;

