-- Consulta: qry_05(1)eProdEstoqueSaldo(Viw)
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [qry_05(1)bProdEstoqueOut(Viw)].ProdutoFullIDfk, Sum([qry_05(1)dProdEstoqueSaldoProdGrpd(Viw)].Saldo) AS TotalIn, Sum([qry_05(1)bProdEstoqueOut(Viw)].TotalProdsOut) AS TotalOut, Sum([Saldo]-[TotalProdsOut]) AS SaldoTotal, Format(Sum([Saldo]-[TotalProdsOut]),"#,###") AS SaldoStr
FROM [qry_05(1)dProdEstoqueSaldoProdGrpd(Viw)] INNER JOIN [qry_05(1)bProdEstoqueOut(Viw)] ON [qry_05(1)dProdEstoqueSaldoProdGrpd(Viw)].ProdutoFullIDfk = [qry_05(1)bProdEstoqueOut(Viw)].ProdutoFullIDfk
GROUP BY [qry_05(1)bProdEstoqueOut(Viw)].ProdutoFullIDfk;

