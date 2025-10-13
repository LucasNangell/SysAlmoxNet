-- Consulta: qry_05(1)dProdEstoqueSaldoProdGrpd(Viw)
-- Tipo: SELECT
SELECT [tbl_05(1)aProdEstoqueIn].ProdutoFullIDfk, Sum([tbl_05(1)aProdEstoqueIn].Prods_Emb) AS SomaDeProds_Emb, Sum([tbl_05(1)aProdEstoqueIn].QtdEmbsIn) AS SomaDeQtdEmbs, [tbl_05(1)aProdEstoqueIn].Saldo
FROM [tbl_05(1)aProdEstoqueIn]
GROUP BY [tbl_05(1)aProdEstoqueIn].ProdutoFullIDfk, [tbl_05(1)aProdEstoqueIn].Saldo;

