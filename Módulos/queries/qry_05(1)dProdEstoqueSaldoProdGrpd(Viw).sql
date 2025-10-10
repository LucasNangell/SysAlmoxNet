-- Consulta: qry_05(1)dProdEstoqueSaldoProdGrpd(Viw)
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [tbl_05(1)aProdEstoqueIn].ProdutoFullIDfk, Sum([tbl_05(1)aProdEstoqueIn].Prods_Emb) AS SomaDeProds_Emb, Sum([tbl_05(1)aProdEstoqueIn].QtdEmbsIn) AS SomaDeQtdEmbs, [tbl_05(1)aProdEstoqueIn].Saldo
FROM [tbl_05(1)aProdEstoqueIn]
GROUP BY [tbl_05(1)aProdEstoqueIn].ProdutoFullIDfk, [tbl_05(1)aProdEstoqueIn].Saldo;

