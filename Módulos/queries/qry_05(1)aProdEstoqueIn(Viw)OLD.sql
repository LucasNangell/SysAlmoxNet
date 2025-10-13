-- Consulta: qry_05(1)aProdEstoqueIn(Viw)OLD
-- Tipo: SELECT
SELECT DISTINCT [tbl_05(1)aProdEstoqueIn].EstoqueInID, [tbl_05(1)aProdEstoqueIn].DataEntrada, [tbl_05(1)aProdEstoqueIn].ProdutoIDfk, [tbl_05(1)aProdEstoqueIn].Lote, [tbl_05(1)aProdEstoqueIn].Validade, [tbl_05(1)aProdEstoqueIn].ProdMarcaIDfk, [tbl_02(2)dProdMarca].ProdMarca, [tbl_05(1)aProdEstoqueIn].UnDaEmb_UnMedIDfk, [tbl_02(3)cProdUnMedida].ProdUnidadeMedida, " " & [ProdUnidadeMedida] AS PrdUnMdRght, [tbl_05(1)aProdEstoqueIn].Prods_Emb, [tbl_05(1)aProdEstoqueIn].QtdCons_Prod, [tbl_05(1)aProdEstoqueIn].Multiplo, [tbl_05(1)aProdEstoqueIn].QtdEmbsIn, [tbl_05(1)aProdEstoqueIn].Preço_Embalagem
FROM [tbl_02(3)cProdUnMedida] INNER JOIN ([tbl_02(2)dProdMarca] INNER JOIN [tbl_05(1)aProdEstoqueIn] ON [tbl_02(2)dProdMarca].ProdMarcaID = [tbl_05(1)aProdEstoqueIn].ProdMarcaIDfk) ON [tbl_02(3)cProdUnMedida].ProdUnMedida = [tbl_05(1)aProdEstoqueIn].UnDaEmb_UnMedIDfk;

