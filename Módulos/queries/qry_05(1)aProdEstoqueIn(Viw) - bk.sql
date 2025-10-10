-- Consulta: qry_05(1)aProdEstoqueIn(Viw) - bk
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12h (finalizando form Estoque).accdb

SELECT DISTINCT [tbl_05(1)aProdEstoqueIn].EstoqueInID, [tbl_05(1)aProdEstoqueIn].DataEntrada, "" AS sep, [tbl_05(1)aProdEstoqueIn].ProdutoFullIDfk, [tbl_05(1)aProdEstoqueIn].Lote, [tbl_05(1)aProdEstoqueIn].Validade, [tbl_05(1)aProdEstoqueIn].ProdMarcaIDfk, [tbl_05(1)aProdEstoqueIn].UnDaEmb_UnMedIDfk, " " & [ProdUnMedidaDescriç] AS PrdUnMdRght, [tbl_05(1)aProdEstoqueIn].Prods_Emb, [tbl_05(1)aProdEstoqueIn].QtdCons_Prod, [tbl_05(1)aProdEstoqueIn].Multiplo, [tbl_05(1)aProdEstoqueIn].QtdEmbsIn, [tbl_05(1)aProdEstoqueIn].Preço_Embalagem
FROM ([qry_02(07)aProdMarca] RIGHT JOIN ([qry_02(10)aProdUnMedida] AS [qry_02(10)aProdUnMedida_1] RIGHT JOIN [tbl_05(1)aProdEstoqueIn] ON [qry_02(10)aProdUnMedida_1].ProdUnMedidaID = [tbl_05(1)aProdEstoqueIn].UnDaEmb_UnMedIDfk) ON [qry_02(07)aProdMarca].ProdMarcaID = [tbl_05(1)aProdEstoqueIn].ProdMarcaIDfk) LEFT JOIN [tbl_04(01)aProdutosFull] ON [tbl_05(1)aProdEstoqueIn].ProdutoFullIDfk = [tbl_04(01)aProdutosFull].ProdutoFullID;

