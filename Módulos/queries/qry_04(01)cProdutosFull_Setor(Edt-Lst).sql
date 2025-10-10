-- Consulta: qry_04(01)cProdutosFull_Setor(Edt-Lst)
-- Tipo: SELECT
-- Banco: ControleEstoque 64bits v12g (passando Recd pro form ao abrir JL).accdb

SELECT [qry_04(01)aProdutosFull(Edt)].ProdutoFullID, [qry_04(01)aProdutosFull(Edt)].ProdutoFull, [qry_10(04)aProdutoFull_SetorJct].SetorIDfk, [qry_04(01)aProdutosFull(Edt)].ProdutoBaseIDfk, [qry_04(01)aProdutosFull(Edt)].Variaçao, [qry_04(01)aProdutosFull(Edt)].ProdCorIDfk, [qry_04(01)aProdutosFull(Edt)].ProdMaterialIDfk, [qry_04(01)aProdutosFull(Edt)].ProdMedidaIDfk, [qry_04(01)aProdutosFull(Edt)].Complemento, [qry_04(01)aProdutosFull(Edt)].UnProdutoIDfk, [qry_04(01)aProdutosFull(Edt)].UnMedConsumoIDfk, [qry_04(01)aProdutosFull(Edt)].UnPedidoIDfk, [qry_04(01)aProdutosFull(Edt)].QtdMinEmEstoque, [qry_04(01)aProdutosFull(Edt)].ProdAplicaçaoIDfk, [qry_04(01)aProdutosFull(Edt)].Inativo
FROM [qry_02(03)aProdutosBase] INNER JOIN ([qry_10(04)aProdutoFull_SetorJct] INNER JOIN [qry_04(01)aProdutosFull(Edt)] ON [qry_10(04)aProdutoFull_SetorJct].ProdutoFullIDfk = [qry_04(01)aProdutosFull(Edt)].ProdutoFullID) ON [qry_02(03)aProdutosBase].ProdutoBaseID = [qry_04(01)aProdutosFull(Edt)].ProdutoBaseIDfk
ORDER BY [qry_04(01)aProdutosFull(Edt)].ProdutoFull;

