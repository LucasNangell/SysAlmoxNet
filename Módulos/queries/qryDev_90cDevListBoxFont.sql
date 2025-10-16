-- Consulta: qryDev_90cDevListBoxFont
-- Tipo: SELECT
SELECT tblDev_90cDevListBoxFont.Código, tblDev_90cDevListBoxFont.Produto, tblDev_90cDevListBoxFont.Saldo, CStr([Saldo]) AS C1, Format([C1],"#,###") AS C2, String(11-Len([C2])," ") AS C3, [C3] & [C2] AS C4
FROM tblDev_90cDevListBoxFont;

