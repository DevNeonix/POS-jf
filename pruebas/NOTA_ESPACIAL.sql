USE VENTAS
GO

ALTER TABLE MOVIMDET
ADD TEXTO VARCHAR(200) NULL
GO

ALTER VIEW [dbo].[VIEW_VENTAS_ARTICULO]
AS

SELECT     TOP (100) PERCENT b1.CODART, LTRIM(RTRIM(dbo.View_ARTICULOS_TIENDA.DESCRI))  COLLATE Modern_Spanish_CI_AI + ' ' + LTRIM(RTRIM(ISNULL(TEXTO,''))) AS DESCRI, b1.PRECIO, b1.DESCUENTO AS DCT, b1.IGV, b1.PORDES, 
                      b1.PRECIO + b1.DESCUENTO AS PVP, b1.SALE, b1.ENTRA, b1.VALE, a1.TIPMOV, a1.CODDOC AS TIPDOC, a1.SERIE + '-' + a1.NUMDOC AS NUMDOC, a1.CLIENTE, 
                      cl1.NOMBRE, a1.FECDOC AS FECHA, a1.OPERACION, a1.DOCORI, a1.SERORI, a1.NUMORI, CASE WHEN (b1.PRECIO + b1.DESCUENTO - b1.IGV) > 0 AND 
                      (b1.SALE + b1.ENTRA + b1.VALE) > 0 THEN (b1.PRECIO + b1.DESCUENTO - b1.IGV) / (b1.SALE + b1.ENTRA + b1.VALE) ELSE 0 END AS LISTA1, b1.TIENDA, b1.ITEM, 
                      dbo.VIEW_GRUPOS.GRUPO, b1.texto
FROM         dbo.movimdet AS b1 INNER JOIN
                      dbo.movimcab AS a1 ON a1.OPERACION = b1.OPERACION AND b1.TIENDA = a1.TIENDA INNER JOIN
                      dbo.View_ARTICULOS_TIENDA ON b1.CODART = dbo.View_ARTICULOS_TIENDA.CODIGO AND b1.TIENDA = dbo.View_ARTICULOS_TIENDA.TIENDA INNER JOIN
                      dbo.VIEW_GRUPOS ON b1.TIENDA = dbo.VIEW_GRUPOS.TIENDA AND b1.CODART = dbo.VIEW_GRUPOS.CODIGO LEFT OUTER JOIN
                      dbo.CLIENTES AS cl1 ON a1.CLIENTE = cl1.CLIENTE
WHERE     (a1.ESTADO = 'A')
ORDER BY b1.CODART
GO