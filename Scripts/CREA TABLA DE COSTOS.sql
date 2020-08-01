CREATE TABLE [dbo].[COSTOS](
	[CODIGO] [char](5)  NOT NULL PRIMARY KEY ,
	[COSTO]  [NUMERIC] (7,2),
	[USUARIO] [char](10)  NOT NULL DEFAULT 'SISTEMAS',
	[FECHA] [smalldatetime] NOT NULL DEFAULT (getdate()),
	[ESTADO] [char](1)  NOT NULL  DEFAULT ('A'),
) ON [PRIMARY]
-----------------------------------------------------------------------------------

INSERT INTO COSTOS 
SELECT DISTINCT LEFT(CODIGO,5), 0, 'AUTOM', GETDATE(), 'A' FROM ARTICULOS
-----------------------------------------------------------------------------------

USE [VENTAS]
GO
/****** Object:  Trigger [dbo].[trig_INSERT_COSTO]    Script Date: 03/13/2015 09:45:11 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE TRIGGER [dbo].[trig_INSERT_COSTO]
ON [dbo].[ARTICULOS]
FOR INSERT
AS
Begin
  SET NOCOUNT ON
	IF EXISTS (SELECT I.codigo FROM inserted AS i left JOIN costos a ON a.codigo = left(ltrim(rtrim(i.codigo)),5) WHERE a.codigo is null )
		BEGIN
			INSERT INTO COSTOS SELECT LEFT(I.CODIGO,5), 0, 'AUTOM', GETDATE(), 'A' FROM INSERTED I
			left JOIN costos a  ON a.codigo = left(ltrim(rtrim(i.codigo)),5)
            WHERE a.codigo is null  
		END	
				
End
-----------------------------------------------------------------------------------
USE VENTAS
GO

CREATE VIEW VIEW_COSTOS 
AS
SELECT	DISTINCT top 100 percent   C.CODIGO, C.COSTO, MIN(REPLACE ( A.AR_CDESCRI , '"' , 'PLG')) AS DESCRI, str(MAX(AR_NPRECI6),12,2) AS PVP
FROM	RSFACCAR.dbo.AL0012ARTI  as A,dbo.COSTOS AS C 
WHERE	LEFT(LTRIM(RTRIM(A.AR_CCODIGO)),5)+'00000' COLLATE Modern_Spanish_CI_AI = LTRIM(RTRIM(C.CODIGO)) +'00000' and c.estado='a'
GROUP BY C.CODIGO, C.COSTO
ORDER BY C.CODIGO

go


