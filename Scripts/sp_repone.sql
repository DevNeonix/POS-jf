set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go



-- ===================================================
-- Author:		Mabel Molina
-- Create date: 13-02-2012
-- Description:	Envia la lista de articulos a reponer
--				de la tienda
-- ===================================================
ALTER PROCEDURE [dbo].[sp_reponer]
	-- Add the parameters for the stored procedure here
	@tda	varchar(2),
	@descri varchar(20),
	@to     varchar(500)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    DECLARE @tableHTML  NVARCHAR(MAX);

SET @tableHTML =
N'<table border="1" cellpadding="5" cellspacing="0">'+
N'<thead>'+
N'<tr style=''background: gainsboro;    color: black;font-weight: bold;''>'+
N'<th>Codigo</th>'+
N'<th>Descripcion</th>'+
N'<th>Cant</th>'+
N'</tr>'+
N'</thead>'+
N'<tbody>'+
CAST ( ( SELECT td = codigo, '',
                td = descri, '',
                td = minimo-stock, ''
              FROM dbo.view_articulos_tienda
              where stock <  minimo and planilla ='1'  
              and tienda = @tda
              FOR XML PATH('tr'), TYPE
    ) AS NVARCHAR(MAX) ) +
N'</tbody>'+
N'<tr style=''background: gainsboro;    color: black;font-weight: bold;''>'+
N'<td colspan =''3''>Esta solicitud proviene de la Tienda: ' + @descri + '</td>'+
N'</tr>'+
N'</table>'

DECLARE @receptores  VARCHAR(500);

set @receptores = @to
EXEC msdb.dbo.sp_send_dbmail
    @profile_name = 'dbMailProfile',--Perfil de correo configurado.
--  @recipients = 'earellano@elmodelador.com.pe;sistemas@elmodelador.com.pe; rmalaga@elmodelador.com.pe; csaba@elmodelador.com.pe', -- A quien se va enviar el correo.
	@recipients = @receptores,
--  @recipients = 'mmolina@elmodelador.com.pe;', -- A quien se va enviar el correo.
    @subject = 'Solicitud para reponer articulos en tienda',
    @body = @tableHTML,
    @body_format = 'HTML',
    @importance = 'High'


END



