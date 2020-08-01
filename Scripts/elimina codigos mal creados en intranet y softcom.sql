USE MODELADOR
GO
-- 98 DE LA INTRANET
SELECT       CODIGO INTO #TEMPO
FROM            AVIOS INNER JOIN
RSFACCAR..AL0001ARTI ON LTRIM(RTRIM(CODIGO)) COLLATE Modern_Spanish_CI_AS = LTRIM(RTRIM(AR_CCODIGO))
WHERE        (USUARIO = 'RBENAVENTE') AND (YEAR(FECHA) = 2017) AND (DESCRIPCION LIKE '%JACINTA%')

SELECT * FROM #TEMPO

Select
   codigo
    , Count(*) As [Total]
From
    avios
Group By
codigo
Having
    Count(*) > 1



SELECT * FROM 
--delete
avios
WHERE ltrim(rtrim(AVIOS.CODIGO)) = (select ltrim(rtrim(codigo)) from #tempo where ltrim(rtrim(avios.codigo))= ltrim(rtrim(#tempo.codigo)))

select * from  
--delete
AVIOSDETA 
WHERE ltrim(rtrim(AVIOSdeta.avio)) = (select ltrim(rtrim(codigo)) from #tempo where ltrim(rtrim(aviosdeta.avio))= ltrim(rtrim(#tempo.codigo)))


SELECT * FROM 
--delete
RSFACCAR..AL0001ARTI 
WHERE ltrim(rtrim(AR_CCODIGO)) collate SQL_Latin1_General_CP1_CI_AI =
 (select ltrim(rtrim(codigo)) from #tempo where ltrim(rtrim(ar_ccodigo)) collate SQL_Latin1_General_CP1_CI_AI
 = ltrim(rtrim(#tempo.codigo)))

drop table #TEMPO
