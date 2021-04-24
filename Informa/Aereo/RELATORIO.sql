



/* SOMA DAS TAXAS */
select 		aliquota, 
		SUM(FRETENACIONAL) FRETENACIONAL, 
		SUM(FRETEREGIONAL) FRETEREGIONAL, 
		SUM(TXORIGEM+TXDESTINO+TXREDESP+TXAGENTE+TXTRANSP+TXOUTROS1+TXOUTROS2),
		SUM(ADVALOREM) ADVALOREM,
		SUM(FRETETOTAL) FRETETOTAL
from 		tb_airawb 
WHERE 		(CANCELADO IS NULL OR CANCELADO <> 'X') AND
		DATA BETWEEN '2003-12-01' AND '2003-12-10'
GROUP BY 	ALIQUOTA 
ORDER BY 	ALIQUOTA DESC








