select 		b.remet_nome, a.numnf, b.dest_nome, count(*) tot
from 		tb_nf_esp a, tb_ctc_esp b 
where 		a.filialctc = b.filialctc and
		b.data between '2002/12/01' and '2002/12/19'
group by 	b.remet_nome, a.numnf, b.dest_nome

