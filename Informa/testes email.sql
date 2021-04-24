update 	tb_codocorr 
set 	env_email = 'S',
	email_cliente = 'S',
	email1 = 'ricardo@intec.com.br',
	email2 = 'cleber@intec.com.br',
	email3 = 'samira@intec.com.br',
	email4 = 'hilton@intec.com.br'
where 	cod_ocorr = '04'

select * from tb_codocorr 

select * from tb_ctc_esp where

delete from tb_ctc_esp

delete from tb_ctc_esp where tem_ocorr <> 'N' and tem_ocorr <> '1'


select * from tb_ctc_esp where remet_cgc = '54516661002490' and tem_ocorr = 'N'

select * from tb_ctc_esp where remet_cgc = '61150819002335' and tem_ocorr = 'N'

select * from tb_ctc_esp where remet_cgc = '53162095001854' and tem_ocorr = 'N'

select * from tb_codocorr where cod_ocorr = '78'
update tb_ocorr set email_enviadoSAC = 'N' where email_enviadoSAC = 'S'
select top 100 * from tb_ocorr








select 	a.codigo,
                a.filialctc,
	a.cod_ocorr,
	a.descr_ocorr,
	a.data,
	b.cgc,
	b.nome,
                c.nfs,
               d.email1,
               d.email2,
               d.email3,
               d.email4
from 	tb_ocorr a,
	tb_cadcli b,
                tb_ctc_esp c,
                tb_codocorr d
where 	a.remet_cgc = b.cgc and 
                a.cod_ocorr = d.cod_ocorr and
                a.filialctc = c.filialctc and
                d.env_email = 'S' and
	a.email_enviadoINT = 'S'
order by   a.remet_cgc,a.filialctc,a.data