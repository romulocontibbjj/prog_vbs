if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_airtabaerea_tb_airciaaerea]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_airtabpreco] DROP CONSTRAINT FK_tb_airtabaerea_tb_airciaaerea
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_airtabgeral_tb_airtabaerea]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_airtabprecogeral] DROP CONSTRAINT FK_tb_airtabgeral_tb_airtabaerea
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_airtabtetc_tb_airtabaerea]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_airtabprecotetc] DROP CONSTRAINT FK_tb_airtabtetc_tb_airtabaerea
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_ctc_esp_tb_cadcli]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_ctc_esp] DROP CONSTRAINT FK_tb_ctc_esp_tb_cadcli
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_ocorr_tb_cadusu]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_ocorr] DROP CONSTRAINT FK_tb_ocorr_tb_cadusu
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_ocorr_tb_codocorr]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_ocorr] DROP CONSTRAINT FK_tb_ocorr_tb_codocorr
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_manifesto_tb_ctc_esp]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_manifesto] DROP CONSTRAINT FK_tb_manifesto_tb_ctc_esp
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_nf_esp_tb_ctc_esp]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_nf_esp] DROP CONSTRAINT FK_tb_nf_esp_tb_ctc_esp
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_ocorr_tb_ctc_esp]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_ocorr] DROP CONSTRAINT FK_tb_ocorr_tb_ctc_esp
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_redaereoitem_tb_redaereo]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_redaereoitem] DROP CONSTRAINT FK_tb_redaereoitem_tb_redaereo
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tb_tabaereoitem_tb_tabaereo]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tb_tabaereoitem] DROP CONSTRAINT FK_tb_tabaereoitem_tb_tabaereo
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_ExportSitlaNao]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_ExportSitlaNao]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_ObsOcorr_Sitla]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_ObsOcorr_Sitla]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_cadcli]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_cadcli]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_cadcli_imp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_cadcli_imp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_cadferiado]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_cadferiado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_cadocorr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_cadocorr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_cadusu]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_cadusu]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_diasprazo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_diasprazo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_manif]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_manif]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_obsentr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_obsentr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_obsocorr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_obsocorr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_ocorr1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_ocorr1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_ocorr1ow]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_ocorr1ow]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_ocorr2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_ocorr2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_ocorrcom]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_ocorrcom]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_alt_temocorr_sn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_alt_temocorr_sn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_excl_cadocorr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_excl_cadocorr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_excl_ocorr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_excl_ocorr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_incl_cadcli]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_incl_cadcli]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_cadferiado]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_cadferiado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_cadocorr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_cadocorr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_cadprazo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_cadprazo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_cadusu]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_cadusu]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_impctc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_impctc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_impnf]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_impnf]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_logusuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_logusuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_manifesto]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_manifesto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_manifestoNull]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_manifestoNull]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_ocorr1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_ocorr1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_ocorr3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_ocorr3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_ocorr4]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_ocorr4]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_ocorr4cod00]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_ocorr4cod00]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_ins_ultconssac]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_ins_ultconssac]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_sel_ctcNFeCGC]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_sel_ctcNFeCGC]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BASE]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BASE]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[est_clientes_bomi]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[est_clientes_bomi]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[est_faixas_de_peso]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[est_faixas_de_peso]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[est_faixas_de_valor]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[est_faixas_de_valor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[prz_clientes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[prz_clientes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[prz_codoc]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[prz_codoc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[prz_cortea]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[prz_cortea]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[prz_corter]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[prz_corter]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[salete]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[salete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_CadFeriado]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_CadFeriado]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_Rel_Diarios]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_Rel_Diarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_aircadcia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_aircadcia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_aircadformulario]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_aircadformulario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_aircadformularioitem]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_aircadformularioitem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_aircadlocal]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_aircadlocal]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_aircadtetc]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_aircadtetc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_airtabpreco]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_airtabpreco]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_airtabprecogeral]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_airtabprecogeral]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_airtabprecotetc]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_airtabprecotetc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_cadcli]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_cadcli]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_cadprazo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_cadprazo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_caduf]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_caduf]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_cadusu]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_cadusu]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_codocorr]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_codocorr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_ctc_esp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_ctc_esp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_genaereo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_genaereo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_logusuario]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_logusuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_manifesto]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_manifesto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_mem]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_mem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_nf_esp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_nf_esp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_ocorr]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_ocorr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_redaereo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_redaereo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_redaereoitem]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_redaereoitem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_tabaereo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_tabaereo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_tabaereoitem]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_tabaereoitem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_transpsub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_transpsub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_ultconssac]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_ultconssac]
GO

CREATE TABLE [dbo].[BASE] (
	[MINUTA] [nvarchar] (7) COLLATE Latin1_General_CI_AS NULL ,
	[REMET_CGC] [nvarchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[REMET_NOME] [nvarchar] (40) COLLATE Latin1_General_CI_AS NULL ,
	[NUMNF] [nvarchar] (6) COLLATE Latin1_General_CI_AS NULL ,
	[DATA] [nvarchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[HORA] [nvarchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[ENTREGA] [nvarchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[RECEBEDOR] [nvarchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[STATUS] [nvarchar] (30) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[est_clientes_bomi] (
	[nome_cliente] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[cgc_cliente] [varchar] (16) COLLATE Latin1_General_CI_AS NULL ,
	[modal] [varchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[diretorio] [varchar] (100) COLLATE Latin1_General_CI_AS NULL ,
	[arquivoA] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[arquivoR] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[base] [varchar] (10) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[est_faixas_de_peso] (
	[peso_inicial] [numeric](19, 4) NULL ,
	[peso_final] [numeric](19, 4) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[est_faixas_de_valor] (
	[valor_inicial] [numeric](19, 4) NULL ,
	[valor_final] [numeric](19, 4) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[prz_clientes] (
	[nome_cliente] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[cgc_cliente] [varchar] (16) COLLATE Latin1_General_CI_AS NULL ,
	[modal] [varchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[diretorio] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[arquivoA] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[arquivoR] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[base] [varchar] (10) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[prz_codoc] (
	[COD] [smallint] NULL ,
	[DESCRICAO] [nvarchar] (81) COLLATE Latin1_General_CI_AS NULL ,
	[COMP] [nvarchar] (6) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[prz_cortea] (
	[REGIAO] [nvarchar] (13) COLLATE Latin1_General_CI_AS NULL ,
	[HORA] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[prz_corter] (
	[REGIAO] [nvarchar] (13) COLLATE Latin1_General_CI_AS NULL ,
	[HORA] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[salete] (
	[nf] [nvarchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[filial] [varchar] (2) COLLATE Latin1_General_CI_AS NULL ,
	[ctc] [varchar] (6) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_CadFeriado] (
	[codigo] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[ano] [numeric](18, 0) NULL ,
	[mes] [numeric](18, 0) NOT NULL ,
	[dia] [numeric](18, 0) NOT NULL ,
	[descricao] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[uf] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cidade] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[tipo] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_Rel_Diarios] (
	[nome_cliente] [varchar] (40) COLLATE Latin1_General_CI_AS NULL ,
	[cgc_cliente] [varchar] (14) COLLATE Latin1_General_CI_AS NULL ,
	[nome_arquivo] [varchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[tipo] [varchar] (3) COLLATE Latin1_General_CI_AS NULL ,
	[diretorio] [varchar] (30) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_aircadcia] (
	[codcia] [varchar] (3) COLLATE Latin1_General_CI_AS NOT NULL ,
	[fantasia] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[descricao] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[estoqueminimo] [int] NULL ,
	[estoqueatual] [int] NULL ,
	[proximonum] [int] NULL ,
	[avisominimo] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[datacadastro] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_aircadformulario] (
	[idcadform] [numeric](10, 0) IDENTITY (1, 1) NOT NULL ,
	[codcia] [varchar] (3) COLLATE Latin1_General_CI_AS NULL ,
	[numinicial] [numeric](18, 0) NULL ,
	[numfinal] [numeric](18, 0) NULL ,
	[datacadastro] [datetime] NULL ,
	[usuariocad] [varchar] (10) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_aircadformularioitem] (
	[idcadform] [numeric](18, 0) NOT NULL ,
	[numero] [numeric](18, 0) NOT NULL ,
	[codcia] [varchar] (3) COLLATE Latin1_General_CI_AS NULL ,
	[status] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[datastatus] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_aircadlocal] (
	[sigla] [varchar] (3) COLLATE Latin1_General_CI_AS NOT NULL ,
	[localidade] [varchar] (40) COLLATE Latin1_General_CI_AS NULL ,
	[uf] [varchar] (2) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_aircadtetc] (
	[codigo] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[descricao] [varchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_airtabpreco] (
	[idtabela] [int] IDENTITY (1, 1) NOT NULL ,
	[codcia] [varchar] (3) COLLATE Latin1_General_CI_AS NULL ,
	[descricao] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[status] [varchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[vigenciainic] [datetime] NULL ,
	[vigenciafim] [datetime] NULL ,
	[cadastro] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_airtabprecogeral] (
	[idgeral] [int] IDENTITY (1, 1) NOT NULL ,
	[idtabela] [int] NULL ,
	[sigladest] [varchar] (3) COLLATE Latin1_General_CI_AS NULL ,
	[taxaminima] [money] NULL ,
	[ate25] [money] NULL ,
	[ate50] [money] NULL ,
	[ate300] [money] NULL ,
	[ate500] [money] NULL ,
	[ate1000] [money] NULL ,
	[acima1000] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_airtabprecotetc] (
	[idtetc] [int] IDENTITY (1, 1) NOT NULL ,
	[idtabela] [int] NULL ,
	[sigladest] [varchar] (3) COLLATE Latin1_General_CI_AS NULL ,
	[codtetc] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[taxaminima] [money] NULL ,
	[porkilo] [money] NULL ,
	[usargeral] [varchar] (1) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_cadcli] (
	[cgc] [varchar] (14) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[nome] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[endereco] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cidade] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[uf] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ie] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[prazo] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[env_emailOco] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[env_emailFer] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[email1] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[email2] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[email3] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[email4] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[email5] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Cli_Dest] [varchar] (1) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_cadprazo] (
	[codigo] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[uf] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[modal] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[prazo_cap] [numeric](18, 0) NULL ,
	[prazo_int] [numeric](18, 0) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_caduf] (
	[uf] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cidade] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_cadusu] (
	[Usuario] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[senha] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Nome] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Departamento] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DataCad] [datetime] NULL ,
	[status] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[stringdireitos] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[expirada] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_codocorr] (
	[cod_ocorr] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[descricao] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[abonaSN] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[env_email] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[email1] [varchar] (35) COLLATE Latin1_General_CI_AS NULL ,
	[email2] [varchar] (35) COLLATE Latin1_General_CI_AS NULL ,
	[email3] [varchar] (35) COLLATE Latin1_General_CI_AS NULL ,
	[email4] [varchar] (35) COLLATE Latin1_General_CI_AS NULL ,
	[email_cliente] [varchar] (1) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_ctc_esp] (
	[filialctc] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[filial] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ctc] [int] NOT NULL ,
	[data] [datetime] NULL ,
	[hora] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[emissor] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[prev_entrega] [datetime] NULL ,
	[remet_cgc] [varchar] (14) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[remet_nome] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[respons_cgc] [varchar] (14) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dest_cgc] [varchar] (14) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dest_nome] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cidade_orig] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cidade_dest] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[uf_dest] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[regiao] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[nfs] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[valmerc] [money] NULL ,
	[peso] [numeric](18, 0) NULL ,
	[volumes] [numeric](18, 0) NULL ,
	[especie] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[natureza] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dimensoes] [varchar] (29) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[fretenacional] [money] NULL ,
	[advalorem] [money] NULL ,
	[txorigem] [money] NULL ,
	[txdestino] [money] NULL ,
	[txredespacho] [money] NULL ,
	[txcoleta] [money] NULL ,
	[txoutros] [money] NULL ,
	[fretetotal] [money] NULL ,
	[tribut] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[aliquota] [money] NULL ,
	[filialmanifesto] [varchar] (8) COLLATE Latin1_General_CI_AS NULL ,
	[fpag] [varchar] (7) COLLATE Latin1_General_CI_AS NULL ,
	[faturar] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[faturanum] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[entrega_data] [datetime] NULL ,
	[recebedor] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[entrega_hora] [char] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[modal] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[obs_emissao] [varchar] (186) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[obs_ocorr] [varchar] (186) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[arq] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[tem_ocorr] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[transp_sub] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[cancelado] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[canc_data] [datetime] NULL ,
	[canc_usu] [varchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[canc_obs] [varchar] (60) COLLATE Latin1_General_CI_AS NULL ,
	[at_ctc_cif] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[at_ctc_cif_data] [datetime] NULL ,
	[codevo] [int] NULL ,
	[at_devol] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[at_edi] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[at_edi_data] [datetime] NULL ,
	[at_edisub] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[at_edisub_data] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_genaereo] (
	[codigo] [varchar] (3) COLLATE Latin1_General_CI_AS NOT NULL ,
	[status] [varchar] (7) COLLATE Latin1_General_CI_AS NOT NULL ,
	[vigenciainic] [datetime] NOT NULL ,
	[vigenciafim] [datetime] NULL ,
	[descricao] [varchar] (30) COLLATE Latin1_General_CI_AS NOT NULL ,
	[fretevalor] [money] NOT NULL ,
	[limitecoleta] [money] NOT NULL ,
	[txcoleta] [money] NOT NULL ,
	[acimalimitecoleta] [money] NOT NULL ,
	[limiteentrega] [money] NOT NULL ,
	[txentrega] [money] NOT NULL ,
	[acimalimiteentrega] [money] NOT NULL ,
	[aplicacubagem] [varchar] (1) COLLATE Latin1_General_CI_AS NOT NULL ,
	[obs1] [varchar] (30) COLLATE Latin1_General_CI_AS NULL ,
	[obs2] [varchar] (30) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_logusuario] (
	[id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[acao] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[data] [datetime] NULL ,
	[usuario] [varchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[descricao] [varchar] (100) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_manifesto] (
	[idcodigo] [int] IDENTITY (1, 1) NOT NULL ,
	[filialctc] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[filialmanifesto] [varchar] (8) COLLATE Latin1_General_CI_AS NOT NULL ,
	[filial] [varchar] (2) COLLATE Latin1_General_CI_AS NULL ,
	[manifesto] [int] NULL ,
	[embarcador] [varchar] (14) COLLATE Latin1_General_CI_AS NULL ,
	[dtemissao] [datetime] NULL ,
	[hsemissao] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[dtsaida] [datetime] NULL ,
	[hssaida] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[placaveic] [varchar] (8) COLLATE Latin1_General_CI_AS NULL ,
	[motorista] [varchar] (25) COLLATE Latin1_General_CI_AS NULL ,
	[cidade_orig] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[at_manif_cif] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[at_manif_cif_data] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_mem] (
	[idmem] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[ctrrelprotocolo] [numeric](18, 0) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_nf_esp] (
	[idcodigo] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[filialctc] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[numnfnum] [numeric](18, 0) NULL ,
	[numnf] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cliente_cgc] [varchar] (14) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cliente_nome] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[emissao_nf] [datetime] NULL ,
	[numpedido] [varchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[dtpedido] [datetime] NULL ,
	[valornf] [money] NULL ,
	[at_cliente] [varchar] (1) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_ocorr] (
	[codigo] [numeric](10, 0) IDENTITY (1, 1) NOT NULL ,
	[filialctc] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cod_ocorr] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[emissaoctc] [datetime] NULL ,
	[remet_cgc] [varchar] (14) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[descr_ocorr] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[usu_ocorr] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[usu_dataocorr] [datetime] NULL ,
	[data] [datetime] NULL ,
	[hora] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[obs_ocorr] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[baixadopre] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtbaixapre] [datetime] NULL ,
	[hsbaixapre] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[recebpre] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[usu_bxpre] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[usu_datapre] [datetime] NULL ,
	[baixadofinal] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[dtbaixa] [datetime] NULL ,
	[hsbaixa] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[receb] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[usu_bx] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[usu_databx] [datetime] NULL ,
	[atualprazo] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[prazoentr] [numeric](18, 0) NULL ,
	[diasuteis] [numeric](18, 0) NULL ,
	[email_enviadoCLI] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[email_dataCLI] [datetime] NULL ,
	[email_enviadoSAC] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[email_dataSAC] [datetime] NULL ,
	[email_enviadoINT] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[email_dataINT] [datetime] NULL ,
	[atual_sitla] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[at_sitla_data] [datetime] NULL ,
	[rel_arquivo] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[rel_arq_data] [datetime] NULL ,
	[rel_arq_num] [numeric](18, 0) NULL ,
	[at_ocorr_cif] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[at_ocorr_cif_data] [datetime] NULL ,
	[canhotonf] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[cartacanhoto] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[datacartacanhoto] [datetime] NULL ,
	[at_edi] [varchar] (1) COLLATE Latin1_General_CI_AS NULL ,
	[at_edi_data] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_redaereo] (
	[idredesp] [int] IDENTITY (1, 1) NOT NULL ,
	[codigo] [varchar] (2) COLLATE Latin1_General_CI_AS NULL ,
	[status] [varchar] (7) COLLATE Latin1_General_CI_AS NULL ,
	[vigenciainic] [datetime] NULL ,
	[vigenciafim] [datetime] NULL ,
	[descricao] [varchar] (30) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_redaereoitem] (
	[idredesp] [int] NOT NULL ,
	[uf_dest] [varchar] (2) COLLATE Latin1_General_CI_AS NOT NULL ,
	[regiao] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[limitepeso] [money] NULL ,
	[taxaredesp] [money] NULL ,
	[acimalimite] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_tabaereo] (
	[idtabela] [int] IDENTITY (1, 1) NOT NULL ,
	[codigo] [varchar] (3) COLLATE Latin1_General_CI_AS NOT NULL ,
	[status] [varchar] (7) COLLATE Latin1_General_CI_AS NOT NULL ,
	[descricao] [varchar] (30) COLLATE Latin1_General_CI_AS NOT NULL ,
	[vigenciainic] [datetime] NOT NULL ,
	[vigenciafim] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_tabaereoitem] (
	[idtabela] [int] NOT NULL ,
	[sigla] [varchar] (3) COLLATE Latin1_General_CI_AS NOT NULL ,
	[localidade] [varchar] (25) COLLATE Latin1_General_CI_AS NOT NULL ,
	[taxamin] [money] NOT NULL ,
	[pesomin] [money] NOT NULL ,
	[porkilo] [money] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_transpsub] (
	[texto] [varchar] (12) COLLATE Latin1_General_CI_AS NULL ,
	[transportador] [varchar] (20) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_ultconssac] (
	[filialctc] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[usuario] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[data] [datetime] NOT NULL ,
	[operacao] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tb_CadFeriado] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_CadFeriado] PRIMARY KEY  CLUSTERED 
	(
		[codigo]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_aircadcia] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_airciaaerea] PRIMARY KEY  CLUSTERED 
	(
		[codcia]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_aircadformulario] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_aircadformulario] PRIMARY KEY  CLUSTERED 
	(
		[idcadform]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_aircadformularioitem] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_aircadformularioitem] PRIMARY KEY  CLUSTERED 
	(
		[idcadform],
		[numero]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_aircadlocal] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_airlocal] PRIMARY KEY  CLUSTERED 
	(
		[sigla]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_airtabpreco] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_airtabaerea] PRIMARY KEY  CLUSTERED 
	(
		[idtabela]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_airtabprecogeral] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_airtabgeral] PRIMARY KEY  CLUSTERED 
	(
		[idgeral]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_airtabprecotetc] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_airtabtetc] PRIMARY KEY  CLUSTERED 
	(
		[idtetc]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_caduf] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_caduf] PRIMARY KEY  CLUSTERED 
	(
		[uf],
		[cidade]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_codocorr] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_codocorr] PRIMARY KEY  CLUSTERED 
	(
		[cod_ocorr]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_logusuario] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_logusuario] PRIMARY KEY  CLUSTERED 
	(
		[id]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_manifesto] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_manifesto] PRIMARY KEY  CLUSTERED 
	(
		[idcodigo]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_nf_esp] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_nf_esp] PRIMARY KEY  CLUSTERED 
	(
		[idcodigo]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_ocorr] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_ocorr] PRIMARY KEY  CLUSTERED 
	(
		[codigo]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_redaereo] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_redaereo] PRIMARY KEY  CLUSTERED 
	(
		[idredesp]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_redaereoitem] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_redaereoitem] PRIMARY KEY  CLUSTERED 
	(
		[idredesp],
		[uf_dest]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_tabaereo] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_tabaereo] PRIMARY KEY  CLUSTERED 
	(
		[idtabela]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_tabaereoitem] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_tabaereoitem] PRIMARY KEY  CLUSTERED 
	(
		[idtabela],
		[sigla]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_cadcli] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_cadcli] PRIMARY KEY  NONCLUSTERED 
	(
		[cgc]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_cadusu] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_cadusu] PRIMARY KEY  NONCLUSTERED 
	(
		[Usuario]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tb_ctc_esp] WITH NOCHECK ADD 
	CONSTRAINT [PK_tb_ctc_esp] PRIMARY KEY  NONCLUSTERED 
	(
		[filialctc]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

 CREATE  INDEX [IX_tb_CadFeriado] ON [dbo].[tb_CadFeriado]([uf], [cidade]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_airtabaerea] ON [dbo].[tb_airtabpreco]([codcia]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_airtabaerea_1] ON [dbo].[tb_airtabpreco]([status]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_airtabgeral] ON [dbo].[tb_airtabprecogeral]([idtabela]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_airtabgeral_2] ON [dbo].[tb_airtabprecogeral]([sigladest]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_airtabtetc] ON [dbo].[tb_airtabprecotetc]([idtabela]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_airtabtetc_2] ON [dbo].[tb_airtabprecotetc]([sigladest]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_airtabtetc_4] ON [dbo].[tb_airtabprecotetc]([codtetc]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_cadcli] ON [dbo].[tb_cadcli]([nome]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_cadcli_1] ON [dbo].[tb_cadcli]([Cli_Dest]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_cadprazo] ON [dbo].[tb_cadprazo]([codigo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_codocorr] ON [dbo].[tb_codocorr]([descricao]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_1] ON [dbo].[tb_ctc_esp]([remet_cgc]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_2] ON [dbo].[tb_ctc_esp]([tem_ocorr]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_3] ON [dbo].[tb_ctc_esp]([uf_dest], [cidade_dest]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_4] ON [dbo].[tb_ctc_esp]([modal]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp] ON [dbo].[tb_ctc_esp]([data]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_5] ON [dbo].[tb_ctc_esp]([arq]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_6] ON [dbo].[tb_ctc_esp]([filialmanifesto]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_7] ON [dbo].[tb_ctc_esp]([codevo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_8] ON [dbo].[tb_ctc_esp]([at_devol]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_9] ON [dbo].[tb_ctc_esp]([at_edi]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ctc_esp_10] ON [dbo].[tb_ctc_esp]([at_edisub]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_genaereo] ON [dbo].[tb_genaereo]([codigo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_genaereo_1] ON [dbo].[tb_genaereo]([status]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_genaereo_2] ON [dbo].[tb_genaereo]([descricao]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_logusuario] ON [dbo].[tb_logusuario]([acao]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_logusuario_1] ON [dbo].[tb_logusuario]([data]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_logusuario_2] ON [dbo].[tb_logusuario]([usuario]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_manifesto] ON [dbo].[tb_manifesto]([filialctc]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_manifesto_1] ON [dbo].[tb_manifesto]([filialmanifesto]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_manifesto_2] ON [dbo].[tb_manifesto]([placaveic]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_manifesto_3] ON [dbo].[tb_manifesto]([manifesto]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_nf_esp] ON [dbo].[tb_nf_esp]([filialctc]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_nf_esp_1] ON [dbo].[tb_nf_esp]([numnf]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_nf_esp_2] ON [dbo].[tb_nf_esp]([cliente_cgc]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_nf_esp_3] ON [dbo].[tb_nf_esp]([numnfnum]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ocorr] ON [dbo].[tb_ocorr]([filialctc]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ocorr_1] ON [dbo].[tb_ocorr]([cod_ocorr]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ocorr_2] ON [dbo].[tb_ocorr]([emissaoctc]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ocorr_3] ON [dbo].[tb_ocorr]([remet_cgc]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ocorr_4] ON [dbo].[tb_ocorr]([email_enviadoCLI]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ocorr_5] ON [dbo].[tb_ocorr]([email_enviadoSAC]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_ocorr_6] ON [dbo].[tb_ocorr]([email_enviadoINT]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_redaereo] ON [dbo].[tb_redaereo]([codigo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_redaereo_1] ON [dbo].[tb_redaereo]([status]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_redaereo_2] ON [dbo].[tb_redaereo]([descricao]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_redaereoitem] ON [dbo].[tb_redaereoitem]([regiao]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_tabaereo] ON [dbo].[tb_tabaereo]([codigo]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_tabaereo_1] ON [dbo].[tb_tabaereo]([status]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tb_tabaereoitem] ON [dbo].[tb_tabaereoitem]([localidade]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[tb_airtabpreco] ADD 
	CONSTRAINT [FK_tb_airtabaerea_tb_airciaaerea] FOREIGN KEY 
	(
		[codcia]
	) REFERENCES [dbo].[tb_aircadcia] (
		[codcia]
	)
GO

ALTER TABLE [dbo].[tb_airtabprecogeral] ADD 
	CONSTRAINT [FK_tb_airtabgeral_tb_airtabaerea] FOREIGN KEY 
	(
		[idtabela]
	) REFERENCES [dbo].[tb_airtabpreco] (
		[idtabela]
	)
GO

ALTER TABLE [dbo].[tb_airtabprecotetc] ADD 
	CONSTRAINT [FK_tb_airtabtetc_tb_airtabaerea] FOREIGN KEY 
	(
		[idtabela]
	) REFERENCES [dbo].[tb_airtabpreco] (
		[idtabela]
	)
GO

ALTER TABLE [dbo].[tb_ctc_esp] ADD 
	CONSTRAINT [FK_tb_ctc_esp_tb_cadcli] FOREIGN KEY 
	(
		[remet_cgc]
	) REFERENCES [dbo].[tb_cadcli] (
		[cgc]
	)
GO

ALTER TABLE [dbo].[tb_manifesto] ADD 
	CONSTRAINT [FK_tb_manifesto_tb_ctc_esp] FOREIGN KEY 
	(
		[filialctc]
	) REFERENCES [dbo].[tb_ctc_esp] (
		[filialctc]
	)
GO

ALTER TABLE [dbo].[tb_nf_esp] ADD 
	CONSTRAINT [FK_tb_nf_esp_tb_ctc_esp] FOREIGN KEY 
	(
		[filialctc]
	) REFERENCES [dbo].[tb_ctc_esp] (
		[filialctc]
	)
GO

ALTER TABLE [dbo].[tb_ocorr] ADD 
	CONSTRAINT [FK_tb_ocorr_tb_cadusu] FOREIGN KEY 
	(
		[usu_bx]
	) REFERENCES [dbo].[tb_cadusu] (
		[Usuario]
	),
	CONSTRAINT [FK_tb_ocorr_tb_codocorr] FOREIGN KEY 
	(
		[cod_ocorr]
	) REFERENCES [dbo].[tb_codocorr] (
		[cod_ocorr]
	),
	CONSTRAINT [FK_tb_ocorr_tb_ctc_esp] FOREIGN KEY 
	(
		[filialctc]
	) REFERENCES [dbo].[tb_ctc_esp] (
		[filialctc]
	)
GO

ALTER TABLE [dbo].[tb_redaereoitem] ADD 
	CONSTRAINT [FK_tb_redaereoitem_tb_redaereo] FOREIGN KEY 
	(
		[idredesp]
	) REFERENCES [dbo].[tb_redaereo] (
		[idredesp]
	)
GO

ALTER TABLE [dbo].[tb_tabaereoitem] ADD 
	CONSTRAINT [FK_tb_tabaereoitem_tb_tabaereo] FOREIGN KEY 
	(
		[idtabela]
	) REFERENCES [dbo].[tb_tabaereo] (
		[idtabela]
	)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

create procedure 	sp_alt_ExportSitlaNao	@filialctc 	varchar(10)
as update 		tb_ocorr
set 			atual_sitla = 'N' 
where 			filialctc = @filialctc and 
			cod_ocorr = '01'
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure sp_alt_ObsOcorr_Sitla 		@obs_ocorr varchar(186),
						@filialctc varchar(10)
as update 	tb_ctc_esp
set 		obs_ocorr = @obs_ocorr 
where 		filialctc = @filialctc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO



CREATE procedure sp_alt_cadcli		@cgc		varchar(14),
					@nome		varchar(40),
					@endereco	varchar(40),
					@cidade	varchar(20),
					@uf		varchar(02),
					@ie		varchar(20),
					@prazo		varchar(06),
					@env_emailOco	varchar(01),
					@env_emailFer	varchar(01),
					@email1	varchar(35),
					@email2	varchar(35),
					@email3	varchar(35),
					@email4	varchar(35),
					@email5	varchar(35)
as update tb_cadcli set 			nome = @nome,
					endereco = @endereco,
					cidade = @cidade,
					uf = @uf,
					ie = @ie,
					prazo = @prazo,
					env_emailOco = @env_emailOco,
					env_emailFer = @env_emailFer,
					email1 = @email1,
					email2 = @email2,
					email3 = @email3,
					email4 = @email4,
					email5 = @email5
where
					cgc = @cgc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure sp_alt_cadcli_imp 	@cgc		varchar(14),
					@nome		varchar(40),
					@endereco	varchar(40),
					@cidade	varchar(20),
					@uf		varchar(02),
					@ie		varchar(20)
as update tb_cadcli set 			nome = @nome,
					endereco = @endereco,
					cidade = @cidade,
					uf = @uf,
					ie = @ie
where
					cgc = @cgc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


create procedure sp_alt_cadferiado		@codigo	int,
						@ano		int,
						@mes		int,
						@dia		int,
						@descricao	varchar(40),
						@cidade	varchar(20),
						@uf		varchar(02),
						@tipo		varchar(01)
as update tb_cadferiado set 	ano = @ano,
				mes = @mes,
				dia = @dia,
				descricao = @descricao,
				uf = @uf,
				cidade = @cidade,
				tipo = @tipo
where				codigo = @codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure sp_alt_cadocorr 	@cod_ocorr	varchar(02),
					@abonaSN	varchar(01),
					@env_email	varchar(01),
					@email1	varchar(35),
					@email2	varchar(35),
					@email3	varchar(35),
					@email4	varchar(35),
					@email_cliente	varchar(01)
as
update tb_codocorr set 			abonaSN = @abonaSN,
					env_email = @env_email,
					email1 = @email1,
					email2 = @email2,
					email3 = @email3,
					email4 = @email4,
					email_cliente = @email_cliente
where
					cod_ocorr = @cod_ocorr
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


create procedure sp_alt_cadusu		@usuario 	varchar(10),
					@nome 	varchar(50),
					@departamento	varchar(25),
					@status		varchar(01),
					@stringdireitos	varchar(100)
as update tb_cadusu set 
					 nome = @nome,
					 departamento = @departamento,
					 status = @status,
					 stringdireitos = @stringdireitos
where 					 usuario = @usuario

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure sp_alt_diasprazo	@codigo	numeric,
					@prazoentr	numeric,
					@diasuteis	numeric
as update tb_ocorr set	diasuteis = @diasuteis, atualprazo = 'N', prazoentr = @prazoentr
where			codigo = @codigo

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure sp_alt_manif	 	@filialmanifesto	varchar(08),
					@filialctc	varchar(10),
					@placaveic	varchar(08),
					@motorista	varchar(25)
as update tb_manifesto with (rowlock) set
				placaveic = @placaveic,
				motorista = @motorista
where
				filialmanifesto = @filialmanifesto and
				filialctc = @filialctc
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure sp_alt_obsentr	@filialctc 	varchar(10),
					@cod_ocorr	varchar(02),
					@obs_ocorr	varchar(300)
as update tb_ocorr with (rowlock) set	
					obs_ocorr = @obs_ocorr
where
					filialctc = @filialctc and
					cod_ocorr = @cod_ocorr

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_alt_obsocorr	@filialctc 	varchar(10),
					@cod_ocorr	varchar(02),
					@data		datetime,
					@hora		varchar(05),
					@obs_ocorr	varchar(300)
as update tb_ocorr with (rowlock) set	
					obs_ocorr = @obs_ocorr
where
					filialctc = @filialctc and
					cod_ocorr = @cod_ocorr and
					data = @data and
					hora = @hora





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


/*   atualiza a pré baixa e os campos Data, Hora e Atualprazo=S. 
      Para casos em que há somente PréBaixa e Não há Baixa Final  */

CREATE procedure sp_alt_ocorr1 	@filialctc	varchar(10),
					@data		datetime,
					@hora		varchar(05),
					@dtbaixapre	datetime,
					@hsbaixapre	varchar(05),
					@recebpre	varchar(25),
					@usu_bxpre	varchar(10),
					@usu_datapre	datetime,
					@atual_sitla	varchar(01),
					@at_sitla_data	datetime
as update tb_ocorr with (rowlock) set
				data = @data,
				hora = @hora,
				dtbaixapre = @dtbaixapre,
				hsbaixapre = @hsbaixapre,
				recebpre = @recebpre,
				usu_bxpre = @usu_bxpre,
				usu_datapre = @usu_datapre,
				baixadopre = 'S',
				atualprazo = 'S',
				atual_sitla = @atual_sitla,
				at_sitla_data = @at_sitla_data
where
				filialctc = @filialctc and
				cod_ocorr = '01'
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


/*   atualiza a pré baixa sem os campos Data, Hora e Atualprazo
      Para casos que já tenham Baixa Final (baixa final é a informação mais confiável)  */

CREATE procedure sp_alt_ocorr1ow 	@filialctc	varchar(10),
					@dtbaixapre	datetime,
					@hsbaixapre	varchar(05),
					@recebpre	varchar(25),
					@usu_bxpre	varchar(10),
					@usu_datapre	datetime
as update tb_ocorr with (rowlock) set
				dtbaixapre = @dtbaixapre,
				hsbaixapre = @hsbaixapre,
				recebpre = @recebpre,
				usu_bxpre = @usu_bxpre,
				usu_datapre = @usu_datapre,
				baixadopre = 'S'
where
				filialctc = @filialctc and
				cod_ocorr = '01'
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_alt_ocorr2 	@filialctc	varchar(10),
					@data		datetime,
					@hora		varchar(05),
					@dtbaixa	datetime,
					@hsbaixa	varchar(05),
					@receb		varchar(25),
					@usu_bx	varchar(10),
					@usu_databx	datetime,
					@atual_sitla	varchar(01),
					@at_sitla_data	datetime,
					@canhotonf	varchar(01)
as update tb_ocorr with (rowlock) set
				data = @data,
				hora = @hora,
				dtbaixa = @dtbaixa,
				hsbaixa = @hsbaixa,
				receb = @receb,
				usu_bx = @usu_bx,
				usu_databx = @usu_databx,
				baixadofinal = 'S',
				atualprazo = 'S',
				atual_sitla = @atual_sitla,
				at_sitla_data = @at_sitla_data,
				rel_arquivo = 'N',
				canhotonf = @canhotonf
where
				filialctc = @filialctc and
				cod_ocorr = '01'
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_alt_ocorrcom	@filialctc	varchar(10),
					@obs_ocorr	varchar(300)
as update tb_ocorr with (rowlock) set
					obs_ocorr = @obs_ocorr
where
					filialctc = @filialctc

				




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_alt_temocorr_sn 	@tem_ocorr	varchar(01),
					@filialctc	varchar(10)
as update tb_ctc_esp with (rowlock) set 	tem_ocorr = @tem_ocorr
where
					filialctc = @filialctc




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_excl_cadocorr	@cod_ocorr	varchar(02)
as delete from tb_codocorr with (rowlock) where
					cod_ocorr = @cod_ocorr



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_excl_ocorr	@codigo 	int
as delete from tb_ocorr with (rowlock) where
				codigo = @codigo
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure sp_incl_cadcli	@cgc		varchar(14),
					@nome		varchar(40),
					@endereco	varchar(40),
					@cidade	varchar(20),
					@uf		varchar(02),
					@ie		varchar(20)
as insert into tb_cadcli
					(cgc,
					nome,
					endereco,
					cidade,
					uf,
					ie,
					prazo,
					env_emailOCO,
					env_emailFER)
values
					(@cgc,
					@nome,
					@endereco,
					@cidade,
					@uf,
					@ie,
					'TAB000',
					'N',
					'N')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


create procedure sp_ins_cadferiado		@ano		int,
						@mes		int,
						@dia		int,
						@descricao	varchar(40),
						@cidade	varchar(20),
						@uf		varchar(02),
						@tipo		varchar(01)
as insert into tb_cadferiado	(ano,
				mes,
				dia,
				descricao,
				uf,
				cidade,
				tipo)
values
				(@ano,
				@mes,
				@dia,
				@descricao,
				@uf,
				@cidade,
				@tipo)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure sp_ins_cadocorr 	@cod_ocorr	varchar(02),
					@descricao	varchar(50),
					@abonaSN	varchar(01),
					@env_email	varchar(01),
					@email1	varchar(35),
					@email2	varchar(35),
					@email3	varchar(35),
					@email4	varchar(35),
					@email_cliente	varchar(01)
as
insert into tb_codocorr
					(cod_ocorr,
					descricao,
					abonaSN,
					env_email,
					email1,
					email2,
					email3,
					email4,
					email_cliente)
values
					(@cod_ocorr,
					@descricao,
					@abonaSN,
					@env_email,
					@email1,
					@email2,
					@email3,
					@email4,
					@email_cliente)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


create procedure sp_ins_cadprazo		@codigo	varchar(06),
						@uf		varchar(02),
						@modal		varchar(01),
						@prazo_cap	int,
						@prazo_int	int
as insert into tb_cadprazo	(codigo,
				uf,
				modal,
				prazo_cap,
				prazo_int)
values
			(@codigo,
			@uf,
			@modal,
			@prazo_cap,
			@prazo_int)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


create procedure sp_ins_cadusu		@usuario 	varchar(10),
					@senha 	varchar(10),
					@nome 	varchar(50),
					@departamento	varchar(25),
					@datacad	datetime,
					@status		varchar(01),
					@stringdireitos	varchar(100),
					@expirada	varchar(01)
as insert into tb_cadusu
					(usuario,
					 senha,
					 nome,
					 departamento,
					 datacad,
					 status,
					 stringdireitos,
					 expirada)
values
					(@usuario,
					 @senha,
					 @nome,
					 @departamento,
					 @datacad,
					 @status,
					 @stringdireitos,
					 @expirada)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure sp_ins_impctc	@filialctc	varchar(10),
					@filial		varchar(02),
					@ctc		int,
					@data		datetime,
					@hora		varchar(05),
					@prev_entrega	datetime,
					@remet_cgc	varchar(15),
					@remet_nome	varchar(40),
					@respons_cgc	varchar(15),
					@dest_cgc	varchar(15),
					@dest_nome	varchar(40),
					@cidade_orig	varchar(20),
					@via		varchar(03),
					@cidade_dest	varchar(20),
					@uf_dest	varchar(02),
					@regiao	varchar(15),
					@nfs		varchar(200),
					@valmerc	money,
					@peso		decimal(9,1),
					@volumes	decimal(9,0),
					@especie	varchar(20),
					@natureza	varchar(20),
					@dimensoes	varchar(29),
					@fretenacional	money,
					@advalorem	money,
					@txorigem	money,
					@txdestino	money,
					@txredespacho	money,
					@txcoleta	money,
					@txoutros	money,
					@fretetotal	money,
					@tribut		varchar(01),
					@aliquota	money,
					@faturar	varchar(01),
					@faturanum	varchar(10),
					@entrega_data	datetime,
					@recebedor	varchar(15),
					@entrega_hora	varchar(05),
					@modal		varchar(15),
					@obs_emissao	varchar(186),
					@obs_ocorr	varchar(124),
					@arq		varchar(12),
					@fpag		varchar(07),
					@emissor	varchar(20),
					@transp_sub	varchar(20)
as insert into tb_ctc_esp	
				 (filialctc,
				 filial,
				 ctc,
				 data,
				 hora,
				 prev_entrega,
				 remet_cgc,
				 remet_nome,
				 respons_cgc,
				 dest_cgc,
				 dest_nome,
				 cidade_orig,
				 via,
				 cidade_dest,
				 uf_dest,
				 regiao,
				 nfs,
				 valmerc,
				 peso,
				 volumes,
				 especie, 
				 natureza, 
				 dimensoes,
				 fretenacional,
				 advalorem,
				 txorigem,
				 txdestino,
				 txredespacho,
				 txcoleta,
				 txoutros,
				 fretetotal,
				 tribut,
				 aliquota,
				 faturar,
				 faturanum,
				 entrega_data,
				 recebedor,
				 entrega_hora,
				 modal,
				 obs_emissao,
				 obs_ocorr,
				 arq,
				 fpag,
				 emissor,
				 tem_ocorr,
				 transp_sub,
				 at_devol,
				 at_ctc_cif)
values
				(@filialctc,
				@filial,
				@ctc,
				@data,
				@hora,
				@prev_entrega,
				@remet_cgc,
				@remet_nome,
				@respons_cgc,
				@dest_cgc,
				@dest_nome,
				@cidade_orig,
				@via,
				@cidade_dest,
				@uf_dest,
				@regiao,
				@nfs,
				@valmerc,
				@peso,
				@volumes,
				@especie,
				@natureza,
				@dimensoes,
				@fretenacional,
				@advalorem,
				@txorigem,
				@txdestino,
				@txredespacho,
				@txcoleta,
				@txoutros,
				@fretetotal,
				@tribut,
				@aliquota,
				@faturar,
				@faturanum,
				@entrega_data,
				@recebedor,
				@entrega_hora,
				@modal	,
				@obs_emissao,
				@obs_ocorr,
				@arq,
				@fpag,
				@emissor,
				'N',
				@transp_sub,
				'',
				'')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE procedure sp_ins_impnf	@numnf	varchar(12),
				@numnfnum	numeric,
				@filialctc	varchar(10),
				@cliente_cgc	varchar(15),
				@cliente_nome	varchar(40)
as insert into tb_nf_esp
				(numnf,
				 numnfnum,
				 filialctc,
				 cliente_cgc,
				 cliente_nome)
values
				(@numnf,
				 @numnfnum,
				 @filialctc,
				 @cliente_cgc,
				 @cliente_nome)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure sp_ins_logusuario		@acao		varchar(15),
						@usuario	varchar(10),
						@descricao	varchar(100)
as insert into tb_logusuario	(acao,
				data,
				usuario,
				descricao)
values
				(@acao,
				 getdate(),
				 @usuario,
				 @descricao)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure sp_ins_manifesto	@filialmanifesto 	varchar(08),
					@filialctc	varchar(10),
					@filial		varchar(02),
					@manifesto	int,
					@embarcador	varchar(14),
					@dtemissao	datetime,
					@hsemissao	varchar(05),
					@dtsaida	datetime,
					@hssaida	varchar(05),
					@placaveic	varchar(08),
					@motorista	varchar(20),
					@cidade_orig	varchar(20)
as insert into tb_manifesto
					(filialmanifesto,
					 filialctc,
					 filial,
					 manifesto,
					 embarcador,
					 dtemissao,
					 hsemissao,
					 dtsaida,
					 hssaida,
					 placaveic,
					 motorista,
					 cidade_orig,
					 at_manif_cif)
values
					(@filialmanifesto,
					 @filialctc,
					 @filial,
					 @manifesto,
					 @embarcador,
					 @dtemissao,
					 @hsemissao,
					 @dtsaida,
					 @hssaida,
					 @placaveic,
					 @motorista,
					 @cidade_orig,
					 '')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure sp_ins_manifestoNull	@filialmanifesto 	varchar(08),
					@filialctc	varchar(10),
					@filial		varchar(02),
					@manifesto	int,
					@embarcador	varchar(14),
					@placaveic	varchar(08),
					@motorista	varchar(20),
					@cidade_orig	varchar(20)
as insert into tb_manifesto
					(filialmanifesto,
					 filialctc,
					 filial,
					 manifesto,
					 embarcador,
					 placaveic,
					 motorista,
					 cidade_orig,
					 at_manif_cif)
values
					(@filialmanifesto,
					 @filialctc,
					 @filial,
					 @manifesto,
					 @embarcador,
					 @placaveic,
					 @motorista,
					 @cidade_orig,
					 '')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_ins_ocorr1 	@filialctc	varchar(10),
					@emissaoctc	datetime,
					@remet_cgc	varchar(14),
					@cod_ocorr	varchar(02),
					@descr_ocorr	varchar(40),
					@data		datetime,
					@hora		varchar(05),
					@dtbaixapre	datetime,
					@hsbaixapre	varchar(05),
					@recebpre	varchar(25),
					@usu_bxpre	varchar(10),
					@usu_datapre	datetime,
					@atual_sitla	varchar(01),
					@at_sitla_data	datetime
as insert into tb_ocorr
				(filialctc,
				emissaoctc,
				remet_cgc,
				cod_ocorr,
				descr_ocorr,
				data,
				hora,
				dtbaixapre,
				hsbaixapre,
				recebpre,
				usu_bxpre,
				usu_datapre,
				baixadopre,
				atualprazo,
				atual_sitla,
				at_sitla_data,
				at_ocorr_cif)
values
				(@filialctc,
				@emissaoctc,
				@remet_cgc,
				@cod_ocorr,
				@descr_ocorr,
				@data,
				@hora,
				@dtbaixapre,
				@hsbaixapre,
				@recebpre,
				@usu_bxpre,
				@usu_datapre,
				'S',
				'S',
				@atual_sitla,
				@at_sitla_data,
				'')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure sp_ins_ocorr3 	@filialctc	varchar(10),
					@emissaoctc	datetime,
					@remet_cgc	varchar(14),
					@cod_ocorr	varchar(02),
					@descr_ocorr	varchar(40),
					@data		datetime,
					@hora		varchar(05),
					@dtbaixapre	datetime,
					@hsbaixapre	varchar(05),
					@recebpre	varchar(25),
					@dtbaixa	datetime,
					@hsbaixa	varchar(05),
					@receb		varchar(25),
					@usuario	varchar(10),
					@usu_data	datetime,
					@atual_sitla	varchar(01),
					@at_sitla_data	datetime,
					@canhotonf 	varchar(01)
as insert into tb_ocorr
				(filialctc,
				emissaoctc,
				remet_cgc,
				cod_ocorr,
				descr_ocorr,
				data,
				hora,
				dtbaixapre,
				hsbaixapre,
				recebpre,
				dtbaixa,
				hsbaixa,
				receb,
				usu_bxpre,
				usu_bx,
				usu_datapre,
				usu_databx,
				baixadopre,
				baixadofinal,
				atualprazo,
				atual_sitla,
				at_sitla_data,
				rel_arquivo,
  				at_ocorr_cif,
				canhotonf)
values
				(@filialctc,
				@emissaoctc,
				@remet_cgc,
				@cod_ocorr,
				@descr_ocorr,
				@data,
				@hora,
				@dtbaixapre,
				@hsbaixapre,
				@recebpre,
				@dtbaixa,
				@hsbaixa,
				@receb,
				@usuario,
				@usuario,
				@usu_data,
				@usu_data,
				'S',
				'S',
				'S',
				@atual_sitla,
				@at_sitla_data,
				'N',
				'',
				@canhotonf)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_ins_ocorr4 	@filialctc	varchar(10),
					@emissaoctc	datetime,
					@remet_cgc	varchar(14),
					@cod_ocorr	varchar(02),
					@descr_ocorr	varchar(40),
					@data		datetime,
					@hora		varchar(05),
					@usu_ocorr	varchar(10),
					@usu_dataocorr 
datetime
as insert into tb_ocorr
				(filialctc,
				emissaoctc,
				remet_cgc,
				cod_ocorr,
				descr_ocorr,
				data,
				hora,
				usu_ocorr,
				usu_dataocorr,
				email_enviadoCLI,
				email_enviadoSAC,
				email_enviadoINT,
				at_ocorr_cif)

values
				(@filialctc,
				@emissaoctc,
				@remet_cgc,
				@cod_ocorr,
				@descr_ocorr,
				@data,
				@hora,		
				@usu_ocorr,
				@usu_dataocorr,
				'N',
				'N',
				'N',
				'')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO



CREATE procedure sp_ins_ocorr4cod00 	@filialctc	varchar(10),
					@emissaoctc	datetime,
					@remet_cgc	varchar(14),
					@cod_ocorr	varchar(02),
					@descr_ocorr	varchar(40),
					@data		datetime,
					@hora		varchar(05),
					@usu_ocorr	varchar(10),
					@usu_dataocorr 
datetime
as insert into tb_ocorr
				(filialctc,
				emissaoctc,
				remet_cgc,
				cod_ocorr,
				descr_ocorr,
				data,
				hora,
				usu_ocorr,
				usu_dataocorr,
				email_enviadoCLI,
				email_enviadoSAC,
				email_enviadoINT,
				rel_arquivo,
				baixadofinal,
				usu_bx,
				usu_databx,
				at_ocorr_cif)
values
				(@filialctc,
				@emissaoctc,
				@remet_cgc,
				@cod_ocorr,
				@descr_ocorr,
				@data,
				@hora,		
				@usu_ocorr,
				@usu_dataocorr,
				'N',
				'N',
				'N',
				'N',
				'R',
				@usu_ocorr,
				@usu_dataocorr,
				'')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE procedure sp_ins_ultconssac	@filialctc	varchar(10),
					@usuario	varchar(10),
					@data		datetime
as insert into tb_ultconssac
				(filialctc,
				usuario,
				data)
values
				(@filialctc,
				@usuario,
				@data)



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure sp_sel_ctcNFeCGC 	@Cliente_CGC	varchar(14),
					@Numnfnum	varchar(12)
as  select 	* 
from 		tb_ctc_esp 
where 		filialctc in 
		(select filialctc 
		 from tb_nf_esp
		 where 	cliente_cgc like @cliente_cgc and 
			numnfnum = @numnfnum) and
                 	tem_ocorr <> 'C'
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

