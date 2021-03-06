if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_cliente]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_cliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_moto]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_moto]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_motoboy]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_motoboy]
GO

CREATE TABLE [dbo].[tb_cliente] (
	[nome] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[cpf_cgc] [varchar] (15) COLLATE Latin1_General_CI_AS NOT NULL ,
	[nome_fantasia] [varchar] (30) COLLATE Latin1_General_CI_AS NULL ,
	[endereco] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[bairro] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[cidade] [varchar] (30) COLLATE Latin1_General_CI_AS NULL ,
	[uf] [char] (2) COLLATE Latin1_General_CI_AS NULL ,
	[telefone] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[cep] [char] (8) COLLATE Latin1_General_CI_AS NULL ,
	[contato] [varchar] (30) COLLATE Latin1_General_CI_AS NULL ,
	[email] [varchar] (40) COLLATE Latin1_General_CI_AS NULL ,
	[tempo_minimo] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_moto] (
	[cod_motoboy] [int] NULL ,
	[marca] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[modelo] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[ano] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[placa] [varchar] (7) COLLATE Latin1_General_CI_AS NOT NULL ,
	[cidade] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[uf] [char] (2) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_motoboy] (
	[cod_motoboy] [int] IDENTITY (1, 1) NOT NULL ,
	[nome] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL ,
	[cpf] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[rg] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[endereco] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL ,
	[cidade] [varchar] (30) COLLATE Latin1_General_CI_AS NULL ,
	[uf] [char] (2) COLLATE Latin1_General_CI_AS NULL ,
	[fone] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[celular] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[cnh] [varchar] (20) COLLATE Latin1_General_CI_AS NOT NULL ,
	[vencimento] [datetime] NULL ,
	[categoria] [char] (2) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

