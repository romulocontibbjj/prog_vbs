if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_motoboys]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_motoboys]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tb_motos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tb_motos]
GO

CREATE TABLE [dbo].[tb_motoboys] (
	[cod] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ,
	[nome] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL ,
	[cpf] [varchar] (15) COLLATE Latin1_General_CI_AS NOT NULL ,
	[rg] [varchar] (15) COLLATE Latin1_General_CI_AS NOT NULL ,
	[endereco] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL ,
	[numero] [int] NOT NULL ,
	[bairro] [char] (5) COLLATE Latin1_General_CI_AS NULL ,
	[uf] [char] (2) COLLATE Latin1_General_CI_AS NULL ,
	[cnh] [varchar] (20) COLLATE Latin1_General_CI_AS NOT NULL ,
	[vencimento] [datetime] NULL ,
	[categoria] [char] (5) COLLATE Latin1_General_CI_AS NULL ,
	[fone] [varchar] (15) COLLATE Latin1_General_CI_AS NULL ,
	[celular] [varchar] (15) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tb_motos] (
	[placa] [varchar] (8) COLLATE Latin1_General_CI_AS NOT NULL ,
	[cod_motoboy] [int] NOT NULL ,
	[ano] [varchar] (10) COLLATE Latin1_General_CI_AS NULL ,
	[uf] [char] (2) COLLATE Latin1_General_CI_AS NULL ,
	[marca] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[modelo] [varchar] (20) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

