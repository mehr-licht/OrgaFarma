SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[alzheimer](
	[Nº de Registo] [numeric](7, 0) NULL,
	[Nome Comercial] [nvarchar](255) COLLATE Latin1_General_CI_AS NULL,
	[DCI] [nvarchar](255) COLLATE Latin1_General_CI_AS NULL,
	[Forma Farmacêutica] [nvarchar](255) COLLATE Latin1_General_CI_AS NULL,
	[Dosagem] [nvarchar](255) COLLATE Latin1_General_CI_AS NULL,
	[Apresentação] [nvarchar](255) COLLATE Latin1_General_CI_AS NULL,
	[% Comp# Regime especial] [numeric](3, 2) NULL
)

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[dados](
	[code] [numeric](8, 0) NOT NULL,
	[nome] [nvarchar](100) COLLATE Latin1_General_CI_AS NULL,
	[dci] [nvarchar](120) COLLATE Latin1_General_CI_AS NULL,
	[forma] [nvarchar](100) COLLATE Latin1_General_CI_AS NULL,
	[dose] [nvarchar](250) COLLATE Latin1_General_CI_AS NULL,
	[qty] [nvarchar](50) COLLATE Latin1_General_CI_AS NULL,
	[comp] [numeric](3, 0) NULL,
	[GH] [numeric](4, 0) NULL,
	[PVP] [numeric](8, 2) NULL,
	[PR] [numeric](8, 2) NULL,
	[gen] [bit] NOT NULL,
	[lab] [nvarchar](100) COLLATE Latin1_General_CI_AS NULL,
	[top5] [numeric](8, 2) NULL,
	[dci_obr] [bit] NOT NULL,
	[pvpold] [numeric](8, 2) NULL,
	[pvp-3] [numeric](8, 2) NULL,
	[CNPEM] [numeric](8, 0) NULL,
	[pvp-4] [numeric](8, 2) NULL,
	[PMA] [numeric](8, 2) NULL,
	[pvp-5] [numeric](8, 2) NULL,
	[pvp-6] [numeric](8, 2) NULL,
	[desp 4250] [bit] NULL,
	[desp 1234] [bit] NULL,
	[desp 21094] [bit] NULL,
	[desp 10279] [bit] NULL,
	[desp 10280] [bit] NULL,
	[desp 10910] [bit] NULL,
	[desp 14123] [bit] NULL,
	[lei 62010] [bit] NULL,
	[unitario] [bit] NULL,
PRIMARY KEY CLUSTERED 
(
	[code] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
)

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Query](
	[desp 4250] [bit] NOT NULL,
	[desp 1234] [bit] NOT NULL,
	[desp 10279] [bit] NOT NULL,
	[desp 10280] [bit] NOT NULL,
	[desp 21094] [bit] NOT NULL,
	[desp 14123] [bit] NOT NULL,
	[desp 10910] [bit] NOT NULL,
	[lei 62010] [bit] NOT NULL,
	[unitario] [bit] NOT NULL
)

GO
