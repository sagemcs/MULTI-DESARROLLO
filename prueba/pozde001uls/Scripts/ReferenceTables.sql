
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tpoReqStateItem]') AND type in (N'U'))
DROP TABLE [dbo].[tpoReqStateItem]
GO

/****** Object:  Table [dbo].[tpoReqStateItem]    Script Date: 07/02/2019 14:36:05 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tpoReqStateItem](
	[stateItemKey] [int] NOT NULL,
	[CompanyID] [varchar](3) NOT NULL,
	[stateItemID] [varchar](25) NOT NULL,
	[descStateItem] [varchar](150) NOT NULL,
 CONSTRAINT [PK_tpoReqStateItem] PRIMARY KEY CLUSTERED 
(
	[stateItemKey] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

GRANT  REFERENCES ,  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [tpoReqStateItem] TO ApplicationDBRole
GO

INSERT INTO tpoReqStateItem
	VALUES 
	(0,'ULS','Licitación','Estatus para las partidas de requisición que se encuentran en estatus de aceptadas para proceso de compra'),
	(1,'ULS','Dictamen Técnico','Estatus para las partidas de requisición que se encuentran en estatus de aceptadas para proceso de compra'),
	(2,'ULS','Legal','Estatus para las partidas de requisición que se encuentran en estatus de aceptadas para proceso de compra'),
	(3,'ULS','Comité Contratación','Estatus para las partidas de requisición que se encuentran en estatus de aceptadas para proceso de compra'),
	(4,'ULS','Inteligencia Comercial','Estatus para las partidas de requisición que se encuentran en estatus de aceptadas para proceso de compra')

GO
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tpoReqTypeBuyItem]') AND type in (N'U'))
DROP TABLE [dbo].[tpoReqTypeBuyItem]
GO

/****** Object:  Table [dbo].[tpoReqTypeBuyItem]    Script Date: 07/02/2019 14:39:21 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tpoReqTypeBuyItem](
	[typeBIKey] [int] NOT NULL,
	[CompanyID] [varchar](3) NOT NULL,
	[typeBIID] [varchar](20) NOT NULL,
	[descBItem] [varchar](150) NOT NULL,
 CONSTRAINT [PK_tpoReqTBItem] PRIMARY KEY CLUSTERED 
(
	[typeBIKey] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

GRANT  REFERENCES ,  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON tpoReqTypeBuyItem TO ApplicationDBRole
GO
INSERT INTO tpoReqTypeBuyItem
 	VALUES 
	(0,'ULS','Nacional','Tipo de Compra para productos contratados con proveedores nacionales'),
	(1,'ULS','Internacional','Tipo de Compra para productos contratados con proveedores internacionales')	
GO

/****** Object:  Table [dbo].[tpoReqStatus]    Script Date: 10/8/2020 12:42:00 PM ******/
DROP TABLE [dbo].[tpoReqStatus]
GO

/****** Object:  Table [dbo].[tpoReqStatus]    Script Date: 10/8/2020 12:42:00 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tpoReqStatus](
	[statusKey] [int] NOT NULL,
	[statusID] [varchar](20) NOT NULL,
	[statusDesc] [varchar](150) NOT NULL,
	[CompanyID] [varchar](3) NOT NULL,
 CONSTRAINT [PK_tpoReqStatus] PRIMARY KEY CLUSTERED 
(
	[statusKey] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
GRANT  REFERENCES ,  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON [dbo].[tpoReqStatus] TO ApplicationDBRole
GO

INSERT INTO [dbo].[tpoReqStatus] select 0,'Pendiente','Estatus para las requisiciones creadas recientemente','ULS'
INSERT INTO [dbo].[tpoReqStatus] select 1,'Autorizada','Estatus para autorizar las requisiciones para aprovación y posterior generación de ordenes de compra','ULS'
INSERT INTO [dbo].[tpoReqStatus] select 2,'Aceptada','Estatus para aceptar requisiciones para la generación de ordenes de compra','ULS'
INSERT INTO [dbo].[tpoReqStatus] select 3,'Cancelada','Estatus para cancelar requisiciones no aprovadas para la generación de ordenes de compra','ULS'