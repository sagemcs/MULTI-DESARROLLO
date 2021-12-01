/****** Object:  Table [dbo].[tctContractLine]    Script Date: 06/20/2019 10:00:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tctContractLine]') AND type in (N'U'))
DROP TABLE [dbo].[tctContractLine]
GO

USE [Unilever_Suchel_Prod]
GO

/****** Object:  Table [dbo].[tctContractLine]    Script Date: 06/20/2019 10:00:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tctContractLine](
	[ContractLineKey] [int] NOT NULL,
	[ContractKey] [int] NOT NULL,
	[DeliveryTime] [int] NOT NULL,
	[Description] [varchar](50) NOT NULL,
	[ItemKey] [int] NOT NULL,
	[LineAmt] [decimal](15, 3) NOT NULL,
	[MaxLot] [decimal](15, 5) NOT NULL,
	[Qty] [decimal](15, 5) NOT NULL,
	[MinLot] [decimal](15, 5) NOT NULL,
	[RoundValue] [decimal](15, 5) NOT NULL,
	[UnitCost] [decimal](15, 5) NOT NULL,
	[UnitMeasKey] [int] NOT NULL,
	[Type] [int] NOT NULL,
	[QtyVariation] [decimal](15, 5) NOT NULL,
	[SeqNo] [int] NOT NULL,
	[CreateDate] [date] NOT NULL,
	[CreateUser] [varchar](50) NOT NULL,
	[UpdateDate] [date] NULL,
	[UpdateUser] [varchar](50) NULL,
 CONSTRAINT [PK_tctContractLine] PRIMARY KEY CLUSTERED 
(
	[ContractLineKey] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

GRANT INSERT, DELETE, SELECT, UPDATE, REFERENCES ON [dbo].[tctContractLine] TO ApplicationDBRole

