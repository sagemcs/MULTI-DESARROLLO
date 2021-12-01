/****** Object:  Table [dbo].[tctContractChgOrder]    Script Date: 08/23/2019 15:12:46 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tctContractChgOrder]') AND type in (N'U'))
DROP TABLE [dbo].[tctContractChgOrder]
GO

/****** Object:  Table [dbo].[tctContractChgOrder]    Script Date: 08/23/2019 15:12:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tctContractChgOrder](
	[ContractChgOrderKey] [int] IDENTITY(1,1) NOT NULL,
	[ContractKey] [int] NOT NULL,
	[CntctKey] [int] NOT NULL,
	[ChngOrdNo] [int] NOT NULL,
	[ChngOrdReason] [varchar](100) NULL,
	[ChngOrdReasonCodeKey] [int] NULL,
	[Clasification] [int] NOT NULL,
	[CompanyID] [varchar](3) NOT NULL,
	[ContractNo] [varchar](50) NOT NULL,
	[ContractID] [varchar](15) NOT NULL,
	[ContractAmt] [decimal](15, 3) NOT NULL,
	[CurrID] [varchar](3) NOT NULL,
	[CountryID] [varchar](3) NOT NULL,
	[Duration] [int] NOT NULL,
	[FOBKey] [int] NOT NULL,
	[Free] [smallint] NOT NULL,
	[ParentContractKey] [int] NULL,
	[PmtTermsKey] [int] NOT NULL,
	[SeqNo] [int] NOT NULL,
	[SignatureDate] [date] NOT NULL,
	[StartDate] [date] NOT NULL,
	[State] [int] NULL,
	[Type] [smallint] NOT NULL,
	[VendorKey] [int] NOT NULL,
	[CreateDate] [date] NOT NULL,
	[CreateUser] [varchar](50) NOT NULL,
	[UpdateDate] [date] NULL,
	[UpdateUser] [varchar](50) NULL,
	[VendClassKey] [int] NOT NULL,
 CONSTRAINT [PK_tctContractChgOrder] PRIMARY KEY CLUSTERED 
(
	[ContractChgOrderKey] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

GRANT INSERT, DELETE, UPDATE, SELECT, REFERENCES ON  [dbo].[tctContractChgOrder] TO ApplicationDBRole

