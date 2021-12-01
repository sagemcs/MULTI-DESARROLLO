
/****** Object:  Table [dbo].[tctContractLog]    Script Date: 09/09/2019 12:40:36 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tctContractLog]') AND type in (N'U'))
DROP TABLE [dbo].[tctContractLog]
GO


/****** Object:  Table [dbo].[tctContractLog]    Script Date: 09/09/2019 12:40:36 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tctContractLog](
	[ContractKey] [int] NOT NULL,
	[ContractID] [varchar](15) NOT NULL,
	[ContractNo] [varchar](50) NOT NULL,
	[CompanyID] [varchar](3) NOT NULL,
	[ContractAmt] [decimal](15, 3) NOT NULL,
	[CurrID] [varchar](3) NULL,
	[SeqNo] [int] NOT NULL,
	[State] [int] NOT NULL,
 CONSTRAINT [PK_tctContractLog] PRIMARY KEY CLUSTERED 
(
	[ContractKey] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


