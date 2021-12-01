/****** Object:  Table [dbo].[tctContractUpgrades]    Script Date: 09/09/2019 12:41:17 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tctContractUpgrades]') AND type in (N'U'))
DROP TABLE [dbo].[tctContractUpgrades]
GO

/****** Object:  Table [dbo].[tctContractUpgrades]    Script Date: 09/09/2019 12:41:18 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tctContractUpgrades](
	[ContractKey] [int] NOT NULL,
	[VendorClassKey] [int] NOT NULL,
	[Duration] [int] NOT NULL,
	[CntctKey] [int] NOT NULL,
	[PmtTermsKey] [int] NOT NULL,
	[FOBKey] [int] NOT NULL,
	[SignatureDate] [date] NOT NULL,
	[StartDate] [date] NOT NULL,
 CONSTRAINT [PK_tctContractUpgrades] PRIMARY KEY CLUSTERED 
(
	[ContractKey] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


