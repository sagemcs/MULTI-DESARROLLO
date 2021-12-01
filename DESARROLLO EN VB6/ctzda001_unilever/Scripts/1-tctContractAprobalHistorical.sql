/****** Object:  Table [dbo].[tctContractAprobalHistorical]    Script Date: 07/05/2019 16:53:49 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tctContractAprobalHistorical]') AND type in (N'U'))
DROP TABLE [dbo].[tctContractAprobalHistorical]
GO

/****** Object:  Table [dbo].[tctContractAprobalHistorical]    Script Date: 07/05/2019 16:53:49 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tctContractAprobalHistorical](
	[HistoricalKey] [int] IDENTITY(1,1) NOT NULL,
	[ContractKey] [int] NOT NULL,
	[AprobalLevelKey] [int] NOT NULL,
	[UserID] [varchar](50) NOT NULL,
	[AprobalDate] [date] NOT NULL,
 CONSTRAINT [PK_tctContractAprobalHistorical] PRIMARY KEY CLUSTERED 
(
	[HistoricalKey] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


GRANT INSERT, DELETE, UPDATE,SELECT, REFERENCES ON tctContractAprobalHistorical TO ApplicationDBRole