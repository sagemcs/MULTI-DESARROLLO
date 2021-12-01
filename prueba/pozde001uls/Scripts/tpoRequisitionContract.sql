GO

/****** Object:  Table [dbo].[tpoRequisitionContract]    Script Date: 08/06/2019 16:17:58 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tpoRequisitionContract]') AND type in (N'U'))
DROP TABLE [dbo].[tpoRequisitionContract]
GO
GO

/****** Object:  Table [dbo].[tpoRequisitionContract]    Script Date: 08/06/2019 16:17:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tpoRequisitionContract](
	[ReqKey] [int] NOT NULL,
	[ContractKey] [int] NOT NULL,
 CONSTRAINT [PK_tpoRequisitionContract] PRIMARY KEY CLUSTERED 
(
	[ReqKey] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


GRANT  REFERENCES ,  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON tpoRequisitionContract TO ApplicationDBRole
GO
