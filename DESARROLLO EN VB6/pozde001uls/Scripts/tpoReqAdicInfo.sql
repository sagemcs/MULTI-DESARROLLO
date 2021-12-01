/****** Object:  Table [dbo].[tpoReqAdicInfo]    Script Date: 07/02/2019 14:26:42 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tpoReqAdicInfo]') AND type in (N'U'))
DROP TABLE [dbo].[tpoReqAdicInfo]
GO

/****** Object:  Table [dbo].[tpoReqAdicInfo]    Script Date: 07/02/2019 14:26:43 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tpoReqAdicInfo](
	[ReqKeyIA] [int] NOT NULL,
	[statusKey] [int] NOT NULL,
	[autorizaReq] [varchar](50) NULL,
	[dateAutorReq] [datetime] NULL,
	[aceptaReq] [varchar](50) NULL,
	[dateAceptReq] [datetime] NULL,
	[rejectReq] [varchar](50) NULL,
	[dateRejectReq] [datetime] NULL,
	[segLvlAutoriza] [int] NOT NULL,
	[segAutorizaReq] [varchar](50) NULL,
	[segDateAutorReq] [datetime] NULL,
	[descriptionStatus] [text] NULL,
	[Type] [smallint] null,
 CONSTRAINT [PK_tpoReqAdicInfo] PRIMARY KEY CLUSTERED 
(
	[ReqKeyIA] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

ALTER TABLE [dbo].[tpoReqAdicInfo] ADD  CONSTRAINT [DF_tpoReqAdicInfo_statusReq]  DEFAULT ((1)) FOR [statusKey]
GO

ALTER TABLE [dbo].[tpoReqAdicInfo] ADD  CONSTRAINT [DF_tpoReqAdicInfo_2doAutoriza]  DEFAULT ((0)) FOR [segLvlAutoriza]
GO

GRANT  REFERENCES, SELECT, UPDATE, INSERT, DELETE ON tpoReqAdicInfo TO ApplicationDBRole

GO