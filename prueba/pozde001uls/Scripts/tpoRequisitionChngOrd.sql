/****** Object:  Table [dbo].[tpoRequisitionChngOrd]    Script Date: 08/27/2019 14:54:55 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[tpoRequisitionChngOrd]') AND type in (N'U'))
DROP TABLE [dbo].[tpoRequisitionChngOrd]
GO

/****** Object:  Table [dbo].[tpoRequisitionChngOrd]    Script Date: 08/27/2019 14:54:55 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[tpoRequisitionChngOrd](
	[ReqKey] [int] NOT NULL,
	[AirFrtJustified] [smallint] NOT NULL,
	[ApprovalDate] [datetime] NULL,
	[ApprovalStatus] [smallint] NOT NULL,
	[BuyerKey] [int] NULL,
	[CompanyID] [varchar](3) NOT NULL,
	[Contact] [varchar](40) NULL,
	[CreateDate] [datetime] NULL,
	[CreateType] [smallint] NOT NULL,
	[CreateUserID] [varchar](30) NULL,
	[CurrExchRate] [float] NOT NULL,
	[CurrExchSchdKey] [int] NULL,
	[CurrID] [varchar](3) NULL,
	[DfltPurchDeptKey] [int] NULL,
	[DfltRequestDate] [datetime] NULL,
	[DfltShipMethKey] [int] NULL,
	[DfltShipToAddrKey] [int] NULL,
	[DfltShipToWhseKey] [int] NULL,
	[DfltShipZoneKey] [int] NULL,
	[DfltTargetCompID] [varchar](3) NOT NULL,
	[Expedite] [smallint] NOT NULL,
	[ExpediteReasonKey] [int] NULL,
	[FirstPOIssueDate] [datetime] NULL,
	[FOBKey] [int] NULL,
	[FreightAmt] [decimal](15, 3) NOT NULL,
	[Hold] [smallint] NOT NULL,
	[HoldReason] [varchar](20) NULL,
	[IntrnlContact] [varchar](20) NULL,
	[Originator] [varchar](40) NULL,
	[PmtTermsKey] [int] NULL,
	[Printed] [smallint] NOT NULL,
	[PurchAddrKey] [int] NULL,
	[PurchAmt] [decimal](15, 3) NOT NULL,
	[PurchVendAddrKey] [int] NULL,
	[RemitToAddrKey] [int] NULL,
	[RemitToVendAddrKey] [int] NULL,
	[ReqFormKey] [int] NULL,
	[Status] [smallint] NOT NULL,
	[STaxAmt] [decimal](15, 3) NOT NULL,
	[STaxTranKey] [int] NULL,
	[TranAmt] [decimal](15, 3) NOT NULL,
	[TranAmtHC] [decimal](15, 3) NOT NULL,
	[TranCmnt] [varchar](50) NULL,
	[TranDate] [datetime] NOT NULL,
	[TranID] [varchar](13) NOT NULL,
	[TranNo] [varchar](10) NOT NULL,
	[TranType] [int] NOT NULL,
	[UpdateCounter] [int] NOT NULL,
	[UpdateDate] [datetime] NULL,
	[UpdateUserID] [varchar](30) NULL,
	[UsedFor] [varchar](40) NULL,
	[UserFld1] [varchar](15) NULL,
	[UserFld2] [varchar](15) NULL,
	[UserFld3] [varchar](15) NULL,
	[UserFld4] [varchar](15) NULL,
	[V1099Box] [varchar](3) NULL,
	[V1099BoxText] [varchar](15) NULL,
	[V1099Form] [smallint] NULL,
	[Type] [int] null,
	[ChangeReason] [varchar](100) NOT NULL,
	[ReasonCodeKey] [int] NOT NULL,
	[ChngNo] [int] NOT NULL,
 CONSTRAINT [PK__tpoRequi__44566AEF2B5AD8E8] PRIMARY KEY CLUSTERED 
(
	[ReqKey] ASC,
	[ChngNo] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

GRANT INSERT, DELETE, UPDATE, DELETE, REFERENCES ON tpoRequisitionChngOrd TO ApplicationDBRole