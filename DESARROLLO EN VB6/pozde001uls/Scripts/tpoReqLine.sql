/****** Object:  View [dbo].[vdvRequisitionTrace]    Script Date: 07/02/2019 14:52:38 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
ALTER TABLE tpoReqLine ADD
	EstimatedPres decimal(18,2) DEFAULT ((0.00)) NOT NULL,
	ReqLineBuyerKey int NULL,
	TypeBIKey int NULL,
	StateBIKey int NULL
GO

ALTER TABLE tpoReqLine SET (LOCK_ESCALATION = TABLE)
GO

GRANT  REFERENCES ,  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON tpoReqLine TO ApplicationDBRole
GO