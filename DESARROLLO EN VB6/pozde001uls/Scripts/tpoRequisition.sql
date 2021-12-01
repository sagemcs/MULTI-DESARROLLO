/****** Object:  View [dbo].[vdvRequisitionTrace]    Script Date: 07/02/2019 14:52:38 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
alter table tpoRequisition
	add  [Type] [int] null

GO

ALTER TABLE tpoRequisition SET (LOCK_ESCALATION = TABLE)
GO

GRANT  REFERENCES ,  SELECT ,  UPDATE ,  INSERT ,  DELETE  ON tpoRequisition TO ApplicationDBRole
GO