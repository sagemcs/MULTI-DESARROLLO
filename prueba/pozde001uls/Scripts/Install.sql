DECLARE @strNo AS INT

select 1 from tsmDataDictTbl where TableName = 'tpoReqAdicInfo'
if @@ROWCOUNT = 0
begin
	insert into tsmDataDictTbl values ('tpoReqAdicInfo', 'ReqKeyIA', 'ReqKeyIA')
end 

select 1 from tsmDataDictCol where ColumnName = 'statusKey' and TableName = 'tpoReqAdicInfo'
if @@ROWCOUNT = 0
begin
	insert into tsmDataDictCol values ('tpoReqAdicInfo', 'statusKey', 'StaticCode', 0,0,0,0,0);
end 

select 1 from tsmDataDictCol where ColumnName = 'Type' and TableName = 'tpoReqAdicInfo'
if @@ROWCOUNT = 0
begin
	insert into tsmDataDictCol values ('tpoReqAdicInfo', 'Type', 'StaticCode', 0,0,0,0,0);
end 

select 1 from tsmDataDictCol where ColumnName = 'TypeBIKey' and TableName = 'tpoReqLine'
if @@ROWCOUNT = 0
begin
	insert into tsmDataDictCol values ('tpoReqLine', 'TypeBIKey', 'StaticCode', 0,0,0,0,0);
end 

select 1 from tsmDataDictCol where ColumnName = 'StateBIKey' and TableName = 'tpoReqLine'
if @@ROWCOUNT = 0
begin
	insert into tsmDataDictCol values ('tpoReqLine', 'StateBIKey', 'StaticCode', 0,0,0,0,0);
end 

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrReqStatus0'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrReqStatus0')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Pendiente')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqAdicInfo' AND tlv.ColumnName = 'statusKey' AND tlv.DBValue = 0
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqAdicInfo','statusKey',0,1,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrReqStatus1'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrReqStatus1')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Autorizada')
END


SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqAdicInfo' AND tlv.ColumnName = 'statusKey' AND tlv.DBValue = 1
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqAdicInfo','statusKey',1,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrReqStatus2'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrReqStatus2')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Aceptada')
END


SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqAdicInfo' AND tlv.ColumnName = 'statusKey' AND tlv.DBValue = 2
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqAdicInfo','statusKey',2,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrReqStatus3'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrReqStatus3')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Cancelada')
END


SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqAdicInfo' AND tlv.ColumnName = 'statusKey' AND tlv.DBValue = 3
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqAdicInfo','statusKey',3,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kpoTBINac'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kpoTBINac')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Nacional')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqLine' AND tlv.ColumnName = 'TypeBIKey' AND tlv.DBValue = 0
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqLine','TypeBIKey',0,1,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kpoTBIInternac'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kpoTBIInternac')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Internacional')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqLine' AND tlv.ColumnName = 'TypeBIKey' AND tlv.DBValue = 1
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqLine','TypeBIKey',1,0,@strNo,0)
END

SELECT @strNo = NULL;
/*------------------------------------------------------------------------------------------------------------------*/
SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kpoStILicit'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kpoStILicit')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Licitación')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqLine' AND tlv.ColumnName = 'StateBIKey' AND tlv.DBValue = 0
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqLine','StateBIKey',0,1,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kpoStIDictTecn'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kpoStIDictTecn')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Dictamen Técnico')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqLine' AND tlv.ColumnName = 'StateBIKey' AND tlv.DBValue = 1
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqLine','StateBIKey',1,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kpoStILegal'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kpoStILegal')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Legal')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqLine' AND tlv.ColumnName = 'StateBIKey' AND tlv.DBValue = 2
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqLine','StateBIKey',2,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kpoStIComContrat'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kpoStIComContrat')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Comité de Contratación')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqLine' AND tlv.ColumnName = 'StateBIKey' AND tlv.DBValue = 3
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqLine','StateBIKey',3,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kpoStIIntelCom'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kpoStIIntelCom')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Inteligencia Comercial')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqLine' AND tlv.ColumnName = 'StateBIKey' AND tlv.DBValue = 4
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqLine','StateBIKey',4,0,@strNo,0)
END

SELECT @strNo = NULL;
/*------------------------------------------------------------------------------------------------------------------*/
SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kPOReqPTMP'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kPOReqPTMP')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'PT/MP/MT')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqAdicInfo' AND tlv.ColumnName = 'Type' AND tlv.DBValue = 1
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqAdicInfo','Type',1,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kPOReqOther'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kPOReqOther')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Otros')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tpoReqAdicInfo' AND tlv.ColumnName = 'Type' AND tlv.DBValue = 2
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tpoReqAdicInfo','Type',2,0,@strNo,0)
END

SELECT @strNo = NULL;

/*--------------------------------------------------------------------------*/

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kSecEventChgStatusReqAutz'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kSecEventChgStatusReqAutz')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Cambiar Estatus de Requisición (Autorizada)')
END

SELECT 1 FROM tsmSecurEvent WHERE SecurEventID = 'CHGPOETREQAUTZ'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEvent VALUES ('CHGPOETREQAUTZ',@strNo,11)
END

SELECT 1 FROM tsmSecurEventPerm  WHERE SecurEventID = 'CHGPOETREQAUTZ' AND UserGroupID = 'SysAdmin'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEventPerm VALUES ('SysAdmin', 'CHGPOETREQAUTZ', 1, NULL, 0)	
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kSecEventChgStatusReqAcpt'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kSecEventChgStatusReqAcpt')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Cambiar Estatus de Requisición (Aceptada)')
END

SELECT 1 FROM tsmSecurEvent WHERE SecurEventID = 'CHGPOETREQACPT'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEvent VALUES ('CHGPOETREQACPT',@strNo,11)
END

SELECT 1 FROM tsmSecurEventPerm  WHERE SecurEventID = 'CHGPOETREQACPT' AND UserGroupID = 'SysAdmin'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEventPerm VALUES ('SysAdmin', 'CHGPOETREQACPT', 1, NULL, 0)	
END


SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kSecEventChgStatusReqRchd'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kSecEventChgStatusReqRchd')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Cambiar Estatus de Requisición (Cancelada)')
END

SELECT 1 FROM tsmSecurEvent WHERE SecurEventID = 'CHGPOETREQRCHZ'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEvent VALUES ('CHGPOETREQRCHZ',@strNo,11)
END

SELECT 1 FROM tsmSecurEventPerm WHERE SecurEventID = 'CHGPOETREQRCHZ' AND UserGroupID = 'SysAdmin'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEventPerm VALUES ('SysAdmin', 'CHGPOETREQRCHZ', 1, NULL, 0)	
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kSecEventRmvRequisition'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kSecEventRmvRequisition')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Eliminar una Requisición')
END

SELECT 1 FROM tsmSecurEvent WHERE SecurEventID = 'DLTPOREQDELETE'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEvent VALUES ('DLTPOREQDELETE',@strNo,11)
END

SELECT 1 FROM tsmSecurEventPerm WHERE SecurEventID = 'DLTPOREQDELETE' AND UserGroupID = 'SysAdmin'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEventPerm VALUES ('SysAdmin', 'DLTPOREQDELETE', 1, NULL, 0)	
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kSecEventGSTReqBuyInfo'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kSecEventGSTReqBuyInfo')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Gestionar Información sobre Compra de Art. en Requisición')
END

SELECT 1 FROM tsmSecurEvent WHERE SecurEventID = 'CHGPOBUYREQINFO'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEvent VALUES ('CHGPOBUYREQINFO',@strNo,11)
END

SELECT 1 FROM tsmSecurEventPerm WHERE SecurEventID = 'CHGPOBUYREQINFO' AND UserGroupID = 'SysAdmin'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEventPerm VALUES ('SysAdmin', 'CHGPOBUYREQINFO', 1, NULL, 0)	
END

SELECT @strNo = NULL;
/*--------------------------------------------------------------------------*/
/****** Object:  View [dbo].[vdvRequisitionTrace]    Script Date: 07/02/2019 14:52:38 ******/
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vdvRequisitionTrace]'))
DROP VIEW [dbo].[vdvRequisitionTrace]
GO
CREATE VIEW [dbo].[vdvRequisitionTrace]
AS
SELECT  dbo.tpoReqTrace.modifyReq AS Modifica, dbo.tpoReqTrace.modifyDate AS Modificada, dbo.tpoReqTrace.dateReq AS [Fecha Req], 
		dbo.tpoReqStatus.statusID AS [Estatus Req], dbo.tpoReqTrace.descriptionStatus AS [Estatus Descripción], dbo.tpoReqTrace.autorizaReq AS Autoriza, 
		dbo.tpoReqTrace.dateAutorReq AS Autorizada, dbo.tpoReqTrace.aceptaReq AS Acepta, dbo.tpoReqTrace.dateAceptReq AS Aceptada, 
		dbo.tpoReqTrace.rejectReq AS Cancela, dbo.tpoReqTrace.dateRejectReq AS Cancelada, dbo.tpoReqTrace.segAutorizaReq AS [Seg Autorizo], 
		dbo.tpoReqTrace.segDateAutorReq AS [Seg Autorización], dbo.tpoReqTrace.ItemID AS Artículo, dbo.tpoReqTrace.QtyReq AS Cantidad, 
		dbo.tpoReqTrace.UnitMeasID AS UoM, dbo.tpoReqTrace.EstimatedPres AS Presupuesto, dbo.tpoReqTrace.ReqLineBuyerID AS Comprador, 
		dbo.tpoReqTrace.ReqLineCmnt AS [Comentario Art], dbo.tpoReqTrace.ReqTraceKey, dbo.tpoReqTrace.ReqLineKey, dbo.tpoReqTrace.ReqKey, 
		dbo.tpoRequisition.TranID AS Transacción, dbo.tpoReqTrace.CompanyID, dbo.tpoReqTrace.Proveedor, dbo.tpoReqTrace.Departamento, 
		dbo.tpoReqTrace.ReqLAlmacen, dbo.tpoReqTrace.ReqLDepartamento, dbo.tpoReqTrace.ReqLDescription AS [Descripción Art], 
		dbo.tpoReqTrace.StateBIID AS [Estado Art], dbo.tpoReqTrace.TypeBIID AS [Tipo de Compra]
FROM    dbo.tpoReqStatus INNER JOIN
		dbo.tpoReqTrace ON dbo.tpoReqStatus.statusKey = dbo.tpoReqTrace.statusKey INNER JOIN
		dbo.tpoRequisition ON dbo.tpoReqTrace.ReqKey = dbo.tpoRequisition.ReqKey
GO
/****** Object:  View [dbo].[vPOReqNo]    Script Date: 11/20/2018 17:16:16 ******/
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vPOReqNo]'))
DROP VIEW [dbo].[vPOReqNo]
GO

/****** Object:  View [dbo].[vPOReqNo]    Script Date: 11/20/2018 17:16:16 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[vPOReqNo]
AS
SELECT     dbo.tpoRequisition.TranNo, dbo.tpoRequisition.TranDate, dbo.tpoReqStatus.statusID, dbo.tpoRequisition.CompanyID, dbo.tpoReqAdicInfo.statusKey, 
                      dbo.tpoRequisition.Originator, dbo.tpoReqAdicInfo.segLvlAutoriza, dbo.tpoReqAdicInfo.segAutorizaReq, dbo.tpoRequisition.Status AS Situación
FROM         dbo.tpoRequisition INNER JOIN
                      dbo.tpoReqAdicInfo ON dbo.tpoRequisition.ReqKey = dbo.tpoReqAdicInfo.ReqKeyIA  INNER JOIN
                      dbo.tpoReqStatus ON dbo.tpoRequisition.CompanyID = dbo.tpoReqStatus.CompanyID AND dbo.tpoReqAdicInfo.statusKey = dbo.tpoReqStatus.statusKey

GO

DELETE tsmDataDictViewCol WHERE  SQLViewName = 'vPOReqNo'
INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'TranNo', 'tpoRequisition','TranNo',NULL)
INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'TranDate', 'tpoRequisition','TranDate',NULL)
--INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'statusID', 'tpoReqStatus','statusID',NULL)
INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'CompanyID', 'tpoRequisition','CompanyID',NULL)
INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'statusKey', 'tpoReqAdicInfo','statusKey',NULL)
INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'Originator', 'tpoRequisition','Originator',NULL)
--INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'segLvlAutoriza', 'tpoReqAdicInfo','segLvlAutoriza',NULL)
--INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'segAutorizaReq', 'tpoReqAdicInfo','segAutorizaReq',NULL)
INSERT tsmDataDictViewCol (SQLViewName, SQLViewColName, TableName, ColumnName, SQLViewColCaption) VALUES ('vPOReqNo', 'Situación', 'tpoReqAdicInfo','statusKey','Situación')

GO

DECLARE @LookupKey INT, @LookupViewKey INT 
 EXEC spsmSaveLookupViewStd @LookupKey OUTPUT, @LookupViewKey OUTPUT, 'vPOReqNo', 1, 'POReqNo', 11, 'TranNo, CompanyID', 'vPOReqNo', 1, NULL,'Standard', 0, 0, 'TranNo, TranDate, Originator, statusID, segLvlAutoriza, segAutorizaReq, Situación','TranNo', 10
 SELECT @LookupKey AS LookupKey, @LookupViewKey as LookupViewKey 
 GO 


