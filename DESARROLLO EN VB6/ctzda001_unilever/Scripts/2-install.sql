DECLARE @strNo AS INT

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrCTContractTypeCont'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrCTContractTypeCont')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Contrato')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tctContract' AND tlv.ColumnName = 'Type' AND tlv.DBValue = 1
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tctContract','Type',1,1,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrCTContractTypeSuple'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrCTContractTypeSuple')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Suplemento')
END


SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tctContract' AND tlv.ColumnName = 'Type' AND tlv.DBValue = 2
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tctContract','Type',2,0,@strNo,0)
END

SELECT @strNo = NULL;
------------------------------------------------

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrCTContractClasPurch'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrCTContractClasPurch')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Artículos')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tctContract' AND tlv.ColumnName = 'Clasification' AND tlv.DBValue = 1
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tctContract','Clasification',1,1,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrCTContractClasServ'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrCTContractClasServ')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Servicio')
END


SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tctContract' AND tlv.ColumnName = 'Clasification' AND tlv.DBValue = 2
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tctContract','Clasification',2,0,@strNo,0)
END

SELECT @strNo = NULL;
-------------------------------------------------
SELECT @strNo = StringNo FROM tsmString WHERE ConstantName = 'kmsgCTDesaactCont'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString ORDER BY StringNo asc
	INSERT INTO tsmString (StringNo,ConstantName)VALUES(@strNo,'kmsgCTDesaactCont');
	INSERT INTO tsmLocalString (StringNo,LanguageID,LocalText)VALUES(@strNo,1033,'Desactivar Contrato/Suplemento')
END

SELECT 1 FROM tsmSecurEvent WHERE SecurEventID = 'CTDesactCont'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEvent(SecurEventID,DescStrNo,ModuleNo)VALUES('CTDesactCont',@strNo,61)
END

SELECT @strNo = NULL;


SELECT @strNo = StringNo FROM tsmString WHERE ConstantName = 'kmsgCTReactCont'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString ORDER BY StringNo asc
	INSERT INTO tsmString (StringNo,ConstantName)VALUES(@strNo,'kmsgCTReactCont');
	INSERT INTO tsmLocalString (StringNo,LanguageID,LocalText)VALUES(@strNo,1033,'Reactivar Contrato/Suplemento')
END

SELECT 1 FROM tsmSecurEvent WHERE SecurEventID = 'CTReactCont'
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmSecurEvent(SecurEventID,DescStrNo,ModuleNo)VALUES('CTReactCont',@strNo,61)
END

SELECT @strNo = NULL;

------------------------------------------------

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrCTContLineTypeAdd'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrCTContLineTypeAdd')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Adicionar')
END

SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tctContractLine' AND tlv.ColumnName = 'Type' AND tlv.DBValue = 1
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tctContractLine','Type',1,1,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrCTContLineTypeDel'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrCTContLineTypeDel')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Eliminar')
END


SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tctContractLine' AND tlv.ColumnName = 'Type' AND tlv.DBValue = 2
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tctContractLine','Type',2,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrCTContLineTypeEditUp'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrCTContLineTypeEditUp')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Modificar Aumento')
END


SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tctContractLine' AND tlv.ColumnName = 'Type' AND tlv.DBValue = 3
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tctContractLine','Type',3,0,@strNo,0)
END

SELECT @strNo = NULL;

SELECT @strNo = ts.StringNo FROM tsmString AS ts WHERE ts.ConstantName = 'kstrCTContLineTypeEditDown'
IF @@ROWCOUNT = 0
BEGIN
	SELECT @strNo = StringNo + 1 FROM tsmString AS ts ORDER BY ts.StringNo
	INSERT INTO tsmString(StringNo,ConstantName)VALUES(@strNo,'kstrCTContLineTypeEditDown')
	INSERT INTO tsmLocalString (StringNo,LanguageID, LocalText)VALUES(@strNo,1033,'Modificar Decremento')
END


SELECT 1 FROM tsmListValidation AS tlv WHERE tlv.TableName = 'tctContractLine' AND tlv.ColumnName = 'Type' AND tlv.DBValue = 4
IF @@ROWCOUNT = 0
BEGIN
	INSERT INTO tsmListValidation(TableName,ColumnName,	DBValue,IsDefault,StringNo,IsHidden)VALUES('tctContractLine','Type',4,0,@strNo,0)
END

SELECT @strNo = NULL;