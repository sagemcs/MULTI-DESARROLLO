/****** Object:  StoredProcedure [dbo].[spctGetNextContractNo]    Script Date: 06/20/2019 10:02:10 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spctGetNextContractNo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spctGetNextContractNo]
GO

/****** Object:  StoredProcedure [dbo].[spctGetNextContractNo]    Script Date: 06/20/2019 10:02:10 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spctGetNextContractNo](@iCompanyID AS VARCHAR(3), @oContractNo AS VARCHAR(15) OUTPUT, @oRetVal AS INT OUTPUT)AS
BEGIN
	DECLARE @tContractNo AS INT
	SELECT @oContractNo = '0000000000';
	SELECT @oRetVal = -1;
	
	SELECT @tContractNo = p.NextTranNo FROM tciTranTypeCompany AS p WHERE p.CompanyID =@iCompanyID AND p.TranType = 6101
	IF @tContractNo IS NULL
	BEGIN
		SELECT @oRetVal = -2;
		RETURN	
	END
	
	SELECT @tContractNo
	SELECT @oContractNo = CONVERT(VARCHAR(15), @tContractNo)
	SELECT 1 FROM tctContractLog AS p WHERE p.ContractID = @oContractNo
	IF @@ROWCOUNT > 0
	BEGIN
		SELECT @tContractNo = p.ContractID FROM tctContractLog AS p ORDER BY p.ContractID ASC
		UPDATE tciTranTypeCompany SET NextTranNo = @tContractNo + 1 WHERE CompanyID =@iCompanyID AND TranType = 6101
		SELECT @oRetVal = -3
		RETURN
	END
	
	UPDATE tciTranTypeCompany SET NextTranNo = @tContractNo + 1 WHERE CompanyID =@iCompanyID AND TranType = 6101
	SELECT @oRetVal = 0;
END
GO


GRANT EXECUTE ON [dbo].[spctGetNextContractNo] TO ApplicationDBRole