/****** Object:  StoredProcedure [dbo].[spctIsValidContract]    Script Date: 06/20/2019 10:03:06 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spctIsValidContract]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spctIsValidContract]
GO

/****** Object:  StoredProcedure [dbo].[spctIsValidContract]    Script Date: 06/20/2019 10:03:06 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spctIsValidContract] (@iCompanyID as Varchar(3), @iContractID as VARCHAR(15), @oContractKey as INT OUTPUT, @oRetVal as INT OUTPUT) AS
BEGIN
	SELECT @oRetVal = -1;
	DECLARE @tContractKey AS INT  
	select @tContractKey = ContractKey FROM tctContractLog WHERE ContractID = @iContractID AND CompanyID = @iCompanyID;
	IF @tContractKey IS NULL
	BEGIN
		EXEC spGetNextSurrogateKey
			'tctContract',
			@tContractKey OUTPUT
	END
	
	SELECT @oContractKey = @tContractKey
	
	SELECT @oRetVal = 0
END
GO



GRANT EXECUTE ON [dbo].[spctIsValidContract] TO ApplicationDBRole