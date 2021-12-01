/****** Object:  StoredProcedure [dbo].[spctUpdateContractLog]    Script Date: 06/20/2019 10:06:56 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[spctUpdateContractLog]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[spctUpdateContractLog]
GO


/****** Object:  StoredProcedure [dbo].[spctUpdateContractLog]    Script Date: 06/20/2019 10:06:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spctUpdateContractLog] (@iContractKey AS INT) AS 
BEGIN
	SELECT 1 FROM tctContractLog WHERE ContractKey = @iContractKey
	IF @@ROWCOUNT = 0
	BEGIN
		INSERT INTO tctContractLog 
		SELECT p.ContractKey, p.ContractID, p.ContractNo, p.CompanyID, p.ContractAmt, p.CurrID, p.SeqNo, p.[State]
		  FROM tctContract AS p WHERE p.ContractKey = @iContractKey;
	END
	ELSE
	BEGIN
		UPDATE s
		SET
			s.ContractAmt = p.ContractAmt,
			s.CurrID = p.CurrID,
			s.SeqNo = p.SeqNo,
			s.[State] = p.[State]
		FROM tctContractLog AS s JOIN tctContract AS p ON p.ContractKey = s.ContractKey WHERE p.ContractKey = @iContractKey
	END
END
GO


GRANT EXECUTE ON [dbo].[spctUpdateContractLog] TO ApplicationDBRole