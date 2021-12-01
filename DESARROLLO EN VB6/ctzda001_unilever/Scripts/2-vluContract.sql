/****** Object:  View [dbo].[vluContract]    Script Date: 06/20/2019 10:09:23 ******/
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[vluContract]'))
DROP VIEW [dbo].[vluContract]
GO


/****** Object:  View [dbo].[vluContract]    Script Date: 06/20/2019 10:09:23 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE VIEW [dbo].[vluContract] as SELECT p.ContractID, p.ContractNo,p.ContractKey,p.CurrID, p.ContractAmt, p.CompanyID, p.VendorKey, s.VendName, p.[Type], t.LocalText "TypeAsString"
                                    FROM tctContract AS p JOIN tapVendor AS s ON p.VendorKey = s.VendKey JOIN vListValidationString AS t ON t.TableName = 'tctContract' AND t.ColumnName = 'Type' AND t.DBValue = p.[Type]


GO


