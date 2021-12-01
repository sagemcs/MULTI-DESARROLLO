DECLARE @LookupKey INT, @LookupViewKey INT 
 EXEC spsmSaveLookupViewStd @LookupKey OUTPUT, @LookupViewKey OUTPUT, 'vluContract', 1, 'Contract', 61, 'ContractKey, ContractID, ContractNo', 'vluContract', 1, NULL,'Standard', 0, 0, 'ContractID, ContractNo, CurrID, ContractAmt, VendName, TypeAsString','', 10
 SELECT @LookupKey AS LookupKey, @LookupViewKey as LookupViewKey 
 GO 


