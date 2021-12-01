Attribute VB_Name = "basCTZDA001"
Option Explicit

Const VBRIG_MODULE_ID_STRING = "CTZDA001.BAS"

    
'    Public Const kColCompanyID = 1
'    Public Const kColContractID = 1
    Public Const kColContractKey = 1
    Public Const kColContractLineKey = 2
    Public Const kColSeqNo = 3
    Public Const kColItemKey = 4
    Public Const kColItemID = 5
    Public Const kColDescription = 6
    Public Const kColUnitCost = 7
    Public Const kColUnitMeasID = 8
    Public Const kColUnitMeasKey = 9
    Public Const kColItemQty = 10
    Public Const kColLineAmt = 11
    Public Const kColMaxLot = 12
    Public Const kColMinLot = 13
    Public Const kColRoundValue = 14
    Public Const kColDeliveryTime = 15
    Public Const kColCreateDate = 16
    Public Const kColCreateUser = 17
    Public Const kColUpdateDate = 18
    Public Const kColUpdateUser = 19
    Public Const kColType = 20
    Public Const kColQtyVariation = 21
    
    
    Public Const kMaxCol = 21
    
    Public Const kColOpertInvLineKey = 1
    Public Const kColOpertPoNo = 2
    Public Const kColOpertInvNo = 4
    Public Const kColOpertQty = 3
    Public Const kColOpertAmt = 6
    Public Const kColOpertDesc = 5
    Public Const kColOpertDate = 7
    
    
    Public Const kSecurityEventDesactContr = "CTDesactCont"
    Public Const kSecurityEventReactContr = "CTReactCont"
    
    
    
    Public Const kContractTypeContract = 1
    Public Const kContractTypeSuplement = 3
    
    Public Const kLineTypeAdd = 1
    Public Const kLineTypeDel = 2
    Public Const kLineTypeModUp = 3
    Public Const kLineTypeModDown = 4
    Public Const kLineTypeModCos = 5
'-- Needed for the currency form
Public mbEnterAsTab As Boolean

Public Sub Main()
    If App.StartMode = vbSModeStandalone Then
        Dim oApp As Object
        Set oApp = CreateObject("CTZDA001.clsCTZDA001")
        StartAppInStandaloneMode oApp, Command$()    ' ** The third parameter MUST be the
                                                     '    TaskID for this project in order for the
                                                     '    Standalone method to be activated for ActiveX Exe's
    End If
End Sub

Private Function sMyName() As String
'+++ VB/Rig Skip +++
    sMyName = "basCTZDA001"
End Function
