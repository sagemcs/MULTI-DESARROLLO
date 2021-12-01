Attribute VB_Name = "basRequistn"
Option Explicit
'**********************************************************************
'     Name: baspurchord
'     Desc: Main Module of this Object containing "Sub Main ()".
' Original: 05-07-1997
'     Mods: mm/dd/yy XXX
'**********************************************************************


    'Public Const ktskPrint = 9998
'Public moAPSTaxCalc             As New clsAPSTaxCalc
Public mfrmMain                 As Form
Public mlFreightSTaxClassKey    As Long
Public mbEnterAsTab             As Boolean              ' use enter as tab?
Public msUserFld(3)             As String
Public msUserFld_Line(1)        As String               'PO line level comment fields.
Public mbDontRunGridClick       As Boolean


Public Const ktskQuickPrintReqs = 184942943                      'PO Print Reqs DLL
Public Const kclsQuickPrintReqs = "poztbdl1.clsPrintReqs"

Public Const kCustomFields = "Custom Fields..."

' tapVendor | Status
    Public Const kvActiveVend       As Integer = 1
    Public Const kvInactiveVend     As Integer = 2
    Public Const kvTemporaryVend    As Integer = 3
    Public Const kvDeletedvend      As Integer = 4

'-- Static List Constants
' timItem | Status
    Public Const kvActiveItem       As Integer = 1
    Public Const kvInactiveItem     As Integer = 2
    Public Const kvDiscontinuedItem As Integer = 3
    Public Const kvDeletedItem      As Integer = 4
    
    ' timItem | ItemType
    Public Const kvFinishedGoods    As Integer = 5
    Public Const kvRawMaterial      As Integer = 6
    Public Const kvCommentOnly      As Integer = 4
    Public Const kvExpense          As Integer = 3
    Public Const kvService          As Integer = 2
    Public Const kvMiscItem         As Integer = 1
    Public Const kvKit              As Integer = 7


    Public mlUniqueKey      As Long
    Public mbCalcTaxes      As Boolean
    Public moUF             As Object		'User Comment Fields Object
    
    Public Type VendDflts
        lBuyerKey           As Long
        dCreditLimit        As Double
        lVendKey            As Long
        lPurchFromAddrKey   As Long
        sPurchFromAddrID    As String
        sPurchFromAddrName  As String
        lDfltItemKey        As Long
        lDfltPurchAcctKey   As Long
        lRemitToAddrKey     As Long
        sRemitToAddrID      As String
        sRemitToAddrName    As String
        iHold               As Integer
        lPrimaryAddrKey     As Long
        dRetntRate          As Double
        iStatus             As Integer
        lPmtTermsKey        As Long
        lPOFormKey          As Long
        lCurrExchSchdKey    As Long
        sCurrID             As String
        lDfltCntctKey       As Long
        lShipMethKey        As Long
        lShipZoneKey        As Long
        lSTaxSchdKey        As Long
        sFOB                As String
        lVendClassKey       As Long
        sV1099Box           As String
        sV1099Control       As String
        iV1099Form          As Integer
        iV1099Type          As Integer
        lClassOvrdSegKey    As Long
        sClassOvrdSegVal    As String
        MatchKey            As Long
    End Type
    
    Public Type ItemDflts
        iStatus             As Integer
        sTargetCompID       As String
        lPOLineKey          As Long
        iPOLineNo           As Integer
        lItemKey            As Long
        lGLAcctKey          As Long
        dQtyRcvd            As Double
        dQtyInvcd           As Double
        dQtyReturned        As Double
        dAmtInvcd           As Double
        dOrigQtyOrd         As Double
        sOrigPromiseDate    As String
        iDropShip           As Integer
        lShipToCustKey      As Long
        sCustid             As String
        lShipToCustAddrKey  As Long
        sCustAddrID         As String
        lShipToWhseKey      As Long
        sWhseID             As String
        sComment            As String
        lToleranceKey       As Long
        sToleranceID        As String
        iRcptMatched        As Integer
        sRcptMatchDate      As String
        iReqRcvr            As Integer
        iOvrdExmpt          As Integer
        lSTaxClassKey       As Long
        sSTaxClassID        As String
        lSTaxSchdKey        As Long
        sSTaxSchdID         As String
        lAddrKey            As Long
        sAddrName           As String
        sAddr(5)            As String
        sCity               As String
        sSTate              As String
        sZIP                As String
        sCountry            As String
        lUpdateCounter      As Long
        lShipMethKey        As Long
        sShipMethID         As String
        lShipZoneKey        As Long
        sShipZoneID         As String
        iAllowCostOvrd      As Integer
        iCommentOnly        As Integer
        iPOLineDetlEntryNo  As Integer
        sUserFld1           As String
        sUserFld2           As String
    End Type
    
 
    

    Public msCurrSymbol         As String
    Public msHomeCurrSymbol     As String
    Public miDigAfterDec        As Integer
    Public miHomeDigAfterDec    As Integer
    Public miRoundPrec          As Integer
    Public miHomeRoundPrec      As Integer
    Public miRoundMeth          As Integer
    Public miHomeRoundMeth      As Integer
    Public mlCurrencyLocale     As Long
    Public mlHomeCurrencyLocale As Long

Const VBRIG_MODULE_ID_STRING = "REQUISTN.BAS"

Public Sub Main()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If App.StartMode = vbSModeStandalone Then
        Dim oApp As Object
        Set oApp = CreateObject("pozde001.clsrequistn")
        StartAppInStandaloneMode oApp, Command$()
    End If

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Main", VBRIG_IS_MODULE
        Err.Raise guSotaErr.Number
'+++ VB/Rig End +++
End Sub

Private Function sMyName() As String
'+++ VB/Rig Skip +++
    sMyName = "basRequistn"
End Function

Public Function lGetNextSurrogateKey(oDB As Object, tblName As String) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lKeyVal As Long
    
    lKeyVal = 0
    With oDB
        .SetInParam tblName
        .SetOutParam lKeyVal
        .ExecuteSP ("spGetNextSurrogateKey")
        lKeyVal = .GetOutParam(2)
        .ReleaseParams
    End With
    
    lGetNextSurrogateKey = lKeyVal
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lGetNextSurrogateKey", VBRIG_IS_MODULE
        Err.Raise guSotaErr.Number
'+++ VB/Rig End +++
End Function

Public Sub MyErrMsg(oclass As Object, sDesc As String, lErr As Long, sSub As String, sProc As String)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
Dim sText As String

    If lErr = guSotaErr.Number And Trim(guSotaErr.Description) <> Trim(sDesc) Then
        sDesc = sDesc & " " & guSotaErr.Description
    End If

    If oclass Is Nothing Then
        sText = sText & " AppName:  " & App.Title
        sText = sText & "    Error < " & Format(lErr)
        sText = sText & " >   occurred at < " & sSub
        sText = sText & " >  in Procedure < " & sProc
        sText = sText & " > " & sDesc
    Else
        sText = "Module: " & oclass.moSysSession.MenuModule
        sText = sText & "  Company: " & oclass.moSysSession.CompanyID
        sText = sText & "  AppName:  " & App.Title
        sText = sText & "    Error < " & Format(lErr)
        sText = sText & " >   occurred at < " & sSub
        sText = sText & " >  in Procedure < " & sProc
        sText = sText & " > " & sDesc
    End If
    
    MsgBox sText, vbExclamation, "Sage 500 ERP"

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "MyErrMsg", VBRIG_IS_MODULE
        Err.Raise guSotaErr.Number
'+++ VB/Rig End +++
End Sub




