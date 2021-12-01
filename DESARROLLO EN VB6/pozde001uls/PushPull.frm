VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F2F2EE3C-0D23-4FC8-944C-7730C86412E3}#67.0#0"; "sotasbar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0FA91D91-3062-44DB-B896-91406D28F92A}#65.0#0"; "SOTACalendar.ocx"
Object = "{C41A85E3-4CB6-40B5-B425-EE9ECC5E6F06}#181.0#0"; "SOTATbar.ocx"
Begin VB.Form frmSelectReqLines 
   Caption         =   "Select Requisition Lines"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   HelpContextID   =   60295
   Icon            =   "PushPull.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6810
   Begin VB.CheckBox chkCostFromReq 
      Caption         =   "Default &Unit Cost from Requisition"
      Enabled         =   0   'False
      Height          =   240
      Left            =   2685
      TabIndex        =   2
      Top             =   600
      Value           =   1  'Checked
      WhatsThisHelpID =   60316
      Width           =   3750
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "&Clear All"
      Height          =   375
      Left            =   1350
      TabIndex        =   5
      Top             =   4620
      WhatsThisHelpID =   60315
      Width           =   1245
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4620
      WhatsThisHelpID =   60314
      Width           =   1125
   End
   Begin FPSpreadADO.fpSpread grdSelectPush 
      Height          =   3645
      Left            =   0
      TabIndex        =   3
      Top             =   900
      WhatsThisHelpID =   60313
      Width           =   6675
      _Version        =   524288
      _ExtentX        =   11774
      _ExtentY        =   6429
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "PushPull.frx":3332
      AppearanceStyle =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3750
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SOTACalendarControl.SOTACalendar CustomDate 
      Height          =   315
      Index           =   0
      Left            =   -30000
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      WhatsThisHelpID =   75
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BackColor       =   -2147483633
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskedText      =   "  /  /    "
      Protected       =   -1  'True
      Text            =   "  /  /    "
   End
   Begin VB.ComboBox CustomCombo 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   -30000
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      WhatsThisHelpID =   64
      Width           =   1245
   End
   Begin VB.OptionButton CustomOption 
      Caption         =   "Option"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      WhatsThisHelpID =   68
      Width           =   1245
   End
   Begin VB.CheckBox CustomCheck 
      Caption         =   "Check"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      WhatsThisHelpID =   62
      Width           =   1245
   End
   Begin VB.CommandButton CustomButton 
      Caption         =   "Button"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   -30000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2730
      Visible         =   0   'False
      WhatsThisHelpID =   61
      Width           =   1245
   End
   Begin VB.Frame CustomFrame 
      Caption         =   "Frame"
      Enabled         =   0   'False
      Height          =   1035
      Index           =   0
      Left            =   -30000
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComCtl2.UpDown CustomSpin 
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      WhatsThisHelpID =   69
      Width           =   195
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
   End
   Begin NEWSOTALib.SOTACurrency CustomCurrency 
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1140
      Visible         =   0   'False
      WhatsThisHelpID =   65
      Width           =   1245
      _Version        =   65536
      _ExtentX        =   2196
      _ExtentY        =   503
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
      text            =   "           0.00"
      sDecimalPlaces  =   2
   End
   Begin NEWSOTALib.SOTANumber CustomNumber 
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      WhatsThisHelpID =   67
      Width           =   1245
      _Version        =   65536
      _ExtentX        =   2196
      _ExtentY        =   503
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
      text            =   "           0.00"
      sDecimalPlaces  =   2
   End
   Begin NEWSOTALib.SOTAMaskedEdit CustomMask 
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      WhatsThisHelpID =   66
      Width           =   1245
      _Version        =   65536
      _ExtentX        =   2196
      _ExtentY        =   503
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin NEWSOTALib.SOTACustomizer picDrag 
      Height          =   330
      Index           =   0
      Left            =   -75000
      TabIndex        =   17
      Top             =   645
      Visible         =   0   'False
      WhatsThisHelpID =   70
      Width           =   345
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   582
      _StockProps     =   0
   End
   Begin SOTAToolbarControl.SOTAToolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      WhatsThisHelpID =   74
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   741
      Style           =   4
   End
   Begin VB.TextBox txtReqNo 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   540
      WhatsThisHelpID =   60299
      Width           =   1185
   End
   Begin StatusBar.SOTAStatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      Top             =   4725
      WhatsThisHelpID =   60298
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   688
      BrowseVisible   =   0   'False
      UserNameVisible =   0   'False
      CompanyIDVisible=   0   'False
      BusinessDateVisible=   0   'False
      StatusVisible   =   0   'False
   End
   Begin VB.Label CustomLabel 
      Caption         =   "Label"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   16
      Top             =   60
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblReqNo 
      Caption         =   "Requisition "
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   570
      Width           =   945
   End
End
Attribute VB_Name = "frmSelectReqLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If CUSTOMIZER Then
    Public moFormCust As Object
#End If

'Private WithEvents moSelectGrid As fpSpread
Private WithEvents moSelectGM                As clsGridMgr               ' grid manager class
Attribute moSelectGM.VB_VarHelpID = -1


Private moClass As Object
Private mlSession As Long

Private WithEvents moPushPull As clsPushPull
Attribute moPushPull.VB_VarHelpID = -1

'Binding Object Variables
Public WithEvents moDmSelectGrid        As clsDmGrid   ' grid object
Attribute moDmSelectGrid.VB_VarHelpID = -1
Public moContextMenu       As clsContextMenu           ' context menu

Public moMapSrch           As New Collection    ' collection for f4/f5 navigators/lookups


Private msReqID             As String
Private mlReqKey            As Long
Private mbDontChkclick      As Boolean                  'global don't run click logic
Private mbScrolling         As Boolean
Private miOldFormHeight As Long
Private miOldFormWidth As Long
Private miMinFormHeight As Long
Private miMinFormWidth As Long



Const kMaxCols = 9
Const kColSelect = 1
Const kColDocLineKey = 2
Const kcolItemID = 3
Const kColDescription = 4
Const kColQty = 5
Const kColReqDate = 6
Const kColVendorID = 7
Const kColDeptID = 8
Const kcolWhseID = 9

' Requisition status constants
Private Const kvReqIncomplete As Integer = 0
Private Const kvReqPendApprvl As Integer = 1
Private Const kvReqOpen As Integer = 2
Private Const kvReqInactive As Integer = 3
Private Const kvReqCanceled As Integer = 4
Private Const kvReqClosed As Integer = 5

'Requisition Line Status Constants
Private Const kvStatusOpen As Integer = 1
Private Const kvStatusClosed As Integer = 2
    
'Collection object for update counter
Private mcolUpdCounter As New Collection

Const kMaxRows = 0

Private lContractKeyUsed As Long 'Agregado por Multiconsulting

'************************************************************************
'   Description:
'       bind grid manager (handles right mouse clicks for grids)
'
'   Param:
'       <none>
'
'   Returns:
'
'************************************************************************

Const VBRIG_MODULE_ID_STRING = "PUSHPULL.FRM"



Private Sub BindGM()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Set moSelectGM = New clsGridMgr

    With moSelectGM
        Set .Grid = grdSelectPush
        Set .Form = frmSelectReqLines
        Set .DM = moDmSelectGrid
' Set the grid type to data sheet so that the context menus work properly.
        .GridType = kGridDataSheetNoAppend
        .GridSortEnabled = False
        .Init
    End With

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "BindGM", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Function sMyName() As String
'+++ VB/Rig Skip +++
    sMyName = Me.Name
End Function

Public Sub Init(oclass As Object, sReqID As String, lReqKey As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Set moClass = oclass
    msReqID = sReqID
    mlReqKey = lReqKey
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Init", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub


Private Sub cmdClearAll_Click()
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If Not moFormCust.onClick(cmdClearAll, True) Then Exit Sub
    End If
    #End If
'+++ End Customizer Code Push +++
    Dim lCtr As Long

    For lCtr = 1 To grdSelectPush.MaxRows
        grdSelectPush.Col = kColSelect
        grdSelectPush.Row = lCtr
        grdSelectPush.Value = 0
    Next

End Sub

Private Sub cmdSelectAll_Click()
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If Not moFormCust.onClick(cmdSelectAll, True) Then Exit Sub
    End If
    #End If
'+++ End Customizer Code Push +++
    Dim lCtr As Long

    For lCtr = 1 To grdSelectPush.MaxRows
        grdSelectPush.Col = kColSelect
        grdSelectPush.Row = lCtr
        grdSelectPush.Value = 1
    Next
'    AfterSelectAll grdSelectPush.MaxRows
End Sub

Private Sub Form_Activate()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
#If CUSTOMIZER Then
    If moFormCust Is Nothing Then
        Set moFormCust = CreateObject("SOTAFormCustRT.clsFormCustRT")
        If Not moFormCust Is Nothing Then
                moFormCust.Initialize Me, goClass
                Set moFormCust.CustToolbarMgr = Nothing
                moFormCust.ApplyFormCust
        End If
    End If
#End If
    txtReqNo = msReqID
    moPushPull.Init
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Form_Activate", VBRIG_IS_FORM
        Select Case VBRIG_IS_FORM_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

'************************************************************************
'   Description:
'       process keydown events on the form.  Used to trap function key
'       presses. NOTE: key preview of the form should be set to True
'
'   Param:
'       KeyCode -   key pressed
'       Shift -     state of the shift/ctl/alt keys
'
'   Returns:
'
'************************************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++

    On Error Resume Next

    Select Case KeyCode

        Case vbKeyF1 To vbKeyF16
            gProcessFKeys Me, KeyCode, Shift

    End Select

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub

'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Form_KeyDown", VBRIG_IS_FORM
        Select Case VBRIG_IS_FORM_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub Form_Load()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    'Set form initial height and Width
    miOldFormHeight = Me.Height
    miOldFormWidth = Me.Width
    miMinFormHeight = miOldFormHeight
    miMinFormWidth = miOldFormWidth
    
    Set moPushPull = New clsPushPull
    moPushPull.LoadImmediate = True
    moPushPull.InitControl moClass
    mlSession = moPushPull.lGetSessionID
    frmRequistn.SetSession = mlSession
   
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Form_Load", VBRIG_IS_FORM
        Select Case VBRIG_IS_FORM_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If Not moFormCust.CanShutdown Then
            Cancel = True
            Exit Sub
        End If
    End If
#End If

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Form_QueryUnload", VBRIG_IS_FORM
        Select Case VBRIG_IS_FORM_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Public Property Get oclass() As Object
'+++ VB/Rig Skip +++
    Set oclass = moClass
End Property
'
'
'Public Function lStartErrorReport(Optional iFileType As Variant, Optional sFileName As Variant) As Long
'
''+++ VB/Rig Begin Push +++
''+++ VB/Rig End +++
'Dim sWhereClause As String
'Dim sTablesUsed As String
'Dim sSelect As String
'Dim sInsert As String
'Dim lRetval As Long
'Dim bValid As Boolean
'Dim iNumTablesUsed As Integer
'Dim RptFileName As String
'Dim lBadRow As Long
'Dim lErrBatchKey As Long
'Dim sSql As String
'Dim rs As Object
''Dim ReportObj As clsReportEngine
''Dim DBObj As Object
''Dim oDDData As clsDDData
''Dim sRealTableCollection As Collection
'
'    On Error GoTo badexit
'
'    lStartErrorReport = kFailure
'
''    ShowStatusBusy frm
'
'    'Set SelectObj = frm.moSelect
'    Set frmRequistn.ReportObj = New clsReportEngine
'    Set frmRequistn.DBObj = moClass.moAppDB
''    lErrBatchKey = moClass.lKey
'
'    frmRequistn.ReportObj.UI = False
'    frmRequistn.ReportObj.AppOrSysDB = kAppDB
'
'    RptFileName = "pozde001.rpt"
'    Set frmRequistn.sRealTableCollection = New Collection
'    With frmRequistn.sRealTableCollection
'        .Add "tciErrorLog" 'the "Driving Table" name
'    End With
'    Set frmRequistn.oDDData = New clsDDData
'    If Not frmRequistn.oDDData.lInitDDData(frmRequistn.sRealTableCollection, moClass.moAppDB, moClass.moAppDB, moClass.moSysSession.CompanyID) = kSuccess Then
''+++ VB/Rig Begin Pop +++
''+++ VB/Rig End +++
'        Exit Function
'    End If
'
'     If (frmRequistn.ReportObj.lInitReport("PO", RptFileName, frmSelectReqLines, frmRequistn.oDDData) = kFailure) Then
''+++ VB/Rig Begin Pop +++
''+++ VB/Rig End +++
'        Exit Function
'    End If
'
'
'
'    On Error GoTo badexit
'
'    '*************** NOTE ********************
'    'THE ORDER OF THE FOLLOWING EVENTS IS IMPORTANT!
'
'    'CUSTOMIZE:  The .RPT file to be used should be set here.  (More than one .RPT file
'    'may exist for the task.)
'    frmRequistn.ReportObj.ReportFileName() = RptFileName
'
'    'Start Crystal print engine, open a print job, and get localized strings from
'    'tsmLocalString table.
'    If (frmRequistn.ReportObj.lSetupReport = kFailure) Then
'        GoTo badexit
'    End If
'
'    'work around if you print without previewing first.
'    'Crystal does not provide a way of getting page orientation
'    'used to create report. use VB constants:
'    'vbPRORPortrait, vbPRORLandscape
'    frmRequistn.ReportObj.Orientation() = vbPRORPortrait
'
'    'CUSTOMIZE:  Set report titles to localized text from tsmLocalString table using call
'    'to gsBuildString with a VB constant defined in StrConst.bas. The subtitles should
'    'not include the format selected by the user, i.e., "Detail" or "Summary".
'    frmRequistn.ReportObj.ReportTitle1() = "Error Log Listing" 'gsBuildString(kVendClassListing, frm.oClass.moAppDB, frm.oClass.moSysSession)
'    frmRequistn.ReportObj.ReportTitle2() = ""
'
'    'CUSTOMIZE:  Include these calls if you have named subtotal & header labels on the report
'    'using the "lbl" convention on formula field names so that label text will handled automatically
'    frmRequistn.ReportObj.UseSubTotalCaptions() = 1
'    frmRequistn.ReportObj.UseHeaderCaptions() = 1
'    'Supress the summary section
'    frmRequistn.ReportObj.lSetSummarySection 0, Nothing
'
'    'set standard formulas, business date, run time, company name etc.
'    'as defined in the template
'    If (frmRequistn.ReportObj.lSetStandardFormulas(frmSelectReqLines) = kFailure) Then
'        GoTo badexit
'    End If
'
'    'Set sort order in .RPT file according to user selections in the Sort grid.
'    'If (ReportObj.lSetSortCriteria(frm.moSort) = kFailure) Then
'        'GoTo badexit
'    'End If
'
'    '********* the following is specific to your report *************'
'    'Select Case RptFileName
'        'Case "XXZYY001.RPT"
'        'Case Else
'    'End Select
'    '********* End of special processing *************'
'
'    'Retrieve the SQL statement stored with the .RPT file and modify it as needed.
'    frmRequistn.ReportObj.BuildSQL
'    frmRequistn.ReportObj.SetSQL
'
'    'CUSTOMIZE:  Include this call if you have named column labels on the report
'    'using the "lbl" convention and wish label text to be handled automatically for you.
'    frmRequistn.ReportObj.SetReportCaptions
'
'    'used in the Summary section on the report: use kLenPortrait or kLenLandscape
'    'ReportObj.SelectString = SelectObj.sGetUserReadableWhereClause(kLenPortrait)
'
'    'CUSTOMIZE:  If using work tables, restrict report data to current Session ID.  If using
'    'real tables, might restrict report data to current company or other criteria.
'    If (frmRequistn.ReportObj.lRestrictBy("{tciErrorLog.SessionID} = " & mlSession & " AND {tciErrorLog.Severity} > 0") = kFailure) Then
'        GoTo badexit
'    End If
'
'    frmRequistn.ReportObj.ProcessReport frmSelectReqLines, kTbPreview, iFileType, sFileName
'
''    ShowStatusNone frmSelectReqLines
'
'    lStartErrorReport = kSuccess
'
''    Set ReportObj = Nothing
''    Set DBObj = Nothing
''    Set oDDData = Nothing
''+++ VB/Rig Begin Pop +++
''+++ VB/Rig End +++
'    Exit Function
'
'badexit:
'    'ReportObj.CleanupWorkTables
'    'Set SelectObj = Nothing
''    Set ReportObj = Nothing
''    Set DBObj = Nothing
''    Set oDDData = Nothing
'    gClearSotaErr
''+++ VB/Rig Begin Pop +++
''+++ VB/Rig End +++
'    Exit Function
''+++ VB/Rig Begin Pop +++
'#If ERRORTRAPON = 0 Then
'Err.Raise Err
'#End If
'VBRigErrorRoutine:
'        gSetSotaErr Err, sMyName, "lStartReport", VBRIG_IS_MODULE
'        Err.Raise guSotaErr.Number
''+++ VB/Rig End +++
'End Function

Private Sub Form_Resize()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then

      'resize Height
       gResizeForm kResizeDown, Me, miOldFormHeight, miMinFormHeight, grdSelectPush
           
      'resize Width
        gResizeForm kResizeRight, Me, miOldFormWidth, miMinFormWidth, grdSelectPush
               
   
        miOldFormHeight = Me.Height
        miOldFormWidth = Me.Width
        

    End If

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Form_Resize", VBRIG_IS_FORM
        Select Case VBRIG_IS_FORM_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub Form_Unload(Cancel As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        moFormCust.UnloadSelf
        Set moFormCust = Nothing
    End If
#End If
    
    ' Clean out the Error Log table for this session
    CleanOutErrorLog mlSession

    moPushPull.Terminate
    DropTempTable
    moDmSelectGrid.UnloadSelf
    moSelectGM.UnloadSelf
    
    Set moContextMenu = Nothing
    Set moClass = Nothing
    Set moDmSelectGrid = Nothing
    Set moSelectGM = Nothing
    Set mcolUpdCounter = Nothing

'    Set frmSelectReqLines = Nothing
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Form_Unload", VBRIG_IS_FORM
        Select Case VBRIG_IS_FORM_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub grdSelectPush_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    moDmSelectGrid.SetRowDirty Row
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdSelectPush_KeyDown", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub grdSelectPush_KeyDown(KeyCode As Integer, Shift As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moSelectGM Is Nothing Then
        moSelectGM.Grid_KeyDown KeyCode, Shift
    End If
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdSelectPush_KeyDown", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++

End Sub

Private Sub moDmSelectGrid_DMGridBeforeUpdate(lRow As Long, bValid As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
' Force the checkbox into the DM (for some reason, it is not happening automatically)
    If gsGridReadCell(grdSelectPush, lRow, kColSelect) = "1" Then
        moDmSelectGrid.SetColumnValue lRow, "Selected", 1
    Else
        moDmSelectGrid.SetColumnValue lRow, "Selected", 0
    End If

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moDmSelectGrid_DMGridBeforeUpdate", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub


Private Sub grdSelectPush_Change(ByVal Col As Long, ByVal Row As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lLastRow As Long
    
    If Not moSelectGM Is Nothing Then
        lLastRow = grdSelectPush.MaxRows
        moSelectGM.Grid_Change Col, Row
        If grdSelectPush.MaxRows > lLastRow Then
            grdSelectPush.MaxRows = lLastRow
        End If
    End If

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdSelectPush_Change", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub AfterSelectAll(lRows As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lLoop As Long
    
    For lLoop = 1 To lRows
        moDmSelectGrid.SetRowDirty (lLoop)
    Next
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moPushPull_AfterSelectAll", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub moPushPull_BindGrid()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
        
    If Not moDmSelectGrid Is Nothing Then
        moDmSelectGrid.UnloadSelf
        Set moDmSelectGrid = Nothing
    End If
    If Not moSelectGM Is Nothing Then
        moSelectGM.UnloadSelf
        Set moSelectGM = Nothing
    End If
    If Not moContextMenu Is Nothing Then
        Set moContextMenu = Nothing
    End If

    Set moDmSelectGrid = New clsDmGrid

    With moDmSelectGrid
        Set .Form = frmSelectReqLines
        Set .Session = moClass.moSysSession
        Set .Grid = grdSelectPush
        Set .Database = moClass.moAppDB
        If moPushPull.LoadImmediate Then
            .UIType = kDmUISingle
        End If
        .Table = "#tpoPushPull"
        .UniqueKey = "DocLineKey"
        .OrderBy = "DocLineKey"
        .SaveOrder = 1

        .BindColumn "DocLineKey", kColDocLineKey, SQL_INTEGER
        .BindColumn "Selected", kColSelect, SQL_SMALLINT
        .BindColumn "Qty", kColQty, SQL_DECIMAL
        .BindColumn "RequestDate", kColReqDate, SQL_DATE
        

        .LinkSource "#tpoPushPullDetail", "#tpoPushPullDetail.LineKey=#tpoPushPull.DocLineKey", kDmJoin
        .Link kColDescription, "Description"
        .Link kColDeptID, "PurchDeptID"
        .Link kColVendorID, "VendID"
        .Link kcolItemID, "ItemID"
        .Link kcolWhseID, "WhseID"
        
        .Init
    End With
    
    ' Remove the last row.  It seems that the grid automatically adds a row to
    ' the end of the grid.
    grdSelectPush.MaxRows = grdSelectPush.MaxRows - 1
    FetchUpdCnt
    BindGM
    BindContextMenu
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moPushPull_BindGrid", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub moPushPull_CreateMoreTempTables()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
   Dim sAPIValidSQL As String
   Dim sPOLineSQL As String
   Dim sPOLineDistSQL As String
   Dim sTaxCodeTranSQL As String
   Dim sTaxDeleteSQL As String
   Dim sTaxTranSQL As String
   Dim sPushPullSQL As String
   
   gAPITables sAPIValidSQL, sPOLineSQL, sPOLineDistSQL, sTaxCodeTranSQL, _
sTaxDeleteSQL, sTaxTranSQL, sPushPullSQL

    On Error Resume Next
   moClass.moAppDB.ExecuteSQL sAPIValidSQL
   moClass.moAppDB.ExecuteSQL sPOLineSQL
   moClass.moAppDB.ExecuteSQL sPOLineDistSQL
   moClass.moAppDB.ExecuteSQL sTaxCodeTranSQL
   moClass.moAppDB.ExecuteSQL sTaxDeleteSQL
   moClass.moAppDB.ExecuteSQL sTaxTranSQL
   moClass.moAppDB.ExecuteSQL sPushPullSQL

'    CreateAPIValidTemp
'    CreatePOLineTemp
'    CreatePOLineDistTemp
'    CreateSTaxCodeTranTemp
'    CreateSTaxTranTemp
'    CreateSTaxDeleteTemp
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moPushPull_CreateMoreTempTables", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CreateSTaxDeleteTemp()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    
    sSql = "CREATE TABLE #tciSTaxDelete" & _
            "(STaxTranKey       int       NOT NULL, " & _
            "NeedsDelete       smallint  NOT NULL)"
            
        
   On Error Resume Next
   moClass.moAppDB.ExecuteSQL sSql


'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CreateSTaxDeleteTemp", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CreateSTaxTranTemp()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    
    sSql = "CREATE TABLE #tciSTaxTran" & _
            "(STaxTranKey       int       NOT NULL, " & _
            "STaxSchdKey       int       NULL, " & _
            "NeedInsert        smallint  NOT NULL)"
        
   On Error Resume Next
   moClass.moAppDB.ExecuteSQL sSql


'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CreateSTaxTranTemp", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CreateSTaxCodeTranTemp()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    
    sSql = "SELECT * " & _
            "INTO   #tciSTaxCodeTran " & _
            "From tciSTaxCodeTran " & _
            "Where 1 = 2 "
        
        On Error Resume Next
   moClass.moAppDB.ExecuteSQL sSql

'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CreateSTaxCodeTranTemp", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub CreatePOLineDistTemp()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++

    Dim sSql As String
    
    sSql = "CREATE TABLE #tpoPOLineDist" & _
            "(POLineDistKey      int       NOT NULL , " & _
            "AcctRefKey         int       NULL , " & _
            "AmtInvcd           dec(15,3) NOT NULL , " & _
            "BlnktPOLineDistKey int       NULL , " & _
            "ClosedForInvc      smallint  NOT NULL , " & _
            "ClosedForRcvg      smallint  NOT NULL , " & _
            "DropShip           smallint  NOT NULL , " & _
            "ExclLeadTime       smallint  NOT NULL , " & _
            "ExclLTReasCodeKey  int       NULL , " & _
            "Expedite           smallint  NOT NULL , " & _
            "ExtAmt             dec(15,3) NOT NULL , " & _
            "FOBKey             int       NULL , " & _
            "ExpediteReasonKey  int       NULL , " & _
            "FreightAmt         dec(15,3) NOT NULL , " & _
            "GLAcctKey          int       NULL , " & _
            "OrigOrdered        dec(16,8) NOT NULL , " & _
            "OrigPromiseDate    datetime  NULL , " & _
            "POLineKey          int       NOT NULL , " & _
            "PurchDeptKey       int       NULL , " & _
            "PromiseDate        datetime  NULL , " & _
            "QtyInvcd           dec(16,8) NOT NULL , " & _
            "QtyOnBO            dec(16,8) NOT NULL , " & _
            "QtyOpenToRcv       dec(16,8) NOT NULL , " & _
            "QtyOrd             dec(16,8) NOT NULL , "
                  sSql = sSql & "QtyRcvd            dec(16,8) NOT NULL , " & _
            "QtyRtrnCredit      dec(16,8) NOT NULL , " & _
            "QtyRtrnReplacement dec(16,8) NOT NULL , " & _
            "RequestDate        datetime  NULL , " & _
            "ShipMethKey        int       NULL , " & _
            "ShipToAddrKey      int       NULL , " & _
            "ShipToCustAddrKey  int       NULL , " & _
            "ShipToCustKey      int       NULL , " & _
            "ShipToWhseKey      int       NULL , " & _
            "ShipZoneKey        int       NULL , " & _
            "Status             smallint  NOT NULL , " & _
            "STaxTranKey        int       NULL , " & _
            "ShipToAddrName     varchar(40)  NULL, " & _
            "ShipToAddrLine1    varchar(40)  NULL, " & _
            "ShipToAddrLine2    varchar(40)  NULL, " & _
            "ShipToAddrLine3    varchar(40)  NULL, " & _
            "ShipToAddrLine4    varchar(40)  NULL, " & _
            "ShipToAddrLine5    varchar(40)  NULL, " & _
            "ShipToCity         varchar(20)  NULL, " & _
            "ShipToState        varchar(03)  NULL, " & _
            "ShipToCountryID    varchar(03)  NULL, " & _
            "ShipToPostalCode   varchar(09)  NULL, " & _
            "CreateShipTo       smallint  NULL, " & _
            "UpdateCounter      int       NOT NULL)"

    On Error Resume Next
   moClass.moAppDB.ExecuteSQL sSql


'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CreatePOLineDistTemp", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub CreateAPIValidTemp()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    
    sSql = "CREATE TABLE #tpoAPIValid" & _
        "(TrackSTax          smallint   NULL, " & _
        "MatchLevel         smallint   NULL, " & _
        "ChkCreditLimit     smallint   NULL, " & _
        "IntegrateWithIM    smallint   NULL, " & _
        "DeptOvrdSegKey     int        NULL, " & _
        "POLineMatchTolKey  int        NULL, " & _
        "POMatchTolKey      int        NULL, " & _
        "PrintPOs           smallint   NULL, " & _
        "UseBlnktRelNos     smallint   NULL, " & _
        "DfltRcvrAddrKey    int        NULL, " & _
        "QtyDecPlaces       smallint   NULL, " & _
        "UnitCostDecPlaces  smallint   NULL, " & _
        "AcctRefUsage       smallint   NULL, " & _
        "AutoAcctAdd        smallint   NULL, " & _
        "HomeCurrID         varchar(03)   NULL, " & _
        "HomeCurrDP         smallint   NULL, " & _
        "CompanyID          varchar(03)   NULL, " & _
        "Use1099            smallint   NULL, " & _
        "AllowInterComp     smallint   NULL, " & _
        "APSTaxSchdKey      int        NULL, " & _
        "UseMultCurr        smallint   NULL, " & _
        "VCOvrdSegKey       int        NULL, " & _
        "RecordNumber       int        NULL, " & _
        "Spid               int        NULL, "
           sSql = sSql & "LastRetVal         smallint   NULL, " & _
        "UFDataType1        smallint   NULL, " & _
        "UFUsage1           smallint   NULL, " & _
        "UFKey1             integer    NULL, " & _
        "UFDataType2        smallint   NULL, " & _
        "UFUsage2           smallint   NULL, " & _
        "UFKey2             integer    NULL, " & _
        "UFDataType3        smallint   NULL, " & _
        "UFUsage3           smallint   NULL, " & _
        "UFKey3             integer    NULL, " & _
        "UFDataType4        smallint   NULL, " & _
        "UFUsage4           smallint   NULL, " & _
        "UFKey4             integer    NULL, " & _
        "UFDataTypel1       smallint   NULL, " & _
        "UFUsagel1          smallint   NULL, " & _
        "UFKeyl1            integer    NULL, " & _
        "UFDataTypel2       smallint   NULL, " & _
        "UFUsagel2          smallint   NULL, " & _
        "UFKeyl2            integer    NULL, " & _
        "POKey              int        NULL, " & _
        "AmtInvcd           dec(15,3)  NULL, " & _
        "ApprovalDate       datetime   NULL, " & _
        "ApprovalStatus     smallint   NULL, " & _
        "BegDate            datetime   NULL, " & _
        "BlnktPOKey         integer    NULL, "
           sSql = sSql & "BlnktRelNo         smallint   NULL, " & _
        "BuyerKey           int        NULL, " & _
        "ChngOrdDate        datetime   NULL, " & _
        "ChngOrdNo          smallint   NULL, " & _
        "ChngReason         varchar(40)   NULL, " & _
        "ChngUserID         varchar(30)   NULL, " & _
        "CloseDate          datetime   NULL, " & _
        "ClosedForInvc      smallint   NULL, " & _
        "ClosedForRcvg      smallint   NULL, " & _
        "CntctKey           integer    NULL, " & _
        "CreateDate         datetime   NULL, " & _
        "CreateType         smallint   NULL, " & _
        "CreateUserID       varchar(30)   NULL, " & _
        "CurrExchRate       float      NULL, " & _
        "CurrExchSchdKey    integer    NULL, " & _
        "CurrID             varchar(03)   NULL, " & _
        "DfltAcctRefKey     integer    NULL, " & _
        "DfltDropShip       smallint   NULL, " & _
        "DfltExclLastCost   smallint   NULL, " & _
        "DfltExclLeadTime   smallint   NULL, " & _
        "DfltExclLTReasKey  integer    NULL, " & _
        "DfltExpedite       smallint   NULL, " & _
        "DfltExpediteRsnKey integer    NULL, " & _
        "DfltPurchDeptKey   integer    NULL, " & _
        "DfltRequestDate    datetime   NULL, "
           sSql = sSql & "DfltShipMethKey    integer    NULL, " & _
        "DfltShipToAddrKey  integer    NULL, " & _
        "DfltShipToCAddrKey integer    NULL, " & _
        "DfltShipToCustKey  integer    NULL, " & _
        "DfltShipToWhseKey  integer    NULL, " & _
        "DfltShipZoneKey    integer    NULL, " & _
        "DfltTargetCompID   varchar(03)   NULL, " & _
        "FOBKey             integer    NULL, " & _
        "FreightAllocMeth   smallint   NULL, " & _
        "FreightAmt         dec(15,3)  NULL, " & _
        "Hold               smallint   NULL, " & _
        "HoldReason         varchar(20)   NULL, " & _
        "ImportLogKey       integer    NULL, " & _
        "IssueDate          datetime   NULL, " & _
        "MatchToleranceKey  integer    NULL, " & _
        "NextLineNo         integer    NULL, " & _
        "OpenAmt            dec(15,3)  NULL, " & _
        "OpenAmtHC          dec(15,3)  NULL, " & _
        "OriginationDate    datetime   NULL, " & _
        "PmtTermsKey        integer    NULL, " & _
        "POFormKey          integer    NULL, " & _
        "PurchAddrKey       integer    NULL, " & _
        "PurchAmt           dec(15,3)  NULL, " & _
        "PurchVendAddrKey   integer    NULL, " & _
        "RecurPOKey         integer    NULL, "
           sSql = sSql & "RemitToAddrKey     integer    NULL, " & _
        "RemitToVendAddrKey integer    NULL, " & _
        "Status             smallint   NULL, " & _
        "STaxAmt            dec(15,3)  NULL, " & _
        "STaxTranKey        integer    NULL, " & _
        "StopDate           datetime   NULL, " & _
        "TranAmt            dec(15,3)  NULL, " & _
        "TranAmtHC          dec(15,3)  NULL, " & _
        "TranCmnt           varchar(50)   NULL, " & _
        "TranDate           datetime   NULL, " & _
        "TranID             varchar(13)   NULL, " & _
        "TranNo             varchar(10)   NULL, " & _
        "TranNoChngOrd      varchar(15)   NULL, " & _
        "TranNoRel          varchar(15)   NULL, " & _
        "TranNoRelChngOrd   varchar(20)   NULL, " & _
        "TranType           integer    NULL, " & _
        "UpdateCounter      integer    NULL, " & _
        "UpdateDate         datetime   NULL, " & _
        "UpdateUserID       varchar(30)   NULL, " & _
        "UserFld1           varchar(15)   NULL, " & _
        "UserFld2           varchar(15)   NULL, " & _
        "UserFld3           varchar(15)   NULL, " & _
        "UserFld4           varchar(15)   NULL, " & _
        "V1099Box           varchar(03)   NULL, " & _
        "V1099BoxText       varchar(15)   NULL, "
           sSql = sSql & "V1099Form          smallint   NULL, " & _
        "VendClassKey       integer    NULL, " & _
        "VendKey            integer    NULL, " & _
        "VendQuoteKey       integer    NULL, " & _
        "ClassOvrdSegValue  varchar(15)   NULL, " & _
        "DeptOvrdSegValue   varchar(15)   NULL, " & _
        "DSTAddrName        varchar(40)   NULL, " & _
        "DSTAddrLine1       varchar(40)   NULL, " & _
        "DSTAddrLine2       varchar(40)   NULL, " & _
        "DSTAddrLine3       varchar(40)   NULL, " & _
        "DSTAddrLine4       varchar(40)   NULL, " & _
        "DSTAddrLine5       varchar(40)   NULL, " & _
        "DSTCity            varchar(20)   NULL, " & _
        "DSTState           varchar(03)   NULL, " & _
        "DSTCountryID       varchar(03)   NULL, " & _
        "DSTPostalCode      varchar(09)   NULL, " & _
        "RTAddrName         varchar(40)   NULL, " & _
        "RTAddrLine1        varchar(40)   NULL, " & _
        "RTAddrLine2        varchar(40)   NULL, " & _
        "RTAddrLine3        varchar(40)   NULL, " & _
        "RTAddrLine4        varchar(40)   NULL, " & _
        "RTAddrLine5        varchar(40)   NULL, " & _
        "RTCity             varchar(20)   NULL, " & _
        "RTState            varchar(03)   NULL, " & _
        "RTCountryID        varchar(03)   NULL, "
           sSql = sSql & "RTPostalCode       varchar(09)   NULL, " & _
        "PurchAddrName      varchar(40)   NULL, " & _
        "PurchAddrLine1     varchar(40)   NULL, " & _
        "PurchAddrLine2     varchar(40)   NULL, " & _
        "PurchAddrLine3     varchar(40)   NULL, " & _
        "PurchAddrLine4     varchar(40)   NULL, " & _
        "PurchAddrLine5     varchar(40)   NULL, " & _
        "PurchCity          varchar(20)   NULL, " & _
        "PurchState         varchar(03)   NULL, " & _
        "PurchCountryID     varchar(03)   NULL, " & _
        "PurchPostalCode    varchar(09)   NULL, " & _
        "CreateShipTo       smallint   NULL, " & _
        "CreateRemitTo      smallint   NULL, " & _
        "CreatePurchFrom    smallint   NULL, " & _
        "UniqueID           varchar(80)   NULL, " & _
        "DocRoundAmt        smallint   NULL, " & _
        "CreditLimit        dec(15,3)  NULL, " & _
        "VendSTaxSchdKey    integer    NULL, " & _
        "TCDeptOvrdSegKey   integer    NULL, " & _
        "TCAcctRefUsage     smallint   NULL, " & _
        "TCDfltRcvrAddrKey  integer    NULL, " & _
        "TCAutoAcctAdd      smallint   NULL, " & _
        "ClassOvrdGL        smallint   NULL, " & _
        "VendID             varchar(12)   NULL, " & _
        "VendAcctKey        integer    NULL, "
           sSql = sSql & "LogSuccessful      smallint   NULL, " & _
        "UserID            varchar(30)   NULL, " & _
        "RequirePOIssue    smallint   NULL)"
   
    On Error Resume Next
   moClass.moAppDB.ExecuteSQL sSql

'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CreateAPIValidTemp", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub CreatePOLineTemp()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++

    Dim sSql As String
    
    sSql = "CREATE TABLE #tpoPOLine" & _
        "(POLineKey          int          NOT NULL, " & _
        "CloseDate          datetime     NULL, " & _
        "ClosedForInvc      smallint     NOT NULL, " & _
        "ClosedForRcvg      smallint     NOT NULL, " & _
        "CmntOnly           smallint     NOT NULL, " & _
        "Description        varchar(40)  NULL, " & _
        "ExclLastCost       smallint     NOT NULL, " & _
        "ExtAmt             dec(15,3)    NULL, " & _
        "ExtCmnt            varchar(255) NULL, " & _
        "ItemKey            int          NULL, " & _
        "MatchToleranceKey  int          NULL, " & _
        "POKey              int          NOT NULL, " & _
        "POLineNo           smallint     NULL, " & _
        "ReqRcvr            smallint     NOT NULL, " & _
        "Status             smallint     NOT NULL, " & _
        "STaxClassKey       int          NULL, " & _
        "TargetCompanyID    varchar(03)     NOT NULL, " & _
        "UnitCost           dec(15,5)    NULL, " & _
        "UnitMeasKey        int          NULL, " & _
        "UpdateCounter      int          NOT NULL, " & _
        "UserFld1           varchar(15)     NULL, " & _
        "UserFld2           varchar(15)     NULL, " & _
        "VendQuoteLineKey   int          NULL, " & _
        "CalcQty            smallint     NULL, "
            sSql = sSql & "CalcCost           smallint     NULL, " & _
        "CalcExtAmt         smallint     NULL, " & _
        "StdUnitCost        dec(15,5)    NULL, " & _
        "ItemGLAcctKey      int          NULL, " & _
        "NoDfltCostRsn      smallint     NULL, " & _
        "TCDeptOvrdSegKey   integer      NULL, " & _
        "TCAcctRefUsage     smallint     NULL, " & _
        "TCDfltRcvrAddrKey  integer      NULL, " & _
        "TCAutoAcctAdd      smallint     NULL, " & _
        "ItemID             varchar(30)  NULL, " & _
        "UnitCostExact      dec(25,13)   NULL) "

   On Error Resume Next
   moClass.moAppDB.ExecuteSQL sSql

'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CreatePOLineTemp", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub moPushPull_FormatGrid()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    gGridSetProperties grdSelectPush, kMaxCols, kGridDataSheetNoAppend
    grdSelectPush.UserResizeCol = SS_USER_RESIZE_OFF

    
    grdSelectPush.MaxCols = kMaxCols
    grdSelectPush.MaxRows = kMaxRows
    
    gGridSetHeader grdSelectPush, kColSelect, "Select"
    gGridSetHeader grdSelectPush, kcolItemID, "Item"
    gGridSetHeader grdSelectPush, kColDescription, "Description"
    gGridSetHeader grdSelectPush, kColQty, "Quantity"
    gGridSetHeader grdSelectPush, kColReqDate, "Request Date"
    gGridSetHeader grdSelectPush, kColVendorID, "Vendor"
    gGridSetHeader grdSelectPush, kColDeptID, "Department"
    gGridSetHeader grdSelectPush, kcolWhseID, "Warehouse"
    
    gGridSetColumnType grdSelectPush, kColSelect, SS_CELL_TYPE_CHECKBOX
    gGridSetColumnType grdSelectPush, kColReqDate, SS_CELL_TYPE_DATE
    
    gGridSetColumnWidth grdSelectPush, kColSelect, 6
    gGridSetColumnWidth grdSelectPush, kcolItemID, 12
    gGridSetColumnWidth grdSelectPush, kColDescription, 18
    gGridSetColumnWidth grdSelectPush, kColQty, 12
    gGridSetColumnWidth grdSelectPush, kColReqDate, 12
    gGridSetColumnWidth grdSelectPush, kColVendorID, 12
    gGridSetColumnWidth grdSelectPush, kColDeptID, 12
    gGridSetColumnWidth grdSelectPush, kcolWhseID, 12
    
    gGridLockColumn grdSelectPush, kcolItemID
    gGridLockColumn grdSelectPush, kColDescription
    gGridLockColumn grdSelectPush, kColQty
    gGridLockColumn grdSelectPush, kColReqDate
    gGridLockColumn grdSelectPush, kColVendorID
    gGridLockColumn grdSelectPush, kColDeptID
    gGridLockColumn grdSelectPush, kcolWhseID
    
    gGridHideColumn grdSelectPush, kColDocLineKey
    
    'Make the Items column frozen
    gGridFreezeCols grdSelectPush, kColSelect
 
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moPushPull_FormatGrid", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub DropTempTable()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    
    sSql = "DROP TABLE #tpoPushPullDetail"
        
    On Error Resume Next
    moClass.moAppDB.ExecuteSQL sSql

'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DropTempTable", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub


Private Sub moPushPull_LoadTempTable()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    
    sSql = "DELETE FROM #tpoPushPull"
    moClass.moAppDB.ExecuteSQL sSql
    
    sSql = "INSERT INTO #tpoPushPull " & _
            "(SessionID, Selected, DocLineKey, Qty, RequestDate) " & _
            "SELECT " & mlSession & ", 0, b.ReqLineKey, c.QtyReq, c.RequestDate " & _
            "FROM tpoRequisition a " & _
            " JOIN tpoReqLine b ON a.ReqKey = b.ReqKey " & _
            " JOIN tpoReqLineDist c ON b.ReqLineKey = c.ReqLineKey " & _
            "WHERE a.TranNo = " & gsQuoted(msReqID) & " " & _
            "AND b.POLineKey IS NULL " & _
            "AND a.CompanyID = " & gsQuoted(moClass.moSysSession.CompanyId)
        
    moClass.moAppDB.ExecuteSQL sSql
    
    sSql = "CREATE TABLE #tpoPushPullDetail(" & _
            "LineKey INTEGER NULL, " & _
            "ItemID varCHAR (30) NULL, " & _
            "Description VARCHAR(40) NULL, " & _
            "VendID varCHAR(12) NULL, " & _
            "PurchDeptID varCHAR (15) NULL, " & _
            "WhseID varCHAR (6) NULL) "

    On Error Resume Next
    moClass.moAppDB.ExecuteSQL sSql
    If Err.Number <> 0 Then
        sSql = "DELETE FROM #tpoPushPullDetail"
        moClass.moAppDB.ExecuteSQL sSql
    End If


    sSql = "INSERT INTO #tpoPushPullDetail " & _
            "SELECT a.DocLineKey, d.ItemID," & _
            " b.Description, e.VendID, f.PurchDeptID, w.WhseID " & _
            " FROM #tpoPushPull a " & _
            " JOIN  tpoReqLine b ON a.DocLineKey = b.ReqLineKey" & _
            " JOIN  tpoReqLineDist c ON b.ReqLineKey = c.ReqLineKey" & _
            " LEFT OUTER JOIN  timItem d ON b.ItemKey = d.ItemKey" & _
            " LEFT OUTER JOIN  tapVendor e ON b.VendKey = e.VendKey" & _
            " LEFT OUTER JOIN  tpoPurchDepartment f ON c.PurchDeptKey = f.PurchDeptKey" & _
            " LEFT OUTER JOIN timWarehouse w ON c.ShipToWhseKey = w.WhseKey"
            
    moClass.moAppDB.ExecuteSQL sSql
    
'+++ VB/Rig Begin Pop +++
        Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moPushPull_LoadTempTable", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub moPushPull_ProcessProceed()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim lRetVal As Long
    Dim lRowCount As Long
    'Agregado por Multiconsulting
    Dim lPOKey As Long
    'Agregado por Multiconsulting
    
   lRowCount = grdSelectPush.MaxRows
' First save the grid to the temp table.
   moDmSelectGrid.Save
' Turn the hourglass back on (Save seems to turn it off)
   SetHourglass True
   
   
' The save may add an extra line at the end if the last row was checked.
' This extra row must be saved.
   grdSelectPush.MaxRows = lRowCount

    If bNoRowsSelected Then

'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
        SetHourglass False
        Exit Sub
    End If
    
    If ValidateUpdCnt(lRowCount) = False Then
        SetHourglass False
        Exit Sub
    End If
    
' Now clean out the Error Log table for this session
    CleanOutErrorLog mlSession
'    lContractKeyUsed = glGetValidLong(moClass.moAppDB.Lookup("ContractKey", "tpoRequisitionContract", "ReqKey =" & mlReqKey))
'   DisplayTempData
' Call the stored procedure to create the stored procedure
    With moClass.moAppDB
        .SetInParam mlSession
        .SetInParam moClass.moSysSession.CompanyId
        .SetInParamInt kTranTypePORQ
        .SetInParamInt kTranTypePOPO
        .SetInParam mlReqKey
        .SetInParam ""
        .SetInParam moClass.moSysSession.UserId
        If chkCostFromReq = vbChecked Then
            .SetInParam 0
        Else
            .SetInParam 1       'Get Cost from IMS
        End If
        .SetOutParam lRetVal
        On Error Resume Next
        .ExecuteSP ("sppoPushPull")
        If Err.Number <> 0 Then
            lRetVal = Err.Number
'            MsgBox "Return Val = " & Err.Description
        Else
            'CUSTOMIZE:  If procedure has output parameters, attempt to retrieve them, checking for
            'errors after each attempt.  After storing all return values in local variables OR
            'encountering an error, release parameters.
            lRetVal = moClass.moAppDB.GetOutParam(9)
'            MsgBox "Return Val = " & lRetval
        End If
        .ReleaseParams

    End With

' Now call the procedure to write the new PO Line Keys back to the
' tpoReqLine table.
'    DisplayTempData
    If lRetVal = 1 Then
        If UpdateExistingLines = 0 Then
            UpdateReqStatus
            UpdatePOReqTable
            'Agregado por Multiconsulting
            If glGetValidLong(moClass.moAppDB.Lookup("count(DocLineKey)", "#tpoPushPull", "Selected = 1 AND ErrorCode IS NULL")) > 0 Then
                lPOKey = moClass.moAppDB.Lookup("POKEY", "tpoPurchOrder", "1 = 1 order by POKey desc")
                If lContractKeyUsed > 0 Then
                    moClass.moAppDB.ExecuteSQL "update tpoPurchOrder set ContractKey = " & lContractKeyUsed & " where POKey =" & lPOKey
                End If
                MsgBox "OC Generada: " & moClass.moAppDB.Lookup("TranNo", "tpoPurchOrder", "POKey = " & lPOKey)
            End If
            'Agregado por Multiconsulting
            DisplayGenerationStatus
        Else
'            giSotaMsgBox Nothing, moClass.moSysSession, 220071
           giSotaMsgBox Me, moClass.moSysSession, kmsgReqGenErrorEncountered
        End If
    Else
'        giSotaMsgBox Nothing, moClass.moSysSession, 220071
        giSotaMsgBox Me, moClass.moSysSession, kmsgReqGenErrorEncountered
    End If
   ' Since the sethourglass function keeps a counter of how many times it was turned on,
   ' I have to turn it off an equal amount of times.
   SetHourglass False

'+++ VB/Rig Begin Pop +++
        Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moPushPull_ProcessProceed", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Function bNoRowsSelected() As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

' Determine if any rows were selected.

    Dim sSql As String
    Dim rs As Object
    sSql = "SELECT DocLineKey FROM #tpoPushPull " & _
            "WHERE Selected = 1 "

    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    If rs.IsEOF Then
        bNoRowsSelected = True
    Else
        bNoRowsSelected = False
    End If
        
    Set rs = Nothing
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bNoRowsSelected", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function


Private Sub DisplayGenerationStatus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

' Determine if all selected lines were generated into PO's.  If so display the successful message.
' Otherwise, display the partially successful message.

    Dim sSql As String
    Dim rs As Object
    sSql = "SELECT DocLineKey FROM #tpoPushPull " & _
            "WHERE Selected = 1 " & _
            "AND ErrorCode IS NOT NULL"

    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    If Not rs.IsEOF Then
'        giSotaMsgBox Nothing, moClass.moSysSession, 220070
        giSotaMsgBox Me, moClass.moSysSession, kmsgReqGenPartialSuccessful
    Else
'        giSotaMsgBox Nothing, moClass.moSysSession, 220069
        giSotaMsgBox Me, moClass.moSysSession, kmsgReqGenSuccessful
    End If
        
    Set rs = Nothing
    
    sSql = "SELECT SessionID FROM tciErrorLog " & _
"WHERE SessionID = " & mlSession & " " & _
"AND Severity > 0"

    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    If Not rs.IsEOF Then
        frmRequistn.PrintErrorReport = True
    End If

    Set rs = Nothing

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DisplayGenerationStatus", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub DisplayTempData()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim sSql As String
    Dim rs As Object
    sSql = "SELECT DocLineKey, Qty, RequestDate, Selected, ReturnValue, ErrorCode FROM #tpoPushPull " & _
            "WHERE Selected = 1"

    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    While Not rs.IsEOF
        MsgBox "DOC " & rs.Field("DocLineKey")
        MsgBox "Selected " & rs.Field("Selected")
        MsgBox "Ret Val " & rs.Field("ReturnValue")
        MsgBox "Error Code " & rs.Field("ErrorCode")
        MsgBox "Request date " & rs.Field("RequestDate")
        MsgBox "Qty " & rs.Field("Qty")
        rs.MoveNext
    Wend
    
    Set rs = Nothing

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DisplayTempData", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub CleanOutErrorLog(lSessionID As Long)
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
'    DisplayTempData
    
    
    sSql = "DELETE FROM tciErrorLog WHERE SessionID = " & lSessionID
    On Error Resume Next
     moClass.moAppDB.ExecuteSQL sSql
     
'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CleanOutErrorLog", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Function UpdateExistingLines()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    Dim rs As Object
    
'    DisplayTempData
'    First, update the generated PO's to change the CreateType to a 1 because the API
'       automatically sets the CreateType to 0.
    sSql = "UPDATE tpoPurchOrder SET tpoPurchOrder.CreateType = 1 " & _
            " FROM #tpoPushPull " & _
            " JOIN tpoPOLine ON #tpoPushPull.ReturnValue = tpoPOLine.POLineKey " & _
            " JOIN tpoPurchOrder ON tpoPOLine.POKey = tpoPurchOrder.POKey " & _
            " WHERE #tpoPushPull.ReturnValue IS NOT NULL " & _
            " AND #tpoPushPull.ErrorCode IS NULL"
    On Error Resume Next
     moClass.moAppDB.ExecuteSQL sSql
     
     UpdateExistingLines = Err.Number
        
    If Err.Number <> 0 Then
'+++ VB/Rig Begin Pop +++
        Exit Function
'+++ VB/Rig End +++
    End If

    sSql = "UPDATE tpoReqLine SET tpoReqLine.POLineKey = #tpoPushPull.ReturnValue, tpoReqLine.UpdateCounter = (tpoReqLine.UpdateCounter + 1) " & _
            "FROM #tpoPushPull" & _
            " JOIN tpoReqLine ON tpoReqLine.ReqLineKey = #tpoPushPull.DocLineKey " & _
            " WHERE #tpoPushPull.ReturnValue IS NOT NULL " & _
            " AND #tpoPushPull.ErrorCode IS NULL"
    On Error Resume Next
     moClass.moAppDB.ExecuteSQL sSql
     
     UpdateExistingLines = Err.Number
    
'+++ VB/Rig Begin Pop +++
Exit Function
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "UpdateExistingLines", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function
Private Sub UpdateReqStatus()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    Dim rs As Object
     
     sSql = "SELECT tpoReqLine.ReqLineKey " & _
            " FROM tpoReqLine " & _
            " JOIN tpoRequisition ON tpoRequisition.ReqKey = tpoReqLine.ReqKey " & _
            " WHERE tpoRequisition.TranNo = " & gsQuoted(msReqID) & " " & _
            " AND tpoRequisition.CompanyID = " & gsQuoted(moClass.moSysSession.CompanyId) & " " & _
            " AND tpoReqLine.POLineKey IS NULL"
        
    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    If rs.IsEOF Then
        sSql = "UPDATE tpoRequisition SET Status = " & kvReqClosed & " " & _
                " WHERE TranNo = " & gsQuoted(msReqID) & _
                " AND CompanyID = " & gsQuoted(moClass.moSysSession.CompanyId)
    On Error Resume Next
        moClass.moAppDB.ExecuteSQL sSql

    End If
    
    Set rs = Nothing
    'Update the status for tpoReqLine when PO generated
    
    sSql = "UPDATE TRL set TRL.Status = " & kvStatusClosed & _
       " FROM tpoReqLine TRL " & _
       " JOIN tpoRequisition TR WITH (NOLOCK) ON TR.ReqKey = TRL.ReqKey " & _
       " WHERE TR.TranNo = " & gsQuoted(msReqID) & " " & _
       " AND TR.CompanyID = " & gsQuoted(moClass.moSysSession.CompanyId) & " " & _
       " AND TRL.Status <> " & kvStatusClosed & _
       " AND TRL.POLineKey IS NOT NULL"
    moClass.moAppDB.ExecuteSQL sSql
  

'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "UpdateReqStatus", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub UpdatePOReqTable()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sSql As String
    Dim rs As Object
    Dim lReqKey As Long
    Dim lPOKey As Long

    lReqKey = 0
    sSql = "SELECT ReqKey FROM tpoRequisition " & _
            "WHERE TranNo = " & gsQuoted(msReqID) & " " & _
            "AND CompanyID = " & gsQuoted(moClass.moSysSession.CompanyId)
    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    If Not rs.IsEOF Then
        lReqKey = rs.Field("ReqKey")
    End If
    Set rs = Nothing
            

    If lReqKey > 0 Then
        sSql = "INSERT INTO tpoPurchOrdReq (POKey, ReqKey)" & _
                "SELECT Distinct a.POKey, " & lReqKey & " " & _
                "FROM tpoPOLine a" & _
                " JOIN #tpoPushPull b ON a.POLineKey = b.ReturnValue " & _
                "WHERE b.ReturnValue IS NOT NULL " & _
                "AND b.ErrorCode IS NULL"
      On Error Resume Next
        moClass.moAppDB.ExecuteSQL sSql
       
        'Scopus incident # 19512
        'Transfer comments from Requisition to PO
        lPOKey = 0
        sSql = "SELECT POKey FROM tpoPurchOrdReq " & _
                "WHERE ReqKey = " & lReqKey
        Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
       While Not rs.IsEOF
            lPOKey = rs.Field("POKey")
    
        If IsNumeric(lPOKey) Then
            
            sSql = "UPDATE tpoPurchOrder set TranCmnt=a.TranCmnt " & _
                    " FROM (select TranCmnt from tpoRequisition where ReqKey=" & lReqKey & ") a " & _
                    " WHERE POKey =" & lPOKey
            
            moClass.moAppDB.ExecuteSQL sSql
    
            sSql = "UPDATE tpoPOLine set ExtCmnt=a.ExtCmnt " & _
                    " FROM tpoReqLine a " & _
                    " WHERE a.ReqKey =" & lReqKey & _
                    " AND a.POLineKey=tpoPOLine.POLineKey"
            
            moClass.moAppDB.ExecuteSQL sSql
            
        End If
        rs.MoveNext
       Wend
       Set rs = Nothing
    End If

    
'+++ VB/Rig Begin Pop +++
Exit Sub
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "UpdatePOReqTable", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub tbrMain_ButtonClick(Button As String)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If moFormCust.ToolbarClick(Button) Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Sub
        End If
    End If
#End If
    
  
    Dim lRowCount As Long
    frmSelectReqLines.SetFocus
    DoEvents
    Select Case Button
        Case kTbProceed
            
            SetHourglass True
            If bIsValid Then
                moPushPull.ProceedPressed
            End If
            SetHourglass False
            frmSelectReqLines.Hide
            
        Case kTbClose
            frmSelectReqLines.Hide
        Case kTbHelp
            gDisplayFormLevelHelp Me
            SetHourglass False
        
    End Select

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "tbrMain_ButtonClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

'************************************************************************
'   FormHelpPrefix will contain the help prefix for the Form Level Help.
'   This is contstructed as:
'                      <ModuleID> & "Z" & <FormType>
'
'   <Module>   is "CI", "AP", "GL", . . .
'   "Z"        is the Sage MAS 500 identifier.
'   <FormType> is "M" = Maintenance, "D" = data entry, "I" = Inquiry,
'                 "P" = PeriodEnd, "R" = Reports, "L" = Listings, . . .
'************************************************************************
Public Property Get FormHelpPrefix() As String
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    FormHelpPrefix = "POZ"         '       Put your prefix here
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "FormHelpPrefix_Get", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Property
        End Select
'+++ VB/Rig End +++
End Property

'************************************************************************
'   WhatHelpPrefix will contain the help prefix for the What's This Help.
'   This is contstructed as:
'                      <ModuleID> & "Z" & <FormType>
'
'   <Module>   is "CI", "AP", "GL", . . .
'   "Z"        is the Sage MAS 500 identifier.
'************************************************************************
Public Property Get WhatHelpPrefix() As String
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    WhatHelpPrefix = "POZ"         '       Put your prefix here
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "WhatHelpPrefix_Get", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Property
        End Select
'+++ VB/Rig End +++
End Property

Private Sub BindContextMenu()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'***************************************************
'  Instantiate Context Menu Class
'***************************************************

    Set moContextMenu = New clsContextMenu  'Instantiate Context Menu Class
    
    With moContextMenu
       .BindGrid moSelectGM, grdSelectPush.hwnd
       .Bind "*APPEND", grdSelectPush.hwnd
       
      'Assign Winhook control to Context Menu Class
        ' **PRESTO ** Set .Hook = Me.WinHook1
        Set .Form = frmSelectReqLines
    
      'Init will set properties of Winhook to intercept WM_RBUTTONDOWN
        .Init
    
    End With

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "BindContextMenu", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub



#If CUSTOMIZER Then
Private Sub picDrag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
    On Error GoTo VBRigErrorRoutine:
#End If
'+++ VB/Rig End +++

    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseDown Index, Button, Shift, x, y
    End If

    Exit Sub

'+++ VB/Rig Begin Pop +++
    Exit Sub

VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "picDrag_MouseDown", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
            Err.Raise guSotaErr.Number
        Case Else
            Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++

End Sub
Private Sub picDrag_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
    On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseMove Index, Button, Shift, x, y
    End If

    Exit Sub

'+++ VB/Rig Begin Pop +++
    Exit Sub

VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "picDrag_MouseMove", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
            Err.Raise guSotaErr.Number
        Case Else
            Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++

End Sub
Private Sub picDrag_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
    On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseUp Index, Button, Shift, x, y
    End If

    Exit Sub

'+++ VB/Rig Begin Pop +++
    Exit Sub

VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "picDrag_MouseUp", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
            Err.Raise guSotaErr.Number
        Case Else
            Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++

End Sub
Private Sub picDrag_Paint(Index As Integer)

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
    On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_Paint Index
    End If

    Exit Sub

'+++ VB/Rig Begin Pop +++
    Exit Sub

VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "picDrag_Paint", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
            Err.Raise guSotaErr.Number
        Case Else
            Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++

End Sub
#End If

Private Sub cmdClearAll_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus cmdClearAll, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cmdClearAll_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cmdClearAll_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus cmdClearAll, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cmdClearAll_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cmdSelectAll_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus cmdSelectAll, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cmdSelectAll_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cmdSelectAll_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus cmdSelectAll, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cmdSelectAll_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub


Private Sub txtReqNo_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange txtReqNo, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtReqNo_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtReqNo_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress txtReqNo, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtReqNo_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtReqNo_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus txtReqNo, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtReqNo_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtReqNo_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus txtReqNo, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtReqNo_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub


Private Sub chkCostFromReq_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.onClick chkCostFromReq, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "chkCostFromReq_Click()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub chkCostFromReq_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus chkCostFromReq, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "chkCostFromReq_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub chkCostFromReq_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus chkCostFromReq, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "chkCostFromReq_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

#If CUSTOMIZER And CONTROLS Then

Private Sub CustomButton_Click(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.onClick CustomButton(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomButton_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomButton_GotFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomButton(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomButton_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomButton_LostFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomButton(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomButton_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCheck_Click(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.onClick CustomCheck(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCheck_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCheck_GotFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomCheck(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCheck_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCheck_LostFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomCheck(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCheck_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCombo_Change(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomCombo(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCombo_Change", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCombo_Click(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.onClick CustomCombo(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCombo_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCombo_DblClick(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomCombo(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCombo_DblClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCombo_GotFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomCombo(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCombo_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCombo_KeyPress(Index As Integer, KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomCombo(Index), KeyAscii
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCombo_KeyPress", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCombo_LostFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomCombo(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCombo_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCurrency_Change(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomCurrency(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCurrency_Change", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCurrency_GotFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomCurrency(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCurrency_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCurrency_KeyPress(Index As Integer, KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomCurrency(Index), KeyAscii
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCurrency_KeyPress", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomCurrency_LostFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomCurrency(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomCurrency_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomFrame_Click(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.onClick CustomFrame(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomFrame_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomFrame_DblClick(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomFrame(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomFrame_DblClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomLabel_Click(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.onClick CustomLabel(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomLabel_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomLabel_DblClick(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomLabel(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomLabel_DblClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomMask_Change(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomMask(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomMask_Change", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomMask_GotFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomMask(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomMask_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomMask_KeyPress(Index As Integer, KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomMask(Index), KeyAscii
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomMask_KeyPress", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomMask_LostFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomMask(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomMask_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
               Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomNumber_Change(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomNumber(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomNumber_Change", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomNumber_GotFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomNumber(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomNumber_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomNumber_KeyPress(Index As Integer, KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomNumber(Index), KeyAscii
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomNumber_KeyPress", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomNumber_LostFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomNumber(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomNumber_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomOption_Click(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.onClick CustomOption(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomOption_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomOption_DblClick(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomOption(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomOption_DblClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomOption_GotFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomOption(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomOption_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomOption_LostFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomOption(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomOption_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomSpin_DownClick(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnSpinDown CustomSpin(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomSpin_DownClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomSpin_UpClick(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnSpinUp CustomSpin(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomSpin_UpClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

#End If
#If CUSTOMIZER And CONTROLS Then

Private Sub CustomDate_Click(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.onClick CustomDate(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomDate_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomDate_DblClick(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomDate(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomDate_DblClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomDate_GotFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomDate(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomDate_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomDate_LostFocus(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomDate(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomDate_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomDate_KeyPress(Index As Integer, KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomDate(Index), KeyAscii
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomDate_KeyPress", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub CustomDate_Change(Index As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomDate(Index)
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CustomDate_Change", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

#End If









Public Property Get MyApp() As Object
    Set MyApp = App
End Property
Public Property Get MyForms() As Object
    Set MyForms = Forms
End Property






 Private Sub grdSelectPush_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    moSelectGM.Grid_LeaveCell Col, Row, NewCol, NewRow
End Sub
 Private Sub grdSelectPush_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    moSelectGM.Grid_TopLeftChange OldLeft, OldTop, NewLeft, NewTop
End Sub

Private Function FetchUpdCnt() As Boolean
    Dim sSql As String
    Dim rsPreFetch As Object
    
    If Not mcolUpdCounter Is Nothing Then
        Set mcolUpdCounter = Nothing
        Set mcolUpdCounter = New Collection
    End If
    
    sSql = "SELECT DocLineKey, rl.UpdateCounter FROM #tpoPushPull temp " & _
           "JOIN tpoReqLine rl WITH (NOLOCK) " & _
           "ON temp.DocLineKey = rl.ReqLineKey  "

    Set rsPreFetch = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
    Do While Not rsPreFetch.IsEOF
        mcolUpdCounter.Add CStr(rsPreFetch.Field("UpdateCounter")), Trim$(rsPreFetch.Field("DocLineKey"))
        rsPreFetch.MoveNext
    Loop
    
    Set rsPreFetch = Nothing
        
End Function

Private Function ValidateUpdCnt(ByVal lGrdMaxrows As Long) As Boolean
    Dim sSql As String
    Dim rsPostFetch As Object
    
    Dim lRow As Long
    Dim lDocLineKey As Long
    Dim lSelected As Long
    
    Dim lPrevUpdateCounter As Long
    Dim lCurUpdateCounter As Long
    
    ValidateUpdCnt = True
        
    For lRow = 1 To lGrdMaxrows
        lSelected = gsGridReadCell(grdSelectPush, lRow, kColSelect)
        
        If lSelected = 1 Then
        
            lDocLineKey = gsGridReadCell(grdSelectPush, lRow, kColDocLineKey)
            lPrevUpdateCounter = mcolUpdCounter.Item(Trim$(lDocLineKey))
            lCurUpdateCounter = 0
                    
            sSql = "SELECT rl.UpdateCounter FROM #tpoPushPull temp" & _
                   " JOIN tpoReqLine rl WITH (NOLOCK)" & _
                   " ON temp.DocLineKey = rl.ReqLineKey " & _
                   " WHERE temp.DocLineKey = " & lDocLineKey & _
                   " AND temp.Selected = 1"
                   
            Set rsPostFetch = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
            
            If Not rsPostFetch.IsEOF Then
                lCurUpdateCounter = rsPostFetch.Field("UpdateCounter")
            End If
                    
            If lPrevUpdateCounter <> lCurUpdateCounter Then
                giSotaMsgBox Me, moClass.moSysSession, kmsgDMNoSaveConcurrency
                If Not rsPostFetch Is Nothing Then rsPostFetch.Close: Set rsPostFetch = Nothing
                ValidateUpdCnt = False
                Exit For
            End If
                
            If Not rsPostFetch Is Nothing Then rsPostFetch.Close: Set rsPostFetch = Nothing
        End If
    Next lRow
    
End Function

'Agregado por Multiconsulting
Private Function bIsValid() As Boolean
    Dim lContractKey As Long
    Dim lType As Long
    Dim lVendKey As Long
    Dim i As Long
    Dim lTemp As Long
    
    Dim bValid As Boolean
    
    On Error GoTo ErrorHandler
    bIsValid = True
    For i = 0 To grdSelectPush.DataRowCnt
        If giGetValidInt(gsGridReadCell(grdSelectPush, i, kColSelect)) = 1 Then
            If lType = 0 Then
                lType = moClass.moAppDB.Lookup("p.Type", "tpoReqAdicInfo as p join tpoReqLine as s on s.ReqKey = p.ReqKeyIA", "s.ReqLineKey =" & glGetValidLong(gsGridReadCell(grdSelectPush, i, kColDocLineKey)))
                lVendKey = glGetValidLong(moClass.moAppDB.Lookup("s.VendKey", "tpoReqLine as s", "s.ReqLineKey =" & glGetValidLong(gsGridReadCell(grdSelectPush, i, kColDocLineKey))))
            Else
                lTemp = moClass.moAppDB.Lookup("p.Type", "tpoReqAdicInfo as p join tpoReqLine as s on s.ReqKey = p.ReqKeyIA", "s.ReqLineKey =" & glGetValidLong(gsGridReadCell(grdSelectPush, i, kColDocLineKey)))
                If lTemp <> lType Then
                    MsgBox "Debe seleccionar líneas que pertenescan a Requisiciones del mismo tipo", vbExclamation, "Alerta"
                    bIsValid = False
                    Exit Function
                End If
                lTemp = glGetValidLong(moClass.moAppDB.Lookup("s.VendKey", "tpoReqLine as s", "s.ReqLineKey =" & glGetValidLong(gsGridReadCell(grdSelectPush, i, kColDocLineKey))))
                If lTemp <> lVendKey Then
                    MsgBox "Debe seleccionar líneas que pertenescan al mismo Proveedor", vbExclamation, "Alerta"
                    bIsValid = False
                    Exit Function
                End If
            End If
        End If
    Next i
    
    If lType = 1 Then
        If Not bValidContract Then
            bIsValid = False
        End If
    Else
        lContractKey = 0
        frmContractAssociate.lVendKey = lVendKey
        Set frmContractAssociate.oclass = moClass
        frmContractAssociate.lkuContract.Enabled = True
        frmContractAssociate.lkuSuplement.Enabled = True
        
        frmContractAssociate.cmdOK.Enabled = True
        
        
        frmContractAssociate.ShowContract lContractKey, bValid
        If bValid And lContractKey > 0 Then
            lContractKeyUsed = lContractKey
            If Not ValidateContractLines Then
                bIsValid = False
            End If
        Else
            MsgBox "Debe seleccionar un Contrato válido", vbExclamation, "Alerta"
            bIsValid = False
            Exit Function
        End If
    End If
    Exit Function
ErrorHandler:
    MsgBox "bIsValid_" & Err.Description, vbCritical, "Sage MAS 500"
    bIsValid = False
End Function

Private Function bValidContract()
    Dim lContractKey As Long
    Dim iRow As Long
    Dim bRetValue As Boolean
    
    bValidContract = True
    bRetValue = True
    If grdSelectPush.DataRowCnt = 0 Then Exit Function
    
    
    For iRow = 1 To grdSelectPush.DataRowCnt
        If gsGridReadCell(grdSelectPush, iRow, kColSelect) = 1 Then
            If lContractKey = 0 Then
                lContractKey = glGetValidLong(moClass.moAppDB.Lookup("p.ContractKey", "tpoRequisitionContract AS p JOIN tpoReqLine AS s ON s.ReqKey = p.ReqKey", "s.ReqLineKey =" & glGetValidLong(gsGridReadCell(grdSelectPush, iRow, kColDocLineKey))))
            Else
                If lContractKey <> glGetValidLong(moClass.moAppDB.Lookup("p.ContractKey", "tpoRequisitionContract AS p JOIN tpoReqLine AS s ON s.ReqKey = p.ReqKey", "s.ReqLineKey =" & glGetValidLong(gsGridReadCell(grdSelectPush, iRow, kColDocLineKey)))) Then
                    bRetValue = False
                End If
            End If
        End If
    Next
    
    If Not bRetValue Then
        MsgBox "Debe selecionar partidas del mismo contrato", vbInformation, "Alerta"
        bValidContract = False
    End If
    
    lContractKeyUsed = lContractKey
End Function

Private Function ValidateContractLines() As Boolean
Dim i As Long
Dim lItemKey As Long
Dim dQty As Double
    ValidateContractLines = True
    For i = 0 To grdSelectPush.DataRowCnt
        If giGetValidInt(gsGridReadCell(grdSelectPush, i, kColSelect)) = 1 Then
            lItemKey = glGetValidLong(moClass.moAppDB.Lookup("ItemKey", "tpoReqLine", "ReqLineKey =" & gsGridReadCell(grdSelectPush, i, kColDocLineKey)))
            If Not bIsItemInContract(lContractKeyUsed, lItemKey) Then
                MsgBox "El artículo " & gsGridReadCell(grdSelectPush, i, kcolItemID) & " no pertenece al contrato seleccionado", vbExclamation, "Alerta"
                ValidateContractLines = False
                Exit Function
            End If
            
            dQty = dGetMaxQtyAllowAtContract(lContractKeyUsed, lItemKey)
            If dQty < gsGridReadCell(grdSelectPush, i, kColQty) Then
                MsgBox "El artículo " & gsGridReadCell(grdSelectPush, i, kcolItemID) & " sobrepasa la cantidad pendiente en el contrato seleccionado", vbExclamation, "Alerta"
                ValidateContractLines = False
                Exit Function
            End If
        End If
    Next i
End Function

Private Function bIsItemInContract(lContractKey As Long, lItemKey As Long) As Boolean
    Dim sSql As String
    Dim rs As Object
    Dim bFirst As Boolean
    Dim sRetVal As String
    Dim lParentKey As Long
    
    On Error GoTo ErrorHandler
    
    bIsItemInContract = False
    
    lParentKey = glGetValidLong(moClass.moAppDB.Lookup("ParentContractKey", "tctContract", "ContractKey=" & lContractKey))
    
    sSql = "SELECT p.ItemKey FROM tctContractLine AS p WHERE p.ContractKey = " & lContractKey
    If lParentKey <> 0 Then
        sSql = sSql & " or p.ContractKey =" & lParentKey
    End If
    
    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    If rs.IsEmpty Then Exit Function
    
    While Not rs.IsEOF
        If glGetValidLong(rs.Field("ItemKey")) = lItemKey Then
            Set rs = Nothing
            bIsItemInContract = True
            Exit Function
        End If
        rs.MoveNext
    Wend
    
    Set rs = Nothing
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error"
End Function

Private Function dGetMaxQtyAllowAtContract(lContractKey As Long, lItemKey As Long) As Double
    Dim dMaxQty As Double
    Dim dActualUsed As Double
    Dim i As Long
    Dim lParentKey As Long
    On Error GoTo ErrorHandler
    dGetMaxQtyAllowAtContract = 0

    lParentKey = glGetValidLong(moClass.moAppDB.Lookup("ParentContractKey", "tctContract", "ContractKey=" & lContractKey))

    If lParentKey <> 0 Then
        dMaxQty = gdGetValidDbl(moClass.moAppDB.Lookup("p.Qty", "tctContractLine AS p join tctContract as s on p.ContractKey = s.ContractKey", "(p.ContractKey = " & lContractKey & " or p.ContractKey =" & lParentKey & ") AND p.ItemKey = " & lItemKey & " order by s.StartDate desc"))
        dActualUsed = gdGetValidDbl(moClass.moAppDB.Lookup("sum(s.QtyReq)", "tpoReqLine AS p JOIN tpoReqLineDist AS s ON s.ReqLineKey = p.ReqLineKey JOIN tpoRequisitionContract AS t ON t.ReqKey = p.ReqKey", "(t.ContractKey = " & lContractKey & " or t.ContractKey = " & lParentKey & ") AND p.ItemKey = " & lItemKey))
        dActualUsed = dActualUsed + gdGetValidDbl(moClass.moAppDB.Lookup("SUM(s.QtyOrd)", "tpoPOLine AS p JOIN tpoPOLineDist AS s ON s.POLineKey = p.POLineKey JOIN tpoPurchOrder AS t ON p.POKey = t.POKey  and t.Status <> 3", "(t.ContractKey = " & lContractKey & " or t.ContractKey = " & lParentKey & ") AND p.ItemKey = " & lItemKey))
    Else
        dMaxQty = gdGetValidDbl(moClass.moAppDB.Lookup("p.Qty", "tctContractLine AS p", "p.ContractKey = " & lContractKey & " AND p.ItemKey = " & lItemKey))
        dActualUsed = gdGetValidDbl(moClass.moAppDB.Lookup("sum(s.QtyReq)", "tpoReqLine AS p JOIN tpoReqLineDist AS s ON s.ReqLineKey = p.ReqLineKey JOIN tpoRequisitionContract AS t ON t.ReqKey = p.ReqKey", "t.ContractKey = " & lContractKey & " AND p.ItemKey = " & lItemKey))
        dActualUsed = dActualUsed + gdGetValidDbl(moClass.moAppDB.Lookup("SUM(s.QtyOrd)", "tpoPOLine AS p JOIN tpoPOLineDist AS s ON s.POLineKey = p.POLineKey JOIN tpoPurchOrder AS t ON p.POKey = t.POKey  and t.Status <> 3", "t.ContractKey = " & lContractKey & " AND p.ItemKey = " & lItemKey))
    End If
    
    If dMaxQty < dActualUsed Then
        dGetMaxQtyAllowAtContract = 0
    Else
        dGetMaxQtyAllowAtContract = dMaxQty - dActualUsed
    End If
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

