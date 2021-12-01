VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCTZDA001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public moFramework          As Object

Public mlContext            As Long

Public moDasSession         As Object

Public moAppDB              As Object

Public moSysSession         As Object

Public miShutDownRequester  As Integer

Public mlError              As Long

Private mlRunFlags          As Long

Private mlUIActive          As Long

Public mfrmMain            As Form

Const VBRIG_MODULE_ID_STRING = "CTZDA001.CLS"
Public Function iMyModule() As Integer
    iMyModule = 61
End Function


Private Function sMyName() As String
'+++ VB/Rig Skip +++
    sMyName = "clsCTZDA001"
End Function

Private Sub Class_Initialize()
    mlUIActive = kChildObjectInactive
    miShutDownRequester = kFrameworkShutDown
End Sub

Public Function InitializeObject(ByVal oFramework As Object, ByVal lContext As Long) As Long
    InitializeObject = kFailure
    DefaultInitializeObject Me, oFramework, lContext, App.ProductName, App.Title
    InitializeObject = kSuccess     ' return success
End Function

Public Function LoadUI(ByVal lContext As Long) As Long
    LoadUI = kFailure

    Set mfrmMain = frmContract

    Set mfrmMain.oClass = Me
    mfrmMain.lRunMode = mlContext And kRunModeMask
    
    Load frmContract
    If mlError Then Err.Raise mlError
    
    'determine if the form load was successful
    If mfrmMain.bLoadSuccess Then
        LoadUI = kSuccess
    Else
        LoadUI = kFailure
    End If

End Function

Public Function GetUIHandle(ByVal lContext As Long) As Long
    If Not mfrmMain Is Nothing Then
        GetUIHandle = mfrmMain.hWnd
    End If
End Function

Public Function DisplayUI(ByVal lContext As Long) As Long
#If InProc = 0 Then
    DisplayUI = kFailure
    If mfrmMain Is Nothing Then Exit Function

    mfrmMain.Show
    DoEvents
    If mlError Then Err.Raise mlError
    
    mfrmMain.SetFocus
    DisplayUI = kSuccess
#Else
    DisplayUI = EFW_CT_MODALEXIT
#End If
End Function

Public Function ShowUI(ByVal lContext As Long) As Long
#If InProc = 0 Then
    ShowUI = kFailure

    If Not mfrmMain Is Nothing Then
        mfrmMain.Show
        DoEvents
    End If
#End If

    ShowUI = kSuccess
End Function

Public Function MinimizeUI(ByVal lContext As Long) As Long
    MinimizeUI = kFailure

    If Not mfrmMain Is Nothing Then
#If InProc = 0 Then
        If gbActiveChildObjects(mfrmMain.moSotaObjects) Then Exit Function
#End If
        mfrmMain.WindowState = vbMinimized
    End If

    MinimizeUI = kSuccess
End Function

Public Function RestoreUI(ByVal lContext As Long) As Long
    RestoreUI = kFailure

    If Not mfrmMain Is Nothing Then
        mfrmMain.WindowState = vbNormal
    End If

    RestoreUI = kSuccess
End Function

Public Function HideUI(ByVal lContext As Long) As Long
#If InProc = 0 Then
    HideUI = kFailure

    If Not mfrmMain Is Nothing Then
        If gbActiveChildObjects(mfrmMain.moSotaObjects) Then Exit Function
        mfrmMain.Hide
    End If
#End If

    HideUI = kSuccess
End Function

Public Function QueryShutDown(lContext As Long) As Long
    QueryShutDown = kFailure

    If Not mfrmMain Is Nothing Then
        If gbActiveChildObjects(mfrmMain.moSotaObjects) Then Exit Function
    End If

    QueryShutDown = kSuccess
End Function

Public Function UnloadUI(ByVal lContext As Long) As Long
    UnloadUI = kFailure

    If Not mfrmMain Is Nothing Then
        If miShutDownRequester = kFrameworkShutDown Then
            Unload mfrmMain
            If mfrmMain.bCancelShutDown Then Exit Function
            Set mfrmMain = Nothing
        Else
          If mlError Then mfrmMain.PerformCleanShutDown
        End If
    End If

    UnloadUI = kSuccess
End Function

Public Function TerminateObject(ByVal lContext As Long) As Long

    TerminateObject = kFailure

    DefaultTerminateObject Me
    
    Set mfrmMain = Nothing
    
    TerminateObject = kSuccess
End Function

Public Property Get lUIActive() As Long
    lUIActive = mlUIActive
End Property
Public Property Let lUIActive(lNewActive As Long)
    mlUIActive = lNewActive
End Property

Public Property Get lRunFlags() As Long
    lRunFlags = mlRunFlags
End Property
Public Property Let lRunFlags(ltRunFlags As Long)
    mlRunFlags = ltRunFlags
End Property

Public Sub NavigateTo(ByVal xml As String)
'+++ VB/Rig Skip +++
'*******************************************************************************
'    Desc: Exposed method for object's navigation
'
'   Parms: xml - XML string containing context.
'
' Returns: N/A
'*******************************************************************************
On Error Resume Next

    DrillAround GetFromXML(xml, "{Element Name Goes Here}")
    ' i.e. DrillAround GetFromXML(xml, "GLAcctNo")
    
    ' Exit and clear error buffer
    Err.Clear
    
End Sub

Public Sub DrillAround(sID As String)
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
'*******************************************************************************
'    Desc: DrillAround will work like the Maintenance form is being
'          invoked from the launch pad, with the exception that an ID
'          may be passed in.
'   Parms: sID = ID to be Displayed/Viewed
' Returns: N/A
'*******************************************************************************
On Error GoTo ExpectedErrorRoutine

  If mfrmMain Is Nothing Then
      If moFramework.LoadUI(Me) = kFailure Then GoTo ShutMeDown
      If moFramework.DisplayUI(Me) = kFailure Then GoTo ShutMeDown
  End If
     
  If Len(Trim(sID)) > 0 Then
    ' {Form specific logic goes here}
    ' i.e. mfrmMain.glaGLAcct = sID
    ' i.e. mfrmMain.VMIsValidKey
  End If
  
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
  Exit Sub
  
ShutMeDown:
  miShutDownRequester = kUnloadSelfShutDown
  moFramework.UnloadSelf EFW_TF_MANSHUTDN
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
  Exit Sub
  
ExpectedErrorRoutine:

'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DrillAround", VBRIG_IS_CLASS
        Err.Raise guSotaErr.Number
'+++ VB/Rig End +++
End Sub

