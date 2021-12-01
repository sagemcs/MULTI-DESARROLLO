VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsrequistn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************
'     Name: clsrequistn
'     Desc: This class handles the user interface for requisitions.
' Original: 03-05-1998
'     Mods: mm/dd/yy XXX
'**********************************************************************
Option Explicit
    '=======================================================================
    '           Public Variables for use by other modules
    '           -----------------------------------------
    '
    '   moFramework     - required to give other modules the ablitiy to
    '                     call the framework methods and properties, such
    '                     as UnloadSelf, LoadSotaObject, . . .
    '                     Object reference is set in InitializeObject.
    '
    '   mlContext       - Contains information of how this class object
    '                     was invoked.  It is set in InitializeObject
    '
    '
    '   moSysSession    - system Manager Session Object which contains
    '                     information specific to the user's session with
    '                     this accounting application.  For example, current
    '                     CompanyID, Business date, UserID, UserName, . . .
    '
    '   miShutDownRequester
    '                   - The ShutDownRequester is the object responsible
    '                     for requesting this object to shut down.  There are
    '                     two possible choices: 1) the Framework requests the
    '                     object to shut itself down (kFrameworkShutDown; and
    '                     2) the User Interface or UnloadSelf call requests
    '                     the framework to shut this object down (kUnloadSelfShutDown).
    '                     The main purpose of this flag is to
    '                     ensure that the framework reference isn't removed (or
    '                     set to nothing) before the UI.
    '=======================================================================

    Public moFramework          As Object
    Public mlContext            As Long
    Public moDasSession         As Object
    Public moAppDB              As Object
    Public moSysSession         As Object
    Public miShutDownRequester  As Integer
    Public mlError              As Long
    Public mbClassShutdown      As Boolean

    '=======================================================================
    'Private Variables for Class Properties
    '=======================================================================
    Private mlRunFlags          As Long
    Private mlUIActive          As Long

    '=======================================================================
    'Private variables for use within this class only
    '=======================================================================
 '   Private mfrmmain            As Form

Const VBRIG_MODULE_ID_STRING = "REQUISTN.CLS"

Private Function sMyName() As String
'+++ VB/Rig Skip +++
    sMyName = "clspurchord"
End Function

Public Property Get lUIActive() As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'*************************************************************************
' Description:
'    lUIActive Property Get informs the parent object of its Active state.
'    A class is Active if it is in the middle of processing a record.
'    This property is **Required** for all Sage MAS 500 Interface 3 classes.
'*************************************************************************
    lUIActive = mlUIActive
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "lUIActive_Get", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Property

Public Property Let lUIActive(lNewActive As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'******//*****************************************************************
' Description:
'    lUIActive Property Let is contains the new Active state of the class.
'    This is set internally within the form or method within this class.
'    This property is **Required** for all Sage MAS 500 Interface 3 classes.
'*************************************************************************
    mlUIActive = lNewActive
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "lUIActive_Let", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Property

Public Property Get lRunFlags() As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'**************************************************************************
' Description:
'    lRunFlags Property Get can be used to inform other objects or modules.
'    of how this object was invoked by the framework or parent object.
'    This property is **Required** for all Sage MAS 500 Interface 3 classes.
'**************************************************************************
    lRunFlags = mlRunFlags
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "lRunFlags_Get", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Property

Public Property Let lRunFlags(ltRunFlags As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'*******************************************************************
' Description:
'    lRunFlags Property Let is set by the Framework no matter which
'    interface Type is being used.  This property is set before any
'    other method call of this class.
'    This property is **Required** for all Sage MAS 500 Interface 3 classes.
'*******************************************************************
    mlRunFlags = ltRunFlags
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "lRunFlags_Let", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Property

Private Sub Class_Initialize()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'************************************************************************
' Description:
'    Class_Initialize re-dimensions the Hide and Minimize Form Arrays
'    to 0.  The purpose of this is to ensure that the UBound function
'    doesn't error if the ShowUI or RestoreUI are called out of order.
'
'    Class_Initialize also sets the ShutDownRequester variable to say
'    that if anything goes wrong initially within this class, that the
'    framework will automatically attempt to shut down this class object.
'************************************************************************
    mlUIActive = kChildObjectInactive
    miShutDownRequester = kFrameworkShutDown
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "Class_Initialize", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Sub

Public Function InitializeObject(ByVal oFramework As Object, ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'***********************************************************************
'     Description:
'       InitalizeObject performs setup required for all applications.
'
'       InitalizeObject sets the context in which this class was invoked.
'       The context contains information from both the framework and
'       application on how this class was invoked.
'       The available user-defined application contexts that can be derived
'       from the lContext argument for clsWarehouseBinMnt are:
'           1) normal           (kContextNormal)
'           2) Add-on-the-fly   (kContextAOF)
'           3) Drill Around     (kContextDA)
'
'       InitalizeObject also sets the Non-UI objects required by all
'       applications.  These Non-UI objects are:
'           1) system Manager Session Object, which provides our applications
'              with user specific session information such as the
'              current CompanyID, Business date, etc and more importantly
'              Application data source Name, and system data source Name.
'           2) DAS Session Object, which supplies an instance of
'              the Sage MAS 500 DAS so our projects can use the methods/Services
'              required for data Access.
'           3) Once we have the DAS Session Object, Application DSN, and
'              system DSN we can open the Application and system Databases,
'              required for data manipulation.
'
'       InitializeObject is called for Sage MAS 500 Inteferface 1, Sage MAS 500 Interface 2
'       and Sage MAS 500 Interface 3.
'
'    Parameters:
'       oFramework <in> - reference to the framework event object which
'                         gives this object access to the public
'                         framework properties and methods.
'       lContext <in>   - The context contains information from both the
'                         framework and application on how this class was
'                         invoked.
'
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     In case of kFailure, the Framework will execute
'                     the TerminateObject method of this class and shut
'                     down this object.
'***********************************************************************
    InitializeObject = kFailure

    DefaultInitializeObject Me, oFramework, lContext, App.ProductName, App.Title

    InitializeObject = kSuccess     ' return success
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "InitializeObject", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function LoadUI(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'***********************************************************************
'     Description:
'       LoadUI is responsible for loading the main form(s) required for
'       user interface of this class.  LoadUI performs the following
'       tasks:
'           1) Instantiate (or set a pointer to) the main form(s) of
'              the class.
'           2) In order for the form to use the class' public
'              variables, a reference back to the class is required.
'              This is done using the form's oClass Property Set and setting
'              the object property to Me (this class instance).
'           3) For the application to use the form in a variety of ways,
'              (Maintenance, Add-on-the-Fly, Drill Around), the Form needs
'              to know the user-defined context in which it was invoked.
'              The form contains a property to Hold this Value which it
'              refers to as the Run Mode.  The Run Mode can
'              be extracted by performing a bit-wise comparison of the
'              context originally passed in through the InitializeObject method.
'
'       LoadUI is called automatically by the framework for SotaInterface 2,3 and
'       the AutoLoadUI run flag is set.
'
'    Parameters:
'           lContext <in>   - Context of the LoadUI or how it is being
'                             invoked by the framework
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     In case of kFailure, the Framework will proceed
'                     to shut down this object starting with the
'                     QueryShutDown method, proceeding to UnloadUI and
'                     TerminateObject.
'***********************************************************************
    LoadUI = kFailure

    Set mfrmMain = frmRequistn

    Set mfrmMain.oclass = Me
'    mfrmMain.lRunMode = mlContext And kRunModeMask
    
    Load mfrmMain
    
    If mlError Then Err.Raise mlError
    
    If mfrmMain.mbCancelLoad Then
        LoadUI = EFW_C_CANCEL
    Else
        LoadUI = kSuccess
    End If

'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "LoadUI", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function GetUIHandle(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not mfrmMain Is Nothing Then
        GetUIHandle = mfrmMain.hWnd
    End If
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "GetUIHandle", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function DisplayUI(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'*********************************************************************
'   Description:
'       The framework will call this method when the class object is
'       initially invoked and the RunFlags are are set to AutoDisplayUI.
'
'       DisplayUI is called automatically by the framework for SotaInterface 2,3 and
'       the AutoDisplayUI run flag is set.
'
'    Parameters:
'           lContext <in>   - Context of the DisplayUI or how it is being
'                             invoked by the framework.
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     In case of kFailure, the Framework will proceed
'                     to shut down this object starting with the
'                     QueryShutDown method, proceeding to UnloadUI and
'                     TerminateObject.
'*********************************************************************
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
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "DisplayUI", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function ShowUI(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'*********************************************************************
'   Description:
'       The framework will call this method when
'       the UI of this object had been hidden in an
'       explicit HideUI because the framework was minimized.
'       When the framework is restored, it requests the UI
'       of its objects to show themselves.
'
'       ShowUI is called automatically by the framework for SotaInterface 2,3 and
'
'    Parameters:
'           lContext <in>   - Context of the ShowUI or how it is being
'                             invoked by the framework.
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     Currently the framework ignores a failure.
'*********************************************************************
#If InProc = 0 Then
    ShowUI = kFailure

    If Not mfrmMain Is Nothing Then
        mfrmMain.Show
        DoEvents
    End If
#End If

    ShowUI = kSuccess
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "ShowUI", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function MinimizeUI(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'*********************************************************************
'   Description:
'       MinimizeUI minimizes the object's main form and hides all other forms
'       launched directly from the main form or from the class.
'       Forms already minimized or already hidden are ignored.
'       Forms launched from loaded objects are also ignored.
'
'
'    Parameters:
'           lContext <in>   - Context of the MinimizeUI or how it is being
'                             invoked by the framework.
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     The framework currently ignores a failure.
'*********************************************************************
    MinimizeUI = kFailure

    If Not mfrmMain Is Nothing Then
#If InProc = 0 Then
        If gbActiveChildObjects(mfrmMain.moSotaObjects) Then Exit Function
#End If
        mfrmMain.WindowState = vbMinimized
    End If

    MinimizeUI = kSuccess
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "MinimizeUI", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function RestoreUI(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'*********************************************************************
'   Description:
'       Calls the general procedure which restores the object's main
'       form and all other forms explicitly minimized in a previous
'       MinimizeUI.
'
'    Parameters:
'           lContext <in>   - Context of the RestoreUI or how it is being
'                             invoked by the framework.
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     The framework currently ignores a failure.
'*********************************************************************
    RestoreUI = kFailure

    If Not mfrmMain Is Nothing Then
        mfrmMain.WindowState = vbNormal
    End If

    RestoreUI = kSuccess
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "RestoreUI", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function HideUI(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'*********************************************************************
'   Description:
'       If any child objects are Active, then this form is the parent
'       and it cannot hide itself.
'
'    Parameters:
'           lContext <in>   - Context of the HIdeUI or how it is being
'                             invoked by the framework.
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     The framework currently ignores a failure.
'*********************************************************************
#If InProc = 0 Then
    HideUI = kFailure

    If Not mfrmMain Is Nothing Then
        If gbActiveChildObjects(mfrmMain.moSotaObjects) Then Exit Function
        mfrmMain.Hide
    End If
#End If

    HideUI = kSuccess
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "HideUI", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

'***********************************************************************
'                   Class Object ShutDown Process
'                   -----------------------------
'
'   Sage MAS 500 Interface 3:
'       If InitializeObject returns kFailure, then the TerminateObject
'       method is called immediately and this object is guaranteed to
'       shutdown whether or not the TerminateObject returned kSuccess
'       or kFailure.
'
'       If LoadUI or DisplayUI returned kFailure, the QueryShutDown method is
'       called, followed by UnloadUI and TerminateObject.  The framework will
'       ignore failures from the QueryShutDown, UnloadUI, and Terminate object,
'       because the object is in a failure.
'
'       If the Framework requests shutdown or the UnloadSelf call requests
'       shutdown, then the QueryShutDown method is invoked.
'       If a successful QueryShutDown is performed, the UnloadUI method is
'       called. And if a successful UnloadUI is performed, TerminateObject
'       method is called.  Once the TerminateObject method is called, the
'       object will be shut down whether or not the TerminateObject returned
'       success or failure.
'
'       Special NOTE:
'           The QueryShutDown is designed to check for Active child objects
'           and return failure to the framework if any have been found.  If
'           there are any Active Child Objects remaining after a successful
'           QueryShutDown, the framework will shut them down prior to the
'           UnloadUI method call.
'
'           The UnloadUI is designed to unload its UI and within that it will unload
'           all its child objects (Sage MAS 500 children).  If there are
'           any Sage MAS 500 child objects remaining after a successful UnloadUI,
'           the framework will shut them down prior to the TerminateObject
'           method call.
'
'           The TerminateObject is designed to unload/remove all Non-UI
'           object references.  If there are any child Non-UI (No Sage MAS 500
'           Interface or Sage MAS 500 Interface 1 (?)) remaining after the
'           TerminateObject call, then the framework will remove or
'           shut them down.
'***********************************************************************
Public Function QueryShutDown(lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'***********************************************************************
'   Description:
'       QueryShutDown method to determine if this form can be
'       shutdown.  This is done by checking all child objects (created
'       via a LoadSotaObject) for the lUIActive property set to
'       kChildObjectActive.
'
'       QueryShutDown is called automatically by the framework for
'       Sage MAS 500 Interface 1, 2, 3.
'
'    Parameters:
'           lContext <in>   - Context of the QueryShutDown or how it is being
'                             invoked by the framework.
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'                     Proceed with the ShutDown
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     Do not proceed with the ShutDown because there is
'                     an Active child object.
'***********************************************************************
    QueryShutDown = kFailure

    If Not mfrmMain Is Nothing Then
        If gbActiveChildObjects(mfrmMain.moSotaObjects) Then Exit Function
    End If

    QueryShutDown = kSuccess
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "QueryShutDown", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function UnloadUI(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'***********************************************************************
'   Description:
'       UnloadUI method attempts to Unload the UI or any forms loaded
'       within the life of this class object.  During the Unload of
'       the form, all Sage MAS 500 child Objects (Sage MAS 500 Interface 2,3) must
'       be unloaded.  This is currently coded in the form's Query_Unload
'       event.
'
'       UnLoadUI is called automatically by the framework for SotaInterface 2,3
'
'    Parameters:
'           lContext <in>   - Context of the UnloadUI or how it is being
'                             invoked by the framework.
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'                     Proceed with the ShutDown
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     Do not proceed with the ShutDown because there is
'                     a problem or the user has cancelled the framwork
'                     shutdown request.  This can be done if a dirty
'                     record is up an the form asks to save changes.
'                     If the user presses the cancel button, then the
'                     user is requesting the entire shutdown of the
'                     application to cancel its shutdown.
'***********************************************************************
    UnloadUI = kFailure

    If Not mfrmMain Is Nothing Then
        mbClassShutdown = True
        If miShutDownRequester = kFrameworkShutDown Then
        ' First close the Select Req Lines
            If mfrmMain.bSelectFormLoaded Then
'                frmSelectReqLines.Hide
                Unload frmSelectReqLines
             End If
            Unload mfrmMain
            If mfrmMain.bCancelShutDown Then Exit Function
        End If
    End If

    UnloadUI = kSuccess
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "UnloadUI", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function

Public Function TerminateObject(ByVal lContext As Long) As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'***********************************************************************
'   Description:
'       TerminateObject method removes all Non-UI object references
'       created in through the InitializeObject method.  If the framework
'       requested to shutdown of this object, then setting the framework
'       reference to Nothing is executed.  Otherwise, the form requested
'       an UnloadSelf with its QueryUnload.  After this TerminatObject
'       method has finished executing, control is relinquished
'       back to the that form event.  At that point, it is okay to set
'       the framework reference to nothing.
'
'    Parameters:
'           lContext <in>   - Context of the TerminateObject or how it is being
'                             invoked by the framework.
'   Return Values:
'       kSuccess    - if the function executed its code without error.
'                     Proceed with the ShutDown
'       kFailure    - if any of the code executed in this
'                     function produced an error or invalid result.
'                     Failure is ignored and the Object is guaranteed
'                     to shutdown at this point.
'***********************************************************************
    TerminateObject = kFailure

    DefaultTerminateObject Me
    
    Set mfrmMain = Nothing
    
    TerminateObject = kSuccess
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "TerminateObject", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Function
Public Sub DrillAround(ByVal sReqID As String)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigLogErrorRoutine
#End If
'+++ VB/Rig End +++
'**************************************************************************
'   Desc: DrillDown returns True if the Template ID was properly added into
'         the Template table.  Otherwise, it returns false.
'
'  Parms: sTemplateID - Key Value of the new Template Record.
'**************************************************************************
    If mfrmMain Is Nothing Then
        If moFramework.LoadUI(Me) = kFailure Then GoTo ShutMeDown
      If moFramework.DisplayUI(Me) = kFailure Then GoTo ShutMeDown
    End If
    
    sReqID = Trim(sReqID)
    If Len(sReqID) > 0 Then
        mfrmMain.lkuMain = sReqID
        mfrmMain.bIsValidReqNum
    End If
    
    
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub

ShutMeDown:
    miShutDownRequester = kUnloadSelfShutDown
    moFramework.UnloadSelf EFW_TF_MANSHUTDN

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigLogErrorRoutine:
        gMainClassErr Err, sMyName, "DrillAround", VBRIG_IS_CLASS
        Resume Next
'+++ VB/Rig End +++
End Sub

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

    Dim lReqKey     As Long
    Dim sTranNo     As String
    
    ' Init
    sTranNo = ""
    
    ' Get the ReqKey in order to get the TranNo
    lReqKey = glGetValidLong(GetFromXML(xml, "ReqKey"))
    
    ' Perform lookup
    If (lReqKey <> 0) Then
        sTranNo = gsGetValidStr(moAppDB.Lookup("TranNo", "tpoRequisition", "ReqKey = " & CStr(lReqKey)))
    End If
    
    DrillAround sTranNo
    
    ' Exit and clear error buffer
    Err.Clear
    
End Sub


