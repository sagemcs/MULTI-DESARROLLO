VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPushPull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private moClass As Object

Private mbLoadImmediate As Boolean

Private mgSelectGrid As fpSpread


' Positioning constants
Const kGridHeightOffset = 585
Const kGridWidthOffset = 75
Const kButtonTopOffset = 480

' Column fields
Const kSelectCol = 1
Const kUserCol01 = 2
Const kUserCol02 = 3
Const kUserCol03 = 4
Const kUserCol04 = 5
Const kUserCol05 = 6
Const kUserCol06 = 7
Const kUserCol07 = 8
Const kUserCol08 = 9
Const kUserCol09 = 10
Const kUserCol10 = 11
Const kUserCol11 = 12
Const kUserCol12 = 13
Const kUserCol13 = 14
Const kUserCol14 = 15
Const kUserCol15 = 16
Const kUserCol16 = 17
Const kUserCol17 = 18
Const kUserCol18 = 19
Const kUserCol19 = 20
Const kUserCol20 = 21

'Event Declarations:
Event FormatGrid()
Event BindGrid()
Event LoadTempTable()
Event ProcessProceed()
Event CreateMoreTempTables()

Public Property Get hWnd() As Long
    hWnd = mgSelectGrid.hWnd
End Property

Property Get moPushPullGrid() As Object
'    Set moPushPullGrid = grdSelectLines
End Property

Public Sub Init()

    If mbLoadImmediate Then
        RaiseEvent LoadTempTable
    End If
    RaiseEvent BindGrid
End Sub
Public Sub InitControl(oclass As Object)
    Set moClass = oclass
    RaiseEvent FormatGrid
    CreateTempTable

End Sub

Public Sub Terminate()
    DropTempTable
    
End Sub
Private Sub DropTempTable()
    Dim sSql As String
    
    sSql = "DROP TABLE #tpoPushPull"
        
    On Error Resume Next
    moClass.moAppDB.ExecuteSQL sSql

End Sub


Private Sub CreateTempTable()

    RaiseEvent CreateMoreTempTables
    
End Sub
Public Function lGetSessionID() As Long
    Dim lKey As Long
    
    With moClass.moAppDB
        .SetInParam "tciErrorLog"
        .SetOutParam lKey
        On Error Resume Next
        .ExecuteSP ("spGetNextSurrogateKey")
        
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbOK, "Sage 500 ERP"
            .ReleaseParams
            Exit Function
        End If
        lKey = .GetOutParamLong(2)
        .ReleaseParams
    End With
    

    lGetSessionID = lKey

End Function


Public Sub ProceedPressed()

' This will save the grid data back to the temp table

    RaiseEvent ProcessProceed

End Sub


Public Property Get LoadImmediate() As Variant
    LoadImmediate = mbLoadImmediate
End Property

Public Property Let LoadImmediate(ByVal vNewValue As Variant)
    mbLoadImmediate = vNewValue
End Property





