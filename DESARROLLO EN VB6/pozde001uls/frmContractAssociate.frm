VERSION 5.00
Object = "{2A076741-D7C1-44B1-A4CB-E9307B154D7C}#185.0#0"; "EntryLookupControls.ocx"
Begin VB.Form frmContractAssociate 
   Caption         =   "Asociar Contrato"
   ClientHeight    =   2145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContractAssociate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   990
   End
   Begin EntryLookupControls.TextLookup lkuSuplement 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
      ForeColor       =   -2147483640
      LookupID        =   "Contract"
      Datatype        =   0
      sSQLReturnCols  =   "ContractKey,,;ContractID,,;ContractNo,,;"
   End
   Begin EntryLookupControls.TextLookup lkuContract 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
      ForeColor       =   -2147483640
      LookupID        =   "Contract"
      Datatype        =   0
      sSQLReturnCols  =   "ContractKey,,;ContractID,,;ContractNo,,;"
   End
   Begin VB.Label lblSuplement 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suplemento"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   840
   End
   Begin VB.Label lblContract 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "frmContractAssociate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlContractKey As Long
Private mbValid As Boolean

Public lVendKey As Long
Public oclass As Object

Public Sub ShowContract(lContractKey As Long, bValid As Boolean)
    Dim lParentKey As Long
    
    If lContractKey > 0 Then
        lParentKey = glGetValidLong(frmChngOrd.oclass.moAppDB.Lookup("p.ParentContractKey", "tctContract AS p", "p.ContractKey =" & lContractKey))
        If lParentKey > 0 Then
            lkuContract.Text = frmChngOrd.oclass.moAppDB.Lookup("p.ContractNo", "tctContract AS p", "p.ContractKey =" & lParentKey)
            lkuSuplement.Text = frmChngOrd.oclass.moAppDB.Lookup("p.ContractNo", "tctContract AS p", "p.ContractKey =" & lContractKey)
        Else
            lkuContract.Text = frmChngOrd.oclass.moAppDB.Lookup("p.ContractNo", "tctContract AS p", "p.ContractKey =" & lContractKey)
        End If
    Else
        lkuContract.Text = ""
        lkuSuplement.Text = ""
        lkuSuplement.RestrictClause = "1=2"
    End If
    
    mbValid = False
    mlContractKey = lContractKey
    Form_Load
    frmContractAssociate.Show vbModal
    If mbValid = True Then
        bValid = True
        lContractKey = mlContractKey
    End If
End Sub

Private Sub cmdOK_Click()
    If Len(Trim$(lkuSuplement)) = 0 Then
        If Len(Trim$(lkuContract)) = 0 Then
            If mlContractKey > 0 Then
                If MsgBox("Desea desasociar el contrato actual?", vbYesNo, "Confirmar") <> vbYes Then
                    Exit Sub
                End If
            Else
                MsgBox "No se ha asociado ningun contrato Válido", vbInformation, "Alerta"
            End If
            mlContractKey = 0
            mbValid = True
            Hide
            Exit Sub
        End If
        mbValid = True
        If lkuContract.KeyValue <= 0 Then
            lkuContract_Validate False
            Exit Sub
        End If
        mlContractKey = lkuContract.KeyValue
        Hide
    Else
        mbValid = True
        mlContractKey = lkuSuplement.KeyValue
        Hide
    End If
End Sub

Private Sub Form_Load()
    With lkuContract
        Set .Framework = frmChngOrd.oclass.moFramework
        Set .SysDB = frmChngOrd.oclass.moAppDB
        Set .AppDatabase = frmChngOrd.oclass.moAppDB
        
       .RestrictClause = IIf(lVendKey > 0, "VendorKey =" & lVendKey & " and ", "") _
        & " ContractKey in (select ContractKey from tctContract where ParentContractKey is null and State =(SELECT TOP 1 s.ContractStateKey FROM tctContractState AS s WHERE s.ContractStateId = 'Activo'))"
    End With
    
    With lkuSuplement
        Set .Framework = frmChngOrd.oclass.moFramework
        Set .SysDB = frmChngOrd.oclass.moAppDB
        Set .AppDatabase = frmChngOrd.oclass.moAppDB
         .RestrictClause = IIf(lVendKey > 0, "VendorKey =" & lVendKey & " and ", "") _
        & " ContractKey in (select ContractKey from tctContract where ParentContractKey is not null or State <> (SELECT TOP 1 s.ContractStateKey FROM tctContractState AS s WHERE s.ContractStateId = 'Activo'))"
    End With
End Sub

Private Sub lkuContract_Validate(Cancel As Boolean)
    Dim lContractKey As Long
    Dim lHabilitado As Long
    Dim sFecha1 As Long
    Dim sEnableDate As String
    Dim dEnableDate As Date
    
    On Error GoTo ErrorHandler
    
    If lkuContract.Text <> lkuContract.Tag Then
        If Len(Trim$(lkuContract.Text)) > 0 Then
        
           Dim MyString As String
           MyString = Mid$(lkuContract.Text, 1, 1)
           
           If IsNumeric(MyString) Then
           lkuContract.Text = frmChngOrd.oclass.moAppDB.Lookup("p.ContractNo", "tctContract AS p", "p.ContractKey =" & lkuContract)
           End If
           
            lContractKey = glGetValidLong(frmChngOrd.oclass.moAppDB.Lookup("p.ContractKey", "tctContract AS p", "p.ContractNo = " & gsQuoted(lkuContract.Text) & " and " & lkuContract.RestrictClause))
            lHabilitado = glGetValidLong(frmChngOrd.oclass.moAppDB.Lookup("p.State", "tctContract AS p", "p.ContractNo = " & gsQuoted(lkuContract.Text) & " and " & lkuContract.RestrictClause))
            
            If lHabilitado > 0 Then
             
          '   sFecha1 = glGetValidLong(frmChngOrd.oClass.moAppDB.Lookup("p.ValidationTime", "tctContractOptions AS p"))
             'sEnableDate = gsGetValidStr(frmChngOrd.oclass.moAppDB.Lookup(" Dateadd(year,Duration,(DATEADD(DAY, (CASE WHEN EnableDate is null then 0 else (Select ValidationTime from tctContractOptions) end),StartDate)))", "tctContract", "ContractKey = " & gsQuoted(lContractKey)))
                   
             'If sEnableDate <> "" Then
              ' dEnableDate = CDate(sEnableDate)
               ' If dEnableDate < DateTime.Now Then
               '  MsgBox "El Contrato está vencido. Consulte ha expirado el tiempo de duración del Contrato."
                ' lkuContract.ClearData
                '  Exit Sub
                
               ' Else
                  'MsgBox "El Contrato estará vigente hasta el " & sEnableDate & " ."
              ' End If
            ' End If
              
            End If
            
          
            If lContractKey <> 0 Then
                lkuContract.Tag = lkuContract.Text
                lkuContract.KeyValue = lContractKey
                
                lkuContract.Text = frmChngOrd.oclass.moAppDB.Lookup("p.ContractNo", "tctContract AS p", "p.ContractKey =" & lContractKey)
                'lkuContract.KeyValue = lContractKey
                lkuSuplement.RestrictClause = "ContractKey in (SELECT p.ContractKey FROM tctContract AS p WHERE p.ParentContractKey = " & lContractKey & " and State = (SELECT TOP 1 s.ContractStateKey FROM tctContractState AS s WHERE s.ContractStateId = 'Activo'))"
                lkuSuplement.Enabled = True
                Exit Sub
            Else
                MsgBox "Debe seleccionar un Contrato Válido", vbInformation, "Alerta"
                lkuContract.Text = ""
                lkuContract.KeyValue = 0
            End If
        End If
        lkuSuplement.Enabled = False
        lkuContract.Tag = lkuContract.Text
        lkuContract.KeyValue = 0
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub lkuSuplement_Validate(Cancel As Boolean)
    Dim lContractKey As Long
    
    On Error GoTo ErrorHandler
    
    If lkuSuplement.Text <> lkuSuplement.Tag Then
        If Len(Trim$(lkuSuplement.Text)) > 0 Then
        
        Dim MyString As String
           MyString = Mid$(lkuSuplement.Text, 1, 1)
           
           If IsNumeric(MyString) Then
           lkuSuplement.Text = frmChngOrd.oclass.moAppDB.Lookup("p.ContractNo", "tctContract AS p", "p.ContractKey =" & lkuSuplement)
           End If
        
            lContractKey = glGetValidLong(oclass.moAppDB.Lookup("p.ContractKey", "tctContract as p", "p.ContractNo = " & gsQuoted(lkuSuplement.Text) & " and ParentContractKey =" & lkuContract.KeyValue & " and p.State = (SELECT TOP 1 s.ContractStateKey FROM tctContractState AS s WHERE s.ContractStateId = 'Activo')"))
            If lContractKey <> 0 Then
                lkuSuplement.Tag = lkuSuplement.Text
                lkuSuplement.KeyValue = lContractKey
                Exit Sub
            Else
                MsgBox "Debe seleccionar un Suplemento Válido", vbInformation, "Alerta"
                lkuSuplement.Text = ""
            End If
        End If
        lkuSuplement.Tag = lkuSuplement.Text
        lkuSuplement.KeyValue = 0
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error"
End Sub
