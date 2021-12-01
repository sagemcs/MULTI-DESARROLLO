VERSION 5.00
Begin VB.Form frmContractBank 
   Caption         =   "Información Bancaria"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   2880
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtBankAddress 
      Height          =   735
      Left            =   1680
      TabIndex        =   6
      Text            =   "Bank address"
      Top             =   2040
      Width           =   5055
   End
   Begin VB.TextBox txtSWIFT 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "SWIFT"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtAccountNo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Text            =   "Account no"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtTitular 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Titular"
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblBankAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección Bancaria"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1305
   End
   Begin VB.Label lblSWIFT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SWIFT"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label lblAccountNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Cuenta"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   945
   End
   Begin VB.Label lblTitular 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titular"
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   570
   End
End
Attribute VB_Name = "frmContractBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents modmForm1 As clsDmForm
Attribute modmForm1.VB_VarHelpID = -1
'Public moDmContractBankInf As clsDmForm

Public oClass As Object

Private Sub cmdAceptar_Click()
'    If Not bValidData Then Exit Sub
'     If modmForm1 Then
'        frmContractBank.Save.Save (True)
'        frmContract.txtAccountID.Text = txtAccountNo.Text
'        frmContract.txtBankAddress.Text = txtBankAddress.Text
'        frmContract.txtSWIFT.Text = txtSWIFT.Text
'        frmContract.txtTitular.Text = txtTitular.Text
'   '  End If
'        txtAccountNo.Text = ""
'        txtBankAddress.Text = ""
'        txtSWIFT.Text = ""
'        txtTitular.Text = ""
'
'    Hide

If Not bValidData Then Exit Sub
'    If modmForm1.IsDirty Then
'        modmForm1.Save (True)
'    End If
        frmContract.txtAccountID.Text = txtAccountNo.Text
        frmContract.txtBankAddress.Text = txtBankAddress.Text
        frmContract.txtSWIFT.Text = txtSWIFT.Text
        frmContract.txtTitular.Text = txtTitular.Text
        
        frmContract.ssAccountNo = txtAccountNo.Text
        frmContract.ssBankAddress = txtBankAddress.Text
        frmContract.ssSWIFT = txtSWIFT.Text
        frmContract.ssTitular = txtTitular.Text

        txtAccountNo.Text = ""
        txtBankAddress.Text = ""
        txtSWIFT.Text = ""
        txtTitular.Text = ""
   ' modmForm1.Clear (True)
    Hide
End Sub

Private Function bValidData() As Boolean
    bValidData = False
   
   If txtAccountNo.Enabled = True Then
     If Len(Trim$(txtAccountNo.Text)) = 0 Then
         MsgBox "Debe introducir el número de Contrato", vbExclamation, "Alerta"
         Exit Function
     End If
    
    
     If Len(Trim$(txtBankAddress.Text)) = 0 Then
        MsgBox "Debe introducir la Dirección del Banco", vbExclamation, "Alerta"
        Exit Function
    End If
    
    
     If Len(Trim$(txtSWIFT.Text)) = 0 Then
        MsgBox "Debe introducir el SWIFT", vbExclamation, "Alerta"
        Exit Function
    End If
    
    
     If Len(Trim$(txtTitular.Text)) = 0 Then
        MsgBox "Debe introducir el Titular", vbExclamation, "Alerta"
        Exit Function
    End If
   End If
   
   
    
    
    
    bValidData = True
End Function


'Public Sub setupComponents(msCompany As String)
'    With lkuContact
'        Set .Framework = oClass.moFramework
'        Set .AppDatabase = oClass.moAppDB
''        .RestrictClause = msLookupRestrict
'    End With
'
'    With lkuVendorClass
'        Set .Framework = oClass.moFramework
'        Set .AppDatabase = oClass.moAppDB
''        .RestrictClause = msLookupRestrict
'    End With
'
'    'sddPaymentTerms.InitDynamicList oClass.moAppDB, "SELECT p.PmtTermsID, p.PmtTermsKey FROM tciPaymentTerms AS p"
'    'sddFOB.InitDynamicList oClass.moAppDB, "SELECT p.FOBID, p.FOBKey FROM tciFOB AS p WHERE p.CompanyID =  " & gsQuoted(msCompany)
'
'End Sub


Private Sub Form_Load()
' On Error GoTo ErrorHandler

'    Set modmForm1 = New clsDmForm
'
'    With modmForm1
'        Set .Form = frmContractBank
'        Set .Session = frmContract.oClass.moSysSession
'        'oClass.moSysSession
'        Set .Database = frmContract.oClass.moAppDB
'        .AppName = gsStripChar(frmContractBank.Caption, ".")
'        .UniqueKey = "ContractKey"
'        .Table = "tctContractBankInf"
'
'        .Bind Nothing, "ContractKey", SQL_INTEGER
'        '.BindLookup frmContractChg.lkuVendorClass
'        '.BindLookup frmContractChg.lkuContact
'        .Bind txtAccountNo, "AccountID", SQL_INTEGER
'        .Bind txtTitular, "Titular", SQL_VARCHAR
'        .Bind txtSWIFT, "SWIFT", SQL_VARCHAR
'        .Bind txtBankAddress, "BankAddress", SQL_VARCHAR
'        '.Bind dtpSignatureDate, "SignatureDate", SQL_DATE
'        '.Bind txtCmnt, "Cmnt", SQL_CHAR
'        .Init
'    End With
'    Exit Sub
'ErrorHandler:
 '   MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '  modmForm.Clear (True)
End Sub

