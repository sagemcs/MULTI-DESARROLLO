VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{BC90D6A3-491E-451B-ADED-8FABA0B8EE36}#57.0#0"; "SOTADropDown.ocx"
Object = "{2A076741-D7C1-44B1-A4CB-E9307B154D7C}#185.0#0"; "EntryLookupControls.ocx"
Begin VB.Form frmContractChg 
   Caption         =   "Datos Generales a Actualizar"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
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
   ScaleHeight     =   3810
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   360
      Left            =   3000
      TabIndex        =   11
      Top             =   3240
      Width           =   990
   End
   Begin VB.Frame frmData 
      Caption         =   "Datos Actualizables"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.TextBox txtCmnt 
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   2040
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpSignatureDate 
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   192020481
         CurrentDate     =   43704
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   192020481
         CurrentDate     =   43704
      End
      Begin NEWSOTALib.SOTANumber nbrDuration 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
         text            =   "           0.00"
         sDecimalPlaces  =   2
      End
      Begin SOTADropDownControl.SOTADropDown sddFOB 
         Height          =   315
         Left            =   4920
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "sddFOB"
      End
      Begin SOTADropDownControl.SOTADropDown sddPaymentTerms 
         Height          =   315
         Left            =   4920
         TabIndex        =   6
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "sddPaymentTerms"
      End
      Begin EntryLookupControls.TextLookup lkuContact 
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         ForeColor       =   -2147483640
         IsSurrogateKey  =   -1  'True
         LookupID        =   "Contact"
         ParentIDColumn  =   "Name"
         ParentKeyColumn =   "CntctKey"
         ParentTable     =   "tciContact"
         BoundColumn     =   "CntctKey"
         BoundTable      =   "tctContractUpgrades"
         IsForeignKey    =   -1  'True
         Datatype        =   0
         sSQLReturnCols  =   "Name,lkuContact,;Title,,;CntctKey,,;"
      End
      Begin EntryLookupControls.TextLookup lkuVendorClass 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         ForeColor       =   -2147483640
         IsSurrogateKey  =   -1  'True
         LookupID        =   "VendorClass"
         ParentIDColumn  =   "VendClassID"
         ParentKeyColumn =   "VendClassKey"
         ParentTable     =   "tapVendClass"
         BoundColumn     =   "VendorClassKey"
         BoundTable      =   "tctContractUpgrades"
         IsForeignKey    =   -1  'True
         Datatype        =   0
         sSQLReturnCols  =   "VendClassID,lkuVendorClass,;"
      End
      Begin VB.Label lblCmnt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón de la Actualización"
         Height          =   195
         Left            =   2640
         TabIndex        =   17
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblSignatureDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Firma"
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblDuration 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vigencia"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblFOB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condiciones de Entrega"
         Height          =   195
         Left            =   4920
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblPaymentTerms 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condiciones de Pago"
         Height          =   195
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblContact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto"
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblVendorClass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clase del Proveedor"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmContractChg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents modmForm As clsDmForm
Attribute modmForm.VB_VarHelpID = -1

Public oClass As Object

Private Sub cmdActualizar_Click()
    If Not bValidData Then Exit Sub
    If modmForm.IsDirty Then
        modmForm.Save (True)
    End If
    modmForm.Clear (True)
    
    
    
    
    Hide
End Sub

Public Sub setupComponents(msCompany As String)
    With lkuContact
        Set .Framework = oClass.moFramework
        Set .AppDatabase = oClass.moAppDB
'        .RestrictClause = msLookupRestrict
    End With

    With lkuVendorClass
        Set .Framework = oClass.moFramework
        Set .AppDatabase = oClass.moAppDB
'        .RestrictClause = msLookupRestrict
    End With

    sddPaymentTerms.InitDynamicList oClass.moAppDB, "SELECT p.PmtTermsID, p.PmtTermsKey FROM tciPaymentTerms AS p"
    sddFOB.InitDynamicList oClass.moAppDB, "SELECT p.FOBID, p.FOBKey FROM tciFOB AS p WHERE p.CompanyID =  " & gsQuoted(msCompany)

End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Set modmForm = New clsDmForm
    
    With modmForm
        Set .Form = frmContractChg
        Set .Session = oClass.moSysSession
        Set .Database = oClass.moAppDB
        .AppName = gsStripChar(frmContractChg.Caption, ".")
        .UniqueKey = "ContractKey"
        .Table = "tctContractUpgrades"
        
        .Bind Nothing, "ContractKey", SQL_INTEGER
        .BindLookup frmContractChg.lkuVendorClass
        .BindLookup frmContractChg.lkuContact
        .Bind nbrDuration, "Duration", SQL_INTEGER
        .Bind sddFOB, "FOBKey", SQL_INTEGER, kDmUseItemData
        .Bind sddPaymentTerms, "PmtTermsKey", SQL_INTEGER, kDmUseItemData
       ' .Bind dtpStartDate, "StartDate", SQL_DATE
       ' .Bind dtpSignatureDate, "SignatureDate", SQL_DATE
        .Bind txtCmnt, "Cmnt", SQL_CHAR
        .Init
    End With
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modmForm.Clear (True)
End Sub

Private Function bValidData() As Boolean
    bValidData = False
    If Len(Trim$(lkuContact)) = 0 Then
        MsgBox "Debe seleccionar un Contacto valido", vbExclamation, "Alerta"
        Exit Function
    End If
    
    If Len(Trim$(lkuVendorClass)) = 0 Then
        MsgBox "Debe seleccionar una Clase de proveedor valida", vbExclamation, "Alerta"
        Exit Function
    End If
    
    If sddFOB.ListIndex = -1 Then
        MsgBox "Debe seleccionar condiciones de entrega validas", vbExclamation, "Alerta"
        Exit Function
    End If
    
    If sddPaymentTerms.ListIndex = -1 Then
        MsgBox "Debe seleccionar condiciones de pago validas", vbExclamation, "Alerta"
        Exit Function
    End If
    
    If nbrDuration <= 0 Then
        MsgBox "La vigencia debe ser mayor que 0", vbExclamation, "Alerta"
        Exit Function
    End If
    
'    If DateDiff("d", dtpSignatureDate.Value, dtpStartDate) < 0 Then
'        MsgBox "La fecha de firma debe ser anterior a la de inicio", vbExclamation, "Alerta"
'        Exit Function
'    End If
    
    If Len(Trim$(txtCmnt.Text)) = 0 Then
        MsgBox "Debe introducir el comentario de la Razón del cambio", vbExclamation, "Alerta"
        Exit Function
    End If
    
    bValidData = True
End Function

Private Sub lkuContact_Change()
  If lkuContact.Text <> "" Then
      frmContract.lkuContact.Text = lkuContact.Text
    End If
End Sub

Private Sub lkuVendorClass_Change()
 If lkuVendorClass.Text <> "" Then
      frmContract.lkuVendClass.Text = lkuVendorClass.Text
    End If
End Sub

Private Sub nbrDuration_Change()
    If nbrDuration.Value <> 0 Then
      frmContract.nbrDuration = nbrDuration.Value
    End If
    
End Sub

Private Sub sddFOB_Validate(Cancel As Boolean)
   If sddFOB.ItemData <> 0 Then
      frmContract.sddFOB.ItemData = sddFOB.ItemData
    End If
End Sub

Private Sub sddPaymentTerms_Validate(Cancel As Boolean)
  If sddPaymentTerms.ItemData <> 0 Then
      frmContract.sddPaymentTerms.ItemData = sddPaymentTerms.ItemData
    End If
End Sub
