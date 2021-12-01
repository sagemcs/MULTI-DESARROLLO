VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{F2F2EE3C-0D23-4FC8-944C-7730C86412E3}#67.0#0"; "sotasbar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "Sysinfo.ocx"
Object = "{8A9C5D3D-5A2F-4C5F-A12A-A955C4FB68C8}#101.0#0"; "LookupView.ocx"
Object = "{2A076741-D7C1-44B1-A4CB-E9307B154D7C}#185.0#0"; "EntryLookupControls.ocx"
Object = "{BC90D6A3-491E-451B-ADED-8FABA0B8EE36}#57.0#0"; "SOTADropDown.ocx"
Object = "{0FA91D91-3062-44DB-B896-91406D28F92A}#65.0#0"; "SOTACalendar.ocx"
Object = "{C41A85E3-4CB6-40B5-B425-EE9ECC5E6F06}#181.0#0"; "SOTATbar.ocx"
Begin VB.Form frmRequistn 
   Caption         =   "Enter Requisitions"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15810
   HelpContextID   =   54712
   Icon            =   "pozde001.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7365
   ScaleWidth      =   15810
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7080
      TabIndex        =   63
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdContract 
      Caption         =   "Contrato"
      Height          =   375
      Left            =   8760
      TabIndex        =   62
      Top             =   6480
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txtAceptaReq 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   13425
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txt2doAutorizaReq 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   13425
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   600
      Width           =   1935
   End
   Begin VB.CheckBox chkb2doAutorizoReq 
      Caption         =   "Autorizado por:"
      Enabled         =   0   'False
      Height          =   315
      Left            =   11880
      TabIndex        =   53
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtUserMod 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   2640
      Width           =   2490
   End
   Begin VB.TextBox txtAutorizaReq 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1305
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtEstatusReqDesc 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   8640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CheckBox chkbEstatusDesc 
      Caption         =   "Estatus Req. Descripciòn:"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   43
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CheckBox chkb2doAuthNeed 
      Caption         =   "Requiere Segunda Autorizaciòn"
      Height          =   375
      Left            =   8640
      TabIndex        =   42
      Top             =   960
      Width           =   2655
   End
   Begin SOTADropDownControl.SOTADropDown sddEstatusReq 
      Height          =   315
      Left            =   9720
      TabIndex        =   40
      Top             =   570
      Width           =   1815
      _ExtentX        =   3201
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
      Style           =   1
      Text            =   "sddEstatusReq"
      StaticListTableName=   "tpoReqAdicInfo"
      StaticListColumnName=   "statusKey"
   End
   Begin VB.CommandButton cmdUserFlds 
      Caption         =   "Custom &Fields..."
      Height          =   285
      Left            =   5040
      TabIndex        =   39
      Top             =   1920
      WhatsThisHelpID =   95457
      Width           =   1725
   End
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1110
      MaxLength       =   40
      TabIndex        =   15
      Top             =   1875
      WhatsThisHelpID =   54733
      Width           =   3450
   End
   Begin SOTACalendarControl.SOTACalendar CustomDate 
      Height          =   315
      Index           =   0
      Left            =   -30000
      TabIndex        =   36
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
   Begin SOTADropDownControl.SOTADropDown cboExpReason 
      Height          =   315
      Left            =   6120
      TabIndex        =   9
      Top             =   1215
      WhatsThisHelpID =   54754
      Width           =   2205
      _ExtentX        =   3889
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
      Text            =   "cboExpReason"
   End
   Begin VB.ComboBox CustomCombo 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   -30000
      Style           =   2  'Dropdown List
      TabIndex        =   22
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
      TabIndex        =   23
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
      TabIndex        =   24
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
      TabIndex        =   25
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
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComCtl2.UpDown CustomSpin 
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   27
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
      TabIndex        =   28
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
      TabIndex        =   29
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
      TabIndex        =   30
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
      TabIndex        =   21
      Top             =   645
      Visible         =   0   'False
      WhatsThisHelpID =   70
      Width           =   345
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   582
      _StockProps     =   0
   End
   Begin EntryLookupControls.TextLookup lkuDept 
      Height          =   285
      Left            =   1110
      TabIndex        =   11
      Top             =   1560
      WhatsThisHelpID =   54743
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   503
      ForeColor       =   -2147483640
      LookupID        =   "POPurchDept"
      Datatype        =   0
      sSQLReturnCols  =   "PurchDeptID,,;"
   End
   Begin SOTACalendarControl.SOTACalendar calDate 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5055
      TabIndex        =   4
      Top             =   900
      WhatsThisHelpID =   54742
      Width           =   2080
      _ExtentX        =   3678
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
      MaskedText      =   "  /  /    "
      Text            =   "  /  /    "
      Object.CausesValidation=   0   'False
   End
   Begin EntryLookupControls.TextLookup lkuMain 
      Height          =   285
      Left            =   1110
      TabIndex        =   0
      Top             =   570
      WhatsThisHelpID =   54741
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   503
      ForeColor       =   -2147483640
      LookupMode      =   0
      MaxLength       =   10
      LookupID        =   "POReqNo"
      BoundColumn     =   "TranNo"
      BoundTable      =   "tpoRequisition"
      sSQLReturnCols  =   "TranNo,,;"
   End
   Begin StatusBar.SOTAStatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      Top             =   6975
      WhatsThisHelpID =   73
      Width           =   15810
      _ExtentX        =   27887
      _ExtentY        =   688
      MessageVisible  =   -1  'True
   End
   Begin SOTAToolbarControl.SOTAToolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      WhatsThisHelpID =   71
      Width           =   15810
      _ExtentX        =   27887
      _ExtentY        =   741
      Style           =   9
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5055
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   570
      WhatsThisHelpID =   54738
      Width           =   2080
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate PO"
      Height          =   390
      Left            =   6840
      TabIndex        =   16
      Top             =   6480
      WhatsThisHelpID =   54737
      Width           =   1290
   End
   Begin VB.CheckBox chkExpedite 
      Alignment       =   1  'Right Justify
      Caption         =   "&Expedite"
      Height          =   195
      Left            =   4080
      TabIndex        =   7
      Top             =   1260
      WhatsThisHelpID =   54735
      Width           =   1170
   End
   Begin VB.TextBox txtContact 
      Height          =   285
      Left            =   1110
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1230
      WhatsThisHelpID =   54734
      Width           =   1890
   End
   Begin VB.TextBox txtOriginator 
      Height          =   285
      Left            =   1110
      MaxLength       =   40
      TabIndex        =   2
      Top             =   900
      WhatsThisHelpID =   54733
      Width           =   1890
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   5520
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtNavReturn 
      Height          =   285
      Left            =   2520
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      WhatsThisHelpID =   54730
      Width           =   1425
   End
   Begin LookupViewControl.LookupView navItemGrid 
      Height          =   285
      Left            =   3870
      TabIndex        =   33
      Top             =   90
      WhatsThisHelpID =   54729
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      LookupMode      =   1
   End
   Begin LookupViewControl.LookupView navVendorGrid 
      Height          =   285
      Left            =   0
      TabIndex        =   34
      Top             =   0
      WhatsThisHelpID =   54728
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      LookupMode      =   1
   End
   Begin LookupViewControl.LookupView navDeptGrid 
      Height          =   285
      Left            =   0
      TabIndex        =   35
      Top             =   0
      WhatsThisHelpID =   54727
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      LookupMode      =   1
   End
   Begin LookupViewControl.LookupView navSTaxGrid 
      Height          =   285
      Left            =   0
      TabIndex        =   37
      Top             =   0
      WhatsThisHelpID =   54725
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      LookupMode      =   1
   End
   Begin EntryLookupControls.TextLookup lkuWarehouse 
      Height          =   285
      Left            =   5040
      TabIndex        =   13
      Top             =   1560
      WhatsThisHelpID =   54724
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   503
      ForeColor       =   -2147483640
      LookupID        =   "Warehouse"
      Datatype        =   0
      sSQLReturnCols  =   "WhseID,lkuWarehouse,;Description,,;"
   End
   Begin LookupViewControl.LookupView navWhseGrid 
      Height          =   285
      Left            =   0
      TabIndex        =   38
      Top             =   0
      WhatsThisHelpID =   54723
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      LookupMode      =   1
   End
   Begin SOTADropDownControl.SOTADropDown sddType 
      Height          =   315
      Left            =   1110
      TabIndex        =   47
      Top             =   2280
      Width           =   2535
      _ExtentX        =   4471
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
      Style           =   1
      Text            =   "sddType"
      StaticListTableName=   "tpoReqAdicInfo"
      StaticListColumnName=   "Type"
   End
   Begin SOTACalendarControl.SOTACalendar calAutorizaReq 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   5040
      TabIndex        =   49
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
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
      Object.CausesValidation=   0   'False
   End
   Begin SOTACalendarControl.SOTACalendar calAceptaReq 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13425
      TabIndex        =   55
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
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
      Object.CausesValidation=   0   'False
   End
   Begin SOTACalendarControl.SOTACalendar cal2doAutorizaReq 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   13425
      TabIndex        =   57
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
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
      Object.CausesValidation=   0   'False
   End
   Begin LookupViewControl.LookupView navBuyerGrid 
      Height          =   285
      Left            =   6600
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   7440
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      LookupMode      =   1
   End
   Begin FPSpreadADO.fpSpread grdReqLineDtl 
      Height          =   1485
      Left            =   840
      TabIndex        =   64
      Top             =   4200
      Visible         =   0   'False
      WhatsThisHelpID =   54755
      Width           =   13935
      _Version        =   524288
      _ExtentX        =   24580
      _ExtentY        =   2619
      _StockProps     =   64
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      MaxRows         =   1
      SpreadDesigner  =   "pozde001.frx":7D32
      AppearanceStyle =   0
   End
   Begin FPSpreadADO.fpSpread grdReqLines 
      Height          =   2820
      Left            =   120
      TabIndex        =   65
      Top             =   3600
      WhatsThisHelpID =   54756
      Width           =   15225
      _Version        =   524288
      _ExtentX        =   26855
      _ExtentY        =   4974
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
      NoBeep          =   -1  'True
      SpreadDesigner  =   "pozde001.frx":816D
      AppearanceStyle =   0
   End
   Begin VB.Label lblAceptaReq 
      AutoSize        =   -1  'True
      Caption         =   "Aceptado por:"
      Height          =   195
      Left            =   12330
      TabIndex        =   60
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label lblFechaAceptaReq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Aceptación:"
      Height          =   195
      Left            =   11985
      TabIndex        =   59
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Autorizo:"
      Height          =   195
      Left            =   12225
      TabIndex        =   58
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label lblRechazaReq 
      AutoSize        =   -1  'True
      Caption         =   "Autorizado por:"
      Height          =   195
      Left            =   120
      TabIndex        =   51
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Label lblFechaAutorizaReq 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Autorizo:"
      Height          =   195
      Left            =   3840
      TabIndex        =   50
      Top             =   3000
      Width           =   1110
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   120
      TabIndex        =   46
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label lblUserMod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modificado"
      Height          =   195
      Left            =   120
      TabIndex        =   45
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label lblStatusReq 
      AutoSize        =   -1  'True
      Caption         =   "Estatus Req.:"
      Height          =   195
      Left            =   8640
      TabIndex        =   41
      Top             =   615
      Width           =   960
   End
   Begin VB.Label lblWarehouse 
      Caption         =   "&Warehouse"
      Height          =   195
      Left            =   4125
      TabIndex        =   12
      Top             =   1620
      Width           =   885
   End
   Begin VB.Label CustomLabel 
      Caption         =   "Label"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   -30000
      TabIndex        =   31
      Top             =   60
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblDept 
      Caption         =   "De&partment"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1620
      Width           =   945
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   240
      Left            =   4125
      TabIndex        =   18
      Top             =   615
      Width           =   765
   End
   Begin VB.Label lblContact 
      Caption         =   "&Contact"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   1275
      Width           =   690
   End
   Begin VB.Label lblReason 
      Caption         =   "Re&ason"
      Height          =   195
      Left            =   5475
      TabIndex        =   8
      Top             =   1275
      Width           =   615
   End
   Begin VB.Label lblOriginator 
      Caption         =   "&Originator"
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   915
      Width           =   840
   End
   Begin VB.Label lblDate 
      Caption         =   "&Date"
      Height          =   195
      Left            =   4125
      TabIndex        =   3
      Top             =   915
      Width           =   615
   End
   Begin VB.Label lblComment 
      Caption         =   "Co&mment"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1935
      Width           =   765
   End
   Begin VB.Label lblReqNum 
      AutoSize        =   -1  'True
      Caption         =   "Requisition"
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   615
      Width           =   780
   End
End
Attribute VB_Name = "frmRequistn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************
'     Name: POZDE001 - Requisition Entry
'     Desc: Requisition Entry
' Original: MAD 03-06-1998
'     Mods:
'************************************************************************************
Option Explicit

#If CUSTOMIZER Then
    Public moFormCust As Object
#End If

    Private mbFocus  As Boolean

    'Private variables of Form Properties
    Public moClass              As Object                   ' class reference
    Private mlRunMode           As Long                     ' run mode of task
    Private mbCancelShutDown    As Boolean                  ' cancel shutdown flag
    Private miSecurityLevel     As Integer                  ' security level
    Private mbScrolling         As Boolean
    Private mbDontSave          As Boolean                  ' Dont save due to errors.

    Public mbCancelLoad         As Boolean
    
    Private mbSelectFormLoaded As Boolean

    ' declare filter state tracking variable
    Private miFilter            As Integer                  ' browse filter

    Private moIMSClass          As New clsIMS
    Private moItem              As clsIMSItem


    'Public Form Variables
    Public moSotaObjects        As New Collection           ' collection of loaded objects
    Public moToolbar            As SOTAToolbar              ' toolbar object
    
    Private mbDontChkclick      As Boolean                  'global don't run click logic
    Private mbPrintingReq       As Boolean                  ' Prevent certain processing if currently printing.
    
    ' private objects for this project
    Public moContextMenu         As clsContextMenu           ' context menu
    Private WithEvents moGM                As clsGridMgr               ' grid manager class
Attribute moGM.VB_VarHelpID = -1
    Private moOptions           As New clsModuleOptions    ' Sage MAS 500 Module Options class
    
    ' public objects for creating the Error Log
    Private moReportObj As clsReportEngine
    Private moDBObj As Object
    Private moDDData As clsDDData
    Private moRealTableCollection As Collection
    Private mbPrintErrorReport As Boolean  ' Flag to indicate whether the error report should be displayed
    Private mlSession As Long              ' The session id for the error report.
    
    

    
    ' Minimum form size
    Private miOldFormHeight As Long
    Private miOldFormWidth As Long
    Private miMinFormHeight As Long
    Private miMinFormWidth As Long


  'Binding Object Variables
    Public moDmForm                 As clsDmForm            ' parent object
    Public moDmReqAdicInfo           As clsDmForm            'Agregado por Multiconsulting
    Public moDmGrid                 As clsDmGrid            ' grid object
    Public moDMSubGrid              As clsDmGrid            ' grid object
    Public moGridNav                As clsGridLookup
    Public moGridNav2                As clsGridLookup
    Public moGridNav3                As clsGridLookup
    Public moGridNav4                As clsGridLookup
    Public moGridNav5                As clsGridLookup           ' grid navigator
    
  'Miscellaneous Variables
    Public msCompanyID              As String               ' company id
    Private msBusinessDate          As String               ' business date
    Private msHomeCurrID            As String
    Private msTranTypeID            As String
    Private mlLanguage              As Long
    Private msCurrentUser           As String
    Private miCostDecPlaces         As Integer
    Private miQtyDecPlaces          As Integer
    Private mbReqClosed             As Boolean
    Private mbTrackSTax             As Boolean
    Private mbIntegrateWithIM       As Boolean              ' Is PO Integrated with IM
    Private mbIntegratedCT          As Boolean              'Agregado por Multiconsulting Jose
    Private msOldItemID             As String
    Private msOldVendID             As String
    Private msOldSTaxClassID        As String
    Private msOldPurchDeptID        As String
    Private msOldWhseID             As String
    Private msOldReqQtyRequested    As String
    Private msOldReqUnitCost        As String
    Private msOldReqExtAmt          As String
    
   'By QQ this flag is set for fix #7872
    Private mbSkipGotFocus         As Boolean
    
    Private mlCurrRow As Long
    
    
    Private msItemDescCaption       As String
    Private msItemIDCaption         As String
    Private msRequestDateCaption    As String
    
    Private miTemp                  As Integer
    Private mbIsInvalid             As Boolean
    Private mbIsPressF5             As Boolean
    
    Private Const kMaxCols = 35                             ' Maximum Number of columns Modificado por Multiconsulting
    Private Const kColReqItemID = 1
    Private Const kColReqItemKey = 2
    Private Const kColReqLineKey = 3
    Private Const kColReqDescription = 4
    Private Const kColReqQtyRequested = 5
    Private Const kColReqUnitMeasKey = 6
    Private Const kColReqUnitMeasID = 7
    Private Const kColReqUnitCost = 8
    Private Const kColReqExtAmt = 9
    Private Const kColReqVendKey = 10
    Private Const kcolReqVendID = 11
    Private Const kColReqCurrid = 12                        'Added column currency to be displayed after vendor
    Private Const kcolReqRequestDate = 13
    Private Const kColReqPurchDeptKey = 14
    Private Const kColReqPurchDeptID = 15
    Private Const kColReqWhseKey = 16
    Private Const kColReqWhseID = 17
    Private Const kColReqSTaxClassKey = 18
    Private Const kColReqSTaxClassID = 19
    Private Const kColReqComment = 20
    Private Const kColReqPOlineCustomFields = 21
    Private Const kColReqPOLineKey = 22
    Private Const kColReqPOKey = 23
    Private Const kColReqPOTranID = 24
    Private Const kColPOUOMType = 25
    Private Const kColPOUserFld1 = 26
    Private Const kColPOUserFld2 = 27
    Private Const kColReqUnitCostExact = 28 ' Stores Un-rounded UnitCost
    'Agregado por Multiconsulting
    Private Const kColReqEstPres = 29
    Private Const kColReqBuyerkey = 30
    Private Const kColReqBuyerID = 31
    Private Const kColReqLStItemKey = 32
    Private Const kColReqLStItemID = 33
    Private Const kColReqLTBItemKey = 34
    Private Const kColReqLTBItemID = 35
    
    Public listRows                As New Collection           'Variable to save all rows was modified in grdReqLines_Change event
    Private lastRowMod          As Long                     'Variable to save the last row was modified in grdReqLines_Change event
    Private msTraceReqChange    As String                   'Variable to save all trace Req was to save requisition
    Private RequisitionChanged  As Boolean                  'Variable flag for denote the requisition modified field anything change
    Private msPersistDescriptionReq    As String                   'Variable presist for store a value of Descript Requested in Row Active in grdReqLines
    Private msPersistCantReq   As String                   'Variable presist for store a value of Cant Item Requested in Row Active in grdReqLines
    Private msPersistPresEstReq    As String                   'Variable presist for store a value of PresEst Requested in Row Active in grdReqLines
    Private msPersistCantRows    As Integer                   'Variable presist for store a value of PresEst Requested in Row Active in grdReqLines
    Public moGridNav6                As clsGridLookup
    
    Private msOldReqEstPres         As String
    Private msOldBuyerID            As String
    Private msOldStateItemID        As String
    Private msOldTypeBuyItemID      As String
    
    Dim sUser As String
    Dim SIDEtReqAutz As String
    Dim SIDEtReqAcpt As String
    Dim SIDEtReqRchz As String
    Dim SIDEtReqDelt As String
    Dim SIDEtReqBuyInfo As String
    
    Dim evSegUserAutz As Integer
    Dim evSegUserAcpt As Integer
    Dim evSegUserRchz As Integer
    Dim evSegUserDelt As Integer
    Dim evSegUserBuyInfo As Integer
    
    Private moStaticListStateItem As clsStaticList      'Lista para almacenamiento de los diferentes estados de los Items durante el proceso de compra
    Private moStaticListTypeBuyItem As clsStaticList    'Lista para almacenamiento de los diferentes tipos de compras para Items durante el proceso de compra
    
    Private Const kReqStatusPending = 0
    Private Const kReqStatusAuthorized = 1
    Private Const kReqStatusAccepted = 2
    Private Const kReqStatusCancel = 3
    'Agregado por Multiconsulting
    
    Private Const kChildPOMaxCols = 8
    Private Const kColChildReqLineDistKey = 1
    Private Const kColChildPurchDeptKey = 2
    Private Const kColChildPurchDeptID = 3
    Private Const kColChildQtyReq = 4
    Private Const kColChildRequestDate = 5
    Private Const kColChildFrtAmt = 6
    Private Const kColChildWhseKey = 7
    Private Const kColChildWhseID = 8

    
    ' Requisition status constants
    Private Const kvReqIncomplete As Integer = 0
    Private Const kvReqPendApprvl As Integer = 1
    Private Const kvReqOpen As Integer = 2
    Private Const kvReqInactive As Integer = 3
    Private Const kvReqCanceled As Integer = 4
    Private Const kvReqClosed As Integer = 5
    
    Private Const kvMaxQtySize As Double = 99999999.9999999
    Private Const kvMaxUnitCost As Double = 9999999999.99999
    Private Const kvMaxExtAmt As Double = 999999999999.999
    Private Const kvDfltLineStatus As Integer = 1
    
 ' Item Type constants
    Private Const kItemTypeMisc = 1
    Private Const kItemTypeExpense = 3
    Private Const kItemTypeComment = 4
    
' Item status constants
    Private Const kItemStatusActive = 1
    
    Private Declare Sub LoHiWord Lib "utils.dll" (ByVal lTaskID As Long, ByRef spLoValue As Integer, ByRef spHiValue As Integer)
    Private Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long, ByVal uFlags As Long) As Long

    'Agregado por Multiconsulting
    Private adicInfo                As Boolean
    Private reqIsClosed             As Boolean
    Private firstCompEstatusReq     As Boolean 'Variable usada para comprobaciones del Estatus inicial de la Requisición
    Private loadNewReq              As Boolean 'Variable para comprobar la carga de New Requisiciones
    Private msButtonInitAcc         As String  'Variable para salvar Acc que inicia la actividad
    'Agregado por Multiconsulting
    
'************************************************************************
'   Description:
'      Check all of the rows for the selected req.  If there is one or more
'      entered and all rows are associated with a PO, set the status of the
'      requisition to closed
'   Param:
'       <none>
'
'************************************************************************


Const VBRIG_MODULE_ID_STRING = "POZDE001.FRM"

'************************************************************************
'   bCancelShutDown tells the framework whether the form has requested
'   the shutdown process to be cancelled.
'************************************************************************

Public Property Get bCancelShutDown()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    bCancelShutDown = mbCancelShutDown

'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bCancelShutDown_Get", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Property
        End Select
'+++ VB/Rig End +++
End Property
Private Sub SetupIMInterfaces()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    '-- Setup the IMS Class
    moIMSClass.Init moClass.moAppDB, moClass.moAppDB, msCompanyID
    Set moItem = moIMSClass.Items.Item
    
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "SetupIMInterfaces", VBRIG_IS_FORM
        Select Case VBRIG_IS_FORM_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++

End Sub

Public Property Get bSelectFormLoaded() As Variant
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    bSelectFormLoaded = mbSelectFormLoaded

'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bSelectFormLoaded_Get", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Property
        End Select
'+++ VB/Rig End +++

End Property



Private Sub CheckStatus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lLoop As Long
    
    ' If there are no rows in the grid, dont close the req. (there will always be
    ' and appended row, so compare to 1 instead of 0.
    If grdReqLines.MaxRows <= 1 Then Exit Sub
    ' Loop through all except the last row (appended row).
    For lLoop = 1 To grdReqLines.MaxRows - 1
        If gsGridReadCellText(grdReqLines, lLoop, kColReqPOTranID) = "" Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Sub
        End If
    Next
    
    moDmForm.SetColumnValue "Status", kvReqClosed
    DisplayStatus

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CheckStatus", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub HideAllNavs()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    navWhseGrid.Visible = False
    navDeptGrid.Visible = False
    navVendorGrid.Visible = False
    navItemGrid.Visible = False
    navSTaxGrid.Visible = False
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "HideAllNavs", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub


Private Function bCalcLineAmts(lCurRow As Long, lCurCol As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

Dim fCurRowReqQty As Double
Dim fCurRowUnitCost As Double
Dim fCurRowUnitCostExact As Double
Dim fCurRowExtAmt As Double
Dim sID         As String
Dim sUser       As String
Dim vPrompt     As Variant
Dim fCurRowEstPres As Double 'Agregado por Multiconsulting Osmel Barreras
fCurRowUnitCostExact = 0
    If lCurCol = kColReqUnitCost Or lCurCol = kColReqExtAmt Or lCurCol = kColReqEstPres Then 'Modificado por Multiconsulting
    ' Subordinaciòn  temporal del Presupuesto Estimado al Permiso de Modificaciòn del Costo Unitario
    '   Check the user id typed in can change the unit cost.
        sID = CStr("CHGPOCOST")
        sUser = CStr(moClass.moSysSession.UserId)
        vPrompt = True
        If moClass.moFramework.GetSecurityEventPerm(sID, sUser, vPrompt) = 0 Then
            bCalcLineAmts = False
            '+++ VB/Rig Begin Pop +++
            '+++ VB/Rig End +++
            Exit Function
        End If
    End If
    Select Case lCurCol
        Case kColReqQtyRequested, kColReqUnitCost
            fCurRowReqQty = gdGetValidDbl(gsGridReadCell(grdReqLines, lCurRow, kColReqQtyRequested))
            fCurRowUnitCost = gdGetValidDbl(gsGridReadCell(grdReqLines, lCurRow, kColReqUnitCost))
            
            'Agregado por MultiCOnsulting Osmel Barreras
            fCurRowEstPres = gdGetValidDbl(gsGridReadCell(grdReqLines, lCurRow, kColReqEstPres))
            If fCurRowEstPres < 0 Then
                giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeNegative, _
                            "Presupuesto Estimado"
                bCalcLineAmts = False
                '+++ VB/Rig Begin Pop +++
                '+++ VB/Rig End +++
                Exit Function
            End If
            'Agregado por MultiCOnsulting Osmel Barreras
            
            If fCurRowUnitCost < 0 Then
                giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeNegative, _
                            "Unit Cost"
                bCalcLineAmts = False
                '+++ VB/Rig Begin Pop +++
                '+++ VB/Rig End +++
                Exit Function
            End If
            fCurRowExtAmt = fCurRowReqQty * fCurRowUnitCost
            If fCurRowExtAmt > kvMaxExtAmt Then
                giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeGreaterThan, _
                            "Ext Amt", kvMaxExtAmt
                bCalcLineAmts = False
                '+++ VB/Rig Begin Pop +++
                '+++ VB/Rig End +++
                Exit Function
            Else
                gGridUpdateCell grdReqLines, lCurRow, kColReqUnitCostExact, CStr(fCurRowUnitCost)
                gGridUpdateCell grdReqLines, lCurRow, kColReqExtAmt, CStr(fCurRowExtAmt)
                gGridUpdateCell grdReqLines, lCurRow, kColReqEstPres, CStr(fCurRowExtAmt)
            End If
        Case kColReqExtAmt
            fCurRowReqQty = CDbl(gsGridReadCell(grdReqLines, lCurRow, kColReqQtyRequested))
            fCurRowExtAmt = CDbl(gsGridReadCell(grdReqLines, lCurRow, kColReqExtAmt))
            If fCurRowExtAmt < 0 Then
                giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeNegative, _
                            "Ext Amt"
                bCalcLineAmts = False
                '+++ VB/Rig Begin Pop +++
                '+++ VB/Rig End +++
                Exit Function
            End If
            fCurRowUnitCost = fCurRowExtAmt / fCurRowReqQty
            fCurRowUnitCostExact = fCurRowExtAmt / fCurRowReqQty
            If fCurRowUnitCost > kvMaxUnitCost Then
                giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeGreaterThan, _
                            "Unit Cost", kvMaxUnitCost
                bCalcLineAmts = False
                '+++ VB/Rig Begin Pop +++
                '+++ VB/Rig End +++
                Exit Function
            Else
                gGridUpdateCell grdReqLines, lCurRow, kColReqUnitCost, CStr(fCurRowUnitCost)
                gGridUpdateCell grdReqLines, lCurRow, kColReqUnitCostExact, CStr(fCurRowUnitCostExact) ' Updating UnitCostExact value
            End If
    End Select
    bCalcLineAmts = True

'+++ VB/Rig Begin Pop +++
        Exit Function
Resume
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bCalcLineAmts", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function


Private Function bLoseFocus(Optional NavCtl) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    bLoseFocus = True
    
    If Me.ActiveControl Is sbrMain Or Me.ActiveControl Is tbrMain Then
        bLoseFocus = False
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If
    
    If Not IsMissing(NavCtl) Then
        If Me.ActiveControl Is NavCtl Then
            bLoseFocus = False
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Function
        End If
    End If
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bLoseFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Function fGetTranAmt() As Double
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lLoop As Long
    Dim fTotal As Double
    Dim sExtAmt As String
    
    
    For lLoop = 1 To grdReqLines.MaxRows
        sExtAmt = gsGridReadCell(grdReqLines, lLoop, kColReqExtAmt)
        If Trim(sExtAmt) <> "" Then
            fTotal = fTotal + CDbl(sExtAmt)
        End If
    Next
    
    fGetTranAmt = fTotal
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "fGetTranAmt", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Function bFocusBack() As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
   
    bFocusBack = False
    
    '-- If using a navigator or toolbar don't shift focus back to textbox
'    If Me.ActiveControl Is navMain Or Me.ActiveControl Is navVendID Then
'        Exit Function
'    End If

    '-- Make sure we are not still in key section
    If Me.ActiveControl Is lkuMain Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If
    
    '-- If the tran no control is not enabled, do
    '-- not validate, it has already been done.
    If (lkuMain.Enabled = False Or lkuMain.Protected = True) Then Exit Function
    
    '-- If tran no is blank and type is not standard
    '-- do not lose focus
    If Len(Trim(lkuMain)) = 0 Then
        bFocusBack = True
    End If
    
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bFocusBack", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function


Private Sub LoadLineDflts(lRow As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
        
        Dim sRqstDate As String
        Dim lItemKey As Long
        Dim iLeepTime As Integer
        Dim lContractKey As Long
' This routine loads the required line defaults.
' Both the visible and hidden grids must have the same values defaulted.

        gGridUpdateCell grdReqLines, lRow, kColReqQtyRequested, 1
        gGridUpdateCell grdReqLineDtl, 1, kColChildQtyReq, 1
        gGridUpdateCell grdReqLines, lRow, kColReqUnitCost, 0
        gGridUpdateCell grdReqLines, lRow, kColReqUnitCostExact, 0
        gGridUpdateCell grdReqLines, lRow, kColReqExtAmt, 0
        gGridUpdateCell grdReqLineDtl, 1, kColChildFrtAmt, 0
        
        'Agregado por MultiConsulting Osmel Barreras
        gGridUpdateCell grdReqLines, lRow, kColReqEstPres, 0
        
        If sddEstatusReq.ItemData = kReqStatusAccepted Then
            gGridLockColumn grdReqLines, kColReqQtyRequested
            gGridLockColumn grdReqLines, kColReqEstPres
            gGridUpdateCell grdReqLines, lRow, kColReqQtyRequested, msPersistCantReq
            gGridUpdateCell grdReqLineDtl, lRow, kColChildQtyReq, msPersistCantReq
            gGridUpdateCell grdReqLines, lRow, kColReqEstPres, msPersistPresEstReq
        End If
        'Agregado por MultiConsulting Osmel Barreras
        
' if selecting a Dept is what caused the line to be created, don't overwrite it.
        If Trim(gsGridReadCell(grdReqLines, lRow, kColReqPurchDeptID)) = "" And _
Trim(lkuDept.Text) <> "" Then
            gGridUpdateCell grdReqLineDtl, 1, kColChildPurchDeptID, lkuDept.Text
            gGridUpdateCell grdReqLines, lRow, kColReqPurchDeptID, lkuDept.Text
            gGridUpdateCell grdReqLineDtl, 1, kColChildPurchDeptKey, _
                             moDmForm.GetColumnValue("DfltPurchDeptKey")
            gGridUpdateCell grdReqLines, lRow, kColReqPurchDeptKey, _
                            moDmForm.GetColumnValue("DfltPurchDeptKey")
            gGridLockCell grdReqLines, kColReqWhseID, lRow

        End If
' if selecting a Warehouse is what caused the line to be created, don't overwrite it.
        If Trim(gsGridReadCell(grdReqLines, lRow, kColReqWhseID)) = "" And _
        Trim(lkuWarehouse.Text) <> "" Then
            gGridUpdateCell grdReqLineDtl, 1, kColChildWhseID, lkuWarehouse.Text
            gGridUpdateCell grdReqLines, lRow, kColReqWhseID, lkuWarehouse.Text
            gGridUpdateCell grdReqLineDtl, 1, kColChildWhseKey, _
            moDmForm.GetColumnValue("DfltShipToWhseKey")
            gGridUpdateCell grdReqLines, lRow, kColReqWhseKey, _
            moDmForm.GetColumnValue("DfltShipToWhseKey")
            gGridLockCell grdReqLines, kColReqPurchDeptID, lRow
        End If
        'Modificado por Multiconsulting
        lItemKey = glGetValidLong(gsGridReadCell(grdReqLines, lRow, kColReqItemKey))
        If mbIntegratedCT And lItemKey <> 0 Then
            lContractKey = lGetReqContract
            If lContractKey > 0 Then
                iLeepTime = giGetValidInt(moClass.moAppDB.Lookup("s.DeliveryTime", "tctContract AS p JOIN tctContractLine AS s ON s.ContractKey = p.ContractKey", "s.ItemKey =" & lItemKey & " and p.ContractKey =" & lContractKey))
                sRqstDate = sGetDate(DateAdd("d", iLeepTime, calDate.Value))
                gGridUpdateCell grdReqLines, lRow, kcolReqRequestDate, sRqstDate
                gGridUpdateCell grdReqLineDtl, lRow, kColChildRequestDate, sRqstDate
            End If
        Else
            If lRow > 1 Then
                sRqstDate = gsGridReadCell(grdReqLines, lRow - 1, kcolReqRequestDate)
                gGridUpdateCell grdReqLineDtl, 1, kColChildRequestDate, sRqstDate
                gGridUpdateCell grdReqLines, lRow, kcolReqRequestDate, sRqstDate
            End If
        End If
        'Modificado por Multiconsulting
        
        gGridUpdateCell grdReqLines, lRow, kColReqLineKey, lGetNextSurrogateKey(moClass.moAppDB, "tpoReqLine")

    
'Set the child row dirty so that it will get saved.
        moDMSubGrid.SetRowDirty 1
        

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "LoadLineDflts", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
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


Private Sub LoadDflts()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++

'This routine should only be called for new rows.
'If this routine is needed for any other reason, the LoadLineDfts call
'will need to be modified to ensure that it does not overwrite important data.

On Error GoTo ExpectedErrorRoutine

    Dim bOldMb      As Boolean
    Dim lKey        As Long

    bOldMb = mbDontChkclick

    
    mbDontChkclick = True
    moDmForm.SetColumnValue "Status", kvReqOpen ' default the status to Open
    DisplayStatus
    mbDontChkclick = bOldMb
    moDmForm.SetColumnValue "CompanyID", msCompanyID
    moDmForm.SetColumnValue "DfltTargetCompID", msCompanyID
    moDmForm.SetColumnValue "TranType", kTranTypePORQ
    calDate.Text = Format(msBusinessDate, gsGetLocalVBDateMask())
'    moDMForm.SetColumnValue "Contact", msCurrentUser
'    moDMForm.SetColumnValue "Originator", msCurrentUser
    txtContact = msCurrentUser
    txtOriginator = msCurrentUser
    
    gGridSetCellType grdReqLines, 1, kcolReqRequestDate, SS_CELL_TYPE_DATE
    
'   Default the values for the first row.  This is needed because DMGridAppend
'   is not called as part of this process.
    LoadLineDflts 1

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Sub

ExpectedErrorRoutine:
mbDontChkclick = bOldMb
MyErrMsg moClass, Err.Description, Err, sMyName, "LoadDflts"
gClearSotaErr

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Sub
    
'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "LoadDflts", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub RemoveLastRowFromGrid()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    'Modificado por Multiconsulting
    If glGetValidLong(moClass.moAppDB.Lookup("count(1)", "tpoReqLine", "ReqKey = " & glGetValidLong(moDmForm.GetColumnValue("ReqKey")))) < grdReqLines.MaxRows Then
        gGridDeleteRow grdReqLines, grdReqLines.MaxRows
    End If
    'Modificado por Multiconsulting
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "RemoveLastRowFromGrid", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub DisplayStatus()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
On Error GoTo ExpectedErrorRoutine

Dim iStatus As Integer
Dim rs As Object
Dim sSql As String
Dim bNotClosed As Boolean
Dim lRow As Long


    iStatus = moDmForm.GetColumnValue("Status")
    sSql = "SELECT tsmLocalString.LocalText FROM tsmLocalString, tsmListValidation "
    sSql = sSql & " WHERE tsmListValidation.TableName = 'tpoRequisition'"
    sSql = sSql & " AND tsmListValidation.ColumnName = 'Status'"
    sSql = sSql & " AND tsmListValidation.DBValue = " & iStatus
    sSql = sSql & " AND tsmListValidation.StringNo = tsmLocalString.StringNo"
    sSql = sSql & " AND tsmLocalString.LanguageID = " & mlLanguage

    
    
    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
    If rs.IsEOF Then
        
        Set rs = Nothing
        txtStatus.Text = ""
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Sub
    
    Else
        txtStatus.Text = rs.Field("LocalText") '**************************************************************+
        
        Set rs = Nothing

    End If
' If the Requisition is closed, disable all fields in the header otherwise, enable them
    bNotClosed = Not (iStatus = kvReqClosed)
    mbReqClosed = Not bNotClosed
    moGM.MenuAdd = bNotClosed
    moGM.MenuDelete = bNotClosed
    txtOriginator.Enabled = bNotClosed
    calDate.Enabled = bNotClosed
    txtContact.Enabled = bNotClosed
    chkExpedite.Enabled = bNotClosed
    cboExpReason.Enabled = (chkExpedite = 1) And bNotClosed
    lkuDept.Enabled = bNotClosed
    If mbIntegrateWithIM Then
        lkuWarehouse.Enabled = bNotClosed
    End If '***************************************************************************************************+++
    txtComment.Enabled = bNotClosed
    cmdGenerate.Enabled = bNotClosed
    cmdUserFlds.Enabled = bNotClosed
' If the req is closed, remove the row automatically appended to the end of the grid
    If mbReqClosed Then
        RemoveLastRowFromGrid
    End If
    
' If the status is not closed change the back color to window
    If bNotClosed Then
        txtOriginator.BackColor = vbWindowBackground
        txtContact.BackColor = vbWindowBackground
        txtComment.BackColor = vbWindowBackground
    Else
' Otherwise, change the back color to buttonface
        txtOriginator.BackColor = vbButtonFace
        txtContact.BackColor = vbButtonFace
        txtComment.BackColor = vbButtonFace
    End If
    If chkExpedite.Enabled And chkExpedite.Value = 1 Then
        cboExpReason.BackColor = vbWindowBackground
    Else
        cboExpReason.BackColor = vbButtonFace
    End If
   'MsgBox ""

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Sub

ExpectedErrorRoutine:
MyErrMsg moClass, Err.Description, Err, sMyName, "DisplayStatus"
gClearSotaErr

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Sub
'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DisplayStatus", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Public Sub DMGridRowLoaded(oDm As Object, lRow As Long)
'+++ VB/Rig Begin Push +++*******************************************************************************aqui
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim dOpenQty    As Double
    Dim sDateString As String
    Dim iStatus     As Integer
    Dim sTranNo     As String
    Dim bAllowCostOvrd As Boolean
    Dim bFoundSTaxClass As Boolean
    Dim bIsCommentOnly As Boolean
        
        
    ' The Request date is originally made into a char column to accomodate the loading
    ' of data from the temp table.  It is dynamically changed after loading.
    If oDm Is moDmGrid Then 'segundo

        'Defect #17870: Get the request date from #tpoReqLineJoin, since the original
        '               method no longer works (produces an empty string).
        'sDateString = Trim(gsGridReadCellText(grdReqLines, lRow, kcolReqRequestDate))
        sDateString = Trim(gsGetValidStr(moClass.moAppDB.Lookup("RequestDate", "#tpoReqLineJoin", _
                    "ReqLineKey=" _
                    & gsGridReadCellText(grdReqLines, lRow, kColReqLineKey))))


        gGridSetCellType grdReqLines, lRow, kcolReqRequestDate, SS_CELL_TYPE_DATE
        Debug.Print Now() & ": Request Date (" & lRow & ") = " & sDateString
        gGridUpdateCell grdReqLines, lRow, kcolReqRequestDate, sDateString
        'moDmGrid.SetDirty True
        
'*********************************************************************************************************************************************************************
'Modificado por MultiConsulting Cuba - Osmel Barreras

        If gsGridReadCell(grdReqLines, lRow, kColReqLStItemKey) <> "" Then
            gGridUpdateCell grdReqLines, lRow, kColReqLStItemID, gsGridReadCell(grdReqLines, lRow, kColReqLStItemKey)
        End If
        
        If gsGridReadCell(grdReqLines, lRow, kColReqLTBItemKey) <> "" Then
            gGridUpdateCell grdReqLines, lRow, kColReqLTBItemID, gsGridReadCell(grdReqLines, lRow, kColReqLTBItemKey)
        End If
        
        'ComprobarGridRow (lRow)
        
'Modificado por MultiConsulting Cuba - Osmel Barreras
'*********************************************************************************************************************************************************************
        
        
'        If the current row is associated with a PO line, lock the entire row
        If gsGridReadCellText(grdReqLines, lRow, kColReqPOTranID) <> "" Then
            gGridLockRow grdReqLines, lRow
        Else
            GetItemInfo gsGridReadCellText(grdReqLines, lRow, kColReqItemKey), bFoundSTaxClass, bIsCommentOnly
            If bIsCommentOnly Then
                gGridLockCell grdReqLines, kColReqQtyRequested, lRow
                gGridLockCell grdReqLines, kColReqUnitCost, lRow
                gGridLockCell grdReqLines, kColReqExtAmt, lRow
            End If
            If bFoundSTaxClass Then
                gGridLockCell grdReqLines, kColReqSTaxClassID, lRow
            End If
            If Trim(gsGridReadCellText(grdReqLines, lRow, kColReqPurchDeptID)) <> "" Then
                gGridLockCell grdReqLines, kColReqWhseID, lRow
            ElseIf Trim(gsGridReadCellText(grdReqLines, lRow, kColReqWhseID)) <> "" Then
                gGridLockCell grdReqLines, kColReqPurchDeptID, lRow
            End If
        End If
    End If
    
'+++ VB/Rig Begin Pop +++
        Exit Sub
Resume
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DMGridRowLoaded", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Public Sub DMGridAppend(oDm As Object, lRow As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    If oDm Is moDmGrid Then
    ' There is a problem with this when the row is appended by selecting a floating
    ' navigator.  If this gets fixed, the overriding of the row should be changed.
        lRow = grdReqLines.Row
        
        moDMSubGrid.AppendRow
    ' The Request date is originally made into a char column to accomodate the loading
    ' of data from the temp table.  When creating a new row, default it to date type.
        gGridSetCellType grdReqLines, lRow, kcolReqRequestDate, SS_CELL_TYPE_DATE
        LoadLineDflts (lRow)
    End If
    

    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DMGridAppend", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Public Function DMGridBeforeInsert(oDm As Object, lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If oDm Is moDmGrid Then

'    Fill in default values for items not available - possibly future use
        oDm.SetColumnValue lRow, "TargetCompanyID", msCompanyID
        oDm.SetColumnValue lRow, "Status", kvDfltLineStatus

    End If
    
    'Agregado por multiconsulting para corregir error de insert en la tabla
    'Revisar
'    If oDm Is moDMSubGrid Then
'        oDm.SetColumnValue lRow, "ReqLineDistKey", glGetNextSurrogateKey(moClass.moAppDB, "tpoReqLineDist")
'    End If
    'Agregado por multiconsulting
    
    DMGridBeforeInsert = True

'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DMGridBeforeInsert", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function


Public Function DMPreSave(oDm As Object)
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
On Error GoTo ExpectedErrorRoutine

    Dim sSql    As String
    Dim lKey    As Long
    Dim lRev    As Long
    Dim fAmt    As Double
    Dim sErr    As String
    Dim lErr    As Long
    'Agregado por Multiconsulting
    Dim msReason As String
    Dim lReasonCodeKey As Long
    Dim bCancel As Boolean
    Dim lChngNo As Long
    'Agregado por Multiconsulting
    
    If Len(Trim(msTranTypeID)) = 0 Then
        sSql = "CompanyID = " & gsQuoted(msCompanyID) & " AND TranType = " & kTranTypePORQ
        
        msTranTypeID = Trim(gsGetValidStr(moClass.moAppDB.Lookup("TranTypeID", _
"tciTranTypeCompany", sSql)))
    End If
        
    moDmForm.SetColumnValue "TranID", msTranTypeID & "-" & Trim(lkuMain.Text)
    fAmt = fGetTranAmt
    moDmForm.SetColumnValue "TranAmt", CStr(fAmt)
    moDmForm.SetColumnValue "TranAmtHC", CStr(fAmt)

 'Check for validations before save.
            If Not bIsValidDirtyCheck Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
                Exit Function
            End If
    'Agregado por Multiconsulting
    If moDmForm.State = kDmStateEdit Then
        'frmChngOrd.Init moClass
        lChngNo = glGetValidLong(moClass.moAppDB.Lookup("count(*)", "tpoRequisitionChngOrd", "ReqKey =" & moDmForm.GetColumnValue("ReqKey")))
        frmChngOrd.CurrentNumber = lChngNo
        frmChngOrd.ShowMe msReason, bCancel ',  True, lReasonCodeKey
        If bCancel Then
            Exit Function
        End If
        moClass.moAppDB.ExecuteSQL "insert into tpoRequisitionChngOrd select *, " & gsQuoted(msReason) & ", " & lReasonCodeKey & ", " & (lChngNo + 1) & " from tpoRequisition where ReqKey =" & moDmForm.GetColumnValue("ReqKey")
    End If
    'Agregado por Multiconsulting
    
'   If all of the lines on the req are associated with a PO, close the req.
    CheckStatus
    
    DMPreSave = True

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Function

ExpectedErrorRoutine:
sErr = Err.Description
lErr = Err
moDmForm.CancelAction
MyErrMsg moClass, sErr, lErr, sMyName, "DMPreSave"
gClearSotaErr

'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DMPreSave", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

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

Private Sub BindGM()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Set moGM = New clsGridMgr

    With moGM
        Set .Grid = grdReqLines
        Set .Form = frmRequistn
        Set .DM = moDmGrid
' Set the grid type to data sheet so that the context menus work properly.
        .GridType = kGridDataSheet
        .GridSortEnabled = True
        Set moGridNav5 = .BindColumn(kColReqSTaxClassID, navSTaxGrid)
        Set moGridNav5.ReturnControl = txtNavReturn
        Set moGridNav4 = .BindColumn(kColReqWhseID, navWhseGrid)
        Set moGridNav4.ReturnControl = txtNavReturn
        Set moGridNav3 = .BindColumn(kColReqPurchDeptID, navDeptGrid)
        Set moGridNav3.ReturnControl = txtNavReturn
        Set moGridNav2 = .BindColumn(kcolReqVendID, navVendorGrid)
        Set moGridNav2.ReturnControl = txtNavReturn
        Set moGridNav = .BindColumn(kColReqItemID, navItemGrid)
        Set moGridNav.ReturnControl = txtNavReturn
        'Agregado por MultiConsulting Osmel Barreras
        Set moGridNav6 = .BindColumn(kColReqBuyerID, navBuyerGrid)
        Set moGridNav6.ReturnControl = txtNavReturn
        'Agregado por MultiConsulting Osmel Barreras
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

'Agregado por Multiconsulting
Private Sub cmdContract_Click()
Dim lContractKey As Long
Dim bValid As Boolean
Dim bUpdateAble As Boolean
    
    On Error GoTo ErrorHandler
    If moDmForm.State = kDmStateEdit Or moDmForm.State = kDmStateAdd Then
        bUpdateAble = bAllowContractChange
        lContractKey = lGetReqContract
        
        frmContractAssociate.lkuContract.Enabled = bUpdateAble
        frmContractAssociate.lkuSuplement.Enabled = bUpdateAble
        frmContractAssociate.cmdOK.Enabled = bUpdateAble
        frmContractAssociate.ShowContract lContractKey, bValid
        If bValid Then
            SetReqContract lContractKey
            gGridLockColumn grdReqLines, kcolReqVendID
            BindNavigators
            If lContractKey > 0 Then
                gGridLockColumn grdReqLines, kColReqUnitMeasID
            Else
                gGridUnlockColumn grdReqLines, kColReqUnitMeasID
            End If
        End If
        If lContractKey > 0 Then
            grdReqLines.Enabled = True
            sddType.Enabled = False
        Else
            grdReqLines.Enabled = False
            sddType.Enabled = True
        End If
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub
'Agregado por Multiconsulting

Private Sub cmdUserFlds_Click()
'+++ VB/Rig Begin Push +++                                                                'Repository Error Rig  {1.1.1.0.0}
#If ERRORTRAPON Then                                                                      'Repository Error Rig  {1.1.1.0.0}
    On Error GoTo VBRigErrorRoutine                                                       'Repository Error Rig  {1.1.1.0.0}
#End If                                                                                   'Repository Error Rig  {1.1.1.0.0}
'+++ VB/Rig End +++                                                                       'Repository Error Rig  {1.1.1.0.0}
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If Not moFormCust.onClick(cmdUserFlds, True) Then Exit Sub
    End If
    #End If
'+++ End Customizer Code Push +++
    
    If bDontClick Then
        Exit Sub
    End If
        
    Dim bChangesMade As Boolean
    
    msUserFld(0) = gsGetValidStr(moDmForm.GetColumnValue("UserFld1"))
    msUserFld(1) = gsGetValidStr(moDmForm.GetColumnValue("UserFld2"))
    msUserFld(2) = gsGetValidStr(moDmForm.GetColumnValue("UserFld3"))
    msUserFld(3) = gsGetValidStr(moDmForm.GetColumnValue("UserFld4"))
    
    moUF.ShowUserflds kEntTypePOPurchOrder, 4, msUserFld(), False, bChangesMade
    
    moDmForm.SetColumnValue "UserFld1", msUserFld(0)
    moDmForm.SetColumnValue "UserFld2", msUserFld(1)
    moDmForm.SetColumnValue "UserFld3", msUserFld(2)
    moDmForm.SetColumnValue "UserFld4", msUserFld(3)

'+++ VB/Rig Begin Pop +++                                                                 'Repository Error Rig  {1.1.1.0.0}
    Exit Sub                                                                              'Repository Error Rig  {1.1.1.0.0}
VBRigErrorRoutine:                                                                        'Repository Error Rig  {1.1.1.0.0}
        gSetSotaErr Err, sMyName, "cmdUserflds_click", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Function bDontClick() As Boolean
'+++ VB/Rig Begin Push +++                                                                'Repository Error Rig  {1.1.1.0.0}
#If ERRORTRAPON Then                                                                      'Repository Error Rig  {1.1.1.0.0}
    On Error GoTo VBRigErrorRoutine                                                       'Repository Error Rig  {1.1.1.0.0}
#End If                                                                                   'Repository Error Rig  {1.1.1.0.0}
'+++ VB/Rig End +++                                                                       'Repository Error Rig  {1.1.1.0.0}
    
    If moDmForm.State = kStateNone Then
        bDontClick = True
    End If
    
    If Len(Trim(lkuMain.Text)) = 0 Then
        bDontClick = True
    End If
    
'+++ VB/Rig Begin Pop +++                                                                 'Repository Error Rig  {1.1.1.0.0}
    Exit Function                                                                         'Repository Error Rig  {1.1.1.0.0}
VBRigErrorRoutine:                                                                        'Repository Error Rig  {1.1.1.0.0}
        gSetSotaErr Err, sMyName, "bDontClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++                                                                       'Repository Error Rig  {1.1.1.0.0}
End Function



Private Sub Command1_Click()

lkuWarehouse.Text = "Nave 4"

End Sub

Private Sub grdReqLines_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    Dim bChangesMade As Boolean
    
    msUserFld_Line(0) = gsGetValidStr(moDmGrid.GetColumnValue(Row, "UserFld1"))
    msUserFld_Line(1) = gsGetValidStr(moDmGrid.GetColumnValue(Row, "UserFld2"))
    
    moUF.ShowUserflds kE
    ntTypePOPOLine , 2, msUserFld_Line(), False, bChangesMade
    
    moDmGrid.SetColumnValue Row, "UserFld1", msUserFld_Line(0)
    moDmGrid.SetColumnValue Row, "UserFld2", msUserFld_Line(1)
        
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_ButtonClicked", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub grdReqLines_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If moDmForm.State = kDmStateAdd And grdReqLines.Row = 1 Then
        If Trim(gsGetValidStr(moDmGrid.GetColumnValue(1, kColReqItemID))) = "" Then
             gGridSetActiveCell grdReqLines, 1, kColReqItemID
             'txtNavReturn.SetFocus
            'gGridUpdateCell grdReqLines, 1, kColReqItemID, ""
             
        End If
    End If
        
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_GotFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub moGM_EnterGridRow(ByVal lRow As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

Dim lItemKey As Long

     
    If moDmForm.State = kDmStateNone Then
        Exit Sub
    End If
    
     If lRow > 0 And lRow <> mlCurrRow Then
        lItemKey = glGetValidLong(gsGridReadCellText(grdReqLines, mlCurrRow, kColReqItemKey))
        If lItemKey <> 0 Then
           msUserFld_Line(0) = gsGetValidStr(moDmGrid.GetColumnValue(mlCurrRow, "UserFld1"))
           msUserFld_Line(1) = gsGetValidStr(moDmGrid.GetColumnValue(mlCurrRow, "UserFld2"))
    
            'Call comment engine to accept user fields for PO level and PO line level before processing.
            If Not moUF.bValidateUserflds(kEntTypePOPOLine, 2, True, msUserFld_Line()) Then
                Exit Sub
            End If
    
            moDmGrid.SetColumnValue mlCurrRow, "UserFld1", msUserFld_Line(0)
            moDmGrid.SetColumnValue mlCurrRow, "UserFld2", msUserFld_Line(1)
    
        End If

        mlCurrRow = lRow
        msOldItemID = Trim(gsGridReadCellText(grdReqLines, lRow, kColReqItemID))
        lItemKey = glGetValidLong(gsGridReadCellText(grdReqLines, lRow, kColReqItemKey))
        Set moItem = moIMSClass.Items(lItemKey)

        msOldVendID = Trim(gsGridReadCellText(grdReqLines, lRow, kcolReqVendID))
        msOldSTaxClassID = Trim(gsGridReadCellText(grdReqLines, lRow, kColReqSTaxClassID))
        msOldPurchDeptID = Trim(gsGridReadCellText(grdReqLines, lRow, kColReqPurchDeptID))
        'If Not mbSkipGotFocus Then
            msOldWhseID = Trim(gsGridReadCellText(grdReqLines, lRow, kColReqWhseID))
        'End If
        'Agregado por MultiConsulting Osmel Barreras
        msOldBuyerID = Trim(gsGridReadCellText(grdReqLines, lRow, kColReqBuyerID))
        msOldStateItemID = Trim(gsGridReadCellText(grdReqLines, lRow, kColReqLStItemID))
        msOldTypeBuyItemID = Trim(gsGridReadCellText(grdReqLines, lRow, kColReqLTBItemID))
        'Agregado por MultiConsulting Osmel Barreras
    End If
    


'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moGM_EnterGridROw", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub moGM_OverrideMenu(MenuAdd As Boolean, MenuDelete As Boolean, MenuDrillDown As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If giGetValidInt(moDmForm.GetColumnValue("Status")) = 5 Then   'status 5-closed
        MenuAdd = False
        MenuDelete = False
        gGridClearSelectRow grdReqLines, grdReqLines.Row
    End If
    If gsGridReadCellText(grdReqLines, grdReqLines.Row, kColReqPOTranID) <> "" Then
        MenuDelete = False
        gGridClearSelectRow grdReqLines, grdReqLines.Row
    End If
    'gGridClearSelectRow grdReqLines, grdReqLines.Row
'+++ VB/Rig Begin Pop +++
Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "GM_OverrideMenu", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

'************************************************************************
'   Description:
'       Setup the lookup fields for the form, including restricting the
'       lookups to only display data for the current company.
'
'   Param:
'       <none>
'
'************************************************************************

Private Sub SetupLookups()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
  'Set the Requisition Number Navigator info
    With lkuMain
        Set .Framework = moClass.moFramework
        Set .SysDB = moClass.moAppDB
        Set .AppDatabase = moClass.moAppDB
        .RestrictClause = "tpoRequisition.CompanyID = " & gsQuoted(msCompanyID)
    End With

  'Set the Department Navigator info
    With lkuDept
        Set .Framework = moClass.moFramework
        Set .SysDB = moClass.moAppDB
        Set .AppDatabase = moClass.moAppDB
        .RestrictClause = "tpoPurchDepartment.CompanyID = " & gsQuoted(msCompanyID)
    End With
    
  'Set the Warehouse Navigator info
    With lkuWarehouse
        Set .Framework = moClass.moFramework
        Set .SysDB = moClass.moAppDB
        Set .AppDatabase = moClass.moAppDB
        .RestrictClause = "CompanyID = " & gsQuoted(msCompanyID) & " AND Transit = 0"
    End With

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "SetupLookups", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub



Public Function lStartErrorReport(Optional iFileType As Variant, Optional sFileName As Variant) As Long
    
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
Dim sWhereClause As String
Dim sTablesUsed As String
Dim sSelect As String
Dim sInsert As String
Dim lRetVal As Long
Dim bValid As Boolean
Dim iNumTablesUsed As Integer
Dim RptFileName As String
Dim lBadRow As Long
Dim lErrBatchKey As Long
Dim sSql As String
Dim rs As Object
'Dim moReportObj As clsReportEngine
'Dim moDBObj As Object
'Dim moDDData As clsDDData
'Dim moRealTableCollection As Collection

    On Error GoTo badexit

    lStartErrorReport = kFailure
    
'    ShowStatusBusy frm

    'Set SelectObj = frm.moSelect
    Set moReportObj = New clsReportEngine
    Set moDBObj = moClass.moAppDB
'    lErrBatchKey = moClass.lKey
    
    moReportObj.UI = False
    moReportObj.AppOrSysDB = kAppDB
 
    RptFileName = "pozde001.rpt"
    Set moRealTableCollection = New Collection
    With moRealTableCollection
        .Add "tciErrorLog" 'the "Driving Table" name
    End With
    Set moDDData = New clsDDData
    If Not moDDData.lInitDDData(moRealTableCollection, moClass.moAppDB, moClass.moAppDB, moClass.moSysSession.CompanyId) = kSuccess Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If

     If (moReportObj.lInitReport("PO", RptFileName, frmSelectReqLines, moDDData) = kFailure) Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If

   
    
    On Error GoTo badexit
    
    '*************** NOTE ********************
    'THE ORDER OF THE FOLLOWING EVENTS IS IMPORTANT!
    
    'CUSTOMIZE:  The .RPT file to be used should be set here.  (More than one .RPT file
    'may exist for the task.)
    moReportObj.ReportFileName() = RptFileName
    
    'Start Crystal print engine, open a print job, and get localized strings from
    'tsmLocalString table.
    If (moReportObj.lSetupReport = kFailure) Then
        GoTo badexit
    End If
    
    'work around if you print without previewing first.
    'Crystal does not provide a way of getting page orientation
    'used to create report. use VB constants:
    'vbPRORPortrait, vbPRORLandscape
    moReportObj.Orientation() = vbPRORPortrait
    
    'CUSTOMIZE:  Set report titles to localized text from tsmLocalString table using call
    'to gsBuildString with a VB constant defined in StrConst.bas. The subtitles should
    'not include the format selected by the user, i.e., "Detail" or "Summary".
    moReportObj.ReportTitle1() = "Error Log Listing" 'gsBuildString(kVendClassListing, frm.oClass.moAppDB, frm.oClass.moSysSession)
    moReportObj.ReportTitle2() = ""
    
    'CUSTOMIZE:  Include these calls if you have named subtotal & header labels on the report
    'using the "lbl" convention on formula field names so that label text will handled automatically
    moReportObj.UseSubTotalCaptions() = 1
    moReportObj.UseHeaderCaptions() = 1
    'Supress the summary section
    moReportObj.lSetSummarySection 0, Nothing
    
    'set standard formulas, business date, run time, company name etc.
    'as defined in the template
    If (moReportObj.lSetStandardFormulas(frmSelectReqLines) = kFailure) Then
        GoTo badexit
    End If

    'Set sort order in .RPT file according to user selections in the Sort grid.
    'If (moReportObj.lSetSortCriteria(frm.moSort) = kFailure) Then
        'GoTo badexit
    'End If
    
    '********* the following is specific to your report *************'
    'Select Case RptFileName
        'Case "XXZYY001.RPT"
        'Case Else
    'End Select
    '********* End of special processing *************'
    
    'Retrieve the SQL statement stored with the .RPT file and modify it as needed.
    moReportObj.BuildSQL
    moReportObj.SetSQL

    'CUSTOMIZE:  Include this call if you have named column labels on the report
    'using the "lbl" convention and wish label text to be handled automatically for you.
    moReportObj.SetReportCaptions
    
    'used in the Summary section on the report: use kLenPortrait or kLenLandscape
    'moReportObj.SelectString = SelectObj.sGetUserReadableWhereClause(kLenPortrait)
    
    'CUSTOMIZE:  If using work tables, restrict report data to current Session ID.  If using
    'real tables, might restrict report data to current company or other criteria.
    If (moReportObj.lRestrictBy("{tciErrorLog.SessionID} = " & mlSession & " AND {tciErrorLog.Severity} > 0") = kFailure) Then
        GoTo badexit
    End If
    
    moReportObj.ProcessReport frmRequistn, kTbPreview, iFileType, sFileName
            
'    ShowStatusNone frmSelectReqLines
    
    lStartErrorReport = kSuccess
    
'    Set moReportObj = Nothing
'    Set moDBObj = Nothing
'    Set moDDData = Nothing
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Function
    
badexit:
    'moReportObj.CleanupWorkTables
    'Set SelectObj = Nothing
'    Set moReportObj = Nothing
'    Set moDBObj = Nothing
'    Set moDDData = Nothing
    gClearSotaErr
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Function
'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lStartReport", VBRIG_IS_MODULE
        Err.Raise guSotaErr.Number
'+++ VB/Rig End +++
End Function


'************************************************************************
'   Description:
'       create data manager objects and bind controls to
'       the appropriate fields on the form
'
'   Param:
'       <none>
'
'   Returns:
'
'************************************************************************

Private Sub BindForm()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    Set moDmForm = New clsDmForm
    Set moDmGrid = New clsDmGrid
    Set moDMSubGrid = New clsDmGrid
    
    With moDmForm
        Set .Form = frmRequistn
        Set .Session = moClass.moSysSession
        .AppName = Me.Caption
        Set .Database = moClass.moAppDB
        .Where = "CompanyID = '" + msCompanyID + "'"
        .Table = "tpoRequisition"
        .UniqueKey = "ReqKey"
        Set .SOTAStatusBar = sbrMain
        .SaveOrder = 1
        .Bind Nothing, "ReqKey", SQL_INTEGER
        .Bind Nothing, "AirFrtJustified", SQL_SMALLINT
        .Bind Nothing, "ApprovalDate", SQL_DATE, kDmSetNull
        .Bind Nothing, "ApprovalStatus", SQL_SMALLINT
        .Bind Nothing, "BuyerKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "CompanyID", SQL_CHAR
        .Bind txtContact, "Contact", SQL_VARCHAR, kDmSetNull
        .Bind Nothing, "CreateType", SQL_SMALLINT
        .Bind Nothing, "CurrExchRate", SQL_DECIMAL
        .Bind Nothing, "CurrExchSchdKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "CurrID", SQL_CHAR, kDmSetNull
        .Bind Nothing, "DfltPurchDeptKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "DfltRequestDate", SQL_DATE, kDmSetNull
        .Bind Nothing, "DfltShipMethKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "DfltShipToAddrKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "DfltShipToWhseKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "DfltShipZoneKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "DfltTargetCompID", SQL_CHAR
        .Bind chkExpedite, "Expedite", SQL_SMALLINT
        .BindComboBox cboExpReason, "ExpediteReasonKey", SQL_INTEGER, kDmSetNull Or kDmUseItemData
        .Bind Nothing, "FOBKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "FirstPOIssueDate", SQL_DATE, kDmSetNull
        .Bind Nothing, "FreightAmt", SQL_DECIMAL
        .Bind Nothing, "Hold", SQL_SMALLINT
        .Bind Nothing, "HoldReason", SQL_CHAR, kDmSetNull
        .Bind txtOriginator, "Originator", SQL_CHAR
        .Bind Nothing, "PmtTermsKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "Printed", SQL_SMALLINT
        .Bind Nothing, "PurchAddrKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "PurchAmt", SQL_DECIMAL
        .Bind Nothing, "PurchVendAddrKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "RemitToAddrKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "RemitToVendAddrKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "ReqFormKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "STaxAmt", SQL_DECIMAL
        .Bind Nothing, "STaxTranKey", SQL_INTEGER, kDmSetNull
        .Bind Nothing, "Status", SQL_SMALLINT
        .Bind Nothing, "TranAmt", SQL_DECIMAL
        .Bind Nothing, "TranAmtHC", SQL_DECIMAL
        .Bind txtComment, "TranCmnt", SQL_VARCHAR
        .Bind calDate, "TranDate", SQL_DATE
        .Bind Nothing, "TranID", SQL_CHAR
        .Bind lkuMain, "TranNo", SQL_CHAR
        .Bind Nothing, "TranType", SQL_INTEGER
        .Bind Nothing, "UpdateCounter", SQL_INTEGER
        .Bind Nothing, "UsedFor", SQL_VARCHAR, kDmSetNull
        .Bind Nothing, "UserFld1", SQL_CHAR, kDmSetNull 'User fields addition
        .Bind Nothing, "UserFld2", SQL_CHAR, kDmSetNull
        .Bind Nothing, "UserFld3", SQL_CHAR, kDmSetNull
        .Bind Nothing, "UserFld4", SQL_CHAR, kDmSetNull
        .Bind Nothing, "V1099Box", SQL_CHAR, kDmSetNull
        .Bind Nothing, "V1099BoxText", SQL_CHAR, kDmSetNull
        .Bind Nothing, "V1099Form", SQL_SMALLINT, kDmSetNull
        
        .LinkSource "tpoPurchDepartment", "tpoPurchDepartment.PurchDeptKey=<<DfltPurchDeptKey>>"
        .Link lkuDept, "PurchDeptID"

        .LinkSource "timWarehouse", "timWarehouse.WhseKey=<<DfltShipToWhseKey>>"
        .Link lkuWarehouse, "WhseID"

        .Init
    
    End With

    'Agregado por Multiconsulting
    Set moDmReqAdicInfo = New clsDmForm
    With moDmReqAdicInfo
        Set .Form = frmRequistn
        Set .Session = moClass.moSysSession
        .AppName = Me.Caption
        Set .Database = moClass.moAppDB
        Set .Parent = moDmForm
        .Table = "tpoReqAdicInfo"
        .UniqueKey = "ReqKeyIA"
        Set .SOTAStatusBar = sbrMain
        .SaveOrder = 2
        .ParentLink "ReqKeyIA", "ReqKey", SQL_INTEGER
        
        .Bind sddEstatusReq, "statusKey", SQL_INTEGER, kDmUseItemData
        .Bind txtAutorizaReq, "autorizaReq", SQL_VARCHAR
        .Bind calAutorizaReq, "dateAutorReq", SQL_DATE
        .Bind chkb2doAuthNeed, "segLvlAutoriza", SQL_BIT
        .Bind txt2doAutorizaReq, "segAutorizaReq", SQL_CHAR
        .Bind cal2doAutorizaReq, "segDateAutorReq", SQL_DATE
        .Bind txtAceptaReq, "aceptaReq", SQL_CHAR
        .Bind calAceptaReq, "dateAceptReq", SQL_DATE
        .Bind txtEstatusReqDesc, "descriptionStatus", SQL_CHAR
        .Bind sddType, "Type", SQL_INTEGER, kDmUseItemData
        .Init
    End With
    'Agregado por Multiconsulting


    With moDmGrid
        Set .Form = frmRequistn
        Set .Session = moClass.moSysSession
        Set .Grid = grdReqLines
        Set .Parent = moDmForm
        Set .Database = moClass.moAppDB
        .Table = "tpoReqLine"
        .UniqueKey = "ReqLineKey"
'        .OrderBy = "ReqLineKey"
        .SaveOrder = 3 'Modificado por Multiconsulting

        .BindColumn "ReqLineKey", kColReqLineKey, SQL_INTEGER
        .BindColumn "CmntOnly", Nothing, SQL_SMALLINT
        .BindColumn "Description", kColReqDescription, SQL_VARCHAR
        .BindColumn "ExtAmt", kColReqExtAmt, SQL_DECIMAL, , kDmSetNull
        .BindColumn "ExtCmnt", kColReqComment, SQL_VARCHAR, , kDmSetNull
        .BindColumn "ItemKey", kColReqItemKey, SQL_INTEGER, , kDmSetNull
        .BindColumn "STaxClassKey", kColReqSTaxClassKey, SQL_INTEGER, , kDmSetNull
        .BindColumn "POLineKey", kColReqPOLineKey, SQL_INTEGER, , kDmSetNull
        .BindColumn "Status", Nothing, SQL_SMALLINT
        .BindColumn "TargetCompanyID", Nothing, SQL_CHAR
        .BindColumn "UnitCost", kColReqUnitCost, SQL_DECIMAL
        .BindColumn "UnitMeasKey", kColReqUnitMeasKey, SQL_INTEGER, , kDmSetNull
        .BindColumn "UpdateCounter", Nothing, SQL_INTEGER
        .BindColumn "UserFld1", kColPOUserFld1, SQL_CHAR, , kDmSetNull
        .BindColumn "UserFld2", kColPOUserFld2, SQL_CHAR, , kDmSetNull
        .BindColumn "VendKey", kColReqVendKey, SQL_INTEGER, , kDmSetNull
        .BindColumn "UnitCostExact", kColReqUnitCostExact, SQL_DECIMAL
        '****************************************************************************************************
        'Agregado por MultiConsulting Osmel Barreras
        .BindColumn "EstimatedPres", kColReqEstPres, SQL_DECIMAL
        .BindColumn "ReqLineBuyerKey", kColReqBuyerkey, SQL_INTEGER, , kDmSetNull
        .BindColumn "StateBIKey", kColReqLStItemKey, SQL_INTEGER, , kDmSetNull
        .BindColumn "TypeBIKey", kColReqLTBItemKey, SQL_INTEGER, , kDmSetNull
        'Agregado por MultiConsulting Osmel Barreras
        '****************************************************************************************************
        .ParentLink "ReqKey", "ReqKey", SQL_INTEGER
        
        .LinkSource "timItem", "tpoReqLine.ItemKey=timItem.ItemKey", kDmJoin, LeftOuter
        .Link kColReqItemID, "ItemID"
        
        .LinkSource "tciSTaxClass", "tpoReqLine.STaxClassKey=tciSTaxClass.STaxClassKey", kDmJoin, LeftOuter
        .Link kColReqSTaxClassID, "STaxClassID"

        .LinkSource "tapVendor", "tpoReqLine.VendKey=tapVendor.VendKey", kDmJoin, LeftOuter
        .Link kcolReqVendID, "VendID"

        'Associate the Currency column to tapVendAddr.CurrID
        .LinkSource "tapVendAddr", "tapVendAddr.AddrKey = tapVendor.PrimaryAddrKey", kDmJoin, LeftOuter
        .Link kColReqCurrid, "CurrID"

'***************************************************************************************************************************
'Modificado por MultiConsulting Osmel Barreras Piñera
        .LinkSource "timBuyer", "tpoReqLine.ReqLineBuyerKey=timBuyer.BuyerKey", kDmJoin, LeftOuter
        .Link kColReqBuyerID, "BuyerID"
        
        .LinkSource "tpoReqStateItem", "tpoReqLine.StateBIKey=tpoReqStateItem.stateItemKey", kDmJoin, LeftOuter
        .Link kColReqLStItemID, "stateItemID"
        
        .LinkSource "tpoReqTypeBuyItem", "tpoReqLine.TypeBIKey=tpoReqTypeBuyItem.typeBIKey", kDmJoin, LeftOuter
        .Link kColReqLTBItemID, "typeBIID"
'Modificado por MultiConsulting Osmel Barreras Piñera
'***************************************************************************************************************************

        .LinkSource "tciUnitMeasure", "tpoReqLine.UnitMeasKey=tciUnitMeasure.UnitMeasKey", kDmJoin, LeftOuter
        .Link kColReqUnitMeasID, "UnitMeasID"
        .Link kColPOUOMType, "MeasType"

        .LinkSource "#tpoReqLineJoin", "tpoReqLine.ReqLineKey=#tpoReqLineJoin.ReqLineKey", kDmJoin, LeftOuter
        .Link kcolReqRequestDate, "RequestDate"
        .Link kColReqPurchDeptKey, "PurchDeptKey"
        .Link kColReqPurchDeptID, "PurchDeptID"
        .Link kColReqWhseKey, "ShipToWhseKey"
        .Link kColReqWhseID, "ShipToWhseID"
        .Link kColReqQtyRequested, "QtyReq"
        .Link kColReqPOLineKey, "POLineKey"
        .Link kColReqPOKey, "POKey"
        .Link kColReqPOTranID, "POTranNo"

        .Init
    End With
    
    With moDMSubGrid
        Set .Form = frmRequistn
        Set .Session = moClass.moSysSession
        Set .Grid = grdReqLineDtl
        Set .Parent = moDmGrid
        Set .Database = moClass.moAppDB
        .SaveOrder = 4 'Modificado por Multiconsulting
        .Table = "tpoReqLineDist"
        .UniqueKey = "tpoReqLineDist.ReqLineKey, ReqLineDistKey"
'        .ManualLoad = True
        .BindColumn "ReqLineDistKey", kColChildReqLineDistKey, SQL_INTEGER
        .BindColumn "PurchDeptKey", kColChildPurchDeptKey, SQL_INTEGER, , kDmSetNull
        .BindColumn "ShipToWhseKey", kColChildWhseKey, SQL_INTEGER, , kDmSetNull
        .BindColumn "RequestDate", kColChildRequestDate, SQL_DATE
        .BindColumn "QtyReq", kColChildQtyReq, SQL_DECIMAL
        .BindColumn "UpdateCounter", Nothing, SQL_INTEGER
        .BindColumn "FreightAmt", kColChildFrtAmt, SQL_DECIMAL
        .ParentLink "ReqLineKey", "ReqLineKey", SQL_INTEGER

        .LinkSource "tpoPurchDepartment", "tpoReqLineDist.PurchDeptKey=tpoPurchDepartment.PurchDeptKey", kDmJoin, LeftOuter
        .Link kColChildPurchDeptID, "PurchDeptID"
        
        .LinkSource "timWarehouse", "tpoReqLineDist.ShipToWhseKey=timWarehouse.WhseKey", kDmJoin, LeftOuter
        .Link kColChildWhseID, "WhseID"

        .Init

         .LinkSource "#tpoReqLineJoin", "tpoReqLineDist.ReqLineKey=#tpoReqLineJoin.ReqLineKey" & _
                                        " AND tpoReqLineDist.ReqLineDistKey=" & _
                                        "#tpoReqLineJoin.ReqLineDistKey", kDmJoin, LeftOuter

        .Init
    End With
    
        '-- Eventually the floating navigator should become part of the data manager grid object !
    '-- That's why I think it belongs in the BindForm
    

    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "BindForm", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub LoadHeaderLists()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    Dim sSQLSelect As String
        
    sSQLSelect = "SELECT ReasonCodeID, ReasonCodeKey FROM tciReasonCode"
    cboExpReason.InitDynamicList moClass.moAppDB, sSQLSelect, msCompanyID
    
    'Agregado por Multiconsulting
    sddEstatusReq.InitStaticList moClass.moAppDB, "tpoReqAdicInfo", "statusKey", mlLanguage
    sddType.InitStaticList moClass.moAppDB, "tpoReqAdicInfo", "Type", mlLanguage
    
    Set moStaticListStateItem = New clsStaticList
    moStaticListStateItem.InitStaticList moClass.moAppDB, mlLanguage, "tpoReqLine", "StateBIKey"
        
    Set moStaticListTypeBuyItem = New clsStaticList
    moStaticListTypeBuyItem.InitStaticList moClass.moAppDB, mlLanguage, "tpoReqLine", "TypeBIKey"
    'Agregado por Multiconsulting
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "LoadHeaderLists", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub BindNavigators()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim sRestrict As String
    Dim lContractKey As Long 'Agregado por Multiconsulting
    
    '-- Only active status of type Misc. or Expense.
    If mbIntegrateWithIM Then
        sRestrict = "CompanyID = " & gsQuoted(msCompanyID) & " AND Status = 1 AND ItemType IN (1,3,4,5,6,8,9)"
    Else
        sRestrict = "CompanyID = " & gsQuoted(msCompanyID) & " AND Status = 1 AND ItemType IN (1,3,4)"
    End If
    'Agregado por Multiconsulting
    If mbIntegratedCT Then
        lContractKey = lGetReqContract
        If lContractKey > 0 Then
            sRestrict = sRestrict & " and ItemKey in (" & sGetAvaliableItemsListFromContract(lContractKey) & ")"
        End If
    End If
    'Agregado por Multiconsulting
    gbLookupInit navItemGrid, moClass, moClass.moAppDB, "Item", sRestrict
    
    '-- Only Classes created via the maintenance
    sRestrict = "STaxCodeClass = 0"
    gbLookupInit navSTaxGrid, moClass, moClass.moAppDB, "STaxClass", sRestrict
    
    '-- Only Vendors for the current company
    gbLookupInit navVendorGrid, moClass, moClass.moAppDB, "Vendor", "CompanyID = " & gsQuoted(msCompanyID)
        
    ' Agregado por MultiConsulting Osmel Barreras
    '-- Only Buyers for the current company
    
    
   gbLookupInit navBuyerGrid, moClass, moClass.moAppDB, "Buyer", "CompanyID = " & gsQuoted(msCompanyID)
    ' Agregado por MultiConsulting Osmel Barreras
       
       'lkuWarehouse.Text = kColReqWhseID.Text
       
    '-- Only Departments for the current company
    gbLookupInit navDeptGrid, moClass, moClass.moAppDB, "POPurchDept", "tpoPurchDepartment.CompanyID = " & gsQuoted(msCompanyID)
        
    '-- Only Warehouses for the current company
    gbLookupInit navWhseGrid, moClass, moClass.moAppDB, "Warehouse", "CompanyID = " & gsQuoted(msCompanyID) & " AND Transit = 0"
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "BindNavigators", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub


Private Function CreateLineJoin(oDm As Object)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
Dim lKey    As Long
Dim sSql    As String
Static bdid As Boolean
Dim iRet As Integer

    
    If oDm Is moDmForm Then
        
        lKey = glGetValidLong(moDmForm.GetColumnValue("ReqKey"))
        
        If Not bdid Then
    
            sSql = "CREATE TABLE #tpoReqLineJoin "
            sSql = sSql & "(ReqLineDistKey      int       NULL, "
            sSql = sSql & "ReqLineKey           int       NULL, "
            sSql = sSql & "RequestDate          varChar(15)  NULL, "
            sSql = sSql & "QtyReq               decimal(16,8)  NULL, "
            sSql = sSql & "PurchDeptKey         int  NULL, "
            sSql = sSql & "PurchDeptID          varchar(15)  NULL, "
            sSql = sSql & "ShipToWhseKey        int  NULL, "
            sSql = sSql & "ShipToWhseID         varchar(6)  NULL, "
            sSql = sSql & "POLineKey            int  NULL, "
            sSql = sSql & "POKey                int  NULL, "
            sSql = sSql & "POTranNo             varchar(15)  NULL) "
            
            moClass.moAppDB.ExecuteSQL sSql
    
            bdid = True
        
        End If
        
        sSql = " EXEC sppoGetReqDetail " & Format(lKey) & ""
        
        moClass.moAppDB.ExecuteSQL (sSql)
    
    End If

'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CreateLineJoin", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function


Private Function sGetNextReqNo() As String
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    Dim sNextReqNo      As String
    Dim iNumChars       As Integer
    Dim iTranTypeToUse  As Integer
    Dim iRetVal         As Integer
      
On Error GoTo ExpectedErrorRoutine2
    
    iTranTypeToUse = kTranTypePORQ 'use the default tran Type
    
    If iNumChars = 0 Or iNumChars > 10 Then
        iNumChars = 10
    End If
    
    iRetVal = 0
    
    On Error GoTo ExpectedErrorRoutine
    With moClass.moAppDB
        .SetInParam msCompanyID
        .SetInParam iNumChars
        .SetInParam iTranTypeToUse
        .SetOutParam sNextReqNo
        .SetOutParam iRetVal
        .ExecuteSP ("sppoGetNextReqNo")
        sNextReqNo = .GetOutParam(4)
        iRetVal = .GetOutParam(5)
        .ReleaseParams
    End With
    On Error GoTo ExpectedErrorRoutine2
    
    If iRetVal = 0 Then
        sNextReqNo = ""
        giSotaMsgBox Me, moClass.moSysSession, kmsgUnexpectedSPReturnValue, _
"sppoGetNextReqNo: " & "0"
    End If
    
    If iRetVal = -1 Then
        sNextReqNo = ""
        giSotaMsgBox Me, moClass.moSysSession, kmsgAllPONumsUsed
    End If
    
    sGetNextReqNo = sNextReqNo
    
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Function

ExpectedErrorRoutine:
giSotaMsgBox Me, moClass.moSysSession, kmsgUnexpectedSPReturnValue, _
"sppoGetNextPONo: " & Err.Description
gClearSotaErr

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Function

ExpectedErrorRoutine2:
MyErrMsg moClass, Err.Description, Err, sMyName, "sGetNextPONo"
gClearSotaErr
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Function

'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "sGetNextReqNo", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Sub BindContextMenu()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'***************************************************
'  Instantiate Context Menu Class
'***************************************************

    Set moContextMenu = New clsContextMenu
    
    With moContextMenu
       .BindGrid moGM, grdReqLines.hwnd
       .Bind "POREQLINES", grdReqLines.hwnd, kEntTypePOReqLines
       
       .Bind "POREQ", lkuMain.hwnd, kEntTypePORequisition
       .Bind "IMWHSE", lkuWarehouse.hwnd, kEntTypeIMWarehouse
       
        Set .Form = frmRequistn
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

Public Function ETWhereClause(ByVal ctl As Object, _
                              ByVal taskID As Long, _
                              ByVal EntityType As Integer, _
                              ByVal naturalKeyCols As String, _
                              ByVal TaskGroup As EntityTask_Groups) As String
'+++ VB/Rig Skip +++
'*******************************************************************************
'   Description: OPTIONAL
'                Entity Task Where Clause. This routine allows the application
'                to specify the where clause to be used against the
'                Host Data View for the specified entity.
'
'                If a where clause is not specified, the where clause will be
'                built using the Natural Key Columns specified for the entity.
'
'                Surrogate Keys as natural keys and other criteria make it
'                impossible without supporting metadata to get these values from
'                the application.
'
'   Parameters:
'                ctl <in> - Control that selected the task
'                taskID <in> - Task ID for the Entity Task
'                entityType <in> - Entity Type bound to the control that selected the task
'                naturalKeyCols <in> - Key columns used to specify where clause (FYI)
'                TaskGroup <in> - Task Group 100-500 (FYI)
'
'   Returns:
'                A specified where clause WITHOUT the 'WHERE' verb. (i.e. "1 = 2")
'*******************************************************************************
On Error Resume Next
    
    Select Case True
        Case ctl Is grdReqLines
            ETWhereClause = "ReqLineKey = " & CStr(glGetValidLong(moDmGrid.GetColumnValue(grdReqLines.ActiveRow, "ReqLineKey")))
        Case ctl Is lkuMain
            ETWhereClause = "ReqKey = " & CStr(glGetValidLong(moDmForm.GetColumnValue("ReqKey")))
        Case Else
            ETWhereClause = ""
    End Select
    
    Err.Clear
    
End Function


'************************************************************************
'   Description:
'       process Data Manager state changes (pass new state to toolbar)
'
'   Param:
'       <none>
'
'   Returns:
'
'************************************************************************
Public Sub DMStateChange(oDm As Object, iOldState As Integer, iNewState As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    gGridLockColumn grdReqLines, kcolReqRequestDate 'Agregado por Multiconsulting
    If oDm Is moDmForm Then
        
        If iNewState = kDmStateNone Then
            moClass.lUIActive = kChildObjectInactive
            tbrMain.ButtonEnabled(kTbMemo) = False
        Else
            tbrMain.ButtonEnabled(kTbMemo) = True
        End If
        
        If iNewState = kDmStateAdd Or iNewState = kDmStateNone Then
            bSetFormCurrencyControls msHomeCurrID
        End If

        'by QQ, to fix #5302
        'gGridSetActiveCell grdReqLines, 1, 1
    End If

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DMStateChange", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Function DMGridValidate(oDm As Object, lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'----------------------------------------------------------
'Description: DMGridValidate is called from the Data Manager before it
'attempts to save the row. Here, entries will be checked for
'validity.
'
'If DMGridValidate returns true, the Data Manager will save the
'record. Otherwise, it will stop processing, and the grid
'will remain in the same state as if the use had not pressed
'the finish button.
'
'Return Values:
'TRUE - the entry is valid
'FALSE - the Entry is invalid because anyone of the
'components tested for validity is invalid.
'----------------------------------------------------------
    Dim lCol As Long
    
    DMGridValidate = False
    
'Validate columns for current row
    If oDm Is moDmGrid Then
        
        lCol = kColReqDescription
        If Not bIsValidItemDesc(lRow) Then GoTo SetCellForEdit
        
        lCol = kColReqQtyRequested
        If Not bIsValidQtyRequested(lRow) Then GoTo SetCellForEdit
        
        lCol = kcolReqRequestDate
        If Not bIsValidRequestedDate(lRow) Then GoTo SetCellForEdit
        
        lCol = kColReqPurchDeptID
        If Not bIsValidWhseDept(lRow) Then GoTo SetCellForEdit
        
        If frmRequistn.ActiveControl.Name = grdReqLines.Name Then
            Select Case grdReqLines.ActiveCol
                Case kColReqItemID, kcolReqVendID, kColReqPurchDeptID, kColReqSTaxClassID, kColReqWhseID
                    mbDontSave = False
                    If lRow = grdReqLines.ActiveRow Then
'                    GNav_CellChange grdReqLines.ActiveRow, grdReqLines.ActiveCol
                        moGM_CellChange lRow, grdReqLines.ActiveCol
                    End If
                    If mbDontSave Then Exit Function
            End Select
        End If
    End If

    DMGridValidate = True
    
    
    Exit Function
SetCellForEdit:
     grdReqLines.redraw = False
     gGridSetActiveCell grdReqLines, lRow, lCol
     grdReqLines.redraw = True
     grdReqLines.Refresh
     
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DMGridValidate", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Function DMValidate(oDm As Object) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'----------------------------------------------------------
' Description: DMValidate is called from the Data Manager
' before it attempts to save the record.
' Here, entries will be checked for validity.
'
' If DMValidate returns true, the Data
' Manager will save the record. Otherwise, it
' will stop processing, and the form will
' remain in the same state as if the user had
' not pressed the finish button.
'
' Return Values: TRUE - the entry is valid
' FALSE - the Entry is invalid because anyone
' of the components tested for validity is
' invalid.
'----------------------------------------------------------
DMValidate = False
'Validate the header information
    'Agregado por Multiconsulting
    If sddType.GetIndexByItemData(sddType.ItemData) < 0 Then
        MsgBox "Debe seleccionar un tipo para la requisición", vbExclamation, "Sage MAS 500"
        Exit Function
    End If
    If Not bIsValidVendors Then Exit Function
    'Agregado por Multiconsulting
    If Not bIsValidOriginator Then Exit Function
    If Not bIsValidContact Then Exit Function
    If Not bIsValidTranDate Then Exit Function
DMValidate = True

'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DMValidate", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Public Function iConfirmUnload(Optional vNoClear As Variant) As Integer
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

Dim i       As Long
Dim iRetVal As Integer
Dim bOldMb  As Boolean

    bOldMb = mbDontChkclick
    
    If IsMissing(vNoClear) Then
        vNoClear = False
    End If

    
    If miSecurityLevel = kSecLevelDisplayOnly Then
        iConfirmUnload = kDmSuccess
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If
    
    
    If moDmForm.IsDirty(True) Then
        mbDontChkclick = True
        iConfirmUnload = moDmForm.ConfirmUnload(vNoClear)
        mbDontChkclick = bOldMb
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If


    If vNoClear Then
    
    Else
        mbDontChkclick = True
        moDmForm.Action kDmCancel
'        ClearForm
        mbDontChkclick = bOldMb
    End If
    
    iConfirmUnload = kDmSuccess

'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "iConfirmUnload", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function
Public Sub DMReposition(oDm As Object)
'+++ VB/Rig Begin Push +++                                                                'Repository Error Rig  {1.1.1.0.0}
#If ERRORTRAPON Then                                                                      'Repository Error Rig  {1.1.1.0.0}
    On Error GoTo VBRigErrorRoutine                                                       'Repository Error Rig  {1.1.1.0.0}
#End If                                                                                   'Repository Error Rig  {1.1.1.0.0}
'+++ VB/Rig End +++                                                                       'Repository Error Rig  {1.1.1.0.0}
    Dim lKeyValue As Long
    
    If oDm Is moDmForm Then
            
        lKeyValue = glGetValidLong(moDmForm.GetColumnValue("ReqKey"))
        '-- Set the Memo toolbar to the correct state
        gSetMemoToolBarState moClass.moAppDB, lKeyValue, kEntTypePORequisition, msCompanyID, tbrMain
    
    End If
     
Exit Sub

'+++ VB/Rig Begin Pop +++                                                                 'Repository Error Rig  {1.1.1.0.0}
    Exit Sub                                                                              'Repository Error Rig  {1.1.1.0.0}
VBRigErrorRoutine:                                                                        'Repository Error Rig  {1.1.1.0.0}
        gSetSotaErr Err, "frmRequistn", "DMReposition", VBRIG_IS_FORM                     'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number                                                        'Repository Error Rig  {1.1.1.0.0}
'+++ VB/Rig End +++                                                                       'Repository Error Rig  {1.1.1.0.0}
End Sub

Public Function bSetFocus(oControl As Object) As Boolean
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++

On Error GoTo ExpectedErrorRoutine

    If Not (oControl Is Nothing) Then
        
        If oControl.Enabled = True Then
            
            If oControl.Visible = True Then
                
                On Error Resume Next
                oControl.SetFocus
            
            End If
        
        End If
    
    End If
    
    bSetFocus = True

ExpectedErrorRoutine:


'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bSetFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Public Sub ClearForm()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'*************************************************************************
' Desc:    Cleans up form.
'*************************************************************************
    lkuMain.Protected = False

    'This should be taken care of by setting the protected property of the lookup control.
    lkuMain.TabStop = True
    
    bSetFocus lkuMain
        
    SetTags
    ChangeToolBar (False)
    txtStatus.Text = ""
    
    'Enable the fields that could be disabled if a closed Req was displayed.
    txtOriginator.Enabled = True
    calDate.Enabled = True
    txtContact.Enabled = True
    chkExpedite.Enabled = True
    cboExpReason.Enabled = True
    lkuDept.Enabled = True
    
    '*****************************************************************************************************
    'Agregado por Multiconsulting Osmel Barreras
    adicInfo = False
    firstCompEstatusReq = True
    sddEstatusReq.ListIndex = -1
    'sddEstatusReq.Locked = True
    txtAutorizaReq.Text = ""
    txt2doAutorizaReq.Text = ""
    txtAceptaReq.Text = ""
    calAceptaReq = ""
    cal2doAutorizaReq = ""
    calAutorizaReq = ""
    chkb2doAutorizoReq.Value = 0
    chkb2doAutorizoReq.Enabled = False
    chkbEstatusDesc.Value = 0
    chkbEstatusDesc.Enabled = False
    txtEstatusReqDesc = ""
    navBuyerGrid.Visible = False
    txtUserMod.Text = ""
    txtEstatusReqDesc.BackColor = vbWindowBackground
    'Agregado por Multiconsulting Osmel Barreras
    '******************************************************************************************************
    
    If mbIntegrateWithIM Then
        lkuWarehouse.Enabled = True
    End If
    
    txtComment.Enabled = True
    grdReqLines.Enabled = True
    grdReqLines.MaxRows = 0
    grdReqLines.Col = kColReqItemID
    cmdGenerate.Enabled = True
    navItemGrid.Visible = False
    navSTaxGrid.Visible = False
    navDeptGrid.Visible = False
    navWhseGrid.Visible = False
    navVendorGrid.Visible = False
    
    'Change the color to window in case the prior Req had a closed status.
    txtOriginator.BackColor = vbWindowBackground
    txtContact.BackColor = vbWindowBackground
    txtComment.BackColor = vbWindowBackground

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "ClearForm", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Function bGetOptions() As Boolean
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
'****************************************************************************
'  Get User Options
'****************************************************************************

    Dim iGlOvrdSeg As Integer
    Dim sSql As String
    Dim miInt As Integer

    On Error GoTo ExpectedErrorRoutine
' Use Phruity's Option retrieval thang to get required info.
    miCostDecPlaces = moOptions.CI("UnitCostDecPlaces")
    miQtyDecPlaces = moOptions.CI("QtyDecPlaces")
    mbTrackSTax = (moOptions.PO("TrackSTaxOnPurch") = 1)
    mbIntegrateWithIM = (moOptions.PO("IntegrateWithIM") = 1)
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Function

ExpectedErrorRoutine:

MyErrMsg moClass, Err.Description, Err, sMyName, "bGetOptions"
gClearSotaErr
mbCancelLoad = True

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Function

'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bGetOptions", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function


Private Sub FormatGrid()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'***********************************************************************
' Desc: Initializes/formats the 'Requisition Lines' grid.
'***********************************************************************

  'Set default properties
    gGridSetProperties grdReqLines, kMaxCols, _
kGridDataSheetNoAppend
                       
    grdReqLines.UserResizeCol = SS_USER_RESIZE_OFF


    gGridSetProperties grdReqLineDtl, kChildPOMaxCols, _
kGridDataSheetNoAppend

'    grdReqLines.UnitType = SS_CELL_UNIT_VGA          ' make VGA base
'    grdReqLines.ScrollBars = SS_SCROLLBAR_BOTH

    grdReqLines.DisplayRowHeaders = True
    grdReqLines.RowHeaderDisplay = SS_HEADER_NUMBERS

    grdReqLines.TypeFloatSeparator = True
    
    grdReqLines.MaxCols = kMaxCols
    grdReqLines.MaxRows = 0

  'Set the header for 'Item'
    gGridSetHeader grdReqLines, kColReqItemID, "Item"

  'Set the header for 'Item Description'
    gGridSetHeader grdReqLines, kColReqDescription, "Description"

  'Set the header for 'Qty Requested'
    gGridSetHeader grdReqLines, kColReqQtyRequested, "Qty Requested"

  'Set the header for 'UOM'
    gGridSetHeader grdReqLines, kColReqUnitMeasID, "UOM"

  'Set the header for 'Unit Cost'
    gGridSetHeader grdReqLines, kColReqUnitCost, "Unit Cost"

  'Set the header for 'Extended Amt'
    gGridSetHeader grdReqLines, kColReqExtAmt, "Extended Amt"

  'Set the header for 'Vendor ID'
    gGridSetHeader grdReqLines, kcolReqVendID, "Vendor"

  'Set the header for 'Currency ID'
    gGridSetHeader grdReqLines, kColReqCurrid, "CURR"
    
  'Set the header for 'Request Date'
    gGridSetHeader grdReqLines, kcolReqRequestDate, "Request Date"

  'Set the header for 'Department'
    gGridSetHeader grdReqLines, kColReqPurchDeptID, "Department"
  
  'Set the header for 'Warehouse'
    gGridSetHeader grdReqLines, kColReqWhseID, "Warehouse"
  
 ' lkuWarehouse =  kColReqWhseID, "Warehouse"
  
  'Set the header for 'Sales Tax Class'
    gGridSetHeader grdReqLines, kColReqSTaxClassID, "Tax Class"

  'Set the header for 'Comment'
    gGridSetHeader grdReqLines, kColReqComment, "Comment"
  
  'Set the header for 'PO Lines Custom Fields'
    gGridSetHeader grdReqLines, kColReqPOlineCustomFields, "PO Lines Custom Fields"

  'Set the header for 'PO ID'
    gGridSetHeader grdReqLines, kColReqPOTranID, "PO"

'Agregado por Multiconsulting Osmel Barreras
    'Set the header for 'Presupuesto Estimado'
    gGridSetHeader grdReqLines, kColReqEstPres, "Presupuesto Estimado"
    'Set the header for 'Comprador'
    gGridSetHeader grdReqLines, kColReqBuyerID, "Comprador"
    'Set the header for 'Estado de Articulo'
    gGridSetHeader grdReqLines, kColReqLStItemID, "Estado Artículo"
    'Set the header for 'Tipo de Compra'
    gGridSetHeader grdReqLines, kColReqLTBItemID, "Tipo Compra"
'Agregado por Multiconsulting Osmel Barreras Osmel
    
  'Set grid Column Widths
    gGridSetColumnWidth grdReqLines, kColReqItemID, 15
    gGridSetColumnWidth grdReqLines, kColReqDescription, 21
    gGridSetColumnWidth grdReqLines, kColReqQtyRequested, 14
    gGridSetColumnWidth grdReqLines, kColReqUnitMeasID, 6
    gGridSetColumnWidth grdReqLines, kColReqUnitCost, 14
    gGridSetColumnWidth grdReqLines, kColReqExtAmt, 14
    gGridSetColumnWidth grdReqLines, kcolReqVendID, 12
    gGridSetColumnWidth grdReqLines, kcolReqRequestDate, 12
    gGridSetColumnWidth grdReqLines, kColReqPurchDeptID, 10
    gGridSetColumnWidth grdReqLines, kColReqWhseID, 8
    gGridSetColumnWidth grdReqLines, kColReqComment, 21
    gGridSetColumnWidth grdReqLines, kColReqPOlineCustomFields, 16
    gGridSetColumnWidth grdReqLines, kColReqPOTranID, 10
    gGridSetColumnWidth grdReqLines, kColReqSTaxClassID, 10
    gGridSetColumnWidth grdReqLineDtl, kColChildRequestDate, 12
    gGridSetColumnWidth grdReqLineDtl, kColChildQtyReq, 14
    gGridSetColumnWidth grdReqLines, kColReqCurrid, 6
    
    'Agregado por Multiconsulting Osmel Barreras Osmel
    gGridSetColumnWidth grdReqLines, kColReqEstPres, 16
    gGridSetColumnWidth grdReqLines, kColReqBuyerID, 20
    gGridSetColumnWidth grdReqLines, kColReqLStItemID, 25
    gGridSetColumnWidth grdReqLines, kColReqLTBItemID, 16
    'Agregado por Multiconsulting Osmel Barreras Osmel
    
  'Align the visible columns
    gGridHAlignColumn grdReqLines, kColReqItemID, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqDescription, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqQtyRequested, SS_CELL_TYPE_FLOAT, SS_CELL_H_ALIGN_RIGHT
    gGridHAlignColumn grdReqLines, kColReqUnitMeasID, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqUnitCost, SS_CELL_TYPE_FLOAT, SS_CELL_H_ALIGN_RIGHT
    gGridHAlignColumn grdReqLines, kColReqExtAmt, SS_CELL_TYPE_FLOAT, SS_CELL_H_ALIGN_RIGHT
    gGridHAlignColumn grdReqLines, kcolReqVendID, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLineDtl, kColChildRequestDate, SS_CELL_TYPE_DATE, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLineDtl, kColChildQtyReq, SS_CELL_TYPE_FLOAT, SS_CELL_H_ALIGN_RIGHT
    gGridHAlignColumn grdReqLineDtl, kColChildReqLineDistKey, SS_CELL_TYPE_INTEGER, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqPurchDeptID, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqWhseID, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqSTaxClassID, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqComment, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqPOlineCustomFields, SS_CELL_TYPE_BUTTON, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqPOTranID, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqCurrid, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_LEFT
    
    'Agregado por Multiconsulting Osmel Barreras Osmel
    gGridHAlignColumn grdReqLines, kColReqEstPres, SS_CELL_TYPE_FLOAT, SS_CELL_H_ALIGN_RIGHT
    gGridHAlignColumn grdReqLines, kColReqBuyerID, SS_CELL_TYPE_EDIT, SS_CELL_H_ALIGN_CENTER
    gGridHAlignColumn grdReqLines, kColReqLStItemID, SS_CELL_TYPE_COMBOBOX, SS_CELL_H_ALIGN_LEFT
    gGridHAlignColumn grdReqLines, kColReqLTBItemID, SS_CELL_TYPE_COMBOBOX, SS_CELL_H_ALIGN_LEFT
    'Agregado por Multiconsulting Osmel Barreras Osmel
    
' Set column type
    gGridSetColumnType grdReqLines, kColReqItemID, SS_CELL_TYPE_EDIT, 30
    gGridSetColumnType grdReqLines, kColReqDescription, SS_CELL_TYPE_EDIT, 40
    gGridSetColumnType grdReqLines, kcolReqRequestDate, SS_CELL_TYPE_DATE
                    grdReqLines.TypeDateMin = "01011900"
                    grdReqLines.TypeDateMax = "12312099"
    gGridSetColumnType grdReqLineDtl, kColChildRequestDate, SS_CELL_TYPE_DATE
                    grdReqLineDtl.TypeDateMin = "01011900"
                    grdReqLineDtl.TypeDateMax = "12312099"
    gGridSetColumnType grdReqLineDtl, kColChildQtyReq, SS_CELL_TYPE_FLOAT, miQtyDecPlaces, 8
    gGridSetColumnType grdReqLines, kColReqUnitCost, SS_CELL_TYPE_FLOAT, miCostDecPlaces, 10
    gGridSetColumnType grdReqLines, kColReqExtAmt, SS_CELL_TYPE_FLOAT, miCostDecPlaces, 12
    gGridSetColumnType grdReqLines, kColReqQtyRequested, SS_CELL_TYPE_FLOAT, miQtyDecPlaces, 8
    gGridSetColumnType grdReqLines, kColReqUnitMeasID, SS_CELL_TYPE_EDIT
    gGridSetColumnType grdReqLines, kColReqSTaxClassID, SS_CELL_TYPE_EDIT, 15
    gGridSetColumnType grdReqLines, kColReqComment, SS_CELL_TYPE_EDIT, 50
    gGridSetColumnType grdReqLines, kColReqPOlineCustomFields, SS_CELL_TYPE_BUTTON, 16
    gGridSetColumnType grdReqLines, kColReqCurrid, SS_CELL_TYPE_EDIT
    gGridSetColumnType grdReqLines, kColReqUnitCostExact, SS_CELL_TYPE_FLOAT, 13, 12
    
    'Agregado por Multiconsulting Osmel Barreras Osmel
    gGridSetColumnType grdReqLines, kColReqEstPres, SS_CELL_TYPE_FLOAT, miQtyDecPlaces, 20
    gGridSetColumnType grdReqLines, kColReqBuyerID, SS_CELL_TYPE_EDIT, 40
    gGridSetColumnType grdReqLines, kColReqLStItemID, SS_CELL_TYPE_COMBOBOX, moStaticListStateItem.ListData
    gGridSetColumnType grdReqLines, kColReqLTBItemID, SS_CELL_TYPE_COMBOBOX, moStaticListTypeBuyItem.ListData
    'Agregado por Multiconsulting Osmel Barreras Osmel
    
''Define the max field size for each of the fields.
'    With grdReqLines
'        .Row = -1
'        .col = kColReqQtyRequested
'        .TypeFloatMax = kvMaxQtySize
'        .Row = -1
'        .col = kColReqUnitCost
'        .TypeFloatMax = kvMaxUnitCost
'        .Row = -1
'        .col = kColReqExtAmt
'        .TypeFloatMax = kvMaxExtAmt
'    End With
'    With grdReqLineDtl
'        .Row = -1
'        .col = kColChildQtyReq
'        .TypeFloatMax = kvMaxQtySize
'    End With

    With grdReqLines
        .Row = -1
        .Col = kColReqPOlineCustomFields
        .TypeButtonText = kCustomFields
    End With

  'Hide the appropriate columns
    gGridHideColumn grdReqLines, kColReqLineKey
    gGridHideColumn grdReqLines, kColReqItemKey
    gGridHideColumn grdReqLines, kColReqUnitMeasKey
    gGridHideColumn grdReqLines, kColReqVendKey
    gGridHideColumn grdReqLines, kColReqPOLineKey
    gGridHideColumn grdReqLines, kColReqPOKey
    gGridHideColumn grdReqLines, kColReqPurchDeptKey
    gGridHideColumn grdReqLines, kColReqWhseKey
    gGridHideColumn grdReqLines, kColReqSTaxClassKey
    gGridHideColumn grdReqLines, kColPOUOMType
    gGridHideColumn grdReqLines, kColPOUserFld1
    gGridHideColumn grdReqLines, kColPOUserFld2
    gGridHideColumn grdReqLines, kColReqUnitCostExact

    'Agregado por Multiconsulting Osmel Barreras Osmel
    gGridHideColumn grdReqLines, kColReqUnitCost
    gGridHideColumn grdReqLines, kColReqExtAmt
    gGridHideColumn grdReqLines, kColReqBuyerkey
    gGridHideColumn grdReqLines, kColReqLStItemKey
    gGridHideColumn grdReqLines, kColReqLTBItemKey
    'Agregado por Multiconsulting Osmel Barreras Osmel

    If Not mbTrackSTax Then
        gGridHideColumn grdReqLines, kColReqSTaxClassID
    End If
    If Not mbIntegrateWithIM Then
        gGridHideColumn grdReqLines, kColReqWhseID
    End If
    
  'Lock the appropriate columns
    gGridLockColumn grdReqLines, kColReqPOTranID
    gGridLockColumn grdReqLines, kColReqUnitMeasID
    gGridLockColumn grdReqLines, kColReqCurrid
    
  'Make the Items column frozen
    gGridFreezeCols grdReqLines, kColReqItemID
    
  'Set Grid Colors
    gGridSetColors grdReqLines
    grdReqLines.Row = 0
  'Set Captions for messages
    grdReqLines.Col = kColReqItemID
    msItemIDCaption = grdReqLines.Text
    grdReqLines.Col = kColReqDescription
    msItemDescCaption = grdReqLines.Text
    grdReqLines.Col = kcolReqRequestDate
    msRequestDateCaption = grdReqLines.Text

        
    'Exit this subroutine
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "FormatGrid", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub calDate_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus calDate, True
    #End If
'+++ End Customizer Code Push +++
'Agregado por Multiconsulting
    Dim i As Long
    Dim iLeepTime As Integer
    Dim lItemtemKey As Long
    Dim sDate As String
'Agregado por Multiconsulting

    If Not calDate.IsValid Then
        giSotaMsgBox Me, moClass.moSysSession, kmsgInvalidDate
        calDate = ""
        calDate.SetFocus
    Else
        'Agregado por Multiconsulting
        For i = 1 To grdReqLines.DataRowCnt
            lItemtemKey = glGetValidLong(gsGridReadCellText(grdReqLines, i, kColReqItemKey))
            If lItemtemKey <> 0 Then
                If mbIntegratedCT And lGetReqContract > 0 Then
                    iLeepTime = giGetValidInt(moClass.moAppDB.Lookup("s.DeliveryTime", "tctContract AS p JOIN tctContractLine AS s ON s.ContractKey = p.ContractKey", "s.ItemKey =" & lItemtemKey & " and p.ContractKey =" & lGetReqContract))
                Else
                    iLeepTime = giGetValidInt(moClass.moAppDB.Lookup("p.TimeDelivery", "timItemSecurityStock AS p", "p.ItemKey =" & lItemtemKey))
                End If
                
                sDate = sGetDate(DateAdd("d", iLeepTime, calDate.Value))

                gGridUpdateCell grdReqLines, i, kcolReqRequestDate, sDate
                gGridUpdateCell grdReqLineDtl, i, kColChildRequestDate, sDate
                moDMSubGrid.SetRowDirty i
                moDmGrid.SetRowDirty i
            End If
        Next i
        'Agregado por Multiconsulting
    End If
        
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "calDate_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub



Private Sub Form_Initialize()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    mbCancelShutDown = False
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Form_Initialize", VBRIG_IS_FORM
        Select Case VBRIG_IS_FORM_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++

    On Error Resume Next
    
    If KeyCode = vbKeyF5 Then
        mbIsPressF5 = True
    Else
        mbIsPressF5 = False
    End If
    
    Select Case KeyCode

        Case vbKeyF4
' if F4 is pressed without any other key call up navigator
            If Shift = 0 Then
                lkuMain.DoLookupClick
' Set the keycode to zero to prevent the calendar from poping on the grid.
                KeyCode = 0
            End If
        Case vbKeyF1 To vbKeyF12
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

'************************************************************************
'   Description:
'       process keypresses on the form.  NOTE: key preview of the form
'       should be set to True
'
'   Param:
'       KeyAscii -  ascii key code of key pressed
'
'   Returns:
'
'************************************************************************

Private Sub Form_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    mbIsPressF5 = False
  
    Select Case KeyAscii
        Case vbKeyReturn
            If mbEnterAsTab Then
                gProcessSendKeys "{Tab}"
                KeyAscii = 0
            End If
        Case Else
            'Other KeyAscii routines.
    End Select
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "Form_KeyPress", VBRIG_IS_FORM
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
'******************************************
'* Set up form
'******************************************
    
   'Set global class object
    Set goClass = moClass
    
    moClass.lUIActive = kChildObjectInactive
  
   'Set the status bar object
   Set sbrMain.Framework = moClass.moFramework
   sbrMain.Status = SOTA_SB_START
  
    'Get Session defaults
    mlLanguage = moClass.moSysSession.Language       'Language ID
    msCompanyID = moClass.moSysSession.CompanyId         'Company ID
    mbEnterAsTab = moClass.moSysSession.EnterAsTab       'Enter Key like Tab Key
    msBusinessDate = moClass.moSysSession.BusinessDate   'Current business date
    msHomeCurrID = moClass.moSysSession.CurrencyID
    msCurrentUser = moClass.moSysSession.UserName
    
    SetupIMInterfaces
    
   'Setup the Module Options class
   Set moOptions.oAppDB = moClass.moAppDB
   Set moOptions.oSysSession = moClass.moSysSession
   moOptions.sCompanyID = msCompanyID
   
    bGetOptions
   'if not integrated with IM, disable the Warehouse lookup
   If Not mbIntegrateWithIM Then
        lkuWarehouse.Enabled = False
   End If
  
    'Set form initial height and Width
    miOldFormHeight = Me.Height
    miOldFormWidth = Me.Width
    miMinFormHeight = miOldFormHeight
    miMinFormWidth = miOldFormWidth
    
  'Load Header level combo boxes
    LoadHeaderLists
    
' Set the form caption
    Me.Caption = gsStripChar(gsBuildString(ksEnterReq, moClass.moAppDB, moClass.moSysSession), kAmpersand)


    BindForm        'Bind the controls to the database
    BindGM          'Bind the grid manager
    BindContextMenu 'Bind the context menu
    BindNavigators  'Bind the floating navigator
  
  'Setup Toolbar - Single row maintenance - Remove Rename and Memo buttons
    tbrMain.Init sotaTB_TRANSACTION, moClass.moSysSession.Language
    With moClass.moFramework
        tbrMain.SecurityLevel = _
.GetTaskPermission(.GetTaskID())
    End With
    tbrMain.RemoveButton kTbRenameId
    tbrMain.ButtonEnabled(kTbPrint) = False
    'Set User Security Level
    miSecurityLevel = giSetAppSecurity(moClass, tbrMain, moDmForm, moDmGrid, moDMSubGrid)
    If miSecurityLevel = kSecLevelDisplayOnly Then
        tbrMain.ButtonEnabled(kTbNextNumber) = False
    End If

'Set up initial call to user fields.
    Set moUF = CreateObject("cizdbdl1.clsUserFld")
    moUF.Init moClass.moSysSession, moClass.moAppDB, moClass.moAppDB, moSotaObjects, moClass.moFramework
    
  'Set up Grid
    FormatGrid
'    MapControls
    SetupLookups
    
'********************************************************************************
'Modificado por MultiConsulting Osmel Barreras
'   Check the user id typed in can change the Estatus Requisition.
    CheckUserStatusReqPermission
    firstCompEstatusReq = True
    
    'Jose
    IsCTIntegrated
    If mbIntegratedCT Then
        cmdContract.Visible = True
        Set frmContractAssociate.oclass = moClass
        Load frmContractAssociate
    End If
'Modificado por MultiConsulting Osmel Barreras
'********************************************************************************
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



Private Sub grdReqLines_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    moGM.Grid_ColWidthChange Col1
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_ColWidthChange", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub grdReqLines_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    moGM.Grid_EditMode Col, Row, Mode, ChangeMade
' Get a snapshot of the value in the column prior to changing so that is can be
' set back to its original value if there is an error
'Debug.Print "Edit Mode - Column = " & col & " Mode = " & Mode
    If Mode = 1 Then
        Select Case Col
            Case kColReqQtyRequested
                msOldReqQtyRequested = gsGridReadCellText(grdReqLines, Row, Col)
                '*************************************************************************************************
'Modificado por MultiConsulting Osmel Barreras
                If sddEstatusReq.ItemData = kReqStatusAccepted Then
                    gGridLockColumn grdReqLines, kColReqQtyRequested
                    gGridUpdateCellText grdReqLines, Row, kColReqQtyRequested, msPersistCantReq
                    gGridUpdateCellText grdReqLineDtl, Row, kColChildQtyReq, msPersistCantReq
                    Exit Sub
                End If

            Case kColReqEstPres
                msOldReqEstPres = gsGridReadCellText(grdReqLines, Row, Col)
                'msPersistPresEstReq = gsGridReadCellText(grdReqLines, Row, kColReqEstPres)
                
                If sddEstatusReq.ItemData = kReqStatusAccepted Then
                    gGridLockColumn grdReqLines, kColReqEstPres
                    gGridUpdateCellText grdReqLines, Row, kColReqEstPres, msPersistPresEstReq
                    Exit Sub
                End If
                
            Case kColReqItemID
                If gsGridReadCellText(grdReqLines, Row, kColReqItemID) = "" Then
                    msPersistDescriptionReq = gsGridReadCellText(grdReqLines, Row, kColReqDescription)
                End If
                
                If (gsGridReadCellText(grdReqLines, Row, kColReqDescription) = "" And gsGridReadCellText(grdReqLines, Row, kColReqQtyRequested) = "" And gsGridReadCellText(grdReqLines, Row, kColReqEstPres) = "") And (sddEstatusReq.ItemData = kReqStatusAccepted And evSegUserBuyInfo = 1) Then
                    gGridDeleteRow grdReqLines, Row
                    navItemGrid.Visible = False
                    grdReqLines.Refresh
                    txtOriginator.SetFocus
                End If
'Modificado por MultiConsulting Osmel Barreras
'*************************************************************************************************
            Case kColReqUnitCost
                msOldReqUnitCost = gsGridReadCellText(grdReqLines, Row, Col)
            Case kColReqExtAmt
                msOldReqExtAmt = gsGridReadCellText(grdReqLines, Row, Col)
        End Select
    End If
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_EditMode", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub






Private Sub grdReqLines_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
  'Grid manager change of top or left row action
    moGM.Grid_TopLeftChange OldLeft, OldTop, NewLeft, NewTop

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_TopLeftChange", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Sub GetItemInfo(sItemKey As String, bFoundSTaxClass As Boolean, bIsCommentOnly As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim rs As Object
    Dim sSql As String
    
    bFoundSTaxClass = False
    bIsCommentOnly = False
    
    If sItemKey <> "" Then
        sSql = "SELECT STaxClassKey, ItemType FROM timItem WHERE ItemKey = "
        sSql = sSql & sItemKey
        sSql = sSql & " AND CompanyID = " & gsQuoted(msCompanyID)
        Set rs = oclass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
        If Not rs.IsEOF Then
            bIsCommentOnly = (rs.Field("ItemType") = kItemTypeComment)
            If mbTrackSTax Then
                If Not IsNull(rs.Field("STaxClassKey")) Then
                    bFoundSTaxClass = (rs.Field("STaxClassKey") > 0)
                End If
            End If

        End If
    End If
    Set rs = Nothing
    
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "GetItemInfo", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Function bLoadUOMDesc(lUOMKey As Long, lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim rs As Object
    Dim sSql As String
    
    If lUOMKey > 0 Then
        sSql = "SELECT UnitMeasID FROM tciUnitMeasure WHERE UnitMeasKey = "
        sSql = sSql & lUOMKey
        sSql = sSql & " AND CompanyID = " & gsQuoted(msCompanyID)
        Set rs = oclass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
        If rs.IsEOF Then
            gGridUpdateCell grdReqLines, lRow, kColReqUnitMeasID, ""
        Else
            gGridUpdateCell grdReqLines, lRow, kColReqUnitMeasID, rs.Field("UnitMeasID")
        End If
    Else
        gGridUpdateCell grdReqLines, lRow, kColReqUnitMeasID, ""
    End If
    Set rs = Nothing
    
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bLoadUOMDesc", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Function bLoadSTaxClassDesc(lSTaxClassKey As Long, lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim rs As Object
    Dim sSql As String
    
    If lSTaxClassKey > 0 Then
        sSql = "SELECT STaxClassID FROM tciSTaxClass WHERE STaxClassKey = "
        sSql = sSql & lSTaxClassKey
        Set rs = oclass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
        If rs.IsEOF Then
            gGridUpdateCell grdReqLines, lRow, kColReqSTaxClassID, ""
        Else
            gGridUpdateCell grdReqLines, lRow, kColReqSTaxClassID, rs.Field("STaxClassID")
        End If
    Else
        gGridUpdateCell grdReqLines, lRow, kColReqSTaxClassID, ""
    End If
    Set rs = Nothing
    
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bLoadSTaxClassDesc", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Function bLoadItemDflts(lItemKey As Long, lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

Dim sItemID             As String
Dim sItemDesc           As String
Dim iStatus             As Integer
Dim iItemType           As Integer
Dim dUnitCostHC         As Double
Dim lUOMKey             As Long
Dim lSTaxClassKey       As Long
Dim lGLAcctKey          As Long
Dim lToleranceKey       As Long
Dim iAllowCostOvrd      As Integer
Dim iBadUOM             As Integer
Dim iRetVal             As Integer
Dim dExchRate           As Double
Dim iUOMType            As Integer
Dim sSql                As String
Dim rs                  As Object
Dim bOldmbClick         As Boolean
Dim iDecPlaces          As Integer
Dim bWasComment         As Boolean
Dim iReqRcvr            As Integer
Dim lContractKey        As Long 'Agregado por multiconsulting

' If this is called for a closed line, don't do anything
     If gsGridReadCellText(grdReqLines, lRow, kColReqPOTranID) <> "" Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
     End If

     With moClass.moAppDB
        .SetInParam msCompanyID
        .SetInParamNull SQL_CHAR
        .SetInParam lItemKey
        .SetInParam mlLanguage
        .SetInParam 0
        .SetInParam 0
        .SetInParamNull SQL_CHAR
        .SetInParam 0
        .SetInParam 0
        .SetOutParam lItemKey
        .SetOutParam sItemID
        .SetOutParam sItemDesc
        .SetOutParam iStatus
        .SetOutParam iItemType
        .SetOutParam dUnitCostHC
        .SetOutParam lUOMKey
        .SetOutParam lSTaxClassKey
        .SetOutParam lGLAcctKey
        .SetOutParam lToleranceKey
        .SetOutParam iAllowCostOvrd
        .SetOutParam iUOMType
        .SetOutParam iBadUOM
        .SetOutParam iReqRcvr
        .SetOutParam iRetVal
        .ExecuteSP "sppoValidItem"
        lItemKey = glGetValidLong(.GetOutParam(10))
        sItemID = gsGetValidStr(.GetOutParam(11))
        sItemDesc = gsGetValidStr(.GetOutParam(12))
        iStatus = giGetValidInt(.GetOutParam(13))
        iItemType = giGetValidInt(.GetOutParam(14))
        dUnitCostHC = gdGetValidDbl(.GetOutParam(15))
        lUOMKey = glGetValidLong(.GetOutParam(16))
        lSTaxClassKey = glGetValidLong(.GetOutParam(17))
        lGLAcctKey = glGetValidLong(.GetOutParam(18))
        lToleranceKey = glGetValidLong(.GetOutParam(19))
        iAllowCostOvrd = giGetValidInt(.GetOutParam(20))
        iUOMType = giGetValidInt(.GetOutParam(21))
        iBadUOM = giGetValidInt(.GetOutParam(22))
        iReqRcvr = giGetValidInt(.GetOutParam(23))
        iRetVal = giGetValidInt(.GetOutParam(24))
        .ReleaseParams
    End With
    If lItemKey > 0 Then
        gGridUpdateCell grdReqLines, lRow, kColReqDescription, sItemDesc
    End If
    gGridUpdateCell grdReqLines, lRow, kColReqUnitCost, Str(dUnitCostHC)
    gGridUpdateCell grdReqLines, lRow, kColReqUnitCostExact, Str(dUnitCostHC)
    bCalcLineAmts lRow, kColReqQtyRequested
    
    'Agregado por multiconsulting
    'Cargar UM del contrato
    If mbIntegratedCT Then
        lContractKey = lGetReqContract
        If lContractKey > 0 And lItemKey > 0 Then
            lUOMKey = lGetUMFromContract(lItemKey, lContractKey)
        End If
    End If
    'Agregado por multiconsulting
    
    If lUOMKey = 0 Then
        gGridUpdateCell grdReqLines, lRow, kColReqUnitMeasKey, ""
    Else
        gGridUpdateCell grdReqLines, lRow, kColReqUnitMeasKey, CStr(lUOMKey)
    End If
    bLoadUOMDesc lUOMKey, lRow
    If lSTaxClassKey > 0 Then
        gGridUpdateCell grdReqLines, lRow, kColReqSTaxClassKey, CStr(lSTaxClassKey)
        bLoadSTaxClassDesc lSTaxClassKey, lRow
        gGridLockCell grdReqLines, kColReqSTaxClassID, lRow
    Else
        gGridUnlockCell grdReqLines, kColReqSTaxClassID, lRow
    End If
    If iItemType = kItemTypeComment Then
        gGridLockCell grdReqLines, kColReqQtyRequested, lRow
        gGridLockCell grdReqLines, kColReqUnitCost, lRow
        gGridLockCell grdReqLines, kColReqExtAmt, lRow
        gGridUpdateCell grdReqLines, lRow, kColReqQtyRequested, "0"
        gGridUpdateCell grdReqLineDtl, 1, kColChildQtyReq, "0"
        moDMSubGrid.SetRowDirty 1

    Else
        gGridUnlockCell grdReqLines, kColReqQtyRequested, lRow
        gGridUnlockCell grdReqLines, kColReqUnitCost, lRow
        gGridUnlockCell grdReqLines, kColReqExtAmt, lRow
    End If
    
    moDmGrid.SetRowDirty (lRow)
    
'+++ VB/Rig Begin Pop +++
        Exit Function
Resume
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bLoadItemDflts", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Sub moGM_GridBeforeDelete(bContinue As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    bContinue = True
    Debug.Print "Before Delete " & msOldItemID
    
    If gsGridReadCellText(grdReqLines, grdReqLines.Row, kColReqPOTranID) <> "" Then
        bContinue = False
        giSotaMsgBox Me, moClass.moSysSession, kmsgCannotDelReqLine
    End If
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moGM_GridBeforeDelete", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
    
End Sub
Private Sub moGM_GridAfterDelete()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    
    If gsGridReadCellText(grdReqLines, grdReqLines.Row, kColReqPOTranID) <> "" Then
        giSotaMsgBox Me, moClass.moSysSession, kmsgCannotDelReqLine
    End If
    grdReqLines.Refresh
    Debug.Print "After Delete " & msOldItemID
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moGM_GridAfterDelete", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
    
End Sub

Private Sub moGM_CellChange(ByVal lRow As Long, ByVal lCol As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lLookupKey    As Long
    Dim sLookupID     As String
    Dim sLocalOldItemID As String
    Static bMsgDisplayed As Boolean
    Dim sMessage As String
    Dim sWhseID As String
    Dim lWhseKey As Long
    'Agregado por Multiconsulting
    Dim lContractKey As Long
    Dim iLeepTime As Integer
    Dim sDate As String
    'Agregado por Multiconsulting
    
    mbIsInvalid = False
    
    ' Only display the message once.
    If Not bMsgDisplayed Then
        With grdReqLines
            Select Case lCol
                Case kColReqItemID
                    sLookupID = Trim(gsGridReadCell(grdReqLines, lRow, lCol))
                    If mbIntegrateWithIM Then
                        lLookupKey = gvCheckNull(moClass.moAppDB.Lookup _
                                                ("ItemKey", "timItem", "ItemID = " & gsQuoted(sLookupID) & " AND ItemType IN (1,3,4,5,6,8,9) " & _
                                                "AND Status = 1 AND CompanyID = " & _
                                                gsQuoted(msCompanyID)), SQL_INTEGER)
                    Else
                        lLookupKey = gvCheckNull(moClass.moAppDB.Lookup _
                                                ("ItemKey", "timItem", "ItemID = " & gsQuoted(sLookupID) & " AND ItemType IN (1,3,4) " & _
                                                "AND Status = 1 AND CompanyID = " & _
                                                gsQuoted(msCompanyID)), SQL_INTEGER)
                    End If
                    'Agregado por Multiconsulting
                    If mbIntegratedCT Then
                        lContractKey = lGetReqContract
                        If lContractKey > 0 Then
                            If Len(Trim(sLookupID)) > 0 And lLookupKey > 0 Then
                                If giGetValidInt(moClass.moAppDB.Lookup("count(*)", "timItem", "ItemKey =" & lLookupKey & " and ItemKey in (" & sGetAvaliableItemsListFromContract(lContractKey) & ")")) = 0 Then
                                    lLookupKey = 0
                                Else
                                    iLeepTime = giGetValidInt(moClass.moAppDB.Lookup("s.DeliveryTime", "tctContract AS p JOIN tctContractLine AS s ON s.ContractKey = p.ContractKey", "s.ItemKey =" & lLookupKey & " and p.ContractKey =" & lContractKey))
                                    sDate = sGetDate(DateAdd("d", iLeepTime, calDate.Value))
                                    gGridUpdateCell grdReqLines, lRow, kcolReqRequestDate, sDate
                                    gGridUpdateCell grdReqLineDtl, lRow, kColChildRequestDate, sDate
                                End If
                            End If
                        End If
                    End If
                    'Agregado por Multiconsulting
                    If Len(Trim(sLookupID)) > 0 And lLookupKey = 0 Then
                        If (Not mbIsPressF5) And Me.ActiveControl.Name <> "navItemGrid" Then
                            bMsgDisplayed = True
                            mbIsInvalid = True
    
                            sLocalOldItemID = msOldItemID
                            grdReqLines.Row = 0
                            ' Using striaght MsgBox because of a bug with the SotaMsgBox not remaining Modal if the
                            ' msg box is triggered when focus leaves the form.
                            'Modificado por Multiconsuting
                            If mbIntegratedCT And lContractKey > 0 Then
                                sMessage = "El artículo debe existir en el Contrato asignado"
                            Else
                                sMessage = gsBuildMessage(kmsgARBadField, mlLanguage, moClass.moAppDB, grdReqLines.Text)
                            End If
                            'Modificado por Multiconsuting
                            MsgBox sMessage, vbOKOnly, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)  'frmRequistn.Caption
                            grdReqLines.Row = lRow
                            gGridUpdateCell grdReqLines, lRow, lCol, sLocalOldItemID
                            'QQ, fixed #12562. set focus back to the cell.
                            HideAllNavs
                            gGridSetActiveCell grdReqLines, lRow, lCol
                            navItemGrid.Visible = True
    
                            bMsgDisplayed = False
                            mbDontSave = True
                        End If
                    Else
                        If sLookupID <> msOldItemID Then
                            sWhseID = Trim(gsGridReadCell(grdReqLines, lRow, kColReqWhseID))
                            lWhseKey = glGetValidLong(gsGridReadCell(grdReqLines, lRow, kColReqWhseKey))
                            If lRow = grdReqLines.MaxRows Then
                                sWhseID = Trim(lkuWarehouse)
                                If sWhseID <> "" Then
                                    lWhseKey = moDmForm.GetColumnValue("DfltShipToWhseKey")
                                End If
                            End If
                            If sWhseID <> "" And sLookupID <> "" Then
                                If Not bIsValidInventoryItem(lLookupKey, lWhseKey) Then
                                    bMsgDisplayed = True
                                    mbIsInvalid = True
                                    sLocalOldItemID = msOldItemID
                                    grdReqLines.Row = 0
                                    ' Using striaght MsgBox because of a bug with the SotaMsgBox not remaining Modal if the
                                    ' msg box is triggered when focus leaves the form.
                                    If lRow = grdReqLines.MaxRows Then
                                        sMessage = gsBuildMessage(kmsgItemInvalidForWhseDflt, mlLanguage, moClass.moAppDB)
                                    Else
                                        sMessage = gsBuildMessage(kmsgItemInvalidForWhse, mlLanguage, moClass.moAppDB)
                                    End If
                                    MsgBox sMessage, vbOKOnly, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)  'frmRequistn.Caption
                                    grdReqLines.Row = lRow
                                    gGridUpdateCell grdReqLines, lRow, lCol, sLocalOldItemID
                                    bMsgDisplayed = False
                                    mbDontSave = True
                                Else
                                    Set moItem = moIMSClass.Items(lLookupKey)
                                    moDmGrid.SetColumnValue lRow, "ItemKey", IIf(lLookupKey = 0, "", lLookupKey)
                                    bLoadItemDflts lLookupKey, lRow
                                End If
                            Else
                                Set moItem = moIMSClass.Items(lLookupKey)
                                 moDmGrid.SetColumnValue lRow, "ItemKey", IIf(lLookupKey = 0, "", lLookupKey)
                                 bLoadItemDflts lLookupKey, lRow
                            End If
                        End If
                    End If
                    msOldItemID = Trim(gsGridReadCellText(grdReqLines, lRow, lCol))
                    Set moItem = moIMSClass.Items(lLookupKey)
                    '   MVB - Check if item is valid for specified whse
                    moGM.LastCellValue = gsGridReadCellText(grdReqLines, lRow, lCol)
                    'Agregado por multiconsulting
                    If mbIntegratedCT Then
                        calDate_LostFocus
                        lContractKey = lGetReqContract
                        If lContractKey > 0 And lLookupKey > 0 Then
                            gGridUpdateCell grdReqLines, lRow, kColReqUnitCost, dGetContractCost(lContractKey, lLookupKey)
                            gGridUpdateCell grdReqLines, lRow, kColReqVendKey, lGetContractVendor(lContractKey)
                            gGridUpdateCell grdReqLines, lRow, kcolReqVendID, gsGetValidStr(moClass.moAppDB.Lookup("VendID", _
                                            "tapVendor", "VendKey=" & lGetContractVendor(lContractKey)))
                                            bCalcLineAmts lRow, kColReqQtyRequested
                        End If
                    End If
                    'Agregado por multiconsulting
                Case kColReqSTaxClassID
                
                    If (Not mbIsPressF5) And Me.ActiveControl.Name <> "navSTaxGrid" Then
                        sLookupID = gsGridReadCell(grdReqLines, lRow, lCol)
            
                        lLookupKey = gvCheckNull(moClass.moAppDB.Lookup _
                                                ("STaxClassKey", "tciSTaxClass", "STaxClassID = " & gsQuoted(sLookupID)), SQL_INTEGER)
            
                        If Len(Trim(sLookupID)) > 0 And lLookupKey = 0 Then
                            bMsgDisplayed = True
                            mbIsInvalid = True
                            grdReqLines.Row = 0
                            ' Using striaght MsgBox because of a bug with the SotaMsgBox not remaining Modal if the
                            ' msg box is triggered when focus leaves the form.
                            sMessage = gsBuildMessage(kmsgARBadField, mlLanguage, moClass.moAppDB, grdReqLines.Text)
                            MsgBox sMessage, vbOKOnly, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)  'frmRequistn.Caption
                            grdReqLines.Row = lRow
                            gGridUpdateCell grdReqLines, lRow, lCol, msOldSTaxClassID
                            HideAllNavs
                            gGridSetActiveCell grdReqLines, lRow, lCol
                            navSTaxGrid.Visible = True
                            bMsgDisplayed = False
                            mbDontSave = True
                        Else
                            If sLookupID <> msOldSTaxClassID Then
                                moDmGrid.SetColumnValue lRow, "STaxClassKey", IIf(lLookupKey = 0, "", lLookupKey)
                            End If
                        End If
    
                        msOldSTaxClassID = Trim(gsGridReadCellText(grdReqLines, lRow, lCol))
                    End If
                    
                Case kcolReqVendID
                    
                   
                       sLookupID = gsGridReadCell(grdReqLines, lRow, lCol)
            
                        lLookupKey = gvCheckNull(moClass.moAppDB.Lookup _
                                                ("VendKey", "tapVendor", "VendID = " & gsQuoted(sLookupID) & " AND CompanyID = " & _
                                                gsQuoted(msCompanyID)), SQL_INTEGER)
                        'Agregado por Multiconsuting
                        If mbIntegratedCT Then
                            lContractKey = lGetReqContract
                            If lContractKey > 0 And lLookupKey > 0 Then
                                If lLookupKey <> lGetContractVendor(lContractKey) Then
                                    lLookupKey = 0
                                End If
                            End If
                        End If
                        'Agregado por Multiconsuting
                     If (Not mbIsPressF5) And Me.ActiveControl.Name <> "navVendorGrid" Then
                        If Len(Trim(sLookupID)) > 0 And lLookupKey = 0 Then
                            bMsgDisplayed = True
                            mbIsInvalid = True
                            grdReqLines.Row = 0
                            ' Using striaght MsgBox because of a bug with the SotaMsgBox not remaining Modal if the
                            ' msg box is triggered when focus leaves the form.
                            'Modificado por Multiconsuting
                            If mbIntegratedCT And lContractKey > 0 Then
                                sMessage = "El proveedor debe coincidir con el del Contrato asignado"
                            Else
                                sMessage = gsBuildMessage(kmsgARBadField, mlLanguage, moClass.moAppDB, grdReqLines.Text)
                            End If
                            'Modificado por Multiconsuting
                            MsgBox sMessage, vbOKOnly, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)   'frmRequistn.Caption
                            grdReqLines.Row = lRow
                            gGridUpdateCell grdReqLines, lRow, lCol, msOldVendID
                            HideAllNavs
                            gGridSetActiveCell grdReqLines, lRow, lCol
                            navVendorGrid.Visible = True
                            bMsgDisplayed = False
                            mbDontSave = True
                        Else
                            If sLookupID <> msOldVendID Then
                                moDmGrid.SetColumnValue lRow, "VendKey", IIf(lLookupKey = 0, "", lLookupKey)
                                bLoadDfltCurrency lLookupKey, lRow
                            End If
                        End If
    
                    ElseIf mbIsPressF5 Then
                        If Len(Trim(sLookupID)) > 0 And lLookupKey <> 0 Then
                                moDmGrid.SetColumnValue lRow, "VendKey", IIf(lLookupKey = 0, "", lLookupKey)
                                bLoadDfltCurrency lLookupKey, lRow
                        End If
                    End If
                    msOldVendID = Trim(gsGridReadCellText(grdReqLines, lRow, lCol))
                    
'Modificado por Multiconsulting Osmel Barreras
                    Case kColReqBuyerID
                    
                    If (Not mbIsPressF5) And Me.ActiveControl.Name <> "navBuyerGrid" Then
                       sLookupID = gsGridReadCell(grdReqLines, lRow, lCol)
            
                        lLookupKey = gvCheckNull(moClass.moAppDB.Lookup _
                                                ("BuyerKey", "timBuyer", "BuyerID = " & gsQuoted(sLookupID) & " AND CompanyID = " & _
                                                gsQuoted(msCompanyID)), SQL_INTEGER)
            
                        If Len(Trim(sLookupID)) > 0 And lLookupKey = 0 Then
                            bMsgDisplayed = True
                            mbIsInvalid = True
                            grdReqLines.Row = 0
                            ' Using striaght MsgBox because of a bug with the SotaMsgBox not remaining Modal if the
                            ' msg box is triggered when focus leaves the form.
                            sMessage = gsBuildMessage(kmsgARBadField, mlLanguage, moClass.moAppDB, grdReqLines.Text)
                            MsgBox sMessage, vbOKOnly, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)   'frmRequistn.Caption
                            grdReqLines.Row = lRow
                            gGridUpdateCell grdReqLines, lRow, lCol, msOldBuyerID
                            HideAllNavs
                            gGridSetActiveCell grdReqLines, lRow, lCol
                            navBuyerGrid.Visible = True
                            bMsgDisplayed = False
                            mbDontSave = True
                        Else
                            If sLookupID <> msOldBuyerID Then
                                moDmGrid.SetColumnValue lRow, "ReqLineBuyerKey", IIf(lLookupKey = 0, "", lLookupKey)
                            End If
                        End If
    
                        msOldBuyerID = Trim(gsGridReadCellText(grdReqLines, lRow, lCol))
                    End If
' Modificado por Multiconsulting Osmel Barreras
                Case kColReqPurchDeptID
                    
                    If (Not mbIsPressF5) And Me.ActiveControl.Name <> "navDeptGrid" Then
                        sLookupID = gsGridReadCell(grdReqLines, lRow, lCol)
                
                        lLookupKey = gvCheckNull(moClass.moAppDB.Lookup _
                                                ("PurchDeptKey", "tpoPurchDepartment", "PurchDeptID = " & gsQuoted(sLookupID) & " AND CompanyID = " & _
                                                gsQuoted(msCompanyID)), SQL_INTEGER)
            
                        If Len(Trim(sLookupID)) > 0 And lLookupKey = 0 Then
                            bMsgDisplayed = True
                            mbIsInvalid = True
    
                            grdReqLines.Row = 0
                            ' Using striaght MsgBox because of a bug with the SotaMsgBox not remaining Modal if the
                            ' msg box is triggered when focus leaves the form.
                            sMessage = gsBuildMessage(kmsgARBadField, mlLanguage, moClass.moAppDB, grdReqLines.Text)
                            MsgBox sMessage, vbOKOnly, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)  'frmRequistn.Caption
                            grdReqLines.Row = lRow
                            gGridUpdateCell grdReqLines, lRow, lCol, msOldPurchDeptID
                            HideAllNavs
                            gGridSetActiveCell grdReqLines, lRow, lCol
                            navDeptGrid.Visible = True
                            bMsgDisplayed = False
                            mbDontSave = True
                        Else
                            
                            If sLookupID <> msOldPurchDeptID And grdReqLines.MaxRows = lRow Then
                                moDmGrid.SetColumnValue lRow, "Description", ""
                            End If
                            gGridUpdateCell grdReqLines, lRow, kColReqPurchDeptKey, IIf(lLookupKey = 0, "", lLookupKey)
                            'gGridUpdateCell grdReqLineDtl, 1, kColChildPurchDeptKey, IIf(lLookupKey = 0, "", lLookupKey)
                            moDMSubGrid.SetColumnValue 1, "PurchDeptKey", IIf(lLookupKey = 0, "", lLookupKey)
                            gGridLockCell grdReqLines, kColReqWhseID, lRow
    
                        End If
    
                        msOldPurchDeptID = Trim(gsGridReadCellText(grdReqLines, lRow, lCol))
                        'write the key to the parent grid.
                        gGridUpdateCell grdReqLineDtl, 1, kColChildPurchDeptKey, _
                                        gsGridReadCell(grdReqLines, lRow, kColReqPurchDeptKey)
                        moDMSubGrid.SetRowDirty 1
                        moDmGrid.SetRowDirty lRow
                    End If
                    
                Case kColReqWhseID
                
                    If (Not mbIsPressF5) And Me.ActiveControl.Name <> "navWhseGrid" Then
                        sLookupID = gsGridReadCell(grdReqLines, lRow, lCol)
                
                        lLookupKey = gvCheckNull(moClass.moAppDB.Lookup _
                                                ("WhseKey", "timWarehouse", "WhseID = " & gsQuoted(sLookupID) & " AND CompanyID = " & _
                                                gsQuoted(msCompanyID) & " AND Transit = 0"), SQL_INTEGER)
                        If Len(Trim(sLookupID)) > 0 And lLookupKey = 0 Then
                            bMsgDisplayed = True
                            mbIsInvalid = True
                            grdReqLines.Row = 0
                            ' Using striaght MsgBox because of a bug with the SotaMsgBox not remaining Modal if the
                            ' msg box is triggered when focus leaves the form.
                            sMessage = gsBuildMessage(kmsgARBadField, mlLanguage, moClass.moAppDB, grdReqLines.Text)
                            MsgBox sMessage, vbOKOnly, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession) ' frmRequistn.Caption
                            grdReqLines.Row = lRow
                            gGridUpdateCell grdReqLines, lRow, lCol, msOldWhseID
                            HideAllNavs
                            gGridSetActiveCell grdReqLines, lRow, lCol
                            navWhseGrid.Visible = True
                            bMsgDisplayed = False
                            mbDontSave = True
                        Else
                            
                            If sLookupID <> msOldWhseID And grdReqLines.MaxRows = lRow Then
                                moDmGrid.SetColumnValue lRow, "Description", ""
                            End If
                            If Trim(gsGridReadCell(grdReqLines, lRow, kColReqItemID)) <> "" Then
                                If Not bIsValidInventoryItem(glGetValidLong(gsGridReadCell(grdReqLines, lRow, kColReqItemKey)), lLookupKey) Then
                                    bMsgDisplayed = True
                                    mbIsInvalid = True
                                   grdReqLines.Row = 0
                                    ' Using striaght MsgBox because of a bug with the SotaMsgBox not remaining Modal if the
                                    ' msg box is triggered when focus leaves the form.
                                    sMessage = gsBuildMessage(kmsgItemNotvalidForWhse, mlLanguage, moClass.moAppDB)
                                    If MsgBox(sMessage, vbYesNo, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)) = vbYes Then
                                        grdReqLines.Row = lRow
                                        gGridUpdateCell grdReqLines, lRow, kColReqItemID, ""
                                        gGridUpdateCell grdReqLines, lRow, kColReqDescription, ""
                                        moDmGrid.SetColumnValue lRow, "ItemKey", ""
                                        msOldItemID = Trim(gsGridReadCellText(grdReqLines, lRow, lCol))
                                        gGridUpdateCell grdReqLines, lRow, kColReqWhseKey, IIf(lLookupKey = 0, "", lLookupKey)
                                        moDMSubGrid.SetColumnValue 1, "ShipToWhseKey", IIf(lLookupKey = 0, "", lLookupKey)
                                        gGridLockCell grdReqLines, kColReqPurchDeptID, lRow
                                    Else
                                        grdReqLines.Row = lRow
                                        gGridUpdateCell grdReqLines, lRow, lCol, msOldWhseID
                                    End If
                                    bMsgDisplayed = False
                                    mbDontSave = True
                                Else
                                    gGridUpdateCell grdReqLines, lRow, kColReqWhseKey, IIf(lLookupKey = 0, "", lLookupKey)
                                    moDMSubGrid.SetColumnValue 1, "ShipToWhseKey", IIf(lLookupKey = 0, "", lLookupKey)
                                    gGridLockCell grdReqLines, kColReqPurchDeptID, lRow
                                End If
                            Else
                                gGridUpdateCell grdReqLines, lRow, kColReqWhseKey, IIf(lLookupKey = 0, "", lLookupKey)
                                moDMSubGrid.SetColumnValue 1, "ShipToWhseKey", IIf(lLookupKey = 0, "", lLookupKey)
                                gGridLockCell grdReqLines, kColReqPurchDeptID, lRow
                            End If
                            moDmGrid.SetRowDirty lRow
                        End If
    
                        msOldWhseID = Trim(gsGridReadCellText(grdReqLines, lRow, lCol))
                        'write the key to the parent grid.
                        gGridUpdateCell grdReqLineDtl, 1, kColChildWhseKey, _
                                        gsGridReadCell(grdReqLines, lRow, kColReqWhseKey)
                        moDmGrid.SetRowDirty lRow
                        moDMSubGrid.SetRowDirty 1
                        
                    End If
                    
            End Select
        
        End With
    End If
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "moGM_CellChange", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Function bIsValidInventoryItem(lItemKey As Long, lWhseKey As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim rs As Object
    Dim sSql As String
    
    bIsValidInventoryItem = False
    'Inventory Item with Active status
    sSql = "SELECT WhseKey FROM timInventory WHERE Status = 1 AND WhseKey = "
    sSql = sSql & lWhseKey
    sSql = sSql & " AND ItemKey = " & lItemKey
    Set rs = oclass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
    If rs.IsEOF Then
'        giSotaMsgBox Nothing, moClass.moSysSession, kmsgARBadField, _
'            gsStripChar(lblWarehouse.Caption, "&")
        bIsValidInventoryItem = False
    Else
        bIsValidInventoryItem = True
    End If
    Set rs = Nothing
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidInventoryItem", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++

End Function

Private Function gvCheckNull(vField, Optional vDataType As Variant, Optional vSetEmpty As Variant) As Variant
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    If IsMissing(vDataType) Then
      vDataType = SQL_CHAR
    End If
    
    If IsMissing(vSetEmpty) Then
      vSetEmpty = False
    End If
    
    If IsEmpty(vField) Or Len(vField) = 0 Then
        If vSetEmpty Then
            gvCheckNull = Empty
        Else
            If vDataType = SQL_CHAR Then
                gvCheckNull = ""
            Else
                gvCheckNull = 0
            End If
       End If
    Else
        If IsNull(vField) Then
            If vSetEmpty Then
                gvCheckNull = Empty
            Else
                If vDataType = SQL_CHAR Then
                    gvCheckNull = ""
                Else
                    gvCheckNull = 0
                End If
            End If
        Else
            gvCheckNull = vField
        End If
    End If
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "gvCheckNull", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function




Private Sub lkuWarehouse_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus lkuWarehouse, True
    #End If
'+++ End Customizer Code Push +++
    bIsValidWhse
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lkuWarehouse_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++

End Sub


Private Sub lkuDept_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus lkuDept, True
    #End If
'+++ End Customizer Code Push +++
    bIsValidPurchDept
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lkuDept_LostFocus", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++



End Sub

Private Sub lkuWarehouse_LookupClick(bCancel As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then
            bCancel = moFormCust.OnLookupClick(lkuWarehouse, True)
            If bCancel Then Exit Sub
        End If
    #End If
'+++ End Customizer Code Push +++

        bCancel = Trim(lkuMain) = ""
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lkuWarehouse_LookupClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++

End Sub


Private Sub lkuDept_LookupClick(bCancel As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then
            bCancel = moFormCust.OnLookupClick(lkuDept, True)
            If bCancel Then Exit Sub
        End If
    #End If
'+++ End Customizer Code Push +++

        bCancel = Trim(lkuMain) = ""
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lkuDept_LookupClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++

End Sub

Private Sub lkuMain_LookupClick(bCancel As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then
            bCancel = moFormCust.OnLookupClick(lkuMain, True)
            If bCancel Then Exit Sub
        End If
    #End If
'+++ End Customizer Code Push +++
    If iConfirmUnload(True) = kDmFailure Then
        bCancel = True
    End If
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lkuMain_LookupClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

'********************************************************************************************************************************
' Agregado por Multiconsulting Osmel Barreras

Private Sub navBuyerGrid_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lActiveRow As Long
    lActiveRow = glGridGetActiveRow(grdReqLines)
    txtNavReturn = gsGridReadCellText(grdReqLines, lActiveRow, kColReqBuyerID)
    
    If txtNavReturn = "" Then
        mbIsPressF5 = False
    End If
    
    gcLookupClick Me, navBuyerGrid, txtNavReturn, "BuyerID"
    moGM.LookupClicked
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "navBuyerGrid_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
' Agregado por MultiConsulting Osmel Barreras
'********************************************************************************************************************************

Private Sub navVendorGrid_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lActiveRow As Long

    lActiveRow = glGridGetActiveRow(grdReqLines)
    txtNavReturn = gsGridReadCellText(grdReqLines, lActiveRow, kcolReqVendID)
    
    gcLookupClick Me, navVendorGrid, txtNavReturn, "VendID"
    moGM.LookupClicked
    
    'Call validation for grid cell change
    moGM_CellChange lActiveRow, kcolReqVendID
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "navVendorGrid_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub navWhseGrid_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    'by QQ to fix #7872, moGridNav.LookupClicked will trigger grdReqLines_GotFocus
    'in run time(run from exe). set this flag to skip it.
    
    Dim lActiveRow As Long

    lActiveRow = glGridGetActiveRow(grdReqLines)
    txtNavReturn = gsGridReadCellText(grdReqLines, lActiveRow, kColReqWhseID)
    
    mbSkipGotFocus = True
    gcLookupClick Me, navWhseGrid, txtNavReturn, "WhseID"
    moGM.LookupClicked
    mbSkipGotFocus = False
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "navWhseGrid_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub navDeptGrid_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    Dim lActiveRow As Long

    lActiveRow = glGridGetActiveRow(grdReqLines)
    txtNavReturn = gsGridReadCellText(grdReqLines, lActiveRow, kColReqPurchDeptID)
    
    gvLookupClick Me, navDeptGrid, txtNavReturn
    moGM.LookupClicked
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "navDeptGrid_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub


Private Sub navItemGrid_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    Dim lActiveRow As Long

    lActiveRow = glGridGetActiveRow(grdReqLines)
    '*************************************************************************************************
'Modificado por Multiconsulting Osmel Barreras
    If sddEstatusReq.ItemData <= kReqStatusPending Then
            
        txtNavReturn = gsGridReadCellText(grdReqLines, lActiveRow, kColReqItemID)
    
        gcLookupClick Me, navItemGrid, txtNavReturn, "ItemID"
        moGM.LookupClicked
        
    ElseIf (sddEstatusReq.ItemData = kReqStatusAccepted And (gsGridReadCellText(grdReqLines, lActiveRow, kColReqItemID) = "" And gsGridReadCellText(grdReqLines, lActiveRow, kColReqDescription) <> "")) Then
        
        txtNavReturn = gsGridReadCellText(grdReqLines, lActiveRow, kColReqItemID)
    
        gcLookupClick Me, navItemGrid, txtNavReturn, "ItemID"
        moGM.LookupClicked
        gGridLockColumn grdReqLines, kColReqEstPres
        gGridLockColumn grdReqLines, kColReqQtyRequested
        gGridUpdateCellText grdReqLines, lActiveRow, kColReqQtyRequested, msPersistCantReq
        gGridUpdateCellText grdReqLineDtl, lActiveRow, kColChildQtyReq, msPersistCantReq
    End If
'Modificado por Multiconsulting Osmel Barreras
'*************************************************************************************************
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "navItemGrid_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub navSTaxGrid_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    Dim lActiveRow As Long

    lActiveRow = glGridGetActiveRow(grdReqLines)
    txtNavReturn = gsGridReadCellText(grdReqLines, lActiveRow, kColReqSTaxClassID)
    
    gvLookupClick Me, navSTaxGrid, txtNavReturn
    moGM.LookupClicked
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "navSTaxGrid_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

'************************************************************************
'   Description:
'       standard Sage MAS 500 shutdown procedure.  Unload all child forms
'       and remove any objects created within this app.
'
'   Param:
'       <none>
'
'   Returns:
'
'************************************************************************

Public Sub PerformCleanShutDown()
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
    On Error GoTo ExpectedErrorRoutine
    On Error Resume Next
    
    'Unload all forms loaded from this main form
    gUnloadChildForms Me
         
    'remove all child collections
    giCollectionDel moClass.moFramework, moSotaObjects, -1
    
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub

ExpectedErrorRoutine:
'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "PerformCleanShutDown", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub



'************************************************************************
'   Description:
'       process form resize event
'
'   Param:
'       <none>
'
'   Returns:
'
'************************************************************************

Private Sub Form_Resize()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then

      'resize Height
       gResizeForm kResizeDown, Me, miOldFormHeight, miMinFormHeight, grdReqLines
          
      'resize Width
        gResizeForm kResizeRight, Me, miOldFormWidth, miMinFormWidth, grdReqLines
               
' Keep the Generate button centered.
        cmdGenerate.Left = frmRequistn.Width / 2 - (cmdGenerate.Width / 2)
        
        miOldFormHeight = Me.Height
        miOldFormWidth = Me.Width
        
' Redraw the grid navigators if necessary
        If grdReqLines.ActiveRow > 0 Then
            moGM.Grid_ColWidthChange grdReqLines.ActiveCol
        End If
        

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

'************************************************************************
'   Description:
'       unload all objects.  removed references to objects created
'       in this app.
'
'   Param:
'       Cancel - used to cancel the unload
'
'   Returns:
'
'************************************************************************

Private Sub Form_Unload(Cancel As Integer)
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        moFormCust.UnloadSelf
        Set moFormCust = Nothing
    End If
#End If

    lkuMain.Terminate
    lkuDept.Terminate
    lkuWarehouse.Terminate
    moIMSClass.Terminate
    Set moIMSClass = Nothing
    
    Set moItem = Nothing
    cboExpReason.Terminate

    On Error Resume Next
    moDmForm.UnloadSelf
    moDmGrid.UnloadSelf
    moDMSubGrid.UnloadSelf
    moGM.UnloadSelf


    Set moDmForm = Nothing
    Set moDmGrid = Nothing
    Set moDMSubGrid = Nothing
    Set moGM = Nothing
    Set moContextMenu = Nothing
    Set moToolbar = Nothing
    Set moGM = Nothing
    Set moOptions = Nothing
    Set moReportObj = Nothing
    Set moDBObj = Nothing
    Set moDDData = Nothing

    'Unload Custom field object.
    If Not moUF Is Nothing Then
        moUF.UnloadMe
        Set moUF = Nothing
    End If

'    Set frmRequistn = Nothing
    
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub

'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
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

'************************************************************************
'   Description:
'       if form is dirty, prompt for save.  handle cancel of the
'       shutdown or process a normal shutdown.
'
'   Param:
'       Cancel -        set to True to cancel the form shutdown
'       UnloadMode -    flag indicating type of shutdown requested
'
'   Returns:
'
'************************************************************************

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
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Sub
        End If
    End If
#End If
    
    Dim i               As Integer
    Dim iConfirmUnloadVar  As Integer
    Dim iInitialState   As Integer
    Dim bOldMb          As Boolean
    
    bOldMb = mbDontChkclick
    

   'Reset the CancelShutDown flag if prior shutdowns were canceled.
    mbCancelShutDown = False

    ' Line Entry hook
        
'    If Not bIsValidDirtyCheck Then GoTo CancelShutDown
    If moClass.mlError = 0 Then
        If miSecurityLevel = kSecLevelDisplayOnly Then
            iConfirmUnloadVar = kDmSuccess
        Else
            mbDontChkclick = True
            iConfirmUnloadVar = iConfirmUnload(True)
            mbDontChkclick = bOldMb
        End If
    
        Select Case iConfirmUnloadVar
            Case kDmSuccess
                'Do Nothing

            Case kDmFailure
                GoTo CancelShutDown

            Case kDmError
                GoTo CancelShutDown
                'If you need data Manager Error Value then you would
                'dimension a variable, such as lError as Long and assign
                'to the DataManager Error property.
                'lError = moDmFormLog.Error

            Case Else
                giSotaMsgBox Me, moClass.moSysSession, kmsgUnexpectedConfirmUnloadRV, iConfirmUnloadVar

        End Select
    
        'Check all other forms  that may have been loaded from this main form.
        'If there are any Visible forms, then this means the form is Active.
        'Therefore, cancel the shutdown.
        If gbActiveChildForms(Me) Then GoTo CancelShutDown
    
        Select Case UnloadMode
            Case vbFormCode
            'Do Nothing.
            'If the unload is caused by form code, then the form
            'code should also have the miShutDownRequester set correctly.
            
            Case Else
                If mlRunMode = kContextDD Then
                                        
                    mbDontChkclick = True
                    moDmForm.Action kDmCancel
                    ClearForm
                    mbDontChkclick = bOldMb
                    Me.Hide
                    GoTo CancelShutDown
                End If
                moClass.miShutDownRequester = kUnloadSelfShutDown
        End Select

    End If
    
    If Not moReportObj Is Nothing Then
        If Not moReportObj.bShutdownEngine Then
            GoTo CancelShutDown
        End If
    End If

    
    'If execution gets to this point, the form and class object of the form
    'will be shut down.  Perform all operations necessary for a clean shutdown.
    PerformCleanShutDown
    
    Select Case moClass.miShutDownRequester
        Case kUnloadSelfShutDown
            moClass.moFramework.UnloadSelf EFW_TF_MANSHUTDN
            Set moClass.moFramework = Nothing

        Case Else 'kFrameworkShutDown
            'Do nothing

    End Select
    
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub

CancelShutDown:
    moClass.miShutDownRequester = kFrameworkShutDown
    mbCancelShutDown = True
    Cancel = True
    
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
'************************************************************************
'   Description:
'       returns form name for debugging information
'
'   Param:
'       <none>
'
'   Returns:
'       form name
'
'************************************************************************

Private Function sMyName() As String
'+++ VB/Rig Skip +++
    sMyName = Me.Name
End Function
'************************************************************************
'   oClass contains the reference to the parent class object.  The form
'   needs this reference to use the public variables created within the
'   class object.
'************************************************************************
Public Property Get oclass() As Object
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Set oclass = moClass
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "oClass_Get", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Property
        End Select
'+++ VB/Rig End +++
End Property

Public Property Set oclass(oNewClass As Object)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Set moClass = oNewClass
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "oClass_Set", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Property
        End Select
'+++ VB/Rig End +++
End Property


Private Sub grdReqLines_Change(ByVal Col As Long, ByVal Row As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    'Agregado por Multiconsulting
    Dim lLookupKey    As Long
    Dim sLookupID     As String
    Dim sLocalOldItemID As String
    
    Dim lItemKey As Long
    Dim lContractKey As Long
    Dim dMaxQty As Double
    
    If gsGridReadCellText(grdReqLines, Row, kColReqItemID) = "" Then
        msPersistDescriptionReq = gsGridReadCellText(grdReqLines, Row, kColReqDescription)
    End If
    
    msPersistCantReq = gsGridReadCellText(grdReqLines, Row, kColReqQtyRequested)
    msPersistPresEstReq = gsGridReadCellText(grdReqLines, Row, kColReqEstPres)
    'Agregado por Multiconsulting
    
        If Not moGM Is Nothing Then
        moGM.Grid_Change Col, Row

        Select Case Col
            Case kColReqItemID ' if the item is cleared, clear the item key
                If Trim(gsGridReadCell(grdReqLines, Row, Col)) = "" Then
                    gGridUpdateCell grdReqLines, Row, kColReqItemKey, ""
                End If
'**************************************************************************************************************************************************************
'************ Agregado por MultiConsulting Cuba Osmel Barreras*************
            If Col = kColReqItemID And (msOldItemID = "" Or gsGridReadCellText(grdReqLines, Row, kColReqItemID) <> msOldItemID) And sddEstatusReq.ItemData = kReqStatusAccepted Then
                gGridUpdateCell grdReqLines, Row, kColReqItemID, msOldItemID
            End If
            Case kColReqBuyerID  ' if the buyer is cleared, clear the buyer key
                If Trim(gsGridReadCell(grdReqLines, Row, Col)) = "" Then
                    gGridUpdateCell grdReqLines, Row, kColReqBuyerkey, ""
                End If
                moDMSubGrid.SetRowDirty 1
                
            Case kColReqLStItemID ' if the state item is cleared, clear the state item key
                If Trim(gsGridReadCell(grdReqLines, Row, Col)) = "" Then
                    gGridUpdateCell grdReqLines, Row, kColReqLStItemKey, ""
                End If

                'Get the New Default Bill To value.
                lLookupKey = grdReqLines.TypeComboBoxCurSel   '0-Lisit, 1-Dict Tecn, 2-Legal, 3-Com Cont, 4-Intel Com
                sLookupID = gsGridReadCellText(grdReqLines, Row, Col)

                If sLookupID <> msOldStateItemID Then
                    moDmGrid.SetColumnValue Row, "StateBIKey", lLookupKey
                End If

                msOldStateItemID = Trim(gsGridReadCellText(grdReqLines, Row, Col))

                moDMSubGrid.SetRowDirty 1

            Case kColReqLTBItemID ' if the type buy is cleared, clear the type buy key
                If Trim(gsGridReadCell(grdReqLines, Row, Col)) = "" Then
                    gGridUpdateCell grdReqLines, Row, kColReqLTBItemKey, ""
                End If
                
                'Get the New Default Bill To value.
                lLookupKey = grdReqLines.TypeComboBoxCurSel   '0-Nacional, 1-Internacional
                sLookupID = gsGridReadCellText(grdReqLines, Row, Col)

                If sLookupID <> msOldTypeBuyItemID Then
                    moDmGrid.SetColumnValue Row, "TypeBIKey", lLookupKey
                End If

                msOldTypeBuyItemID = Trim(gsGridReadCellText(grdReqLines, Row, Col))

                moDMSubGrid.SetRowDirty 1
'************ Agregado por MultiConsulting Osmel Barreras
'********************************************************************************************************************************
        
'**************************************************************************************************************************************************************
 
            Case kcolReqVendID ' if the vendor is cleared, clear the vendor key
                If Trim(gsGridReadCell(grdReqLines, Row, Col)) = "" Then
                    gGridUpdateCell grdReqLines, Row, kColReqVendKey, ""
                End If
            Case kColReqPurchDeptID ' if the department is cleared, clear the dept key
            
                If Trim(gsGridReadCell(grdReqLines, Row, Col)) = "" Then
                    gGridUpdateCell grdReqLines, Row, kColReqPurchDeptKey, ""
                End If
                gGridUpdateCell grdReqLineDtl, 1, kColChildPurchDeptKey, _
                                gsGridReadCell(grdReqLines, Row, kColReqPurchDeptKey)
                moDMSubGrid.SetRowDirty 1
                
            Case kColReqWhseID ' if the warehouse is cleared, clear the whse key
            
                If Trim(gsGridReadCell(grdReqLines, Row, Col)) = "" Then
                    gGridUpdateCell grdReqLines, Row, kColReqWhseKey, ""
                End If
                gGridUpdateCell grdReqLineDtl, 1, kColChildWhseKey, _
                            gsGridReadCell(grdReqLines, Row, kColReqWhseKey)
                moDMSubGrid.SetRowDirty 1
                
           Case kcolReqRequestDate
           
               gGridUpdateCell grdReqLineDtl, 1, kColChildRequestDate, _
                            gsGridReadCell(grdReqLines, Row, Col)
                moDMSubGrid.SetRowDirty 1
                
            Case kColReqQtyRequested
                'Agregado por Multiconsulting
                If mbIntegratedCT Then
                    lContractKey = lGetReqContract
                    lItemKey = glGetValidLong(gsGridReadCell(grdReqLines, Row, kColReqItemKey))
                    If lItemKey <> 0 And lContractKey > 0 Then
                        dMaxQty = dGetMaxQtyAllowAtContract(lContractKey, lItemKey, glGetValidLong(moDmForm.GetColumnValue("ReqKey")), Row)
                        If dMaxQty < msPersistCantReq Then
                            msPersistCantReq = dMaxQty
                            gGridUpdateCell grdReqLines, Row, Col, msPersistCantReq
                            gGridUpdateCell grdReqLineDtl, Row, Col, msPersistCantReq
                            If dMaxQty > 0 Then
                                MsgBox "La Cantidad máxima restante en contrato es de " & dMaxQty, vbInformation, "Alerta"
                            Else
                                MsgBox "No quedan cantidades restantes por ejecutar en el contrato para este artículo", vbInformation, "Alerta"
                            End If
                        End If
                    End If
                End If
                'Agregado por Multiconsulting
                If bCalcLineAmts(Row, Col) Then
                    gGridUpdateCell grdReqLineDtl, 1, kColChildQtyReq, _
                                    gsGridReadCell(grdReqLines, Row, Col)
                    'Agregado por Multiconsulting
                    gGridUpdateCell grdReqLines, Row, kColReqEstPres, _
                                    Round(gdGetValidDbl(gsGridReadCell(grdReqLines, Row, kColReqUnitCost)) * msPersistCantReq, 2)
                    'Agregado por Multiconsulting
                    moDMSubGrid.SetRowDirty 1
                Else
                    ' If there is a problem in calculating, reset the value
                    gGridUpdateCell grdReqLines, Row, Col, msOldReqQtyRequested
                End If
                
            Case kColReqUnitCost
            
                If Not bCalcLineAmts(Row, Col) Then
                    ' If there is a problem in calculating, reset the value
                    gGridUpdateCell grdReqLines, Row, Col, gsStripChar(msOldReqUnitCost, ",")
                End If
                
             Case kColReqUnitCostExact
            
                If Not bCalcLineAmts(Row, Col) Then
                    ' If there is a problem in calculating UnitCostExact , reset the value to UnitCost
                    gGridUpdateCell grdReqLines, Row, Col, gsStripChar(msOldReqUnitCost, ",")
                End If
            Case kColReqExtAmt
            
                If Not bCalcLineAmts(Row, Col) Then
                    ' If there is a problem in calculating, reset the value
                    gGridUpdateCell grdReqLines, Row, Col, gsStripChar(msOldReqExtAmt, ",")
                End If

        End Select
    End If
    
    
    
    '============================================================================================================
    
    moGM_CellChange Row, Col
    
    '********************************************************************************************************************************
    ' Modificado por MultiConsulting Osmel Barreras
    RequisitionChanged = True
    lastRowMod = grdReqLines.ActiveRow
    
'    If grdReqLines.MaxRows > 2 Then
'        RemoveLastRowFromGrid
'    End If
        
    If listRows.Count + 1 < grdReqLines.DataRowCnt Or lastRowMod > listRows.Count Then
        Dim i As Integer
        For i = listRows.Count + 1 To grdReqLines.DataRowCnt
            listRows.Add 0, "r" & i
        Next
    End If
    
'    MsgBox listRows.Item("r" & grdReqLines.ActiveRow)
    
    If listRows.Item("r" & grdReqLines.ActiveRow) <> grdReqLines.ActiveRow Then
        listRows.Remove ("r" & grdReqLines.ActiveRow)
        listRows.Add grdReqLines.ActiveRow, "r" & grdReqLines.ActiveRow
'        MsgBox listRows.Item("r" & grdReqLines.ActiveRow)
    End If
    ' Modificado por MultiConsulting Osmel Barreras
    '********************************************************************************************************************************
'    Dim sNavID As String
'    Dim lItemKey As Long
'    Dim lNextCol As Long
'
'    If mbPrintingReq Then
''+++ VB/Rig Begin Pop +++
''+++ VB/Rig End +++
'        Exit Sub
'    End If
'
'     If Not moGM Is Nothing Then
'
'
'        If Row = 0 Then Exit Sub
'
'        'If the row in not associated with a PO, do the check for navigator field changes
'        If gsGridReadCellText(grdReqLines, Row, kColReqPOTranID) = "" Then
'
'            sNavID = Trim(gsGridReadCellText(grdReqLines, Row, Col))
'
'            Select Case Col
'                Case kColReqItemID
'                    If sNavID <> msOldItemID Then
'                        moGM_CellChange Row, Col
'                    End If
'                    msOldItemID = Trim(gsGridReadCellText(grdReqLines, Row, kColReqItemID))
'                    lItemKey = glGetValidLong(gsGridReadCellText(grdReqLines, Row, kColReqItemKey))
'                    Set moItem = moIMSClass.Items(lItemKey)
'
'                Case kColReqSTaxClassID
'                    If sNavID <> "" And _
'                        sNavID <> msOldSTaxClassID Then
'                            moGM_CellChange Row, Col
'                    Else
'                        msOldSTaxClassID = Trim(gsGridReadCellText(grdReqLines, Row, kColReqSTaxClassID))
'                    End If
'                Case kcolReqVendID
'                    If sNavID <> "" And _
'                        sNavID <> msOldVendID Then
'                        moGM_CellChange Row, Col
'                    Else
'                        msOldVendID = Trim(gsGridReadCellText(grdReqLines, Row, kcolReqVendID))
'                    End If
'                Case kColReqPurchDeptID
'                    If sNavID <> "" And _
'                        sNavID <> msOldPurchDeptID Then
'                            moGM_CellChange Row, Col
'                            ' Make sure the column is still filled in after being validated.
'                            If NewCol = kColReqWhseID And Row = NewRow And _
'                            Trim(gsGridReadCell(grdReqLines, Row, Col)) <> "" Then
'                                HideAllNavs
'                                lNextCol = kColReqSTaxClassID
'                                If bGridCellLocked(grdReqLines, lNextCol, Row) Then
'                                    lNextCol = kColReqComment
'                                End If
'                                gGridSetActiveCell grdReqLines, Row, lNextCol
'                                moGM.Grid_LeaveCell Col, Row, lNextCol, NewRow
'                            End If
'                    Else
'                        msOldPurchDeptID = Trim(gsGridReadCellText(grdReqLines, Row, kColReqPurchDeptID))
'                        If sNavID = "" And mbIntegrateWithIM Then
'                            ' Enable the warehouse column
'                            gGridUnlockCell grdReqLines, kColReqWhseID, Row
'                            lNextCol = kColReqSTaxClassID
'                            If bGridCellLocked(grdReqLines, lNextCol, Row) Then
'                                lNextCol = kColReqComment
'                            End If
'                            If NewRow = Row And NewCol = lNextCol Then
'                                HideAllNavs
'                                gGridSetActiveCell grdReqLines, Row, kColReqWhseID
'                                moGM.Grid_LeaveCell Col, Row, kColReqWhseID, NewRow
'                                moGM.Grid_LeaveCell Col, Row, kColReqWhseID, NewRow
'                            End If
'                        End If
'                    End If
'                Case kColReqWhseID
'                    If sNavID <> "" And _
'                        sNavID <> msOldWhseID Then
'                            moGM_CellChange Row, Col
'                            ' Make sure the column is still filled in after being validated.
'                            If NewCol = kColReqPurchDeptID And Row = NewRow And _
'                            Trim(gsGridReadCell(grdReqLines, Row, Col)) <> "" Then
'                                HideAllNavs
'                                gGridSetActiveCell grdReqLines, Row, kcolReqRequestDate
'                                moGM.Grid_LeaveCell Col, Row, kcolReqRequestDate, NewRow
'                                moGM.Grid_LeaveCell Col, Row, kcolReqRequestDate, NewRow
'                            End If
'                    Else
'                        msOldWhseID = Trim(gsGridReadCellText(grdReqLines, Row, kColReqWhseID))
'                        If sNavID = "" Then
'                            ' Enable the department column
'                            gGridUnlockCell grdReqLines, kColReqPurchDeptID, Row
'                            If NewRow = Row And NewCol = kcolReqRequestDate Then
'                                gGridSetActiveCell grdReqLines, Row, kColReqPurchDeptID
'                                moGM.Grid_LeaveCell Col, Row, kColReqPurchDeptID, NewRow
'                                moGM.Grid_LeaveCell Col, Row, kColReqPurchDeptID, NewRow
'                            End If
'                        End If
'                    End If
'            End Select
'        End If
'        If Row <> NewRow And NewRow > 0 Then
'            msOldItemID = Trim(gsGridReadCellText(grdReqLines, NewRow, kColReqItemID))
'            lItemKey = glGetValidLong(gsGridReadCellText(grdReqLines, NewRow, kColReqItemKey))
'            Set moItem = moIMSClass.Items(lItemKey)
'            msOldSTaxClassID = Trim(gsGridReadCellText(grdReqLines, NewRow, kColReqSTaxClassID))
'            msOldVendID = Trim(gsGridReadCellText(grdReqLines, NewRow, kcolReqVendID))
'            msOldPurchDeptID = Trim(gsGridReadCellText(grdReqLines, NewRow, kColReqPurchDeptID))
'            msOldWhseID = Trim(gsGridReadCellText(grdReqLines, NewRow, kColReqWhseID))
'        End If
'    End If
    
    
    
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_Change", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub grdReqLines_Click(ByVal Col As Long, ByVal Row As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    
    Dim lOldRowCnt As Long

    If Not moGM Is Nothing Then
        lOldRowCnt = grdReqLines.MaxRows
        moGM.Grid_Click Col, Row
        If grdReqLines.MaxRows > lOldRowCnt Then
            grdReqLines.MaxRows = lOldRowCnt
        End If
    End If



'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub grdReqLines_KeyDown(KeyCode As Integer, Shift As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moGM Is Nothing Then
        moGM.Grid_KeyDown KeyCode, Shift
    End If
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_KeyDown", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Public Function bGridCellLocked(Grid As Control, lCol As Long, lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'************************************************************************
'     Desc: Returns the locked status for a cell.
'    Parms: Grid    Control     Name of the grid control
'           lCol    long        Column number to unlock
'           lRow    long        Row number to unlock
'************************************************************************

    With Grid
        .Col = lCol: .Row = lRow
        bGridCellLocked = .Lock
    End With
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bGridCellLocked", VBRIG_IS_MODULE
        Err.Raise guSotaErr.Number
'+++ VB/Rig End +++
End Function


Private Sub grdReqLines_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Cancel = mbIsInvalid
    mbIsInvalid = False
    If Not Cancel Then
        Cancel = Not moGM.Grid_LeaveCell(Col, Row, NewCol, NewRow)

    End If
    

'+++ VB/Rig Begin Pop +++
        Exit Sub
Resume
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_LeaveCell", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub grdReqLines_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    If Not moGM Is Nothing Then
        moGM.Grid_LeaveRow Row, NewRow
    End If
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "grdReqLines_LeaveRow", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuMain_BeforeLookupReturn(colSQLReturnVal As Collection, bCancel As Boolean)

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then
            bCancel = moFormCust.OnBeforeLookupReturn(lkuMain, True)
            If bCancel Then Exit Sub
        End If
    #End If
'+++ End Customizer Code Push +++
    If bLoseFocus() Then
        '********************************************************************************
        'Agregado por MultiConsulting Osmel Barreras
        '   Check the user id typed in can change the Estatus Requisition.
                CheckUserStatusReqPermission
                firstCompEstatusReq = True
                adicInfo = True
                msPersistCantRows = grdReqLines.DataRowCnt
        'Agregado por MultiConsulting Osmel Barreras
        '********************************************************************************
        If colSQLReturnVal(1) <> lkuMain.Tag Then
            moDmForm.Clear True
            ClearForm
        End If
        If Not bIsValidReqNum Then
            
            If bFocusBack Then
                bSetFocus lkuMain
            End If
    
        End If
    
    Else
        
        bSetFocus lkuMain
    
    End If
    

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lkuMain_BeforeLookupReturn", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuMain_LostFocusText()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus lkuMain, True
    #End If
'+++ End Customizer Code Push +++
    If bLoseFocus() Then
        
        If Not bIsValidReqNum Then
            
            If bFocusBack Then
                bSetFocus lkuMain
            End If
        
        Else
            If txtOriginator.Enabled Then
                txtOriginator.SetFocus
            End If
            '*************************************************************************
            'Modificado Multiconsulting Osmel Barreras
                    chkbEstatusDesc.Enabled = True
            
                '   Check the user id typed in can change the StatusReq.
                    UpdateSegEstatusReq (False)
                    msPersistCantRows = grdReqLines.DataRowCnt
            'Modificado Multiconsulting Osmel Barreras
            '*************************************************************************
        End If
    
    Else
        
        bSetFocus lkuMain
    
    End If

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lkuMain_LostFocusText", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub



Public Function bSetFormCurrencyControls(sCurrID As String) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'****************************************************************************
' Desc: sets up the Currency Controls w decimal places & Currency Controls
'****************************************************************************
    Dim iDecPlaces      As Integer
    Static sLastCurrID  As String
    Dim uCurrInfo       As CurrencyInfo
    Static bDone        As Boolean
    
  'Currency Id same as last time do not change controls
    If sLastCurrID = sCurrID Then
        bSetFormCurrencyControls = True
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If
  
  'Get Current Number of decimal places
    iDecPlaces = miDigAfterDec
    
  'Set Return Code
    bSetFormCurrencyControls = False
    
  'Get the Currency attributes for this Currency ID
  'sSQL = "SELECT * FROM tmcCurrency WHERE CurrID = %s"
    If sCurrID = msHomeCurrID Then
        
        If bDone Then
        
            msCurrSymbol = msHomeCurrSymbol
            miDigAfterDec = miHomeDigAfterDec
            miRoundPrec = miHomeRoundPrec
            miRoundMeth = miHomeRoundMeth
            mlCurrencyLocale = mlHomeCurrencyLocale
        
        Else
        
          'Setup home currency (1st time)
            If gbGetCurrInfo(moClass, sCurrID, uCurrInfo) Then
                mlHomeCurrencyLocale = moClass.moSysSession.CurrencyLocale
                miHomeRoundPrec = uCurrInfo.iRoundPrecision
                miHomeRoundMeth = uCurrInfo.iRoundMeth
                msHomeCurrSymbol = uCurrInfo.sCurrSymbol
                miHomeDigAfterDec = uCurrInfo.iDecPlaces
            Else
              'An error occured assume US
                mlHomeCurrencyLocale = 1033
                miHomeRoundPrec = 1
                miHomeRoundMeth = 1
                msHomeCurrSymbol = "$"
                miHomeDigAfterDec = 2
            End If
        
            bDone = True
        
            msCurrSymbol = msHomeCurrSymbol
            miDigAfterDec = miHomeDigAfterDec
            miRoundPrec = miHomeRoundPrec
            miRoundMeth = miHomeRoundMeth
            mlCurrencyLocale = mlHomeCurrencyLocale
        
        End If
        
    Else
        
        If Not gbGetCurrInfo(moClass, sCurrID, uCurrInfo) Then
            msCurrSymbol = msHomeCurrSymbol
            miDigAfterDec = miHomeDigAfterDec
            miRoundPrec = miHomeRoundPrec
            miRoundMeth = miHomeRoundMeth
            mlCurrencyLocale = mlHomeCurrencyLocale
        Else
            mlCurrencyLocale = mlHomeCurrencyLocale
            miRoundPrec = uCurrInfo.iRoundPrecision
            miRoundMeth = uCurrInfo.iRoundMeth
            msCurrSymbol = uCurrInfo.sCurrSymbol
            miDigAfterDec = uCurrInfo.iDecPlaces
        End If
        
    End If
    
    
  'Setup form currency controls
   
'    gbSetCurrCtls moClass, sCurrID, uCurrInfo, curPurchAmt(0), curPurchAmt(1), curPurchAmt(2), _
'        curFreightAmt, curExtAmt, curSTaxAmt, curTranAmt, curAmtRcvd, curOpenAmt, _
'        curInvcdAmt
'
'    'Set grid Alignment Types
'    If mbTrackTaxes Then
'        gGridSetColumnType frmSalesTaxDetail.grdSalesTaxDetail, kColTaxCodeTxblPurchAmt, _
'                SS_CELL_TYPE_FLOAT, miDigAfterDec
'        gGridSetColumnType frmSalesTaxDetail.grdSalesTaxDetail, kColTaxCodeTxblFreightAmt, _
'                SS_CELL_TYPE_FLOAT, miDigAfterDec
'        gGridSetColumnType frmSalesTaxDetail.grdSalesTaxDetail, kColTaxCodeTxblSTaxAmt, _
'                SS_CELL_TYPE_FLOAT, miDigAfterDec
'        gGridSetColumnType frmSalesTaxDetail.grdSalesTaxDetail, kColTaxCodeExmptAmt, _
'                SS_CELL_TYPE_FLOAT, miDigAfterDec
'        gGridSetColumnType frmSalesTaxDetail.grdSalesTaxDetail, kColTaxCodeActSTaxAmt, _
'                SS_CELL_TYPE_FLOAT, miDigAfterDec
'        gGridSetColumnType frmSalesTaxDetail.grdSalesTaxDetail, kColTaxCodeActUseTaxAmt, _
'                SS_CELL_TYPE_FLOAT, miDigAfterDec
'        gGridSetColumnType frmSalesTaxDetail.grdSalesTaxDetail, kColTaxCodeActNonRecAmt, _
'                SS_CELL_TYPE_FLOAT, miDigAfterDec
'
'    End If
'
'    gGridSetColumnType grdDetail, kColPOExtAmt, SS_CELL_TYPE_FLOAT, miDigAfterDec                               'Grid sales amount
'
'    gGridSetColumnType grdDetail, kColPOAmtInvcd, SS_CELL_TYPE_FLOAT, miDigAfterDec                               'Grid sales amount
    
  'Set return Code
    bSetFormCurrencyControls = True
    sLastCurrID = sCurrID

    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bSetFormCurrencyControls", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function
Public Function DMDataDisplayed(oDm As Object)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    'Agregado por Multiconsulting
    gGridLockColumn grdReqLines, kcolReqRequestDate 'Agregado por Multiconsulting
    If (evSegUserBuyInfo = 1 And loadNewReq = True And firstCompEstatusReq = True) Then
        gGridLockGrid grdReqLines
        gGridLockGrid grdReqLineDtl
        lkuWarehouse.Enabled = True
        txtComment.Enabled = False
        calDate.Enabled = False
        cboExpReason.Enabled = False
        chkExpedite.Enabled = False
        cmdGenerate.Enabled = False
    End If
    'Agregado por Multiconsulting
    
    'lkuWarehouse.Text
    
    
    'lkuWarehouse.Text = "Nave 4"

    
    CreateLineJoin oDm
    DisplayStatus
    
' If Dept is filled in, disable warehouse
    If Len(lkuDept) > 0 Then
        lkuWarehouse.Enabled = False
' If warehouse is filled in, disable Dept
    ElseIf Len(lkuWarehouse) > 0 Then
        lkuDept.Enabled = False
    End If
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "DMDataDisplayed", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Sub SetTags()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    lkuMain.Tag = lkuMain.Text
    lkuDept.Tag = lkuDept.Text
    lkuWarehouse.Tag = lkuWarehouse.Text
    
    calDate.Tag = calDate.Text
    
    chkExpedite.Tag = chkExpedite.Value
    
    txtComment.Tag = txtComment.Text
    txtContact.Tag = txtContact.Text
    txtOriginator.Tag = txtOriginator.Text
    txtStatus.Tag = txtStatus.Text
    
    cboExpReason.Tag = cboExpReason.ListIndex
    
    'Agregado por Multiconsulting Osmel Barreras
    sddEstatusReq.Tag = sddEstatusReq.ListIndex
    txtAutorizaReq.Tag = txtAutorizaReq.Text
    txt2doAutorizaReq.Tag = txt2doAutorizaReq.Text
    txtAceptaReq.Tag = txtAceptaReq.Text
    calAceptaReq.Tag = calAceptaReq.Text
    calAutorizaReq.Tag = calAutorizaReq.Text
    cal2doAutorizaReq.Tag = cal2doAutorizaReq.Text
    txtEstatusReqDesc.Tag = txtEstatusReqDesc.Text
    'Agregado por Multiconsulting Osmel Barreras
    
'    SetDetailTags
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "SetTags", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Public Function bIsValidWhse() As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim rs As Object
    Dim sSql As String
    
    If Len(Trim(lkuWarehouse)) > 0 Then
        sSql = "SELECT WhseKey FROM timWarehouse WHERE WhseID = "
        sSql = sSql & gsQuoted(lkuWarehouse)
        sSql = sSql & " AND CompanyID = " & gsQuoted(msCompanyID) & " AND Transit = 0"
        Set rs = oclass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
        If rs.IsEOF Then
            giSotaMsgBox Me, moClass.moSysSession, kmsgARBadField, _
gsStripChar(lblWarehouse.Caption, "&")
            lkuWarehouse.Text = lkuWarehouse.Tag
            lkuWarehouse.SetFocus
            bIsValidWhse = False
        Else
            lkuWarehouse.Tag = lkuWarehouse.Text
            moDmForm.SetColumnValue "DfltShipToWhseKey", rs.Field("WhseKey")
            lkuDept.Enabled = False
            bIsValidWhse = True
        End If
    Else
        lkuWarehouse.Tag = lkuWarehouse.Text
        moDmForm.SetColumnValue "DfltShipToWhseKey", ""
        lkuDept.Enabled = True
        bIsValidWhse = True
    End If
    Set rs = Nothing
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidWhse", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Public Function bIsValidPurchDept() As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim rs As Object
    Dim sSql As String
    
    If Len(Trim(lkuDept)) > 0 Then
        sSql = "SELECT PurchDeptKey FROM tpoPurchDepartment WHERE PurchDeptID = "
        sSql = sSql & gsQuoted(lkuDept)
        sSql = sSql & " AND CompanyID = " & gsQuoted(msCompanyID)
        Set rs = oclass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
        If rs.IsEOF Then
            giSotaMsgBox Me, moClass.moSysSession, kmsgARBadField, _
gsStripChar(lblDept.Caption, "&")
            lkuDept.Text = lkuDept.Tag
            lkuDept.SetFocus
            bIsValidPurchDept = False
        Else
            lkuDept.Tag = lkuDept.Text
            moDmForm.SetColumnValue "DfltPurchDeptKey", rs.Field("PurchDeptKey")
            lkuWarehouse.Enabled = False
            bIsValidPurchDept = True
        End If
    Else
        lkuDept.Tag = lkuDept.Text
        moDmForm.SetColumnValue "DfltPurchDeptKey", ""
        If mbIntegrateWithIM Then
            lkuWarehouse.Enabled = True
        End If
        bIsValidPurchDept = True
    End If
    Set rs = Nothing
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidPurchDept", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function
Public Function bIsValidReqNum() As Boolean
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
'************************************************************************
' Desc:  Pad Req NUmber with zero, Validate it can be used if existing
'        or create in log if it new, perform key change logic
' Returns: True - good req Num/Key Change   False - Not Good Req Num/Key Change
'************************************************************************
    
    
    Dim iRetVal         As Integer  'Stored Procedure Return Code
    Dim lReqKey          As Long     'Purchase Order Surrogate Key
    Dim iStatus         As Integer  'Status of Purchase Order Chosen
    Dim iKeyChangeCode  As Integer  'Key Change Return Code
    Dim sReqNum          As String   'Purchase Order Number TranNo
    Dim bOldMb          As Boolean
    Dim bOldActivity    As Boolean
    Dim lContractKey As Long 'Agregado por Multiconsulting
    
    mbReqClosed = False
    bOldMb = mbDontChkclick
'    bOldActivity = mbActivity
    
    bIsValidReqNum = False                           'Set initial return Code
    
    If Len(Trim(lkuMain.Text)) = 0 Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If
    
    If lkuMain.Text = lkuMain.Tag Then
        bIsValidReqNum = True
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If
    
    lkuMain.Text = sPadWithZero(10, lkuMain)      'Pad Req Number with Leading zeros
    
    If miSecurityLevel = kSecLevelDisplayOnly Then
        
        lReqKey = glGetValidLong(moClass.moAppDB.Lookup("ReqKey", "tpoRequisition WITH (NOLOCK)", _
"TranNo = " & gsQuoted(lkuMain.Text)))
    
    Else
      'Run stored procedure to validate/create (in the log) the PO Number input
        On Error GoTo ExpectedErrorRoutine
        With moClass.moAppDB
            .SetInParam lkuMain.Text                   'Req Number
            .SetInParam kTranTypePORQ                   'Req Tran Type
            .SetInParam gsFormatDateToDB(msBusinessDate) 'Business Date (For Log)
            .SetInParam msCompanyID                     'Company ID
            .SetOutParam lReqKey                         'Output New/Existing PO Key
            .SetOutParam iRetVal                        'Return Code 0-Bad 1-New 2-Existing
            .ExecuteSP "sppoValidateReq"
            lReqKey = .GetOutParam(5)                    'Get New/Existing PO Key
            iRetVal = .GetOutParam(6)                   'Get Return Code
            .ReleaseParams
        End With
        On Error GoTo ExpectedErrorRoutine2
    
       'Check Return Values
        If iRetVal = 0 Then                 '0- Unexpected error occured
            giSotaMsgBox Me, moClass.moSysSession, kmsgSPBadReturn, "sppoValidatePO", "0"
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Function
        
        ElseIf iRetVal = 2 Then             '2-Existing PO Log was ot there
      
          'Incomplete state may have been created by another user
            If iStatus = 1 Then
                
                If giSotaMsgBox(Nothing, moClass.moSysSession, kmsgPOIncomplete) = kretNo Then
                   'User has indicated do not load
                   lkuMain.Text = lkuMain.Tag
                   bSetFocus lkuMain
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
                    Exit Function
                End If

          'Purged/Deleted status do not load
            ElseIf iStatus = 4 Or iStatus = 5 Then
                giSotaMsgBox Me, moClass.moSysSession, kmsgPOPurged
                lkuMain.Text = lkuMain.Tag
                bSetFocus lkuMain
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
                Exit Function
            End If
    
        End If
        
    End If   'Return code has been checked and ok'd

' Perform Key change code
    
    moDmForm.SetColumnValue "ReqKey", lReqKey     'Set Primary Key Field
    mbDontChkclick = True
'    mbActivity = False
    iKeyChangeCode = moDmForm.KeyChange()       'Perform Key change code
    If mbReqClosed And miSecurityLevel <> kSecLevelDisplayOnly Then
        RemoveLastRowFromGrid
    End If
    mbDontChkclick = bOldMb
  
  'Reset the activity flag under certain(error) circumstances
    Select Case iKeyChangeCode
        Case kDmKeyFound, kDmKeyNotFound
           '-- Do nothing - continue on
           'Agregado por Multiconsulting
           If mbIntegratedCT Then
                cmdContract.Visible = True
                lContractKey = lGetReqContract
                If sddType.ItemData = 1 Then
                    If lContractKey > 0 Then
                        gGridLockColumn grdReqLines, kcolReqVendID
                        gGridLockColumn grdReqLines, kColReqUnitMeasID
                        grdReqLines.Enabled = True
                    Else
                        grdReqLines.Enabled = False
                        gGridUnlockColumn grdReqLines, kColReqUnitMeasID
                    End If
                    gGridLockColumn grdReqLines, kcolReqRequestDate
                    cmdContract.Enabled = True
                Else
                    grdReqLines.Enabled = True
                    cmdContract.Enabled = False
                    gGridUnlockColumn grdReqLines, kcolReqVendID
                    gGridUnlockColumn grdReqLines, kColReqUnitMeasID
                    gGridUnlockColumn grdReqLines, kcolReqRequestDate
                End If
                BindNavigators
           Else
                gGridUnlockColumn grdReqLines, kColReqUnitMeasID
                gGridUnlockColumn grdReqLines, kcolReqVendID
                gGridUnlockColumn grdReqLines, kcolReqRequestDate
                cmdContract.Visible = False
           End If
           If lContractKey > 0 Then
                grdReqLines.Enabled = True
                sddType.Enabled = False
            Else
                grdReqLines.Enabled = False
                sddType.Enabled = True
            End If
           'Agregado por Multiconsulting
        Case kDmNotAllowed
            If miSecurityLevel = kSecLevelDisplayOnly Then
                '-- Do nothing again - continue on
            End If
         Case Else
'            mbActivity = bOldActivity
    End Select
    
    Select Case iKeyChangeCode
        
        Case kDmKeyNotFound                         'No Key
           
           'Should be an existing record - (Open Status in log)
            If iRetVal = 2 And iStatus <> 1 Then
                If giSotaMsgBox(Nothing, moClass.moSysSession, kmsgPOLogBad) = kretNo Then
                  'reset form
                    sReqNum = lkuMain.Text
                    lkuMain.SetFocus
                    moDmForm.Action kDmCancel
                    lkuMain.Text = sReqNum
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
                    Exit Function
                End If
            End If
            
         'Load up defaults & Set up form
            moClass.lUIActive = kChildObjectActive
            lkuMain.Protected = True
'           This should be taken care of by setting the protected property of the lookup control
            lkuMain.TabStop = False

            mbDontChkclick = True
            LoadDflts
            mbDontChkclick = bOldMb
            CreateLineJoin moDmForm         'Create table for pulling up existing rows
            SetTags
            
            ' Line Entry hook
            bIsValidReqNum = True
    
        Case kDmKeyFound
            moClass.lUIActive = kChildObjectActive
            SetupExistingReq
            bIsValidReqNum = True
        
        Case kDmKeyNotComplete
            '-- key not completely filled in
            bIsValidReqNum = False
            
        Case kDmError
            '-- database error occurred trying to get row
            bIsValidReqNum = False
         
        Case Else
            If miSecurityLevel = kSecLevelDisplayOnly And iKeyChangeCode = kDmNotAllowed Then
                'No messages here
            Else
                giSotaMsgBox Me, moClass.moSysSession, kmsgUnexpectedKeyChangeCode, iKeyChangeCode
            End If
            
            bIsValidReqNum = False
    
    End Select
    
    mlCurrRow = 0
    
    ChangeToolBar (True)

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Function

ExpectedErrorRoutine:
giSotaMsgBox Me, moClass.moSysSession, kmsgUnexpectedSPReturnValue, "sppoValidatePO", Err.Description
gClearSotaErr
mbDontChkclick = bOldMb
'mbActivity = bOldActivity

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Function
ExpectedErrorRoutine2:
'MyErrMsg moClass, Err.Description, Err, sMyName, "bIsValidPoNum"
gClearSotaErr
mbDontChkclick = bOldMb
'mbActivity = bOldActivity

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
Exit Function

'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidReqNum", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Sub ChangeToolBar(bNewVal As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
   If miSecurityLevel <> kSecLevelDisplayOnly Then
        tbrMain.ButtonEnabled(kTbFinish) = bNewVal
        tbrMain.ButtonEnabled(kTbSave) = bNewVal
        tbrMain.ButtonEnabled(kTbCancel) = bNewVal
        tbrMain.ButtonEnabled(kTbDelete) = bNewVal
   End If
        tbrMain.ButtonEnabled(kTbPrint) = bNewVal

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "ChangeToolBar", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub


Private Sub SetupExistingReq()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    
    Dim bOldMb  As Boolean
    Dim lPOKey      As Long
    Dim lActivity   As Long
    
    bOldMb = mbDontChkclick
    
  'Protect The PO Number
    lkuMain.Protected = True
'   This should be taken care of by setting the protected property of the lookup control
    lkuMain.TabStop = False
    bSetFormCurrencyControls msHomeCurrID
    
  'Set The form tags
    SetTags

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "SetupExistingReq", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub


Private Function sPadWithZero(iLen As Integer, ctlMask As Control) As String
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++
'************************************************************************
' Desc:  Load a masked control with a zero padded Number if
'        a Number is passed in
' Parms: iLen - Maximum Number of characters to use if No mask Set
'        ctlMask - Masked edit control
' Returns: Value sent in padded with zeros if a Number is passed in
'************************************************************************
    Dim sSetText        As String
    Dim bIsANumber      As Boolean
    Dim TheNumber       As Long
    
    bIsANumber = True
    sSetText = Trim$(ctlMask.Text)
    
    If Len(sSetText) > 0 Then
        On Error GoTo ExpectedErrorRoutine
        TheNumber = CLng(sSetText)
        
        If TheNumber <= 0 Then 'invalid - must be a Number > 0
            bIsANumber = False
        End If
    
    End If
    
    If bIsANumber Then
        sPadWithZero = String(iLen - Len(sSetText), "0") & sSetText
    Else
        sPadWithZero = ctlMask.Text
    End If
    
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Function
    
ExpectedErrorRoutine:
    bIsANumber = False
    gClearSotaErr
    Resume Next

'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "sPadWithZero", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Sub cmdGenerate_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If Not moFormCust.onClick(cmdGenerate, True) Then Exit Sub
    End If
    #End If
'+++ End Customizer Code Push +++
'************************************************************************
' Desc: Call the push/pull form and generate PO's for the selected lines
'       by calling the PO API.  After the return, update each of the lines
'       with the PO key and then check whether the Req status needs to be
'       updated.
' Parms: None
'
'************************************************************************
Dim sID         As String
Dim sUser       As String
Dim vPrompt     As Variant
'Agregado por Multiconsulting Osmel Barreras
Dim valRet      As Integer

    If msPersistCantRows <> grdReqLines.DataRowCnt And evSegUserBuyInfo = 1 And sddEstatusReq.ItemData = kReqStatusAccepted Then
        valRet = MsgBox("El usuario autenticado no tiene privilegios de seguridad para eliminar partidas de esta Requisiciòn", vbCritical, "Eliminar partidas de Requisiciòn")
        Exit Sub
    End If
    
    If Not ComprobarProveedor And grdReqLines.DataRowCnt = 1 Then
        Exit Sub
    End If
'Agregado por Multiconsulting Osmel Barreras
  'Check the user id typed in can override (no cancel hit)
    sID = CStr("REQGENRTPO")
    sUser = CStr(moClass.moSysSession.UserId)
    vPrompt = True
    If moClass.moFramework.GetSecurityEventPerm(sID, sUser, vPrompt) = 0 Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        valRet = MsgBox("El usuario autenticado no posee privilegios de seguridad para la generación de Òrdenes de Compra", vbOKOnly, "Generar Òrdenes de Compra") 'Agregado por Multiconsulting
        Exit Sub
    End If
    
    'Agregado por Multiconsulting Osmel Barreras
    If sddEstatusReq.ItemData <> kReqStatusAccepted Then
        valRet = MsgBox("La Requisición no posee el estatus de aceptada la generación de una Orden de Compra", vbOKOnly, "Cambiar Estatus de Requisiciòn")
        Exit Sub
    End If
    'Agregado por Multiconsulting Osmel Barreras
    
    If moDmForm.IsDirty(True) Then
        'If giSotaMsgBox(frmRequistn, moClass.moSysSession, kmsgReqSaveChanges, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)) = kretCancel Then
        If giSotaMsgBox(frmRequistn, moClass.moSysSession, kmsgReqSaveChanges, frmRequistn.Caption) = kretCancel Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Sub
        Else
            If moDmForm.Save(True) <> kDmSuccess Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
                Exit Sub
            End If
        End If
    End If

    mbSelectFormLoaded = True
    
    frmSelectReqLines.Init moClass, lkuMain.Text, moDmForm.GetColumnValue("ReqKey")
    'Agregado por Multiconsulting
    If mbIntegratedCT And lGetReqContract > 0 Then
        frmSelectReqLines.chkCostFromReq.Value = 1
        frmSelectReqLines.chkCostFromReq.Enabled = False
    Else
        frmSelectReqLines.chkCostFromReq.Value = 0
        frmSelectReqLines.chkCostFromReq.Enabled = True
    End If
    'Agregado por Multiconsulting
    frmSelectReqLines.Show vbModal
    
    mbSelectFormLoaded = False

' If the form was closed by closing the launcher, unload the form
    If Not moClass.mbClassShutdown Then
        ' Now force the reloading of the record so that the po's generated will be
        ' displayed.
        lkuMain.Tag = ""
        bIsValidReqNum
        If mbPrintErrorReport Then
            lStartErrorReport
            mbPrintErrorReport = False
        End If
    Else
        Unload frmRequistn
    End If
    
'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "cmdGenerate_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub









    











Private Sub chkExpedite_Click()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.onClick chkExpedite, True
    #End If
'+++ End Customizer Code Push +++

'   If the req num is not entered, don't allow
    If Trim(lkuMain) = "" Then
        chkExpedite.Value = 0
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Sub
    End If
    If chkExpedite.Value = 0 Then
        cboExpReason.Enabled = False
        cboExpReason.ListIndex = -1
        cboExpReason.BackColor = vbButtonFace
        
    Else
        cboExpReason.Enabled = True
        cboExpReason.BackColor = vbWindowBackground
        
    End If

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "chkExpedite_Click", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub



Private Sub sbrMain_ButtonClick(sButton As String)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    HandleToolBarClick sButton

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "sbrMain_ButtonClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_CONTROL_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
'Agregado por multiconsulting
Private Sub sddEstatusReq_Click(Cancel As Boolean, ByVal PrevIndex As Long, ByVal NewIndex As Long)
''+++ VB/Rig Begin Push +++
'#If ERRORTRAPON Then
'On Error GoTo VBRigErrorRoutine
'#End If
''+++ VB/Rig End +++
'
'    Dim statusSelected As Integer
'    Dim sIDEtReq As String
'    Dim sUser As String
'    Dim vPrompt As Boolean
'    Dim valRet  As Integer
'
'    If NewIndex = PrevIndex Then
'        Exit Sub
'    End If
'
'    'Agregado por multiconsulting
'    If PrevIndex = -1 And moDmForm.State = kDmStateEdit Then
'        If Len(Trim$(moDmReqAdicInfo.GetColumnValue("StatusKey"))) = kReqStatusPending Or glGetValidLong(moDmReqAdicInfo.GetColumnValue("StatusKey")) <> glGetValidLong(sddEstatusReq.ItemData(NewIndex)) Then
'            Cancel = True
'            Exit Sub
'        End If
'
'    End If
'    'Agregado por multiconsulting
'
'    'Revisar
'    If PrevIndex = -1 And moDmForm.State = kDmStateEdit Then
'        If Len(Trim$(moDmReqAdicInfo.GetColumnValue("StatusKey"))) = kReqStatusPending Or glGetValidLong(moDmReqAdicInfo.GetColumnValue("StatusKey")) <> glGetValidLong(sddEstatusReq.ItemData(NewIndex)) Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
'
'    If msPersistCantRows <> grdReqLines.DataRowCnt And grdReqLines.DataRowCnt > 0 And evSegUserBuyInfo = 1 And sddEstatusReq.ItemData = kReqStatusAccepted Then
'        sddEstatusReq.ListIndex = PrevIndex
'        valRet = MsgBox("El usuario autenticado no tiene privilegios de seguridad para eliminar partidas de esta Requisiciòn", vbCritical, "Eliminar partidas de Requisiciòn")
'        Exit Sub
'    End If
'
'    '************************************************************************
'    ' Desc: Se ejecuta el procedimiento para la definición del Estatus de la
'    '       Requisición para la generación de Ordenes de Compras.
'    ' Parms: None
'    ' Develop: Osmel Barreras (MultiConsulting Cuba)
'    '
'    '************************************************************************
'
'        If reqIsClosed = True Then
'            EnabledFormFields (False)
'            sddEstatusReq.ListIndex = PrevIndex
'            sddEstatusReq.Enabled = False
'            tbrMain.ButtonEnabled(kTbSave) = False
'            tbrMain.ButtonEnabled(kTbFinish) = False
'            GoTo CheckGenTrace
'        Else
'            If sddEstatusReq.ItemData = kReqStatusCancel Then
'                EnabledFormFields (False)
'            Else
'                EnabledFormFields (True)
'            End If
'        End If
'
'        If PrevIndex = -1 And sddEstatusReq.ItemData(NewIndex) = kReqStatusPending And firstCompEstatusReq = True Then
'            UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'            Exit Sub
'        ElseIf loadNewReq And firstCompEstatusReq Then
'            UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'            firstCompEstatusReq = False
'            Exit Sub
'        Else
'            loadNewReq = False
'        End If
'
'    '   Check the user id typed in can change the Estatus Requisition.
'            sUser = CStr(moClass.moSysSession.UserId)
'            vPrompt = False
'
'    Select Case sddEstatusReq.ItemData(NewIndex)
'    Case kReqStatusAuthorized
'            If firstCompEstatusReq = True And evSegUserAutz = 0 Then
'                txtAceptaReq.BackColor = &H8000000F
'                calAceptaReq.BackColor = &H8000000F
'                txtAutorizaReq.BackColor = &H8000000F
'                calAutorizaReq.BackColor = &H8000000F
'                txt2doAutorizaReq.BackColor = &H8000000F
'                cal2doAutorizaReq.BackColor = &H8000000F
'                UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'                firstCompEstatusReq = False
'                Exit Sub
'            End If
'
'            If evSegUserAutz = 1 And sddEstatusReq.ItemData(PrevIndex) = kReqStatusAccepted And NewIndex <> -1 Then
'                statusSelected = MsgBox("No usuario autenticado no tiene privilegios funcionales para cambiar el Estatus de la Requisición", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                loadNewReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                Exit Sub
'            End If
'
'            If evSegUserAutz = 1 Or PrevIndex = -1 Then       'Comprobando si el usuario tiene permiso para Autorizar la Requisiciòn
'                chkb2doAuthNeed.Enabled = True
'                chkb2doAutorizoReq.Enabled = True
'                txtAutorizaReq.BackColor = &H80000005
'                calAutorizaReq.BackColor = &H80000005
'
'                If firstCompEstatusReq = False Or Len(Trim$(txtAutorizaReq.Text)) = 0 Then
'                    txtAutorizaReq.Text = msCurrentUser
'                    calAutorizaReq.Text = Format(msBusinessDate, gsGetLocalVBDateMask())
'                End If
'
'                UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'                firstCompEstatusReq = False
'            Else
'                statusSelected = MsgBox("El usuario autenticado no tiene privilegios de seguridad para cambiar el Estatus la Requisiciòn a Autorizada", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                loadNewReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                txtAutorizaReq.BackColor = &H8000000F
'                calAutorizaReq.BackColor = &H8000000F
'                txt2doAutorizaReq.BackColor = &H8000000F
'                cal2doAutorizaReq.BackColor = &H8000000F
'            End If
'
'    Case kReqStatusAccepted
'            If firstCompEstatusReq = True Then
'                If evSegUserBuyInfo = 0 Then
'                    txtAceptaReq.BackColor = &H8000000F
'                    calAceptaReq.BackColor = &H8000000F
'                    txtAutorizaReq.BackColor = &H8000000F
'                    calAutorizaReq.BackColor = &H8000000F
'                    txt2doAutorizaReq.BackColor = &H8000000F
'                    cal2doAutorizaReq.BackColor = &H8000000F
'                    UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'                    firstCompEstatusReq = False
'                    Exit Sub
'                Else
'                    PrevIndex = sddEstatusReq.GetIndexByItemData(kReqStatusAuthorized)
'                End If
'            End If
'
'            If evSegUserAcpt = 1 Then       'Comprobando si el usuario tiene permiso para Aceptar la Requisiciòn
'                If sddEstatusReq.ItemData(PrevIndex) <> kReqStatusAuthorized And PrevIndex <> -1 Then
'                    statusSelected = MsgBox("La requisiciòn debe estar Autorizada para poder cambiar su Estatus por Aceptada", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                    firstCompEstatusReq = True
'                    loadNewReq = True
'                    sddEstatusReq.ListIndex = PrevIndex
'                    Exit Sub
'                End If
'
'                If sddEstatusReq.ItemData(PrevIndex) = kReqStatusAuthorized And chkb2doAuthNeed.Value = 1 And chkb2doAutorizoReq.Value <> 1 Then
'                    statusSelected = MsgBox("La requisiciòn necesita de doble Autorizo para poder cambiar su Estatus por Aceptada", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                    firstCompEstatusReq = True
'                    loadNewReq = True
'                    sddEstatusReq.ListIndex = PrevIndex
'                    Exit Sub
'                End If
'
'                txtAceptaReq.BackColor = &H80000005
'                calAceptaReq.BackColor = &H80000005
'
'                If firstCompEstatusReq = False Then
'                    txtAceptaReq.Text = msCurrentUser
'                    calAceptaReq.Text = Format(msBusinessDate, gsGetLocalVBDateMask())
'                End If
'
'                UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'                firstCompEstatusReq = False
'            Else
'                statusSelected = MsgBox("El usuario autenticado no tiene privilegios de seguridad para cambiar el Estatus la Requisiciòn a Aceptada", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                loadNewReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                txtAceptaReq.BackColor = &H8000000F
'                calAceptaReq.BackColor = &H8000000F
'            End If
'
'    Case 3:
'            If firstCompEstatusReq = True And evSegUserRchz = 0 Then
'                txtAceptaReq.BackColor = &H8000000F
'                calAceptaReq.BackColor = &H8000000F
'                txtAutorizaReq.BackColor = &H8000000F
'                calAutorizaReq.BackColor = &H8000000F
'                txt2doAutorizaReq.BackColor = &H8000000F
'                cal2doAutorizaReq.BackColor = &H8000000F
'                UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'                firstCompEstatusReq = False
'                Exit Sub
'            End If
'
'            If evSegUserRchz = 0 Then        'Comprobando si el usuario tiene permiso para Cancelar la Requisiciòn
'                statusSelected = MsgBox("El usuario autenticado no tiene privilegios de seguridad para cambiar el Estatus la Requisiciòn a Cancelada", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                loadNewReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                Exit Sub
'            ElseIf txtOriginator.Text <> msCurrentUser And evSegUserAutz = 0 Then
'                statusSelected = MsgBox("El usuario autenticado no es el creador de la Requisición por tanto no tiene privilegios funcionales para cambiar el Estatus a Cancelada", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                Exit Sub
'            End If
'
'            txtAceptaReq.BackColor = &H8000000F
'            calAceptaReq.BackColor = &H8000000F
'            txtAutorizaReq.BackColor = &H8000000F
'            calAutorizaReq.BackColor = &H8000000F
'            EnabledFormFields (False)
'            chkb2doAutorizoReq.Enabled = False
'            chkbEstatusDesc.Value = 0
'            chkbEstatusDesc.Enabled = False
'            txt2doAutorizaReq.BackColor = &H8000000F
'            cal2doAutorizaReq.BackColor = &H8000000F
'            firstCompEstatusReq = False
'
'    Case Else:
'            If firstCompEstatusReq = True Then
'                txtAceptaReq.BackColor = &H8000000F
'                calAceptaReq.BackColor = &H8000000F
'                txtAutorizaReq.BackColor = &H8000000F
'                calAutorizaReq.BackColor = &H8000000F
'                txt2doAutorizaReq.BackColor = &H8000000F
'                cal2doAutorizaReq.BackColor = &H8000000F
'                UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'                firstCompEstatusReq = False
'                Exit Sub
'            End If
'
'            If (evSegUserAcpt = 1 Or evSegUserBuyInfo = 1) And sddEstatusReq.ItemData(PrevIndex) = kReqStatusAccepted And (sddEstatusReq.ItemData(NewIndex) = kReqStatusAuthorized Or sddEstatusReq.ItemData(NewIndex) = kReqStatusCancel) Then
'                statusSelected = MsgBox("No usuario autenticado no tiene privilegios funcionales para cambiar el Estatus de la Requisición", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                loadNewReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                Exit Sub
'            ElseIf (evSegUserAcpt = 1 Or evSegUserBuyInfo = 1) And sddEstatusReq.ItemData(PrevIndex) = kReqStatusAuthorized And (sddEstatusReq.ItemData(NewIndex) <> kReqStatusAccepted And NewIndex <> -1) Then
'                statusSelected = MsgBox("No usuario autenticado no tiene privilegios funcionales para cambiar el Estatus de la Requisición", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                loadNewReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                Exit Sub
'            ElseIf (evSegUserAcpt = 0 And evSegUserBuyInfo = 1) And sddEstatusReq.ItemData(PrevIndex) = kReqStatusPending And (NewIndex <> -1) Then
'                statusSelected = MsgBox("No usuario autenticado no tiene privilegios funcionales para cambiar o modificar la Requisición", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                loadNewReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                Exit Sub
'            End If
'
'            If evSegUserAutz = 1 And sddEstatusReq.ItemData(PrevIndex) = kReqStatusAccepted And NewIndex <> -1 Then
'                statusSelected = MsgBox("No usuario autenticado no tiene privilegios funcionales para cambiar el Estatus de la Requisición", vbOKOnly, "Cambiar Estatus de Requisiciòn")
'                firstCompEstatusReq = True
'                loadNewReq = True
'                sddEstatusReq.ListIndex = PrevIndex
'                Exit Sub
'            End If
'
'            EnabledFormFields (True)
'            UpdateEstatusFields (sddEstatusReq.ItemData(NewIndex))
'
'            txtAceptaReq.BackColor = &H8000000F
'            calAceptaReq.BackColor = &H8000000F
'            txtAutorizaReq.BackColor = &H8000000F
'            calAutorizaReq.BackColor = &H8000000F
'            txt2doAutorizaReq.BackColor = &H8000000F
'            cal2doAutorizaReq.BackColor = &H8000000F
'    End Select
'
'CheckGenTrace:
'        If Not firstCompEstatusReq And Not loadNewReq Then
'                RequisitionChanged = True
'                Dim elRow
'                Dim i As Integer
'                For i = 1 To grdReqLines.DataRowCnt
'                    If i <= listRows.Count Then
'                        listRows.Remove ("r" & i)
'                    End If
'
'                    If i <= grdReqLines.DataRowCnt Then
'                        listRows.Add i, "r" & i
'                    End If
'                    lastRowMod = i
'                Next
'        End If
'
''+++ VB/Rig Begin Pop +++
'        Exit Sub
'
'VBRigErrorRoutine:
'        gSetSotaErr Err, sMyName, "sddEstatusReq_Click", VBRIG_IS_FORM
'        Select Case VBRIG_IS_CONTROL_EVENT
'        Case VBRIG_IS_NON_EVENT
'                Err.Raise guSotaErr.Number
'        Case Else
'                Call giErrorHandler: Exit Sub
'        End Select
''+++ VB/Rig End +++

End Sub
'Agregado por multiconsulting

'Agregado por Multiconsulting
Private Sub sddType_Click(Cancel As Boolean, ByVal PrevIndex As Long, ByVal NewIndex As Long)
    Dim lContractKey As Long
    If NewIndex = -1 Then Exit Sub
        
    If Not bAllowContractChange Then
        Cancel = True
        Exit Sub
    End If
        
    If mbIntegratedCT Then
         lContractKey = lGetReqContract
         If sddType.ItemData = 1 Then
             If lContractKey > 0 Then
                 gGridLockColumn grdReqLines, kcolReqVendID
                 gGridLockColumn grdReqLines, kColReqUnitMeasID
                 grdReqLines.Enabled = True
             Else
                 grdReqLines.Enabled = False
                 gGridUnlockColumn grdReqLines, kColReqUnitMeasID
             End If
             gGridLockColumn grdReqLines, kcolReqRequestDate
             cmdContract.Enabled = True
         Else
             grdReqLines.Enabled = True
             cmdContract.Enabled = False
             gGridUnlockColumn grdReqLines, kcolReqVendID
             gGridUnlockColumn grdReqLines, kColReqUnitMeasID
             gGridUnlockColumn grdReqLines, kcolReqRequestDate
         End If
         BindNavigators
    Else
         gGridUnlockColumn grdReqLines, kColReqUnitMeasID
    End If
End Sub
'Agregado por Multiconsulting

Private Sub tbrMain_ButtonClick(Button As String)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'********************************************************************************
'Agregado por MultiConsulting Osmel Barreras
'   Check the user id typed in can change the Estatus Requisition.
    CheckUserStatusReqPermission
'Agregado por MultiConsulting Osmel Barreras
'********************************************************************************

    HandleToolBarClick Button
    
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

Private Function bIsValidContact() As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

' The contact must be filled in.
    
    bIsValidContact = True

    If Len(Trim(txtContact)) = 0 Then
        bIsValidContact = False
        giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeBlank, gsStripChar(lblContact.Caption, "&")
        ' get rid of any spaces
        txtContact.Text = ""
        txtContact.SetFocus
    End If
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidContact", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function
Private Function bIsValidTranDate() As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

' The Tran date must be filled in.
    
    bIsValidTranDate = True

    If Len(calDate) = 0 Then
        bIsValidTranDate = False
        giSotaMsgBox Me, moClass.moSysSession, kmsgRequiredField, gsStripChar(lblDate.Caption, "&")
        calDate.SetFocus
    Else
        If Not calDate.IsValid Then
            bIsValidTranDate = False
            giSotaMsgBox Me, moClass.moSysSession, kmsgInvalidDate
            calDate = ""
            calDate.SetFocus
        End If
    End If
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidTranDate", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Function bIsValidOriginator() As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

' The originator must be filled in.
    
    bIsValidOriginator = True

    If Len(Trim(txtOriginator)) = 0 Then
        bIsValidOriginator = False
        giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeBlank, gsStripChar(lblOriginator.Caption, "&")
        ' get rid of any spaces
        txtOriginator.Text = ""
        txtOriginator.SetFocus
    End If
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidOriginator", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function
Private Function bIsValidItemDesc(lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lOldRow As Long, lOldCol As Long
' The description must be filled in if the Item id is not present.
' Also check that if a whse is filled in, that the Item ID is also present.
    
    bIsValidItemDesc = True
    
    If Len(Trim(gsGridReadCell(grdReqLines, lRow, kColReqDescription))) = 0 And _
Len(gsGridReadCell(grdReqLines, lRow, kColReqItemID)) = 0 Then
            bIsValidItemDesc = False
' Set the column to the description column
            lOldRow = grdReqLines.ActiveRow
            lOldCol = grdReqLines.ActiveCol
            giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeBlank, msItemDescCaption
            MoveToGridField lOldCol, lOldRow, kColReqDescription, lRow
    ElseIf Len(Trim(gsGridReadCell(grdReqLines, lRow, kColReqWhseID))) > 0 And _
Len(gsGridReadCell(grdReqLines, lRow, kColReqItemID)) = 0 Then
            bIsValidItemDesc = False
' Set the column to the item column
            lOldRow = grdReqLines.ActiveRow
            lOldCol = grdReqLines.ActiveCol
            giSotaMsgBox Me, moClass.moSysSession, kmsgNoBlankItemWithWhse
            MoveToGridField lOldCol, lOldRow, kColReqItemID, lRow
    End If
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidItemDesc", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function
Private Function bIsValidQtyRequested(lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

' The quantity must be greater than zero if the qty requested is editable.
    Dim sQtyReq As String
    Dim fQtyReq As Double
    Dim lOldRow As Long, lOldCol As Long
    Dim lItemKey As Long

    bIsValidQtyRequested = True
    
' if the qty requested is locked, don't check the value.
    With grdReqLines
        .Col = kColReqQtyRequested
        .Row = lRow
        If .Lock Then
            Exit Function
        End If
    End With
    
    sQtyReq = gsGridReadCell(grdReqLines, lRow, kColReqQtyRequested)
    If Len(sQtyReq) = 0 Then
        fQtyReq = 0
    Else
        fQtyReq = CDbl(sQtyReq)
    End If
    
    If fQtyReq <= 0 Then
            bIsValidQtyRequested = False
            giSotaMsgBox Me, moClass.moSysSession, kmsgMustBeGreatZero
' Set the column to the quantity column
            lOldRow = grdReqLines.ActiveRow
            lOldCol = grdReqLines.ActiveCol
            MoveToGridField lOldCol, lOldRow, kColReqQtyRequested, lRow
    End If

    If Int(fQtyReq) <> fQtyReq Then
        lItemKey = glGetValidLong(gsGridReadCellText(grdReqLines, lRow, kColReqItemKey))
        Set moItem = moIMSClass.Items(lItemKey)
        If Not moItem.AllowDecimalQty Then
            bIsValidQtyRequested = False
            MsgBox "Decimal quantities are not allowed for " & moItem.ItemID, vbExclamation, gsBuildString(kSotaTitle, moClass.moAppDB, moClass.moSysSession)
            lOldRow = grdReqLines.ActiveRow
            lOldCol = grdReqLines.ActiveCol
            MoveToGridField lOldCol, lOldRow, kColReqQtyRequested, lRow
        End If
    End If
    
    If fQtyReq > kvMaxQtySize Then
        giSotaMsgBox Me, moClass.moSysSession, kmsgCannotBeGreaterThan, "Qty Requested", kvMaxQtySize
        bIsValidQtyRequested = False
        lOldRow = grdReqLines.ActiveRow
        lOldCol = grdReqLines.ActiveCol
        MoveToGridField lOldCol, lOldRow, kColReqQtyRequested, lRow
    End If
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidQtyRequested", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Function bIsValidWhseDept(lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
Dim lItemKey As Long
Dim lWhseKey As Long
Dim lDeptKey As Long
Dim lOldCol As Long
Dim lOldRow As Long

    bIsValidWhseDept = False
    
    lItemKey = glGetValidLong(gsGridReadCellText(grdReqLines, lRow, kColReqItemKey))
    lWhseKey = glGetValidLong(gsGridReadCellText(grdReqLines, lRow, kColReqWhseKey))
    lDeptKey = glGetValidLong(gsGridReadCellText(grdReqLines, lRow, kColReqPurchDeptKey))
    
    Set moItem = moIMSClass.Items(lItemKey)
    If Not moItem.NonInventory And lWhseKey = 0 And lDeptKey = 0 Then
        giSotaMsgBox Me, moClass.moSysSession, kmsgPOMissingWhseOrDept, _
moItem.ItemID
        lOldRow = grdReqLines.ActiveRow
        lOldCol = grdReqLines.ActiveCol
        MoveToGridField lOldCol, lOldRow, kColReqPurchDeptKey, lRow
        Exit Function
    End If
    
    bIsValidWhseDept = True
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidWhseDept", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Function bIsValidRequestedDate(lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

' The description must be filled in if the Item id is not present.
    Dim sRequestDate As String, lOldRow As Long, lOldCol As Long
    
    bIsValidRequestedDate = True
    sRequestDate = gsGridReadCell(grdReqLines, lRow, kcolReqRequestDate)
    
    If Len(sRequestDate) = 0 Then
           bIsValidRequestedDate = False
' Set the column to the date column
            lOldRow = grdReqLines.ActiveRow
            lOldCol = grdReqLines.ActiveCol
        MsgBox "Request Date required." & vbCrLf & "Please enter a valid date consistent with your date format.", vbInformation, "Sage 500 ERP"
            'giSotaMsgBox Nothing, moClass.moSysSession, kmsgRequiredField, msRequestDateCaption
            MoveToGridField lOldCol, lOldRow, kcolReqRequestDate, lRow
    End If
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidRequestedDate", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function
Private Sub MoveToGridField(lOldCol As Long, lOldRow As Long, lNewCol As Long, lNewRow As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

            
    Dim bCancel As Boolean
    grdReqLines.SetFocus
    gGridSetActiveCell grdReqLines, lNewRow, lNewCol
    grdReqLines_LeaveCell lOldCol, lOldRow, lNewCol, lNewRow, bCancel

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "MoveToGridField", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub
Private Function bIsValidDirtyCheck() As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If Not bIsValidOriginator Then Exit Function
    If Not bIsValidContact Then Exit Function
    
    'Call comment engine to accept user fields for PO level and PO line level before processing.
    msUserFld(0) = gsGetValidStr(moDmForm.GetColumnValue("UserFld1"))
    msUserFld(1) = gsGetValidStr(moDmForm.GetColumnValue("UserFld2"))
    msUserFld(2) = gsGetValidStr(moDmForm.GetColumnValue("UserFld3"))
    msUserFld(3) = gsGetValidStr(moDmForm.GetColumnValue("UserFld4"))
    
    If Not moUF.bValidateUserflds(kEntTypePOPurchOrder, 4, True, msUserFld()) Then
            Exit Function
    End If
    
    moDmForm.SetColumnValue "UserFld1", msUserFld(0)
    moDmForm.SetColumnValue "UserFld2", msUserFld(1)
    moDmForm.SetColumnValue "UserFld3", msUserFld(2)
    moDmForm.SetColumnValue "UserFld4", msUserFld(3)
   

    If mlCurrRow > 0 Then
        If glGetValidLong(gsGridReadCellText(grdReqLines, mlCurrRow, kColReqItemKey)) <> 0 Then
        msUserFld_Line(0) = gsGetValidStr(moDmGrid.GetColumnValue(mlCurrRow, "UserFld1"))
        msUserFld_Line(1) = gsGetValidStr(moDmGrid.GetColumnValue(mlCurrRow, "UserFld2"))
        
        'Call comment engine to accept user fields for PO level and PO line level before processing.
        If Not moUF.bValidateUserflds(kEntTypePOPOLine, 2, True, msUserFld_Line()) Then
            Exit Function
        End If
        
        moDmGrid.SetColumnValue mlCurrRow, "UserFld1", msUserFld_Line(0)
        moDmGrid.SetColumnValue mlCurrRow, "UserFld2", msUserFld_Line(1)
        End If
    End If
    
    bIsValidDirtyCheck = True

'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bIsValidDirtyCheck", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function


'************************************************************************
'   Description:
'       process all toolbar clicks (as well as hotkey shortcuts to
'       toolbar buttons)
'
'   Param:
'       sKey -  token returned from toolbar.  indicates what function
'               is to be executed
'
'   Returns:
'
'************************************************************************

Public Sub HandleToolBarClick(sKey As String)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If moFormCust.ToolbarClick(sKey) Then
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Sub
        End If
    End If
#End If
    Dim iConfirmUnloadVar   As Integer
    Dim iActionCode         As Integer
    Dim iFiltered           As Integer
    Dim lRet                As Long
    Dim sNewKey             As String
    Dim vParseRet           As Variant
    Dim bOldmbClick         As Boolean
    
    bOldmbClick = mbDontChkclick
    mbDontChkclick = True
    
    ' VB 5 does not automatically fire LostFocus event when pressing toolbar
    Me.SetFocus
    
    DoEvents
    
    'make the floating navigators disappear
    navItemGrid.Visible = False
    navSTaxGrid.Visible = False
    navDeptGrid.Visible = False
    navWhseGrid.Visible = False
    navVendorGrid.Visible = False
    
'***********************************************************************************************************
'Agregado por MultiConsulting Osmel Barreras
    navBuyerGrid.Visible = False
    Dim retVal As Integer
'Agregado por MultiConsulting Osmel Barreras
'***********************************************************************************************************

    'Determine whether the grid column needs to be validated prior to saving the form.
    Select Case sKey
        Case kTbFinish, kTbFinishExit
'***********************************************************************************************************
'Agregado por Multiconsulting Osmel Barreras
            If msPersistCantRows <> grdReqLines.DataRowCnt And evSegUserBuyInfo = 1 And sddEstatusReq.ItemData = kReqStatusAccepted Then
                retVal = MsgBox("El usuario autenticado no tiene privilegios de seguridad para eliminar partidas de esta Requisiciòn", vbCritical, "Eliminar partidas de Requisiciòn")
                Exit Sub
            End If
            
            UpdateSegEstatusReq (True)
            
            Select Case sddEstatusReq.ItemData
                Case kReqStatusAuthorized
                    If evSegUserAutz = 0 And sddEstatusReq.ItemData = kReqStatusAuthorized Then
                        retVal = MsgBox("El usuario autenticado no tiene privilegios de seguridad para editar la informaciòn en el Estatus de Requisiciòn Autorizada", vbOKOnly, "Editar Requisiciòn")
                        Exit Sub
                    End If
                
                Case kReqStatusAccepted
                    If (evSegUserBuyInfo = 0 And evSegUserAcpt = 0) And sddEstatusReq.ItemData = kReqStatusAccepted Then
                        retVal = MsgBox("El usuario autenticado no tiene privilegios de seguridad para editar la informaciòn en el Estatus de Requisiciòn Aceptada", vbOKOnly, "Editar Requisiciòn")
                        Exit Sub
                    End If
            End Select
                                                            
'            ComprobarFechaReq
            txtUserMod.Text = msCurrentUser
            
            msPersistCantRows = 0
            msPersistCantReq = ""
            msPersistPresEstReq = ""
            msPersistDescriptionReq = ""
            
'Agregado por Multiconsulting Osmel Barreras
'***********************************************************************************************************
            
' This check is needed to validate the Purch Dept prior to saving
            If bIsValidPurchDept Then
Debug.Print "modmSubGrid.Isdirty " & moDMSubGrid.IsDirty
                iActionCode = moDmForm.Action(kDmFinish)

                If iActionCode = kDmSuccess Then
                    ClearForm
                    If sKey = kTbFinishExit Then
                        Me.Hide
                    End If
                End If
            End If
            
        Case kTbCancel, kTbCancelExit
            iActionCode = moDmForm.Action(kDmCancel)

            If iActionCode = kDmSuccess Then
                ClearForm
                If sKey = kTbCancelExit Then
                    Me.Hide
                End If
            End If
        
        Case kTbDelete
'*******************************************************************************************
''Agregado por Multiconsulting Osmel Barreras
            If reqIsClosed Or evSegUserDelt = 0 Then
                If evSegUserDelt = 0 Then
                    retVal = MsgBox("El usuario no tiene privilegios de Seguridad para Eliminar una Requisición", vbInformation, "Sage MAS 500 Información de Seguridad")
                Else
                    retVal = MsgBox("Las Requisiciones en situación Cerrada no se Eliminan", vbInformation, "Sage MAS 500 Información de Seguridad")
                End If
                                
                Exit Sub
            End If
''Agregado por Multiconsulting Osmel Barreras
'*******************************************************************************************
            iActionCode = moDmForm.DeleteRow

            If iActionCode = kDmSuccess Then
                ClearForm
            End If
            
        Case kTbSave
'*******************************************************************************************
''Agregado por Multiconsulting Osmel Barreras
            If msPersistCantRows <> grdReqLines.DataRowCnt And evSegUserBuyInfo = 1 And sddEstatusReq.ItemData = kReqStatusAccepted Then
                retVal = MsgBox("El usuario autenticado no tiene privilegios de seguridad para eliminar partidas de esta Requisiciòn", vbCritical, "Eliminar partidas de Requisiciòn")
                Exit Sub
            End If
            UpdateSegEstatusReq (True)
            
            Select Case sddEstatusReq.ItemData
                Case kReqStatusAuthorized
                    If evSegUserAutz = 0 And sddEstatusReq.ItemData = kReqStatusAuthorized Then
                        retVal = MsgBox("El usuario autenticado no tiene privilegios de seguridad para editar la informaciòn en el Estatus de Requisiciòn Autorizada", vbOKOnly, "Editar Requisiciòn")
                        Exit Sub
                    End If
                
                Case kReqStatusAccepted
                    If (evSegUserBuyInfo = 0 And evSegUserAcpt = 0) And sddEstatusReq.ItemData = kReqStatusAccepted Then
                        retVal = MsgBox("El usuario autenticado no tiene privilegios de seguridad para editar la informaciòn en el Estatus de Requisiciòn Aceptada", vbOKOnly, "Editar Requisiciòn")
                        Exit Sub
                    End If
            End Select
                        
'            ComprobarFechaReq
            txtUserMod.Text = msCurrentUser
''Agregado por Multiconsulting Osmel Barreras
'*******************************************************************************************

' This check is needed to validate the Purch Dept prior to saving
            If bIsValidPurchDept Then
                iActionCode = moDmForm.Save(True)
Debug.Print "moDmForm.IsDirty " & moDmForm.IsDirty

                If iActionCode = kDmSuccess Then
                    mbDontChkclick = bOldmbClick
                    CreateLineJoin moDmForm
                    'QQ, fixed 14389
                    lkuMain.Tag = ""
                    bIsValidReqNum
                    'QQ, Fixed #29,  In some case after .save, IsDirty flag is not reset.
                    moDmForm.SetDirty False, True
                    'After save, the active grid row may have changed position.  Set old value to current
                    'position so that a false change is not detected.
                    msOldItemID = Trim(gsGridReadCellText(grdReqLines, grdReqLines.ActiveRow, kColReqItemID))
                    SetHourglass False
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
                    Exit Sub
                End If
            End If
                    
        Case kTbHelp
            gDisplayFormLevelHelp Me
            mbDontChkclick = bOldmbClick
            SetHourglass False
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Sub

        Case kTbNextNumber
            If moDmForm.State <> kDmStateNone Then
                If iConfirmUnload() = kDmSuccess Then
                    moDmForm.Clear True
                    ClearForm
                Else
                    mbDontChkclick = bOldmbClick
                    SetHourglass False
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
                    Exit Sub
                End If
            End If

            lkuMain.Protected = False
'           This should be taken care of by setting the protected property of the lookup control
            lkuMain.TabStop = True
            bSetFocus lkuMain
            lkuMain.Text = sGetNextReqNo()

            bIsValidReqNum
            txtOriginator.SetFocus
            iActionCode = kDmSuccess
'***************************************************************************************
' Agregado Multiconsulting Osmel Barreras
            chkbEstatusDesc.Enabled = True
            firstCompEstatusReq = False
            loadNewReq = False
            sddEstatusReq.ListIndex = sddEstatusReq.GetIndexByItemData(kReqStatusPending)
            UpdateEstatusFields (kReqStatusPending)
    
    '   Check the user id typed in can change the StatusReq.
            UpdateSegEstatusReq (False)
            
    '        adicInfo = False
' Agregado Multiconsulting Osmel Barreras
'***************************************************************************************
        Case kTbPrint
            Dim oPrintReq As Object
                If iConfirmUnload(True) = kDmSuccess Then
                    mbPrintingReq = True
'                   Launch the Quick Print task.
                    Set oPrintReq = goGetSOTAChild(moClass.moFramework, moSotaObjects, _
kclsQuickPrintReqs, ktskQuickPrintReqs, kAOFRunFlags, kContextAOF)
                    If Not (oPrintReq Is Nothing) Then
                        oPrintReq.QuickPrintReq lkuMain.Text
                        Set oPrintReq = Nothing
                    End If
                    mbPrintingReq = False
                    Me.SetFocus
                    
                End If
            iActionCode = kDmSuccess

        Case kTbMemo
            CMMemoSelected

    
        Case kTbFirst, kTbPrevious, kTbLast, kTbNext
            
            If mbFocus Then Exit Sub
            
            'Process requested browse move
            If iConfirmUnload(True) = kDmFailure Then Exit Sub
            'Execute requested move
            If sbrMain.Filtered Then
                iFiltered = RSID_FILTERED
            Else
                iFiltered = RSID_UNFILTERED
            End If
            lRet = glLookupBrowse(lkuMain, sKey, iFiltered, sNewKey)
            'Evaluate outcome of requested browse move
            Select Case lRet
                Case MS_SUCCESS

                    vParseRet = gvParseLookupReturn(sNewKey)
                    If IsNull(vParseRet) Then Exit Sub
                    If Trim(lkuMain.Text) <> Trim(vParseRet(1)) Then
                        lkuMain.Text = Trim(vParseRet(1))
                        bIsValidReqNum
                    End If

'***************************************************************************************
'Agregado Multiconsulting Osmel Barreras
                '   Check the user id typed in can change the StatusReq.
                    UpdateSegEstatusReq (False)
                    msPersistCantRows = grdReqLines.DataRowCnt
'Agregado Multiconsulting Osmel Barreras
'***************************************************************************************
                    mbFocus = True
                    If txtOriginator.Enabled Then
                        txtOriginator.SetFocus
                    End If
                    mbFocus = False

                Case Else

                    mbFocus = True
                    gLookupBrowseError lRet, Me, moClass
                    mbFocus = False

            End Select
        
        Case kTbMemo

        Case Else
'            Debug.Print "Calling generic toolbar handler for key: " & sKey
'            moToolbar.GenericHandler sKey, Me, moDmGrid, moClass
'            mbDontChkclick = bOldmbClick
'            SetHourglass False
'            Exit Sub

    End Select

    Select Case iActionCode
        
        Case kDmSuccess
            ' Line Entry hook
'            moLE.InitDataReset

        Case kDmFailure
'            Debug.Print "this blew up (failure)"
            mbDontChkclick = bOldmbClick
            SetHourglass False
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Sub
            
        Case kDmError
             
'            Debug.Print "this blew up"
            mbDontChkclick = bOldmbClick
            SetHourglass False
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
            Exit Sub
    
    End Select
    
    mbDontChkclick = bOldmbClick
    SetHourglass False

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "HandleToolBarClick", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Private Function sGetClassName(hwnd As Long) As String

' This was taken from clsContextMenu for use in the WinHook2 procedure

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
        Dim sClassName As String * 64
        Dim retVal As Integer
        
        retVal = GetClassName(hwnd, sClassName, 64)
        sGetClassName = Left(sClassName, retVal)
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "sGetClassName", VBRIG_IS_CLASS
        Err.Raise guSotaErr.Number
'+++ VB/Rig End +++
End Function



Private Function hGetWindowHandle(lFromhwnd As Long, lTohwnd As Long, tpnt As POINTAPI) As Long

' This was taken from clsContextMenu for use in the WinHook2 procedure


'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
Dim tSavedpnt As POINTAPI
tSavedpnt = tpnt
Dim sClassName As String

Dim lChandle As Long

hGetWindowHandle = 0 ' default return value incase the point is outside the parent

MapWindowPoints lFromhwnd, lTohwnd, tpnt, 1
lChandle = ChildWindowFromPointEx(lTohwnd, tpnt.x, tpnt.y, 1)
If lChandle = 0 Then Exit Function
sClassName = sGetClassName(lChandle)
If sClassName <> "ThunderComboBox" And sClassName <> "SPR32X30_SpreadSheet" And sClassName <> "SPR32X30_SpreadSheet" Then
    If lChandle <> lTohwnd Then
        lChandle = hGetWindowHandle(lFromhwnd, lChandle, tSavedpnt) 'recursive call to get to the bottom
        tpnt = tSavedpnt
        hGetWindowHandle = lChandle
'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
        Exit Function
    End If
Else
    MapWindowPoints lFromhwnd, lChandle, tSavedpnt, 1
    tpnt = tSavedpnt
End If

hGetWindowHandle = lChandle
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "hGetWindowHandle", VBRIG_IS_CLASS
        Err.Raise guSotaErr.Number
'+++ VB/Rig End +++
End Function

#If CUSTOMIZER Then
Private Sub picDrag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'+++ VB/Rig Begin Push +++
'+++ VB/Rig End +++

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
    On Error GoTo VBRigErrorRoutine:
#End If
'+++ VB/Rig End +++

    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseDown Index, Button, Shift, x, y
    End If

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub


'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
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
'+++ VB/Rig End +++

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
    On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseMove Index, Button, Shift, x, y
    End If

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub


'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
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
'+++ VB/Rig End +++

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
    On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseUp Index, Button, Shift, x, y
    End If

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub


'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
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
'+++ VB/Rig End +++

'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
    On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_Paint Index
    End If

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub


'+++ VB/Rig Begin Pop +++
#If ERRORTRAPON = 0 Then
Err.Raise Err
#End If
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
#If CUSTOMIZER Then
Private Sub Form_Activate()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++


    If moFormCust Is Nothing Then
        Set moFormCust = CreateObject("SOTAFormCustRT.clsFormCustRT")
        If Not moFormCust Is Nothing Then
                moFormCust.Initialize Me, goClass
                Set moFormCust.CustToolbarMgr = moToolbar
                moFormCust.ApplyDataBindings moDmForm
                moFormCust.ApplyFormCust
        End If
    End If

'+++ VB/Rig Begin Pop +++
'+++ VB/Rig End +++
    Exit Sub


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
#End If


Private Sub cmdUserFlds_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus cmdUserFlds, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cmdUserFlds_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cmdUserFlds_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus cmdUserFlds, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cmdUserFlds_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cmdGenerate_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus cmdGenerate, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cmdGenerate_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cmdGenerate_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus cmdGenerate, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cmdGenerate_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub


Private Sub txtComment_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange txtComment, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtComment_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress txtComment, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtComment_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtComment_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus txtComment, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtComment_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtComment_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus txtComment, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtComment_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtStatus_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange txtStatus, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtStatus_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtStatus_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress txtStatus, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtStatus_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtStatus_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus txtStatus, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtStatus_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtStatus_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus txtStatus, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtStatus_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtContact_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange txtContact, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtContact_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress txtContact, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtContact_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtContact_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus txtContact, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtContact_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtContact_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus txtContact, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtContact_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtOriginator_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange txtOriginator, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtOriginator_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtOriginator_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress txtOriginator, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtOriginator_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtOriginator_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus txtOriginator, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtOriginator_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtOriginator_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus txtOriginator, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtOriginator_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtNavReturn_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange txtNavReturn, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtNavReturn_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtNavReturn_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress txtNavReturn, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtNavReturn_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtNavReturn_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus txtNavReturn, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtNavReturn_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub txtNavReturn_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus txtNavReturn, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "txtNavReturn_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuDept_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange lkuDept, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuDept_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuDept_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress lkuDept, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuDept_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuDept_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus lkuDept, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuDept_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuDept_BeforeLookupReturn(colSQLReturnVal As Collection, bCancel As Boolean)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then
            bCancel = moFormCust.OnBeforeLookupReturn(lkuDept, True)
            If bCancel Then Exit Sub
        End If
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuDept_BeforeLookupReturn()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuDept_LookupClicked()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLookupClicked lkuDept, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuDept_LookupClicked()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuMain_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange lkuMain, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuMain_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuMain_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress lkuMain, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuMain_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuMain_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus lkuMain, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuMain_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuMain_LookupClicked()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLookupClicked lkuMain, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuMain_LookupClicked()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuWarehouse_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange lkuWarehouse, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuWarehouse_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuWarehouse_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress lkuWarehouse, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuWarehouse_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuWarehouse_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus lkuWarehouse, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuWarehouse_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuWarehouse_BeforeLookupReturn(colSQLReturnVal As Collection, bCancel As Boolean)


'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then
            bCancel = moFormCust.OnBeforeLookupReturn(lkuWarehouse, True)
            If bCancel Then Exit Sub
        End If
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuWarehouse_BeforeLookupReturn()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub lkuWarehouse_LookupClicked()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLookupClicked lkuWarehouse, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "lkuWarehouse_LookupClicked()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub


Private Sub chkExpedite_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus chkExpedite, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "chkExpedite_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub chkExpedite_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus chkExpedite, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "chkExpedite_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cboExpReason_Click(bCancel As Boolean, ByVal lPrevIndex As Long, ByVal lNewIndex As Long)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then
            bCancel = moFormCust.OnDropDownClick(cboExpReason, lPrevIndex, lNewIndex, True)
            If bCancel Then Exit Sub
        End If
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cboExpReason_Click()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cboExpReason_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress cboExpReason, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cboExpReason_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cboExpReason_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus cboExpReason, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cboExpReason_GotFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub cboExpReason_LostFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnLostFocus cboExpReason, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "cboExpReason_LostFocus()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub calDate_Change()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnChange calDate, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "calDate_Change()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub calDate_KeyPress(KeyAscii As Integer)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnKeyPress calDate, KeyAscii, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "calDate_KeyPress()", VBRIG_IS_FORM
    Select Case VBRIG_IS_CONTROL_EVENT
    Case VBRIG_IS_NON_EVENT
        Err.Raise guSotaErr.Number
    Case Else
        Call giErrorHandler: Exit Sub
    End Select
'+++ VB/Rig End +++
End Sub

Private Sub calDate_GotFocus()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++

'+++ Customizer Code Push +++
    #If CUSTOMIZER Then
        If Not moFormCust Is Nothing Then moFormCust.OnGotFocus calDate, True
    #End If
'+++ End Customizer Code Push +++

'+++ VB/Rig Begin Pop +++
    Exit Sub
VBRigErrorRoutine:
    gSetSotaErr Err, sMyName, "calDate_GotFocus()", VBRIG_IS_FORM
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


Public Property Let SetSession(ByVal vNewValue As Variant)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    mlSession = vNewValue
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "SetSession_Let", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Property
        End Select
'+++ VB/Rig End +++
End Property

Public Property Let PrintErrorReport(ByVal vNewValue As Variant)
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    mbPrintErrorReport = vNewValue
'+++ VB/Rig Begin Pop +++
        Exit Property

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "PrintErrorReport_Let", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Property
        End Select
'+++ VB/Rig End +++

End Property










Public Property Get MyApp() As Object
    Set MyApp = App
End Property
Public Property Get MyForms() As Object
    Set MyForms = Forms
End Property

Public Sub CMMemoSelected()
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
    Dim lKeyValue As Long
    
    lKeyValue = glGetValidLong(moDmForm.GetColumnValue("ReqKey"))
    
    Me.Enabled = False
    
    '-- Launch the Memo object
    gLaunchMemo Me, moClass.moFramework, moSotaObjects, kEntTypePORequisition, lKeyValue, _
                "Requisition:", lkuMain.Text, moClass.moSysSession.BusinessDate

    '-- Set the Memo toolbar to the correct state
    gSetMemoToolBarState moClass.moAppDB, lKeyValue, kEntTypePORequisition, msCompanyID, tbrMain
    
    Me.Enabled = True

'+++ VB/Rig Begin Pop +++
        Exit Sub

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "CMMemoSelected", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Sub
        End Select
'+++ VB/Rig End +++
End Sub

Public Sub OfficeInitialization()
'+++ VB/Rig Skip +++
On Error Resume Next
'*************************************************************************
'      Desc:  The Office class has been initialized elsewhere and has
'             already received this form’s reference. Therefore,
'             this method is additive if the client task wishes to add
'             clsDMForm and/or clsDMGrid references.
'     Parms:  N/A
'   Returns:  N/A
'************************************************************************
    With tbrMain.Office
        .AddFormObject moDmForm, "Requisition"
        .AddGridObject moDmGrid, "RequisitionLines"
        .AddGridObject moDMSubGrid, "LinesDist"
    End With
    Err.Clear
End Sub

Private Function bLoadDfltCurrency(lVendKey As Long, lRow As Long) As Boolean
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'*************************************************************************
'      Desc:  The function bLoadDfltCurrency is used to default the Vendor
'             default currency from tapVenAddr based on the input VendorKey
'             for every row of the requisition line.
'     Parms:  Vendkey, Row
'   Returns:  N/A
'************************************************************************
    Dim rs As Object
    Dim sSql As String
    
    If lVendKey > 0 Then
        sSql = "SELECT CurrID FROM tapVendAddr WITH (NOLOCK) WHERE Addrkey = (SELECT PrimaryAddrKey FROM tapVendor WITH (NOLOCK) WHERE Vendkey = "
        sSql = sSql & lVendKey
        sSql = sSql & " AND CompanyID = " & gsQuoted(msCompanyID) & ")"
        Set rs = oclass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    
        If rs.IsEOF Then
            gGridUpdateCell grdReqLines, lRow, kColReqCurrid, ""
        Else
            gGridUpdateCell grdReqLines, lRow, kColReqCurrid, rs.Field("CurrID")
        End If
    Else
        gGridUpdateCell grdReqLines, lRow, kColReqCurrid, ""
    End If
    Set rs = Nothing
    
    
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "bLoadDfltCurrency", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

'Agregado por Multiconsulting
Private Sub CheckUserStatusReqPermission()
    SIDEtReqAutz = CStr("CHGPOETREQAUTZ")
    SIDEtReqAcpt = CStr("CHGPOETREQACPT")
    SIDEtReqRchz = CStr("CHGPOETREQRCHZ")
    SIDEtReqDelt = CStr("DLTPOREQDELETE")
    SIDEtReqBuyInfo = CStr("CHGPOBUYREQINFO")
    
    sUser = CStr(moClass.moSysSession.UserId)
    
    evSegUserAutz = moClass.moFramework.GetSecurityEventPerm(SIDEtReqAutz, sUser, False)
    evSegUserAcpt = moClass.moFramework.GetSecurityEventPerm(SIDEtReqAcpt, sUser, False)
    evSegUserRchz = moClass.moFramework.GetSecurityEventPerm(SIDEtReqRchz, sUser, False)
    evSegUserDelt = moClass.moFramework.GetSecurityEventPerm(SIDEtReqDelt, sUser, False)
    evSegUserBuyInfo = moClass.moFramework.GetSecurityEventPerm(SIDEtReqBuyInfo, sUser, False)
    
    tbrMain.ButtonEnabled(kTbNextNumber) = IIf(evSegUserAcpt = 0 And evSegUserAutz = 0 And evSegUserBuyInfo = 0, True, False)
    
End Sub

Private Sub UpdateSegEstatusReq(pCompSeg As Boolean)
On Error GoTo ExpectedErrorRoutine

'************************************************************************
'   Description:
'      Funcionalidad desarrollada para comprobar si el usuario puede generar una OC
'      a partir del Estatus de la Requisición
'
'   Param:
'       <none>
'
'   Develop:
'       Osmel Barreras Piñera (Cuba)
'************************************************************************
'Exit Sub
        'MsgBox "Seg_perm: " & evSegUserAutz & " - " & evSegUserAcpt & " - " & evSegUserRchz
        If reqIsClosed Then
            EnabledFormFields (False)
            'sddEstatusReq.Locked = True
            Exit Sub
        End If
        
        If evSegUserAutz = 0 And evSegUserAcpt = 0 And evSegUserBuyInfo = 0 And evSegUserRchz = 0 And _
            pCompSeg = True And reqIsClosed = False Then
            'sddEstatusReq.Locked = True
            Exit Sub
        End If
        
        Dim vPrompt As Boolean
        vPrompt = False
        
        Select Case sddEstatusReq.ItemData
        Case kReqStatusAccepted:
            If evSegUserBuyInfo = 0 Then
                cmdGenerate.Enabled = IIf((evSegUserBuyInfo = 1 Or evSegUserAcpt = 1), True, False)
                EnabledFormFields (False)
            Else
                EnabledFormFields (True)
                
                If sddEstatusReq.ItemData = kReqStatusAccepted Then
                    cmdGenerate.Enabled = IIf((evSegUserBuyInfo = 1 Or evSegUserAcpt = 1), True, False)
                    UpdateEstatusFields (kReqStatusAccepted)
                End If
            End If
            
        Case kReqStatusAuthorized:
            If evSegUserAutz = 0 Then
                EnabledFormFields (False)
                chkb2doAutorizoReq.Enabled = False
                    
                If txt2doAutorizaReq.Text = "" Then
                    cal2doAutorizaReq = ""
                End If
                
                If txtAutorizaReq.Text = "" Then
                    calAutorizaReq = ""
                End If
                
                If txtAceptaReq.Text = "" Then
                    calAceptaReq = ""
                End If
            Else
                EnabledFormFields (True)
                UpdateEstatusFields (kReqStatusAuthorized)
                
                If sddEstatusReq.ItemData = kReqStatusAuthorized Then
                    chkb2doAutorizoReq.Enabled = True
                End If
            End If
        End Select
Exit Sub
ExpectedErrorRoutine:
        moClass.moAppDB.Rollback
        MsgBox "Error:UpdateSegEstatusReq " & Err.Description
End Sub

Private Function ComprobarProveedor() As Boolean
    'Dim ComprobarProveedor As Boolean
    Dim iRow As Long
    ComprobarProveedor = True
    
    For iRow = 1 To grdReqLines.DataRowCnt
        If gsGridReadCellText(grdReqLines, iRow, kcolReqVendID) = "" Then
            ComprobarProveedor = False
        End If
    Next
    
    Dim retVal As Integer
    If Not ComprobarProveedor Then
        retVal = MsgBox("Existen partidas que no tienen PROVEEDOR asignado y es un campo obligatorio para generar la Orden de Compra", vbExclamation, "SAGE MAS 500 - Editar de Requisiciòn")
    End If
End Function

Private Sub EnabledFormFields(pVal As Boolean)
    pVal = pVal And Not mbReqClosed 'Agregado por Multiconsulting
' If the Requisition is closed or rejected, disable all fields in the header otherwise, enable them
'    mbReqClosed = Not pVal
    moGM.MenuAdd = pVal
    moGM.MenuDelete = pVal
    calDate.Enabled = pVal
    chkExpedite.Enabled = pVal
    cboExpReason.Enabled = (chkExpedite = 1) And pVal
    lkuDept.Enabled = pVal
    
    If mbIntegrateWithIM Then
        lkuWarehouse.Enabled = pVal
    End If
    
    txtComment.Enabled = pVal
    cmdGenerate.Enabled = IIf((evSegUserBuyInfo = 1 Or evSegUserAcpt = 1), True, False)
        
' If the req is closed, remove the row automatically appended to the end of the grid
    If mbReqClosed Then
        RemoveLastRowFromGrid
    End If
    
    If pVal Then
        gGridUnlockGrid grdReqLines
        gGridUnlockGrid grdReqLineDtl
    Else
        gGridLockGrid grdReqLines
        gGridLockGrid grdReqLineDtl
    End If
'    UpdateSegEstatusReq (True)
    
' If the status is not closed change the back color to window
    If pVal Then
        txtOriginator.BackColor = vbWindowBackground
        txtContact.BackColor = vbWindowBackground
        txtComment.BackColor = vbWindowBackground
        'txtAceptaReq.BackColor = vbWindowBackground
        'txtAutorizaReq.BackColor = vbWindowBackground
    Else
' Otherwise, change the back color to buttonface
        txtOriginator.BackColor = vbButtonFace
        txtContact.BackColor = vbButtonFace
        txtComment.BackColor = vbButtonFace
        txtAceptaReq.BackColor = vbButtonFace
        txtAutorizaReq.BackColor = vbButtonFace
    End If
    
    If chkExpedite.Enabled And chkExpedite.Value = 1 Then
        cboExpReason.BackColor = vbWindowBackground
    Else
        cboExpReason.BackColor = vbButtonFace
    End If
End Sub

Private Sub UpdateEstatusFields(pVal As Integer)  ', pValPrev As Integer)

'MsgBox "UpdateStatusFields: " & pVal
    
    Select Case pVal
        
        Case kReqStatusPending
            If ((evSegUserBuyInfo = 1 And loadNewReq = True And firstCompEstatusReq = False) Or evSegUserAcpt = 1) Then
                gGridLockGrid grdReqLines
                gGridLockGrid grdReqLineDtl
                lkuDept.Enabled = True
                lkuWarehouse.Enabled = True
                txtComment.Enabled = False
                calDate.Enabled = False
                cboExpReason.Enabled = False
                chkExpedite.Enabled = False
                txtAutorizaReq.Text = ""
                calAutorizaReq = ""
                txt2doAutorizaReq.Text = ""
                cal2doAutorizaReq = ""
                txtAceptaReq.Text = ""
                calAceptaReq = ""
                chkbEstatusDesc.Enabled = True
                chkb2doAuthNeed.Enabled = False
                cmdGenerate.Enabled = False
                loadNewReq = False
                
                If (evSegUserAcpt = 1 And evSegUserBuyInfo = 0) Then
                    gGridLockGrid grdReqLines
                    gGridLockGrid grdReqLineDtl
                    gGridUnlockColumn grdReqLines, kColReqBuyerID
                End If
                
                Exit Sub
            End If
            
            
            If firstCompEstatusReq = False Or (firstCompEstatusReq = True And pVal = 0) Then
                lkuDept.Enabled = True
                txtComment.Enabled = True
                calDate.Enabled = True
                cboExpReason.Enabled = True
                chkExpedite.Enabled = True
                gGridUnlockGrid grdReqLines
                gGridUnlockGrid grdReqLineDtl
                gGridLockColumn grdReqLines, kcolReqVendID
                gGridLockColumn grdReqLines, kColReqBuyerID
                gGridLockColumn grdReqLines, kColReqLStItemID
                gGridLockColumn grdReqLines, kColReqLTBItemID
                chkb2doAutorizoReq.Value = 0
                chkb2doAutorizoReq.Enabled = False
                chkb2doAuthNeed.Enabled = False
                chkb2doAuthNeed.Value = 0
                txtAutorizaReq.Text = ""
                calAutorizaReq = ""
                txt2doAutorizaReq.Text = ""
                cal2doAutorizaReq = ""
                txtAceptaReq.Text = ""
                calAceptaReq = ""
                chkbEstatusDesc.Enabled = True
                
                If msButtonInitAcc = "kTbNextNumber" Then
                    txtUserMod.Text = msCurrentUser
                    sddEstatusReq.ListIndex = sddEstatusReq.GetIndexByItemData(kReqStatusPending)
                End If
            End If
            
           ' sddEstatusReq.Locked = IIf(evSegUserBuyInfo = 0 Or evSegUserAcpt = 1, False, True)
        Case kReqStatusAuthorized
            gGridLockGrid grdReqLines
            gGridLockGrid grdReqLineDtl
            lkuDept.Enabled = False
            lkuWarehouse.Enabled = False
            txtComment.Enabled = False
            calDate.Enabled = False
            cboExpReason.Enabled = False
            chkExpedite.Enabled = False
            txtAceptaReq.Text = ""
            calAceptaReq = ""
            chkbEstatusDesc.Enabled = True
                        
            If (chkb2doAuthNeed.Value = 1 And evSegUserAutz = 1) Or evSegUserAutz = 1 Then
                chkb2doAuthNeed.Enabled = True
                chkb2doAutorizoReq.Enabled = True
            Else
                chkb2doAuthNeed.Enabled = False
                chkb2doAutorizoReq.Enabled = False
            End If
            
          '  sddEstatusReq.Locked = IIf((evSegUserAcpt = 1 Or evSegUserAutz = 1), False, True)
                    
        Case kReqStatusAccepted
            If evSegUserBuyInfo = 0 And evSegUserAcpt = 0 Then
                gGridLockGrid grdReqLines
                gGridLockGrid grdReqLineDtl
                lkuDept.Enabled = True
                lkuWarehouse.Enabled = True
                calAutorizaReq = ""
                txt2doAutorizaReq.Text = ""
                cal2doAutorizaReq = ""
                txtAceptaReq.Text = ""
                calAceptaReq = ""
                chkExpedite.Enabled = False
                cboExpReason.Enabled = False
                chkbEstatusDesc.Enabled = True
                chkb2doAuthNeed.Enabled = False
                cmdGenerate.Enabled = IIf((evSegUserBuyInfo = 1 Or evSegUserAcpt = 1), True, False)
                loadNewReq = False
                Exit Sub
            End If
            
            chkb2doAuthNeed.Enabled = False
            chkb2doAutorizoReq.Enabled = False
            lkuWarehouse.Enabled = True
            lkuWarehouse.EnabledLookup = True
            lkuDept.Enabled = True
            lkuDept.EnabledLookup = True
'            'Agregado por multiconsulting
            If mbIntegratedCT And lGetReqContract > 0 Then
                gGridLockColumn grdReqLines, kcolReqVendID
            Else
                gGridUnlockColumn grdReqLines, kcolReqVendID
            End If
'            'Agregado por multiconsulting
            If Not mbReqClosed Then
                gGridUnlockColumn grdReqLines, kColReqWhseID
                gGridUnlockColumn grdReqLines, kColReqPurchDeptID
                gGridUnlockColumn grdReqLines, kColReqItemID
                gGridUnlockColumn grdReqLines, kColReqLStItemID
                gGridUnlockColumn grdReqLines, kColReqLTBItemID
            End If
            If evSegUserAcpt = 0 Then
                gGridLockColumn grdReqLines, kColReqBuyerID
                loadNewReq = False
            End If
            
            gGridLockColumn grdReqLines, kColReqDescription
            gGridLockColumn grdReqLines, kColReqQtyRequested
            gGridLockColumn grdReqLines, kColReqUnitMeasID
            gGridLockColumn grdReqLines, kcolReqRequestDate
            gGridLockColumn grdReqLines, kColReqEstPres
                        
            txtComment.Enabled = False
            calDate.Enabled = False
            cboExpReason.Enabled = False
            chkExpedite.Enabled = False
            chkbEstatusDesc.Enabled = True
            cmdGenerate.Enabled = IIf((evSegUserBuyInfo = 1 Or evSegUserAcpt = 1), True, False)
            'sddEstatusReq.Locked = IIf((evSegUserAcpt = 1 Or evSegUserAutz = 1), False, True)
            
            If (evSegUserAcpt = 1 And evSegUserBuyInfo = 0) Then
                gGridLockGrid grdReqLines
                gGridLockGrid grdReqLineDtl
                gGridUnlockColumn grdReqLines, kColReqBuyerID
            End If
    End Select
    
End Sub

'Contratación
Public Function lGetReqContract() As Long
    Dim lContractKey As Long
    
    On Error GoTo ErrorHandler
    
    lGetReqContract = -1
        
    lContractKey = glGetValidLong(moClass.moAppDB.Lookup("ContractKey", "tpoRequisitionContract", "ReqKey =" & glGetValidLong(moDmForm.GetColumnValue("ReqKey"))))
    If lContractKey <> 0 Then
        lGetReqContract = lContractKey
    End If
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
End Function

Public Sub SetReqContract(lContractKey As Long)
    Dim sSql As String
    On Error GoTo ErrorHandler
    If lContractKey > 0 Then
        If giGetValidInt(moClass.moAppDB.Lookup("count(*)", "tpoRequisitionContract", "ReqKey =" & glGetValidLong(moDmForm.GetColumnValue("ReqKey")))) = 0 Then
            sSql = "insert into tpoRequisitionContract values (" & glGetValidLong(moDmForm.GetColumnValue("ReqKey")) & ", " & lContractKey & ")"
        Else
            sSql = "update tpoRequisitionContract set ContractKey = " & lContractKey & " where ReqKey =" & glGetValidLong(moDmForm.GetColumnValue("ReqKey"))
        End If
    Else
        sSql = "delete from tpoRequisitionContract where ReqKey=" & glGetValidLong(moDmForm.GetColumnValue("ReqKey"))
    End If
    
    moClass.moAppDB.ExecuteSQL sSql
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
End Sub


Public Function bAllowContractChange() As Boolean
    bAllowContractChange = False
    If grdReqLines.DataRowCnt > 0 Then
        If Not (grdReqLines.DataRowCnt = 1 And Len(Trim$(gsGridReadCellText(grdReqLines, 1, kColReqItemID))) = 0) Then
            Exit Function
        End If
    End If
    
    bAllowContractChange = True
End Function

Private Sub IsCTIntegrated()
    On Error GoTo ErrorHandler
    mbIntegratedCT = gbGetValidBoolean(moClass.moAppDB.Lookup("IntegrateWithCT", "tpoOptions", "CompanyID=" & gsQuoted(msCompanyID)))
    Exit Sub
ErrorHandler:
    mbIntegratedCT = False
End Sub


Private Function sGetAvaliableItemsListFromContract(lContractKey As Long) As String
    Dim sSql As String
    Dim rs As Object
    Dim bFirst As Boolean
    Dim sRetVal As String
    Dim lParentKey As Long
    
    On Error GoTo ErrorHandler
    
    sGetAvaliableItemsListFromContract = ""
    sRetVal = ""
    lkuWarehouse = msOldWhseID
    lParentKey = glGetValidLong(moClass.moAppDB.Lookup("ParentContractKey", "tctContract", "ContractKey=" & lContractKey))
    
    sSql = "SELECT p.ItemKey FROM tctContractLine AS p WHERE p.ContractKey = " & lContractKey
    If lParentKey <> 0 Then
        sSql = sSql & " or p.ContractKey =" & lParentKey
    End If
    
    Set rs = moClass.moAppDB.OpenRecordset(sSql, kSnapshot, kOptionNone)
    If rs.IsEmpty Then Exit Function
    
    While Not rs.IsEOF
        If bFirst Then
            sRetVal = sRetVal & ","
        Else
            bFirst = True
        End If
        sRetVal = sRetVal & rs.Field("ItemKey")
        rs.MoveNext
    Wend
    
    Set rs = Nothing
    sGetAvaliableItemsListFromContract = sRetVal
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error"
End Function

Private Function sGetDate(tDate As Date) As String
Dim sDate As String
    On Error GoTo ErrorHandler
    sGetDate = ""
    sDate = ""
    If Len(gsGetValidStr(DateTime.Month(tDate))) = 1 Then sDate = "0"
    sDate = sDate & gsGetValidStr(DateTime.Month(tDate))
    If Len(gsGetValidStr(DateTime.Day(tDate))) = 1 Then sDate = sDate & "0"
    sDate = sDate & gsGetValidStr(DateTime.Day(tDate)) & DateTime.Year(tDate)
    
    sGetDate = sDate
    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Function dGetMaxQtyAllowAtContract(lContractKey As Long, lItemKey As Long, lMainKey As Long, lRow As Long) As Double
    Dim dMaxQty As Double
    Dim dActualUsed As Double
    Dim i As Long
    Dim lParentKey As Long
    On Error GoTo ErrorHandler
    dGetMaxQtyAllowAtContract = 0

    lParentKey = glGetValidLong(moClass.moAppDB.Lookup("ParentContractKey", "tctContract", "ContractKey=" & lContractKey))

    If lParentKey <> 0 Then
     ' Correccion error
        dMaxQty = gdGetValidDbl(moClass.moAppDB.Lookup("sum(p.Qty)", "tctContractLine AS p join tctContract as s on p.ContractKey = s.ContractKey", "(p.ContractKey = " & lContractKey & " or p.ContractKey =" & lParentKey & ") AND p.ItemKey = " & lItemKey))
        dActualUsed = gdGetValidDbl(moClass.moAppDB.Lookup("sum(s.QtyReq)", "tpoReqLine AS p JOIN tpoReqLineDist AS s ON s.ReqLineKey = p.ReqLineKey JOIN tpoRequisitionContract AS t ON t.ReqKey = p.ReqKey", "(t.ContractKey = " & lContractKey & " or t.ContractKey = " & lParentKey & ") AND p.ItemKey = " & lItemKey & " and t.ReqKey <> " & lMainKey))
        dActualUsed = dActualUsed + gdGetValidDbl(moClass.moAppDB.Lookup("SUM(s.QtyOrd)", "tpoPOLine AS p JOIN tpoPOLineDist AS s ON s.POLineKey = p.POLineKey JOIN tpoPurchOrder AS t ON p.POKey = t.POKey  and t.Status <> 3", "(t.ContractKey = " & lContractKey & " or t.ContractKey = " & lParentKey & ") AND p.ItemKey = " & lItemKey))
    Else
        dMaxQty = gdGetValidDbl(moClass.moAppDB.Lookup("p.Qty", "tctContractLine AS p", "p.ContractKey = " & lContractKey & " AND p.ItemKey = " & lItemKey))
        dActualUsed = gdGetValidDbl(moClass.moAppDB.Lookup("sum(s.QtyReq)", "tpoReqLine AS p JOIN tpoReqLineDist AS s ON s.ReqLineKey = p.ReqLineKey JOIN tpoRequisitionContract AS t ON t.ReqKey = p.ReqKey", "t.ContractKey = " & lContractKey & " AND p.ItemKey = " & lItemKey & " and t.ReqKey <> " & lMainKey))
        dActualUsed = dActualUsed + gdGetValidDbl(moClass.moAppDB.Lookup("SUM(s.QtyOrd)", "tpoPOLine AS p JOIN tpoPOLineDist AS s ON s.POLineKey = p.POLineKey JOIN tpoPurchOrder AS t ON p.POKey = t.POKey  and t.Status <> 3", "t.ContractKey = " & lContractKey & " AND p.ItemKey = " & lItemKey))
    End If
    For i = 1 To grdReqLines.DataRowCnt
        If glGetValidLong(gsGridReadCell(grdReqLines, i, kColReqItemKey)) = lItemKey And i <> lRow Then
            dActualUsed = dActualUsed + gsGridReadCell(grdReqLines, i, kColReqQtyRequested)
        End If
    Next i
    
    If dMaxQty < dActualUsed Then
        dGetMaxQtyAllowAtContract = 0
    Else
        dGetMaxQtyAllowAtContract = dMaxQty - dActualUsed
    End If
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Function dGetContractCost(lContractKey As Long, lItemKey As Long) As Double
    Dim dCost As Double
    Dim lParentKey As Long
    On Error GoTo ErrorHandler
    
    lParentKey = glGetValidLong(moClass.moAppDB.Lookup("ParentContractKey", "tctContract", "ContractKey=" & lContractKey))
    If lParentKey <> 0 Then
        dCost = gdGetValidDbl(moClass.moAppDB.Lookup("p.UnitCost", "tctContractLine AS p join tctContract as s on s.ContractKey = p.ContractKey", "p.ItemKey = " & lItemKey & " AND (p.ContractKey = " & lContractKey & " or p.ContractKey =" & lParentKey & ") order by s.StartDate desc"))
    Else
        dCost = gdGetValidDbl(moClass.moAppDB.Lookup("p.UnitCost", "tctContractLine AS p ", "p.ItemKey = " & lItemKey & " AND p.ContractKey = " & lContractKey))
    End If
    
    dGetContractCost = dCost
    Exit Function
ErrorHandler:
    dGetContractCost = 0
    MsgBox Err.Description
End Function

Private Function lGetContractVendor(lContractKey As Long) As Long
    Dim lVendKey As Long
    On Error GoTo ErrorHandler
    lGetContractVendor = 0
    lVendKey = moClass.moAppDB.Lookup("VendorKey", "tctContract", "ContractKey=" & lContractKey)
    lGetContractVendor = lVendKey
    Exit Function
ErrorHandler:
    lGetContractVendor = 0
End Function

Private Function lGetUMFromContract(lItemKey As Long, lContractKey As Long) As Long
    Dim lParentKey As Long
    On Error GoTo ErrorHandler
    lParentKey = glGetValidLong(moClass.moAppDB.Lookup("ParentContractKey", "tctContract", "ContractKey=" & lContractKey))
    If lParentKey <> 0 Then
        lGetUMFromContract = giGetValidInt(moClass.moAppDB.Lookup("p.UnitMeasKey", "tctContractLine AS p join tctContract as s on s.ContractKey = p.ContractKey", "p.ItemKey = " & lItemKey & " AND (p.ContractKey =" & lContractKey & " or p.ContractKey = " & lParentKey & ") order by s.StartDate desc"))
    Else
        lGetUMFromContract = giGetValidInt(moClass.moAppDB.Lookup("p.UnitMeasKey", "tctContractLine AS p", "p.ItemKey = " & lItemKey & " AND p.ContractKey =" & lContractKey))
    End If
    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function
Private Function bIsValidQty() As Boolean
    Dim i As Long
    On Error GoTo ErrorHandler
    bIsValidQty = False
    For i = 1 To grdReqLines.DataRowCnt
        If glGetValidLong(gsGridReadCell(grdReqLines, i, kColReqItemKey)) <> 0 Then
            If gdGetValidDbl(gsGridReadCell(grdReqLines, i, kColReqQtyRequested)) <= 0 Then
                MsgBox "No se pueden tener partidas con 0 en una requisición", vbExclamation, "Alerta"
                Exit Function
            End If
        End If
    Next i
    bIsValidQty = True
    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Function bIsValidWhses() As Boolean
    Dim i As Long
    On Error GoTo ErrorHandler
    bIsValidWhses = False
    For i = 1 To grdReqLines.DataRowCnt
        If glGetValidLong(gsGridReadCell(grdReqLines, i, kColReqItemKey)) <> 0 Then
            If Len(Trim$(gsGetValidStr(gsGridReadCell(grdReqLines, i, kColReqWhseID)))) = 0 Then
                MsgBox "No se pueden tener partidas sin Almacén", vbExclamation, "Alerta"
                Exit Function
            End If
        End If
    Next i
    bIsValidWhses = True
    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Function bIsValidVendors() As Boolean
    Dim i As Long
    On Error GoTo ErrorHandler
    bIsValidVendors = False
    For i = 1 To grdReqLines.DataRowCnt
        If glGetValidLong(gsGridReadCell(grdReqLines, i, kColReqItemKey)) <> 0 Then
            If Len(Trim$(gsGetValidStr(gsGridReadCell(grdReqLines, i, kcolReqVendID)))) = 0 Then
                MsgBox "No se pueden tener partidas sin Proveedor", vbExclamation, "Alerta"
                Exit Function
            End If
        End If
    Next i
    bIsValidVendors = True
    Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function
'Agregado por Multiconsulting


