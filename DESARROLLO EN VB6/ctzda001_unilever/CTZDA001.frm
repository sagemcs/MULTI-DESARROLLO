VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{CAF0FDE4-8332-11CF-BC13-0020AFD6738C}#1.0#0"; "newsota.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{BC90D6A3-491E-451B-ADED-8FABA0B8EE36}#57.0#0"; "SOTADropDown.ocx"
Object = "{2A076741-D7C1-44B1-A4CB-E9307B154D7C}#185.0#0"; "EntryLookupControls.ocx"
Object = "{C41A85E3-4CB6-40B5-B425-EE9ECC5E6F06}#181.0#0"; "SOTATbar.ocx"
Object = "{F2F2EE3C-0D23-4FC8-944C-7730C86412E3}#67.0#0"; "sotasbar.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{0FA91D91-3062-44DB-B896-91406D28F92A}#65.0#0"; "SOTACalendar.ocx"
Object = "{9504980C-B928-4BF5-A5D0-13E1F649AECB}#45.0#0"; "SOTAVM.ocx"
Object = "{8A9C5D3D-5A2F-4C5F-A12A-A955C4FB68C8}#101.0#0"; "LookupView.ocx"
Begin VB.Form frmContract 
   Caption         =   "Gestionar Contratos"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   Icon            =   "CTZDA001.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6765
   ScaleWidth      =   10470
   Begin SOTADropDownControl.SOTADropDown sddClasification 
      Height          =   315
      Left            =   8040
      TabIndex        =   85
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
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
      Text            =   "sddClasification"
   End
   Begin NEWSOTALib.SOTAMaskedEdit txtChgNo 
      Height          =   315
      Left            =   3960
      TabIndex        =   68
      Top             =   600
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   556
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      bProtected      =   -1  'True
   End
   Begin LookupViewControl.LookupView lkuNav 
      Height          =   285
      Left            =   2640
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   600
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      LookupID        =   "Contract"
   End
   Begin NEWSOTALib.SOTAMaskedEdit txtContract 
      Height          =   315
      Left            =   1080
      TabIndex        =   58
      Top             =   600
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   556
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SOTADropDownControl.SOTADropDown sddType 
      Height          =   315
      Left            =   4920
      TabIndex        =   33
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
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
      StaticListTableName=   "tctContract"
   End
   Begin VB.Frame CustomFrame 
      Caption         =   "Frame"
      Enabled         =   0   'False
      Height          =   1035
      Index           =   0
      Left            =   -10000
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton CustomButton 
      Caption         =   "Button"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   -10000
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2730
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CheckBox CustomCheck 
      Caption         =   "Check"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   -10000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2310
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ComboBox CustomCombo 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H80000012&
      Height          =   315
      Index           =   0
      Left            =   -10000
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "CustomCombo"
      Top             =   1530
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.OptionButton CustomOption 
      Caption         =   "Option"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   -10000
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      Width           =   1245
   End
   Begin StatusBar.SOTAStatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      ToolTipText     =   "7.30.1001"
      Top             =   6375
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   688
   End
   Begin SOTAToolbarControl.SOTAToolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   741
   End
   Begin SOTAVM.SOTAValidationMgr valMgr 
      Left            =   -10000
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   661
      CtlsCount       =   0
   End
   Begin NEWSOTALib.SOTAMaskedEdit CustomMask 
      Height          =   285
      Index           =   0
      Left            =   -10000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
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
   End
   Begin NEWSOTALib.SOTANumber CustomNumber 
      Height          =   285
      Index           =   0
      Left            =   -10000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
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
      mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
      text            =   "           0.00"
      sDecimalPlaces  =   2
   End
   Begin NEWSOTALib.SOTACurrency CustomCurrency 
      Height          =   285
      Index           =   0
      Left            =   -10000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1140
      Visible         =   0   'False
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
      mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
      text            =   "           0.00"
      sDecimalPlaces  =   2
   End
   Begin MSComCtl2.UpDown CustomSpin 
      Height          =   285
      Index           =   0
      Left            =   -10000
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
   End
   Begin SOTACalendarControl.SOTACalendar CustomDate 
      Height          =   315
      Index           =   0
      Left            =   -10000
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2880
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
   Begin NEWSOTALib.SOTACustomizer picDrag 
      Height          =   615
      Index           =   0
      Left            =   -10000
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1085
      _StockProps     =   0
   End
   Begin TabDlg.SSTab tabDataEntry 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   8705
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&General"
      TabPicture(0)   =   "CTZDA001.frx":532B
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "pnlTab(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalles"
      TabPicture(1)   =   "CTZDA001.frx":5347
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "pnlTab(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Operaciones"
      TabPicture(2)   =   "CTZDA001.frx":5363
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pnlTab(2)"
      Tab(2).ControlCount=   1
      Begin Threed.SSPanel pnlTab 
         Height          =   4500
         Index           =   0
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   10080
         _Version        =   65536
         _ExtentX        =   17780
         _ExtentY        =   7937
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.Frame fraVendor 
            Caption         =   "Proveedor"
            Height          =   4215
            Left            =   4320
            TabIndex        =   46
            Top             =   240
            Width           =   5775
            Begin VB.TextBox txtContratokey1 
               Height          =   375
               Left            =   3480
               TabIndex        =   97
               Top             =   2640
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.CommandButton cmdBankInfo 
               Caption         =   "Información Bancaria"
               Height          =   360
               Left            =   2280
               TabIndex        =   96
               Top             =   3240
               Width           =   1815
            End
            Begin VB.TextBox txtContractBankInfKey 
               Height          =   375
               Left            =   3600
               TabIndex        =   95
               Top             =   2160
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtBankAddress 
               Height          =   375
               Left            =   3480
               TabIndex        =   94
               Top             =   1680
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtSWIFT 
               Height          =   375
               Left            =   3480
               TabIndex        =   93
               Top             =   1200
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtAccountID 
               Height          =   285
               Left            =   3480
               TabIndex        =   92
               Top             =   840
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtTitular 
               Height          =   375
               Left            =   3480
               TabIndex        =   91
               Top             =   360
               Visible         =   0   'False
               Width           =   1215
            End
            Begin SOTADropDownControl.SOTADropDown sddFOB 
               Height          =   315
               Left            =   1320
               TabIndex        =   74
               Top             =   1800
               Width           =   2055
               _ExtentX        =   3625
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
            Begin EntryLookupControls.TextLookup lkuVendClass 
               Height          =   315
               Left            =   1320
               TabIndex        =   57
               Top             =   720
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               ForeColor       =   -2147483640
               IsSurrogateKey  =   -1  'True
               LookupID        =   "VendorClass"
               ParentIDColumn  =   "VendClassID"
               ParentKeyColumn =   "VendClassKey"
               ParentTable     =   "tapVendClass"
               BoundColumn     =   "VendClassKey"
               BoundTable      =   "tctContract"
               IsForeignKey    =   -1  'True
               Datatype        =   0
               sSQLReturnCols  =   "VendClassID,,;"
            End
            Begin EntryLookupControls.TextLookup lkuContact 
               Height          =   315
               Left            =   1320
               TabIndex        =   48
               Top             =   1080
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               ForeColor       =   -2147483640
               IsSurrogateKey  =   -1  'True
               LookupID        =   "Contact"
               ParentIDColumn  =   "Name"
               ParentKeyColumn =   "CntctKey"
               ParentTable     =   "tciContact"
               BoundColumn     =   "CntctKey"
               BoundTable      =   "tctContract"
               IsForeignKey    =   -1  'True
               Datatype        =   0
               sSQLReturnCols  =   "Name,,;Title,,;CntctKey,,;"
            End
            Begin EntryLookupControls.TextLookup lkuVendor 
               Height          =   315
               Left            =   1320
               TabIndex        =   75
               Top             =   360
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   556
               ForeColor       =   -2147483640
               IsSurrogateKey  =   -1  'True
               LookupID        =   "Vendor"
               ParentIDColumn  =   "VendID"
               ParentKeyColumn =   "VendKey"
               ParentTable     =   "tapVendor"
               BoundColumn     =   "VendorKey"
               BoundTable      =   "tctCountract"
               IsForeignKey    =   -1  'True
               Datatype        =   0
               sSQLReturnCols  =   "VendID,lkuVendor,;VendName,lblVendorName,;"
            End
            Begin SOTADropDownControl.SOTADropDown sddPaymentTerms 
               Height          =   315
               Left            =   1320
               TabIndex        =   77
               Top             =   2160
               Width           =   2055
               _ExtentX        =   3625
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
               Style           =   2
               Text            =   "sddPaymentTerms"
            End
            Begin SOTADropDownControl.SOTADropDown SddCurrID 
               Height          =   315
               Left            =   1320
               TabIndex        =   79
               Top             =   2640
               Width           =   975
               _ExtentX        =   1720
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
               Text            =   "SddCurrID"
            End
            Begin VB.CommandButton cmdGenParams 
               Caption         =   "Parametros Generales"
               Height          =   360
               Left            =   120
               TabIndex        =   90
               Top             =   3240
               Width           =   1815
            End
            Begin VB.Label lblVendorName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   555
               Left            =   3120
               TabIndex        =   81
               Top             =   360
               Width           =   2445
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblMoneda 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda"
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   2640
               Width           =   585
            End
            Begin VB.Label lblCondDe 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cond de Pago"
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   2160
               Width           =   1020
            End
            Begin VB.Label lblProveedor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedor"
               Height          =   195
               Left            =   120
               TabIndex        =   76
               Top             =   360
               Width           =   735
            End
            Begin VB.Label lblFOB 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cond Entrega"
               Height          =   195
               Left            =   120
               TabIndex        =   73
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lblVendClass 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Clase "
               Height          =   195
               Left            =   120
               TabIndex        =   56
               Top             =   720
               Width           =   435
            End
            Begin VB.Label txtCountryID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   1320
               TabIndex        =   50
               Top             =   1440
               Width           =   1485
            End
            Begin VB.Label lblCountryID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "País"
               Height          =   195
               Left            =   120
               TabIndex        =   49
               Top             =   1440
               Width           =   330
            End
            Begin VB.Label lblContact 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contacto"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   1080
               Width           =   645
            End
         End
         Begin VB.Frame fraDates 
            Caption         =   "Fechas"
            Height          =   1935
            Left            =   240
            TabIndex        =   37
            Top             =   2520
            Width           =   3735
            Begin NEWSOTALib.SOTANumber nbrDuration 
               Height          =   300
               Left            =   1080
               TabIndex        =   43
               Top             =   1080
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   529
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
               mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               sDecimalPlaces  =   2
            End
            Begin SOTACalendarControl.SOTACalendar sclStartDate 
               Height          =   300
               Left            =   1080
               TabIndex        =   42
               Top             =   720
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
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
            End
            Begin SOTACalendarControl.SOTACalendar sclSignatureDate 
               Height          =   300
               Left            =   1080
               TabIndex        =   41
               Top             =   360
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
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
            End
            Begin VB.Label txtOutDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   1080
               TabIndex        =   45
               Top             =   1440
               Width           =   1845
            End
            Begin VB.Label lblOutDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vencimiento"
               Height          =   195
               Left            =   120
               TabIndex        =   44
               Top             =   1440
               Width           =   870
            End
            Begin VB.Label lblDuration 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vigencia"
               Height          =   195
               Left            =   120
               TabIndex        =   40
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label lblStartDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Inicio"
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   720
               Width           =   375
            End
            Begin VB.Label lblSignatureDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Firma"
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame fraContract 
            Caption         =   "Contrato"
            Height          =   2295
            Left            =   240
            TabIndex        =   34
            Top             =   120
            Width           =   3735
            Begin VB.TextBox currAmt 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               Height          =   315
               Left            =   840
               TabIndex        =   104
               Top             =   1440
               Width           =   2775
            End
            Begin EntryLookupControls.TextLookup lkuParentContract 
               Height          =   315
               Left            =   1440
               TabIndex        =   54
               Top             =   720
               Visible         =   0   'False
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               ForeColor       =   -2147483640
               IsSurrogateKey  =   -1  'True
               LookupID        =   "Contract"
               ParentIDColumn  =   "ContractNo"
               ParentKeyColumn =   "ContractKey"
               ParentTable     =   "tctContract"
               BoundColumn     =   "ParentContractKey"
               BoundTable      =   "tctContract"
               IsForeignKey    =   -1  'True
               Datatype        =   0
               sSQLReturnCols  =   "ContractKey,,;ContractID,,;ContractNo,lkuParentContract,;"
            End
            Begin SOTADropDownControl.SOTADropDown sddState 
               Height          =   315
               Left            =   1440
               TabIndex        =   52
               Top             =   1080
               Width           =   2055
               _ExtentX        =   3625
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
               Text            =   "sddState"
            End
            Begin NEWSOTALib.SOTAMaskedEdit txtContractNo 
               Height          =   285
               Left            =   1440
               TabIndex        =   35
               Top             =   360
               Width           =   2055
               _Version        =   65536
               _ExtentX        =   3625
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
            End
            Begin VB.Label txtAprobalLevel 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   1800
               TabIndex        =   83
               Top             =   1920
               Width           =   1725
            End
            Begin VB.Label lblAprobalLevel 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nivel Aprobación"
               Height          =   195
               Left            =   240
               TabIndex        =   82
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label lblTotalDel 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Importe"
               Height          =   195
               Left            =   240
               TabIndex        =   72
               Top             =   1440
               Width           =   525
            End
            Begin VB.Label txtParentContract 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   3600
               TabIndex        =   55
               Top             =   1440
               Width           =   45
            End
            Begin VB.Label lblParentContract 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Contrato"
               Height          =   195
               Left            =   240
               TabIndex        =   53
               Top             =   720
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblState 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estado"
               Height          =   195
               Left            =   240
               TabIndex        =   51
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label lblNoContrato 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Contrato"
               Height          =   195
               Left            =   240
               TabIndex        =   36
               Top             =   360
               Width           =   855
            End
         End
      End
      Begin Threed.SSPanel pnlTab 
         Height          =   4500
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   390
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   7937
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.Frame frmLineEntry 
            Height          =   2355
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   9885
            Begin VB.TextBox snOtrosCostos 
               Height          =   405
               Left            =   6960
               TabIndex        =   103
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox snSeguro 
               Height          =   375
               Left            =   5160
               TabIndex        =   102
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox snFlete 
               Height          =   375
               Left            =   3360
               TabIndex        =   101
               Top             =   1680
               Width           =   1695
            End
            Begin NEWSOTALib.SOTANumber nbrQtyVariation 
               Height          =   315
               Left            =   2400
               TabIndex        =   89
               Top             =   1080
               Width           =   1260
               _Version        =   65536
               _ExtentX        =   2222
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               mask            =   "<ILH>#|,###|,###|,##<ILp0>#<IRp0>|.#####"
               text            =   "         0.00000"
               dMaxValue       =   1E+20
               sIntegralPlaces =   10
               sDecimalPlaces  =   5
            End
            Begin SOTADropDownControl.SOTADropDown sddLineType 
               Height          =   315
               Left            =   120
               TabIndex        =   88
               Top             =   1680
               Width           =   2295
               _ExtentX        =   4048
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
               Text            =   "sddLineType"
            End
            Begin SOTADropDownControl.SOTADropDown sddUM 
               Height          =   315
               Left            =   120
               TabIndex        =   71
               Top             =   1080
               Width           =   855
               _ExtentX        =   1508
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
               Style           =   2
               Text            =   "sddUM"
            End
            Begin NEWSOTALib.SOTACurrency currUnitCost 
               Height          =   315
               Left            =   1080
               TabIndex        =   69
               Top             =   1080
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               mask            =   "<HL> <ILH>###|,###|,##<ILp0>#<IRp0>|.#####"
               text            =   "        0.00000"
               sIntegralPlaces =   9
               sDecimalPlaces  =   5
            End
            Begin NEWSOTALib.SOTANumber numDeliveryTime 
               Height          =   315
               Left            =   5400
               TabIndex        =   67
               Top             =   480
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               dMaxValue       =   9999999
               sDecimalPlaces  =   2
            End
            Begin NEWSOTALib.SOTANumber numRoundValue 
               Height          =   315
               Left            =   6960
               TabIndex        =   30
               Top             =   480
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           1.00"
               dMinValue       =   1
               dMaxValue       =   9999999
               sDecimalPlaces  =   2
            End
            Begin VB.TextBox txtDescription 
               Height          =   285
               Left            =   2400
               TabIndex        =   66
               Top             =   480
               Width           =   2895
            End
            Begin NEWSOTALib.SOTACurrency currItemAmt 
               Height          =   315
               Left            =   3750
               TabIndex        =   64
               Top             =   1080
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               bProtected      =   -1  'True
               mask            =   "<HL> <ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           0.00"
               sDecimalPlaces  =   2
            End
            Begin EntryLookupControls.TextLookup lkuItem 
               Height          =   285
               Left            =   120
               TabIndex        =   61
               Top             =   480
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   503
               ForeColor       =   -2147483640
               LookupID        =   "Item"
               Datatype        =   0
               sSQLReturnCols  =   "ItemID,lkuItem,;ShortDesc,txtDescription,;ItemKey,,;"
            End
            Begin VB.CommandButton cmdUndo 
               Caption         =   "&Cancelar"
               Height          =   315
               Left            =   8640
               TabIndex        =   11
               Top             =   1080
               WhatsThisHelpID =   35807
               Width           =   1005
            End
            Begin VB.CommandButton cmdOK 
               Caption         =   "&Aceptar"
               Height          =   315
               Left            =   8640
               TabIndex        =   9
               Top             =   480
               WhatsThisHelpID =   35806
               Width           =   1005
            End
            Begin NEWSOTALib.SOTANumber numMaxLot 
               Height          =   315
               Left            =   6960
               TabIndex        =   24
               Top             =   1080
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           1.00"
               dMinValue       =   1
               dMaxValue       =   9999999
               sDecimalPlaces  =   2
            End
            Begin NEWSOTALib.SOTANumber numMinLot 
               Height          =   315
               Left            =   5400
               TabIndex        =   26
               Top             =   1080
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               mask            =   "<ILH>###|,###|,###|,##<ILp0>#<IRp0>|.##"
               text            =   "           1.00"
               dMinValue       =   1
               dMaxValue       =   9999999
               sDecimalPlaces  =   2
            End
            Begin NEWSOTALib.SOTANumber numItemQty 
               Height          =   315
               Left            =   2400
               TabIndex        =   28
               Top             =   1080
               Width           =   1260
               _Version        =   65536
               _ExtentX        =   2222
               _ExtentY        =   556
               _StockProps     =   93
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               mask            =   "<ILH>#|,###|,###|,##<ILp0>#<IRp0>|.#####"
               text            =   "         0.00000"
               sIntegralPlaces =   10
               sDecimalPlaces  =   5
            End
            Begin VB.Label lblSeguro 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Seguro"
               Height          =   195
               Left            =   5280
               TabIndex        =   100
               Top             =   1440
               Width           =   510
            End
            Begin VB.Label lblFlete 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Flete"
               Height          =   195
               Left            =   3480
               TabIndex        =   99
               Top             =   1440
               Width           =   345
            End
            Begin VB.Label lbl_otros_costos 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Otros Costos"
               Height          =   195
               Left            =   6960
               TabIndex        =   98
               Top             =   1440
               Width           =   900
            End
            Begin VB.Label lblLineType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo"
               Height          =   195
               Left            =   120
               TabIndex        =   87
               Top             =   1440
               Width           =   315
            End
            Begin VB.Label lblUM 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UM"
               Height          =   195
               Left            =   120
               TabIndex        =   70
               Top             =   840
               Width           =   255
            End
            Begin VB.Label lblDescription 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción"
               Height          =   255
               Left            =   2400
               TabIndex        =   65
               Top             =   240
               Width           =   3000
            End
            Begin VB.Label lblItemAmt 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Importe"
               Height          =   315
               Left            =   3750
               TabIndex        =   63
               Top             =   840
               Width           =   525
            End
            Begin VB.Label lblItemID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Artículo"
               Height          =   195
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   555
            End
            Begin VB.Label lblUnitCost 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Unitario"
               Height          =   195
               Left            =   1080
               TabIndex        =   31
               Top             =   840
               Width           =   990
            End
            Begin VB.Label lblChildRoundValue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor de Redondeo"
               Height          =   195
               Left            =   6960
               TabIndex        =   29
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblItemQty 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cantidad"
               Height          =   195
               Left            =   2400
               TabIndex        =   27
               Top             =   840
               Width           =   630
            End
            Begin VB.Label lblChildMinLot 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lote Mínimo"
               Height          =   195
               Left            =   5400
               TabIndex        =   25
               Top             =   840
               Width           =   885
            End
            Begin VB.Label lblChildMaxLot 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lote Máximo"
               Height          =   195
               Left            =   6960
               TabIndex        =   23
               Top             =   840
               Width           =   900
            End
            Begin VB.Label lblDeliveryTime 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo de Entrega"
               Height          =   195
               Left            =   5400
               TabIndex        =   22
               Top             =   240
               Width           =   1350
            End
            Begin VB.Shape shpFocusRect 
               Height          =   285
               Left            =   120
               Top             =   240
               Visible         =   0   'False
               Width           =   585
            End
         End
         Begin FPSpreadADO.fpSpread grdMain 
            Height          =   1875
            Left            =   120
            TabIndex        =   16
            Top             =   2400
            Width           =   9810
            _Version        =   524288
            _ExtentX        =   17304
            _ExtentY        =   3307
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "CTZDA001.frx":537F
            AppearanceStyle =   0
         End
      End
      Begin Threed.SSPanel pnlTab 
         Height          =   4500
         Index           =   2
         Left            =   -74940
         TabIndex        =   18
         Top             =   390
         Width           =   9840
         _Version        =   65536
         _ExtentX        =   17357
         _ExtentY        =   7937
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin FPSpreadADO.fpSpread grdOperations 
            Height          =   3615
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   9615
            _Version        =   524288
            _ExtentX        =   16960
            _ExtentY        =   6376
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "CTZDA001.frx":5769
            AppearanceStyle =   0
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label lblClasification 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clasificación"
      Height          =   195
      Left            =   7080
      TabIndex        =   84
      Top             =   600
      Width           =   885
   End
   Begin VB.Label lblExtra 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Cambio"
      Height          =   195
      Left            =   3000
      TabIndex        =   59
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   4440
      TabIndex        =   32
      Top             =   600
      Width           =   315
   End
   Begin VB.Label lblContractID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Identificador"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   870
   End
   Begin VB.Label CustomLabel 
      Caption         =   "Label"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   -10000
      TabIndex        =   17
      Top             =   60
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "frmContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public moSotaObjects            As New Collection

Private moContextMenu           As clsContextMenu

Private moLE                    As clsLineEntry

Private moOptions               As New clsModuleOptions

Private mbEnterAsTab            As Boolean

Private msCompanyID             As String

Private msBusDate              As String

Private msUserID                As String

Private miSecurityLevel         As Integer

Private miFilter                As Integer

Private mlLanguage              As Long

Private miMinFormHeight         As Long

Private miMinFormWidth          As Long

Private miOldFormHeight         As Long

Private miOldFormWidth          As Long

Public moClass                  As Object

Private moAppDB                 As Object

Private moSysSession            As Object

Private mbCancelShutDown        As Boolean

Private mbLoadSuccess           As Boolean

Private mbSaved                 As Boolean

Private mlRunMode               As Long

Public WithEvents moDmHeader    As clsDmForm
Attribute moDmHeader.VB_VarHelpID = -1

Private WithEvents moDmDetl     As clsDmGrid
Attribute moDmDetl.VB_VarHelpID = -1
Private moDmOperations          As clsDmGrid


Private Const kiHeaderTab = 0

Private Const kiDetailTab = 1

Private Const kiTotalsTab = 2

Const VBRIG_MODULE_ID_STRING = "frmContract.FRM"

Private mbInBrowse As Boolean

Private msNatCurrID As String

Private msHomeCurrID As String

Private muHomeCurrInfo As CurrencyInfo

Private muNatCurrInfo As CurrencyInfo

Private msContractNo As String
Private msContractNo1 As String
Private mlContractKey As Long

Private mdCurrExchRate As Double
Private mbFixedExchRate As Boolean
Private mlCurrExchSchdKey As Long
Private mlCurrExchSchdKeyOld As Long

Private mlActiveRow As Long

Private msLookupRestrict As String
Private bNum001 As Boolean

Private nOtrosCostos As Double
Private nFlete As Double
Private nSeguro As Double



Public sTit As String

Public ssAccountNo As String
Public ssBankAddress As String
Public ssSWIFT As String
Public ssTitular As String
Public bGrand As Boolean
Public bGrandD As Boolean

'Esta es la que agregue yo para la informacion del Banco
'Public moDmContractBankInf As clsDmForm


#If CUSTOMIZER Then
    Public moFormCust As Object
#End If

Public Property Get FormHelpPrefix() As String
    FormHelpPrefix = "CTZ"
End Property

Public Property Get WhatHelpPrefix() As String
    WhatHelpPrefix = "CTZ"
End Property

Public Property Get oClass() As Object
    Set oClass = moClass
End Property

Public Property Set oClass(oNewClass As Object)
    Set moClass = oNewClass
End Property

Public Property Get lRunMode() As Long
    lRunMode = mlRunMode
End Property

Public Property Let lRunMode(lNewRunMode As Long)
    mlRunMode = lNewRunMode
End Property

Public Property Get bCancelShutDown()
    bCancelShutDown = mbCancelShutDown
End Property

Public Property Get bLoadSuccess() As Boolean
    bLoadSuccess = mbLoadSuccess
End Property

Public Property Get bSaved() As Boolean
    bSaved = mbSaved
End Property

Public Property Let bSaved(bNewSaved As Boolean)
    mbSaved = bNewSaved
End Property

Private Function sMyName() As String
'+++ VB/Rig Skip +++
    sMyName = Me.Name
End Function

Private Sub SetupDropDowns()

    '-- Set up drop down properties (see APZDA001 for reference)
    sddType.InitStaticList moClass.moAppDB, "tctContract", "Type", mlLanguage
    sddLineType.InitStaticList moClass.moAppDB, "tctContractLine", "Type", mlLanguage
    sddClasification.InitStaticList moClass.moAppDB, "tctContract", "Clasification", mlLanguage
    sddPaymentTerms.InitDynamicList moClass.moAppDB, "SELECT p.PmtTermsID, p.PmtTermsKey FROM tciPaymentTerms AS p"
    sddUM.InitDynamicList moClass.moAppDB, "SELECT p.UnitMeasID, p.UnitMeasKey FROM tciUnitMeasure AS p"
    sddState.InitDynamicList moAppDB, "SELECT p.ContractStateId, p.ContractStateKey FROM tctContractState AS p "
    'WHERE p.CompanyId = " & gsQuoted(msCompanyID)
    sddState.InitDynamicList moAppDB, "SELECT p.ContractStateId, p.ContractStateKey FROM tctContractState AS p "
    'WHERE p.CompanyId = " & gsQuoted(msCompanyID)
    sddFOB.InitDynamicList moAppDB, "SELECT p.FOBID, p.FOBKey FROM tciFOB AS p "
    'WHERE p.CompanyID =  " & gsQuoted(msCompanyID)
    SddCurrID.InitDynamicList moAppDB, "SELECT p.CurrID FROM tmcCurrency AS p WHERE p.IsUsed = 1"
    
End Sub

Private Sub SetupLookups()
Dim oFramework As Object
    '-- Set up lookup properties (see APZDA001 for reference)
    Set oFramework = moClass.moFramework

   
    bSetupLkuNav
    
    With lkuVendor
        Set .Framework = oFramework
        Set .AppDatabase = moClass.moAppDB
        .RestrictClause = "Status = 1"
    End With
    
    With lkuContact
        Set .Framework = oFramework
        Set .AppDatabase = moClass.moAppDB
'        .RestrictClause = msLookupRestrict
    End With
    
    With lkuVendClass
        Set .Framework = oFramework
        Set .AppDatabase = moClass.moAppDB
'        .RestrictClause = msLookupRestrict
    End With
    
    With lkuParentContract
        Set .Framework = oFramework
        Set .AppDatabase = moClass.moAppDB
        .RestrictClause = sGetParentRestrict
    End With
    
    With lkuItem
        Set .Framework = oFramework
        Set .AppDatabase = moClass.moAppDB
        .RestrictClause = "Status = 1"
    End With
    
     Set oFramework = Nothing
End Sub


Private Sub SetupBars()

    Set sbrMain.Framework = moClass.moFramework
    

    With tbrMain
        .Style = sotaTB_TRANSACTION
        .RemoveButton kTbCopyFrom
        .RemoveButton kTbRenameId
        .AddSeparator .GetIndex(kTbHelp)
        sbrMain.BrowseVisible = True

        .LocaleID = mlLanguage
    End With
End Sub

Private Sub SetupModuleVars()
    '-- Currency variables
    msNatCurrID = msHomeCurrID
    
    '-- Assign the initial filter value
    miFilter = RSID_UNFILTERED

End Sub

Private Sub BindForm()
    BindHeader
    BindDetail
   ' BindBankInf
End Sub

Private Sub BindHeader()
    '-- Bind the parent DM object
    Set moDmHeader = New clsDmForm

    With moDmHeader
        Set .Form = frmContract
        Set .Session = moSysSession
        Set .Database = moAppDB
        Set .Toolbar = tbrMain
        Set .SOTAStatusBar = sbrMain
        Set .ValidationMgr = valMgr
        .AppName = gsStripChar(Me.Caption, ".")
        .UniqueKey = "ContractKey"
        .Bind Nothing, "CompanyID", SQL_CHAR

        .OrderBy = "CompanyID, ContractID"
        .AccessType = kDmBuildQueries
        .Table = "tctContract"
        
        .Bind txtContract, "ContractID", SQL_CHAR
        .Bind Nothing, "ContractKey", SQL_INTEGER
     '   .Bind txtContratokey1, "ContractKey", SQL_INTEGER
        .Bind txtCountryID, "CountryID", SQL_CHAR
        .Bind txtChgNo, "ChngOrdNo", SQL_INTEGER
        .Bind sclSignatureDate, "SignatureDate", SQL_DATE
        .Bind sclStartDate, "StartDate", SQL_DATE
        .Bind nbrDuration, "Duration", SQL_INTEGER
        .Bind txtContractNo, "ContractNo", SQL_CHAR
        .Bind currAmt, "ContractAmt", SQL_DECIMAL
        .Bind Nothing, "SeqNo", SQL_INTEGER
        .BindLookup lkuVendor
        .BindLookup lkuContact
        .BindLookup lkuVendClass
        .BindLookup lkuParentContract, kDmSetNull
        .Bind sddPaymentTerms, "PmtTermsKey", SQL_INTEGER, kDmUseItemData
        .Bind sddType, "Type", SQL_INTEGER, kDmUseItemData
        .Bind sddClasification, "Clasification", SQL_INTEGER, kDmUseItemData
        .Bind sddState, "State", SQL_INTEGER, kDmUseItemData
        .Bind sddFOB, "FOBKey", SQL_INTEGER, kDmUseItemData
        .Bind SddCurrID, "CurrID", SQL_CHAR
        
        '.Bind txtContratokey1, "ContractKey", SQL_INTEGER
        .Bind txtTitular, "Titular", SQL_VARCHAR
        .Bind txtAccountID, "AccountID", SQL_VARCHAR
        .Bind txtSWIFT, "SWIFT", SQL_VARCHAR
        .Bind txtBankAddress, "BankAddress", SQL_VARCHAR
       ' .Bind txtContractBankInfKey, "ContractBankInfKey", SQL_CHAR
        .Bind snOtrosCostos, "OtherExpenses", SQL_DECIMAL
        .Bind snFlete, "Flete", SQL_DECIMAL
        .Bind snSeguro, "Seguro", SQL_DECIMAL
        .Bind Nothing, "CreateDate", SQL_DATE
        .Bind Nothing, "UpdateDate", SQL_DATE
        .Bind Nothing, "CreateUser", SQL_CHAR
        .Bind Nothing, "UpdateUser", SQL_CHAR, kDmSetNull
        .Bind Nothing, "ChngOrdReason", SQL_CHAR
        .Bind Nothing, "ChngOrdReasonCodeKey", SQL_INTEGER
        .LinkSource "tapVendor", "VendKey = <<VendorKey>>"
        .Link lblVendorName, "VendName", SQL_CHAR
        .Init
    End With
    
    
'    Set moDmContractBankInf = New clsDmForm
'
'    With moDmContractBankInf
'        Set .Session = moSysSession
'        Set .Database = moAppDB
'        Set .Form = Me
'
'        'Set .Grid = grdOperations
'        .Table = "tctContractBankInf"
'        .UniqueKey = "ContractBankInfKey"
'
'        Set .Parent = moDmHeader
'        .ParentLink "ContractKey", "ContractKey", SQL_INTEGER
'        .Bind Nothing, "ContractKey", SQL_INTEGER
'        .Bind txtTitular, "Titular", SQL_CHAR
'        .Bind txtAccountID, "AccountID", SQL_INTEGER
'        .Bind txtSWIFT, "SWIFT", SQL_CHAR
'        .Bind txtBankAddress, "BankAddress", SQL_DECIMAL
'        .Bind Nothing, "ContractBankInfKey", SQL_CHAR
'
'
'
'        .Init
'    End With
    
    
    
End Sub

'Private Sub BindBankInf()
'  Set moDmContractBankInf = New clsDmForm
'
'    With moDmContractBankInf
'        Set .Session = moSysSession
'        Set .Database = moAppDB
'        Set .Form = Me
'
'        'Set .Grid = grdOperations
'        .Table = "tctContractBankInf"
'        .UniqueKey = "ContractBankInfKey"
'
'        .ParentLink "ContractKey", "ContractKey", SQL_INTEGER
'
'       ' Set .Parent = moDmHeader
'        '.ParentLink "ContractKey", "ContractKey", SQL_INTEGER
'        .Bind Nothing, "ContractKey", SQL_INTEGER
'        .Bind txtTitular, "Titular", SQL_CHAR
'        .Bind txtAccountID, "AccountID", SQL_INTEGER
'        .Bind txtSWIFT, "SWIFT", SQL_CHAR
'        .Bind txtBankAddress, "BankAddress", SQL_DECIMAL
'        .Bind Nothing, "ContractBankInfKey", SQL_CHAR
'
'
'
'        .Init
'    End With
'End Sub

Private Sub BindDetail()
    
    Set moDmDetl = New clsDmGrid

    With moDmDetl
        Set .Session = moSysSession
        Set .Database = moAppDB
        Set .Form = frmContract
        Set .Parent = moDmHeader
        .UniqueKey = "ContractLineKey"
        .OrderBy = "ContractLineKey"

        '-- Bind the detail columns
        Set .Grid = grdMain
        .Table = "tctContractLine"
        .ParentLink "ContractKey", "ContractKey", SQL_INTEGER
        
        .BindColumn "ContractLineKey", kColContractLineKey, SQL_INTEGER
'       .BindColumn "ContractKey", kColContractKey, SQL_INTEGER
        .BindColumn "SeqNo", kColSeqNo, SQL_INTEGER
        .BindColumn "ItemKey", kColItemKey, SQL_INTEGER
        .BindColumn "Description", kColDescription, SQL_CHAR, txtDescription
        .BindColumn "UnitMeasKey", kColUnitMeasKey, SQL_INTEGER
        .BindColumn "UnitCost", kColUnitCost, SQL_DECIMAL, currUnitCost
        .BindColumn "Qty", kColItemQty, SQL_DECIMAL, numItemQty
        .BindColumn "LineAmt", kColLineAmt, SQL_DECIMAL, currItemAmt
        
        
        .BindColumn "MaxLot", kColMaxLot, SQL_DECIMAL, numMaxLot
        .BindColumn "MinLot", kColMinLot, SQL_DECIMAL, numMinLot
        
        .BindColumn "RoundValue", kColRoundValue, SQL_DECIMAL, numRoundValue
        .BindColumn "DeliveryTime", kColDeliveryTime, SQL_INTEGER, numRoundValue
        .BindColumn "Type", kColType, SQL_INTEGER
        .BindColumn "QtyVariation", kColQtyVariation, SQL_DECIMAL, nbrQtyVariation
        
        .BindColumn "CreateDate", kColCreateDate, SQL_DATE
        .BindColumn "CreateUser", kColCreateUser, SQL_CHAR
        .BindColumn "UpdateDate", kColUpdateDate, SQL_DATE
        .BindColumn "UpdateUser", kColUpdateUser, SQL_CHAR
        
        

        .LinkSource "timItem", "tctContractLine.ItemKey=timItem.ItemKey", kDmJoin, LeftOuter
        .Link kColItemID, "ItemID", lkuItem
        .LinkSource "tciUnitMeasure", "tctContractLine.UnitMeasKey = tciUnitMeasure.UnitMeasKey", kDmJoin, LeftOuter
        .Link kColUnitMeasID, "UnitMeasID", Nothing, SQL_CHAR
        .Init
    End With
    
    moAppDB.ExecuteSQL "create table #tctContractOperations ([VoucherLineKey] [int] not null, [PoNo] [varchar](15) not null, [InvoiceNo] [varchar](15) not null, [Qty] [decimal](15,3) not null, [Description] [varchar](100) not null, [Amt] [decimal](15,3) not null, [AppDate] [datetime] not null) "
    
    Set moDmOperations = New clsDmGrid
    
    With moDmOperations
        Set .Session = moSysSession
        Set .Database = moAppDB
        Set .Form = frmContract
        Set .Grid = grdOperations
        .Table = "#tctContractOperations"
        .UniqueKey = "VoucherLineKey"
        
        .BindColumn "VoucherLineKey", kColOpertInvLineKey, SQL_INTEGER
        .BindColumn "PoNo", kColOpertPoNo, SQL_CHAR
        .BindColumn "InvoiceNo", kColOpertInvNo, SQL_CHAR
        .BindColumn "Qty", kColOpertQty, SQL_DECIMAL
        .BindColumn "Description", kColOpertDesc, SQL_CHAR
        .BindColumn "Amt", kColOpertAmt, SQL_DECIMAL
        .BindColumn "AppDate", kColOpertDate, SQL_DATE
        .Init
    End With
    
  
    
End Sub

Private Sub BindLE()

    Set moLE = New clsLineEntry
    
    With moLE
        Set .Grid = grdMain
        .GridType = kGridLineEntry
        Set .TabControl = tabDataEntry
        .TabDetailIndex = kiDetailTab
        Set .DM = moDmDetl
        Set .Ok = cmdOK
        Set .Undo = cmdUndo
        .SeqNoCol = 0
        .AllowGotoLine = True
        Set .FocusRect = shpFocusRect
        Set .Form = frmContract
        .Init
    End With
End Sub

Private Sub BindContextMenu()
    Set moContextMenu = New clsContextMenu

    With moContextMenu
        .BindGrid moLE, grdMain.hwnd
        Set .Form = frmContract
         .Init
    End With
End Sub

Public Sub HandleToolbarClick(sKey As String)
    Dim lKey As Long
    'On Error GoTo CancelInsert
    'bValid = True
    
    
' With moClass.moAppDB
'    .SetInParamStr "tctContractBankInf"
'    .SetOutParam lKey
'    .ExecuteSP "spGetNextSurrogateKey"
'    lKey = .GetOutParam(2)
'    .ReleaseParams
' End With
 'moDmDetl.SetColumnValue lRow, "ContractBankInfKey", lKey

'  txtContractBankInfKey.Text = lKey
'  txtContratokey1.Text = moDmHeader.GetColumnValue("ContractKey")
  
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If moFormCust.ToolbarClick(sKey) Then
            Exit Sub
        End If
    End If
#End If
'********************************************************************
' Description:
'    This routine is the central routine for processing events when a
'    button on the Toolbar is clicked.
'********************************************************************
    Select Case sKey
        Case kTbFinish, kTbFinishExit
            If Not valMgr.ValidateForm Then Exit Sub
            If Not moLE.GridEditDone() = kLeSuccess Then Exit Sub
            If moDmHeader.Action(kDmFinish) = kDmSuccess Then
                ClearDetlFields
            End If
            
            
            
            tabDataEntry.Tab = 0
        Case kTbSave
            If Not valMgr.ValidateForm Then Exit Sub
            If Not moLE.GridEditDone() = kLeSuccess Then Exit Sub
'            If txtAccountID.Text = "" Then
                If moDmHeader.Save(kDmSuccess) Then
                    'moDmContractBankInf.Save (True)
                End If
'            Else
'                 If moDmContractBankInf.Save(kDmSuccess) Then
'                    'moDmContractBankInf.Save (True)
'                End If
'                 If txtAccountID.Text <> "" And moDmHeader.GetColumnValue("ContractKey") Then
'                    moAppDB.ExecuteSQL "select Titular into : from tctContractBankInf where ContractKey =" & moDmHeader.GetColumnValue("ContractKey")
'
'                    moAppDB.ExecuteSQL "delete from tctContractAprobals where ContractKey =" & moDmHeader.GetColumnValue("ContractKey")
'                    moAppDB.ExecuteSQL "update tctContract set Free = 0 where ContractKey =" & moDmHeader.GetColumnValue("ContractKey")
'                 End If
'
                
                
           ' End If
            
        Case kTbCancel, kTbCancelExit
            ProcessCancel
            
        Case kTbDelete
           ' moDmHeader.Action (kDmDelete)
           ' moDmContractBankInf.Action (kDmDelete)
           ' ClearDetlFields
        Case kTbNextNumber
            
            
            HandleNextNumberClick
        
        Case kTbMemo
            '-- The Memo button was pressed by the user
            'CMMemoSelected lkuVendID

        Case kTbHelp
            gDisplayFormLevelHelp Me
        
        Case kTbCopyFrom
            '-- this is a future enhancement to Data Manager
            'moDmHeader.CopyFrom

        Case kTbRenameId
            moDmHeader.RenameID
            
        Case kTbFilter
            miFilter = giToggleLookupFilter(miFilter)
            
        '-- Trap the browse control buttons
        Case kTbFirst, kTbPrevious, kTbLast, kTbNext
            HandleBrowseClick sKey
                        
        Case Else
            tbrMain.GenericHandler sKey, Me, moDmHeader, moClass
    End Select
End Sub

Private Sub HandleNextNumberClick()
    Dim sNextNo As String
    Dim sNum As String
    Dim nNum2 As Long
    
    On Error GoTo ErrorHandler
    
    If Not bConfirmUnload(0) Then Exit Sub
    
    
    moDmHeader.Clear True
    ClearDetlFields
    
    If sddType.ListIndex = -1 Then
        sddType.ListIndex = 0
    End If
    
    '-- Clear the form
'    ProcessCancel
    
    
     Dim tiRetVal As Integer
    '-- Set the detail controls as valid so the old values are correct
    With moClass.moAppDB
       
        .SetOutParam tiRetVal
        .ExecuteSP ("spctActualizarUltimoNoContrato")
        'tiRetVal = .GetOutParam(3)
        'tsContractNO = .GetOutParam(3)
        .ReleaseParams
    End With
    
'
    If GetNextContractNo Then
        sNextNo = msContractNo
        If (Len(Trim$(sNextNo)) > 0) Then
            txtContract.Text = sNextNo
            txtContratokey1.Text = moDmHeader.GetColumnValue("ContractKey")

            '-- Fire the KeyChange event now
            If IsValidContract Then
                valMgr_KeyChange
            End If

'            valMgr_KeyChange
        End If

     If sddType.ItemData = 2 Then
       GetNextContractNo1
'        If sNum <> "" Then
'          txtContractNo.Text = sNum
'          bNum001 = True
'          End If
      ElseIf sddType.ItemData = 1 Then
      
      GetNextContractNo1
'        sNum = gsGetValidStr(moAppDB.Lookup("REPLACE(nextic,SUBSTRING(Nextic, CHARINDEX('USI',Nextic)+3,4), RIGHT('0000' + CONVERT(VARCHAR, CAST(SUBSTRING(Nextic, CHARINDEX('USI',Nextic)+3,4) AS INT)),4))", "tctContractGConfig whith (NOLOCK)", "1 = 1"))
'        moAppDB.ExecuteSQL "UPDATE tctcontractgConfig SET nextic = (SELECT REPLACE(nextic,SUBSTRING(Nextic, CHARINDEX('USI',Nextic)+3,4), RIGHT('0000' + CONVERT(VARCHAR, CAST(SUBSTRING(Nextic, CHARINDEX('USI',Nextic)+3,4) AS INT)+1),4)) FROM tctcontractgConfig whith (NOLOCK))"
'
'         If sNum <> "" Then
'          txtContractNo.Text = sNum
'          bNum001 = True
'          End If
     End If


'     If sddType.ItemData = 2 Then
'      ElseIf sddType.ItemData = 1 Then
'      End If
'      bNum001 = False


    End If
    Exit Sub
ErrorHandler:
    MsgBox "HandleNextNumberClick()_" & Err.Description, vbCritical, "Error"
End Sub

Public Sub HandleBrowseClick(sKey As String)

    Dim lRet                  As Long
    Dim sNewKey          As String
    Dim iConfirmUnload  As Integer

    Select Case sKey
        Case kTbFilter
            miFilter = giToggleLookupFilter(miFilter)

        Case kTbFirst, kTbPrevious, kTbLast, kTbNext

            bConfirmUnload iConfirmUnload, True
            If Not iConfirmUnload = kDmSuccess Then Exit Sub

                lRet = glLookupBrowse(lkuNav, sKey, miFilter, sNewKey)
           
                Select Case lRet

                    Case MS_SUCCESS

                        If lkuNav.ReturnColumnValues.Count = 0 Then Exit Sub

                        '!!! TO DO: Replace ReturnColumnValues Key "CustID" with ColumnName to be placed into the control.
                        If StrComp(Trim(txtContract.Text), Trim(lkuNav.ReturnColumnValues("ContractID")), vbTextCompare) <> 0 Then
             
                            txtContract.Text = Trim(lkuNav.ReturnColumnValues("ContractID"))
                            If IsValidContract Then
                                txtContract.Tag = txtContract.Text
                                valMgr_KeyChange
                            End If
                        End If
           
                    Case Else

                        gLookupBrowseError lRet, Me, moClass

                End Select

        End Select

End Sub

Private Sub ProcessCancel()

  Dim tiRetVal As Integer
    '-- Set the detail controls as valid so the old values are correct
    With moClass.moAppDB
       
        .SetOutParam tiRetVal
        .ExecuteSP ("spctActualizarUltimoNoContrato")
        'tiRetVal = .GetOutParam(3)
        'tsContractNO = .GetOutParam(3)
        .ReleaseParams
    End With
    
    moDmHeader.Clear True
    valMgr.Reset
    LoadOperations
    
    ClearDetlFields
    
     
    

End Sub

Private Function bConfirmUnload(iConfirmUnload As Integer, Optional ByVal bNoClear As Boolean = False) As Boolean
    Dim bValid As Boolean
    
    bConfirmUnload = False
    
    If Not valMgr.ValidateForm Then Exit Function
    
    iConfirmUnload = moDmHeader.ConfirmUnload(bNoClear)
    
    If (iConfirmUnload = kDmSuccess) Then
        bConfirmUnload = True
    End If
End Function

Public Function CMAppendContextMenu(ctl As Control, hmenu As Long) As Boolean
'************************************************************************************
'      Desc: Called to append menu items to right mouse popup menu.
'     Parms: Ctl:   The control that received the right mouse button down message.
'            hmenu: The handle to the popup menu
'   Returns: True if menu item added; False if menu item not added.
'************************************************************************************
    Dim sMenuText As String

    CMAppendContextMenu = False

'-- This will need to be specific to your application

End Function

Public Function CMMenuSelected(ctl As Control, lTaskID As Long) As Boolean
'************************************************************************************
'      Desc: Called when a popup context menu item is selected.
'            Called because menu item was added by CMAppendContextMenu event.
'     Parms: Ctl:   The control that received the right mouse button down message.
'            lTaskID: The Task ID of the selected menu item.
'   Returns: True if successful; False if unsuccessful.
'************************************************************************************
On Error GoTo ExpectedErrorRoutine

    Dim lVoucherKey As Long
    Dim iPos As Integer
    Dim oDDN As Object
    Dim sClassID As String
    Dim lRealTaskID As Long
    
    CMMenuSelected = False

'-- This will need to be specific to your application
   
    CMMenuSelected = True

    Exit Function

ExpectedErrorRoutine:
    SetHourglass False
    
    If oDDN Is Nothing Then
        iPos = InStr(sClassID, ".")
        giSotaMsgBox Nothing, moClass.moSysSession, kmsgDDNNotRegistered, _
            Left$(sClassID, iPos) & "dll"
    
        gClearSotaErr
        Exit Function
    End If
    
End Function


Public Sub CMMemoSelected(oCtl As Control)
'-- This will need to be specific to your application


End Sub

Private Sub cmdBankInfo_Click()

    Dim nContractKey As Long
    Dim sTitular As String
    Dim nAccountID As String
    Dim sSWIFT As String
    Dim sBankAddress As String
    Dim nContractBankInfKey As Long
    Dim lParentContractKey As Long
   Dim bLiberado As Long
   
    
    On Error GoTo ErrorHandler
    
    lParentContractKey = moDmHeader.GetColumnValue("ContractKey")
    'lParentContractKey = glGetValidLong(lkuParentContract.KeyValue)
    
    'If Not moDmHeader.IsDirty() And lParentContractKey <> 0 Then
       ' If frmContractBank.modmForm.KeyChange = kDmKeyNotFound Then
       
        If sddType.ItemData = 3 And lkuParentContract.Text <> "" Then
            
            nContractKey = glGetValidLong(moAppDB.Lookup("ContractKey", "tctContract", "ContractNo =" & gsQuoted(lkuParentContract.Text)))
            If nContractKey <> 0 Then
                
'                If lkuParentContract.Text = "" Then
                
                sTitular = gsGetValidStr(moAppDB.Lookup("Titular", "tctContract", "ContractKey =" & nContractKey))
                nAccountID = gsGetValidStr(moAppDB.Lookup("AccountID", "tctContract", "ContractKey =" & nContractKey))
                sSWIFT = gsGetValidStr(moAppDB.Lookup("SWIFT", "tctContract", "ContractKey =" & nContractKey))
                sBankAddress = gsGetValidStr(moAppDB.Lookup("BankAddress", "tctContract", "ContractKey =" & nContractKey))

'                End If
'
            End If
            
        End If
       
       
          If Len(Trim$(txtTitular.Text)) = 0 And Len(Trim$(txtAccountID.Text)) = 0 And Len(Trim$(txtSWIFT.Text)) = 0 And Len(Trim$(txtBankAddress.Text)) = 0 Then
              txtTitular.Text = sTitular
              txtAccountID.Text = nAccountID
              txtSWIFT.Text = sSWIFT
              txtBankAddress.Text = sBankAddress
          End If
         
         frmContractBank.txtTitular.Text = txtTitular.Text
         frmContractBank.txtAccountNo.Text = txtAccountID.Text
         frmContractBank.txtSWIFT.Text = txtSWIFT.Text
         frmContractBank.txtBankAddress.Text = txtBankAddress.Text
        
         bLiberado = glGetValidLong(moAppDB.Lookup("Free", "tctContract", "ContractKey = " & mlContractKey))
         
         If bLiberado = 1 Then
            frmContractBank.txtTitular.Enabled = False
            frmContractBank.txtAccountNo.Enabled = False
            frmContractBank.txtSWIFT.Enabled = False
            frmContractBank.txtBankAddress.Enabled = False
            
           
            
            
            Else
            
            frmContractBank.txtTitular.Enabled = True
            frmContractBank.txtAccountNo.Enabled = True
            frmContractBank.txtSWIFT.Enabled = True
            frmContractBank.txtBankAddress.Enabled = True
         End If
   
        
      '  frmContractBank.setupComponents msCompanyID
       ' frmContractBank.modmForm1.SetColumnValue "ContractKey", moDmHeader.GetColumnValue("ContractKey")
        'If frmContractChg.modmForm.KeyChange = kDmKeyNotFound Then
           ' lVendClassKey = glGetValidLong(moAppDB.Lookup("p.VendClassKey", "tctContract as p join tapVendClass as s on p.VendClassKey = s.VendClassKey", "p.ContractKey =" & lParentContractKey))
            'sVendClass = gsGetValidStr(moAppDB.Lookup("VendClassID", "tctContract as p join tapVendClass as s on p.VendClassKey = s.VendClassKey", "p.ContractKey =" & lParentContractKey))
            
           ' lContactKey = glGetValidLong(moAppDB.Lookup("p.CntctKey", "tctContract as p join tciContact as s on p.CntctKey = s.CntctKey", "p.ContractKey =" & lParentContractKey))
            'sContactID = gsGetValidStr(moAppDB.Lookup("s.Name", "tctContract as p join tciContact as s on p.CntctKey = s.CntctKey", "p.ContractKey =" & lParentContractKey))
            
           ' lPaymentTermsKey = glGetValidLong(moAppDB.Lookup("p.PmtTermsKey", "tctContract as p join tciPaymentTerms as s on p.PmtTermsKey = s.PmtTermsKey", "p.ContractKey =" & lParentContractKey))
            
           ' lFOBKey = glGetValidLong(moAppDB.Lookup("s.FOBKey", "tctContract as p join tciFOB as s on p.FOBKey = s.FOBKey", "p.ContractKey =" & lParentContractKey))
           ' dDuration = gdGetValidDbl(moAppDB.Lookup("Duration", "tctContract", "ContractKey =" & lParentContractKey))
            
          '  tStartDate = CDate(moAppDB.Lookup("StartDate", "tctContract", "ContractKey =" & lParentContractKey))
          '  tSignatureDate = CDate(moAppDB.Lookup("SignatureDate", "tctContract", "ContractKey =" & lParentContractKey))
            
           ' frmContractChg.lkuVendorClass.Text = sVendClass
           ' frmContractChg.lkuVendorClass.KeyValue = lVendClassKey
            'frmContractChg.lkuContact.SetTextAndKeyValue sContactID, lContactKey
            
           ' frmContractChg.sddPaymentTerms.ItemData = lPaymentTermsKey
           ' frmContractChg.sddFOB.ItemData = lFOBKey
            'frmContractChg.nbrDuration = dDuration
           ' frmContractChg.dtpSignatureDate.Value = tSignatureDate
           ' frmContractChg.dtpStartDate.Value = tSignatureDate
       ' End If
        frmContractBank.Show vbModal, Me
        
        moDmHeader.SetDirty True
        
    'Else
      '  MsgBox "La Información del Banco no tiene cambios Pendientes", vbInformation, "Alerta"
   ' End If
    Exit Sub
ErrorHandler:
    MsgBox "cmdGenParams_Click_" & Err.Description, vbInformation, "Error"
    
End Sub

Private Sub cmdGenParams_Click()
    Dim lPaymentTermsKey As Long
    Dim lFOBKey As Long
    Dim sVendClass As String
    Dim lVendClassKey As Long
    Dim sContactID As String
    Dim lContactKey As Long
    Dim lParentContractKey As Long
    Dim dDuration As Double
    Dim tStartDate As Date
    Dim tSignatureDate As Date
    
    Dim bLiberado As Long
      
      
    On Error GoTo ErrorHandler
    
    lParentContractKey = glGetValidLong(lkuParentContract.KeyValue)
    
    
    
    If Not moDmHeader.IsDirty() And lParentContractKey <> 0 Then
        frmContractChg.setupComponents msCompanyID
        frmContractChg.modmForm.SetColumnValue "ContractKey", moDmHeader.GetColumnValue("ContractKey")
        If frmContractChg.modmForm.KeyChange = kDmKeyNotFound Then
            lVendClassKey = glGetValidLong(moAppDB.Lookup("p.VendClassKey", "tctContract as p join tapVendClass as s on p.VendClassKey = s.VendClassKey", "p.ContractKey =" & lParentContractKey))
            sVendClass = gsGetValidStr(moAppDB.Lookup("VendClassID", "tctContract as p join tapVendClass as s on p.VendClassKey = s.VendClassKey", "p.ContractKey =" & lParentContractKey))
            
            lContactKey = glGetValidLong(moAppDB.Lookup("p.CntctKey", "tctContract as p join tciContact as s on p.CntctKey = s.CntctKey", "p.ContractKey =" & lParentContractKey))
            sContactID = gsGetValidStr(moAppDB.Lookup("s.Name", "tctContract as p join tciContact as s on p.CntctKey = s.CntctKey", "p.ContractKey =" & lParentContractKey))
            
            lPaymentTermsKey = glGetValidLong(moAppDB.Lookup("p.PmtTermsKey", "tctContract as p join tciPaymentTerms as s on p.PmtTermsKey = s.PmtTermsKey", "p.ContractKey =" & lParentContractKey))
            
            lFOBKey = glGetValidLong(moAppDB.Lookup("s.FOBKey", "tctContract as p join tciFOB as s on p.FOBKey = s.FOBKey", "p.ContractKey =" & lParentContractKey))
            dDuration = gdGetValidDbl(moAppDB.Lookup("Duration", "tctContract", "ContractKey =" & lParentContractKey))
            
            tStartDate = CDate(moAppDB.Lookup("StartDate", "tctContract", "ContractKey =" & lParentContractKey))
            tSignatureDate = CDate(moAppDB.Lookup("SignatureDate", "tctContract", "ContractKey =" & lParentContractKey))
            
            If sVendClass <> "" Then frmContractChg.lkuVendorClass.Text = sVendClass
            
            If lVendClassKey <> 0 Then frmContractChg.lkuVendorClass.KeyValue = lVendClassKey
            If lContactKey <> 0 Then frmContractChg.lkuContact.SetTextAndKeyValue sContactID, lContactKey
            
            If lPaymentTermsKey <> 0 Then frmContractChg.sddPaymentTerms.ItemData = lPaymentTermsKey
            If lFOBKey <> 0 Then frmContractChg.sddFOB.ItemData = lFOBKey
            If dDuration <> 0 Then frmContractChg.nbrDuration = dDuration
            If tSignatureDate <> Null Then frmContractChg.dtpSignatureDate.Value = tSignatureDate
            If tSignatureDate <> Null Then frmContractChg.dtpStartDate.Value = tSignatureDate
        End If
        
         bLiberado = glGetValidLong(moAppDB.Lookup("Free", "tctContract", "ContractKey = " & mlContractKey))
         
         If bLiberado = 1 Then
            frmContractChg.lkuVendorClass.Enabled = False
            frmContractChg.lkuContact.Enabled = False
            frmContractChg.sddPaymentTerms.Enabled = False
            frmContractChg.sddFOB.Enabled = False
            frmContractChg.nbrDuration.Enabled = False
            frmContractChg.txtCmnt.Enabled = False
           
            
            
            Else
            
           frmContractChg.lkuVendorClass.Enabled = True
            frmContractChg.lkuContact.Enabled = True
            frmContractChg.sddPaymentTerms.Enabled = True
            frmContractChg.sddFOB.Enabled = True
            frmContractChg.nbrDuration.Enabled = True
            frmContractChg.txtCmnt.Enabled = True
            
         End If
        
        frmContractChg.Show vbModal, Me
    Else
        MsgBox "El suplemento no puede tener cambios Pendientes", vbInformation, "Alerta"
    End If
    Exit Sub
ErrorHandler:
    MsgBox "cmdGenParams_Click_" & Err.Description, vbInformation, "Error"
End Sub

Private Sub cmdOK_Click()
    moLE.GridEditOk
End Sub

Private Sub cmdUndo_Click()
    moLE.GridEditUndo
End Sub

Private Function bSetupNumerics() As Boolean
'*************************************************************************
' This routine will properly format all the currency controls on the form.
'*************************************************************************
    Dim bValid As Boolean
    Dim iIntegralPlaces As Integer
    
    bSetupNumerics = False

    '-- Set attributes for home currency controls
    If (Len(Trim$(msHomeCurrID)) > 0) Then
        SetHomeCurrCtrls msHomeCurrID
    Else
        '-- Apparently, the home currency hasn't been set
        giSotaMsgBox Nothing, moSysSession, kmsgSetCurrControlsError, _
            gsBuildString(kNone, moAppDB, moSysSession)
        Exit Function
    End If

    '-- Set attributes for natural currency controls
    If (Len(Trim$(msNatCurrID)) > 0) Then
        SetNatCurrCtrls msNatCurrID, False
    Else
        '-- Apparently, the natural currency hasn't been set
        giSotaMsgBox Nothing, moSysSession, kmsgSetCurrControlsError, _
            gsBuildString(kNone, moAppDB, moSysSession)
        Exit Function
    End If

    '-- Set properties for number fields
    'nbrQuantity.DecimalPlaces = moOptions.CI("QtyDecPlaces")
    
    '-- Setting the integral places at run-time sets it back to 12!! We only want 8.
    'iIntegralPlaces = IIf(kiQtyLen - moOptions.CI("QtyDecPlaces") > kiQtyIntegralPlaces, _
    '    kiQtyIntegralPlaces, kiQtyLen - moOptions.CI("QtyDecPlaces"))
    'nbrQuantity.IntegralPlaces = iIntegralPlaces
    
    bSetupNumerics = True
End Function


Private Sub SetHomeCurrCtrls(sCurrID As String)
    Dim bValid As Boolean
    Dim iIntegralPlaces As Integer

'-- The HC controls on the form need to be identified here
    

    '-- Set attributes for home currency controls
    bValid = gbSetCurrCtls(moClass, sCurrID, muHomeCurrInfo)
    
    If bValid Then
        '-- Set integral places for Decimal(15,3) fields
    End If

    '-- Display the currency descriptions
    
End Sub


Private Sub SetNatCurrCtrls(sCurrID As String, Optional bRefresh As Boolean = True)
    Dim bValid As Boolean
    Dim iIntegralPlaces As Integer
    
'-- The NC controls on the form need to be identified here
    
    '-- Set attributes for natural currency controls
    
    If bValid Then
        '-- Set integral places for Decimal(15,3) fields

    End If
End Sub

Private Sub InitializeDetlGrid()
'*******************************************
' Desc: Initializes/formats the Detail grid.
'*******************************************
    Dim iGridType As Integer
    Dim sTitle As String


    gGridSetProperties grdMain, kMaxCol, kGridLineEntry
    gGridSetColors grdMain
    gGridSetMaxCols grdMain, kMaxCol
    '-- Set default grid properties
    'Headers
    gGridSetHeader grdMain, kColSeqNo, "No"
    gGridSetHeader grdMain, kColItemID, "Artículo"
    gGridSetHeader grdMain, kColDescription, "Descripción"
    gGridSetHeader grdMain, kColUnitMeasID, "UM"
    gGridSetHeader grdMain, kColUnitCost, "Costo Unitario"
    gGridSetHeader grdMain, kColItemQty, "Cantidad"
    gGridSetHeader grdMain, kColLineAmt, "Importe"
    gGridSetHeader grdMain, kColMaxLot, "Lote Maximo"
    gGridSetHeader grdMain, kColMinLot, "Lote Mínimo"
    gGridSetHeader grdMain, kColDeliveryTime, "Tiempo de Entrega"
    gGridSetHeader grdMain, kColRoundValue, "Valor de Redondeo"
    
    gGridSetHeader grdOperations, kColOpertPoNo, "OC"
    gGridSetHeader grdOperations, kColOpertInvNo, "Factura"
    gGridSetHeader grdOperations, kColOpertDate, "Fecha Aplicación"
    gGridSetHeader grdOperations, kColOpertAmt, "Importe"
    gGridSetHeader grdOperations, kColOpertQty, "Cantidad"
    gGridSetHeader grdOperations, kColOpertDesc, "Descripción"

    'Visivility
    gGridHideColumn grdMain, kColContractKey
    gGridHideColumn grdMain, kColContractLineKey
    gGridHideColumn grdMain, kColItemKey
    gGridHideColumn grdMain, kColUnitMeasKey
    gGridHideColumn grdMain, kColCreateDate
    gGridHideColumn grdMain, kColUpdateDate
    gGridHideColumn grdMain, kColCreateUser
    gGridHideColumn grdMain, kColUpdateUser
    gGridHideColumn grdMain, kColType
    gGridHideColumn grdMain, kColQtyVariation
    
    gGridHideColumn grdOperations, kColOpertInvLineKey
    
    'Types
    
    
    
'    gGridSetColumnType grdMain, kColChildContractKey, SS_CELL_TYPE_EDIT
'    gGridSetColumnType grdMain, kColChildCreateDate, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildCreateDate, "Create Date"
'    gGridSetColumnWidth grdMain, kColChildCreateDate, "10"
'    gGridSetColumnType grdMain, kColChildCreateUser, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildCreateUser, "Create User"
'    gGridSetColumnWidth grdMain, kColChildCreateUser, "50"
'    gGridSetColumnType grdMain, kColChildDeliveryTime, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildDeliveryTime, "Delivery Time"
'    gGridSetColumnWidth grdMain, kColChildDeliveryTime, "4"
'    gGridSetColumnType grdMain, kColChildDescription, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildDescription, "Description"
'    gGridSetColumnWidth grdMain, kColChildDescription, "50"
'    gGridSetColumnType grdMain, kColChildItemKey, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildItemKey, "Item Key"
'    gGridSetColumnWidth grdMain, kColChildItemKey, "4"
'    gGridSetColumnType grdMain, kColChildMaxLot, SS_CELL_TYPE_FLOAT, 2
'    gGridSetHeader grdMain, kColChildMaxLot, "Max Lot"
'    gGridSetColumnWidth grdMain, kColChildMaxLot, "19"
'    gGridSetColumnType grdMain, kColChildMinLot, SS_CELL_TYPE_FLOAT, 2
'    gGridSetHeader grdMain, kColChildMinLot, "Min Lot"
'    gGridSetColumnWidth grdMain, kColChildMinLot, "19"
'    gGridSetColumnType grdMain, kColChildQty, SS_CELL_TYPE_FLOAT, 2
'    gGridSetHeader grdMain, kColChildQty, "Qty"
'    gGridSetColumnWidth grdMain, kColChildQty, "19"
'    gGridSetColumnType grdMain, kColChildRoundValue, SS_CELL_TYPE_FLOAT, 2
'    gGridSetHeader grdMain, kColChildRoundValue, "Round Value"
'    gGridSetColumnWidth grdMain, kColChildRoundValue, "19"
'    gGridSetColumnType grdMain, kColChildSeqNo, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildSeqNo, "Seq No"
'    gGridSetColumnWidth grdMain, kColChildSeqNo, "4"
'    gGridSetColumnType grdMain, kColChildUnitCost, SS_CELL_TYPE_FLOAT, 2
'    gGridSetHeader grdMain, kColChildUnitCost, "Unit Cost"
'    gGridSetColumnWidth grdMain, kColChildUnitCost, "19"
'    gGridSetColumnType grdMain, kColChildUnitMeasKey, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildUnitMeasKey, "Unit Measure Key"
'    gGridSetColumnWidth grdMain, kColChildUnitMeasKey, "4"
'    gGridSetColumnType grdMain, kColChildUpdateDate, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildUpdateDate, "Update Date"
'    gGridSetColumnWidth grdMain, kColChildUpdateDate, "10"
'    gGridSetColumnType grdMain, kColChildUpdateUser, SS_CELL_TYPE_EDIT
'    gGridSetHeader grdMain, kColChildUpdateUser, "Update User"
'    gGridSetColumnWidth grdMain, kColChildUpdateUser, "50"

    gGridSetColumnType grdOperations, kColOpertDate, SS_CELL_TYPE_DATE

    gGridSetColumnWidth grdOperations, kColOpertPoNo, 15
    gGridSetColumnWidth grdOperations, kColOpertDesc, 25

    gGridLockColumn grdOperations, kColOpertInvNo
    gGridLockColumn grdOperations, kColOpertPoNo
    gGridLockColumn grdOperations, kColOpertAmt
    gGridLockColumn grdOperations, kColOpertDesc
    gGridLockColumn grdOperations, kColOpertDate
    gGridLockColumn grdOperations, kColOpertQty

    gGridSetMaxCols grdOperations, 7
    

    '-- Setup the grid's Currency columns based on the Natural Currency ID
    bSetupGridNumerics
End Sub

Private Function bSetupGridNumerics() As Boolean
    bSetupGridNumerics = True
    
'-- The numeric columns in the grid need to be identified here
    

    '-- Setup the grid's Currency columns based on the Natural Currency ID

End Function

Public Function CadenaSinCeros(sEntrada As String) As String
 Dim i As Integer, j As Integer, sIntermedio As String
 For j = 1 To Len(sEntrada)
    If Mid(sEntrada, j, 1) <> "0" Then
     sIntermedio = Mid(sEntrada, j)
    Exit For
    End If
 Next
 CadenaSinCeros = sIntermedio
End Function

Private Sub currAmt_Change()
 
 
'If currAmt.Text <> "" Then
'     If CDec(currAmt) > 1E+16 Then
'       dd = CDec(currAmt)
'       dd = dd / 1E+16
'        dd2 = dd - CLng(dd)
'        dd2 = dd2 * 1E+16
'
'        dd = CLng(dd)
'    '   dd = CDbl(currAmt) - (CDbl(currAmt) / 1E+16)
'        If (dd2 < 0) Then
'         dd2 = dd2 * (-1)
'        End If
'
'
'       cc = CStr(FormatNumber(dd, 0, , vbFalse, vbFalse))
'       cc2 = CStr(FormatNumber(dd2, 2, , vbFalse, vbFalse))
'      currAmt.Text = cc + cc2
'     ' currAmt.Text = CadenaSinCeros(Format(currAmt, "###############################000000000000000000000000000000.0000"))
'
' Else
'
' End If
'End If

End Sub

Private Sub currAmt_Validate(Cancel As Boolean)
   Dim dd As Double
 Dim dd2 As Double
 
 Dim cc As String
 Dim cc2 As String
   
   If currAmt.Text <> "" Then
     If CDbl(currAmt) > 100000000000000# Then
       dd = CDbl(currAmt)
       dd = dd / 100000000000000#
        dd2 = dd - FormatNumber(dd, 0, vbFalse, vbFalse)
        dd2 = dd2 * 100000000000000#
        
        dd = FormatNumber(dd, 0, vbFalse, vbFalse)
    '   dd = CDbl(currAmt) - (CDbl(currAmt) / 1E+16)
        If (dd2 < 0) Then
         dd2 = dd2 * (-1)
        End If
        
        
       cc = CStr(FormatNumber(dd, 0, , vbFalse, vbFalse))
       cc2 = CStr(FormatNumber(dd2, 2, , vbFalse, vbFalse))
      currAmt.Text = cc + cc2
     ' currAmt.Text = CadenaSinCeros(Format(currAmt, "###############################000000000000000000000000000000.0000"))
      
     Else
        If (currAmt < 0) Then
        currAmt.Text = "0.00"
        End If
     End If
    End If
End Sub

Private Sub currItemAmt_Change()
'  moLE.GridEditChange currItemAmt
End Sub

Private Sub Form_Initialize()
    mbCancelShutDown = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF16
            gProcessFKeys Me, KeyCode, Shift
         
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            If mbEnterAsTab Then
               gProcessSendKeys "{TAB}"
               KeyAscii = 0
            End If
    
    End Select
End Sub

Private Sub Form_Load()
    mbLoadSuccess = False
    
    Set frmContractChg.oClass = moClass
    
    Set moSysSession = moClass.moSysSession
    With moSysSession
        msCompanyID = .CompanyId
        msUserID = .UserId
        mbEnterAsTab = .EnterAsTab
        mlLanguage = .Language
        msHomeCurrID = .CurrencyID
        msBusDate = .BusinessDate
    End With
    Set moAppDB = moClass.moAppDB
    msLookupRestrict = "CompanyID = " & gsQuoted(msCompanyID)

    With Me
        miMinFormHeight = .Height
        miMinFormWidth = .Width
        miOldFormHeight = .Height
        miOldFormWidth = .Width
    End With

    Me.Caption = "Gestionar Contratos"
        
     With moOptions
        Set .oSysSession = moSysSession
        Set .oAppDB = moAppDB
        .sCompanyID = msCompanyID
    End With
    
     SetupModuleVars
    
    '-- Setup currency controls
    If Not bSetupNumerics() Then Exit Sub
    
     SetupBars
    
     SetupLookups
    
     SetupDropDowns
    
     BindForm
    
     miSecurityLevel = giSetAppSecurity(moClass, tbrMain, moDmHeader, moDmDetl)
            
     BindLE
    
     InitializeDetlGrid
    
     BindContextMenu
        
     'SetFieldStates
    
     'SetTBButtonStates
    
    '-- Make sure the header tab is shown to the user first
    'tabVoucher.Tab = kiHeaderTab
    
    tabDataEntry.Tab = kiHeaderTab
    
  '-- Disable controls on hidden tabs
    Dim i As Integer
   pnlTab(kiHeaderTab).Enabled = True
   For i = 0 To pnlTab.Count - 1
        pnlTab(i).Enabled = False
   Next i
 
    '-- Grid starts out with 500 rows.Remove the rows then add the header
    grdMain.MaxRows = 0
    grdMain.MaxRows = 1
    
    bNum001 = False
    
    
    
    '-- We made it through Form_Load successfully
    mbLoadSuccess = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        If Not moFormCust.CanShutdown Then
            Cancel = True
            Exit Sub
        End If
    End If
#End If
    Dim iConfirmUnload As Integer
    Dim bValid As Boolean
   
    '-- Reset the CancelShutDown flag if prior shutdowns were canceled.
    mbCancelShutDown = False

    If (moClass.mlError = 0) Then
        '-- If the form is dirty, prompt the user to save the record
         bValid = bConfirmUnload(iConfirmUnload)
        
        Select Case iConfirmUnload
            Case kDmSuccess
                'Do Nothing
                
            Case kDmFailure, kDmError
                GoTo CancelShutDown
        
        End Select
      
        '-- Clear the form
        ProcessCancel
        
        '-- Check all other forms that may have been loaded from this main form.
        '-- If there are any Visible forms, then this means the form is Active.
        '-- Therefore, cancel the shutdown.
        If gbActiveChildForms(Me) Then GoTo CancelShutDown
    
        Select Case UnloadMode
            Case vbFormCode
                'Do Nothing
            
            Case Else
                '-- Most likely the user has requested to shut down the form.
                '-- If the context is normal or Drill-Around, have the object unload itself.
                moClass.miShutDownRequester = kUnloadSelfShutDown
        End Select
    End If
    
    '-- If execution gets to this point, the form and class object of the form
    '-- will be shut down. Perform all operations necessary for a clean shutdown.


    PerformCleanShutDown
        
    Select Case moClass.miShutDownRequester
        Case kUnloadSelfShutDown
            moClass.moFramework.UnloadSelf EFW_TF_MANSHUTDN
            Set moClass.moFramework = Nothing
            
        Case Else
            'Do Nothing
            
    End Select

    Exit Sub
    
CancelShutDown:
    moClass.miShutDownRequester = kFrameworkShutDown
    mbCancelShutDown = True
    Cancel = True
    Exit Sub
End Sub

Public Sub PerformCleanShutDown()

#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        moFormCust.UnloadSelf
        Set moFormCust = Nothing
    End If
#End If

'-- Unload all forms loaded from this main form
    gUnloadChildForms Me
         
    '-- Remove all child collections
    giCollectionDel moClass.moFramework, moSotaObjects, -1
    
    '-- Clean up object references
    If Not moDmHeader Is Nothing Then
        moDmHeader.UnloadSelf
        Set moDmHeader = Nothing
    End If
    
    If Not moDmDetl Is Nothing Then
        moDmDetl.UnloadSelf
        Set moDmDetl = Nothing
    End If
    

    '-- Clean up Line Entry references
    If Not moLE Is Nothing Then
        moLE.UnloadSelf
        Set moLE = Nothing
    End If
    
    '-- Fire Terminate event for lookups
    TerminateControls Me

    '-- Clean up other objects
    Set moOptions = Nothing
    Set moSotaObjects = Nothing
    Set moContextMenu = Nothing
    Set moAppDB = Nothing
    Set moSysSession = Nothing

End Sub

Private Sub Form_Resize()
    '-- Resize the grid and associated controls
    If (Me.WindowState = vbNormal Or Me.WindowState = vbMaximized) Then
    '    '-- Move the controls down
        gResizeForm kResizeDown, Me, miOldFormHeight, miMinFormHeight, _
                        pnlTab(kiHeaderTab), pnlTab(kiDetailTab), pnlTab(kiTotalsTab), _
                        grdMain, tabDataEntry

        '-- Move the controls to the right
        gResizeForm kResizeRight, Me, miOldFormWidth, miMinFormWidth, _
                        pnlTab(kiHeaderTab), pnlTab(kiDetailTab), pnlTab(kiTotalsTab), _
                        grdMain, tabDataEntry
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
 '   Set moClass = Nothing
End Sub

Public Function LECheckRequiredData(oLE As clsLineEntry) As Boolean
'***********************************************************************
' Desc: Soft touch validation in case hard validation at the
'       source control was click-bypassed.
' Parameters: oLE - the Line Entry class object calling this procedure.
'***********************************************************************
    LECheckRequiredData = False

'-- Validate data from Line Entry here before pushing it into the grid
    If Len(Trim$(lkuItem.Text)) = 0 Then
        MsgBox "Debe seleccionar un artículo.", vbInformation, "Alerta"
        Exit Function
    End If
    
    If Len(Trim$(txtDescription.Text)) = 0 Then
        MsgBox "Debe introducir una Descripción válida.", vbInformation, "Alerta"
        Exit Function
    End If
    
    If sddUM.ListIndex = -1 Then
        MsgBox "Debe seleccionar una UM.", vbInformation, "Alerta"
        Exit Function
    End If
    
    If sddLineType.ListIndex = -1 Then
        MsgBox "Debe seleccionar un Tipo de Partida.", vbInformation, "Alerta"
        Exit Function
    End If
    
    If currUnitCost.Amount <= 0 And sddLineType.ItemData <> kLineTypeModCos Then
        MsgBox "Debe introducir un costo unitario mayor que 0.", vbInformation, "Alerta"
        Exit Function
    End If
    
    If (sddLineType.Text <> "Eliminar" And sddLineType.Text <> "Modificar Costo") And (numItemQty.Value <= 0 And nbrQtyVariation.Value <= 0) Then
'        If sddLineType.ItemData <> 2 And (numItemQty.Value <= 0 Or nbrQtyVariation.Value <= 0) Then
            MsgBox "Debe introducir una cantidad mayor que 0.", vbInformation, "Alerta"
            Exit Function
'        End If
'    Else
'        If (numItemQty.Value <= 0 Or nbrQtyVariation.Value <= 0) And sddLineType.ItemData <> kLineTypeModCos Then
'            MsgBox "Debe introducir una cantidad mayor que 0.", vbInformation, "Alerta"
'            Exit Function
'        End If
    End If
    
      If lkuItem.KeyValue <> 0 And numItemQty.Value < numMinLot.Value And numMinLot.Value <> 0 And (numItemQty.Value <> 0 Or nbrQtyVariation.Value <> 0) And sddLineType.ItemData <> kLineTypeModCos Then

          MsgBox "La cantidad debe ser superior al Lote Mínimo.", vbExclamation, "Alerta"
          bSetFocus nbrQtyVariation
           Exit Function

    End If
    
    LECheckRequiredData = True
End Function

Public Sub LEClearDetailControls(oLE As Object)
'***********************************************************************
' Description: Clear detail fields on the Line Entry object.
' Parameters: oLE - the Line Entry class object calling this procedure.
'***********************************************************************
    'ClearDetlFields True
    CleanDetails
    
End Sub

Public Sub LESetDetailControlDefaults(oLE As clsLineEntry)
'************************************************************************************
'   Description:
'           Detail edit controls loading point.
'   Parameter:
'           oLE - the line entry class object calling this procedure.
'************************************************************************************

'-- Set up default values for Line Entry here
'    CleanDetails

    mlActiveRow = moLE.ActiveRow
End Sub

Public Sub LEDetailToGrid(oLE As clsLineEntry)
'***********************************************************************
' Desc: If any detail entry controls need to be manually linked to
'       the grid after entry, then do it here.
' Parameters: oLE - the Line Entry class object calling this procedure.
'***********************************************************************

'-- Update any unbound columns in the grid here
    Dim iSeqNo As Integer
    

    gGridUpdateCellText grdMain, oLE.ActiveRow, kColItemID, lkuItem.Text
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColItemKey, lkuItem.KeyValue
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColDescription, txtDescription.Text
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColUnitCost, currUnitCost.Amount
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColItemQty, numItemQty.Value
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColLineAmt, currItemAmt.Amount
    
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColUnitMeasKey, sddUM.ItemData
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColUnitMeasID, sddUM.Text
    
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColMinLot, numMinLot.Value
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColMaxLot, numMaxLot.Value
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColDeliveryTime, numDeliveryTime.Value
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColRoundValue, numRoundValue.Value
    
    iSeqNo = giGetValidInt(gsGridReadCellText(grdMain, oLE.ActiveRow, kColSeqNo))
    If iSeqNo = 0 Then iSeqNo = GetNextSeqNo
    
    gGridUpdateCellText grdMain, oLE.ActiveRow, kColSeqNo, gsGetValidStr(iSeqNo)
    gGridUpdateCell grdMain, oLE.ActiveRow, kColType, sddLineType.ItemData
    gGridUpdateCell grdMain, oLE.ActiveRow, kColQtyVariation, nbrQtyVariation.Value
    
    If sddLineType.ItemData = kLineTypeModCos Then
     nbrQtyVariation.Enabled = False
     numItemQty.Enabled = False
   Else
      nbrQtyVariation.Enabled = True
     numItemQty.Enabled = True
    End If
    
    
    CalcContractAmt
    
  
    
End Sub

Public Sub LEGridToDetail(oLE As clsLineEntry)
Dim iUMKey As Integer
    On Error GoTo ErrorHandler
'***********************************************************************
' Desc: If any grid columns need to be manually linked to the detail
'       entry controls, then do it here.
' Parameters: oLE - the Line Entry class object calling this procedure.
'***********************************************************************

'-- Set up Line Entry from the grid here
    
    lkuItem.SetTextAndKeyValue gsGridReadCellText(grdMain, oLE.ActiveRow, kColItemID), glGetValidLong(gsGridReadCellText(grdMain, oLE.ActiveRow, kColItemKey))
    txtDescription.Text = gsGridReadCellText(grdMain, oLE.ActiveRow, kColDescription)
    currUnitCost.Amount = gdGetValidDbl(gsGridReadCellText(grdMain, oLE.ActiveRow, kColUnitCost))
    numItemQty.Value = gdGetValidDbl(gsGridReadCellText(grdMain, oLE.ActiveRow, kColItemQty))
    UpdateUnitOfMeas
    iUMKey = giGetValidInt(gsGridReadCellText(grdMain, oLE.ActiveRow, kColUnitMeasKey))
    If iUMKey > 0 And sddUM.ListCount > 0 Then sddUM.ItemData = iUMKey
    numMinLot.Value = gdGetValidDbl(gsGridReadCellText(grdMain, oLE.ActiveRow, kColMinLot))
    numMaxLot.Value = gdGetValidDbl(gsGridReadCellText(grdMain, oLE.ActiveRow, kColMaxLot))
    numDeliveryTime.Value = gdGetValidDbl(gsGridReadCellText(grdMain, oLE.ActiveRow, kColDeliveryTime))
    numRoundValue.Value = gdGetValidDbl(gsGridReadCellText(grdMain, oLE.ActiveRow, kColRoundValue))
    If giGetValidInt(gsGridReadCell(grdMain, oLE.ActiveRow, kColType)) <> 0 Then
        sddLineType.ItemData = giGetValidInt(gsGridReadCell(grdMain, oLE.ActiveRow, kColType))
    Else
        sddLineType.ListIndex = -1
    End If
    nbrQtyVariation.Value = gdGetValidDbl(gsGridReadCell(grdMain, oLE.ActiveRow, kColQtyVariation))
    '-- Set the detail controls as valid so the old values are correct
    
      If lkuItem.Text = "" Then
               lkuItem.Enabled = True
                txtDescription.Enabled = False
                numDeliveryTime.Enabled = True
                numRoundValue.Enabled = True
                sddUM.Enabled = True
                currUnitCost.Enabled = True
                numMinLot.Enabled = True
                numMaxLot.Enabled = True
                numItemQty.Enabled = True
                
                
    End If
    
    
    mlActiveRow = oLE.ActiveRow
    valMgr.Reset
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Public Function LEGridBeforeDelete(oLE As clsLineEntry, lRow As Long) As Boolean
'************************************************************************
'   Description:
'       called from line entry class when a row has been successfully
'       deleted from the grid
'   Param:
'       oLE -   line entry class making the call
'       lRow -  grid row deleted
'************************************************************************
    
    LEGridBeforeDelete = False
    
'-- Perform processing before a line is deleted from the grid here

    Dim lStateKey As Long
    
    
    lStateKey = glGetValidLong(moAppDB.Lookup("ContractStateKey", "tctContractState", "ContractStateId = 'Activo'"))
    If lStateKey <> 0 Then
        If lStateKey = sddState.ItemData Then
            MsgBox "No se pueden eliminar partidas de un contrato activo.", vbExclamation, "Alerta"
            Exit Function
        End If
    End If

    
    LEGridBeforeDelete = True
End Function

Public Sub LEGridAfterDelete(oLE As clsLineEntry, lRow As Long)
'************************************************************************
'   Description:
'       called from line entry class when a row has been successfully
'       deleted from the grid
'   Param:
'       oLE -   line entry class making the call
'       lRow -  grid row deleted
'************************************************************************

'-- Perform processing after a line is deleted from the grid here

End Sub

Public Sub LESetDetailFocus(oLE As clsLineEntry, Optional Col As Variant)
'***********************************************************************
' Desc: Double-clicking on any grid column will drive focus to the
'       detail oControl representing the Selected grid column.
' Parameters: oLE - the Line Entry class object calling this procedure.
'             Col - The column to set the focus into.
'***********************************************************************
    Dim oControl As Object

'-- Link up columns to controls in Line Entry here
End Sub


Private Sub txtContract_LookupClick(bCancel As Boolean)
    bCancel = Not bConfirmUnload(0, True)

End Sub

Private Sub lkuContact_LostFocus()
    If Len(Trim$(lkuContact.Text)) > 0 Then
        If lkuContact.Tag <> lkuContact.Text Then
        End If
    End If
End Sub

Private Sub lkuContact_Validate(Cancel As Boolean)
    If Len(Trim$(lkuContact.Text)) > 0 Then
        If lkuContact.Tag <> lkuContact.Text Then
        End If
    End If
End Sub

Private Sub lkuItem_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    Dim lItemKey As Long
    Dim lParentKey As Long
    Dim lBaseUnit As Long
    Dim i As Long
    Dim bExist As Boolean
    Dim bFromParent As Boolean
    
    If Len(Trim$(lkuItem.Text)) > 0 Then
        If lkuItem.Text <> lkuItem.Tag Then
            lItemKey = glGetValidLong(moAppDB.Lookup("ItemKey", "timItem", "ItemID = " & gsQuoted(lkuItem)))
            bExist = False
            For i = 0 To grdMain.DataRowCnt
                If i <> moLE.ActiveRow Then
                    If glGetValidLong(gsGridReadCell(grdMain, i, kColItemKey)) = lItemKey Then
                        bExist = True
                    End If
                End If
            Next i
            
            If lItemKey > 0 And Not bExist Then
                lParentKey = glGetValidLong(lkuParentContract.KeyValue)
                bFromParent = False
                If lParentKey <> 0 Then
                    If giGetValidInt(moAppDB.Lookup("count(*)", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey)) > 0 Then
                        bFromParent = True
                    End If
                End If
                lkuItem.KeyValue = lItemKey
                If sddType.ItemData = 3 And bFromParent Then
                    currUnitCost.Amount = gdGetValidDbl(moAppDB.Lookup("UnitCost", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey))
                    txtDescription.Text = gsGetValidStr(moAppDB.Lookup("Description", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey))
                    numDeliveryTime.Value = gsGetValidStr(moAppDB.Lookup("DeliveryTime", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey))
                    numRoundValue.Value = gsGetValidStr(moAppDB.Lookup("RoundValue", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey))
                    numMinLot.Value = gsGetValidStr(moAppDB.Lookup("MinLot", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey))
                    numMaxLot.Value = gsGetValidStr(moAppDB.Lookup("MaxLot", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey))
                    lBaseUnit = glGetValidLong(moAppDB.Lookup("UnitMeasKey", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey))
                    sddLineType.ItemData = 3
                Else
                    currUnitCost.Amount = gdGetValidDbl(moAppDB.Lookup("stdUnitCost", "timItem", "ItemKey = " & lItemKey))
                    txtDescription.Text = gsGetValidStr(moAppDB.Lookup("ShortDesc", "timItemDescription", "ItemKey = " & lItemKey))
                    numDeliveryTime.Value = gsGetValidStr(moAppDB.Lookup("isnull(TimeDelivery,0)", "timItemSecurityStock", "ItemKey = " & lItemKey))
                    numRoundValue.Value = gsGetValidStr(moAppDB.Lookup("isnull(RoundQty,0)", "timItemSecurityStock", "ItemKey = " & lItemKey))
                    numMinLot.Value = gsGetValidStr(moAppDB.Lookup("isnull(MinQty,0)", "timItemSecurityStock", "ItemKey = " & lItemKey))
                    lBaseUnit = glGetValidLong(moAppDB.Lookup("s.UnitMeasKey", "timItemUnitOfMeas AS p JOIN tciUnitMeasure AS s ON p.TargetUnitMeasKey = s.UnitMeasKey", "s.Base = 1 and p.ItemKey = " & lItemKey))
                    sddLineType.ItemData = 1
                End If
                UpdateUnitOfMeas
                If lBaseUnit > 0 Then sddUM.ItemData = lBaseUnit
            Else
                If bExist Then
                    MsgBox "El artículo ya existe en el contrato", vbExclamation, "Alerta"
                Else
                    MsgBox "Debe seleccionar un artículo valido", vbExclamation, "Alerta"
                End If
                
                lkuItem.SetTextAndKeyValue "", 0
                txtDescription.Text = ""
                numDeliveryTime.Value = 0
                numRoundValue.Value = 0
                numMinLot.Value = 0
                sddUM.ListIndex(False) = -1
            End If
        End If
    Else
        txtDescription.Text = ""
        currUnitCost.Amount = 0
    End If
    lkuItem.Tag = lkuItem.Text
    
    If lkuItem.Text <> "" And sddLineType.ItemData = kLineTypeAdd Then
              
                numItemQty.Protected = False
                 numItemQty.Visible = True
                 nbrQtyVariation.Visible = False
    End If
    
    Exit Sub
ErrorHandler:
    gSetSotaErr Err, "frmContract", "lkuItem_Validate", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number
        lkuItem.Text = ""
End Sub

Private Sub lkuNav_Click()
    Dim sRestrict As String
    On Error GoTo ErrorHandler
    sRestrict = "CompanyID = " & gsQuoted(msCompanyID)
    If Len(Trim$(lkuVendor.Text)) > 0 Then
        sRestrict = sRestrict & " and VendorKey =" & lkuVendor.KeyValue
    End If
    
    lkuNav.RestrictClause = sRestrict
    
    gcLookupClick Me, lkuNav, txtContract, "ContractID"
    Exit Sub
ErrorHandler:
    MsgBox "lkuNav_Click()_" & Err.Description, vbCritical, "Error"
End Sub

Private Sub lkuParentContract_Change()
 Dim lParentContractKey As Long
    
'      If Len(Trim$(lkuVendor.Text)) > 0 Then
'                lParentContractKey = glGetValidLong(moAppDB.Lookup("p.ContractKey", "tctContract as p join tapVendor as s on s.VendKey = p.VendorKey", "p.ContractNo =" & gsQuoted(lkuParentContract.Text) & " and s.VendID = " & gsQuoted(lkuVendor.Text)))
'            Else
'         End If
'
    
    
    If lkuParentContract.Text <> "" Then
    
             lParentContractKey = glGetValidLong(moAppDB.Lookup("ContractKey", "tctContract", "ContractNo =" & gsQuoted(lkuParentContract.Text)))

            ssTitular = gsGetValidStr(moAppDB.Lookup("p.Titular", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
            ssAccountNo = gsGetValidStr(moAppDB.Lookup("p.AccountID", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
            ssSWIFT = gsGetValidStr(moAppDB.Lookup("p.SWIFT", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
            ssBankAddress = gsGetValidStr(moAppDB.Lookup("p.BankAddress", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
     
           If ssTitular <> "" Then txtTitular.Text = ssTitular
           If ssAccountNo <> "" Then txtAccountID.Text = ssAccountNo
           If ssSWIFT <> "" Then txtSWIFT.Text = ssSWIFT
           If ssBankAddress <> "" Then txtBankAddress.Text = ssBankAddress
     
     
     End If
End Sub

Private Sub lkuParentContract_LostFocus()
   Dim lParentContractKey As Long
   Dim nCoun As Long
   Dim sNumCon As String
   Dim sGuin As String
   
   sGuin = "_"
   If lkuParentContract.Text <> "" Then
     lParentContractKey = glGetValidLong(moAppDB.Lookup("ContractKey", "tctContract", "ContractNo =" & gsQuoted(lkuParentContract.Text)))
     sNumCon = gsGetValidStr(moAppDB.Lookup("ContractNo +" & gsQuoted(sGuin) & "+ CONVERT(VARCHAR(5),(SELECT Count(*)+1 FROM tctContract WHERE ParentContractKey = " & lParentContractKey & "))", "tctContract", "ContractKey =" & lParentContractKey))
     txtContractNo.Text = sNumCon
   End If
End Sub

Private Sub lkuParentContract_Validate(Cancel As Boolean)
Dim lParentContractKey As Long
    On Error GoTo ErrorHandler
    
    If Len(Trim$(lkuParentContract.Text)) > 0 Then
        If lkuParentContract.Tag <> lkuParentContract.Text Then
            If Len(Trim$(lkuVendor.Text)) > 0 Then
                lParentContractKey = glGetValidLong(moAppDB.Lookup("p.ContractKey", "tctContract as p join tapVendor as s on s.VendKey = p.VendorKey", "p.ContractNo =" & gsQuoted(lkuParentContract.Text) & " and s.VendID = " & gsQuoted(lkuVendor.Text)))
            Else
                lParentContractKey = glGetValidLong(moAppDB.Lookup("ContractKey", "tctContract", "ContractNo =" & gsQuoted(lkuParentContract.Text)))
            End If
            If lParentContractKey <> 0 Then
                bLoadParentContratDftl lParentContractKey
                lkuParentContract.SetTextAndKeyValueNoValidate lkuParentContract.Text, lParentContractKey
                lkuVendor.Enabled = False
            Else
                MsgBox "El número de contrato no es valido", vbExclamation, "Alerta"
                lkuParentContract.SetTextAndKeyValue "", 0
                lkuVendor.Enabled = True
            End If
        End If
    Else
        lkuVendor.Enabled = True
        lkuParentContract.SetTextAndKeyValue "", 0
    End If
    lkuParentContract.Tag = lkuParentContract.Text
    lkuParentContract.RestrictClause = sGetParentRestrict
    
    
       If sddType.Text = "Suplemento" Then
                    lkuVendor.Enabled = False
                    sddType.Enabled = False
                  
                    sddPaymentTerms.Enabled = False
                  
                    lkuContact.Enabled = False
                    lkuVendClass.Enabled = False
                    sddFOB.Enabled = False
                    SddCurrID.Enabled = False
                  
                    txtDescription.Enabled = False
                    sddClasification.Enabled = False
                       Else
                   lkuVendor.Enabled = True
                    sddType.Enabled = True
                  
                    lkuParentContract.Enabled = True
'
                    lkuContact.Enabled = True
                    lkuVendClass.Enabled = True
                    sddFOB.Enabled = True
                    SddCurrID.Enabled = True
                    
'
                    txtDescription.Enabled = True
                    sddClasification.Enabled = True
                   End If
    
'
    
    Exit Sub
ErrorHandler:
    gSetSotaErr Err, sMyName, "bLoadParentContratDftl", VBRIG_IS_FORM                    'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number
End Sub

Private Sub lkuVendClass_Validate(Cancel As Boolean)
    Dim lVendClassKey As Long
    If Len(Trim$(lkuVendClass.Text)) > 0 Then
        If lkuVendClass.Tag <> lkuVendClass.Text Then
            lVendClassKey = moAppDB.Lookup("VendClassKey", "tapVendClass", "VendClassID =" & gsQuoted(lkuVendClass.Text))
            If lVendClassKey = 0 Then
                MsgBox "La clase del proveedor no es valida", vbExclamation, "Alerta"
                lkuVendClass.SetTextAndKeyValue "", 0
            End If
        End If
    End If
    lkuVendClass.Tag = lkuVendClass.Text
End Sub

Private Sub lkuVendor_LostFocus()
    Dim tlVendKey As Long
    If Len(Trim$(lkuVendor.Text)) > 0 Then
        If lkuVendor.Text <> lkuVendor.Tag Then
            tlVendKey = glGetValidLong(moAppDB.Lookup("VendKey", "tapVendor", "CompanyID = " & gsQuoted(msCompanyID) & " and VendID = " & gsQuoted(lkuVendor.Text)))
            If tlVendKey = 0 Then
                MsgBox "El Proovedor seleccionado no es valido", vbExclamation, "Alerta"
                lkuVendor.Text = ""
                bSetFocus lkuVendor
                GoTo Invalid
            End If
            lkuVendor.KeyValue = tlVendKey
            lkuVendor.Tag = lkuVendor.Text
            lblVendorName.Caption = moAppDB.Lookup("VendName", "tapVendor", "CompanyID = " & gsQuoted(msCompanyID) & " and VendID = " & gsQuoted(lkuVendor.Text))
            
            LoadDftlVendorValues
            lkuParentContract.RestrictClause = sGetParentRestrict
        End If
        Exit Sub
    Else
        lkuParentContract.RestrictClause = sGetParentRestrict
    End If
    
Invalid:
    lkuVendor.Tag = ""
    lkuVendor.KeyValue = 0
    lblVendorName.Caption = ""
End Sub


Private Sub moDmDetl_DMGridAfterDelete(lRow As Long, bValid As Boolean)
    bValid = False

'-- Perform processing after a row is deleted from the detail table here


    bValid = True
End Sub

Private Sub moDmDetl_DMGridAfterInsert(lRow As Long, bValid As Boolean)
    bValid = False

'-- Perform processing after a row is inserted into the detail table here

    
    bValid = True
End Sub

Private Sub moDmDetl_DMGridAfterUpdate(lRow As Long, bValid As Boolean)
    bValid = False

'-- Perform processing after a row is updated in the detail table here

    
    bValid = True
End Sub

Private Sub moDmDetl_DMGridAppend(lRow As Long)
    '-- This adds a new row to the grandchild detail grid


'-- Perform processing after a row is appended to the detail grid here

End Sub


Private Sub moDmDetl_DMGridRowLoaded(lRow As Long)
'************************************************************************************
'   Description:
'      Fires when each row of data is loaded by the DM into the grid.
'      This is a manual movement of a copy of the data from the DM to the correct
'      grid cell on the same row.
'************************************************************************************
Dim DBValue As Integer

'-- Perform processing as each row is loaded into the grid
End Sub

Private Sub modmForm_DMPreSave(bValid As Boolean)
    MsgBox ""
End Sub

Private Sub moDmHeader_DMAfterDelete(bValid As Boolean)
    bValid = False
    On Error GoTo ErrorHandler
'-- Perform processing after a record is deleted from the main table here
    moAppDB.ExecuteSQL "delete from tctContractLine where ContractKey =" & moDmHeader.GetColumnValue("ContractKey")
ErrorHandler:
    bValid = True
End Sub

Private Sub moDmHeader_DMAfterInsert(bValid As Boolean)
'      If sddType.ItemData = 2 Then
'         moAppDB.ExecuteSQL "UPDATE tctcontractgConfig SET nextnc = (SELECT REPLACE(nextnc,SUBSTRING(NextNC, CHARINDEX('USN',NextNC)+3,4), RIGHT('0000' + CONVERT(VARCHAR, CAST(SUBSTRING(NextNC, CHARINDEX('USN',NextNC)+3,4) AS INT)+1),4)) FROM tctcontractgConfig)"
'      ElseIf sddType.ItemData = 1 Then
'          moAppDB.ExecuteSQL "UPDATE tctcontractgConfig SET nextic = (SELECT REPLACE(nextic,SUBSTRING(Nextic, CHARINDEX('USI',Nextic)+3,4), RIGHT('0000' + CONVERT(VARCHAR, CAST(SUBSTRING(Nextic, CHARINDEX('USI',Nextic)+3,4) AS INT)+1),4)) FROM tctcontractgConfig)"
'      End If
'      bNum001 = False
      
End Sub

Private Sub moDmHeader_DMBeforeInsert(bValid As Boolean)
    Dim lKey As Long
    On Error GoTo CancelInsert
    bValid = True

 '  bValid = False

 
' With moClass.moAppDB
'    .SetInParamStr "tctContractBankInf"
'    .SetOutParam lKey
'    .ExecuteSP "spGetNextSurrogateKey"
'    lKey = .GetOutParam(2)
'    .ReleaseParams
' End With
' 'moDmDetl.SetColumnValue lRow, "ContractBankInfKey", lKey
'
'  txtContractBankInfKey.Text = lKey
 
 
    moDmHeader.SetColumnValue "CreateUser", msUserID
    moDmHeader.SetColumnValue "CreateDate", Format(DateTime.Now, gsGetLocalVBDateMask())
    
    Exit Sub
CancelInsert:
    bValid = False
End Sub

Private Sub moDmHeader_DMBeforeUpdate(bValid As Boolean)
    bValid = False
     Dim lKey As Long
    
    
'     With moClass.moAppDB
'    .SetInParamStr "tctContractBankInf"
'    .SetOutParam lKey
'    .ExecuteSP "spGetNextSurrogateKey"
'    lKey = .GetOutParam(2)
'    .ReleaseParams
'     End With
' 'moDmDetl.SetColumnValue lRow, "ContractBankInfKey", lKey
'
'  txtContractBankInfKey.Text = lKey
'  txtContratokey1.Text = moDmHeader.GetColumnValue("ContractKey")
'
    
'-- Perform processing before a record is updated in the main table here
    moDmHeader.SetColumnValue "UpdateUser", msUserID
    moDmHeader.SetColumnValue "UpdateDate", Format(DateTime.Now, gsGetLocalVBDateMask())
    
    bValid = True
End Sub

Private Sub moDmHeader_DMAfterUpdate(bValid As Boolean)
    bValid = False

'-- Perform processing after a record is updated in the main table here


    bValid = True
End Sub

Private Sub moDmHeader_DMBeforeDelete(bValid As Boolean)
    Dim lContractKey As Long
    bValid = False
    
    If grdOperations.DataRowCnt > 0 Then
         MsgBox "Imposible eliminar. El contrato tiene transacciones asociadas."
         bValid = False
    Else
            lContractKey = glGetValidLong(moDmHeader.GetColumnValue("ContractKey"))
    '-- Perform processing before a record is deleted from the main table here
        If lContractKey <> 0 Then
            moAppDB.ExecuteSQL "delete from tctContractLine where ContractKey =" & lContractKey
        End If
        bValid = True
    End If
    
End Sub

'Private Sub moDmHeader_DMDataDisplayed(oChild As Object)
'    If oChild Is moDmHeader Then
'
''-- Perform processing as a record is loaded onto the form after data is displayed here
'
'
'    End If
'
'End Sub

Private Sub moDmHeader_DMPostSave(bValid As Boolean)
    Dim lDesactStateKey As Long
    bValid = False
    lDesactStateKey = glGetValidLong(moAppDB.Lookup("ContractStateKey", "tctContractState", "ContractStateId = 'Inactivo'"))
'-- Perform processing after a record is saved to the main table here
    With moClass.moAppDB
       .SetInParam moDmHeader.GetColumnValue("ContractKey")
       .ExecuteSP "spctUpdateContractLog"
       .ReleaseParams
    End With
    
    If lDesactStateKey <> 0 Then
        If sddState.ItemData = lDesactStateKey Then
            moAppDB.ExecuteSQL "insert into tctContractAprobalHistorical (ContractKey, AprobalLevelKey, UserID, AprobalDate) select * from tctContractAprobals where ContractKey =" & moDmHeader.GetColumnValue("ContractKey")
            moAppDB.ExecuteSQL "delete from tctContractAprobals where ContractKey =" & moDmHeader.GetColumnValue("ContractKey")
            moAppDB.ExecuteSQL "update tctContract set Free = 0 where ContractKey =" & moDmHeader.GetColumnValue("ContractKey")
        End If
    End If
    
    bValid = True
End Sub

Private Sub moDmHeader_DMPreSave(bValid As Boolean)
    On Error GoTo ErrorHandler
    Dim sReason As String
    Dim bCancel As Boolean
    Dim lReasonCode As Long
    
    If Not ValidateFields Then
        bValid = False
        Exit Sub
    End If
    
    If txtAccountID.Text <> "" Then
      Exit Sub
    End If
    
    If txtAccountID.Text = "" Then
        txtAccountID.Text = ssAccountNo
        txtBankAddress.Text = ssBankAddress
        txtSWIFT.Text = ssSWIFT
        txtTitular.Text = ssTitular
        
        
    End If
    
    If moDmHeader.State = kDmStateEdit Then
        Set frmChngOrd.oClass = moClass
'        frmChngOrd.oClass = moClass
        frmChngOrd.CurrentNumber = gsGetValidStr(moDmHeader.GetColumnValue("ChngOrdNo"))
        frmChngOrd.ShowMe sReason, bCancel
       ' True, lReasonCode
        If bCancel Then
            bValid = False
            Exit Sub
        End If
    moDmHeader.SetColumnValue "ChngOrdReason", sReason
    moDmHeader.SetColumnValue "ChngOrdReasonCodeKey", lReasonCode
    moDmHeader.SetColumnValue "ChngOrdNo", giGetValidInt(moDmHeader.GetColumnValue("ChngOrdNo")) + 1

    txtChgNo.Text = moDmHeader.GetColumnValue("ChngOrdNo")
    
    moAppDB.ExecuteSQL "insert into tctContractChgOrder  SELECT *FROM tctContract AS p WHERE p.ContractKey = " & moDmHeader.GetColumnValue("ContractKey")
    End If
    Exit Sub
ErrorHandler:
   ' MsgBox "err - " & Err.Description
    Exit Sub
    gSetSotaErr Err, "frmContract", "moDmHeader_DMPreSave", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number
End Sub

Private Sub moDmHeader_DMReposition(oChild As Object)
'    If oChild Is moDmHeader Then
'
''-- Perform processing as a record is loaded onto the form before data is displayed here
'
'
'    End If
End Sub

Private Sub moDmHeader_DMStateChange(iOldState As Integer, iNewState As Integer)
    '
    If iNewState = kDmStateNone Then
        If Not moLE Is Nothing Then
            If moLE.State <> kGridNone Then
                On Error Resume Next
                moLE.InitDataReset
            End If
        End If
    End If
End Sub

Private Sub moDmHeader_DMBeforeTransaction(bValid As Boolean)
'*******************************************************************
' Description:
'    This routine will be called by Data Manager before the record
'    is saved. This is where form-level validation should occur.
'*******************************************************************
    bValid = False

'-- Perform validations and other processes that need to occur before the DB transaction here

        
    bValid = True
End Sub

Private Sub nbrDuration_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If nbrDuration.Tag <> gsGetValidStr(nbrDuration.Value) Then
        nbrDuration.Tag = gsGetValidStr(nbrDuration.Value)
        CalcFinishDate
    End If
    Exit Sub
ErrorHandler:
    gSetSotaErr Err, "frmContract", "nbrDuration_Validate", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number
End Sub

Private Sub nbrQtyVariation_Change()
    Dim lParentKey As Long
    Dim lItemKey As Long
    Dim dParentQty As Double
    
    On Error GoTo ErrorHandler
    
    lParentKey = glGetValidLong(lkuParentContract.KeyValue)
    lItemKey = lkuItem.KeyValue
    If lParentKey <> 0 And lItemKey <> 0 Then
        If giGetValidInt(moAppDB.Lookup("count(*)", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey)) > 0 Then
            dParentQty = gdGetValidDbl(moAppDB.Lookup("Qty", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey))
            If sddLineType.ItemData = kLineTypeModUp Then
                numItemQty.Value = dParentQty + nbrQtyVariation.Value
            End If
            If sddLineType.ItemData = kLineTypeModDown Then
                If nbrQtyVariation.Value > dParentQty Then
                    MsgBox "La cantidad a disminuir del artículo no puede ser mayor que la que posee en el contrato padre", vbExclamation, "Alerta"
                    nbrQtyVariation.Value = 0
                    nbrQtyVariation.Visible = False
                     numItemQty.Visible = True
                     numItemQty.Value = dParentQty
                Else
                    numItemQty.Value = dParentQty - nbrQtyVariation.Value
                End If

            End If
            If sddLineType.ItemData = kLineTypeModCos Then
                numItemQty.Value = dParentQty
                numItemQty.Visible = True
                nbrQtyVariation.Visible = False

            End If

            Exit Sub
        End If
    End If
    
'   numItemQty.Value = nbrQtyVariation.Value
     CalcLineAmt
    
'
'    If lItemKey <> 0 And nbrQtyVariation.Value < numMinLot.Value And numMinLot.Value <> 0 And nbrQtyVariation.Value <> 0 Then
'
'          MsgBox "La cantidad debe ser superior al Lote Mínimo.", vbExclamation, "Alerta"
'          bSetFocus nbrQtyVariation
'
'
'    End If
    Exit Sub
ErrorHandler:
'    gSetSotaErr Err, "frmContract", "nbrDuration_Validate", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
'        Err.Raise guSotaErr.Number
End Sub

Private Sub nbrQtyVariation_KeyPress(KeyAscii As Integer)
'  If lkuItem.KeyValue <> 0 And nbrQtyVariation.Value < numMinLot.Value And numMinLot.Value <> 0 And nbrQtyVariation.Value <> 0 Then
'
'          MsgBox "La cantidad debe ser superior al Lote Mínimo.", vbExclamation, "Alerta"
'          bSetFocus nbrQtyVariation
'
'
'    End If
End Sub

Private Sub nbrQtyVariation_LostFocus()
' If lkuItem.KeyValue <> 0 And nbrQtyVariation.Value < numMinLot.Value And numMinLot.Value <> 0 And nbrQtyVariation.Value <> 0 Then
'
'          MsgBox "La cantidad debe ser superior al Lote Mínimo.", vbExclamation, "Alerta"
'          bSetFocus nbrQtyVariation
'
'
'    End If
End Sub

Private Sub nbrQtyVariation_Validate(Cancel As Boolean)
 If sddLineType.ItemData = kLineTypeModCos Then
   nbrQtyVariation.Enabled = False
   numItemQty.Enabled = False
   nbrQtyVariation.Visible = False
   
   Else
    nbrQtyVariation.Enabled = True
   numItemQty.Enabled = True
    nbrQtyVariation.Visible = True
 End If
 
End Sub

Private Sub numDeliveryTime_KeyPress(KeyAscii As Integer)
    If gsGetValidStr(numDeliveryTime.Value) <> numDeliveryTime.Tag Then
        moLE.GridEditChange numDeliveryTime
        Exit Sub
    End If
    numDeliveryTime.Tag = gsGetValidStr(numDeliveryTime.Value)
End Sub

Private Sub numItemQty_Validate(Cancel As Boolean)
   numItemQty.Value = gdGetValidDbl(numItemQty.Text)
    CalcLineAmt
End Sub

Private Sub sbrMain_ButtonClick(sButton As String)
    HandleToolbarClick sButton
End Sub

Private Sub sclStartDate_Validate(Cancel As Boolean)
    On Error GoTo ErrorHandler
    If sclStartDate.Value <> sclStartDate.Tag Then
        sclStartDate.Tag = sclStartDate.Value
        CalcFinishDate
    End If
    Exit Sub
ErrorHandler:
    gSetSotaErr Err, "frmContract", "sclStartDate_Validate", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number
End Sub

Private Sub sddClasification_Click(Cancel As Boolean, ByVal PrevIndex As Long, ByVal NewIndex As Long)
    On Error GoTo ErrorHandler
    If NewIndex > -1 Then
        If sddClasification.ItemData(NewIndex) = 1 Then
            numDeliveryTime.Enabled = True
            numMinLot.Enabled = True
            numMaxLot.Enabled = True
            numRoundValue.Enabled = True
            lkuItem.RestrictClause = "ItemKey In (SELECT p.ItemKey FROM timItem AS p WHERE p.ItemType <> 3 AND p.[Status] = 1)"
        Else
            numDeliveryTime.Enabled = False
            numMinLot.Enabled = False
            numMaxLot.Enabled = False
            numRoundValue.Enabled = False
            lkuItem.RestrictClause = "ItemKey In (SELECT p.ItemKey FROM timItem AS p WHERE p.ItemType = 3 AND p.[Status] = 1)"
        End If
    
    End If
    Exit Sub
ErrorHandler:
    gSetSotaErr Err, "frmContract", "sddLineType_Click", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
    Cancel = True
End Sub

Private Sub sddLineType_Click(Cancel As Boolean, ByVal PrevIndex As Long, ByVal NewIndex As Long)
    Dim lParentKey As Long
    Dim lItemKey As Long
    Dim bFromParent As Boolean
    
    On Error GoTo ErrorHandler
    
    If NewIndex > -1 And Len(Trim$(lkuItem.Text)) > 0 Then
        lParentKey = glGetValidLong(lkuParentContract.KeyValue)
        lItemKey = lkuItem.KeyValue
        bFromParent = False
        If lParentKey <> 0 Then
            If giGetValidInt(moAppDB.Lookup("count(*)", "tctContractLine", "ItemKey = " & lItemKey & " and ContractKey = " & lParentKey)) > 0 Then
                bFromParent = True
            End If
        End If
        
        If sddLineType.ItemData(NewIndex) > 1 And Not bFromParent Then
            MsgBox "Los artículos que no se encuentran en el contrato padre solo pueden ser agregados", vbExclamation, "Alerta"
            Cancel = True
        End If
        
        If sddLineType.ItemData(NewIndex) = 1 And bFromParent Then
            MsgBox "Los artículos que se encuentran en el contrato padre no pueden ser agregados", vbExclamation, "Alerta"
            Cancel = True
        End If
        
        Select Case sddLineType.ItemData(NewIndex)
            Case kLineTypeAdd
                numItemQty.Visible = True
                nbrQtyVariation.Visible = False
                nbrQtyVariation.Value = 0
                numItemQty.Visible = True
            
                lkuItem.Enabled = True
                txtDescription.Enabled = False
                numDeliveryTime.Enabled = True
                numRoundValue.Enabled = True
                sddUM.Enabled = True
                currUnitCost.Enabled = True
                numMinLot.Enabled = True
                numMaxLot.Enabled = True
                
            Case kLineTypeDel
                
                numItemQty.Visible = True
                nbrQtyVariation.Visible = False
                nbrQtyVariation.Value = 0
                numItemQty.Visible = True
                
                
                   
                lkuItem.Enabled = False
                txtDescription.Enabled = False
                numDeliveryTime.Enabled = False
                numRoundValue.Enabled = False
                sddUM.Enabled = False
                currUnitCost.Enabled = False
                numMinLot.Enabled = False
                numMaxLot.Enabled = False
                
                
                
             Case kLineTypeModCos
                numItemQty.Visible = True
                nbrQtyVariation.Visible = False
                nbrQtyVariation.Value = 0
                numItemQty.Visible = True
                numItemQty.Enabled = False
                nbrQtyVariation.Enabled = False
                
                
                 
                  lkuItem.Enabled = True
                txtDescription.Enabled = False
                numDeliveryTime.Enabled = True
                numRoundValue.Enabled = True
                sddUM.Enabled = True
                currUnitCost.Enabled = True
                numMinLot.Enabled = True
                numMaxLot.Enabled = True
                
             
                
                 
                
            Case Else
            
            
                numItemQty.Visible = False
                nbrQtyVariation.Visible = True
              '  numItemQty.Value = 0
              
              
              
               
                  lkuItem.Enabled = True
                txtDescription.Enabled = False
                numDeliveryTime.Enabled = True
                numRoundValue.Enabled = True
                sddUM.Enabled = True
                currUnitCost.Enabled = True
                numMinLot.Enabled = True
                numMaxLot.Enabled = True
        End Select
        
        Select Case sddLineType.ItemData(NewIndex)
            Case kLineTypeAdd
                numItemQty.Protected = False
                 numItemQty.Visible = True
                 nbrQtyVariation.Visible = False
            Case kLineTypeDel
                numItemQty.Value = 0
                numItemQty.Protected = True
                numItemQty.Visible = True
                nbrQtyVariation.Visible = False
                
                  lkuItem.Enabled = False
                txtDescription.Enabled = False
                numDeliveryTime.Enabled = False
                numRoundValue.Enabled = False
                sddUM.Enabled = False
                currUnitCost.Enabled = False
                numMinLot.Enabled = False
                numMaxLot.Enabled = False
                
                
            Case kLineTypeModCos
                numItemQty.Protected = False
                numItemQty.Visible = True
                nbrQtyVariation.Visible = False
             Case kLineTypeModDown Or kLineTypeModUp
                numItemQty.Visible = False
                nbrQtyVariation.Visible = True
                nbrDuration.Enabled = True
                'nbrQtyVariation.Visible = False
        End Select
    End If
    If Cancel Then
        sddLineType.Tag = PrevIndex
    Else
        sddLineType.Tag = NewIndex
    End If
    
'        Select Case sddLineType.Text
'            Case "Adicionar"
'                numItemQty.Protected = False
'            Case "Eliminar"
'                numItemQty.Value = 0
'                numItemQty.Protected = True
'            Case "Modificar Costo"
'                numItemQty.Protected = False
'                numItemQty.Visible = True
'                nbrQtyVariation.Visible = False
'             Case "Modificar Aumento" Or "Modificar Decremento"
'                numItemQty.Visible = False
'                nbrQtyVariation.Visible = True
'                nbrDuration.Enabled = True
'                'nbrQtyVariation.Visible = False
'        End Select
    
    
    Exit Sub
    
 
    
ErrorHandler:
    gSetSotaErr Err, "frmContract", "sddLineType_Click", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
    Cancel = True
End Sub

Private Sub sddLineType_Validate(Cancel As Boolean)
'   If sddLineType.ItemData = kLineTypeModCos Then
'
'              '  numItemQty.Protected = False
'                 numItemQty.Enabled = False
'                 nbrQtyVariation.Enabled = False
'
'                 Else
'                 numItemQty.Enabled = True
'                  nbrQtyVariation.Enabled = True
'   End If
  
End Sub

Private Sub sddState_Click(Cancel As Boolean, ByVal PrevIndex As Long, ByVal NewIndex As Long)
    Dim lDesactStateKey As Long
    Dim lCreateStateKey As Long
    Dim lActiveStateKey As Long
    On Error GoTo ErrorHandler
    
    
    If PrevIndex = NewIndex Then Exit Sub
    
    If NewIndex > -1 And PrevIndex > -1 Then
        lDesactStateKey = glGetValidLong(moAppDB.Lookup("ContractStateKey", "tctContractState", "ContractStateId = 'Inactivo'"))
        lCreateStateKey = glGetValidLong(moAppDB.Lookup("ContractStateKey", "tctContractState", "ContractStateId = 'Creado'"))
        lActiveStateKey = glGetValidLong(moAppDB.Lookup("ContractStateKey", "tctContractState", "ContractStateId = 'Activo'"))
        
        If lCreateStateKey <> 0 Then
            If lCreateStateKey = sddState.ItemData(NewIndex) Then
                MsgBox "No se puede devolver el contrato a ese estado.", vbExclamation, "Alerta"
                Cancel = True
                Exit Sub
            End If
        End If
        
        If lActiveStateKey <> 0 Then
            If lActiveStateKey = sddState.ItemData(NewIndex) Then
                MsgBox "No se puede activar el contrato manualmente.", vbExclamation, "Alerta"
                Cancel = True
                Exit Sub
            End If
        End If
        
        If lActiveStateKey <> 0 And lDesactStateKey <> 0 Then
            If lActiveStateKey = sddState.ItemData(PrevIndex) And lDesactStateKey <> sddState.ItemData(NewIndex) Then
                MsgBox "Los contratos activos solo pueden ser desactivados.", vbExclamation, "Alerta"
                Cancel = True
                Exit Sub
            End If
        End If
        
        If lDesactStateKey <> 0 Then
            If lDesactStateKey = sddState.ItemData(NewIndex) Then
                If moClass.moFramework.GetSecurityEventPerm(kSecurityEventDesactContr, msUserID, True) = 0 Then
                    MsgBox "No tiene permiso para desactivar este contrato", vbExclamation, "Alerta"
                    Cancel = True
                    Exit Sub
                End If
                If HavePendingOperations Then
                    MsgBox "Imposible desactivar el contrato/suplemento, aún existen órdenes de compra o requisiciones pendientes asociadas a el", vbExclamation, "Alerta"
                    Cancel = True
                    Exit Sub
                End If
            End If
            
            If lDesactStateKey = sddState.ItemData(PrevIndex) Then
                If moClass.moFramework.GetSecurityEventPerm(kSecurityEventReactContr, msUserID, True) = 0 Then
                    MsgBox "No tiene permiso para reactivar este contrato", vbExclamation, "Alerta"
                    Cancel = True
                    Exit Sub
                End If
                
                moAppDB.ExecuteSQL "DELETE FROM tctContractAprobals WHERE ContractKey = " & moDmHeader.GetColumnValue("ContractKey")
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
    gSetSotaErr Err, "frmContract", "sddState_Click", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
    Cancel = True
End Sub

Private Sub sddType_Click(Cancel As Boolean, ByVal PrevIndex As Long, ByVal NewIndex As Long)
    If sddType.ItemData = kContractTypeSuplement Then
        lkuParentContract.Visible = True
        'lblParentContract.Visible = True
        lkuParentContract.RestrictClause = sGetParentRestrict
        lblParentContract.Visible = True
        lkuParentContract.Enabled = True
        lblNoContrato.Caption = "Suplemento"
        lblLineType.Visible = True
        sddLineType.Visible = True
       ' nbrQtyVariation.Visible = True
        numItemQty.Visible = False
        cmdGenParams.Visible = True
        
      

    'Visivility
    gGridHideColumn grdMain, kColUnitMeasID
    gGridHideColumn grdMain, kColItemQty
    
       
        

    Else
        lkuParentContract.Visible = False
        lblParentContract.Visible = False
        txtContratokey1.Text = lkuParentContract.Text
        
        lblNoContrato.Caption = "No Contrato"
        If Len(Trim$(lkuParentContract.Text)) > 0 Then
            lkuParentContract.Text = ""
            lkuParentContract.KeyValue = 0
        End If
        lblLineType.Visible = False
        sddLineType.Visible = False
        nbrQtyVariation.Visible = False
        numItemQty.Visible = True
        cmdGenParams.Visible = False
        
         'Visivility
    gGridShowColumn grdMain, kColUnitMeasID
    gGridShowColumn grdMain, kColItemQty
    
        
    End If
End Sub

Private Sub sddType_LostFocus()
 Dim sNum As String
  If bNum001 = True Then
     If sddType.ItemData = 2 Then
'        sNum = gsGetValidStr(moAppDB.Lookup("REPLACE(nextnc,SUBSTRING(NextNC, CHARINDEX('USN',NextNC)+3,4), RIGHT('0000' + CONVERT(VARCHAR, CAST(SUBSTRING(NextNC, CHARINDEX('USN',NextNC)+3,4) AS INT)),4))", "tctContractGConfig", "1 = 1"))
'        If sNum <> "" Then
'          txtContractNo.Text = sNum
'          bNum001 = True
'          End If

         GetNextContractNo1
      ElseIf sddType.ItemData = 1 Then
'        sNum = gsGetValidStr(moAppDB.Lookup("REPLACE(nextic,SUBSTRING(Nextic, CHARINDEX('USI',Nextic)+3,4), RIGHT('0000' + CONVERT(VARCHAR, CAST(SUBSTRING(Nextic, CHARINDEX('USI',Nextic)+3,4) AS INT)),4))", "tctContractGConfig", "1 = 1"))
'         If sNum <> "" Then
'          txtContractNo.Text = sNum
'          bNum001 = True
'          End If

        GetNextContractNo1
       End If
  End If

End Sub

Private Sub sddType_Validate(Cancel As Boolean)
 GetNextContractNo1
End Sub

Private Sub sddUM_Click(Cancel As Boolean, ByVal PrevIndex As Long, ByVal NewIndex As Long)
    On Error GoTo ErrorHandler
    Dim dPrevFactor As Double
    Dim dNewFactor As Double
    Dim dQty As Double
    Dim dCost As Double
    
    If mlActiveRow <> moLE.ActiveRow Then Exit Sub
    If lkuItem.KeyValue = 0 Then Exit Sub
    If PrevIndex <> NewIndex And PrevIndex > -1 And NewIndex <> -1 Then
        dQty = numItemQty.Value
        dCost = currUnitCost.Amount
        dPrevFactor = gdGetValidDbl(moAppDB.Lookup("ConversionFactor", "timItemUnitOfMeas", "ItemKey =" & lkuItem.KeyValue & " and TargetUnitMeasKey = " & sddUM.ItemData(PrevIndex)))
        dNewFactor = gdGetValidDbl(moAppDB.Lookup("ConversionFactor", "timItemUnitOfMeas", "ItemKey =" & lkuItem.KeyValue & " and TargetUnitMeasKey = " & sddUM.ItemData(NewIndex)))
        dQty = dQty * dPrevFactor / dNewFactor
        dCost = dCost / dPrevFactor * dNewFactor
        numItemQty.Value = dQty
        currUnitCost.Amount = dCost
    End If
    Exit Sub
ErrorHandler:
    gSetSotaErr Err, "frmContract", "sddUM_Click", VBRIG_IS_FORM                         'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number
End Sub


Public Sub Validar(DatosActuales As String, Caracter As Integer)
  ' Salimos si se ha pulsado la tecla de Retroceso
  If Caracter = 8 Then Exit Sub
  ' Salimos si es de 0 a 9
  If InStr("0123456789", Chr$(Caracter)) Then Exit Sub
  ' Si es punto y no está en el contenido salimos
  If Caracter = 46 And InStr(DatosActuales, ".") = 0 Then Exit Sub
  ' Borramos el Caracter introducido
  Caracter = 0
End Sub
Private Sub snFlete_KeyPress(KeyAscii As Integer)
   Validar snFlete.Text, KeyAscii
End Sub

Private Sub snFlete_LostFocus()
  CalcContractAmt
End Sub

Private Sub snFlete_Validate(Cancel As Boolean)
  
  Dim dd As Double
  Dim dd2 As Double
  
  Dim cc As String
  Dim cc2 As String
  
    If snFlete.Text <> "" Then
    
       If CDbl(snFlete) > 100000000000000# Then
       dd = CDbl(snFlete)
       dd = dd / 100000000000000#
        dd2 = dd - FormatNumber(dd, 0, vbFalse, vbFalse)
        dd2 = dd2 * 100000000000000#

        dd = FormatNumber(dd, 0, vbFalse, vbFalse)
    '   dd = CDbl(currAmt) - (CDbl(currAmt) / 1E+16)
        If (dd2 < 0) Then
         dd2 = dd2 * (-1)
        End If


       cc = CStr(FormatNumber(dd, 0, , vbFalse, vbFalse))
       cc2 = CStr(FormatNumber(dd2, 2, , vbFalse, vbFalse))
      snFlete.Text = cc + cc2
      
      
     ' snFlete.Text = Format(snFlete, "###############################000000000000000000000000000000.0000")
      Else
    
      End If
  
   End If
End Sub

Private Sub snOtrosCostos_KeyPress(KeyAscii As Integer)
   Validar snOtrosCostos.Text, KeyAscii
End Sub

Private Sub snOtrosCostos_LostFocus()
  CalcContractAmt
End Sub

Private Sub snOtrosCostos_Validate(Cancel As Boolean)
 Dim dd As Double
 Dim dd2 As Double
 
 Dim cc As Double
 Dim cc2 As Double
 
 
  If snOtrosCostos.Text <> "" Then
     If CDbl(snOtrosCostos) > 100000000000000# Then
     
       dd = CDbl(snOtrosCostos)
       dd = dd / 100000000000000#
        dd2 = dd - FormatNumber(dd, 0, vbFalse, vbFalse)
        dd2 = dd2 * 100000000000000#

        dd = FormatNumber(dd, 0, vbFalse, vbFalse)
    '   dd = CDbl(currAmt) - (CDbl(currAmt) / 1E+16)
        If (dd2 < 0) Then
         dd2 = dd2 * (-1)
        End If


       cc = CStr(FormatNumber(dd, 0, , vbFalse, vbFalse))
       cc2 = CStr(FormatNumber(dd2, 2, , vbFalse, vbFalse))
      snOtrosCostos.Text = cc + cc2
     
      Else
    
      End If
  End If
 
End Sub



Private Sub snSeguro_KeyPress(KeyAscii As Integer)
 Validar snSeguro.Text, KeyAscii
End Sub

Private Sub snSeguro_LostFocus()
  CalcContractAmt
End Sub

Private Sub snSeguro_Validate(Cancel As Boolean)

 Dim dd As Double
 Dim dd2 As Double
 
 Dim cc As String
 Dim cc2 As String
  
 
  If snSeguro.Text <> "" Then
      If CDbl(snSeguro) > 100000000000000# Then
      
       dd = CDbl(snSeguro)
       dd = dd / 100000000000000#
        dd2 = dd - FormatNumber(dd, 0, vbFalse, vbFalse)
        dd2 = dd2 * 100000000000000#

        dd = FormatNumber(dd, 0, vbFalse, vbFalse)
    '   dd = CDbl(currAmt) - (CDbl(currAmt) / 1E+16)
        If (dd2 < 0) Then
         dd2 = dd2 * (-1)
        End If


       cc = CStr(FormatNumber(dd, 0, , vbFalse, vbFalse))
       cc2 = CStr(FormatNumber(dd2, 2, , vbFalse, vbFalse))
      snSeguro.Text = cc + cc2
      
      Else
    
      End If
  End If

End Sub

Private Sub txtContract_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
    If moDmHeader.State = kDmStateNone Then
        If Len(Trim$(txtContract.Text)) > 0 Then
            If txtContract.Text <> txtContract.Tag Then
                If IsValidContract Then
                    txtContract.Tag = txtContract.Text
                    valMgr_KeyChange
                End If
            End If
        End If
    End If
    Exit Sub
ErrorHandler:
        gSetSotaErr Err, sMyName, "txtContract_Validate", VBRIG_IS_FORM                    'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number
End Sub

Private Sub txtContractNo_Validate(Cancel As Boolean)
    If txtContractNo.Tag <> txtContractNo.Text Then
        If bIsNumberInUse Then
            txtContractNo.Text = ""
            txtContractNo.Tag = ""
             lkuParentContract_LostFocus
            Exit Sub
        End If
        txtContractNo.Tag = txtContractNo.Text
    End If
End Sub

Private Sub txtDescription_Change()
    If txtDescription.Text <> txtDescription.Tag Then
        moLE.GridEditChange txtDescription
        Exit Sub
    End If
    txtDescription.Tag = txtDescription.Text
    
  '  nbrQtyVariation_Change
    
    
End Sub

Private Sub valMgr_KeyChange()
    Dim iKeyChangeCode  As Integer
    Dim iActualTab As Integer
    Dim lActiveStateKey As Long
    Dim bLiberado As Long
    
    
    Dim lState As Long
    
    On Error GoTo ErrorHandler
    ClearDetlFields
    moDmHeader.SetColumnValue "ContractKey", mlContractKey
    moDmHeader.SetColumnValue "CompanyID", msCompanyID

    iKeyChangeCode = moDmHeader.KeyChange

    
    lActiveStateKey = glGetValidLong(moAppDB.Lookup("ContractStateKey", "tctContractState", "ContractStateId = 'Activo'"))
    
    bLiberado = glGetValidLong(moAppDB.Lookup("Free", "tctContract", "ContractKey = " & mlContractKey))
    
    Select Case iKeyChangeCode
        Case kDmKeyFound, kDmKeyNotFound
            txtContract.Enabled = False
            
            If moDmHeader.State = kDmStateAdd Then
                txtChgNo.Text = "0"
                sclSignatureDate.Value = msBusDate
                sclStartDate.Value = msBusDate
                lState = glGetValidLong(moAppDB.Lookup("ContractStateKey", "tctContractState", "ContractStateId = 'Creado'"))
                If lState > 0 Then sddState.ItemData = lState
                sddClasification.ItemData = 1
            End If
            CalcFinishDate
            moLE.InitDataLoaded
            moLE.Grid_Click 1, 1
            LoadOperations
            
            currAmt_Validate (True)
            snFlete_Validate (True)
            
           ssAccountNo = ""
           ssBankAddress = ""
           ssSWIFT = ""
           ssTitular = ""
            
            If tabDataEntry <> 1 Then
                iActualTab = tabDataEntry.Tab
                tabDataEntry.Tab = 1
                tabDataEntry.Tab = iActualTab
            End If
            If sddState.ItemData = lActiveStateKey Or bLiberado = 1 Then
                DisableFields
            Else
                EnableFields
                
                 If sddType.Text = "Suplemento" Then
                    lkuVendor.Enabled = False
                    sddType.Enabled = False
                   ' txtContractNo.Enabled = False
                    sddPaymentTerms.Enabled = False
                  '  lkuParentContract.Enabled = False
'                    sclSignatureDate.Enabled = False
'                    sclStartDate.Enabled = False
'                    nbrDuration.Enabled = False
                    lkuContact.Enabled = False
                    lkuVendClass.Enabled = False
                    sddFOB.Enabled = False
                    SddCurrID.Enabled = False
                    
'                    lkuItem.Enabled = False
'                    txtDescription.Enabled = False
'                    numDeliveryTime.Enabled = False
'                    numRoundValue.Enabled = False
'                    sddUM.Enabled = False
'                    currUnitCost.Enabled = False
'                    numItemQty.Enabled = False
'                    numMinLot.Enabled = False
'                    numMaxLot.Enabled = False
'                    sddLineType.Enabled = False
'                    nbrQtyVariation.Enabled = False
'                    grdMain.Enabled = False
                    
                  '  sddState.Enabled = False
                   ' snOtrosCostos.Enabled = False
                    txtDescription.Enabled = False
                    sddClasification.Enabled = False
                       Else
                   lkuVendor.Enabled = True
                    sddType.Enabled = True
                   ' txtContractNo.Enabled = True
                  '  sddPaymentTerms.Enabled = True
                    lkuParentContract.Enabled = True
'                    sclSignatureDate.Enabled = False
'                    sclStartDate.Enabled = False
'                    nbrDuration.Enabled = False
                    lkuContact.Enabled = True
                    lkuVendClass.Enabled = True
                    sddFOB.Enabled = True
                    SddCurrID.Enabled = True
                    
'                    lkuItem.Enabled = False
'                    txtDescription.Enabled = False
'                    numDeliveryTime.Enabled = False
'                    numRoundValue.Enabled = False
'                    sddUM.Enabled = False
'                    currUnitCost.Enabled = False
'                    numItemQty.Enabled = False
'                    numMinLot.Enabled = False
'                    numMaxLot.Enabled = False
'                    sddLineType.Enabled = False
'                    nbrQtyVariation.Enabled = False
'                    grdMain.Enabled = False
                    
                  '  sddState.Enabled = True
                  '  snOtrosCostos.Enabled = True
                    txtDescription.Enabled = True
                    sddClasification.Enabled = True
                   End If
                
                
                
            End If
            
            
        If sddLineType.ItemData = kLineTypeModCos Then
          nbrQtyVariation.Enabled = False
          numItemQty.Enabled = False
           nbrQtyVariation.Visible = False
        ElseIf sddLineType.ItemData = kLineTypeDel Then
        
                lkuItem.Enabled = False
                txtDescription.Enabled = False
                numDeliveryTime.Enabled = False
                numRoundValue.Enabled = False
                sddUM.Enabled = False
                currUnitCost.Enabled = False
                numMinLot.Enabled = False
                numMaxLot.Enabled = False
                
                nbrQtyVariation.Enabled = False
                numItemQty.Enabled = False
                nbrQtyVariation.Visible = False
        
        Else
        
         nbrQtyVariation.Enabled = True
          numItemQty.Enabled = True
           nbrQtyVariation.Visible = True
        End If
          
        Case kDmKeyNotComplete
    End Select
    Exit Sub
ErrorHandler:
    MsgBox "valMgr_KeyChange()_" & Err.Description, vbCritical, "Error"
End Sub


Private Sub valMgr_Validate(oControl As Object, iReturn As SOTAVM.SOTA_VALID_RETURN_TYPES, sMessage As String, ByVal iLevel As SOTAVM.SOTA_VALIDATION_LEVELS)
    'mbFromCode = True
    
    'iReturn = SOTA_INVALID
    
'-- Call individual validation routines for controls here
End Sub

Private Sub ClearDetlFields()

'-- Clear out controls that are not bound to Line Entry here
    txtContract.Enabled = True
    lkuVendor.Enabled = True
    sddType.Enabled = True
    lblVendorName.Caption = ""
    txtOutDate.Caption = ""
    lkuContact.Tag = ""
    lkuVendor.Tag = ""
    txtContract.Tag = ""
    currAmt.Text = "0"
    sddState.ListIndex(False) = -1
    txtAprobalLevel.Caption = ""
    '-- Set the detail controls as valid so the old values are correct
    valMgr.Reset
    
End Sub

Private Sub tabSubDetl_Click(PreviousTab As Integer)
    '-- Disable controls on hidden tabs
    HandleSubTabClick
End Sub

Private Sub tbrMain_ButtonClick(Button As String)
    HandleToolbarClick Button
End Sub

Private Sub tabDataEntry_Click(PreviousTab As Integer)
    Dim bValid As Boolean
    Static bInTabChange As Boolean
    
    If bInTabChange Then Exit Sub
    
    bInTabChange = True
    
    '-- Do nothing if the tab is the same as last time
    'If (tabVoucher.Tab = PreviousTab) Then
    '    bInTabChange = False
    '    Exit Sub
    'End If
        
    bValid = True
    
    If (PreviousTab = kiDetailTab) And (tabDataEntry.Tab <> PreviousTab) Then
        bValid = moLE.TabChange(tabDataEntry.Tab, PreviousTab)
    End If
    
    If bValid Then

'-- Handle specific tab change logic here

    End If
    
    '-- Disable controls on hidden tabs
    pnlTab(PreviousTab).Enabled = False
    pnlTab(tabDataEntry.Tab).Enabled = True
    
    bInTabChange = False
End Sub

Private Sub HandleTabClick()
    Dim i As Integer
    
    '-- Disable all panels
End Sub

Private Sub HandleSubTabClick()
    Dim i As Integer
    
    '-- Disable all panels
'    '-- Enable the visible panel
End Sub

Private Sub Form_Activate()
#If CUSTOMIZER Then
    If moFormCust Is Nothing Then
        Set moFormCust = CreateObject("SOTAFormCustRT.clsFormCustRT")
        If Not moFormCust Is Nothing Then
                moFormCust.Initialize Me, goClass
                Set moFormCust.CustToolbarMgr = tbrMain
                moFormCust.ApplyDataBindings moDmHeader
                moFormCust.ApplyFormCust
        End If
    End If
#End If

    Dim i As Integer

    '-- Setup the form (validation) manager control
    With valMgr
        Set .Framework = moClass.moFramework
        '.Keys.Add lkuVouchNo
    .Keys.Add txtContract
        .Init
    End With

End Sub


Public Sub grdMain_Click(ByVal Col As Long, ByVal Row As Long)
    If (moLE.GridEditDone(Col, Row) <> kLeSuccess) Then
        Exit Sub
    End If

    '-- Select the row that the user Selected
    moLE.Grid_Click Col, Row
End Sub

Private Sub grdMain_DblClick(ByVal Col As Long, ByVal Row As Long)
    moLE.Grid_DblClick Col, Row
End Sub

Private Sub grdMain_DragDrop(Source As Control, x As Single, y As Single)
    moLE.Grid_DragDrop x, y
End Sub

Private Sub grdMain_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    moLE.Grid_DragOver x, y, State
End Sub

Private Sub grdMain_GotFocus()
    moLE.Grid_GotFocus
End Sub

Public Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer)
    moLE.Grid_KeyDown KeyCode, Shift
End Sub

Private Sub grdMain_LostFocus()
    moLE.Grid_LostFocus
End Sub

Private Sub grdMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    moLE.Grid_MouseDown Button, Shift, x, y
End Sub

Private Sub grdMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   moLE.Grid_MouseMove Button, Shift, x, y
End Sub

Private Sub moDmDetl_DMGridBeforeInsert(lRow As Long, bValid As Boolean)

 bValid = False
 Dim lKey As Long
'
' With moClass.moAppDB
'    .SetInParamStr "tctContractLine"
'    .SetOutParam lKey
'    .ExecuteSP "spGetNextSurrogateKey"
'    lKey = .GetOutParam(2)
'    .ReleaseParams
' End With
 
     lKey = glGetValidLong(moAppDB.Lookup("top 1 (ContractLineKey)", "tctContractLine", " 1 = 1 order by ContractLineKey desc"))
     lKey = lKey + 1
 moDmDetl.SetColumnValue lRow, "ContractLineKey", lKey
 moDmDetl.SetColumnValue lRow, "ContractKey", moDmHeader.GetColumnValue("ContractKey")
 moDmDetl.SetColumnValue lRow, "CreateUser", msUserID
 moDmDetl.SetColumnValue lRow, "CreateDate", Format(DateTime.Now, gsGetLocalVBDateMask())
 bValid = True
End Sub

Public Property Get MyApp() As Object
    Set MyApp = App
End Property

Public Property Get MyForms() As Object
    Set MyForms = Forms
End Property

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
    
    ETWhereClause = ""
    
    ' Specific checks go here
    
    Err.Clear
    
End Function


Public Property Get UseHTMLHelp() As Boolean
    UseHTMLHelp = True   ' Form uses HTML Help (True/False)
End Property


Private Sub CustomButton_Click(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnClick CustomButton(Index)
#End If
End Sub

Private Sub CustomButton_GotFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomButton(Index)
#End If
End Sub

Private Sub CustomCurrency_Change(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomCurrency(Index)
#End If
End Sub

Private Sub CustomFrame_DblClick(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomFrame(Index)
#End If
End Sub

Private Sub CustomOption_Click(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnClick CustomOption(Index)
#End If
End Sub

Private Sub CustomSpin_UpClick(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnSpinUp CustomSpin(Index)
#End If
End Sub

Private Sub picDrag_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseDown Index, Button, Shift, x, y
    End If
#End If
End Sub

Private Sub picDrag_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseMove Index, Button, Shift, x, y
    End If
#End If
End Sub

Private Sub picDrag_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_MouseUp Index, Button, Shift, x, y
    End If
#End If
End Sub

Private Sub picDrag_Paint(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then
        moFormCust.picDrag_Paint Index
    End If
#End If
End Sub

Private Sub CustomButton_LostFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomButton(Index)
#End If
End Sub

Private Sub CustomCheck_Click(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnClick CustomCheck(Index)
#End If
End Sub

Private Sub CustomCheck_GotFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomCheck(Index)
#End If
End Sub

Private Sub CustomCheck_LostFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomCheck(Index)
#End If
End Sub

Private Sub CustomCombo_Change(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomCombo(Index)
#End If
End Sub

Private Sub CustomCombo_Click(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnClick CustomCombo(Index)
#End If
End Sub

Private Sub CustomCombo_DblClick(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomCombo(Index)
#End If
End Sub

Private Sub CustomCombo_GotFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomCombo(Index)
#End If
End Sub

Private Sub CustomCombo_KeyPress(Index As Integer, KeyAscii As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomCombo(Index), KeyAscii
#End If
End Sub

Private Sub CustomCombo_LostFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomCombo(Index)
#End If
End Sub

Private Sub CustomCurrency_GotFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomCurrency(Index)
#End If
End Sub

Private Sub CustomCurrency_KeyPress(Index As Integer, KeyAscii As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomCurrency(Index), KeyAscii
#End If
End Sub

Private Sub CustomCurrency_LostFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomCurrency(Index)
#End If
End Sub

Private Sub CustomDate_Change(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomDate(Index)
#End If
End Sub

Private Sub CustomDate_Click(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnClick CustomDate(Index)
#End If
End Sub

Private Sub CustomDate_DblClick(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomDate(Index)
#End If
End Sub

Private Sub CustomDate_GotFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomDate(Index)
#End If
End Sub

Private Sub CustomDate_KeyPress(Index As Integer, KeyAscii As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomDate(Index), KeyAscii
#End If
End Sub

Private Sub CustomDate_LostFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomDate(Index)
#End If
End Sub

Private Sub CustomFrame_Click(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnClick CustomFrame(Index)
#End If
End Sub

Private Sub CustomLabel_Click(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnClick CustomLabel(Index)
#End If
End Sub

Private Sub CustomLabel_DblClick(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomLabel(Index)
#End If
End Sub

Private Sub CustomMask_Change(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomMask(Index)
#End If
End Sub

Private Sub CustomMask_GotFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomMask(Index)
#End If
End Sub

Private Sub CustomMask_KeyPress(Index As Integer, KeyAscii As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomMask(Index), KeyAscii
#End If
End Sub

Private Sub CustomMask_LostFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomMask(Index)
#End If
End Sub

Private Sub CustomNumber_Change(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnChange CustomNumber(Index)
#End If
End Sub

Private Sub CustomNumber_GotFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomNumber(Index)
#End If
End Sub

Private Sub CustomNumber_KeyPress(Index As Integer, KeyAscii As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnKeyPress CustomNumber(Index), KeyAscii
#End If
End Sub

Private Sub CustomNumber_LostFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomNumber(Index)
#End If
End Sub

Private Sub CustomOption_DblClick(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnDblClick CustomOption(Index)
#End If
End Sub

Private Sub CustomOption_GotFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnGotFocus CustomOption(Index)
#End If
End Sub

Private Sub CustomOption_LostFocus(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnLostFocus CustomOption(Index)
#End If
End Sub

Private Sub CustomSpin_DownClick(Index As Integer)
#If CUSTOMIZER Then
    If Not moFormCust Is Nothing Then moFormCust.OnSpinDown CustomSpin(Index)
#End If
End Sub

Private Sub txtChildDescription_Change()    'LE Control Change Event
    moLE.GridEditChange txtDescription
    Exit Sub
End Sub

Private Sub numMaxLot_Change()    'LE Control Change Event
    If gsGetValidStr(numMaxLot.Value) <> numMaxLot.Tag Then
        moLE.GridEditChange numMaxLot
        Exit Sub
    End If
    numMaxLot.Tag = gsGetValidStr(numMaxLot.Value)
End Sub

Private Sub numMinLot_Change()    'LE Control Change Event
    If gsGetValidStr(numMinLot.Value) <> numMinLot.Tag Then
        moLE.GridEditChange numMinLot
        Exit Sub
    End If
    numMinLot.Tag = gsGetValidStr(numMinLot.Value)
End Sub

Private Sub numItemQty_Change()    'LE Control Change Event
'   If nbrQtyVariation.Value > 0 Then numItemQty.Value = nbrQtyVariation.Value
    
    
    If numItemQty.Value <> gdGetValidDbl(numItemQty.Tag) Then
        numItemQty.Tag = gsGetValidStr(numItemQty.Value)
        moLE.GridEditChange numItemQty
        CalcLineAmt
       ' moLE.GridEditChange numItemQty
        Exit Sub
    End If
    
'    If lkuItem.KeyValue <> 0 And numItemQty.Value < numMinLot.Value And numMinLot.Value <> 0 And numItemQty.Value <> 0 Then
'
'          MsgBox "La cantidad debe ser superior al Lote Mínimo.", vbExclamation, "Alerta"
'          bSetFocus nbrQtyVariation
'
'
'    End If
End Sub

Private Sub numRoundValue_Change()    'LE Control Change Event
    If gsGetValidStr(numRoundValue.Value) <> numRoundValue.Tag Then
        moLE.GridEditChange numRoundValue
        Exit Sub
    End If
    numRoundValue.Tag = gsGetValidStr(numRoundValue.Value)
End Sub


Private Sub currUnitCost_Change()    'LE Control Change Event
    If currUnitCost.Amount <> gdGetValidDbl(currUnitCost.Tag) Then
        currUnitCost.Tag = gsGetValidStr(currUnitCost.Amount)
        moLE.GridEditChange currUnitCost
        CalcLineAmt
        Exit Sub
    End If
End Sub
    'End of Sage MAS 500 Generated Code
    
Private Function IsValidContract() As Boolean
    On Error GoTo ErrorHandler
    Dim tlContractKey As Integer
    Dim tiRetVal As Integer
    
    IsValidContract = False
    
    txtContract.Text = sPadWithZero(10, txtContract)
    
    With moClass.moAppDB
        .SetInParam msCompanyID
        .SetInParam txtContract.Text
        .SetOutParam tlContractKey
        .SetOutParam tiRetVal
        .ExecuteSP ("spctIsValidContract")
        tiRetVal = .GetOutParam(4)
        tlContractKey = .GetOutParam(3)
        .ReleaseParams
    End With
    
    Select Case tiRetVal
        Case 0
            If tlContractKey > 0 Then
                mlContractKey = tlContractKey
                IsValidContract = True
                Exit Function
            End If
    End Select
    
    mlContractKey = 0
    MsgBox "Error Pendiente"
    
    Exit Function
ErrorHandler:
    MsgBox "IsValidContract()_" & Err.Description, vbCritical, "Error"
    IsValidContract = False
End Function

Private Function GetNextContractNo() As Boolean
    Dim tsContractNO As String
    Dim tiRetVal As Integer
    On Error GoTo ErrorHandler
    
    msContractNo = ""
    GetNextContractNo = False
    
    With moClass.moAppDB
        .SetInParam msCompanyID
        
        .SetOutParam tsContractNO
        .SetOutParam tiRetVal
        .ExecuteSP ("spctGetNextContractNo")
        tiRetVal = .GetOutParam(3)
        tsContractNO = .GetOutParam(2)
        .ReleaseParams
    End With
    
    Select Case tiRetVal
        Case 0
            msContractNo = tsContractNO
            GetNextContractNo = True
    End Select
    
    Exit Function
ErrorHandler:
    MsgBox "GetNextContractNo()_" & Err.Description, vbCritical, "Error"
    msContractNo = ""
    GetNextContractNo = False
End Function
Private Function GetNextContractNo1()
    Dim tsContractNO As String
    Dim tiRetVal As Integer
    On Error GoTo ErrorHandler
    
    msContractNo1 = ""
 '   GetNextContractNo1 = False
    
    With moClass.moAppDB
        .SetInParam msCompanyID
        .SetInParam sddType.Text
      
        .SetOutParam tsContractNO
        .SetOutParam tiRetVal
        .ExecuteSP ("spctGetNextContractNo1")
        'tiRetVal = .GetOutParam(3)
        tsContractNO = .GetOutParam(3)
        .ReleaseParams
    End With
    
'    Select Case tiRetVal
'        Case 0
            txtContractNo.Text = tsContractNO
      '      GetNextContractNo1 = True
'    End Select
    
    Exit Function
ErrorHandler:
    MsgBox "GetNextContractNo1()_" & Err.Description, vbCritical, "Error"
    msContractNo = ""
'    GetNextContractNo = False
End Function

Private Sub LoadDftlVendorValues()
    Dim lVendorKey As Long
    Dim lContactKey As Long
    Dim lVendClassKey As Long
    Dim lPaymentCond As Long
    Dim sContactID As String
    Dim sVendClassID As String
    Dim sCountry As String
    Dim sCurrID As String
    Dim sPT As String
    Dim sCE As String
    
    
    
    On Error GoTo ErrorHandler
    
    
    lVendorKey = lkuVendor.KeyValue
    
    lContactKey = glGetValidLong(moAppDB.Lookup("PrimaryCntctKey", "tapVendor", "VendKey = " & lVendorKey))
    If lContactKey > 0 Then
        sContactID = moAppDB.Lookup("Name", "tciContact", "CntctKey = " & lContactKey)
        lkuContact.SetTextAndKeyValue sContactID, lContactKey
    End If
    
   sCurrID = gsGetValidStr(moAppDB.Lookup(" a.CurrID ", "tapVendor v with (NOLOCK)  INNER JOIN tapVendAddr a with (NOLOCK)  ON a.VendKey = v.VendKey ", "v.VendKey = " & lVendorKey))
    If sCurrID <> "" Then
      
        SddCurrID.Text = sCurrID
    End If
    
     sPT = gsGetValidStr(moAppDB.Lookup(" t.PmtTermsID ", "tapVendor v with (NOLOCK)  INNER JOIN tapVendAddr a with (NOLOCK)  ON a.VendKey = v.VendKey INNER JOIN tciPaymentTerms t ON t.PmtTermsKey =  v.PmtTermsKey ", "v.VendKey = " & lVendorKey))
    If sPT <> "" Then
       ' sContactID = moAppDB.Lookup("Name", "tciContact", "CntctKey = " & lContactKey)
        sddPaymentTerms.Text = sPT
    End If
    
      sCE = gsGetValidStr(moAppDB.Lookup(" t.FOBID ", "tapVendor v with (NOLOCK)  INNER JOIN tapVendAddr a with (NOLOCK)  ON a.VendKey = v.VendKey INNER JOIN tciFOB t ON t.FOBKey =  a.FOBKey", "v.VendKey = " & lVendorKey))
    If sCE <> "" Then
       ' sContactID = moAppDB.Lookup("Name", "tciContact", "CntctKey = " & lContactKey)
        sddFOB.Text = sCE
    End If
    
    lVendClassKey = glGetValidLong(moAppDB.Lookup("VendClassKey", "tapVendor", "VendKey = " & lVendorKey))
    If lVendClassKey > 0 Then
        sVendClassID = gsGetValidStr(moAppDB.Lookup("VendClassID", "tapVendClass", "VendClassKey = " & lVendClassKey))
        lkuVendClass.SetTextAndKeyValue sVendClassID, lVendClassKey
        
        lPaymentCond = glGetValidLong(moAppDB.Lookup("PmtTermsKey", "tapVendClass", "VendClassKey = " & lVendClassKey))
        If lPaymentCond > 0 Then sddPaymentTerms.ItemData = lPaymentCond
       If sCurrID = "" Then
                sCurrID = gsGetValidStr(moAppDB.Lookup("isnull(CurrID, " & gsQuoted(msHomeCurrID) & ")", "tapVendClass", "VendClassKey = " & lVendClassKey))

       End If
        
        
    Else
       If sCurrID = "" Then
        
        sCurrID = msHomeCurrID
        
        End If
    End If
    
    
    
    sCountry = moAppDB.Lookup("s.CountryID", "tapVendor AS p JOIN tciAddress AS s ON s.AddrKey = p.PrimaryAddrKey", "p.VendKey = " & lVendorKey)
    txtCountryID.Caption = sCountry
    
    
    
    
    CalcFinishDate
    
    moDmHeader.SetColumnValue "SeqNo", 0
    SddCurrID.Text = sCurrID
    Exit Sub
ErrorHandler:
    MsgBox "LoadDftlVendorValues()_" & Err.Description, vbCritical, "Error"
End Sub



Private Function sPadWithZero(iLen As Integer, ctlMask As Control) As String
On Error GoTo ExpectedErrorRoutine
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
    
    Exit Function
    
ExpectedErrorRoutine:
    bIsANumber = False
    gClearSotaErr
    Resume Next

'+++ VB/Rig Begin Pop +++                                                                 'Repository Error Rig  {1.1.1.0.0}
#If ERRORTRAPON = 0 Then                                                                  'Repository Error Rig  {1.1.1.0.0}
    Err.Raise Err                                                                         'Repository Error Rig  {1.1.1.0.0}
#End If                                                                                   'Repository Error Rig  {1.1.1.0.0}
VBRigErrorRoutine:                                                                        'Repository Error Rig  {1.1.1.0.0}
        gSetSotaErr Err, "frmContract", "sPadWithZero", VBRIG_IS_FORM                     'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number                                                        'Repository Error Rig  {1.1.1.0.0}
'+++ VB/Rig End +++                                                                       'Repository Error Rig  {1.1.1.0.0}
End Function


Private Function lGetCurrExchSchdKey() As Long
'+++ VB/Rig Begin Push +++
#If ERRORTRAPON Then
On Error GoTo VBRigErrorRoutine
#End If
'+++ VB/Rig End +++
'**************************************************************
' Desc: Determine the appropriate Currency Exchange Schedule
'       Key to use on a new voucher based on existing values.
'**************************************************************
    
    'If PO has fixed exchange rate, or there is a receiver (rate is fixed in this case also).
    If mbFixedExchRate Then
        mlCurrExchSchdKey = 0
    Else
        If (moDmHeader.State = kDmStateAdd) And (mlCurrExchSchdKey = 0) Then
            '-- Use the Vendor Address Currency Exchange Schedule Key if possible
'            If (mlAddrCurrExchSchdKey > 0) Then
'                mlCurrExchSchdKey = mlAddrCurrExchSchdKey
'            Else
                '-- Use the MC Options Currency Exchange Schedule Key
                If (msHomeCurrID <> msNatCurrID) Then
                    mlCurrExchSchdKey = moOptions.MC("BuyExchSchdKey")
                End If
'            End If
        End If
    End If
    
    mlCurrExchSchdKeyOld = mlCurrExchSchdKey

    lGetCurrExchSchdKey = mlCurrExchSchdKey
'+++ VB/Rig Begin Pop +++
        Exit Function

VBRigErrorRoutine:
        gSetSotaErr Err, sMyName, "lGetCurrExchSchdKey", VBRIG_IS_FORM
        Select Case VBRIG_IS_NON_EVENT
        Case VBRIG_IS_NON_EVENT
                Err.Raise guSotaErr.Number
        Case Else
                Call giErrorHandler: Exit Function
        End Select
'+++ VB/Rig End +++
End Function

Private Function bSetupLkuNav() As Boolean
'+++ VB/Rig Begin Push +++                                                                'Repository Error Rig  {1.1.1.0.0}
#If ERRORTRAPON Then                                                                      'Repository Error Rig  {1.1.1.0.0}
    On Error GoTo VBRigErrorRoutine                                                       'Repository Error Rig  {1.1.1.0.0}
#End If                                                                                   'Repository Error Rig  {1.1.1.0.0}
'+++ VB/Rig End +++                                                                       'Repository Error Rig  {1.1.1.0.0}
    
    Dim sWhere As String

    bSetupLkuNav = False
    
    If Len(Trim(lkuNav.Tag)) = 0 Then
    
        sWhere = "CompanyID = " & gsQuoted(msCompanyID)
        
        bSetupLkuNav = gbLookupInit(lkuNav, moClass, moClass.moAppDB, "Contract", sWhere)

        
        If bSetupLkuNav Then
            lkuNav.Tag = 1
        End If
        
    Else
    
        bSetupLkuNav = True
        
    End If

'+++ VB/Rig Begin Pop +++                                                                 'Repository Error Rig  {1.1.1.0.0}
    Exit Function                                                                         'Repository Error Rig  {1.1.1.0.0}
VBRigErrorRoutine:                                                                        'Repository Error Rig  {1.1.1.0.0}
        gSetSotaErr Err, sMyName, "bSetupLkuNav", VBRIG_IS_FORM                    'Repository Error Rig  {1.1.1.0.0}
        Err.Raise guSotaErr.Number                                                        'Repository Error Rig  {1.1.1.0.0}
'+++ VB/Rig End +++                                                                       'Repository Error Rig  {1.1.1.0.0}
End Function

Private Sub UpdateUnitOfMeas()
    If lkuItem.KeyValue > 0 Then
        sddUM.SQLStatement = "SELECT s.UnitMeasID, s.UnitMeasKey FROM timItemUnitOfMeas AS p JOIN tciUnitMeasure AS s ON p.TargetUnitMeasKey = s.UnitMeasKey WHERE p.ItemKey = " & lkuItem.KeyValue
        sddUM.Refresh False
    End If
End Sub

Private Sub CalcLineAmt()
    If sddLineType.ListIndex > -1 Then
        If sddLineType.ListIndex > 1 Then
            currItemAmt.Amount = currUnitCost.Amount * nbrQtyVariation.Value
        Else
           If sddLineType.ItemData = kLineTypeModCos Then
             currItemAmt.Amount = currUnitCost.Amount * numItemQty.Value
            Else
              currItemAmt.Amount = currUnitCost.Amount * numItemQty.Value
              
            End If
            
          
        End If
    
     
    
    CalcContractAmt
    End If
    
    
      If sddLineType.ItemData = kLineTypeModCos Then
   
              '  numItemQty.Protected = False
                 numItemQty.Enabled = False
                 nbrQtyVariation.Enabled = False
                 
                 Else
                 numItemQty.Enabled = True
                  nbrQtyVariation.Enabled = True
   End If
End Sub

Private Sub CalcContractAmt()
    
    
    Dim i As Long
    Dim dContAmt As Double
    Dim dContAmtS As Double
    Dim sAprobalLevel As String
    
    dContAmt = 0
    
    For i = 0 To grdMain.DataRowCnt
        
        If glGetValidLong(gsGridReadCellText(grdMain, i, kColType)) = 4 Or glGetValidLong(gsGridReadCellText(grdMain, i, kColType)) = 2 Then
             
             If glGetValidLong(gsGridReadCellText(grdMain, i, kColType)) = 4 Then
                dContAmt = dContAmt - gdGetValidDbl(gsGridReadCellText(grdMain, i, kColLineAmt))
             Else
                dContAmt = dContAmt - gdGetValidDbl(moAppDB.Lookup("p.LineAmt", "tctContractLine AS p", "p.ContractKey = " & glGetValidLong(lkuParentContract.KeyValue) & " And p.ItemKey = " & glGetValidLong(gsGridReadCellText(grdMain, i, kColItemKey))))

             End If
             
             
             
        Else
             dContAmt = dContAmt + gdGetValidDbl(gsGridReadCellText(grdMain, i, kColLineAmt))
        End If
       
        
        
    Next i
    
      If sddType.ItemData = 3 Then
        
       ' dContAmtS = gdGetValidDbl(moAppDB.Lookup("p.ContractAmt", "tctContract AS p", "p.ContractKey = " & glGetValidLong(lkuParentContract.KeyValue)))
    
    End If
    
      '  dContAmt = dContAmt + dContAmtS
    
    
    'If dContAmt <> currAmt.Amount Then
    Dim dd As Double
    Dim dd1 As Double
    
    Dim ddAnt As Double
    
    Dim dd3 As Long
    Dim ente1 As Long
    Dim ente2 As Long
    'CORRECCION
    nFlete = 0
    nOtrosCostos = 0
    nSeguro = 0
    
    If snFlete.Text <> "" Then nFlete = Val(snFlete.Text)
    If snSeguro.Text <> "" Then nSeguro = Val(snSeguro.Text)
    If snOtrosCostos.Text <> "" Then nOtrosCostos = Val(snOtrosCostos.Text)
    
    dd = dContAmt + nOtrosCostos + nFlete + nSeguro
    
    currAmt.Text = CStr(dd)
    

    If dd > 100000000000000# Then
        bGrand = True
    Else
        bGrand = False
    End If
    currAmt.Text = CStr(dd)
    
'    sAprobalLevel = gsGetValidStr(moAppDB.Lookup("p.AprobalLevelID", "tctAprobalLevel AS p", "p.StartAmt <= " & dContAmt & " AND p.EndAmt >=" & dContAmt))
'    If Len(Trim$(sAprobalLevel)) = 0 And dContAmt > 0 Then
'        sAprobalLevel = "No definido"
'    End If
    txtAprobalLevel.Caption = sAprobalLevel
    currAmt_Validate True
    
    
End Sub

Public Function HexToString(ByVal HexToStr As String) As String

Dim strTemp As String
Dim strReturn As String
Dim i As Long

  For i = 1 To Len(HexToStr) Step 3
   strTemp = Chr$(Val("&H" & Mid$(HexToStr, i, 2)))
   strReturn = strReturn & strTemp
  Next i
  HexToString = strReturn

End Function


Private Sub CalcFinishDate()
    If Len(Trim$(sclStartDate)) > 0 Then
        txtOutDate.Caption = DateTime.DateAdd("YYYY", gdGetValidDbl(nbrDuration.Value), sclStartDate.Value)
    End If
End Sub

Private Sub CleanDetails()
    lkuItem.Tag = ""
    txtDescription.Tag = ""
    lkuItem.SetTextAndKeyValue "", ""
    txtDescription.Text = ""
    numItemQty.Value = 0
    currUnitCost.Amount = 0
    numMaxLot.Value = 0
    numMinLot.Value = 0
    numDeliveryTime.Value = 0
    numRoundValue.Value = 0
    sddUM.Clear
    sddLineType.ListIndex = -1
    nbrQtyVariation.Value = 0
    txtAccountID.Text = ""
    txtSWIFT.Text = ""
    txtTitular.Text = ""
    txtBankAddress.Text = ""
End Sub

Private Function GetNextSeqNo() As Integer
    Dim iSeqNo As Integer
    iSeqNo = giGetValidInt(moDmHeader.GetColumnValue("SeqNo")) + 1
    GetNextSeqNo = iSeqNo
    moDmHeader.SetColumnValue "SeqNo", iSeqNo
End Function

Public Function bSetFocus(oControl As Object) As Boolean
'+++ VB/Rig Skip +++

On Error GoTo ExpectedErrorRoutine

DoEvents
    If Not (oControl Is Nothing) Then
        
        If oControl.Enabled = True Then
            
            If oControl.Visible = True Then
                
                On Error Resume Next
                If tabDataEntry <> 0 Then tabDataEntry.Tab = 0
                oControl.SetFocus
                On Error GoTo ExpectedErrorRoutine
                
            End If
        
        End If
    
    End If
    
    
    
   
    bSetFocus = True
    Exit Function
ExpectedErrorRoutine:
    MsgBox "bSetFocus_" & Err.Description, vbCritical, "Error"
End Function

Private Function HavePendingOperations() As Boolean
    On Error GoTo ExpectedErrorRoutine
    HavePendingOperations = True
    If giGetValidInt(moAppDB.Lookup("COUNT(*)", "tpoPurchOrder AS p", "p.[Status] IN (0,1) and  p.ContractKey =" & moDmHeader.GetColumnValue("ContractKey"))) > 0 Then Exit Function
    If giGetValidInt(moAppDB.Lookup("COUNT(*)", "tpoRequisition AS p JOIN tpoRequisitionContract AS s ON s.ReqKey = p.ReqKey JOIN tpoReqLine AS t ON t.ReqKey = s.ReqKey", "t.POLineKey IS NULL AND s.ContractKey =" & moDmHeader.GetColumnValue("ContractKey"))) > 0 Then Exit Function
    HavePendingOperations = False
    Exit Function
ExpectedErrorRoutine:
    HavePendingOperations = False
    MsgBox "HavePendingOperations_" & Err.Description, vbCritical, "Error"
End Function

Private Function ValidateFields() As Boolean

Dim bExist As Boolean

    ValidateFields = False
    
    If sddPaymentTerms.ListIndex = -1 Then
        MsgBox "Debe seleccionar la condición de pago.", vbInformation, "Alerta"
        bSetFocus sddPaymentTerms
        Exit Function
    End If
    
    If DateTime.DateDiff("d", sclStartDate, sclSignatureDate) > 0 Then
        MsgBox "La fecha de inicio no puede ser anterior a la fecha de firma.", vbInformation, "Alerta"
        bSetFocus sclStartDate
        Exit Function
    End If
    
    If nbrDuration.Value = 0 Then
        MsgBox "La vigencia del contrato debe ser mayor que 0.", vbInformation, "Alerta"
        bSetFocus nbrDuration
        Exit Function
    End If
    
    If Len(Trim$(lkuContact.Text)) = 0 Then
        MsgBox "Debe seleccionar un contacto.", vbInformation, "Alerta"
        Exit Function
    End If
    
    If Len(Trim$(lkuVendClass.Text)) = 0 Then
        MsgBox "Debe seleccionar una clase de proveedor.", vbInformation, "Alerta"
        bSetFocus lkuVendClass
        Exit Function
    End If
    
    If lkuParentContract.Visible Then
        If Len(Trim$(lkuParentContract.Text)) = 0 Then
            MsgBox "Debe seleccionar un contrato al cual asociar este suplemento.", vbInformation, "Alerta"
            bSetFocus lkuParentContract
            Exit Function
        End If
    End If
    
     If Len(Trim$(txtContractNo.Text)) = 0 Then
        MsgBox "Debe introducir un No de Contrato.", vbInformation, "Alerta"
        bSetFocus txtContractNo
        Exit Function
    End If
       
      If lkuParentContract.Text <> "" Then
         bExist = gbGetValidBoolean(moAppDB.Lookup("1", "tctContract AS p", "[TYPE] = 2 AND ContractKey in (Select ContractKey from tctContract where ContractNo = " & gsQuoted(lkuParentContract.Text) & ")"))
      End If
    
    If Len(Trim$(sddFOB.Text)) = 0 And sddType.ItemData <> 2 And bExist = False And sddClasification.ItemData <> 2 Then
        MsgBox "Debe seleccionar una condición de entrega.", vbInformation, "Alerta"
        bSetFocus SddCurrID
        Exit Function
    End If
       
    If Len(Trim$(SddCurrID.Text)) = 0 Then
        MsgBox "Debe seleccionar una moneda.", vbInformation, "Alerta"
        bSetFocus SddCurrID
        Exit Function
    End If
    
    If currAmt.Text <> "" Then
        If CDbl(currAmt.Text) <= 0 Then
            If sddType.ItemData = 1 Then
                MsgBox "El contrato no puede tener saldo 0.", vbInformation, "Alerta"
                tabDataEntry.Tab = 1
                Exit Function
            End If
        End If
    End If
    
    If Len(Trim$(currAmt.Text)) > 38 Then
        MsgBox "El importe del Contrato no puede exceder las 38 cifras.", vbInformation, "Alerta"
       
        Exit Function
    End If
    
     If Len(Trim$(snFlete.Text)) > 38 Then
        MsgBox "El Flete del Contrato no puede exceder las 38 cifras.", vbInformation, "Alerta"
        bSetFocus snFlete
        Exit Function
    End If
    
     If Len(Trim$(snSeguro.Text)) > 38 Then
        MsgBox "El Seguro del Contrato no puede exceder las 38 cifras.", vbInformation, "Alerta"
        bSetFocus snSeguro
        Exit Function
    End If
    
     If Len(Trim$(snOtrosCostos.Text)) > 38 Then
        MsgBox "Los Otros Costos del Contrato no puede exceder las 38 cifras.", vbInformation, "Alerta"
        bSetFocus snOtrosCostos
        Exit Function
    End If
    
    ValidateFields = True
End Function

Private Sub EnableFields()
    lkuVendor.Enabled = True
    sddType.Enabled = True
    txtContractNo.Enabled = True
    sddPaymentTerms.Enabled = True
    If sddType.ItemData = kContractTypeSuplement Then
        lkuParentContract.Visible = True
        lblParentContract.Visible = True
        
        If Len(Trim$(lkuParentContract.Text)) > 0 Then
            lkuVendor.Enabled = False
        Else
            lkuVendClass.Enabled = True
        End If
    Else
        lkuParentContract.Visible = False
        lblParentContract.Visible = False
        lkuVendClass.Enabled = True
    End If
    sclSignatureDate.Enabled = True
    sclStartDate.Enabled = True
    nbrDuration.Enabled = True
    lkuContact.Enabled = True
    lkuVendClass.Enabled = True
    sddFOB.Enabled = True
    SddCurrID.Enabled = True
    
    lkuItem.Enabled = True
    txtDescription.Enabled = True
    numDeliveryTime.Enabled = True
    numRoundValue.Enabled = True
    sddUM.Enabled = True
    currUnitCost.Enabled = True
    numItemQty.Enabled = True
    numMinLot.Enabled = True
    numMaxLot.Enabled = True
    sddLineType.Enabled = True
    nbrQtyVariation.Enabled = True
    
    sddState.Enabled = True
    snOtrosCostos.Enabled = True
    snFlete.Enabled = True
    snSeguro.Enabled = True
    
    txtDescription.Enabled = True
    sddClasification.Enabled = True
    
     
    
    
   grdMain.Enabled = True
End Sub

Private Sub DisableFields()
    lkuVendor.Enabled = False
    sddType.Enabled = False
    txtContractNo.Enabled = False
    sddPaymentTerms.Enabled = False
    lkuParentContract.Enabled = False
    sclSignatureDate.Enabled = False
    sclStartDate.Enabled = False
    nbrDuration.Enabled = False
    lkuContact.Enabled = False
    lkuVendClass.Enabled = False
    sddFOB.Enabled = False
    SddCurrID.Enabled = False
    
    lkuItem.Enabled = False
    txtDescription.Enabled = False
    numDeliveryTime.Enabled = False
    numRoundValue.Enabled = False
    sddUM.Enabled = False
    currUnitCost.Enabled = False
    numItemQty.Enabled = False
    numMinLot.Enabled = False
    numMaxLot.Enabled = False
    sddLineType.Enabled = False
    nbrQtyVariation.Enabled = False
'    grdMain.Enabled = True
    
    sddState.Enabled = False
    snOtrosCostos.Enabled = False
    snFlete.Enabled = False
    snSeguro.Enabled = False
    txtDescription.Enabled = False
    sddClasification.Enabled = False
    
    
     
End Sub


Private Function bLoadParentContratDftl(lParentContractKey As Long) As Boolean
    Dim sVendID As String
    Dim lVendKey As Long
   
    Dim lPaymentTermsKey As Long
    Dim lFOBKey As Long
    Dim sVendClass As String
    Dim lVendClassKey As Long
    Dim sContactID As String
    Dim lContactKey As Long
    
    Dim sCurrID As String
    Dim sCountry As String
    
    
    
    On Error GoTo ErrorHandler
    
    If Len(Trim$(lkuParentContract.Text)) Then
        If lkuParentContract.Tag <> lkuParentContract.Text Then
            
            sVendID = gsGetValidStr(moAppDB.Lookup("VendID", "tctContract as p join tapVendor as s on p.VendorKey = s.VendKey", "p.ContractKey =" & lParentContractKey))
            lVendKey = glGetValidLong(moAppDB.Lookup("p.VendorKey", "tctContract as p join tapVendor as s on p.VendorKey = s.VendKey", "p.ContractKey =" & lParentContractKey))
            
            lVendClassKey = glGetValidLong(moAppDB.Lookup("p.VendClassKey", "tctContract as p join tapVendClass as s on p.VendClassKey = s.VendClassKey", "p.ContractKey =" & lParentContractKey))
            sVendClass = gsGetValidStr(moAppDB.Lookup("VendClassID", "tctContract as p join tapVendClass as s on p.VendClassKey = s.VendClassKey", "p.ContractKey =" & lParentContractKey))
            
            lContactKey = glGetValidLong(moAppDB.Lookup("p.CntctKey", "tctContract as p join tciContact as s on p.CntctKey = s.CntctKey", "p.ContractKey =" & lParentContractKey))
            sContactID = gsGetValidStr(moAppDB.Lookup("s.Name", "tctContract as p join tciContact as s on p.CntctKey = s.CntctKey", "p.ContractKey =" & lParentContractKey))
            
            lPaymentTermsKey = glGetValidLong(moAppDB.Lookup("p.PmtTermsKey", "tctContract as p join tciPaymentTerms as s on p.PmtTermsKey = s.PmtTermsKey", "p.ContractKey =" & lParentContractKey))
            
            lFOBKey = glGetValidLong(moAppDB.Lookup("s.FOBKey", "tctContract as p join tciFOB as s on p.FOBKey = s.FOBKey", "p.ContractKey =" & lParentContractKey))
            
            sCountry = gsGetValidStr(moAppDB.Lookup("p.CountryID", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
            sCurrID = gsGetValidStr(moAppDB.Lookup("p.CurrID", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
            
            ssTitular = gsGetValidStr(moAppDB.Lookup("p.Titular", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
            ssAccountNo = gsGetValidStr(moAppDB.Lookup("p.AccountID", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
            ssSWIFT = gsGetValidStr(moAppDB.Lookup("p.SWIFT", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
            ssBankAddress = gsGetValidStr(moAppDB.Lookup("p.BankAddress", "tctContract as p ", "p.ContractKey =" & lParentContractKey))
           
            
            
            If lVendKey <> 0 Then lkuVendor.SetTextAndKeyValue sVendID, lVendKey
            If sVendID <> "" Then lkuVendor.Tag = sVendID
            If sVendClass <> "" Then lkuVendClass.Text = sVendClass
            If lVendClassKey <> 0 Then lkuVendClass.KeyValue = lVendClassKey
            If lContactKey <> 0 Then lkuContact.SetTextAndKeyValue sContactID, lContactKey
            
            If lPaymentTermsKey <> 0 Then sddPaymentTerms.ItemData = lPaymentTermsKey
            If lFOBKey <> 0 Then sddFOB.ItemData = lFOBKey
           If sCurrID <> "" Then SddCurrID.Text = sCurrID
           If sCountry <> "" Then txtCountryID.Caption = sCountry
           
           If ssTitular <> "" Then txtTitular.Text = ssTitular
           If ssAccountNo <> "" Then txtAccountID.Text = ssAccountNo
           If ssSWIFT <> "" Then txtSWIFT.Text = ssSWIFT
           If ssBankAddress <> "" Then txtBankAddress.Text = ssBankAddress
           
            
        End If
    End If
    Exit Function
ErrorHandler:
    MsgBox Err.Description
'    gSetSotaErr Err, Me.Name, "bLoadParentContratDftl", VBRIG_IS_FORM                    'Repository Error Rig  {1.1.1.0.0}
'        Err.Raise guSotaErr.Number
End Function

Private Function sGetParentRestrict() As String
    sGetParentRestrict = "Type in (1,2) and CompanyID = " & gsQuoted(msCompanyID) '& " and VendorKey =" & lkuVendor.KeyValue
    If Len(Trim$(lkuVendor.Text)) > 0 Then
        sGetParentRestrict = sGetParentRestrict & " and VendorKey = " & lkuVendor.KeyValue
    End If
End Function

Private Sub LoadOperations()
     On Error GoTo ErrorHandler
     moAppDB.ExecuteSQL "delete from #tctContractOperations"
    
    If moDmHeader.State = kDmStateEdit Then
        moAppDB.ExecuteSQL "insert into #tctContractOperations SELECT s.VoucherLineKey, fv.TranID, p.TranID, t.Quantity, s.[Description], s.ExtAmt, p.PostDate" _
                        & " FROM tapVoucher AS p JOIN tapVoucherDetl AS s ON s.VoucherKey = p.VoucherKey " _
                        & " JOIN tapVoucherLineDist AS t ON t.VoucherLineKey = s.VoucherLineKey " _
                        & " JOIN tpoPOLine AS f ON f.POLineKey = s.POLineKey " _
                        & " JOIN tpoPurchOrder AS fv ON fv.POKey = f.POKey WHERE fv.ContractKey =" & moDmHeader.GetColumnValue("ContractKey") _
                        & " and f.ItemKey in (SELECT st.ItemKey FROM tctContractLine AS st  WHERE st.ContractKey = " & moDmHeader.GetColumnValue("ContractKey") & ")"
    
        moAppDB.ExecuteSQL "insert into #tctContractOperations SELECT s.VoucherLineKey, fv.TranID, p.TranID, t.Quantity, s.[Description], s.ExtAmt, p.PostDate" _
                        & " FROM tapVoucher AS p JOIN tapVoucherDetl AS s ON s.VoucherKey = p.VoucherKey " _
                        & " JOIN tapVoucherLineDist AS t ON t.VoucherLineKey = s.VoucherLineKey " _
                        & " JOIN tpoPOLine AS f ON f.POLineKey = s.POLineKey " _
                        & " JOIN tpoPurchOrder AS fv ON fv.POKey = f.POKey " _
                        & " join tctContract as sx on sx.ContractKey = fv.ContractKey" _
                        & " WHERE sx.ParentContractKey =" & moDmHeader.GetColumnValue("ContractKey") _
                        & " and f.ItemKey in (SELECT st.ItemKey FROM tctContractLine AS st  WHERE st.ContractKey = " & moDmHeader.GetColumnValue("ContractKey") & ")"
    
    End If
    
    moDmOperations.Init
    gGridDeleteRow grdOperations, grdOperations.DataRowCnt + 1
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Function bIsNumberInUse() As Boolean
    On Error GoTo ErrorHandler
    bIsNumberInUse = True
    If sddType.ItemData = 1 Then
        If giGetValidInt(moAppDB.Lookup("count(*)", "tctContract", "ContractNo = " & gsQuoted(txtContractNo) & " and ContractKey <> " & moDmHeader.GetColumnValue("ContractKey"))) > 0 Then
            MsgBox "Este número de contrato ya esta en uso", vbInformation, "Alerta"
            Exit Function
        End If
    ElseIf sddType.ItemData = 3 Then
        If Len(Trim$(lkuParentContract.Text)) = 0 Then
            MsgBox "Debe seleccionar primeramente el contrato al que se asociara este suplemento", vbInformation, "Alerta"
            Exit Function
        End If
        If giGetValidInt(moAppDB.Lookup("count(*)", "tctContract", "ContractNo = " & gsQuoted(txtContractNo) & " and ParentContractKey =" & lkuParentContract.KeyValue & " and ContractKey <> " & moDmHeader.GetColumnValue("ContractKey"))) > 0 Then
            MsgBox "El número de suplemento ya esta en uso", vbInformation, "Alerta"
           ' lkuParentContract_LostFocus
            Exit Function
        End If
    End If
    bIsNumberInUse = False
    Exit Function
ErrorHandler:
    bIsNumberInUse = True
    MsgBox "bIsNumberInUse_" & Err.Description, vbCritical, "Error"
End Function
