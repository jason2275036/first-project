VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Object = "{3ED2D85D-59DA-11D1-8168-006097C8673E}#1.1#0"; "NEdbox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form DWE2532B 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BARCODE LABEL 출력(PEAK)"
   ClientHeight    =   7350
   ClientLeft      =   8355
   ClientTop       =   3345
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Ebrima"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7350
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox D_PROD_KB 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9855
      TabIndex        =   63
      Top             =   3600
      Visible         =   0   'False
      Width           =   990
   End
   Begin Threed.SSPanel Title 
      Height          =   555
      Left            =   555
      TabIndex        =   24
      Top             =   0
      Width           =   8040
      _Version        =   65536
      _ExtentX        =   14182
      _ExtentY        =   979
      _StockProps     =   15
      Caption         =   "BARCODE LABEL 출력 (PEAK)"
      ForeColor       =   12648447
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel Key_Panel 
      Height          =   5715
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   9165
      _Version        =   65536
      _ExtentX        =   16166
      _ExtentY        =   10081
      _StockProps     =   15
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ebrima"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      Begin VB.Frame Frame_Ship 
         Caption         =   "Shipment"
         Height          =   735
         Left            =   1320
         TabIndex        =   79
         Top             =   4920
         Width           =   3135
         Begin VB.OptionButton Option3 
            Caption         =   "S+G"
            Height          =   255
            Left            =   2160
            TabIndex        =   82
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Peak"
            Height          =   255
            Left            =   1200
            TabIndex        =   81
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "FRC"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CheckBox PEAK_Ship_Check 
         Caption         =   "PEAK Shipment"
         Height          =   255
         Left            =   1920
         TabIndex        =   78
         Top             =   5400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox EOL_Print_Check 
         Caption         =   "EOL Label Print"
         Height          =   255
         Left            =   4800
         TabIndex        =   77
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox D_SWWCUST_CD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "D_SWWCUST_CD"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox D_SCUSTOM_CD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "D_SCUSTOM_CD"
         Top             =   5040
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox D_ENDLABEL_KB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "D_ENDLABEL_KB"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Timer Timer1 
         Left            =   8640
         Top             =   120
      End
      Begin VB.TextBox D_CUSTOM_NM 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   50
         Width           =   3400
      End
      Begin VB.Frame FrameForm 
         Caption         =   "출력양식"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   660
         Left            =   4780
         TabIndex        =   58
         Top             =   3480
         Width           =   4280
         Begin VB.OptionButton PRT_FORM 
            Caption         =   "C/T Box"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   260
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton PRT_FORM 
            Caption         =   "Inner Label"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   1320
            TabIndex        =   59
            Top             =   260
            Width           =   1335
         End
         Begin NEdBox.NumEdBox3D I_BUNDLE_QT 
            Height          =   315
            Left            =   3800
            TabIndex        =   61
            Top             =   200
            Visible         =   0   'False
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            Text            =   ""
            FormatS         =   ""
            MaxLength       =   0
            BackColor       =   12648447
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bundle Qty :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   24
            Left            =   2670
            TabIndex        =   62
            Top             =   260
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin Threed.SSPanel ProgPnl 
         Height          =   735
         Left            =   2880
         TabIndex        =   56
         Top             =   1080
         Visible         =   0   'False
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Data Gathering"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   4
         BevelInner      =   1
         Alignment       =   6
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   300
            Left            =   135
            TabIndex        =   57
            Top             =   315
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin VB.ComboBox C_LOT_ID 
         Height          =   315
         ItemData        =   "DWE2532B.frx":0000
         Left            =   6120
         List            =   "DWE2532B.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox C_ITEM_NO 
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1365
         TabIndex        =   11
         Top             =   2805
         Width           =   3270
      End
      Begin VB.ComboBox C_LOT_NO 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1380
         TabIndex        =   4
         Tag             =   "Y"
         Text            =   "C_LOT_NO"
         Top             =   1580
         Width           =   3255
      End
      Begin VB.CheckBox CT_PRINT_CHK 
         Alignment       =   1  'Right Justify
         Caption         =   "C/T # Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7560
         TabIndex        =   2
         Top             =   4080
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.ComboBox C_INV_NO 
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6120
         TabIndex        =   6
         Text            =   "C_INV_NO"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.ComboBox C_PO_NO 
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1380
         TabIndex        =   5
         Text            =   "C_PO_NO"
         Top             =   2040
         Width           =   3255
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   1320
         Picture         =   "DWE2532B.frx":0004
         ScaleHeight     =   435
         ScaleWidth      =   3075
         TabIndex        =   39
         Top             =   4395
         Width           =   3135
      End
      Begin VB.Frame FramePRT 
         Caption         =   "출력구분"
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   660
         Left            =   4780
         TabIndex        =   38
         Top             =   2805
         Width           =   4215
         Begin VB.CheckBox CHK_INSMP 
            Caption         =   "SMP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   2070
            TabIndex        =   53
            Top             =   240
            Width           =   750
         End
         Begin NEdBox.NumEdBox3D I_SET_QT 
            Height          =   330
            Left            =   3615
            TabIndex        =   52
            Top             =   210
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   582
            Text            =   ""
            FormatS         =   ""
            MaxLength       =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Ebrima"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.OptionButton PRT_OP 
            Caption         =   "국외"
            BeginProperty Font 
               Name            =   "Ebrima"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1050
            TabIndex        =   1
            Top             =   285
            Width           =   885
         End
         Begin VB.OptionButton PRT_OP 
            Caption         =   "국내"
            BeginProperty Font 
               Name            =   "Ebrima"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   0
            Top             =   270
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Set Qty:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   19
            Left            =   2955
            TabIndex        =   51
            Top             =   250
            Width           =   555
         End
      End
      Begin VB.TextBox I_COLOR_NM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7980
         TabIndex        =   10
         Top             =   2430
         Width           =   990
      End
      Begin VB.ComboBox C_CUSTOM_CD 
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Y"
         Top             =   1215
         Width           =   7620
      End
      Begin NEdBox.NumEdBox3D I_TOTAL_CT_NO 
         Height          =   315
         Left            =   3960
         TabIndex        =   15
         Top             =   3615
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Text            =   ""
         FormatS         =   "###"
         MaxLength       =   3
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox I_REV_NO 
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4095
         MaxLength       =   3
         TabIndex        =   8
         Top             =   2430
         Width           =   525
      End
      Begin Threed.SSPanel D_BARCODE_TX 
         Height          =   150
         Left            =   1320
         TabIndex        =   41
         Top             =   4920
         Width           =   3105
         _Version        =   65536
         _ExtentX        =   5477
         _ExtentY        =   265
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   -2147483637
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ebrima"
            Size            =   9.02
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin VB.TextBox I_METERIAL_NM 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6135
         TabIndex        =   9
         Top             =   2430
         Width           =   915
      End
      Begin NEdBox.NumEdBox3D I_BOXUNIT_NB 
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   3225
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         Text            =   "50"
         FormatS         =   "######"
         MaxLength       =   6
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox  I_PART_NO 
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2400
         Width           =   1980
      End
      Begin VB.TextBox I_DESC_NM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1380
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   435
         Width           =   7590
      End
      Begin NEdBox.NumEdBox3D I_START_CT_NO 
         Height          =   315
         Left            =   1365
         TabIndex        =   14
         Top             =   3615
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         Text            =   "1"
         FormatS         =   "###"
         MaxLength       =   3
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin fpBtnAtlLibCtl.fpBtn CMD_SAVE 
         Height          =   495
         Left            =   7245
         TabIndex        =   19
         Top             =   5040
         Width           =   1860
         _Version        =   131072
         _ExtentX        =   3281
         _ExtentY        =   873
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "DWE2532B.frx":4B5A
      End
      Begin Threed.SSPanel D_TOT_LABEL_CNT 
         Height          =   330
         Left            =   5880
         TabIndex        =   44
         Top             =   5205
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   582
         _StockProps     =   15
         ForeColor       =   0
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin NEdBox.NumEdBox3D I_END_CT_NO 
         Height          =   315
         Left            =   2490
         TabIndex        =   16
         Top             =   3615
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Text            =   ""
         FormatS         =   "###"
         MaxLength       =   3
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin NEdBox.NumEdBox3D I_TOTAL_QT 
         Height          =   315
         Left            =   1365
         TabIndex        =   12
         Top             =   3210
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Text            =   ""
         FormatS         =   "######"
         MaxLength       =   6
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin NEdBox.NumEdBox3D I_TOTAL_WEIGHT 
         Height          =   315
         Left            =   1365
         TabIndex        =   17
         Top             =   4005
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Text            =   ""
         FormatS         =   "######.##"
         MaxLength       =   9
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin NEdBox.NumEdBox3D I_PCNT_NO 
         Height          =   315
         Left            =   3960
         TabIndex        =   18
         Top             =   4005
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Text            =   ""
         FormatS         =   "###"
         MaxLength       =   3
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Ebrima"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker SHIP_DATE 
         Height          =   345
         Left            =   1440
         TabIndex        =   65
         Top             =   0
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   140771329
         CurrentDate     =   41600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ship To:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   20
         Left            =   4800
         TabIndex        =   69
         Top             =   45
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ship Date :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   21
         Left            =   120
         TabIndex        =   64
         Top             =   80
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LOT ID:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   4800
         TabIndex        =   55
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EA"
         BeginProperty Font 
            Name            =   "Ebrima"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   6840
         TabIndex        =   50
         Top             =   5280
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Package Count :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   2715
         TabIndex        =   49
         Top             =   4065
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Weight :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   315
         TabIndex        =   48
         Top             =   4080
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Q'TY:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   435
         TabIndex        =   47
         Top             =   3270
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Invoice NO :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   4800
         TabIndex        =   46
         Top             =   2100
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   7380
         TabIndex        =   35
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CUSTOMER :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "~"
         Height          =   180
         Index           =   13
         Left            =   2190
         TabIndex        =   45
         Top             =   3660
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "LABEL QTY :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   17
         Left            =   4650
         TabIndex        =   40
         Top             =   5175
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "REV :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   3540
         TabIndex        =   33
         Top             =   2490
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C/T:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   945
         TabIndex        =   43
         Top             =   3660
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "of"
         Height          =   195
         Index           =   14
         Left            =   3480
         TabIndex        =   42
         Top             =   3660
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   5205
         TabIndex        =   34
         Top             =   2400
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Box Unit Q'ty:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   2880
         TabIndex        =   37
         Top             =   3285
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LOT# :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   330
         TabIndex        =   30
         Top             =   1635
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PO# :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   870
         TabIndex        =   31
         Top             =   2085
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PART# :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   645
         TabIndex        =   32
         Top             =   2490
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ITEM# :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   660
         TabIndex        =   36
         Top             =   2880
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ITEM :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   330
         TabIndex        =   28
         Top             =   495
         Width           =   900
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn CMD_EXIT 
      Height          =   555
      Left            =   8625
      TabIndex        =   22
      Top             =   0
      Width           =   555
      _Version        =   131072
      _ExtentX        =   979
      _ExtentY        =   979
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "DWE2532B.frx":4D40
   End
   Begin fpBtnAtlLibCtl.fpBtn CMD_HELP 
      Height          =   555
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   555
      _Version        =   131072
      _ExtentX        =   979
      _ExtentY        =   979
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "DWE2532B.frx":70F5
   End
   Begin Threed.SSPanel Msg 
      Height          =   345
      Left            =   0
      TabIndex        =   26
      Top             =   6960
      Width           =   9180
      _Version        =   65536
      _ExtentX        =   16192
      _ExtentY        =   609
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ebrima"
         Size            =   9.76
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      FloodColor      =   255
      FloodShowPct    =   0   'False
      Alignment       =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   585
      Left            =   -15
      TabIndex        =   27
      Top             =   6360
      Width           =   9180
      _Version        =   65536
      _ExtentX        =   16192
      _ExtentY        =   1032
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ebrima"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox ChkLANPRT 
         Caption         =   "LAN (ITC)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   400
         Left            =   5760
         TabIndex        =   73
         Top             =   120
         Width           =   720
      End
      Begin VB.Frame Frm_Port 
         Caption         =   "Printer Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   550
         Left            =   2160
         TabIndex        =   66
         Top             =   0
         Width           =   3600
         Begin VB.OptionButton PRT_SET 
            Caption         =   "MAIN(COM1)"
            BeginProperty Font 
               Name            =   "Ebrima"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Value           =   -1  'True
            Width           =   1300
         End
         Begin VB.OptionButton PRT_SET 
            Caption         =   "SUB(COM3)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   2
            Left            =   1455
            TabIndex        =   75
            Top             =   240
            Width           =   1250
         End
         Begin VB.OptionButton PRT_SET 
            Caption         =   "LPT"
            BeginProperty Font 
               Name            =   "Ebrima"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Index           =   0
            Left            =   2745
            TabIndex        =   67
            Top             =   210
            Width           =   795
         End
      End
      Begin fpBtnAtlLibCtl.fpBtn CMD_CLEAR 
         Height          =   495
         Left            =   840
         TabIndex        =   21
         Top             =   45
         Width           =   1245
         _Version        =   131072
         _ExtentX        =   2196
         _ExtentY        =   873
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "DWE2532B.frx":83B9
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   -15
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   -1  'True
      End
      Begin MSWinsockLib.Winsock Lantype_Sck 
         Left            =   600
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin fpBtnAtlLibCtl.fpBtn CMD_PRINT 
         Height          =   495
         Left            =   6600
         TabIndex        =   74
         Top             =   0
         Width           =   2610
         _Version        =   131072
         _ExtentX        =   4604
         _ExtentY        =   873
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         GrayAreaColor   =   12632256
         BorderShowDefault=   -1  'True
         ButtonType      =   0
         NoPointerFocus  =   0   'False
         Value           =   0   'False
         GroupID         =   0
         GroupSelect     =   0
         DrawFocusRect   =   2
         DrawFocusRectCell=   -1
         GrayAreaPictureStyle=   0
         Static          =   0   'False
         BackStyle       =   1
         AutoSize        =   0
         AutoSizeOffsetTop=   0
         AutoSizeOffsetBottom=   0
         AutoSizeOffsetLeft=   0
         AutoSizeOffsetRight=   0
         DropShadowOffsetX=   3
         DropShadowOffsetY=   3
         DropShadowType  =   0
         DropShadowColor =   0
         Redraw          =   -1  'True
         ButtonDesigner  =   "DWE2532B.frx":9687
      End
   End
End
Attribute VB_Name = "DWE2532B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
' Module Name : Print Barcode Lable
' Module ID   : DWE2532B
' Creater     :
' Create Date :
' Contents    : Print Barcode Lable
'               BAR Code Type : Bar code A(Code3of9)
'               This bar code consists of upper-case alphanumeric characters.
'               Columns are variable in length
'               Start/stop codes "*" are given automatically by printer.
'               BAR Code Printer Model : T-4 manufactured ITOCHU ELECTRONICS CORP.
'------------------------------------------------------------
' Modify Date :
' Modifier    :
' Contents    :
'============================================================

'====================================================================================================================================
' 업체별로 Barcode에 Display되는 내용을 다르게 요청하므로
' Barcode에 뿌려지는 내용을 수정
'
'====================================================================================================================================
Option Explicit

Dim Label_Cnt           As Long
Dim Label_Data()        As String

Dim Field_Check_SW      As Integer
Dim Column_Cancel       As Integer
Dim Field_Name()        As String
Dim Field_Check()       As Integer
Dim Field_Object()      As Control

Dim New_index           As Integer
Dim New_Focus           As Control

Dim sFormatTx           As String
Dim SSDATA              As New ADODB.Recordset
Dim SSCODE              As New ADODB.Recordset

Dim varWWCustomCD       As String
Dim varBranchCD             As String
Dim varEndwwc_CD          As String
Dim isAgentPeak           As Boolean
Dim strLotNo            As String
Dim Result_QT           As Long
Dim Box_QT              As Long

Dim varMoldNo           As String

Dim PnlMsg              As String

Private Sub Form_Control_Search()
    Dim i As Integer
    Dim j As Integer

    ReDim Field_Name(Me.count)
    ReDim Field_Check(Me.count)
    ReDim Field_Object(Me.count)

    For i = 0 To Me.count - 1
        If UCase(Mid(Me.Controls(i).Name, 1, 2)) = "I_" Or _
           UCase(Mid(Me.Controls(i).Name, 1, 2)) = "D_" Or _
           UCase(Me.Controls(i).Name) = "CMD_SAVE" Then
                j = Me.Controls(i).TabIndex
                Field_Name(j) = Me.Controls(i).Name
                Set Field_Object(j) = Me.Controls(i)
        End If
    Next

End Sub

Private Sub FIELD_VERIFY(Optional GOT_FOCUS As Control)
    Dim I_idx           As Integer
    
    Dim sql             As String
    Dim strMsg          As String
    
    Set New_Focus = Screen.ActiveForm.ActiveControl

    New_index = New_Focus.TabIndex
    Field_Check(New_index) = False
    
    I_idx = 0
    If New_Focus.TabIndex = I_idx Then Exit Sub
    
    Field_Check_SW = True
    Do Until I_idx > New_Focus.TabIndex Or Field_Check_SW = False
        If Field_Name(I_idx) > "" And Field_Check(I_idx) = False Then
            Select Case UCase(Field_Name(I_idx))
                Case " I_PART_NO"
                    GoSub  I_PART_NO_RTN
                Case "I_REV_NO"
                    GoSub I_REV_NO_RTN
                Case "I_DESC_NM"
                    GoSub I_DESC_NM_RTN
                Case "C_ITEM_NO"
                    GoSub C_ITEM_NO_RTN
                Case "I_METERIAL_NM"
                    GoSub I_METERIAL_NM_RTN
                Case "C_PO_NO"
                    GoSub C_PO_NO_RTN
                Case "C_LOT_NO"
                    GoSub C_LOT_NO_RTN
                Case "I_BOXUNIT_NB"
                    GoSub I_BOXUNIT_NB_RTN
                Case "I_END_CT_NO"
                    GoSub I_END_CT_NO_RTN
                Case "I_TOTAL_CT_NO"
                    GoSub I_TOTAL_CT_NO_RTN
            End Select
            
            If Field_Check_SW = False Then
                Beep
                Field_Check(I_idx) = False
                Set New_Focus = Field_Object(I_idx)
                'New_Focus.SetFocus
                Exit Sub
            Else
                Field_Check(I_idx) = True
            End If
        End If
        I_idx = I_idx + 1
    Loop
    Exit Sub

 I_PART_NO_RTN:
    Msg = ""
    
    sql$ = "select * from PRD_MAST_1_VIEW "
    sql$ = sql$ + " where iPART_NO = '" +  I_PART_NO + "' "
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User

    If SSDATA.RecordCount = 0 Then
        SSDATA.Close
        Msg = GLO_MSGLINE(6)
        Field_Check_SW = False
        Return
    End If

    SSDATA.Close

    Return
                
I_REV_NO_RTN:
    PnlMsg = ""
    
    'Check input Lot Production Info & Part#, Rev
    sql$ = " select distinct a.prod_cd, a.pmatr_kb, a.iPART_NO, a.Prev_No, a.prod_gubun, a.pdesc_nm, a.lowheigh_kb, " & vbCrLf
    sql$ = sql$ & " nvl((select f.mcode_nm from SYS_CODPD_06 f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and  f.mcode_cd = a.bake_kb and mcode_class = a.prod_gubun), ' ') bake, " & vbCrLf
        
    If C_PO_NO = "" Then
        ' 각 제품별 기본 Box Packing 수량을 SYS_CODPD_22 에서 읽어 온다
        ' TR : Tray lowheigh_kb 이용
        ' TE : EndCap psize_nm 이용
        ' 나머지 제품은 prod_gubun 이용
        sql$ = sql$ & " decode(a.prod_gubun, 'TR', (select distinct ref02 from SYS_CODPD_22  f, PRD_MAST_2 g where g.comp_cd = f.comp_cd and g.bsns_cd = f.bsns_cd and g.ipart_no = a.ipart_no and   f.mcode_cc = g.lowheigh_kb), " & vbCrLf
        sql$ = sql$ & "                      'TE', (select distinct ref02 from SYS_CODPD_22  f, PRD_MAST_2 g where g.comp_cd = f.comp_cd and g.bsns_cd = f.bsns_cd and g.ipart_no = a.ipart_no and   f.mcode_nm = trim(g.psize_nm)), " & vbCrLf
        sql$ = sql$ & "                           (select distinct ref02 from SYS_CODPD_22  f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and mcode_cc = a.prod_gubun)) box_unit_qt, " & vbCrLf
    Else
        'PO가 있을 경우 MKT_CUST_2의 Box Pack 수량을 읽어온다.
        If CHK_INSMP.Value = 1 Then
            sql$ = sql$ & " nvl((select distinct max(f.box_pack_qt) from MKT_CUST_2 f, MKT_ORD_1 g " & vbCrLf
            sql$ = sql$ & "      where f.custom_cd = '" & Right(C_CUSTOM_CD, 5) & "' "
            sql$ = sql$ & "      and   f.prod_cd = a.prod_cd  " & vbCrLf
            sql$ = sql$ & "      and   g.comp_cd = f.comp_cd and g.bsns_cd = f.bsns_cd and f.custom_cd = g.insmp_custom_cd  " & vbCrLf
            sql$ = sql$ & "      and   g.insmp_po_no = '" & C_PO_NO & "' " & vbCrLf
            sql$ = sql$ & "      and   rownum = 1),1) box_unit_qt," & vbCrLf
            
            sql$ = sql$ & " nvl((select distinct f.cdesc_tx4 from MKT_CUST_2 f, MKT_ORD_1 g " & vbCrLf
            sql$ = sql$ & "      where f.custom_cd = '" & Right(C_CUSTOM_CD, 5) & "' "
            sql$ = sql$ & "      and   f.prod_cd = a.prod_cd  " & vbCrLf
            sql$ = sql$ & "      and   g.comp_cd = f.comp_cd and g.bsns_cd = f.bsns_cd and f.custom_cd = g.insmp_custom_cd  " & vbCrLf
            sql$ = sql$ & "      and   g.insmp_po_no = '" & C_PO_NO & "' " & vbCrLf
            sql$ = sql$ & "      and   rownum = 1),' ') cdesc_tx4," & vbCrLf
        Else
            sql$ = sql$ & " nvl((select distinct max(f.box_pack_qt) from MKT_CUST_2 f, MKT_ORD_1 g " & vbCrLf
            sql$ = sql$ & "      where f.custom_cd = '" & Right(C_CUSTOM_CD, 5) & "' "
            sql$ = sql$ & "      and   f.prod_cd = a.prod_cd  " & vbCrLf
            sql$ = sql$ & "      and   g.comp_cd = f.comp_cd and g.bsns_cd = f.bsns_cd and f.custom_cd = g.custom_cd  " & vbCrLf
            sql$ = sql$ & "      and   g.po_no = '" & C_PO_NO & "' " & vbCrLf
            sql$ = sql$ & "      and   rownum = 1),1) box_unit_qt," & vbCrLf
            
            sql$ = sql$ & " nvl((select distinct f.cdesc_tx4 from MKT_CUST_2 f, MKT_ORD_1 g " & vbCrLf
            sql$ = sql$ & "      where f.custom_cd = '" & Right(C_CUSTOM_CD, 5) & "' "
            sql$ = sql$ & "      and   f.prod_cd = a.prod_cd  " & vbCrLf
            sql$ = sql$ & "      and   g.comp_cd = f.comp_cd and g.bsns_cd = f.bsns_cd and f.custom_cd = g.custom_cd  " & vbCrLf
            sql$ = sql$ & "      and   g.po_no = '" & C_PO_NO & "' " & vbCrLf
            sql$ = sql$ & "      and   rownum = 1),' ') cdesc_tx4," & vbCrLf
        End If
    End If
    
    'Check SID # from MKT_ORD_11
    sql$ = sql$ + " nvl((select distinct sid_no from MKT_ORD_11  " & vbCrLf
    sql$ = sql$ + " where order_no in (select order_no from MKT_ORD_1 " & vbCrLf
    sql$ = sql$ + "                    where custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' " & vbCrLf
    sql$ = sql$ + "                    and   deli_kb <> '2'" & vbCrLf
    sql$ = sql$ + "                    and   comp_cd = '" & GLO_COMP_CD & "' and bsns_cd = '" & GLO_BSNS_CD & "' and po_no = '" & C_PO_NO & "' )" & vbCrLf
    sql$ = sql$ + " and  comp_cd = a.comp_cd and bsns_cd = a.bsns_cd and prod_cd = a.prod_cd and rownum = 1),'0') item_no," & vbCrLf
    sql$ = sql$ & " nvl((select distinct f.smp_mater_nm from PRD_SMP_1 f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.prod_cd = a.prod_cd and f.mold_no = a.mold_no and f.rev_no = a.Prev_No ),' ') smp_mater_nm, " & vbCrLf
    sql$ = sql$ & " nvl((select b.mcode_yn from SYS_CODPD_16 b where b.comp_cd = a.comp_cd and b.bsns_cd = a.bsns_cd and b.mcode_cd = trim(a.pmatr_kb) and b.mcode_class = a.prod_gubun), ' ') mater_nm, " & vbCrLf
    sql$ = sql$ & " nvl((select b.ref01 from SYS_CODPD_16 b where b.comp_cd = a.comp_cd and b.bsns_cd = a.bsns_cd and b.mcode_cd = trim(a.pmatr_kb) and b.mcode_class = a.prod_gubun), ' ') color_nm, " & vbCrLf
    
    'Tray - Check Mold No on Lot#
    sql$ = sql$ & " decode(a.prod_gubun, 'TR', (select count(*) from PRD_WORK_1 f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.prod_cd = a.prod_cd and f.lot_no like '" & cDll.Gf_UnFormat("M", C_LOT_NO) & "%'), " & vbCrLf
    sql$ = sql$ & "                           1) lot_count " & vbCrLf
    
    sql$ = sql$ & " from PRD_MAST_1_VIEW a " & vbCrLf
    sql$ = sql$ & " where a.iPART_NO = '" +  I_PART_NO + "' " & vbCrLf
    sql$ = sql$ & " and   a.Prev_No = '" + I_REV_NO + "' " & vbCrLf
    sql$ = sql$ & " and a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User

    If SSDATA.RecordCount = 0 Then
        SSDATA.Close
        PnlMsg = GLO_MSGLINE(6)
        Field_Check_SW = False
        Return
    End If
    
    If Trim(SSDATA!lot_count) = 0 Then
        strMsg = IIf(GLO_KOR_VISIBLE, "입력한 Lot#는 ERP상에 생산된 근거가 없습니다.", "There are no Production Information for the Lot No")    ' Define message.
        strMsg = strMsg & vbCrLf & IIf(GLO_KOR_VISIBLE, "계속 진행하시겠습니까?", "Do you want to keep going?")
        If MsgBox(strMsg, vbYesNo, "Data Check") = vbNo Then
            Field_Check_SW = False
            Return
        End If
    End If
    D_PROD_KB = Trim(SSDATA!prod_gubun)
    If I_DESC_NM = "" Then I_DESC_NM = Trim(SSDATA!cdesc_tx4)
    If I_DESC_NM = "" Then I_DESC_NM = Trim(SSDATA!pdesc_nm) + vbCrLf + Trim(SSDATA!bake)  ' Manually input
    I_BOXUNIT_NB = Trim(SSDATA!box_unit_qt)
    If SSDATA!prod_gubun = "TR" Then Call SET_TRAY_BOXUNIT(Trim(SSDATA!ipart_no), Trim(SSDATA!Prev_No))  ' Tray인 경우 NTK에 대한 Box Unit
    
    If Trim(SSDATA!mater_nm) = "" Then
        MsgBox IIf(GLO_KOR_VISIBLE, "업무표준화코드에서 원료에 관한 정보가 없습니다." & vbCrLf & "업무표준화 정보를 확인하십시요", _
                                    "There are no data on Code Table." & vbCrLf & "Check Code Table")
    Else
        I_METERIAL_NM = IIf(isAgentPeak = True And Trim(SSDATA!smp_mater_nm) <> "", Trim(SSDATA!smp_mater_nm), Trim(SSDATA!mater_nm))
        I_COLOR_NM = Trim(SSDATA!color_nm)
    End If

    Call ComboLoad_ITEM_NO(SSDATA!PROD_CD, Right(C_CUSTOM_CD, 5))
    Call GS_ComboSelect(C_ITEM_NO, SSDATA!item_no)
    If C_ITEM_NO.ListIndex = -1 Then C_ITEM_NO.ListIndex = 0
    SSDATA.Close
    
    Return

I_DESC_NM_RTN:
    PnlMsg = ""

    If I_DESC_NM = "" Then
        PnlMsg = GLO_MSGLINE(21)
        Field_Check_SW = False
        Return
    End If

    If cDll.Gf_InstrCount(I_DESC_NM, vbCrLf) > 1 Then
        PnlMsg = IIf(GLO_KOR_VISIBLE, "두줄 까지만 허용 됩니다.", "Only 2 lines are allowed")
        Field_Check_SW = False
        Return
    End If
    Return

C_ITEM_NO_RTN:
    If C_ITEM_NO.ListCount = 0 Then
        MsgBox ("Customer 정보에 Item # 정보가 없습니다. 거래처정보를 확인하십시오")
    ElseIf C_ITEM_NO.ListCount > 1 Then
        If C_ITEM_NO.ListIndex = -1 Then MsgBox (C_ITEM_NO.ListCount & "개의 Item #를 가지고 있습니다." & vbCrLf & "Item#를 선택하세요")               ' Item No가 하나 이상인 경우
    End If
    
    Return
                
I_METERIAL_NM_RTN:
    PnlMsg = ""
    
    If I_METERIAL_NM = "" Then
        PnlMsg = GLO_MSGLINE(21)
        Field_Check_SW = False
        Return
    End If
    
    
    Return
                
C_PO_NO_RTN:
    PnlMsg = ""

    If C_PO_NO = "" Then
        PnlMsg = GLO_MSGLINE(21)
        Field_Check_SW = False
        Return
    End If

    Return
                
C_LOT_NO_RTN:
    PnlMsg = ""

    If C_LOT_NO = "" Then
        PnlMsg = GLO_MSGLINE(21)
        Field_Check_SW = False
        Return
    End If

    Return
                
I_BOXUNIT_NB_RTN:
    PnlMsg = ""

    If Val(I_BOXUNIT_NB) = 0 Then
        PnlMsg = IIf(GLO_KOR_VISIBLE, "'0'은 안됩니다.", "'0' is not allowed.")
        Field_Check_SW = False
        Return
    End If

    Return
                
I_END_CT_NO_RTN:
    PnlMsg = ""

    If Val(I_START_CT_NO) = 0 Or _
       Val(I_START_CT_NO) > Val(I_END_CT_NO) Or _
       Val(I_END_CT_NO) > Val(I_TOTAL_CT_NO) Then
        PnlMsg = IIf(GLO_KOR_VISIBLE, "번호 오류입니다.", "Number Fail")
        Field_Check_SW = False
        Return
    End If

I_TOTAL_CT_NO_RTN:
    PnlMsg = ""

    If Val(I_TOTAL_CT_NO) = 0 Then
        PnlMsg = IIf(GLO_KOR_VISIBLE, "'0'은 안됩니다.", "'0' is not allowed.")
        Field_Check_SW = False
        Return
    End If

    If Val(I_TOTAL_CT_NO) < Val(I_END_CT_NO) Or Val(I_TOTAL_CT_NO) < Val(I_START_CT_NO) Then
        PnlMsg = IIf(GLO_KOR_VISIBLE, "번호 오류입니다.", "Number Fail")
        Field_Check_SW = False
        Return
    End If

    Return

End Sub
Private Sub C_CUSTOM_CD_Click()
    Call C_CUSTOM_CD_LostFocus
End Sub

Private Sub C_CUSTOM_CD_GotFocus()
    FIELD_VERIFY
End Sub

Private Sub C_CUSTOM_CD_LostFocus()
    Dim strTemp     As String
    Dim sql         As String
       
    ' Display Shipping Date for the C_LOT_NO for Delivery Order for the Lot No
    sql$ = "select distinct a.outlot_no, nvl(a.custom_line_cd,' ') as custom_line_cd, a.outorder_no "
    sql$ = sql$ & " from MKT_OUT_11 a " & vbCrLf
    sql$ = sql$ & " where a.outorder_no in " & vbCrLf
    sql$ = sql$ & "                         (select b.outorder_no from MKT_OUT_1 b, MKT_ORD_1 c " & vbCrLf 'Oem Pro
    sql$ = sql$ & "                          where  c.comp_cd = '" & GLO_COMP_CD & "' and c.bsns_cd = '" & GLO_BSNS_CD & "' "
    If CHK_INSMP.Value = 1 Then
        sql$ = sql$ & "                      and    c.insmp_kb = '1'" & vbCrLf
        sql$ = sql$ & "                      and    c.insmp_custom_cd = '" & Trim(Right(C_CUSTOM_CD, 10)) & "' " & vbCrLf
    Else
        sql$ = sql$ & "                      and    c.insmp_kb = '0'" & vbCrLf
        sql$ = sql$ & "                      and    c.custom_cd = '" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    End If
    sql$ = sql$ & "                          and    b.order_no = c.order_no " & vbCrLf
    sql$ = sql$ & "                          and    b.output_dt = '" & Format(SHIP_DATE, "yyyymmdd") & "') " & vbCrLf
    sql$ = sql$ & " and   a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and   a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    sql$ = sql$ & " order by a.outlot_no "
    
    Set SSCODE = Nothing
    SSCODE.Open sql$, DB_User
    
    C_LOT_NO.Clear
    C_LOT_ID.Clear
    If SSCODE.RecordCount > 0 Then
        Do Until SSCODE.EOF
            strTemp = Gf_LotNo_Format("TR", SSCODE!outlot_no)
            C_LOT_NO.AddItem strTemp
            If Trim(Right(C_CUSTOM_CD, 10)) = "20289" Or Trim(Right(C_CUSTOM_CD, 10)) = "21311" Then 'SPIL
                C_LOT_ID.AddItem SSCODE!outorder_no
            Else 'SCK 등
                C_LOT_ID.AddItem SSCODE!custom_line_cd
            End If
            SSCODE.MoveNext
        Loop
        C_LOT_NO.ListIndex = 0
        C_LOT_ID.ListIndex = 0
    End If
    SSCODE.Close

    ' Check WWCustomerCD, and arrange data
    sql$ = "select wwcustom_cd from MKT_CUST_1 "
    sql$ = sql$ & " where custom_cd = '" & Trim(Right(C_CUSTOM_CD, 10)) & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSCODE = Nothing
    SSCODE.Open sql$, DB_User
    
    If SSCODE.RecordCount = 0 Then
        varWWCustomCD = ""
    Else
        varWWCustomCD = Trim(SSCODE!wwcustom_cd)
    End If
    SSCODE.Close
    
    Label1(4).Visible = False
    C_INV_NO.Visible = False
        
    Label1(15).Visible = False
    Label1(16).Visible = False
    I_TOTAL_WEIGHT.Visible = False
    I_PCNT_NO.Visible = False
    
    If Trim(Right(C_CUSTOM_CD, 10)) = "20015" Or Trim(Right(C_CUSTOM_CD, 10)) = "20163" Then '20006: ASEK 일 경우 RT No를 가져온다. 2013-12-23 by S+G kimsi
        Label1(22).Visible = True
        Label1(22).Caption = "RT#:"
        C_LOT_ID.Visible = True
    ElseIf GLO_COMP_CD = "FU" And (Trim(Right(C_CUSTOM_CD, 10)) = "21347" Or Trim(Right(C_CUSTOM_CD, 10)) = "20888" Or Trim(Right(C_CUSTOM_CD, 10)) = "22134") Then
        Label1(22).Visible = True
        Label1(22).Caption = "RT#:"
        C_LOT_ID.Visible = True
    ElseIf Trim(Right(C_CUSTOM_CD, 10)) = "10426" Or Trim(Right(C_CUSTOM_CD, 5)) = "10005" Then 'SCK LOT ID 표시
        Label1(22).Visible = True
        Label1(22).Caption = "LOT ID:"
        C_LOT_ID.Visible = True
    ElseIf Trim(Right(C_CUSTOM_CD, 10)) = "21022" Then 'TOMOE 고객의 경우 KYOCERA KST PO# 항목 표시
        Label1(22).Visible = True
        Label1(22).Caption = "KST PO#:"
        C_LOT_ID.Visible = True
    Else
        Label1(22).Visible = False
        C_LOT_ID.Visible = False
    End If
    '===================================
    ' Invoice NO Field 보이기
    '===================================
    If ((Trim(Right(C_CUSTOM_CD, 10)) = "20180" Or Trim(Right(C_CUSTOM_CD, 10)) = "21844" Or Trim(Right(C_CUSTOM_CD, 10)) = "10507" Or Trim(Right(C_CUSTOM_CD, 10)) = "21832" Or Trim(Right(C_CUSTOM_CD, 10)) = "10519" Or Trim(Right(C_CUSTOM_CD, 10)) = "10536" Or Trim(Right(C_CUSTOM_CD, 10)) = "10558") And GLO_COMP_CD = "FU") Then    'TSMC by FRC Liam 2019.01.11
        Label1(4).Visible = True
        C_INV_NO.Visible = True
    End If
    '-------------------------------------------------------------------
    
    Select Case True
        Case PRT_OP(0)
            I_SET_QT.Enabled = True
            I_SET_QT = 1
        Case PRT_OP(1)
            I_SET_QT.Enabled = True
            I_SET_QT = 1
    End Select

End Sub

Private Sub C_INV_NO_LostFocus()

    Dim sql         As String
    
    sql$ = "select a.shipto_cd, a.outorder_no, a.order_no, "
    sql$ = sql$ & "  nvl((select distinct(c.wwcustom_cd) from MKT_ORD_1 b, MKT_CUST_2 c "
    sql$ = sql$ & "        where b.order_no  = a.order_no"
    sql$ = sql$ & "        and   b.comp_cd = a.comp_cd and b.bsns_cd = a.bsns_cd "
    If CHK_INSMP.Value = 1 Then
        sql$ = sql$ & "        and  b.insmp_custom_cd = c.custom_cd), ' ') wwcustom_cd"
    Else
        sql$ = sql$ & "        and  b.custom_cd = c.custom_cd), ' ') wwcustom_cd"
    End If
    sql$ = sql$ & " from MKT_OUT_1 a "
    sql$ = sql$ & " where a.po_no = '" & C_PO_NO & "' "
    sql$ = sql$ & " and   a.invoice_no = '" & C_INV_NO & "' "
    sql$ = sql$ & " and   a.outorder_no in (select b.outorder_no from MKT_OUT_11 b "
    sql$ = sql$ & "                         where b.comp_cd = a.comp_cd and b.bsns_cd = a.bsns_cd and b.outlot_no = '" & strLotNo & "')"
    sql$ = sql$ & " and a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSCODE = Nothing
    SSCODE.Open sql$, DB_User
    
    If SSCODE.RecordCount > 0 Then
        
    End If
    SSCODE.Close
End Sub

Private Sub C_PO_NO_GotFocus()
    Call FIELD_VERIFY
End Sub

Private Sub SET_TRAY_BOXUNIT(varPartNo As String, varRevNo As String)
    ' Customer 가 NTK(20094)이며, Thickness가 High일 경우 30이 아닌 60을 사용
    ' NTK 일 경우 11D-1919-226/B, 11M-5334-N2F/B인 경우에만 위에서 정의한 것과 같이 수량을 60으로 하고,
    ' 나머지는 High 일 경우 30개, Low일 경우 50개로 한다. -> Box Unit 수량을 따른다.
    '- 2004. 12. 21 by hklee requested by hslee
    If Right(C_CUSTOM_CD, 5) <> "21120" Then Exit Sub
    If (varPartNo = "11D-1919-226" And varRevNo = "B") _
    Or (varPartNo = "11M-5334-N2F" And varRevNo = "B") Then
        I_BOXUNIT_NB = 60
    End If
End Sub

Private Sub C_PO_NO_LostFocus()
    ' Check Data with PO NO
    Dim sql                 As String
    Dim sCustomer           As String  'Ship to
    Dim varMKTOUT1          As String
    Dim varMKTOUT11         As String
    Dim varMKTORD1          As String
    
    '초기화
    varMKTOUT1 = "MKT_OUT_1"
    varMKTOUT11 = "MKT_OUT_11"
    varMKTORD1 = "MKT_ORD_1"
    
    If C_PO_NO.ListCount > 1 Then
        MsgBox "선택하신 제품으로 출하지시된 PO가 2건 이상 존재합니다. PO를 반드시 확인하시기 바랍니다."
    End If
    
    'Check Ship To Customer Data
    sql = " select distinct b.shipto_cd from MKT_ORD_1 a, MKT_OUT_1 b"
    If CHK_INSMP.Value = 1 Then
        sql = sql & " where a.insmp_po_no = '" & C_PO_NO & "'"
        sql = sql & " and   a.insmp_custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' " 'from
    Else
        sql = sql & " where a.po_no = '" & C_PO_NO & "'"
        sql = sql & " and   a.custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' " 'from
    End If
    sql = sql & " and   a.order_no = b.order_no"
    sql = sql & " and   b.output_dt ='" & Format(SHIP_DATE, "yyyymmdd") & "'"
    sql = sql & " and   a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql = sql & " and   a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    sql = sql & " and   a.comp_cd = b.comp_cd and a.bsns_cd = b.bsns_cd " & vbCrLf

    Set SSCODE = Nothing
    SSCODE.Open sql, DB_User
    
    If SSCODE.RecordCount > 0 Then
        sCustomer = SSCODE!shipto_cd
    End If
    SSCODE.Close
   
    '========================================================
    ' Check Delivery Data for selected Customer, input Lot #, Selected PO #
    ' and fill Part #/Rev/Material/Color/Item#
    '========================================================
    sql$ = "select distinct a.prod_cd, NVL(a.outdirect_qt, 0) outdirect_qt, b.iPART_NO, b.Prev_No, b.prod_gubun, b.mater_nm, b.pdesc_nm, b.lowheigh_kb, a.sid_no item_no, " & vbCrLf
    sql$ = sql$ & " (select f.mcode_yn from SYS_CODPD_16 f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.mcode_cd = trim(b.pmatr_kb) and f.mcode_class = b.prod_gubun) material, " & vbCrLf
    sql$ = sql$ & " (select f.ref01 from SYS_CODPD_16 f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.mcode_cd = trim(b.pmatr_kb) and f.mcode_class = b.prod_gubun) color, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct f.branch_cd from MKT_ORD_1 f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and " & IIf(CHK_INSMP.Value = 1, " f.insmp_custom_cd ", " f.custom_cd") & " = '" & Right(C_CUSTOM_CD, 5) & "' and f.po_no = '" & C_PO_NO & "' " & vbCrLf
    sql$ = sql$ & "                                                 and f.order_no in (select d.order_no from MKT_OUT_1 d  where d.comp_cd = a.comp_cd and d.bsns_cd = a.bsns_cd and d.outorder_no = a.outorder_no)), ' ') branch_cd, " & vbCrLf
    sql$ = sql$ & " nvl((select max(d.item_no) from MKT_CUST_2 d " & vbCrLf
    sql$ = sql$ & "      where d.custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' " & vbCrLf
    sql$ = sql$ & "      and   d.comp_cd = a.comp_cd and d.bsns_cd = a.bsns_cd and d.prod_cd = b.prod_cd), ' ') max_item_no, " & vbCrLf
    sql$ = sql$ & " (select count(d.item_no) from MKT_CUST_2 d " & vbCrLf
    sql$ = sql$ & "  where d.custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' " & vbCrLf
    sql$ = sql$ & "  and d.comp_cd = a.comp_cd and d.bsns_cd = a.bsns_cd and   d.prod_cd = b.prod_cd) item_no_count, " & vbCrLf
    sql$ = sql$ & " nvl((select e.mcode_nm from SYS_CODPD_06 e where e.comp_cd = a.comp_cd and e.bsns_cd = a.bsns_cd and e.mcode_cd = b.bake_kb and e.mcode_class = b.prod_gubun), ' ') bake, " & vbCrLf
  ' PO에 대한 각 제품별 기본 Box Pack 수량은 MKT_CUST_2에서 읽어 온다
    sql$ = sql$ & " nvl((select max(f.box_pack_qt) from MKT_CUST_2 f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' and  f.prod_cd = b.prod_cd),1) box_unit_qt " & vbCrLf
    sql$ = sql$ & " from MKT_OUT_11 a, PRD_MAST_1_VIEW b " & vbCrLf
    sql$ = sql$ & " where a.outlot_no = '" & strLotNo & "' " & vbCrLf
    sql$ = sql$ & " and   a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and   a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    sql$ = sql$ & " and   a.comp_cd = b.comp_cd and a.bsns_cd = b.bsns_cd " & vbCrLf
    sql$ = sql$ & " and   a.outorder_no in (select distinct d.outorder_no from MKT_ORD_1 c, MKT_OUT_1 d " & vbCrLf
    If CHK_INSMP.Value = 1 Then
        sql$ = sql$ & "                         where c.insmp_po_no = '" & C_PO_NO & "' " & vbCrLf
        sql$ = sql$ & "                         and   c.insmp_custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' " & vbCrLf
    Else
        sql$ = sql$ & "                         where c.po_no = '" & C_PO_NO & "' " & vbCrLf
        sql$ = sql$ & "                         and   c.custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' " & vbCrLf
    End If
    sql$ = sql$ & "                         and   c.comp_cd = a.comp_cd and c.bsns_cd = a.bsns_cd " & vbCrLf
    sql$ = sql$ & "                         and   d.order_no = c.order_no) " & vbCrLf
    sql$ = sql$ & " and    a.prod_cd = b.prod_cd "

    Set SSCODE = Nothing
    SSCODE.Open sql$, DB_User
    
    If SSCODE.RecordCount > 0 Then
         I_PART_NO = Trim(SSCODE!ipart_no)
        I_REV_NO = Trim(SSCODE!Prev_No)
        I_METERIAL_NM = Trim(SSCODE!material)
        
        varBranchCD = Trim(SSCODE!BRANCH_CD)
        isAgentPeak = IIf(varBranchCD = "15", True, False)  'AGENT PEAK
        Call ComboLoad_ITEM_NO(SSCODE!PROD_CD, Trim(Right(C_CUSTOM_CD, 5)))
        Call GS_ComboSelect(C_ITEM_NO, SSCODE!item_no)
        If C_ITEM_NO.ListIndex = -1 Then C_ITEM_NO.ListIndex = 0
        
        I_COLOR_NM = Trim(SSCODE!Color)
        I_TOTAL_QT = Trim(SSCODE!outdirect_qt)
        I_BOXUNIT_NB = Trim(SSCODE!box_unit_qt)
        If SSCODE!prod_gubun = "TR" Then Call SET_TRAY_BOXUNIT(Trim(SSCODE!ipart_no), Trim(SSCODE!Prev_No))  ' Tray인 경우 NTK에 대한 Box Unit
        
         I_PART_NO.Enabled = True
        I_REV_NO.Enabled = True
        I_METERIAL_NM.Enabled = True
        I_COLOR_NM.Enabled = True
    Else
         I_PART_NO = ""
        I_REV_NO = ""
        I_METERIAL_NM = ""
        
        C_ITEM_NO.Clear
        I_COLOR_NM = ""
        I_DESC_NM = ""
    End If
    SSCODE.Close
    
    '=====================================================
    ' Invoice NO Field가 보이는 경우 Invoice Data Check
    '
    ' 선택된 Customer, 입력된 Lot #, 선택되어진 PO #가
    ' 출하지시 내려진 것인지 확인하여 Invoice #가 필요한 경우 채운다.
    '
    ' WWCUSTOMCode = 21155 CYPRESS          AMD 23395
    ' WWCustomCD = GF_getCode("WWC_SPIL") Then 'WWCUSTOM SPIL
    ' CUstomCode   = 20015 ASE(T) KHH 추가 2005.09.14 sjsung
    '                20888 ASE SH,  TI (T) 20562
    '                ASE ADVANCED SHA 거래선(22134)
    '                20006, 20916 CARSEM(M)-S, CARSEM(M)-S 2010-06-29 jwkim
    '                20110 ASE(M) 추가 2010-12-14 jwkim
    '=====================================================
    '2010-07 Sample 진행중
    If varWWCustomCD = GF_GetCode("WWC_CYPRESS") Or varWWCustomCD = GF_GetCode("WWC_AMD") Or varWWCustomCD = GF_GetCode("WWC_SPIL") Or Trim(Right(C_CUSTOM_CD, 10)) = "20015" Or _
        Trim(Right(C_CUSTOM_CD, 10)) = "20888" Or Trim(Right(C_CUSTOM_CD, 10)) = "22134" Or Trim(Right(C_CUSTOM_CD, 5)) = "20562" Or Trim(Right(C_CUSTOM_CD, 10)) = "20163" Or Trim(Right(C_CUSTOM_CD, 10)) = "20001" Or Trim(Right(C_CUSTOM_CD, 10)) = "20002" Or Trim(Right(C_CUSTOM_CD, 10)) = "20574" Or _
        ((Trim(Right(C_CUSTOM_CD, 10)) = "20180" Or Trim(Right(C_CUSTOM_CD, 10)) = "21844" Or Trim(Right(C_CUSTOM_CD, 10)) = "10507" Or Trim(Right(C_CUSTOM_CD, 10)) = "21832") And GLO_COMP_CD = "FU") Then    'FRC TSMC  2019.01.11 by Liam
                   
'        If CHK_INSMP.Value = 1 Then
'            If Trim(Right(C_FROMSMP_NM, 10)) = GF_GetCode("SITE_TECREACH") Then 'TRERP Invocie 조회
'                varMKTOUT1 = "TRERP.MKT_OUT_1"
'                varMKTOUT11 = "TRERP.MKT_OUT_11"
'                varMKTORD1 = "TRERP.MKT_ORD_1"
'            End If
'        End If
        sql$ = "select a.invoice_no from " & varMKTOUT1 & " a " & vbCrLf
        sql$ = sql$ & " where a.order_no in (select b.order_no from " & varMKTORD1 & " b " & vbCrLf
        sql$ = sql$ & "                      where a.comp_cd = b.comp_cd and a.bsns_cd = b.bsns_cd " & vbCrLf
        If CHK_INSMP.Value = 1 Then
            sql$ = sql$ & "                      and b.insmp_custom_cd = '" & Right(C_CUSTOM_CD, 5) & "' " & vbCrLf
            sql$ = sql$ & "                      and b.insmp_po_no = '" & C_PO_NO & "' ) " & vbCrLf
        Else
            sql$ = sql$ & "                      and b.custom_cd = '" & Right(C_CUSTOM_CD, 5) & "' " & vbCrLf
            sql$ = sql$ & "                      and b.po_no = '" & C_PO_NO & "' ) " & vbCrLf
        End If
        sql$ = sql$ & " and   a.outorder_no in (select b.outorder_no from " & varMKTOUT11 & " b  where b.comp_cd = a.comp_cd and b.bsns_cd = a.bsns_cd and b.outlot_no = '" & strLotNo & "') " & vbCrLf
        sql$ = sql$ & " and  a.output_dt ='" & Format(SHIP_DATE, "yyyymmdd") & "'" & vbCrLf
        'sql$ = sql$ & " and a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
        sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
        Set SSCODE = Nothing
        SSCODE.Open sql$, DB_User
        
        C_INV_NO.Clear
        If SSCODE.RecordCount > 0 Then
            Do Until SSCODE.EOF
                ' 읽어온 InvoiceN o가 Null또는 빈문자열일 경우 Invoice Combo에 입력하지 않는다.
                If Trim(SSCODE!invoice_no) <> "" Then
                    C_INV_NO.AddItem Mid(Trim(SSCODE!invoice_no), 1, Len(Trim(SSCODE!invoice_no)) - 2) & "-" & Right(Trim(SSCODE!invoice_no), 2)
                End If
                SSCODE.MoveNext
            Loop
        End If
        'Invoice가 없을 경우 자동 생성안함 (Order 정보를 알 수 없음)
        C_INV_NO.ListIndex = -1
        If C_INV_NO.ListCount > 0 Then C_INV_NO.ListIndex = 0
        
        SSCODE.Close
    End If
    '---------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------
    ' 2020.03.13 CPL jenny & JHKIM 요청에 따라 End Cust 제품의 경우 Ship To를 기준으로 라벨 출력을 진행 한다.
    D_CUSTOM_NM = Trim(Left(C_CUSTOM_CD, 30))
    D_SCUSTOM_CD = Trim(Right(C_CUSTOM_CD, 10))  'Ship To
    D_SWWCUST_CD = varWWCustomCD
    
    'EndCustomer 정보 확인
    sql$ = "select distinct endlabel_kb, scustom_cd, endwwc_cd, fnc_wwc_nm(" & IIf(GLO_KOR_VISIBLE, "'KOR'", "'ENG'") & ", 'Y', a.endwwc_cd) endwwc_nm, " & vbCrLf
    sql$ = sql$ & " (select distinct f.wwcustom_cd from MKT_CUST_1 f where f.custom_cd = a.scustom_cd and f.comp_cd = a.comp_cd and f.bsns_cd =a.bsns_cd ) shipwwc_cd"
    sql$ = sql$ & " from MKT_ORD_1 a " & vbCrLf
    sql$ = sql$ & " where jit_kb = '0'" & vbCrLf 'Jit
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    If CHK_INSMP.Value = 1 Then
        sql$ = sql$ & "  and  a.insmp_po_no = '" & C_PO_NO & "' " & vbCrLf
        sql$ = sql$ & "  and  a.insmp_custom_cd = '" & Trim(Right(C_CUSTOM_CD, 10)) & "' " & vbCrLf
    Else
        sql$ = sql$ & "  and  a.po_no = '" & C_PO_NO & "' " & vbCrLf
        sql$ = sql$ & "  and  a.custom_cd = '" & Trim(Right(C_CUSTOM_CD, 10)) & "' " & vbCrLf
    End If
    sql$ = sql$ & " order by endlabel_kb desc "
    
    Set SSCODE = Nothing
    SSCODE.Open sql$, DB_User
    
    If SSCODE.RecordCount > 0 Then
        D_ENDLABEL_KB = SSCODE!endlabel_kb
        varEndwwc_CD = SSCODE!endwwc_cd
        If SSCODE!endlabel_kb = "1" Then
            D_CUSTOM_NM = SSCODE!endwwc_nm
            D_SCUSTOM_CD = SSCODE!scustom_cd
            D_SWWCUST_CD = SSCODE!shipwwc_cd
        End If
    End If
    SSCODE.Close
    '-------------------------------------------------------------------
    
    Exit Sub

ERR_MSG_RTN:

End Sub

Private Sub ComboLoad_ITEM_NO(varProdCD As String, varCustomCd As String)
    Dim sql         As String
    Dim ssTemp      As New ADODB.Recordset
    
    sql$ = " select item_no from MKT_CUST_2 "
    sql$ = sql$ & " where custom_cd = '" & varCustomCd & "'"
    sql$ = sql$ & " and  prod_cd = '" & varProdCD & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set ssTemp = Nothing
    ssTemp.Open sql$, DB_User
    
    With C_ITEM_NO
        .Clear
    
        If ssTemp.RecordCount = 0 Then
            ssTemp.Close
            .AddItem ""
        Else
            If ssTemp.RecordCount >= 2 Then
                MsgBox IIf(GLO_KOR_VISIBLE, "거래처별 ITEM 정보에 등록된 내역이 두개 이상입니다." & vbCrLf & "고객 SID#(ItemNo)를 CPL팀에 확인 후 선택 바랍니다.", _
                    "Two or more SID#(Item no) registered in customer list," & vbCrLf & "Please check the SID#(Item no) with CPL team."), _
                    vbExclamation, "Check Data - SID#"
            End If
            Do Until ssTemp.EOF
                .AddItem ssTemp!item_no
                ssTemp.MoveNext
            Loop
            ssTemp.Close
        End If
        .ListIndex = 0
    End With
End Sub



Private Sub CHK_INSMP_Click()
    If CHK_INSMP.Value = 1 Then
        Frame_Ship.Visible = True
    Else
        Frame_Ship.Visible = False
    End If
End Sub

Private Sub I_BOXUNIT_NB_LostFocus()
    Dim LngTotProdQT        As Long
    Dim LngBoxUnit          As Long
    
    I_BOXUNIT_NB = Trim(I_BOXUNIT_NB)
    If I_BOXUNIT_NB = "" Then Exit Sub
    If I_BOXUNIT_NB = "0" Then Exit Sub
    
    If I_TOTAL_QT = "" Then Exit Sub
        
    LngTotProdQT = Val(I_TOTAL_QT)
    LngBoxUnit = Val(I_BOXUNIT_NB)
    
    I_END_CT_NO = IIf(LngTotProdQT Mod LngBoxUnit > 0, LngTotProdQT \ LngBoxUnit + 1, LngTotProdQT \ LngBoxUnit) + Val(I_START_CT_NO) - 1
    I_TOTAL_CT_NO = I_END_CT_NO
0
End Sub

Private Sub I_BUNDLE_QT_Change()
    I_SET_QT = Val(I_BUNDLE_QT)
    If Val(I_BUNDLE_QT) = 0 Or Trim(I_BUNDLE_QT) = "" Then I_SET_QT = 1
End Sub

Private Sub I_COLOR_NM_LostFocus()
    I_COLOR_NM = Trim(I_COLOR_NM)
End Sub

Private Sub I_DESC_NM_GotFocus()
    Call FIELD_VERIFY
End Sub

Private Sub I_DESC_NM_LostFocus()
    I_DESC_NM = Trim(I_DESC_NM)
End Sub

Private Sub I_END_CT_NO_GotFocus()
    FIELD_VERIFY
End Sub

Private Sub I_END_CT_NO_LostFocus()
    I_END_CT_NO = Trim(I_END_CT_NO)
End Sub

Private Sub C_ITEM_NO_GotFocus()
    Call FIELD_VERIFY
End Sub

Private Sub C_ITEM_NO_LostFocus()
    C_ITEM_NO = Trim(C_ITEM_NO)
End Sub

Private Sub I_METERIAL_NM_LostFocus()
    I_METERIAL_NM = Trim(I_METERIAL_NM)
End Sub

Private Sub  I_PART_NO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub I_PCNT_NO_LostFocus()
    I_PCNT_NO = Trim(I_PCNT_NO)
End Sub

Private Sub I_REV_NO_LostFocus()
    I_REV_NO = Trim(I_REV_NO)
    If I_REV_NO = "" Then Exit Sub
    
    ' Check Lot#, Part# with Rev
    If CHECK_INPUT_DATA( I_PART_NO, I_REV_NO, C_LOT_NO) = False Then
        I_REV_NO = "1"
        MsgBox IIf(GLO_KOR_VISIBLE, "Lot#의 입력이 Part No, Rev과 맞지 않습니다." & vbCrLf & "Lot#, Part No, Rev을 확인하시기 바랍니다.", _
                                    "Lot# is not correct with Part No&Rev." & vbCrLf & "Check Lot#, Part No, Rev.")
        Exit Sub
    End If
    
    Call FIELD_VERIFY
End Sub

Private Sub I_SET_QT_LostFocus()
    Dim strMsg      As String
    
    I_SET_QT = Trim(I_SET_QT)
    If I_SET_QT = "" Then Exit Sub
    
    If I_SET_QT = 0 Then
        MsgBox IIf(GLO_KOR_VISIBLE, "Set 수량은 1개 이상이어야 합니다.", "Set Qty should be more then 1")
        I_SET_QT = 1
        Exit Sub
    End If
    
    If I_SET_QT > 2 Or I_SET_QT < 1 Then
        strMsg = IIf(GLO_KOR_VISIBLE, " 인쇄하시려는 라벨 Set 수량이" & I_SET_QT & " 개가 맞습니까?", " Lable Qty for Print is " & I_SET_QT & ", is it right?")
        
        If MsgBox(strMsg, 292, "Print") = vbNo Then Exit Sub
        I_SET_QT = IIf(PRT_OP(0).Value = True, 1, 2)
        PnlMsg = ""
    End If
End Sub

Private Sub I_START_CT_NO_LostFocus()
    Dim LngTotProdQT        As Long
    Dim LngBoxUnit          As Long
    
    I_START_CT_NO = Trim(I_START_CT_NO)
    If I_START_CT_NO = "" Then Exit Sub
    
    If I_TOTAL_QT = "" Then Exit Sub
    If I_BOXUNIT_NB = "" Then Exit Sub
    If I_BOXUNIT_NB = "0" Then Exit Sub
    
    LngTotProdQT = Val(I_TOTAL_QT)
    LngBoxUnit = Val(I_BOXUNIT_NB)
    
    I_END_CT_NO = IIf(LngTotProdQT Mod LngBoxUnit > 0, LngTotProdQT \ LngBoxUnit + 1, LngTotProdQT \ LngBoxUnit) + Val(I_START_CT_NO) - 1
    I_TOTAL_CT_NO = I_END_CT_NO
End Sub

Private Sub I_TOTAL_CT_NO_LostFocus()
    I_TOTAL_CT_NO = Trim(I_TOTAL_CT_NO)
End Sub

Private Sub I_TOTAL_QT_LostFocus()
    Dim Result_QT   As Integer
    Dim Box_QT      As Integer
    
    I_TOTAL_QT = Trim(I_TOTAL_QT)
    If I_TOTAL_QT = "" Then Exit Sub
    
    If Val(I_TOTAL_QT) > 0 And Val(I_TOTAL_QT) < Val(I_BOXUNIT_NB) Then I_BOXUNIT_NB = I_TOTAL_QT
    
    Result_QT = 0
    Box_QT = 0
    If Val(I_TOTAL_QT) > 0 And Val(I_BOXUNIT_NB) > 0 Then
        Result_QT = Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB)
        Box_QT = Val(I_TOTAL_QT) \ Val(I_BOXUNIT_NB)
            
        If Result_QT = 0 Then
            I_TOTAL_CT_NO = Box_QT
        Else
            I_TOTAL_CT_NO = Box_QT + 1
        End If
        I_END_CT_NO = I_TOTAL_CT_NO
        I_START_CT_NO = 1
    End If
End Sub

Private Sub Field_Clear()
    Dim i       As Integer
    Dim strTemp As String
    
    On Error Resume Next
    
    For i = 0 To Me.count - 1
        strTemp = UCase(Me.Controls(i).Name)
        If Len(strTemp) > 2 Then
            If Left(strTemp, 2) = "I_" Then
                Me.Controls(i).Text = ""
                Me.Controls(i).Caption = ""
                Me.Controls(i).Enabled = True
                Me.Controls(i).BackColor = &HFFFFFF
            ElseIf Left(strTemp, 2) = "D_" Then
                If UCase(Me.Controls(i).Container.Name) <> "KEY_PANEL" Then Me.Controls(i).Text = ""
                If UCase(Me.Controls(i).Container.Name) <> "KEY_PANEL" Then Me.Controls(i).Caption = ""
                If UCase(Me.Controls(i).Container.Name) <> "KEY_PANEL" Then Me.Controls(i).BackColor = &HE0E0E0
            ElseIf Left(strTemp, 2) = "C_" Then
                If Me.Controls(i).Tag <> "Y" Then Me.Controls(i).ListIndex = -1
                Me.Controls(i).Enabled = True
            ElseIf Left(strTemp, 4) = "CHK_" Then
                If UCase(Me.Controls(i).Container.Name) <> "KEY_FRAME" Then Me.Controls(i).Value = 0
                If UCase(Me.Controls(i).Container.Name) <> "KEY_FRAME" Then Me.Controls(i).Enabled = True
            ElseIf Left(strTemp, 4) = "OPT_" Then
                If UCase(Me.Controls(i).Container.Name) <> "KEY_FRAME" Then Me.Controls(i).Value = False
                Me.Controls(i).Enabled = True
            End If
        End If
    Next i
    
    PRT_FORM(0).Value = True    'C/T BOX 기본선택
    
End Sub

Private Function PORT_OPEN(Optional iPrtPort As Integer)
    Dim sPortNO     As Integer
    Dim sSetting    As String
    
    On Error Resume Next
    
    sPortNO = IIf(iPrtPort <> 0, iPrtPort, Val(myINI.portcfg))
    sSetting = "9600,N,8,1"

    MSComm1.CommPort = sPortNO
    MSComm1.Settings = sSetting
    MSComm1.OutBufferSize = 1024

    MSComm1.PortOpen = True
End Function

Private Function SEND_DATA(sSTRING As String)
            
    On Error GoTo ErrorRTN
            
    MSComm1.Output = sSTRING
    
    Exit Function
ErrorRTN:
    MsgBox IIf(GLO_KOR_VISIBLE, "환경설정을 다시 하세요.", "Set Configuration again."), vbCritical, "Error"

End Function

Private Function PORT_CLOSE()
    On Error Resume Next
    MSComm1.PortOpen = False
End Function

Private Sub CMD_Clear_Click()

    Call Field_Clear
    
    C_PO_NO.Clear
    C_INV_NO.Clear
    CT_PRINT_CHK.Value = 1
    
    I_METERIAL_NM = "PES"
    I_BOXUNIT_NB = 50
    I_START_CT_NO = 1
    Label_Cnt = 0
    D_TOT_LABEL_CNT = 0
    
    I_SET_QT = IIf(PRT_OP(0).Value = True, 1, 2)
    
    ReDim Label_Data(3000)
End Sub

Private Sub CMD_Exit_Click()
    Unload Me
End Sub

Private Sub CMD_PRINT_Click()

    Dim filehandle As Integer, tempNums As Integer

    Screen.MousePointer = 11
    CMD_SAVE.Enabled = False
    Key_Panel.Enabled = False
    
    ProgPnl.Caption = ""                    ' Dummy Coding
    ProgPnl.Visible = True
    ProgPnl.Caption = ""
     
    ProgressBar1.Min = 1
    ProgressBar1.Max = IIf(Val(D_TOT_LABEL_CNT) = 0, 2, Val(D_TOT_LABEL_CNT) + 1)
    
    Label_Cnt = 1
        
   'LAN PRINT(IT Center)를 선택한 경우에만 소켓 연결을 확인한다. FRC는 LAN PRINT 만을 지원한다.
    If ChkLANPRT.Value = 1 Or GLO_COMP_CD = "FU" Then   'FU는 LAN PRINT로 출력한다.
        If GLO_COMP_CD = "DW" And GLO_DEPT_CD = "10011" Then '(현재는 IT Center에서만 라벨 출력 테스트 용도로 사용중)
            Call GF_SocketConnect(Lantype_Sck, "BC_PRT0048_IP") 'IT센터 바코드 프린터 (테스트장비)
        ElseIf GLO_COMP_CD = "FU" And ChkLANPRT.Value <> 1 Then
            If EOL_Print_Check = 1 Then
                Call GF_SocketConnect(Lantype_Sck, "FRC_BCPRT_3F_EOL_IP")   'FRC Barcode Printer EOL IP For PEAK Label
            Else
                Select Case Trim(Right(C_CUSTOM_CD, 10))
                Case "20180", "21844", "10507", "21832", "10536", "10519"  '2014-04-07 TSMC (TAIWAN SEMICONDUCTOR MANUFACTURING CO.,L) ?? by ICS_jwkim&IS_Junmo, 2018-11-06 add Sites TSMC FAB2, TSMC FAB30 by Jim Chen, 2020-04-13 add Sites TSMC Fab07 and P010
                    Call GF_SocketConnect(Lantype_Sck, "FRC_BCPRT_3F_IP")   'FRC Barcode Printer 3F IP For DW Label
                Case Else
                    Call GF_SocketConnect(Lantype_Sck, "FRC_BCPRT_3F_2_IP")   'FRC Barcode Printer 3F IP For PEAK Label
                End Select
            End If
        Else
            Call GF_SocketConnect(Lantype_Sck, "BC_PRT0048_IP")
        End If
        
        If Lantype_Sck.State = sckConnected Then  'Socket 연결확인
            'MsgBox "이미 연결되어 있습니다."
        Else
            MsgBox "연결이 해제되어 있습니다." & vbCrLf & "LAN TYPE 바코드 프린터 연결상태를 확인바랍니다."
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        For Label_Cnt = 1 To Val(D_TOT_LABEL_CNT)
            Lantype_Sck.SendData (Label_Data(Label_Cnt))
        Next
    Else
        If PRT_SET(0).Value = True Then     'LPT방식
            '-----------------------------------------------------
            'LPT 프린트 (S+G 방식) '-------------------------
            filehandle = FreeFile
            Open "LPT1: " For Output As #filehandle
        
            For Label_Cnt = 1 To Val(D_TOT_LABEL_CNT)
                Print #filehandle, Label_Data(Label_Cnt)
            Next
            Close #filehandle
            '-----------------------------------------------------
        Else    'Daewon 출력방식
            'Serial Port로 출력테스트 (Datamax Print)
            For Label_Cnt = 1 To Val(D_TOT_LABEL_CNT)
                Call SEND_DATA(Label_Data(Label_Cnt))
            Next
        End If
    End If

    Screen.MousePointer = 0
    CMD_SAVE.Enabled = True
    CMD_PRINT.Enabled = False
    
    Key_Panel.Enabled = True
    
    Label_Cnt = 0
    D_TOT_LABEL_CNT = "0"
    I_END_CT_NO = "1"
    ReDim Label_Data(3000)
    
    '=============================================================
    ProgPnl.Caption = ""                    ' Dummy Coding
    ProgPnl.Visible = False
    ProgPnl.Caption = ""
    '=============================================================
End Sub
Private Function CHECK_OUTData() As String
    '선택한 고객, 제품, PO의 정보를 확인하여 PartNo, Rev, item_no, line_code, invoice 정보가 출하지시 정보와 일치하는지 체크
    Dim sql         As String
    Dim ssTemp      As New ADODB.Recordset
    Dim po_no       As String
    Dim lot_no      As String
    Dim item_no     As String
    Dim line_code   As String
    Dim part_no     As String
    Dim rev_no      As String
    Dim invoice_no  As String
    Dim custom_cd   As String
    Dim output_dt   As String
    Dim strMsg      As String
    Dim returnStr   As String
    Dim isError     As Boolean
                    
    On Error GoTo ERROR_RTN
        
    po_no = Trim(C_PO_NO)
    lot_no = Trim(C_LOT_NO)
    item_no = Trim(C_ITEM_NO)
    line_code = Trim(C_LOT_ID)
    part_no = Trim( I_PART_NO)
    rev_no = Trim(I_REV_NO)
    invoice_no = Replace(Trim(C_INV_NO), "-", "")
    custom_cd = Trim(Right(C_CUSTOM_CD, 10))
    output_dt = Format(SHIP_DATE, "yyyymmdd")
        
    sql = " select distinct order_no, nvl(invoice_no,' ' ) invoice_no , nvl(sid_no,' ' ) sid_no, nvl(custom_line_cd,' ') custom_line_cd, ipart_no, Prev_No" & vbCrLf
    sql = sql & " from mkt_out_1 a , mkt_out_11 b" & vbCrLf
    sql = sql & " Where a.outorder_no = b.outorder_no" & vbCrLf
    sql = sql & " and a.po_no ='" & po_no & "'" & vbCrLf
    sql = sql & " and a.output_dt ='" & output_dt & "'" & vbCrLf
    sql = sql & " and b.outlot_no ='" & lot_no & "'" & vbCrLf
         
    ssTemp.Open sql, DB_User
    
    If ssTemp.RecordCount <= 0 Then
        ssTemp.Close
        CHECK_OUTData = "출하지시된 정보가 없습니다."
        Exit Function
    End If
                
    isError = False
    
    If ssTemp.RecordCount = 1 Then
    
        If Trim(item_no) <> Trim(ssTemp!sid_no) Then
            strMsg = strMsg & "ITEM#가 일치하지 않습니다. (출하지시 정보:" & ssTemp!sid_no & ")" & vbCrLf
            isError = True
        End If
        
        If Trim(part_no) <> Trim(ssTemp!ipart_no) Then
            strMsg = strMsg & "PART#가 일치하지 않습니다. (출하지시 정보:" & ssTemp!ipart_no & ")" & vbCrLf
            isError = True
        End If
        
        If Trim(rev_no) <> Trim(ssTemp!Prev_No) Then
            strMsg = strMsg & "REV#가 일치하지 않습니다. (출하지시 정보:" & ssTemp!Prev_No & ")" & vbCrLf
            isError = True
        End If
        
        If Trim(line_code) <> "" And Trim(line_code) <> Trim(ssTemp!custom_line_cd) Then
            strMsg = strMsg & Label1(22).Caption & "가 일치하지 않습니다. (출하지시 정보:" & ssTemp!custom_line_cd & ")" & vbCrLf
            isError = True
        End If
                
        If Trim(invoice_no) <> "" And Trim(invoice_no) <> Trim(ssTemp!invoice_no) Then
            strMsg = strMsg & "Invoice가 일치하지 않습니다. (출하지시 정보:" & ssTemp!invoice_no & ")" & vbCrLf
            isError = True
        End If
    Else
        ssTemp.Close
        
        sql = " select distinct order_no, nvl(invoice_no,' ' ) invoice_no , nvl(sid_no,' ' ) sid_no, nvl(custom_line_cd,' ') custom_line_cd, ipart_no, Prev_No" & vbCrLf
        sql = sql & " from mkt_out_1 a , mkt_out_11 b" & vbCrLf
        sql = sql & " Where a.outorder_no = b.outorder_no" & vbCrLf
        sql = sql & " and a.po_no ='" & po_no & "'" & vbCrLf
        sql = sql & " and a.output_dt ='" & output_dt & "'" & vbCrLf
        sql = sql & " and b.outlot_no ='" & lot_no & "'" & vbCrLf
        sql = sql & " and b.ipart_no ='" & part_no & "'" & vbCrLf
        sql = sql & " and b.Prev_No ='" & rev_no & "'" & vbCrLf
        sql = sql & " and trim(b.sid_no) ='" & item_no & "'" & vbCrLf
        If Trim(line_code) <> "" Then sql = sql & " and trim(b.custom_line_cd) ='" & line_code & "'" & vbCrLf
        If Trim(invoice_no) <> "" Then sql = sql & " and trim(b.invoice_no) ='" & invoice_no & "'" & vbCrLf
                     
        ssTemp.Open sql, DB_User
        
        If ssTemp.RecordCount <= 0 Then
            ssTemp.Close
            CHECK_OUTData = "입력한 정보로 출하지시된 정보가 없습니다."
            Exit Function
        End If
    End If
    
    ssTemp.Close
    
    If isError = True Then
        returnStr = "입력한 정보와 출하지시 정보가 일치하지 않습니다." & vbCrLf
        returnStr = returnStr & "------------------------------------------------" & vbCrLf
        returnStr = returnStr & strMsg
        returnStr = returnStr & "------------------------------------------------" & vbCrLf
        CHECK_OUTData = returnStr
    Else
        CHECK_OUTData = ""
    End If
    
    Exit Function
    
ERROR_RTN:
    CHECK_OUTData = "출하지시 정보 확인중 오류가 발생했습니다. IT센터에 문의하세요"
End Function
Private Sub CMD_Save_Click()
    Dim i               As Integer
    Dim j               As Integer
    
    Dim Style           As String
    Dim sql             As String
    Dim CheckMsg        As String
    
    If C_LOT_NO = "" Then
        MsgBox ("LOT NO not input")
        Exit Sub
    ElseIf C_PO_NO = "" Then
        MsgBox ("PO NO not input")
        Exit Sub
    ElseIf  I_PART_NO = "" Then
        MsgBox ("Part NO not input")
        Exit Sub
    ElseIf I_REV_NO = "" Then
        MsgBox ("REV NO not input")
        Exit Sub
    ElseIf I_METERIAL_NM = "" Then
        MsgBox ("METERIAL_NM not input")
        Exit Sub
    ElseIf I_COLOR_NM = "" Then
        MsgBox ("COLOR not input")
        Exit Sub
    ElseIf C_ITEM_NO = "" Then
        Msg = "Item No not input,"                     ' Define message.
        Msg = Msg & vbCrLf & " Delete Data?"
        Style = vbYesNo + vbCritical + vbDefaultButton2     ' Define buttons.
        Title = "Delete Data"                               ' Define title.
     
        If MsgBox(Msg, Style, Title) = vbNo Then
            C_ITEM_NO.SetFocus
            Exit Sub
        End If
    ElseIf I_DESC_NM = "" Then
        MsgBox ("Description not input")
        Exit Sub
    ElseIf I_END_CT_NO = "" Or I_TOTAL_CT_NO = "" Then
        MsgBox ("Label Count not input")
        Exit Sub
    End If
    
    '제품 출하지시 정보 확인 (출하지시 정보와 일치하지 않을 경우 메세지 표시)
   ' CheckMsg = CHECK_OUTData
   ' If CheckMsg <> "" Then
   '     If MsgBox(CheckMsg & vbCrLf & "저장하시겠습니까?", vbYesNo, "Info") = vbNo Then
   '         Exit Sub
   '     End If
   ' End If
    
    '=============================================================
    ' Default Type에서
    ' Tray와 EndCap, Customer에 따라 D_Barcode를 달리 Pront 한다.
    '=============================================================
    sql$ = " select distinct a.prod_cd, a.prod_gubun, "
    sql$ = sql$ + " DECODE(NVL((select distinct mold_no from prd_work_11 where comp_cd = '" & GLO_COMP_CD & "' and bsns_cd = '" & GLO_BSNS_CD & "' and  outlot_no ='" & strLotNo & "' and iPART_NO = '" &  I_PART_NO & "' and Prev_No ='" & I_REV_NO & "' and trim(mold_no) <> '' ), ' '), ' ',"
    sql$ = sql$ + "   (select distinct mold_no from CIM_MACHI_LOT_INFO where comp_cd = '" & GLO_COMP_CD & "' and bsns_cd = '" & GLO_BSNS_CD & "' and lot_no ='" & strLotNo & "' and iPART_NO = '" &  I_PART_NO & "' and Prev_No ='" & I_REV_NO & "'),"
    sql$ = sql$ + "   (select distinct mold_no from prd_work_11 where comp_cd = '" & GLO_COMP_CD & "' and bsns_cd = '" & GLO_BSNS_CD & "' and outlot_no ='" & strLotNo & "' and iPART_NO = '" &  I_PART_NO & "' and Prev_No ='" & I_REV_NO & "' and trim(mold_no) <> '' ) )  mold_no"
    sql$ = sql$ + " from PRD_MAST_1_VIEW a "
    sql$ = sql$ + " where a.iPART_NO = '" &  I_PART_NO & "' "
    sql$ = sql$ + " and   a.Prev_No = '" & I_REV_NO & "' "
    sql$ = sql$ & " and a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSCODE = Nothing
    SSCODE.Open sql$, DB_User
    
    If SSCODE.RecordCount > 0 Then
        varMoldNo = Trim(IIf(IsNull(SSCODE!mold_no) = True, "", SSCODE!mold_no)) 'Recycle 입고의 경우 Mold No를 알수 없음
        If varMoldNo <> "" Then
            varMoldNo = Gf_MoldDesc_Load(varMoldNo)     'Mold Desc
        End If
    Else
        varMoldNo = ""
    End If
    
    Select Case SSCODE!prod_gubun
        Case "TR"
            D_BARCODE_TX = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & I_REV_NO +  I_PART_NO
        Case "TE"
            D_BARCODE_TX = "00" + I_REV_NO +  I_PART_NO
        Case "TC"
        
    End Select
    SSCODE.Close
        
    'WW 업체별 선택
    If varWWCustomCD = "20200" Then ' AMKOR  WWCUSTOM_CD :20200 2010-07-05 jwkim
        If Trim(Right(C_CUSTOM_CD, 10)) = "10001" Or Trim(Right(C_CUSTOM_CD, 10)) = "10011" Or Trim(Right(C_CUSTOM_CD, 10)) = "10003" Then 'ATK-1, ATK-3, ATK-4 고객만 해당 양식 사용 2014.01.07 S+G kimsi
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_ATK_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
        Else
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
        End If
    ElseIf varWWCustomCD = "22960" Then ' NANYA  WWCUSTOM_CD :22960 2018-05-23
        If PRT_FORM(1).Value = True Then                               'InnerLabel 프린트
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call TR_NANYA_InnerLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
        ElseIf PRT_FORM(0).Value = True Then                           'BoxLabel 프린트
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call TR_NANYA_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
        End If
    ElseIf varWWCustomCD = "21665" Then     'SPIL  WWCUSTOM_CD :21665 WWC 통합적용 2020.05.27 by CPL_orim
        For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
            For j = 1 To Val(I_SET_QT)
                Call T4_SPIL_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                Label_Cnt = Label_Cnt + 1
                Label_Data(Label_Cnt) = sFormatTx
            Next j
        Next i
    Else
        'SITE 별 선택
        Select Case Trim(Right(C_CUSTOM_CD, 10))
        
        Case "20180", "21844", "10507", "21832", "10519", "10536", "10558"  '2014-04-07 TSMC (TAIWAN SEMICONDUCTOR MANUFACTURING CO.,L) 라벨 by ICS_jwkim&IS_Junmo, 2018-11-06 add Sites TSMC FAB2, TSMC FAB30 by Jim Chen
             If C_INV_NO = "" And GLO_COMP_CD = "FU" Then   '20181008 jeenie & liam by FRCUNION
                MsgBox IIf(GLO_KOR_VISIBLE, "INVOICE NO를 입력하지 않았습니다", "You have not entered INVOICE NO"), vbExclamation
                Exit Sub
            End If
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call TR_TSMC_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
        Case "20172", "21876" 'SCC, SCCJ 2015.07.29 S+G Ella Kim(김새롬) 신규양식 요청 건
            If PRT_FORM(0).Value = True Then                                    'BoxLabel 프린트
                For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                    For j = 1 To Val(I_SET_QT)
                        Call T4_SCC_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                        Label_Cnt = Label_Cnt + 1
                        Label_Data(Label_Cnt) = sFormatTx
                    Next j
                Next i
            ElseIf PRT_FORM(1).Value = True Then                               'InnerLabel 프린트 한장에 같이 출력 (1/2씩) 같은내용 반복
                For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                    For j = 1 To Val(I_SET_QT)
                        Call T4_SCC_InnerLabel_Print(i, Val(I_TOTAL_CT_NO))
                        Label_Cnt = Label_Cnt + 1
                        Label_Data(Label_Cnt) = sFormatTx
                    Next j
                Next i
            End If
            
        Case "20038" 'SCS 2017-03-08 S+G Ella Kim(김새롬) 신규양식 요청 건
            If PRT_FORM(0).Value = True Then                                    'BoxLabel 프린트
                For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                    For j = 1 To Val(I_SET_QT)
                        Call T4_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                        Label_Cnt = Label_Cnt + 1
                        Label_Data(Label_Cnt) = sFormatTx
                    Next j
                Next i
            ElseIf PRT_FORM(1).Value = True Then                               'InnerLabel 프린트 한장에 같이 출력 (1/2씩) 같은내용 반복
                For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                    For j = 1 To Val(I_SET_QT)
                        Call T4_SCS_InnerLabel_Print(i, Val(I_TOTAL_CT_NO))
                        Label_Cnt = Label_Cnt + 1
                        Label_Data(Label_Cnt) = sFormatTx
                    Next j
                Next i
            End If
'        Case "20289", "21311", "21316" 'SPIL TW, SPIL SZ  '2013-05-13 by ICS_jwkim & S+G Aiden Kim , 2015.07.29 S+G Ella Kim(김새롬) 추가 변경 요청 반영 (출하LOT), 2019.12.04 FRC CPL Liam request SPIL DF label
'            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
'                For j = 1 To Val(I_SET_QT)
'                    Call T4_SPIL_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
'                    Label_Cnt = Label_Cnt + 1
'                    Label_Data(Label_Cnt) = sFormatTx
'                Next j
'            Next i
        Case "21669" 'SFA KOREA  '2014-11-26 by S+G kimsi
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_STS_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
        
       Case "20209"    'King Yuan Electronics Co.,LTD.   '2014-10-20 by S+G kimsi
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call TR_KYEC_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
            
        Case "10426", "10005" 'SCK  '2011-12-21 by ICS_jwkim & S+G jhbaek , JSCK 2016.04.01 요청에 따라 추가
        
            If Trim(C_LOT_ID) = "" Then
                If MsgBox("출하 LOT ID가 없습니다. 영업팀에 문의 하시기 바랍니다. " & vbCrLf & "계속진행하시겠습니까?", vbYesNo) = vbNo Then Exit Sub
            End If
        
            If PRT_FORM(1).Value = True Then   'InnerLabel 프린트
               For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                   For j = 1 To Val(I_SET_QT)
                       Call TR_SCK_innerLabel_Print(i, Val(I_TOTAL_CT_NO))
                       Label_Cnt = Label_Cnt + 1
                       Label_Data(Label_Cnt) = sFormatTx
                   Next j
               Next i
            Else
               For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                    For j = 1 To Val(I_SET_QT)
                        Call T4_SCK_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                        Label_Cnt = Label_Cnt + 1
                        Label_Data(Label_Cnt) = sFormatTx
                    Next j
                Next i
            End If
                       
        Case "20690" 'Renesas JPN  '2013-05-13 by ICS_jwkim & S+G Aiden Kim
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_Renesas_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
            
        Case "20684"    'ST-Micro(Malta)   '2011-12-21 by ICS_jwkim & S+G jhbaek
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_ST_MALTA_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
            
        Case "20015", "20163"  'ASEK, ASET , ASE(S)   '2011-12-26 by ICS_jwkim & S+G jhbaek
            If Trim(C_LOT_ID) = "" Then
                If MsgBox("RT No가 없습니다. 계속진행하시겠습니까?", vbYesNo) = vbNo Then Exit Sub
            End If
                
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_ASEK_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
        
        Case "20233" '2021-06-30 ASE(S)BY FRC Jim Chen
            If Trim(C_LOT_ID) = "" Then
                If MsgBox("RT No가 없습니다. 계속진행하시겠습니까?", vbYesNo) = vbNo Then Exit Sub
            End If
                
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_ASE_SGP_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
            
        Case "20779", "20276"  'ASECL   '2014-01-22 by S+G kimsi, 'ASE MATL CL (20276) 2020-02-17 by FRC IT Jim Chen & CPL Liam Bae
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_ASECL_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
            
        Case "20562"    'TITL 2017-03-07 by srkim
        
            If PRT_FORM(0).Value = True Then                                    'BoxLabel 프린트
                For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                    For j = 1 To Val(I_SET_QT)
                        Call T4_TITL_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                        Label_Cnt = Label_Cnt + 1
                        Label_Data(Label_Cnt) = sFormatTx
                    Next j
                Next i
            ElseIf PRT_FORM(1).Value = True Then                               'InnerLabel 프린트 한장에 같이 출력 (1/2씩) 같은내용 반복
                For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                    For j = 1 To Val(I_SET_QT)
                        Call T4_TITL_InnerLabel_Print(i, Val(I_TOTAL_CT_NO))
                        Label_Cnt = Label_Cnt + 1
                        Label_Data(Label_Cnt) = sFormatTx
                    Next j
                Next i
            End If
            
        Case "21979", "20879", "22204" And GLO_COMP_CD = "FU"  'MICRON TAIWAN   '2019-11-08 by FRC IT Jim Chen & CPL Liam Bae, 2022-01-06 Kangawa Liam
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_MICRONTW_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
        
        Case "20910" And GLO_COMP_CD = "FU"  'MICRON SINGAPORE   '2021-06-11 by FRC IT Jim Chen & CPL Liam Bae
        For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
            For j = 1 To Val(I_SET_QT)
                Call T4_Micron_SGP_BoxLabel_NEW_Print(i, Val(I_TOTAL_CT_NO))
                Label_Cnt = Label_Cnt + 1
                Label_Data(Label_Cnt) = sFormatTx
            Next j
        Next i
        
        Case "22206" And GLO_COMP_CD = "FU"  'MICRON Malasia Sdn Bhd   '2021-06-11 by FRC IT Jim Chen & CPL Liam Bae
        For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
            For j = 1 To Val(I_SET_QT)
                Call T4_Micron_SGP_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                Label_Cnt = Label_Cnt + 1
                Label_Data(Label_Cnt) = sFormatTx
            Next j
        Next i
        
        Case "20921" And GLO_COMP_CD = "FU"  'Walton   '2021-06-18 by FRC IT Jim Chen & CPL Liam Bae
        For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
            For j = 1 To Val(I_SET_QT)
                Call T4_Walton_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                Label_Cnt = Label_Cnt + 1
                Label_Data(Label_Cnt) = sFormatTx
            Next j
        Next i
            
        Case "20860" And GLO_COMP_CD = "FU"  'POWERTECH(T)20860   '2020-05-14 by FRC IT Jim Chen & CPL Liam Bae
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_POWERTECH_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
            Next j
        Next i
            
         Case "20888", "22134" And GLO_COMP_CD = "FU"   'ASKSH 2019-11-08 by FRC Jim Chen & CPL Liam Seo
            If Trim(C_LOT_ID) = "" Then
                If MsgBox("RT No hasn't been registered? Continue?", vbYesNo) = vbNo Then Exit Sub
            End If
            
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_ASEK_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                    Call T4_ASESH_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i

        Case "21347" And GLO_COMP_CD = "FU"  'ASEKS   '2019-11-08 by FRC IT Jim Chen & CPL Liam Bae
            If Trim(C_LOT_ID) = "" Then
                If MsgBox("RT No hasn't been registered? Continue?", vbYesNo) = vbNo Then Exit Sub
            End If
                
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_ASEKS_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
                Next j
            Next i
            
        Case "20908", "21547", "21987" And GLO_COMP_CD = "FU"  'SIGURD 20908, 21547, 21987 '2022-07-18 by FRC IT JJ Cheng & CPL Liam Bae
            For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                For j = 1 To Val(I_SET_QT)
                    Call T4_SIGUARD_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                    Label_Cnt = Label_Cnt + 1
                    Label_Data(Label_Cnt) = sFormatTx
            Next j
        Next i
            
        Case Else   'Default
            Select Case True
                Case PRT_OP(0).Value
                    '국내
                    For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                        For j = 1 To Val(I_SET_QT)
                            Call T4_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                            Label_Cnt = Label_Cnt + 1
                            Label_Data(Label_Cnt) = sFormatTx
                        Next j
                    Next i
                Case PRT_OP(1).Value
                    For i = Val(I_START_CT_NO) To Val(I_END_CT_NO)
                        For j = 1 To Val(I_SET_QT)
                            Call T4_BoxLabel_Print(i, Val(I_TOTAL_CT_NO))
                            Label_Cnt = Label_Cnt + 1
                            Label_Data(Label_Cnt) = sFormatTx
                        Next j
                    Next i
            End Select
         End Select
    End If
    
    D_TOT_LABEL_CNT = Format(Label_Cnt, "###0")
    CMD_PRINT.Enabled = True
End Sub

Private Function MKT_LABEL_SEQ(wwcustomCD As String, customCD As String, LotNo As String, ShipDT As String) As Integer
    Dim sql             As String
    Dim RecordsAffected As Integer
    
    Dim LotSeq          As Integer
    
    sql$ = "select lot_seq from MKT_LABEL_SEQ "
    sql$ = sql$ & " where wwcustom_cd = '" & wwcustomCD & "'"
    sql$ = sql$ & " and   custom_cd = '" & customCD & "' "
    sql$ = sql$ & " and   lot_No = '" & LotNo & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    If SSDATA.RecordCount <= "0" Then
        LotSeq = 1
        sql$ = "insert into MKT_LABEL_SEQ (comp_cd, bsns_cd, WWCUSTOM_CD, CUSTOM_CD, LOT_NO, LOT_SEQ,JOB_SEQ,SHIP_DT,REGIST_DT, EMP_NO,IP_ADDR) "
        sql$ = sql$ & " values ("
        sql$ = sql$ & " '" & GLO_COMP_CD & "', "                 'comp_cd
        sql$ = sql$ & " '" & GLO_BSNS_CD & "', "                   'bsns_cd
        sql$ = sql$ & " '" & wwcustomCD & "',"
        sql$ = sql$ & " '" & customCD & "',"
        sql$ = sql$ & " '" & LotNo & "',"
        sql$ = sql$ & " '" & LotSeq & "',"
        sql$ = sql$ & " 0,"
        sql$ = sql$ & " '" & ShipDT & "',"
        sql$ = sql$ & " sysdate,"
        sql$ = sql$ & " '" & GLO_USER_ID & "',"
        sql$ = sql$ & " '" & GLO_LOCAL_IP & "')"
        
        DB_User.Execute sql$, RecordsAffected
        
        If RecordsAffected = 0 Then
            MsgBox "Box Count error ", vbExclamation, "Error!!"
            
            MKT_LABEL_SEQ = 0
            Exit Function
        End If
    Else
        LotSeq = SSDATA.Fields(0) + 1
            
        sql$ = "update MKT_LABEL_SEQ set "
        sql$ = sql$ & "lot_seq = '" & LotSeq & "' "
        sql$ = sql$ & "where wwcustom_cd = '" & wwcustomCD & "'"
        sql$ = sql$ & "and   custom_cd = '" & customCD & "' "
        sql$ = sql$ & "and   lot_No = '" & LotNo & "'"
        sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
        sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
        DB_User.Execute sql$, RecordsAffected
    
        If RecordsAffected = 0 Then
            MsgBox "Box Count error ", vbExclamation, "Error!!"
            
            MKT_LABEL_SEQ = 0
            Exit Function
        End If
    End If
    SSDATA.Close
                    
    MKT_LABEL_SEQ = LotSeq
End Function

Private Function CHECK_INPUT_DATA(varPartNo As String, varRevNo As String, varLotNo As String) As Boolean
    ' Lot No/Part No/Rev이 모두 입력된 경우 해당 값이 일치하는지 확인
    ' 3개 모두 입력되고 일치하면, True
    ' 3개 모두 입력되고, 값이 일치하지 않을때는 False
    ' 2005. 11. 3 requested 김영호, programed hklee
    Dim sql         As String
    Dim ssTemp      As New ADODB.Recordset
    Dim i           As Integer
    
    On Error GoTo ERROR_RTN
    
    CHECK_INPUT_DATA = True
    If varPartNo = "" Then Exit Function
    If varRevNo = "" Then Exit Function
    If varLotNo = "" Then Exit Function
      
    CHECK_INPUT_DATA = False

    sql$ = " select distinct a.outlot_no, a.iPART_NO, a.Prev_No, a.prod_cd, a.prod_gubun " & vbCrLf
    sql$ = sql$ & " from PRD_WORK_11 a " & vbCrLf
    sql$ = sql$ & " where a.outlot_no ='" & Replace(varLotNo, "-", "") & "' " & vbCrLf
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    sql$ = sql$ & " Union " & vbCrLf
    sql$ = sql$ & " select distinct lot_no as outlot_no, ipart_no, Prev_No, prod_cd, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct f.prod_gubun from PRD_MAST_1 f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.prod_cd = a.prod_cd), ' ') prod_gubun " & vbCrLf
    sql$ = sql$ & " From CIM_MACHI_LOT_INFO a " & vbCrLf
    sql$ = sql$ & " where lot_no = '" & Replace(varLotNo, "-", "") & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf

    Set ssTemp = Nothing
    ssTemp.Open sql$, DB_User

    If ssTemp.RecordCount = 0 Then
        ssTemp.Close
        GoTo ERROR_RTN
    End If

    For i = 0 To ssTemp.RecordCount - 1
        If varPartNo = ssTemp!ipart_no And varRevNo = ssTemp!Prev_No Then
            CHECK_INPUT_DATA = True
            Exit For
        End If

        ssTemp.MoveNext
    Next i
    ssTemp.Close
        
    Exit Function
    
ERROR_RTN:
    CHECK_INPUT_DATA = False
    MsgBox IIf(GLO_KOR_VISIBLE, "Lot#의 입력이 Part No, Rev과 맞지 않습니다." & vbCrLf & "Lot#, Part No, Rev을 확인하시기 바랍니다.", _
                                "Lot# is not correct with Part No&Rev." & vbCrLf & "Check Lot#, Part No, Rev.")
End Function

Private Sub CMD_SAVE_GotFocus()
    Call FIELD_VERIFY
End Sub

Private Sub Form_Load()
    Dim sIniFile    As String
    Dim sql         As String

    Screen.MousePointer = 11
    
    Top = 1000 '(Screen.Height - Me.Height) / 3
    Left = 3000 '(Screen.Width - Me.Width) / 2
    Call Gs_Form_Caption(Me)
    Call Form_Control_Search
    
    Call PORT_OPEN

    PRT_OP(1).Value = True
    Call PRT_OP_GotFocus(1)
    
    ' Shipping Date 표시
    SHIP_DATE = Format(Curr_Date, "&&&&-&&-&&")
    
    Call Field_Clear
    C_PO_NO.Clear
    C_INV_NO.Clear
    CT_PRINT_CHK.Value = 1
    
    Timer1.Enabled = False
    Label_Cnt = 0
    D_TOT_LABEL_CNT = Format(Label_Cnt, "#,##0")
    ReDim Label_Data(3000)
    
    '==================================
    ' AMD Shipping Label
    '==================================
    Label1(15).Visible = False
    Label1(16).Visible = False
    I_TOTAL_WEIGHT.Visible = False
    I_PCNT_NO.Visible = False
    I_SET_QT = 1
    CMD_PRINT.Enabled = False
    Label1(22).Visible = False
    C_LOT_ID.Visible = False
    
    GLO_COMP_CO = IIf(GLO_COMP_CD = "FU", "MADE IN TAIWAN", "MADE IN KOREA")    'Country of Original 변경
    
    'Hide Shippment option fram by default
    Frame_Ship.Visible = False
    
    'Shippment Option Default on FRC
    Option1.Value = True
    
    Title.Caption = IIf(GLO_KOR_VISIBLE, "BARCODE LABEL 출력", "BARCODE LABEL PRINT")
    FramePRT.Caption = IIf(GLO_KOR_VISIBLE, "출력구분", "Print Sort")
    FrameForm.Caption = IIf(GLO_KOR_VISIBLE, "출력양식", "Print Form")
    PRT_OP(0).Caption = IIf(GLO_KOR_VISIBLE, "국내", "Dome..")
    PRT_OP(1).Caption = IIf(GLO_KOR_VISIBLE, "국외", "Fore..")
    CMD_CLEAR.Text = IIf(GLO_KOR_VISIBLE, "초기화", "Initialize")
    CMD_SAVE.Text = IIf(GLO_KOR_VISIBLE, "Label내용 저장", "Save Label")
    CMD_PRINT.Text = IIf(GLO_KOR_VISIBLE, "LABEL 내역 일괄 출력", "Label Batch Print")
    Label1(17) = IIf(GLO_KOR_VISIBLE, "", "Label Q'ty")
    
    If GLO_DEPT_CD = "10011" Then   '현재는 IT Center만 LanType으로 진행가능 (테스트라벨 출력 용도)
        ChkLANPRT.Visible = True
        ChkLANPRT.Value = 0
    Else
        ChkLANPRT.Visible = False
        ChkLANPRT.Value = 0
    End If
        
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call PORT_CLOSE
    Call GF_SocketClose(Lantype_Sck)
End Sub

Private Sub I_TOTAL_CT_NO_GotFocus()
    Call FIELD_VERIFY
End Sub

Private Sub C_LOT_NO_GotFocus()
    Call FIELD_VERIFY
End Sub

Private Sub C_LOT_NO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub C_LOT_NO_LostFocus()
    Dim sql     As String
        
    ' 입력한 Lot#를 이용 Part#, Rev을 Check 한다.
    If CHECK_INPUT_DATA( I_PART_NO, I_REV_NO, C_LOT_NO) = False Then
        C_LOT_NO = ""
        MsgBox IIf(GLO_KOR_VISIBLE, "Lot#의 입력이 Part No, Rev과 맞지 않습니다." & vbCrLf & "Lot#, Part No, Rev을 확인하시기 바랍니다.", _
                                    "Lot# is not correct with Part No&Rev." & vbCrLf & "Check Lot#, Part No, Rev.")
        Exit Sub
    End If
        
    C_LOT_ID.ListIndex = C_LOT_NO.ListIndex 'LOT-ID 맞춰주기
    C_LOT_ID.Enabled = False
    
    strLotNo = cDll.Gf_UnFormat("M", C_LOT_NO)
    Call ComboLoad_PO_NO
End Sub

Private Sub ComboLoad_PO_NO()
    Dim sql         As String
    Dim ssTemp      As New ADODB.Recordset
    
    ' Customer를 Ship To 기준에서 From기준을 변경함.(2002/10/2)
    ' 예를 들어 AMD에서 Order한후 AMKOR로 Shipping 하더라도 AMD로 Label Print 한다. (2003. 7. 23 by hkle, hyjang)
    If CHK_INSMP.Value = 1 Then
        sql$ = " select a.insmp_po_no po_no from MKT_ORD_1 a "
        sql$ = sql$ & " where insmp_custom_cd = '" + Trim(Right(C_CUSTOM_CD, 10)) + "'"
    Else
        sql$ = " select a.po_no from MKT_ORD_1 a "
        sql$ = sql$ & " where custom_cd = '" + Trim(Right(C_CUSTOM_CD, 10)) + "'"
    End If
    sql$ = sql$ & " and   comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and   bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    sql$ = sql$ & " and   order_no in (select b.order_no from MKT_OUT_1 b, MKT_OUT_11 c "
    sql$ = sql$ & "                    where c.outorder_no = b.outorder_no "
    If C_LOT_ID.Visible = True And C_LOT_ID <> "" Then
        If Trim(Right(C_CUSTOM_CD, 10)) = "20289" Or Trim(Right(C_CUSTOM_CD, 10)) = "21311" Then 'SPIL TW (outorder_no)
            sql$ = sql$ & "             and c.outorder_no = '" & C_LOT_ID & "'"
        Else 'custom_line_cd
            sql$ = sql$ & "             and c.custom_line_cd = '" & C_LOT_ID & "'"
        End If
    End If
    sql$ = sql$ & "                    and   b.comp_cd = c.comp_cd and b.bsns_cd = c.bsns_cd "
    sql$ = sql$ & "                    and   c.outlot_no = '" + strLotNo + "' and b.output_dt = '" & Format(SHIP_DATE, "yyyymmdd") & "') "
            
    Set ssTemp = Nothing
    ssTemp.Open sql$, DB_User
    
    With C_PO_NO
        .Clear
        If ssTemp.RecordCount = 0 Then
            ssTemp.Close
            .ListIndex = -1
        Else
            Do Until ssTemp.EOF
                .AddItem Trim(ssTemp!po_no)
                ssTemp.MoveNext
            Loop
            ssTemp.Close
            .ListIndex = 0
        End If
    End With
    
End Sub

Private Sub  I_PART_NO_LostFocus()
    Dim sql         As String
    
    ' 입력한 Part#를 이용 Lot#, Rev을 Check 한다.
     I_PART_NO = Trim( I_PART_NO)
    If  I_PART_NO = "" Then Exit Sub
    
    sql$ = "select * from PRD_MAST_1_VIEW "
    sql$ = sql$ + " where ipart_no = '" +  I_PART_NO + "' "
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User

    If SSDATA.RecordCount = 0 Then
        SSDATA.Close
        Msg = GLO_MSGLINE(6)
        MsgBox IIf(GLO_KOR_VISIBLE, "등록된 코드가 아닙니다. 다시 확인하십시요", "It is not registered code. Check it again.")
        Field_Check_SW = False
        Exit Sub
    End If
    SSDATA.Close
    
    If CHECK_INPUT_DATA( I_PART_NO, I_REV_NO, C_LOT_NO) = False Then
         I_PART_NO = ""
        MsgBox IIf(GLO_KOR_VISIBLE, "Lot#의 입력이 Part No, Rev과 맞지 않습니다." & vbCrLf & "Lot#, Part No, Rev을 확인하시기 바랍니다.", _
                                    "Lot# is not correct with Part No&Rev." & vbCrLf & "Check Lot#, Part No, Rev.")
        Exit Sub
    End If
End Sub

Private Sub I_REV_NO_GotFocus()
    Call FIELD_VERIFY
End Sub

Private Sub I_REV_NO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub I_TOTAL_WEIGHT_LostFocus()
    I_TOTAL_WEIGHT = Trim(I_TOTAL_WEIGHT)
End Sub




Private Sub PRT_FORM_Click(Index As Integer)
    If Index = 1 Then   'Inner Label 선택시
        Label1(24).Visible = True
        I_BUNDLE_QT.Visible = True
        I_BUNDLE_QT.BackColor = &HC0FFFF
        I_BUNDLE_QT.Enabled = True
        I_BUNDLE_QT.Text = "2"  '기본 Bundle 수량 (영업요청사항)
        If (Trim(Right(C_CUSTOM_CD, 5)) = "10426" Or Trim(Right(C_CUSTOM_CD, 5)) = "10005") And Trim(Right(D_PROD_KB, 10)) = "TR" Then I_SET_QT = "3" 'SCK Tray의 경우 셋트수량을 기본 3으로 지정 DS_YHKIM 요청 사항
    Else
        Label1(24).Visible = False
        I_BUNDLE_QT.Visible = False
        I_BUNDLE_QT.Text = "0"
    End If
End Sub

Private Sub PRT_OP_Click(Index As Integer)
    Call PRT_OP_GotFocus(Index)
End Sub

Private Sub PRT_OP_GotFocus(Index As Integer)
    Dim sql         As String
    Dim ssTemp      As New ADODB.Recordset
    
    Call CMD_Clear_Click
    
    If Index = 0 Then
        PRT_OP(0).Value = True
        I_SET_QT = 1

        sql$ = "select custom_cd, custom_en  custom_nm, custom_eyn custom_yn "
        sql$ = sql$ & " from MKT_CUST_1 a"
        sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
        sql$ = sql$ & " and   a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        'sql$ = sql$ & " and   cust_sal = '11' "
        sql$ = sql$ & " and   a.nation_kb in (select g.cp_nation_kb from SYS_COMPANY_1 g where g.comp_cd = a.comp_cd and g.bsns_cd = a.bsns_cd)" & vbCrLf
        sql$ = sql$ & " order by custom_yn "
    Else
        PRT_OP(1).Value = True
        I_SET_QT = 1
         
        sql$ = "select custom_cd, custom_en custom_nm, custom_eyn custom_yn "
        sql$ = sql$ & " from MKT_CUST_1 a"
        sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
        sql$ = sql$ & " and   a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        'sql$ = sql$ & " and   cust_sal = '12' "
        sql$ = sql$ & " and   a.nation_kb not in (select g.cp_nation_kb from SYS_COMPANY_1 g where g.comp_cd = a.comp_cd and g.bsns_cd = a.bsns_cd)" & vbCrLf
        sql$ = sql$ & " order by custom_yn "
    End If

    Set ssTemp = Nothing
    ssTemp.Open sql$, DB_User

    With C_CUSTOM_CD
        .Clear
        If ssTemp.RecordCount > 0 Then
            Do Until ssTemp.EOF
                .AddItem Trim(ssTemp!custom_yn) + Space(60 - Len(Trim(ssTemp!custom_yn))) + Trim(ssTemp!custom_nm) + Space(40) + ssTemp!custom_cd
                ssTemp.MoveNext
            Loop
            .ListIndex = 0
        End If
        ssTemp.Close
    End With

End Sub

Private Function LABEL_DB_SAVE()
    Dim sql             As String
    Dim RecordsAffected As Integer
    
    On Error GoTo ERROR_RTN
    
    LABEL_DB_SAVE = True
    
    sql$ = " select * from MAT_LABEL_2 "
    sql$ = sql$ & " where custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' "
    sql$ = sql$ & " and   invoice_no = '" & C_INV_NO & "' "
    sql$ = sql$ & " and   outlot_no = '" & C_LOT_NO & "' "
    sql$ = sql$ & " and   po_no = '" & C_PO_NO & "' "
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSCODE = Nothing
    SSCODE.Open sql$, DB_User

    If SSCODE.RecordCount = 0 Then
        sql$ = "Insert into MAT_LABEL_2 (comp_cd, bsns_cd, custom_cd, invoice_no, outlot_no, po_no, item_no, iPART_NO, Prev_No, "
        sql$ = sql$ & " item_nm, inbox_qty,  total_kg, pkg_count, print_dt, emp_no)"
        sql$ = sql$ & "  values (" & vbCrLf
        sql$ = sql$ & " '" & GLO_COMP_CD & "', "                 'comp_cd
        sql$ = sql$ & " '" & GLO_BSNS_CD & "', "                   'bsns_cd
        sql$ = sql$ & " '" & Trim(Right(C_CUSTOM_CD, 5)) & "', "
        sql$ = sql$ & " '" & C_INV_NO & "', "
        sql$ = sql$ & " '" & C_LOT_NO & "',"
        sql$ = sql$ & " '" & C_PO_NO & "', "
        sql$ = sql$ & " '" & C_ITEM_NO & "', "
        sql$ = sql$ & " '" &  I_PART_NO & "',"
        sql$ = sql$ & " '" & I_REV_NO & "', "
        sql$ = sql$ & " '" & I_DESC_NM & "', "
        sql$ = sql$ & " " & Val(I_BOXUNIT_NB) & ","
        sql$ = sql$ & " " & Val(I_TOTAL_WEIGHT) & ", " & Val(I_PCNT_NO) & ", "
        sql$ = sql$ & " Getdate(), "
        sql$ = sql$ & " '" & GLO_USER_ID & "') "
    
        DB_User.Execute sql$, RecordsAffected
        If RecordsAffected = 0 Then
            GoTo ERROR_RTN
        End If
    
    Else
        sql$ = " update MAT_LABEL_2 set "
        sql$ = sql$ & " inbox_qty = " & Val(I_BOXUNIT_NB) & ", "
        sql$ = sql$ & " total_kg = " & Val(I_TOTAL_WEIGHT) & ", "
        sql$ = sql$ & " pkg_count = " & Val(I_PCNT_NO) & "  "
        sql$ = sql$ & " where custom_cd = '" & Trim(Right(C_CUSTOM_CD, 5)) & "' "
        sql$ = sql$ & " and   invoice_no = '" & C_INV_NO & "' "
        sql$ = sql$ & " and   outlot_no = '" & C_LOT_NO & "' "
        sql$ = sql$ & " and   po_no = '" & C_PO_NO & "' "
        sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
        sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
        DB_User.Execute sql$, RecordsAffected
        If RecordsAffected = 0 Then
            GoTo ERROR_RTN
        End If
    End If
    SSCODE.Close
    
    Exit Function
    
ERROR_RTN:
    LABEL_DB_SAVE = False

End Function

Private Sub PRT_SET_Click(Index As Integer)
    Call PORT_CLOSE
    Select Case Index
    Case 0  'Com 1 Port
        Call PORT_OPEN(1)
    Case 1  'Com 3 Port
        Call PORT_OPEN(3)
    End Select
End Sub

Private Sub SHIP_DATE_Change()
    'Call Field_Clear
    Call C_CUSTOM_CD_LostFocus
End Sub

'=============> Label Case
Private Sub T4_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2011. 12. 21
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' 기본라벨형식 (STD)
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim charRowAdd(5 To 48)     As Integer
    Dim varDWG_no               As String
    Dim varCUSTDWG_no           As String
    Dim varPrdITEM_no           As String
    Dim varQApassDT             As String
    
    Dim varEXPDT                As String
    
    'QA Pass Date
    varQApassDT = Format(SHIP_DATE, "yyyy.mm.dd")
    
    varLot_no = C_LOT_NO

   
    varEXPDT = Format(DateAdd("yyyy", 3, SHIP_DATE), "MM-dd-YYYY")

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    '===========================================================================
    ' SMP 정보 확인, Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, nvl(a.cust_dwg_no,' ') as cust_dwg_no," & vbCrLf
    sql$ = sql$ & " (select distinct nvl(item_no, ' ') item_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as prd_item_no, " & vbCrLf
    sql$ = sql$ & " (select distinct nvl(dwg_no, ' ') dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as dwg_no" & vbCrLf
    sql$ = sql$ & " from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    varCUSTDWG_no = ""
    varPrdITEM_no = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
        varDWG_no = SSDATA!dwg_no
        varPrdITEM_no = SSDATA!prd_item_no
        varCUSTDWG_no = SSDATA!cust_dwg_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
    
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'CUST'R
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUST'R" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
    
   'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0240" + "0980"
    sFormatTx = sFormatTx + "CUST'R PO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + vbCr

    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0310" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0310" + "0730"
    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr
    
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '  I_PART_NO + vbCr
    
    'PEAK DWG#, PEAK DESC.
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0450" + "0980"
    sFormatTx = sFormatTx + "PEAK DESC." + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0450" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
    
    'DWG REV#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0980"
    sFormatTx = sFormatTx + "DWG REV#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiRev_NO + vbCr
    
    If Trim(Right(C_CUSTOM_CD, 10)) = "20301" Or Trim(Right(C_CUSTOM_CD, 10)) = "21458" Then  'ATT의 경우 DWG NO# 추가 by FRC jeenie & Liam / 2019.04.08 add Jim
        'DWG NO#
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
        sFormatTx = sFormatTx + "DWG NO#" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
        sFormatTx = sFormatTx + ": " + varCUSTDWG_no + vbCr  '2019.08.14 request by Liam  varPrdITEM_no >> varCUSTDWG_no
    
        'SHIP'G LOT NO
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0980"
        sFormatTx = sFormatTx + "SHIP'G LOT NO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0730"
        sFormatTx = sFormatTx + ": " + C_PO_NO + "-" + Format(SHIP_DATE, "mmddyy") + vbCr
        
        'QTY
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
        sFormatTx = sFormatTx + "QTY" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
        sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
        
        'COO
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0800" + "0980"
        sFormatTx = sFormatTx + "COO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0800" + "0730"
        sFormatTx = sFormatTx + ": " + GLO_COMP_CO + vbCr
        
        'QA Pass Date
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0840" + "0980" + "( QA PASS : " + varQApassDT + " )" + vbCr
        
    '2021-11-10 CPL Liam request KINSUS format
    ElseIf Trim(Right(C_CUSTOM_CD, 10)) = "10504" Or Trim(Right(C_CUSTOM_CD, 10)) = "21126" Then

        'SHIP'G LOT NO
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
        sFormatTx = sFormatTx + "SHIP'G LOT NO." + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
        sFormatTx = sFormatTx + ": " + C_PO_NO + "-" + Format(SHIP_DATE, "mmddyy") + vbCr
        
        'EXP Date
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0980"
        sFormatTx = sFormatTx + "EXP DATE" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0730"
        sFormatTx = sFormatTx + ": " + varEXPDT + vbCr
        
        'COO
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
        sFormatTx = sFormatTx + "COO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
        sFormatTx = sFormatTx + ": " + "MADE IN TAIWAN" + vbCr
        
        'QA Pass Date
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0800" + "0980" + "( QA PASS : " + varQApassDT + " )" + vbCr

    Else
        'SHIP'G LOT NO
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
        sFormatTx = sFormatTx + "SHIP'G LOT NO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
        sFormatTx = sFormatTx + ": " + C_PO_NO + "-" + Format(SHIP_DATE, "mmddyy") + vbCr
        
        'QTY
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0980"
        sFormatTx = sFormatTx + "QTY" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0730"
        sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
        
        If Trim(Right(C_CUSTOM_CD, 10)) = "21022" And varEndwwc_CD = "21355" Then  '2020.07.22 by CPL Orim ENDWWC_CD가 KYOCERA 인경우만 END CUST PO를 출력
            sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
            sFormatTx = sFormatTx + "KST PO#" + vbCr
            sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
            sFormatTx = sFormatTx + ": " + Trim(C_LOT_ID) + vbCr
            
            'QA Pass Date
            'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0800" + "0980" + "( QA PASS : " + varQApassDT + " )" + vbCr
        End If
                
        If GLO_COMP_CD <> "FU" Then '2018.05.09 Jeenie & Liam 요청으로 표준라벨에서 COO 항목을 제외한다.
            sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0800" + "0980"
            sFormatTx = sFormatTx + "COO" + vbCr
            sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0800" + "0730"
            sFormatTx = sFormatTx + ": " + GLO_COMP_CO + vbCr
        End If
        
        
        'QA Pass Date
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0840" + "0980" + "( QA PASS : " + varQApassDT + " )" + vbCr
        
    End If

    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub TR_KYEC_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2014. 10. 20
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' 기본라벨형식+EXPIRE DATE 출하일로부터 3년
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    Dim varQApassDT             As String

    
    Dim charRowAdd(5 To 48)     As Integer
    
        
    'QA Pass Date
    varQApassDT = Format(SHIP_DATE, "yyyy.mm.dd")
    
    varLot_no = C_LOT_NO
        
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    'EXP DATE (Ship 당일 날짜 + 3년) << 출하일 기준으로 생성요청 SNG_Kimsi
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, Trim(SHIP_DATE), Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    yyyy = Left(varEXPDT, 4)
    MM = Mid(varEXPDT, 6, 2)
    dd = Right(varEXPDT, 2)
    varEXPDT = MM & "/" & dd & "/" & yyyy

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If

    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================

    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If

'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0890" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
       '2D BARCODE(QR CODE)   'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로
   sFormatTx = sFormatTx & "3W1D44000072003002HM," & " " & Trim(C_ITEM_NO) & ";" & C_PO_NO & "-" & Format(SHIP_DATE, "mmddyy") & ";" & varUnitQty & ";" & Trim(varEXPDT) & ";" & C_PO_NO & Chr$(13)
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'CUST'R
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUST'R" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
    
   'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0240" + "0980"
    sFormatTx = sFormatTx + "CUST'R PO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + vbCr

    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0310" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0310" + "0730"
    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr
    
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '  I_PART_NO + vbCr
    
    'PEAK DWG#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0450" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0450" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
    
    'DWG REV#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0980"
    sFormatTx = sFormatTx + "DWG REV#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiRev_NO + vbCr
    
    'SHIP'G LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
    sFormatTx = sFormatTx + "SHIP'G LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + "-" + Format(SHIP_DATE, "mmddyy") + vbCr
    
    'EXPIRE DATE (MM/DD/YYYY)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0980"
    sFormatTx = sFormatTx + "EXPIRE DATE" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0730"
    sFormatTx = sFormatTx + ": " + varEXPDT + vbCr
    
    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
    
    '여백
    
    'COO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0800" + "0980"
    sFormatTx = sFormatTx + "COO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0800" + "0730"
    sFormatTx = sFormatTx + ": " + GLO_COMP_CO + vbCr
    
    
    'QA Pass Date
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0870" + "0980" + "( QA PASS : " + varQApassDT + " )" + vbCr
    
    
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0950" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0950" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    
    
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub TR_NANYA_InnerLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2014.04.24 / INNERLABEL
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim varSizeNm               As String   'Size
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varMatrNM             As String
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
            
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, varMFGDT, Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    If varUnitQty Mod Val(I_BUNDLE_QT) <> 0 Or Trim(I_BUNDLE_QT) = 0 Or Trim(I_BUNDLE_QT) = "" Then
        MsgBox "Bundle 수량을 재확인하세요.", vbExclamation
        I_BUNDLE_QT.SetFocus
        Exit Sub
    Else
        varUnitQty = IIf(PRT_FORM(1).Value = True, varUnitQty \ Val(I_BUNDLE_QT), varUnitQty)   'I_BOX_UNIT_QT를 Bundle 수량만큼 나눈다. (정수로 표시)
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
    '============Material Force PP13 into MPPO Jim Chen 2018/10/05 Request by PEAK KH Claire
    varMatrNM = IIf(GLO_COMP_CD = "FU" And Trim(I_METERIAL_NM) = "PP13", "MPPO", I_METERIAL_NM)
    
    '---------------------------------------------------------------------------
    ' Size 정보를 가져옴
    '---------------------------------------------------------------------------
    sql$ = "select distinct nvl(psize_nm, ' ') psize_nm from PRD_MAST_2 " & vbCr
    sql$ = sql$ + " where iPART_NO = '" &  I_PART_NO & "' " & vbCr
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf

    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varSizeNm = Trim(SSDATA!PSIZE_NM)
    SSDATA.Close
    '---------------------------------------------------------------------------

    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If

'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
    '2D BARCODE(QR CODE)   'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로
    sFormatTx = sFormatTx & "3W1D44000052003502HM," & " PEAK" & "," & Trim(C_ITEM_NO) & "," & varSizeNm & "," & varMatrNM & "," & Left(Replace(varMFGDT, "-", ""), 6) & "," & Left(Replace(varEXPDT, "-", ""), 6) & "," & varLot_no & "," & IIf(GLO_COMP_CD = "FU", sText1, varOEMiPART_NO) & "," & varUnitQty & Chr$(13)
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0920"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'Supplier Name
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "Supplier Name" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + "PEAK" + vbCr
    
   'Part No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0240" + "0980"
    sFormatTx = sFormatTx + "Part No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(C_ITEM_NO) + vbCr

    'Size
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0310" + "0980"
    sFormatTx = sFormatTx + "Size" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0310" + "0730"
    sFormatTx = sFormatTx + ": " + varSizeNm + vbCr
           
    'Material
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "Material" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0730"
    sFormatTx = sFormatTx + ": " + varMatrNM + vbCr  '
    
    'MFG Date (YYYYMM)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0450" + "0980"
    sFormatTx = sFormatTx + "MFG Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0450" + "0730"
    sFormatTx = sFormatTx + ": " + Left(Replace(varMFGDT, "-", ""), 6) + vbCr
    
    'EXP Date (YYYYMM)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0980"
    sFormatTx = sFormatTx + "EXP Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
    sFormatTx = sFormatTx + ": " + Left(Replace(varEXPDT, "-", ""), 6) + vbCr
    
    'SLI No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
    sFormatTx = sFormatTx + "SLI No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
    sFormatTx = sFormatTx + ": " + varLot_no + vbCr
    
    'Quantity
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0980"
    sFormatTx = sFormatTx + "Quantity" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
       
    If GLO_COMP_CD = "FU" Then ' By CPL Liam's request 2019.08.13
        'DESC   sText1 = I_DESC_NM
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
        sFormatTx = sFormatTx + "Item" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
        sFormatTx = sFormatTx + ": " + sText1 + vbCr
     Else
        'P/N
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
        sFormatTx = sFormatTx + "P/N" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
        sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr
    End If
'    If CT_PRINT_CHK.Value = 1 Then 'NANYA C/T # Print Function requested by FRC CPL Liam    '2019.04.11
'        'C/NO
'        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
'        sFormatTx = sFormatTx + "C/NO" + vbCr
'        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
'        sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
'    End If

    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub TR_NANYA_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2014.04.24 / BOXLABEL
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim varSizeNm               As String   'Size
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varMatrNM             As String
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, varMFGDT, Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
    '============Material Force PP13 into MPPO Jim Chen 2018/10/05 Request by PEAK KH Claire
    varMatrNM = IIf(GLO_COMP_CD = "FU" And Trim(I_METERIAL_NM) = "PP13", "MPPO", I_METERIAL_NM)
    
    '---------------------------------------------------------------------------
    ' Size 정보를 가져옴
    '---------------------------------------------------------------------------
    sql$ = "select distinct nvl(psize_nm, ' ') psize_nm from PRD_MAST_2 " & vbCr
    sql$ = sql$ + " where iPART_NO = '" &  I_PART_NO & "' " & vbCr
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf

    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varSizeNm = Trim(SSDATA!PSIZE_NM)
    SSDATA.Close
    '---------------------------------------------------------------------------

    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If

'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
    '2D BARCODE(QR CODE)   'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로
    sFormatTx = sFormatTx & "3W1D44000052003502HM," & " PEAK" & "," & Trim(C_ITEM_NO) & "," & varSizeNm & "," & varMatrNM & "," & Left(Replace(varMFGDT, "-", ""), 6) & "," & Left(Replace(varEXPDT, "-", ""), 6) & "," & varLot_no & "," & IIf(GLO_COMP_CD = "FU", sText1, varOEMiPART_NO) & "," & varUnitQty & Chr$(13)
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0920"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'Supplier Name
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "Supplier Name" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + "PEAK" + vbCr
    
   'Part No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0240" + "0980"
    sFormatTx = sFormatTx + "Part No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(C_ITEM_NO) + vbCr

    'Size
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0310" + "0980"
    sFormatTx = sFormatTx + "Size" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0310" + "0730"
    sFormatTx = sFormatTx + ": " + varSizeNm + vbCr
           
    'Material
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "Material" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0730"
    sFormatTx = sFormatTx + ": " + varMatrNM + vbCr  '
    
    'MFG Date (YYYYMM)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0450" + "0980"
    sFormatTx = sFormatTx + "MFG Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0450" + "0730"
    sFormatTx = sFormatTx + ": " + Left(Replace(varMFGDT, "-", ""), 6) + vbCr
    
    'EXP Date (YYYYMM)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0980"
    sFormatTx = sFormatTx + "EXP Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
    sFormatTx = sFormatTx + ": " + Left(Replace(varEXPDT, "-", ""), 6) + vbCr
    
    'SLI No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
    sFormatTx = sFormatTx + "SLI No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
    sFormatTx = sFormatTx + ": " + varLot_no + vbCr
    
    'Quantity
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0980"
    sFormatTx = sFormatTx + "Quantity" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
       
    If GLO_COMP_CD = "FU" Then ' By CPL Liam's request 2019.03.27
        'DESC   sText1 = I_DESC_NM
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
        sFormatTx = sFormatTx + "Item" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
        sFormatTx = sFormatTx + ": " + sText1 + vbCr
     Else
        'P/N
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
        sFormatTx = sFormatTx + "P/N" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
        sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr
    End If
       
    'C/NO
    If CT_PRINT_CHK.Value = 1 Then 'NANYA C/T # Print Function requested by FRC CPL Liam '2019.04.11
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
        sFormatTx = sFormatTx + "C/NO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
        sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    End If
    sFormatTx = sFormatTx + "E" + vbCr
    
End Sub

Private Sub T4_ST_MALTA_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2011. 12. 21
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' 업  체 : ST-Micro(Malta)
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion ratea
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'CUST'R
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUST'R" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
    
   'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0240" + "0980"
    sFormatTx = sFormatTx + "CUST'R PO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0730"
    sFormatTx = sFormatTx + ": " + "CONSIGNMENT" + vbCr 'C_PO_NO + vbCr

    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0310" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0310" + "0730"
    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr
    
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '  I_PART_NO + vbCr
    
    'PEAK DWG#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0450" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0450" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
    
    'DWG REV#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0980"
    sFormatTx = sFormatTx + "DWG REV#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiRev_NO + vbCr
    
    'SHIP'G LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
    sFormatTx = sFormatTx + "SHIP'G LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
    sFormatTx = sFormatTx + ": " + "CONSIGNMENT" + "-" + Format(SHIP_DATE, "mmddyy") + vbCr
    
    '여백
    
    'COO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0840" + "0980"
    sFormatTx = sFormatTx + "COO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0840" + "0730"
    sFormatTx = sFormatTx + ": " + GLO_COMP_CO + vbCr
    
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_ASEK_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2013. 05. 07
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' SCK용 2D BarCode (QR CODE)
    ' 업  체 : SCK
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim Peak_varLot_no          As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varPeakMFGDT            As String
    Dim varEXPDT                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMMold_no           As String
    Dim varOEMiRev_NO           As String
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, Trim(varMFGDT), Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
        
    'EXPIRED DATE duration change from 5 years to 3 years By Jim Chen 2018.09.19 Requested by CPL Liam
    'varEXPDT = Replace(DateAdd("yyyy", IIf(GLO_COMP_CD = "FU", 3, 5), varMFGDT), "-", "")
    'varEXPDT = Replace(DateAdd("yyyy", 5, varMFGDT), "-", "")   '2018.10.24 FRC 요청으로 5년 통일

    varMFGDT = Replace(varMFGDT, "-", "")
    varEXPDT = Replace(varEXPDT, "-", "")
    
    'PEAK Shipment Check for Lot# and DOM
    
    'If PEAK_Ship_Check = 1 Then
        'Peak_varLot_no = Format(Trim(SHIP_DATE), "mmddyy")
        'varPeakMFGDT = Left(varMFGDT, 6) & "01"
    'End If
    
    'PEAK Shipment Check for Lot# and DOM
    If Option2.Value = True Then
        Peak_varLot_no = Format(Trim(SHIP_DATE), "mmddyy")
        varPeakMFGDT = Left(varMFGDT, 6) & "01"
    'S+G Shipment Check for Lot#
    ElseIf Option3.Value = True Then
        Peak_varLot_no = Replace(Replace(C_LOT_NO, "-", ""), "F", "D")
        
    End If
    
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    varLot_no = Replace(C_LOT_NO, "-", "")
    
    '===========================================================================
    ' Mold No
    '===========================================================================
    sql$ = "select nvl(a.SMP_MOLD_NO,' ') as SMP_MOLD_NO from PRD_SMP_1 a " & vbCrLf
    sql$ = sql$ & " where a.mold_no in (select distinct mold_no from PRD_MAST_1 f " & vbCrLf
    sql$ = sql$ & "                     where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.mold_desc = '" & Trim(varMoldNo) & "' and f.Prev_No = a.rev_no and f.ipart_no = a.ipart_no and f.prod_cd = a.prod_cd)" & vbCrLf
    sql$ = sql$ & " and a.rev_no = '" & I_REV_NO & "' and a.ipart_no = '" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf

    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    If SSDATA.RecordCount > 0 Then
        varOEMMold_no = SSDATA!SMP_MOLD_NO
    End If
    SSDATA.Close
    
 'Assembly P-Bin Management and system will no write any main part information in Prd_work_31. Thus we need to get main part in Prd_BOM_1
    'FRC Jim Chen 2020/02/06
    If GLO_COMP_CD = "FU" Then
        If UCase(varMoldNo) = "ASSEMBLY" Then     'FRCUNION은 Add varOEMMold_no by jim
             sql$ = " select distinct a.prod_cd, a.mold_no, a.mold_desc " & vbCrLf
             sql$ = sql$ & " from PRD_MAST_1_VIEW a" & vbCrLf
             sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
             sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
             sql$ = sql$ & " and a.mold_no = (SELECT DISTINCT MASTER_MOLD_NO from PRD_WORK_1 where LOT_NO  ='" & Trim(C_LOT_NO) & "')" & vbCrLf
         
             Set SSDATA = Nothing
             SSDATA.Open sql$, DB_User
             
             If SSDATA.RecordCount = 0 Then
                 varOEMMold_no = ""
             Else
                 varOEMMold_no = str(Int(Right(Trim(SSDATA!mold_desc), 2)))
             End If
             SSDATA.Close
        ElseIf UCase(varMoldNo) = "" Then
            sql$ = " select distinct a.prod_cd, a.mold_no, a.mold_desc " & vbCrLf
            sql$ = sql$ & " from PRD_MAST_1_VIEW a" & vbCrLf
            sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
            sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
            sql$ = sql$ & " and a.prod_cd = (select child_cd from PRD_BOM_1 where parent_cd =(select distinct P_PROD_CD from prd_work_31 where LOT_NO ='" & Trim(C_LOT_NO) & "')" & vbCrLf
            sql$ = sql$ & "and child_cd like 'TR%')" & vbCrLf
    
             Set SSDATA = Nothing
             SSDATA.Open sql$, DB_User
             
             If SSDATA.RecordCount = 0 Then
                 varOEMMold_no = ""
             Else
                 varOEMMold_no = str(Int(Right(Trim(SSDATA!mold_desc), 2)))
             End If
             SSDATA.Close
        Else
             varOEMMold_no = str(Int(Right(Trim(varMoldNo), 2)))
        End If
    End If
        
    If (Trim(Right(C_CUSTOM_CD, 10)) = "20015" Or Trim(Right(C_CUSTOM_CD, 10)) = "20163") And isAgentPeak = True Then 'AgentPeak(기존S+G)의 경우 ASE KHH 금형 번호 고정
        varOEMMold_no = "S1"
    End If
    
    'FRC Part# FT189010
    If Trim( I_PART_NO.Text) = "FT189010" And CHK_INSMP = 1 Then
        varOEMMold_no = "S1"
        varLot_no = Replace(varLot_no, "F", "D")
    End If
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If

    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
    
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    
    ' 가로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0010" + "l" + "0980" + "0004" + vbCr
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0010" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0010" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0010" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0010" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0010" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0800" + "0010" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0880" + "0265" + "l" + "0980" + "0004" + vbCr

    '세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0010" + "0735" + "l" + "0004" + "0940" + vbCr
    'MOLD 세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0320" + "l" + "0004" + "0120" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0190" + "l" + "0004" + "0120" + vbCr
    'RT 세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0510" + "l" + "0004" + "0120" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0440" + "l" + "0004" + "0120" + vbCr
    'C/NO 세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0400" + "l" + "0004" + "0130" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0265" + "l" + "0004" + "0130" + vbCr
            
    'GRCode 세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0800" + "0265" + "l" + "0004" + "0150" + vbCr
                        
   'SUPPLIER
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "1000"
    sFormatTx = sFormatTx + "SUPPLIER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0710"
    sFormatTx = sFormatTx + "PEAK" + vbCr
        
   'MAT'L NAME  (고정문구)
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "1000"
    sFormatTx = sFormatTx + "MAT'L NAME" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0710"
    sFormatTx = sFormatTx + "TRAY" + vbCr
    
   'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0260" + "1000"
    sFormatTx = sFormatTx + "P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0265" + "0710"   'Code 128 SubSET B
       sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0305" + "0710"
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr

    'LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "1000"
    sFormatTx = sFormatTx + "LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0710"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & IIf(Option2.Value = True Or Option3.Value = True, Peak_varLot_no, varLot_no) & vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0425" + "0710"
    sFormatTx = sFormatTx & IIf(Option2.Value = True Or Option3.Value = True, Peak_varLot_no, varLot_no) & vbCr
    
    'Mold No
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0317"
    sFormatTx = sFormatTx + "MOLD" + vbCr
    If GLO_COMP_CD = "FU" Then 'FRCUNION은 MOLD NO 바코드를 제외한다. remove Mold Barcode by jim
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0425" + "0170"
        sFormatTx = sFormatTx & IIf(CHK_INSMP.Value = 1, Right(Trim(varOEMMold_no), 2), Right(Trim(varMoldNo), 2)) & vbCr
    Else
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0170"   'Code 128 SubSET B
        sFormatTx = sFormatTx & "B" & Right(Trim(varOEMMold_no), 2) & vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0425" + "0170"
        sFormatTx = sFormatTx & Right(Trim(varOEMMold_no), 2) & vbCr
    End If
    
    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "1000"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0710"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0710"
    sFormatTx = sFormatTx + varUnitQty + vbCr
    
    'RT
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "0500"
    sFormatTx = sFormatTx + "RT" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0420"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & Trim(C_LOT_ID) + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0420"
    sFormatTx = sFormatTx + Trim(C_LOT_ID) + vbCr
    
    'EXP DATE
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0620" + "1000"
    sFormatTx = sFormatTx + "EXP DATE" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0710"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & varEXPDT + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0665" + "0710"
    sFormatTx = sFormatTx + varEXPDT + vbCr
    
    'DOM    (YYYYMMdd) 'LOT 생성일자
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0740" + "1000"
    sFormatTx = sFormatTx + "DOM" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0745" + "0710"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & IIf(Option2.Value = True, varPeakMFGDT, varMFGDT) + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0785" + "0710"
    sFormatTx = sFormatTx + IIf(Option2.Value = True, varPeakMFGDT, varMFGDT) + vbCr
    
    'C/T
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0740" + "0390"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0745" + "0245"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & varCartonNo & vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0785" + "0245"
    sFormatTx = sFormatTx & varCartonNo & vbCr
    
    'CUSTOMER  (Use fixed)
    sText1 = IIf(Trim(Right(C_CUSTOM_CD, 10)) = 20015, "ASEK", Trim(Left(C_CUSTOM_CD, 30)))  'ASE KHH (Use fixed)
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0860" + "1000"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0860" + "0710"
    sFormatTx = sFormatTx & sText1 & vbCr   'trim(Left(C_CUSTOM_CD, 30))
    
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0935" + "1000"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0935" + "0710"
    sFormatTx = sFormatTx & varOEMiPART_NO & vbCr ' I_PART_NO & vbCr
    
    '2D QR Code
    '2D BARCODE(QR CODE)   'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로
    'QR코드는 숫자 최대 7089자, 문자(ASCII) 최대 4296자, 이진8비트 최대 2953바이트, 한자등 아시아 문자 최대 1817자를 담을 수 있다.
    sFormatTx = sFormatTx & "3W1D33000095002302HM, " & C_ITEM_NO & ";" & Trim(C_LOT_ID) & ";" & IIf(Option2.Value = True Or Option3.Value = True, Peak_varLot_no, varLot_no) & ";" & varUnitQty & ";" & varEXPDT & ";" & IIf(Option2.Value = True, varPeakMFGDT, varMFGDT) & Chr$(13)
    
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_ASECL_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2013. 05. 07
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' SCK용 2D BarCode (QR CODE)
    ' 업  체 : SCK
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMMold_no           As String
    Dim varOEMiRev_NO           As String
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, Trim(varMFGDT), Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.SMP_MOLD_NO,' ') as SMP_MOLD_NO from PRD_SMP_1 a " & vbCrLf
    sql$ = sql$ & " where a.mold_no in (select distinct mold_no from PRD_MAST_1 f " & vbCrLf
    sql$ = sql$ & "                     where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.mold_desc = '" & Trim(varMoldNo) & "' and f.Prev_No = a.rev_no and f.ipart_no = a.ipart_no and f.prod_cd = a.prod_cd)" & vbCrLf
    sql$ = sql$ & " and a.rev_no = '" & I_REV_NO & "' and a.ipart_no = '" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    If SSDATA.RecordCount > 0 Then
        varOEMMold_no = SSDATA!SMP_MOLD_NO
    End If
    SSDATA.Close
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close

'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    If GLO_COMP_CD = "FU" Then  ''ASECL (20276) 2020-02-17 by FRC IT Jim Chen & CPL Liam Bae
       '2D BARCODE(Data Matrix)
        'a W b[b] c d eee ffff gggg hhh i jjj kkk ll…l1 | W1 d 4 4 000 0010 0010This is the data portion
     
        ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
        '3W1C66000 0400 0980 0029 2000044044
        sFormatTx = sFormatTx & "3W1c66000" + "0400" + "0980" + "2000" + "044" + "044" & C_LOT_ID & "|" & C_ITEM_NO & "|" & C_LOT_NO & "|" & varUnitQty & "|" & Replace(Trim(varMFGDT), "-", "/") & "|" & varCartonNo & vbCr
            
       'RT NO
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "0980"
        sFormatTx = sFormatTx + "Batch NO:" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "0500"
        sFormatTx = sFormatTx + Trim(C_LOT_ID) + vbCr
        
       'CUST'R P/N
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0580" + "0980"
        sFormatTx = sFormatTx + "Part NO:" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0580" + "0500"
        sFormatTx = sFormatTx + C_ITEM_NO + vbCr
    
        'LOT NO
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0980"
        sFormatTx = sFormatTx + "Vender Lot:" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0500"
        sFormatTx = sFormatTx & C_LOT_NO & vbCr
        
        'QTY
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0740" + "0980"
        sFormatTx = sFormatTx + "QTY:" + vbCr
        sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0740" + "0500"
        sFormatTx = sFormatTx & varUnitQty + vbCr
        
        'MFG DATE
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0820" + "0980"
        sFormatTx = sFormatTx + "Manufacture Date:" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0820" + "0500"
        sFormatTx = sFormatTx + Replace(Trim(varMFGDT), "-", "/") + vbCr
               
        'C/T (Box No)
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0900" + "0980"
        sFormatTx = sFormatTx + "Box NO:" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0900" + "0500"
        sFormatTx = sFormatTx & varCartonNo & vbCr
    Else
        ' 가로줄
        sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0004" + vbCr
        sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0030" + "l" + "0980" + "0004" + vbCr
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr
        sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0030" + "l" + "0980" + "0004" + vbCr
        sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr
        sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0030" + "l" + "0980" + "0004" + vbCr
        '세로줄
        sFormatTx = sFormatTx + "1X11" + "000" + "0010" + "0715" + "l" + "0004" + "0940" + vbCr
        'QTY 세로줄
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0380" + "l" + "0004" + "0120" + vbCr
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0245" + "l" + "0004" + "0120" + vbCr
                   
       'SUPPLIER
        sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0980"
        sFormatTx = sFormatTx + "SUPPLIER" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0690"
        sFormatTx = sFormatTx + "PEAK" + vbCr
            
       'MAT'L NAME  (고정문구)
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0980"
        sFormatTx = sFormatTx + "MAT'L NAME" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0690"
        sFormatTx = sFormatTx + "TRAY" + vbCr
        
       'CUST'R P/N
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0260" + "0980"
        sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0265" + "0690"   'Code 128 SubSET B
           sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0305" + "0690"
        sFormatTx = sFormatTx + C_ITEM_NO + vbCr
    
        'LOT NO
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0980"
        sFormatTx = sFormatTx + "LOT NO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0690"   'Code 128 SubSET B
        sFormatTx = sFormatTx & "B" & Format(SHIP_DATE, "mmddyy") & vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0425" + "0690"
        sFormatTx = sFormatTx & Format(SHIP_DATE, "mmddyy") & vbCr
        
        'QTY
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0370"
        sFormatTx = sFormatTx + "QTY" + vbCr
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0225"   'Code 128 SubSET B
        sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0425" + "0225"
        sFormatTx = sFormatTx & varUnitQty + vbCr
        
        'MFG DATE
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "0980"
        sFormatTx = sFormatTx + "MFG DATE" + vbCr
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0690"   'Code 128 SubSET B
        sFormatTx = sFormatTx & "B" & Format(SHIP_DATE, "yyyy/mm/dd") + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0690"
        sFormatTx = sFormatTx + Format(SHIP_DATE, "yyyy/mm/dd") + vbCr
           
        'PEAK P/N
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0620" + "0980"
        sFormatTx = sFormatTx + "PEAK P/N" + vbCr
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0690"   'Code 128 SubSET B
        sFormatTx = sFormatTx & "B" & varOEMiPART_NO + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0665" + "0690"
        sFormatTx = sFormatTx + varOEMiPART_NO + vbCr
                 
        'CUSTOMER  (Use fixed)
        sText1 = IIf(Trim(Right(C_CUSTOM_CD, 10)) = 20779, "ASE-CL", Trim(Left(C_CUSTOM_CD, 30))) 'ASE CL (Use fixed)
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0860" + "0980"
        sFormatTx = sFormatTx + "CUSTOMER" + vbCr
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0800" + "0690"   'Code 128 SubSET B
        sFormatTx = sFormatTx & "B" & sText1 & vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0860" + "0690"
        sFormatTx = sFormatTx & sText1 & vbCr
    End If
    
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_MICRONTW_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2020. 02. 06
    ' DEVELOPER : FRC JIM CHEN
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim charRowAdd(5 To 48)     As Integer
    Dim varDWG_no               As String
    Dim varPrdITEM_no           As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    
    Dim varMoldStamp            As String
    Dim varOEMMold_no           As String

    varLot_no = C_LOT_NO
        
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot ???濠? 陛螳褥
    varEXPDT = Replace(DateAdd("yyyy", 5, varMFGDT), "-", "")
    varMFGDT = Replace(varMFGDT, "-", "/")

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    '===========================================================================
    ' SMP 薑爾 ?檣, Drawing No ?檣
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " (select distinct nvl(item_no, ' ') item_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as prd_item_no, " & vbCrLf
    sql$ = sql$ & " (select distinct nvl(dwg_no, ' ') dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as dwg_no" & vbCrLf
    sql$ = sql$ & " from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    varPrdITEM_no = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
        varDWG_no = SSDATA!dwg_no
        varPrdITEM_no = SSDATA!prd_item_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
    
    ' MOLD STAMP
    '===========================================================================
    sql$ = " select distinct a.prod_cd, b.prod_gubun, a.mold_no, a.mold_desc, " & vbCrLf
    sql$ = sql$ & " nvl((select mold_stamp from prd_mold_1 f where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.mold_no =  b.mold_no), ' ') mold_stamp " & vbCrLf
    sql$ = sql$ & " from PRD_MAST_1_VIEW a, PRD_WORK_11 b " & vbCrLf
    sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    sql$ = sql$ & " and a.prod_cd = b.prod_cd "
    sql$ = sql$ & " and a.comp_cd = b.comp_cd and a.bsns_cd = b.bsns_cd " & vbCrLf
    If varMoldNo = "ASSEMBLY" Then   'Assembly 薯?曖 ?辦 Main 薯?曖 LOT煎 綠掖??.
        sql$ = sql$ & " and   b.outlot_no in (select distinct outlot_no from prd_work_31 where lot_no= '" & Replace(Trim(C_LOT_NO), "-", "") & "' and main_kb ='1' ) " & vbCrLf
    Else
        sql$ = sql$ & " and   b.outlot_no = '" & Replace(Trim(C_LOT_NO), "-", "") & "' " & vbCrLf
    End If
    
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
   
    If SSDATA.RecordCount = 0 Then
        varMoldStamp = ""
    Else
        varMoldStamp = Trim(SSDATA!mold_stamp)
    End If
    SSDATA.Close
    
    Set SSDATA = Nothing

    'Assembly P-Bin Management and system will no write any main part information in Prd_work_31. Thus we need to get main part in Prd_BOM_1
    'FRC Jim Chen 2020/02/06
   If varMoldNo = "ASSEMBLY" Then
        sql$ = " select distinct a.prod_cd, a.mold_no, a.mold_desc " & vbCrLf
        sql$ = sql$ & " from PRD_MAST_1_VIEW a" & vbCrLf
        sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
        sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        sql$ = sql$ & " and a.prod_cd = (select child_cd from PRD_BOM_1 where parent_cd =(select distinct P_PROD_CD from prd_work_31 where LOT_NO ='" & Trim(C_LOT_NO) & "')" & vbCrLf
        sql$ = sql$ & "and child_cd like 'TR%')" & vbCrLf
    
        Set SSDATA = Nothing
        SSDATA.Open sql$, DB_User
        
        If SSDATA.RecordCount = 0 Then
            varOEMMold_no = ""
        Else
            varOEMMold_no = str(Int(Right(Trim(SSDATA!mold_desc), 2)))
        End If
        SSDATA.Close
   End If

'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 ?陽 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    Dim varRate     As String
    varRate = "52"
    
    ' Barcode 還 ?晦    -- reserved --   -Row-    -Col-   -line-  -望檜 + 舒眷-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    ' Barcode ?衛           -轎?陝- -type-  --舒眷--   -堪檜-   -row-    -col-          * 舒眷 : 掃擎 舒眷 + 整擎 舒眷
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'CUST'R
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0160" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0160" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
    
   'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0220" + "0980"
    sFormatTx = sFormatTx + "CUST'R PO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0220" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + vbCr
    
    'MICRON SGP and MICRON US add customer PO code 128 barcode
    If Trim(Right(C_CUSTOM_CD, 10)) = "20910" Then 'Or Trim(Right(C_CUSTOM_CD, 10)) = "20879" Then
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0220" + "0400"   'Code 128 SubSET B
        sFormatTx = sFormatTx & "B" & C_PO_NO + vbCr
    Else
    
    End If
    
    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0280" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0280" + "0730"
    sFormatTx = sFormatTx + ":" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0280" + "0700"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0340" + "0700"
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr
    
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0400" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0400" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '  I_PART_NO + vbCr
    
    'PEAK DWG#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0460" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0460" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr

    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0980"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0520" + "0450"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
               
    'SHIP'G LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0580" + "0980"
    sFormatTx = sFormatTx + "BATCH NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0580" + "0730"
    sFormatTx = sFormatTx + ": " + varLot_no + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0640" + "0710"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & varLot_no + vbCr
    
    'EXP DATE
    sText1 = "12/31/9999"   'Fixed Date
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0980"
    sFormatTx = sFormatTx + "EXP DATE" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0720" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0720" + "0450"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" + sText1 + vbCr
        
    'Mold NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0780" + "0980"
    sFormatTx = sFormatTx + "MOLD NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0780" + "0730"
    sFormatTx = sFormatTx + ": " + IIf(Trim(varMoldStamp) = "", Trim(varOEMMold_no), Trim(varMoldStamp)) + vbCr
                
    'COO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0840" + "0980"
    sFormatTx = sFormatTx + "COO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0840" + "0730"
    sFormatTx = sFormatTx + ": " + "MADE IN TAIWAN" + vbCr
    
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_POWERTECH_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2020. 05. 14
    ' DEVELOPER : FRC JIM CHEN
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim charRowAdd(5 To 48)     As Integer
    Dim varDWG_no               As String
    Dim varPrdITEM_no           As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    
    Dim varMoldStamp            As String
    Dim varOEMMold_no           As String

    varLot_no = C_LOT_NO
        
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot ???濠? 陛螳褥
    varEXPDT = Replace(DateAdd("yyyy", 5, varMFGDT), "-", "")
    varMFGDT = Replace(varMFGDT, "-", "/")

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    '===========================================================================
    ' SMP 薑爾 ?檣, Drawing No ?檣
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " (select distinct nvl(item_no, ' ') item_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as prd_item_no, " & vbCrLf
    sql$ = sql$ & " (select distinct nvl(dwg_no, ' ') dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as dwg_no" & vbCrLf
    sql$ = sql$ & " from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    varPrdITEM_no = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
        varDWG_no = SSDATA!dwg_no
        varPrdITEM_no = SSDATA!prd_item_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
    
    ' MOLD STAMP
    '===========================================================================
    sql$ = " select distinct a.prod_cd, b.prod_gubun, a.mold_no, a.mold_desc, " & vbCrLf
    sql$ = sql$ & " nvl((select mold_stamp from prd_mold_1 f where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.mold_no =  b.mold_no), ' ') mold_stamp " & vbCrLf
    sql$ = sql$ & " from PRD_MAST_1_VIEW a, PRD_WORK_11 b " & vbCrLf
    sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    sql$ = sql$ & " and a.prod_cd = b.prod_cd "
    sql$ = sql$ & " and a.comp_cd = b.comp_cd and a.bsns_cd = b.bsns_cd " & vbCrLf
    If varMoldNo = "ASSEMBLY" Then   'Assembly 薯?曖 ?辦 Main 薯?曖 LOT煎 綠掖??.
        sql$ = sql$ & " and   b.outlot_no in (select distinct outlot_no from prd_work_31 where lot_no= '" & Replace(Trim(C_LOT_NO), "-", "") & "' and main_kb ='1' ) " & vbCrLf
    Else
        sql$ = sql$ & " and   b.outlot_no = '" & Replace(Trim(C_LOT_NO), "-", "") & "' " & vbCrLf
    End If
    
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
   
    If SSDATA.RecordCount = 0 Then
        varMoldStamp = ""
    Else
        varMoldStamp = Trim(SSDATA!mold_stamp)
    End If
    SSDATA.Close
    
    Set SSDATA = Nothing

    'Assembly P-Bin Management and system will no write any main part information in Prd_work_31. Thus we need to get main part in Prd_BOM_1
    'FRC Jim Chen 2020/02/06
   If varMoldNo = "ASSEMBLY" Then
        sql$ = " select distinct a.prod_cd, a.mold_no, a.mold_desc " & vbCrLf
        sql$ = sql$ & " from PRD_MAST_1_VIEW a" & vbCrLf
        sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
        sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        sql$ = sql$ & " and a.prod_cd = (select child_cd from PRD_BOM_1 where parent_cd =(select distinct P_PROD_CD from prd_work_31 where LOT_NO ='" & Trim(C_LOT_NO) & "')" & vbCrLf
        sql$ = sql$ & "and child_cd like 'TR%')" & vbCrLf
    
        Set SSDATA = Nothing
        SSDATA.Open sql$, DB_User
        
        If SSDATA.RecordCount = 0 Then
            varOEMMold_no = ""
        Else
            varOEMMold_no = str(Int(Right(Trim(SSDATA!mold_desc), 2)))
        End If
        SSDATA.Close
   End If

'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 ?陽 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    Dim varRate     As String
    varRate = "52"
    
    ' Barcode 還 ?晦    -- reserved --   -Row-    -Col-   -line-  -望檜 + 舒眷-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    ' Barcode ?衛           -轎?陝- -type-  --舒眷--   -堪檜-   -row-    -col-          * 舒眷 : 掃擎 舒眷 + 整擎 舒眷
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'CUST'R
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0160" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0160" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
    
   'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0220" + "0980"
    sFormatTx = sFormatTx + "CUST'R PO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0220" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + vbCr

    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0280" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0280" + "0730"
    sFormatTx = sFormatTx + ":" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0280" + "0700"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0340" + "0700"
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr
    
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0400" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0400" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '  I_PART_NO + vbCr
    
    'PEAK DWG#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0460" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0460" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr

    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0980"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0520" + "0450"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
               
    'SHIP'G LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0580" + "0980"
    sFormatTx = sFormatTx + "BATCH NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0580" + "0730"
    sFormatTx = sFormatTx + ": " + varLot_no + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0640" + "0710"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & varLot_no + vbCr
    
    'EXP DATE
    sText1 = "12/31/9999"   'Fixed Date
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0980"
    sFormatTx = sFormatTx + "EXP DATE" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0720" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0720" + "0450"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" + sText1 + vbCr
        
    'Mold NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0780" + "0980"
    sFormatTx = sFormatTx + "MOLD NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0780" + "0730"
    sFormatTx = sFormatTx + ": " + IIf(Trim(varMoldStamp) = "", Trim(varOEMMold_no), Trim(varMoldStamp)) + vbCr
                
    'COO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0840" + "0980"
    sFormatTx = sFormatTx + "COO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0840" + "0730"
    sFormatTx = sFormatTx + ": " + "MADE IN TAIWAN" + vbCr
    
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_TITL_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2017. 03. 07
    ' TITL용
    ' 업  체 : TITL
    ' 개발자 : sjsung
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If

    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
    
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add

    Dim varRate     As String
    varRate = "52"

    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'CUST'R
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUST'R" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
        
   'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0240" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr

    'PEAK DWG#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0310" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0310" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
        
    'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "CUST'R PO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + vbCr '  I_PART_NO + vbCr
    
    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0450" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0450" + "0730"
    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0520" + "0980"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & C_ITEM_NO & vbCr
    
    'SHIP'G LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
    sFormatTx = sFormatTx + "SHIP'G LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + "-" + Format(SHIP_DATE, "mmddyy") + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0660" + "0980"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & C_PO_NO + "-" + Format(SHIP_DATE, "mmddyy") & vbCr
    
    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0980"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0720" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
    '여백
            
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_TITL_InnerLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2017-03-08
    ' 업  체 : SCS InnerLabel
    ' 개발자 : sjsung
    '======================================================================
    Dim sql                As String
    Dim sText               As String
    Dim varDescNm       As String
    Dim varDWG_no       As String
    Dim varLotNo        As String
    Dim varMFGDT        As String
    Dim varOEMiPART_NO  As String
    Dim varOEMiRev_NO   As String
    Dim sTemp_nm        As String
    Dim sMsg                As String
    Dim varUnitQty      As String
        
    '-------------------------------------------------------------
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    If varUnitQty Mod Val(I_BUNDLE_QT) <> 0 Or Trim(I_BUNDLE_QT) = 0 Or Trim(I_BUNDLE_QT) = "" Then
        MsgBox "Bundle 수량을 재확인하세요.", vbExclamation
        I_BUNDLE_QT.SetFocus
        Exit Sub
    Else
        varUnitQty = IIf(PRT_FORM(1).Value = True, varUnitQty \ Val(I_BUNDLE_QT), varUnitQty)   'I_BOX_UNIT_QT를 Bundle 수량만큼 나눈다. (정수로 표시)
    End If
    
    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
    '-------------------------------------------------------------
    
    '============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"

    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr

    '                      -출력각     --- Font ---    -Row-    -Col-
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0090" + "0980"
    sFormatTx = sFormatTx + "TI Part number : " & IIf(C_ITEM_NO = "", " ", C_ITEM_NO) & vbCr   'sText1
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0160" + "0980"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & IIf(C_ITEM_NO = "", " ", C_ITEM_NO) & vbCr
        
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0980"
    sFormatTx = sFormatTx + "Supplier Lot No : LN" & C_LOT_NO & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0310" + "0980"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "BLN" & C_LOT_NO & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0390" + "0980"
    sFormatTx = sFormatTx + "Expiration date : " & DateAdd("yyyy", 3, Format(SHIP_DATE, "YYYY-MM-DD")) & vbCr '출하일로 부터 3년
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0460" + "0980"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & DateAdd("yyyy", 3, Format(SHIP_DATE, "YYYY-MM-DD")) & vbCr
                
    '분리선 ----------------------------------------------------------------------- (위/아래 동일한 내용 출력)
    sFormatTx = sFormatTx + "1X11" + "000" + "0480" + "0030" + "l" + "0980" + "0002" + vbCr
    
     '                      -출력각     --- Font ---    -Row-    -Col-
    
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0980"
    sFormatTx = sFormatTx + "TI Part number : " & IIf(C_ITEM_NO = "", " ", C_ITEM_NO) & vbCr   'sText1
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0660" + "0980"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & IIf(C_ITEM_NO = "", " ", C_ITEM_NO) & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0740" + "0980"
    sFormatTx = sFormatTx + "Supplier Lot No : LN" & C_LOT_NO & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0810" + "0980"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "BLN" & C_LOT_NO & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0890" + "0980"
    sFormatTx = sFormatTx + "Expiration date : " & DateAdd("yyyy", 3, Format(SHIP_DATE, "YYYY-MM-DD")) & vbCr '출하일로 부터 3년
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0960" + "0980"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & DateAdd("yyyy", 3, Format(SHIP_DATE, "YYYY-MM-DD")) & vbCr
        
    sFormatTx = sFormatTx + "E" + vbCr
End Sub
Private Sub T4_STS_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2011. 12. 21
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' SCK용 2D BarCode (QR CODE)
    ' 업  체 : SCK
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varDWG_no               As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, varMFGDT, Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
        
    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as dwg_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    If SSDATA.RecordCount > 0 Then
        varDWG_no = SSDATA!dwg_no
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
 
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
 
 
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' 외부 Box 그리기           -- reserved --   -Row-    -Col-   -box- -수평길이+수직길이- -- line 폭 --
    ' sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0030" + "b" + "0980" + "0950" + "0004" + "0004" + vbCr
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    
     ' 가로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0004" + vbCr
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0800" + "0030" + "l" + "0980" + "0004" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0880" + "0030" + "l" + "0980" + "0004" + vbCr

    '세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0010" + "0760" + "l" + "0004" + "0920" + vbCr
        
    'QTY 세로줄
    'sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0380" + "l" + "0004" + "0120" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0245" + "l" + "0004" + "0120" + vbCr
               
   'SUPPLIER
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0980"
    sFormatTx = sFormatTx + "Supplier" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0740"
    sFormatTx = sFormatTx + "PEAK" + vbCr
        
   'Customer
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0980"
    sFormatTx = sFormatTx + "Customer" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0740"
    sFormatTx = sFormatTx + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
   
   'DESC
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0260" + "0980"
    sFormatTx = sFormatTx + "DESC" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0260" + "0740"
    sFormatTx = sFormatTx + sText1 + vbCr

    'LOT NO
    sText1 = C_PO_NO & "-" & Format(SHIP_DATE, "mmddyy") & "-" & Format(CStr(varCartonNo), "00#")
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0980"
    sFormatTx = sFormatTx + "LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0740"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & sText1 & vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0425" + "0740"
    sFormatTx = sFormatTx & sText1 & vbCr
    
    'ITEM CODE
    sText1 = C_ITEM_NO
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "0980"
    sFormatTx = sFormatTx + "Item Code" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0740"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0740"
    sFormatTx = sFormatTx & sText1 + vbCr
    
    'QTY
    sText1 = varUnitQty
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0620" + "0980"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0740"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0665" + "0740"
    sFormatTx = sFormatTx + sText1 + "PCS" + vbCr
    
    'SHIPDATE
    sText1 = Trim(SHIP_DATE)
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0740" + "0980"
    sFormatTx = sFormatTx + "Ship Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0745" + "0740"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0785" + "0740"
    sFormatTx = sFormatTx + sText1 + vbCr
       
    'MFG_DATE
    sText1 = varMFGDT
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0870" + "0980"
    sFormatTx = sFormatTx + "MFG Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0875" + "0740"   'Code 128 SubSET B
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0915" + "0740"
    sFormatTx = sFormatTx + sText1 + vbCr
   
        
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_SPIL_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2013. 05. 07
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' SCK용 2D BarCode (QR CODE)
    ' 업  체 : SPIL TW
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varMoldSeq              As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varDWG_no               As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim strLOT_ID               As String
    
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    varMoldSeq = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2)
    
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@@@@@") 'Lot 생성일자를 가져옴
    'varMFGDT = Left(varMFGDT, 6) & "01" '20170918 생산월 01로 고정 표시 ihson , CPL hilee 2020.11.24 실제 생산일자를 표시함 (모든 SPIL SITE 공통적용)
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
      ' 해당월의 출하 횟수를 확인
    ' Display Shipping Date for the C_LOT_NO for Delivery Order for the Lot No
    
    sql$ = "select ipart_no , Prev_No,   'P' || '" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "' || trim(to_char(ROWNUM, '00')) as LOT_ID"
    sql$ = sql$ & " from ( select distinct ipart_no, Prev_No from MKT_OUT_11 a where a.comp_cd = 'SG' and a.bsns_cd = '11' and a.outorder_no ='" & Trim(C_LOT_ID) & "'"
    sql$ = sql$ & ")"
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
            
    strLOT_ID = ""
    If SSDATA.RecordCount > 0 Then
        Do Until SSDATA.EOF
            If Trim(SSDATA!ipart_no) = Trim( I_PART_NO) And Trim(SSDATA!Prev_No) = Trim(I_REV_NO) Then strLOT_ID = Trim(SSDATA!LOT_ID)
            SSDATA.MoveNext
        Loop
    End If
    SSDATA.Close
    If Trim(strLOT_ID) = "" Then strLOT_ID = "P" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "01"

    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no,  nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as dwg_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    If SSDATA.RecordCount > 0 Then
        varDWG_no = SSDATA!dwg_no
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
 
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
 
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' 외부 Box 그리기           -- reserved --   -Row-    -Col-   -box- -수평길이+수직길이- -- line 폭 --
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0030" + "b" + "0980" + "0950" + "0004" + "0004" + vbCr

    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    
    '2D QR Code Request by FRC CPL Liam 2020-06-05
    If GLO_COMP_CD = "FU" Then
        sFormatTx = sFormatTx & "3W1D66000042002902HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)
    End If

    '2D BARCODE(QR CODE)    'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로, b,c = Module Size
    'sFormatTx = sFormatTx & "3W1D33000030502302HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)  'FRC '2D QR Code Request by FRC CPL Liam 2020-06-05
    'sFormatTx = sFormatTx & "3W1D55000063002202HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)

    ' 가로줄 Horizontal Line
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0004" + vbCr
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0030" + "l" + "0980" + "0004" + vbCr
    If GLO_COMP_CD = "FU" Then
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0330" + "l" + "0980" + "0004" + vbCr
    Else
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr
    End If
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr
    'horizontal line under qty
    'sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0245" + "l" + "0765" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0740" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0810" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0880" + "0030" + "l" + "0980" + "0004" + vbCr
    
    '세로줄
    
    '2nd vertial line
    'sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0245" + "l" + "0004" + "0380" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0760" + "l" + "0004" + "0850" + vbCr
    
    '3rd vertial line
    If GLO_COMP_CD = "FU" Then
       'If Trim(Right(C_CUSTOM_CD, 10)) = "21311" Then
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0365" + vbCr
        'Else
            sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0240" + vbCr
        'End If
    End If
    '4th vertial line
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0245" + "l" + "0004" + "0140" + vbCr
    

        
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0980"
    sFormatTx = sFormatTx + "" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0770"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" + vbCr
        
   'CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0740"
    sFormatTx = sFormatTx + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
        
    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0260" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0265" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0305" + "0740"
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr

    'LOT NO (P+YMMDDXX) : XX = 출하횟수 -> 20170918 실제 LOT 표시 ihson
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & Replace(varLot_no, "-", "") & vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0425" + "0740"
    sFormatTx = sFormatTx + Replace(varLot_no, "-", "") + vbCr
            
    'MFG DATE
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0980"
    sFormatTx = sFormatTx + "MFG DATE" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & varMFGDT + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0740"
    sFormatTx = sFormatTx + varMFGDT + vbCr

    'Adding COO on SPIL SUZ Shipping label reqeusted by CPL Liam 2020.0704
    'If Trim(Right(C_CUSTOM_CD, 10)) = "21311" And GLO_COMP_CD = "FU" Then
        'QTY
        'sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "0320"
        'sFormatTx = sFormatTx + "QTY" + vbCr
        'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0210"  'Code 128
        'sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0200"
        'sFormatTx = sFormatTx + varUnitQty + vbCr
        
        'PEAK DWG
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0620" + "0980"
        'sFormatTx = sFormatTx + "PEAK DWG" + vbCr
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0620" + "0740"
        'sFormatTx = sFormatTx & I_DESC_NM & vbCr

        
        'CUST'R PO
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0980"
        'sFormatTx = sFormatTx + "CUST'R PO" + vbCr
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0740"
        'sFormatTx = sFormatTx & C_PO_NO & vbCr
        
        'Rev
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0230"
        sFormatTx = sFormatTx & "REV# : " & Trim(varOEMiRev_NO) & vbCr
        
    
        'PEAK P/N
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0790" + "0980"
        'sFormatTx = sFormatTx + "PEAK P/N" + vbCr
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0790" + "0740"
        'sFormatTx = sFormatTx & varOEMiPART_NO & vbCr ' I_PART_NO & vbCr
        
        'Mold No
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0790" + "0230"
        'sFormatTx = sFormatTx & "MOLD# : " & IIf(GLO_COMP_CD = "FU", Right(varMoldSeq, 2), "S" & Right(varMoldSeq, 1)) & vbCr  'FRC는 S를 제거한다. 2019.11.12 by Liam
        
        'COO
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0980"
        'sFormatTx = sFormatTx + "COO" + vbCr
        'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0740"
        'sFormatTx = sFormatTx & "MADE IN TAIWAN" + vbCr
        

    'Else
        'QTY
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0620" + "0980"
        sFormatTx = sFormatTx + "QTY" + vbCr
        sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0740"  'Code 128
        sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0665" + "0740"
        sFormatTx = sFormatTx + varUnitQty + vbCr
        
                              
        'PEAK DWG
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0980"
        sFormatTx = sFormatTx + "PEAK DWG" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0740"
        sFormatTx = sFormatTx & I_DESC_NM & vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0230"
        sFormatTx = sFormatTx & "REV# : " & Trim(varOEMiRev_NO) & vbCr
        
        'CUST'R PO
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0790" + "0980"
        sFormatTx = sFormatTx + "CUST'R PO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0790" + "0740"
        sFormatTx = sFormatTx & C_PO_NO & vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0790" + "0230"
        'sFormatTx = sFormatTx & "MOLD# : " & IIf(GLO_COMP_CD = "FU", "", "S") & Right(varMoldSeq, 1) & vbCr 'FRC는 S를 제거한다. 2019.11.12 by Liam
        sFormatTx = sFormatTx & "MOLD# : " & IIf(GLO_COMP_CD = "FU", Right(varMoldSeq, 2), "S" & Right(varMoldSeq, 1)) & vbCr  'FRC는 S를 제거한다. 2019.11.12 by Liam
    
        'PEAK P/N
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0980"
        sFormatTx = sFormatTx + "PEAK P/N" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0740"
        sFormatTx = sFormatTx & varOEMiPART_NO & vbCr ' I_PART_NO & vbCr
        
    'End If
    
    'MsgBox (varOEMiPART_NO)
    
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0935" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0935" + "0740"
    sFormatTx = sFormatTx & varCartonNo & vbCr ' I_PART_NO & vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr

End Sub

Private Sub T4_SCC_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2015. 07. 29
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' 업  체 : SCC
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varMoldSeq              As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varDescNm               As String
    Dim varOEMiRev_NO           As String
    Dim varOEMiPART_NO          As String
    
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    varMoldSeq = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2)
    
    'varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@@@@@") 'Lot 생성일자를 가져옴
    varMFGDT = Format(SHIP_DATE, "yyyy-mm-dd")   '출하일을 제조일로 한다.
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, varMFGDT, Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    varMFGDT = Replace(varMFGDT, "-", "")
    varEXPDT = Replace(varEXPDT, "-", "")
    
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
        
    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct pdesc_nm from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as pdesc_nm " & vbCrLf
    sql$ = sql$ & " from MKT_CUST_2 a " & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDescNm = ""
    If SSDATA.RecordCount > 0 Then
        varDescNm = SSDATA!pdesc_nm
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
 
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If

'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' 외부 Box 그리기           -- reserved --   -Row-    -Col-   -box- -수평길이+수직길이- -- line 폭 --
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0030" + "b" + "0980" + "0950" + "0004" + "0004" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    ' 가로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0004" + vbCr
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0030" + "l" + "0980" + "0004" + vbCr
    
   ' sFormatTx = sFormatTx + "1X11" + "000" + "0740" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0790" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0870" + "0030" + "l" + "0980" + "0004" + vbCr

    '세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0350" + "l" + "0004" + "0120" + vbCr 'QTY1
    sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0260" + "l" + "0004" + "0120" + vbCr 'QTY2
    sFormatTx = sFormatTx + "1X11" + "000" + "0790" + "0245" + "l" + "0004" + "0080" + vbCr 'REV
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0760" + "l" + "0004" + "0850" + vbCr
        
   'Title / CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0870"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL  (" & Trim(Left(C_CUSTOM_CD, 30)) & ")" + vbCr            '
        
    'DESC
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "DESC" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0740"
    sFormatTx = sFormatTx + varDescNm + vbCr
        
    'CUST'R P/N
    sText1 = C_ITEM_NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0265" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0265" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0305" + "0740"
    sFormatTx = sFormatTx + sText1 + vbCr
    
    'DPL 원그리기 Rotation,Multiplier(1X11)--Fill Pattern(fff)--Row center point(rrrr)--Column center point(cccc)--Circle ID (C)--Fixed Value(001 0001)--Radius circle(rrrr)
                 '1 X 11 000 0100 0100 C 001 0001 0025<CR>
    sText1 = IIf(Trim(I_METERIAL_NM) = "PES", "GP/E", "GP")
    sFormatTx = sFormatTx + "1X11" + "000" + "0253" + "0140" + "C" + "001" + "0001" + "0052" + vbCr   'Circle 그리기 (원그리기)
    sFormatTx = sFormatTx + Angle180 + "911" + "S01" + "0275" + "0180" + IIf(sText1 = "GP", "P020P015", "P018P010") + sText1 + vbCr 'GP, GP/E 폰트크기조정
    'sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0250" + "0220" + sText1 + vbCr
    

    'LOTNO
    sText1 = Format(SHIP_DATE, "yyyymmdd")
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0385" + "0980"
    sFormatTx = sFormatTx + "LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & sText1 & vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0425" + "0740" + sText1 + vbCr

    'QTY
    sText1 = varUnitQty
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0385" + "0340"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0245" 'Code 128
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0425" + "0245" + sText1 + vbCr
    
    'CUST'R PO
    sText1 = C_PO_NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0505" + "0980"
    sFormatTx = sFormatTx + "CUST'R PO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0740" + sText1 + vbCr

    'MFG DATE       'Ship Date를 지정 by ihson 2019.03.22
    sText1 = Trim(Format(varMFGDT, "@@@@.@@.@@"))
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0625" + "0980"
    sFormatTx = sFormatTx + "MFG Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0665" + "0740" + sText1 + vbCr

    'Exp Date       'Expire Date =  Ship Date를 지정 +3년(MKT_CUST_2), by ihson 2019.03.22
    sText1 = Trim(Format(varEXPDT, "@@@@.@@.@@"))
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0745" + "0980"
    sFormatTx = sFormatTx + "EXP Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0745" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0785" + "0740" + sText1 + vbCr
                          
    'PEAK DWG
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0850" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0850" + "0740"
    sFormatTx = sFormatTx & I_DESC_NM & vbCr ' I_PART_NO & vbCr
    'REV
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0850" + "0230"
    sFormatTx = sFormatTx & "REV# : " & Trim(varOEMiRev_NO) & vbCr
    

    'COO
    'FRC removes COO for C/NO because of Chinese Customs Diplomatic issue
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0930" + "0980"
    'sFormatTx = sFormatTx + IIf(GLO_COMP_CD = "FU", "C/NO", "COO") + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0930" + "0740"

    'COO / FRC(C/NO)
    If GLO_COMP_CD = "FU" Then  'FRC removes COO for C/NO because of Chinese Customs Diplomatic issue
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0930" + "0980"
        sFormatTx = sFormatTx + "C/NO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0930" + "0740"
        sFormatTx = sFormatTx + ": " & varCartonNo & "/" & I_END_CT_NO & vbCr
    Else
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0930" + "0980"
        sFormatTx = sFormatTx + "COO" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0930" + "0740"
        sFormatTx = sFormatTx & "MADE IN KOREA" & vbCr ' I_PART_NO & vbCr
    End If

    
    'If GLO_COMP_CD = "FU" Then
        'sFormatTx = sFormatTx & "MADE IN TAIWAN" & vbCr ' I_PART_NO & vbCr
        'sFormatTx = sFormatTx + ": " & varCartonNo & "/" & I_END_CT_NO & vbCr
    'Else
        'sFormatTx = sFormatTx & "MADE IN KOREA" & vbCr ' I_PART_NO & vbCr
    'End If
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_SCC_InnerLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2015.07.29
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' 업  체 : SCC
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                As String
    Dim sText               As String
    Dim varDescNm       As String
    Dim varDWG_no       As String
    Dim varOEMiPART_NO  As String
    Dim varOEMiRev_NO   As String
    Dim sTemp_nm        As String
    Dim sMsg                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    
    'varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@@@@@") 'Lot 생성일자를 가져옴
    varMFGDT = Format(SHIP_DATE, "yyyy-mm-dd")   '출하일을 제조일로 한다.
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, varMFGDT, Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    varMFGDT = Replace(varMFGDT, "-", "")
    varEXPDT = Replace(varEXPDT, "-", "")
    '-------------------------------------------------------------
    ' Drawing No/ SMP Part No 확인 / Description 정보를 가져옴 (Package & Size & Thinkness)
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as dwg_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct trim((select mcode_nm from SYS_CODPD_03 where comp_cd = f.comp_cd and bsns_cd = f.bsns_cd and mcode_cd = f.ptype_kb and mcode_class = f.prod_gubun ) ||' '|| psize_nm ||' '|| pthick_nm) " & vbCrLf
    sql$ = sql$ & "                                                                             from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as desc_nm " & vbCrLf
    sql$ = sql$ & " from MKT_CUST_2 a where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    varDescNm = ""
    If SSDATA.RecordCount > 0 Then
        varDWG_no = SSDATA!dwg_no
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
        varDescNm = Trim(SSDATA!desc_nm)
    End If
    SSDATA.Close
    '-------------------------------------------------------------
    ' 온도 확인 MKT_CUST_2_Desc_tx4 (고객별 추가 정보컬럼)
    sql$ = "select distinct a.cdesc_tx4 from MKT_CUST_2 a " & vbCrLf
    sql$ = sql$ + " where a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ + " and   a.custom_cd = '" & Trim(Right(C_CUSTOM_CD, 10)) & "' " & vbCrLf
    sql$ = sql$ + " and   a.item_no = '" + IIf(C_ITEM_NO = "", " ", C_ITEM_NO) + "' "
    sql$ = sql$ & " and a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    If SSDATA.RecordCount = 0 Then
        sMsg = "거래처별 ITEM 정보에 Temperature가 미등록 되어 있습니다. " & vbCrLf
        sMsg = sMsg & "영업담당자에게 아래 제품 정보를 재확인하시기 바랍니다." & vbCrLf
        sMsg = sMsg & "Part No: " &  I_PART_NO & ",  Rev: " & I_REV_NO & vbCrLf
        sMsg = sMsg & vbCrLf & "Temperature 없이 출력을 계속 진행하시겠습니까?"
        If MsgBox(sMsg, vbYesNo + vbQuestion, "SCC Label Print") = vbNo Then Exit Sub
        sTemp_nm = " "
    Else
        sTemp_nm = Trim(SSDATA!cdesc_tx4)
    End If
    SSDATA.Close
    '============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"

    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr

    '                      -출력각     --- Font ---    -Row-    -Col-
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0060" + "0980"
    sFormatTx = sFormatTx + "Vendor: PEAK(KOREA)" & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0100" + "0980"
    sFormatTx = sFormatTx + "Part No : " & I_DESC_NM & vbCr 'sText1
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0140" + "0980"
    sFormatTx = sFormatTx + "Package&Size&Thickness: " & varDescNm & "MM" & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0180" + "0980"
    sFormatTx = sFormatTx + "Temperature : " & sTemp_nm & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0220" + "0980"
    sFormatTx = sFormatTx + "Drawing No : " & I_DESC_NM & vbCr
        
    'Material Code
    sText = C_ITEM_NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0270" + "0980"
    sFormatTx = sFormatTx + "Material Code : " & sText & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "054" + "0325" + "0980"
    sFormatTx = sFormatTx & "B" & sText & vbCr
    
    'Lot No
    sText = Format(SHIP_DATE, "yyyymmdd")
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0270" + "0450"
    sFormatTx = sFormatTx + "Lot No : " & sText & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "054" + "0325" + "0450"
    sFormatTx = sFormatTx & "B" & sText & vbCr
                           
    'Manufacture Date =  Ship Date를 지정 2015.09.14 S+G 추가 요청 건
    sText = Trim(Format(varMFGDT, "@@@@.@@.@@"))  ' Format(varMfgDt, "@@@@-@@-@@")
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0375" + "0980"
    sFormatTx = sFormatTx + "Manufacture Date : " & sText & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "054" + "0425" + "0980"
    sFormatTx = sFormatTx & "B" & sText & vbCr
    
    'Expire Date =  Ship Date를 지정 +3년(MKT_CUST_2)    by ihson 2019.03.22
    sText = Trim(Format(varEXPDT, "@@@@.@@.@@"))
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0375" + "0450"
    sFormatTx = sFormatTx + "Expire Date : " & sText & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "054" + "0425" + "0450"
    sFormatTx = sFormatTx & "B" & sText & vbCr
    
    '분리선 ----------------------------------------------------------------------- (위/아래 동일한 내용 출력)
    sFormatTx = sFormatTx + "1X11" + "000" + "0480" + "0030" + "l" + "0980" + "0002" + vbCr
    
     '                      -출력각     --- Font ---    -Row-    -Col-
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0560" + "0980" '60
    sFormatTx = sFormatTx + "Vendor: PEAK(KOREA)" & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0600" + "0980"
    sFormatTx = sFormatTx + "Part No : " & I_DESC_NM & vbCr 'sText1
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0640" + "0980"
    sFormatTx = sFormatTx + "Package&Size&Thickness: " & varDescNm & "MM" & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0680" + "0980"
    sFormatTx = sFormatTx + "Temperature : " & sTemp_nm & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0720" + "0980"
    sFormatTx = sFormatTx + "Drawing No : " & I_DESC_NM & vbCr
        
    'Material Code
    sText = C_ITEM_NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0770" + "0980"
    sFormatTx = sFormatTx + "Material Code : " & sText & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "054" + "0825" + "0980"
    sFormatTx = sFormatTx & "B" & sText & vbCr
    
    'Lot No
    sText = Format(SHIP_DATE, "yyyymmdd")
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0770" + "0450"
    sFormatTx = sFormatTx + "Lot No : " & sText & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "054" + "0825" + "0450"
    sFormatTx = sFormatTx & "B" & sText & vbCr
                                                      
    'Manufacture Date =  Ship Date를 지정 2015.09.14 S+G 추가 요청 건
    sText = Trim(Format(varMFGDT, "@@@@.@@.@@"))
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0875" + "0980"
    sFormatTx = sFormatTx + "Manufacture Date : " & sText & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "054" + "0925" + "0980"
    sFormatTx = sFormatTx & "B" & sText & vbCr
    
    'Expire Date =  Ship Date를 지정 +3년(MKT_CUST_2)    by ihson 2019.03.22
    sText = Trim(Format(varEXPDT, "@@@@.@@.@@"))
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0875" + "0450"
    sFormatTx = sFormatTx + "Expire Date : " & sText & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "054" + "0925" + "0450"
    sFormatTx = sFormatTx & "B" & sText & vbCr
        
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_SCS_InnerLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2017-03-08
    ' 업  체 : SCS InnerLabel
    ' 개발자 : sjsung
    '======================================================================
    Dim sql                As String
    Dim sText               As String
    Dim varDescNm       As String
    Dim varDWG_no       As String
    Dim varLotNo        As String
    Dim varMFGDT        As String
    Dim varOEMiPART_NO  As String
    Dim varOEMiRev_NO   As String
    Dim sTemp_nm        As String
    Dim sMsg                As String
    Dim varUnitQty      As String
        
    '-------------------------------------------------------------
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    If varUnitQty Mod Val(I_BUNDLE_QT) <> 0 Or Trim(I_BUNDLE_QT) = 0 Or Trim(I_BUNDLE_QT) = "" Then
        MsgBox "Bundle 수량을 재확인하세요.", vbExclamation
        I_BUNDLE_QT.SetFocus
        Exit Sub
    Else
        varUnitQty = IIf(PRT_FORM(1).Value = True, varUnitQty \ Val(I_BUNDLE_QT), varUnitQty)   'I_BOX_UNIT_QT를 Bundle 수량만큼 나눈다. (정수로 표시)
    End If
    '-------------------------------------------------------------
    
    '============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"

    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr

    '                      -출력각     --- Font ---    -Row-    -Col-
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0100" + "0980"
    sFormatTx = sFormatTx + "Scs Partnumber : " & IIf(C_ITEM_NO = "", " ", C_ITEM_NO) & vbCr  'sText1
        
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0250" + "0980"
    sFormatTx = sFormatTx + "Peak Description : " & I_DESC_NM & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0400" + "0980"
    sFormatTx = sFormatTx + "Qty Per Bundle : " & varUnitQty & "PCS" + vbCr
                
    '분리선 ----------------------------------------------------------------------- (위/아래 동일한 내용 출력)
    sFormatTx = sFormatTx + "1X11" + "000" + "0480" + "0030" + "l" + "0980" + "0002" + vbCr
    
     '                      -출력각     --- Font ---    -Row-    -Col-
    
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0600" + "0980"
    sFormatTx = sFormatTx + "Scs Partnumber : " & IIf(C_ITEM_NO = "", " ", C_ITEM_NO) & vbCr  'sText1
    
    
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0750" + "0980"
    sFormatTx = sFormatTx + "Peak Description : " & I_DESC_NM & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0900" + "0980"
    sFormatTx = sFormatTx + "Qty Per Bundle : " & varUnitQty & "PCS" + vbCr
        
    sFormatTx = sFormatTx + "E" + vbCr
End Sub
Private Sub T4_SCK_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2011. 12. 21
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' SCK용 2D BarCode (QR CODE)
    ' 업  체 : SCK
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim varCustomNM             As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, varMFGDT, Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    
'    Select Case Trim(C_ITEM_NO)
'    Case "R019-2279C", "R019-1527C", "R019-2433C", "R019-2137C"
'        varEXPDT = DateAdd("yyyy", 1, varMFGDT)      '추가요청에 따라 Part# : R019-2279C,R019-1527C,R019-2433C,R019-2137C 인 경우는 유효일자를 1년으로 한다.
'    Case Else
'        varEXPDT = DateAdd("yyyy", 5, varMFGDT)      'Exp.Date = Lot 생성일자 + 2년 (날짜 형태를 @@@@-@@-@@로 정의 해야 2년을 더할 수 있음 )
'    End Select
    
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If


    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
        
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
        
    '===========================================================================
    '고객정보 가져오기
    '===========================================================================
    sql$ = "select CUSTOM_CD, WWCUSTOM_CD, custom_en CUSTOM_NM, CUSTOM_YN, ADDRE_TX from MKT_CUST_1 " & vbCrLf
    sql$ = sql$ & " where custom_cd = '" & Trim(Right(C_CUSTOM_CD, 10)) & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User

    If SSDATA.RecordCount > 0 Then
        varCustomNM = Trim(SSDATA!custom_nm)    'Custom_NM
    Else
        varCustomNM = Trim(Left(C_CUSTOM_CD, 30))     'Custom_YN
    End If
    SSDATA.Close

'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
       
    '여백
    'QR코드는 숫자 최대 7089자, 문자(ASCII) 최대 4296자, 이진8비트 최대 2953바이트, 한자등 아시아 문자 최대 1817자를 담을 수 있다.
    '2D BARCODE(QR CODE)   'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로
    'a W1 b c d eee ffff gggg hh…h 1 W1 d 4 4 000 0010 0010This is the data portion
    '                               3 W1 D 4 4 000 0320 0200 2HM,
    sFormatTx = sFormatTx & "3W1D55000065002702HM," & C_ITEM_NO & "/" & C_LOT_ID & "/" & varUnitQty & "/" & C_PO_NO & "/" & varMFGDT & Chr$(13)
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0985"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + varCustomNM + vbCr
    
   'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0240" + "0985"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '  I_PART_NO + vbCr

    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0315" + "0985"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0315" + "0730"
    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr
    
    'REV (2015.12.31 S+G 김미진 요청으로 추가 진행)
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0315" + "0260"
    sFormatTx = sFormatTx + "Rev. " + varOEMiRev_NO + vbCr
    
    'SHIP'G LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0390" + "0985"
    sFormatTx = sFormatTx + "SHIP'G LOT#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0390" + "0730"
    sFormatTx = sFormatTx + ": " + C_LOT_ID + vbCr
    
    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0460" + "0985"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0460" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
    
    'P/O NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0530" + "0985"
    sFormatTx = sFormatTx + "P/O NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0530" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + vbCr
    
    'MFG DATE (출하일 하루전 통일) 'QA 성적서와 Mfg Date 정보가 틀어지는 경우 발생  'Format(SHIP_DATE, "@@@@-@@-@@"))
    'varMfgDt LOG생성일자 기준으로 변경 2012-08-13 Sng 전대열, 백지훈 요청사항
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0600" + "0985"
    sFormatTx = sFormatTx + "MFG DATE" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0600" + "0730"
    sFormatTx = sFormatTx + ": " + varMFGDT + vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0670" + "0985"
    sFormatTx = sFormatTx + "EXP DATE" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0670" + "0730"
    sFormatTx = sFormatTx + ": " + varEXPDT + vbCr
    
    'COO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0740" + "0985"
    sFormatTx = sFormatTx + "COO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0740" + "0730"
    sFormatTx = sFormatTx + ": " + GLO_COMP_CO + vbCr


' ------------------------------------------------------------------------------------------------------
' 2013.01.11 SNG_Evan_baek SCK요청으로 CUST'R P/N , SHIP'G LOT NO Font 2배확대 함 (샘플확인후 이전양식으로 재요청함, 보류처리)
' ------------------------------------------------------------------------------------------------------
'    'CUST'R P/N
'    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0350" + "0990"
'    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
'    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0350" + "0605"
'    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr
'
'    'SHIP'G LOT NO
'    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0440" + "0990"
'    sFormatTx = sFormatTx + "SHIP'G LOT#" + vbCr
'    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0440" + "0605"
'    sFormatTx = sFormatTx + ": " + C_PO_NO + "-" + Format(SHIP_DATE, "mmddyy") + vbCr
'
'    'QTY
'    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0985"
'    sFormatTx = sFormatTx + "QTY" + vbCr
'    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
'    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
'
'    'P/O NO
'    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0985"
'    sFormatTx = sFormatTx + "P/O NO" + vbCr
'    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
'    sFormatTx = sFormatTx + ": " + C_PO_NO + vbCr
'
'    'MFG DATE (출하일 하루전 통일) 'QA 성적서와 Mfg Date 정보가 틀어지는 경우 발생
'    'varMfgDt LOG생성일자 기준으로 변경 2012-08-13 Sng 전대열, 백지훈 요청사항
'    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0985"
'    sFormatTx = sFormatTx + "MFG DATE" + vbCr
'    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0730"
'    sFormatTx = sFormatTx + ": " + varMfgDt + vbCr
'
'    '여백
'    'QR코드는 숫자 최대 7089자, 문자(ASCII) 최대 4296자, 이진8비트 최대 2953바이트, 한자등 아시아 문자 최대 1817자를 담을 수 있다.
'    '2D BARCODE(QR CODE)   'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로
'    'a W1 b c d eee ffff gggg hh…h 1 W1 d 4 4 000 0010 0010This is the data portion
'    '                               3 W1 D 4 4 000 0320 0200 2HM,
'    sFormatTx = sFormatTx & "3W1D55000072002602HM," & C_ITEM_NO & "/" & C_PO_NO + "-" + Format(SHIP_DATE, "mmddyy") & "/" & varUnitQty & "/" & C_PO_NO & "/" & varMfgDt & Chr$(13)
'
'    'COO
'    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0985"
'    sFormatTx = sFormatTx + "COO" + vbCr
'    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
'    sFormatTx = sFormatTx + ": " + GLO_COMP_CO  + vbCr
' ------------------------------------------------------------------------------------------------------

    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0985"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & "/" & I_END_CT_NO & vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub TR_SCK_innerLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2013.11.08
    ' SCK CoverTape InnerLabel
    ' 수정자 :  by jwkim
    '======================================================================
    Dim sText1                  As String
    Dim sText2                  As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim varCustomName           As String
    Dim varPO                   As String
    Dim varLotNo                As String   'Lot#
    Dim varPrdDesc              As String
    Dim varDrawingNo            As String
    Dim varSPECNo               As String
    Dim varITEMNo               As String   'Item_no
    Dim varIPartNo              As String   'IPart_no
    Dim varMFGDT                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
        
    Dim sql                     As String
        
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), C_LOT_NO), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    
    '===========================================================================
    ' SMP 정보 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
        
    '===========================================================================
    '고객정보 가져오기
    '===========================================================================
    sql$ = "select CUSTOM_CD, WWCUSTOM_CD, CUSTOM_EN , CUSTOM_YN, ADDRE_TX from MKT_CUST_1 " & vbCrLf
    sql$ = sql$ & " where custom_cd = '" & Trim(Right(C_CUSTOM_CD, 10)) & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
    
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User

    If SSDATA.RecordCount > 0 Then
        varCustomName = Trim(SSDATA!CUSTOM_EN)     'Custom_NM
    Else
        varCustomName = Trim(Left(C_CUSTOM_CD, 30))     'Custom_YN
    End If
    SSDATA.Close
    
    varPO = C_PO_NO
    varLotNo = C_LOT_NO
    'varPrdDesc = GF_TrucateCR(I_DESC_NM) & " " & I_METERIAL_NM
    varITEMNo = C_ITEM_NO
    varIPartNo = varOEMiPART_NO
     
    ' Angle 정의
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"

    ' Horizontal expansion rate 정의
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

    ' Vertical expansion rate 정의
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

    ' Font 9 일때 Size 정의
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

    ' Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
        
    '초기 정의값
    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
'=========================================================================================================
'    1       234     567     8901     2345     6
'Angle180 + "911" + "003" + "0065" + "0950" + "Description " + vbCr
'1              : 출력각 조절        1 = 0:  2 = 90:   3 = 180:    4 = 270
'2              : Font 결정          0 ~ 8 : System Font     9 : Smooth Font
'3              : Font의 가로비율    1 ~ 9, A ~ O  (1 ~ 24 까지 의미 1~9, 10~24)
'4              : Font의 세로비율    1 ~ 9, A ~ O  (1 ~ 24 까지 의미 1~9, 10~24)
'5, 6, 7        : Font가 Smooth 일때 값 결정(2번째 자리가 9일때)  001(6 Pt) ~ 010(48 Pt)
'8, 9, 10, 11   : Row 의 위치 0000~9999  단위 : 0.01 inch/step
'12, 13, 14, 15 : Col 의 위치 0000~0416  단위 : 0.01 inch/step
'16             : 문자 입력
    
    Dim varRate     As String
    varRate = "52"
     
    sText1 = C_ITEM_NO
    sFormatTx = sFormatTx & Angle180 & "911" & "004" & "0070" & "0560" & "SAP Code : " & sText1 & vbCr
    sFormatTx = sFormatTx & Angle180 & "E" & varRate & "080" & "0180" & "0560"   '1D Barcode Type "CODE 128"로 변경
    sFormatTx = sFormatTx + "B" + sText1 & vbCr 'Code 128 SubSET B
    
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_ATK_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2011. 12. 21
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' OEM P/N 사용, Address 출력시 문자열 길이에 따라 ROW 조절
    ' 업  체 : ATK
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim varCustomCd             As String
    Dim varPdescNm              As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    
    Dim subSpace                As String
    
    Dim charRowAdd(5 To 48)     As Integer
    varLot_no = C_LOT_NO

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no , " & vbCrLf
    sql$ = sql$ & " (select distinct pdesc_nm from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.ipart_no=f.ipart_no ) as pdesc_nm "
    sql$ = sql$ & " from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varPdescNm = ""
    If SSDATA.RecordCount > 0 Then
        varPdescNm = SSDATA!pdesc_nm
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
                
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0005" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0800" + "0030" + "l" + "0980" + "0004" + vbCr
    
    '중간 가로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0230" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0300" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0370" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0440" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0510" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0580" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0650" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0720" + "0030" + "l" + "0980" + "0004" + vbCr
        
    '중간 세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0370" + "0320" + "l" + "0004" + "0075" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0510" + "0320" + "l" + "0004" + "0075" + vbCr
    
    '세로줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "1005" + "l" + "0005" + "0700" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0005" + "0700" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0720" + "l" + "0004" + "0700" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'ATK Batch No (공백)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "ATK Batch No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0700"
    sFormatTx = sFormatTx + vbCr
    
   'Item No(Sid No)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0280" + "0980"
    sFormatTx = sFormatTx + "Item No(Sid No)" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0280" + "0700"
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr

    'Description
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0350" + "0980"
    sFormatTx = sFormatTx + "Description" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0350" + "0700"
    sFormatTx = sFormatTx + varPdescNm + vbCr
    
    'Carton No/Qty
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0980"
    sFormatTx = sFormatTx + "Carton No / Qty" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0420" + "0700"
    sFormatTx = sFormatTx & varCartonNo & "/" & varTotalCartonNo & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0310"
    sFormatTx = sFormatTx + varUnitQty + " PCS" + vbCr
    
    'Drawing No / Rev No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0495" + "0980"
    sFormatTx = sFormatTx + "Drawing No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0495" + "0700"
    sFormatTx = sFormatTx + I_DESC_NM + vbCr

    'PO No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0565" + "0980"
    sFormatTx = sFormatTx + "PO NO / Rev" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0565" + "0700"
    sFormatTx = sFormatTx + C_PO_NO & "-" & "000" & vbCr   '000 고정값 사용 (S+G 김세일 사원요청 건 2013.12.12)
    'sFormatTx = sFormatTx + C_PO_NO & "-" & Format(varOEMiRev_NO, "00#") & vbCr  ' I_PART_NO
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0560" + "0310"
    sFormatTx = sFormatTx + Trim(varOEMiRev_NO) + vbCr ' Format(varOEMiRev_NO, "00#") + vbCr
    
    'Peak P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0640" + "0980"
    sFormatTx = sFormatTx + "Peak P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0640" + "0700"
    sFormatTx = sFormatTx + varOEMiPART_NO + vbCr  '
    
    'Supplier Lot No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0710" + "0980"
    sFormatTx = sFormatTx + "Supplier Lot No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0710" + "0700"
    sFormatTx = sFormatTx & C_PO_NO & "-" & "000" & "-" & Format(SHIP_DATE, "mmddyy") & vbCr    '000 고정값 사용 (S+G 김세일 사원요청 건 2013.12.12)
    'sFormatTx = sFormatTx & C_PO_NO & "-" & Format(varOEMiRev_NO, "00#") & "-" & Format(SHIP_DATE, "mmddyy") & vbCr
        
    'Supplier Name
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0780" + "0980"
    sFormatTx = sFormatTx + "Supplier Name" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0780" + "0700"
    sFormatTx = sFormatTx + "PEAK" + vbCr
    
    '업체별 주소가져오기
    varCustomCd = Trim(Right(C_CUSTOM_CD, 10))
    'TEST varCustomCD = "100001"
    sql$ = "select custom_cd, CUSTOM_EN  custom_nm, addre_tx from MKT_CUST_1 "
    sql$ = sql$ & " where custom_cd = '" & varCustomCd & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf

    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User

    If SSDATA.RecordCount > 0 Then
        'Address
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0980"
        sFormatTx = sFormatTx + "Address" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0840"
        sFormatTx = sFormatTx + ":  " & SSDATA!custom_nm & vbCr
        If Len(Trim(SSDATA!addre_tx)) > 46 Then '주소길이에 따라 3줄로 표시한다.
            sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0910" + "0980"
            sFormatTx = sFormatTx + " " & Left(SSDATA!addre_tx, 46) & vbCr
            sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0960" + "0980"
            sFormatTx = sFormatTx + " " & Mid(SSDATA!addre_tx, 47, Len(SSDATA!addre_tx)) & vbCr
        Else
            sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0910" + "0980"
            sFormatTx = sFormatTx + " " & SSDATA!addre_tx & vbCr
        End If
    Else
        '업체정보가 존재하지 않음 (공백으로 표시)
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0980"
        sFormatTx = sFormatTx + "Address : " + vbCr
    End If
    SSDATA.Close
    
    sFormatTx = sFormatTx + "E" + vbCr
    
End Sub

Private Sub T4_Renesas_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2013. 05. 13
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' 업  체 : Renesas
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim sMonth                  As String
    Dim varLot_no               As String
    Dim varMoldSeq              As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim varCustomCd             As String
    Dim subSpace                As String
    
    Dim charRowAdd(5 To 48)     As Integer
    varLot_no = C_LOT_NO
    varMoldSeq = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2)

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = varMoldSeq & I_REV_NO
        Else
            sText2 = sText2 & " " & varMoldSeq & I_REV_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = varMoldSeq & I_REV_NO
    End If
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
   
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0030" + "0030" + "l" + "0980" + "0006" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0230" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0930" + "0030" + "l" + "0980" + "0006" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
         
   'ITEM
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0100" + "0980"
    sFormatTx = sFormatTx + "ITEM" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0100" + "0710"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
    
   'CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0180" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0180" + "0710"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr

    Select Case Format(SHIP_DATE, "mm")           'M
        Case "10":  sMonth = "X"
        Case "11":  sMonth = "Y"
        Case "12":  sMonth = "Z"
        Case Else:  sMonth = Format(SHIP_DATE, "mm")
    End Select
    
    'RENESAS LOT#
    sText1 = "G" & Right(Format(SHIP_DATE, "yyyy"), 1) & sMonth & Format(SHIP_DATE, "dd") & Right(varMoldSeq, 1)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0260" + "0980"
    sFormatTx = sFormatTx + "RENESAS LOT#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0260" + "0710"
    sFormatTx = sFormatTx + ": " & sText1 & vbCr
    
    'LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0340" + "0980"
    sFormatTx = sFormatTx + "LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0340" + "0710"
    sFormatTx = sFormatTx + ": " & varLot_no & vbCr
       
    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0980"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0710"
    sFormatTx = sFormatTx + ": " + varUnitQty + " PCS" + vbCr
    
    'PO No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0980"
    sFormatTx = sFormatTx + "PO NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0710"
    sFormatTx = sFormatTx + ": " + C_PO_NO & vbCr ' I_PART_NO

    'MOLD NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0580" + "0980"
    sFormatTx = sFormatTx + "MOLD NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0580" + "0710"
    sFormatTx = sFormatTx + ": " & "SG-" & Right(varMoldSeq, 1) & vbCr '
    
    'SUPPLIER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0980"
    sFormatTx = sFormatTx + "SUPPLIER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0710"
    sFormatTx = sFormatTx + ": " & "S+G COMPANY" & vbCr
    
    'MFG COUNTRY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0740" + "0980"
    sFormatTx = sFormatTx + "MFG COUNTRY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0740" + "0710"
    sFormatTx = sFormatTx + ": " + "MADE IN KOREA" + vbCr
    
    'C/T NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0820" + "0980"
    sFormatTx = sFormatTx + "C/T NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0820" + "0710"
    sFormatTx = sFormatTx + ": " & varCartonNo & "/" & varTotalCartonNo & vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr
    
End Sub

Private Sub T4_ATK_old_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2011. 12. 21
    ' Daewon기준으로 Code 변경 (S+G : LPT 출력)
    ' OEM P/N 사용, Address 출력시 문자열 길이에 따라 ROW 조절
    ' 업  체 : ATK
    ' 개발자 : ICS_jwkim
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim varCustomCd             As String
    Dim varDWG_no               As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    
    Dim subSpace                As String
    
    Dim charRowAdd(5 To 48)     As Integer
    varLot_no = C_LOT_NO

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " (select distinct dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ) as dwg_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    If SSDATA.RecordCount > 0 Then
        varDWG_no = SSDATA!dwg_no
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
    
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
    
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:


    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0230" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0800" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'ATK Batch No (공백)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "ATK Batch No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + vbCr
    
   'Item No(Sid No)
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0280" + "0980"
    sFormatTx = sFormatTx + "Item No(Sid No)" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0280" + "0730"
    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr

    'Description
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0350" + "0980"
    sFormatTx = sFormatTx + "Description" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0350" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
    
    'Carton No/Qty
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0980"
    sFormatTx = sFormatTx + "Carton No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0420" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & "/" & varTotalCartonNo & vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0320"
    sFormatTx = sFormatTx + "Qty : " + varUnitQty + " PCS" + vbCr

    'Drawing No / Rev No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0490" + "0980"
    sFormatTx = sFormatTx + "Drawing No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0490" + "0730"
    sFormatTx = sFormatTx + ": " + I_DESC_NM + vbCr
    
    'PO No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0560" + "0980"
    sFormatTx = sFormatTx + "PO NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0560" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO & "-" & Format(varOEMiRev_NO, "00#") & vbCr ' I_PART_NO
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0560" + "0320"
    sFormatTx = sFormatTx + "Rev No : " + Format(varOEMiRev_NO, "00#") + vbCr
    
    'Peak P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0640" + "0980"
    sFormatTx = sFormatTx + "Peak P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0640" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '
    
    'Supplier Lot No
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0710" + "0980"
    sFormatTx = sFormatTx + "Supplier Lot No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0710" + "0730"
    sFormatTx = sFormatTx + ": " & C_PO_NO & "-" & Format(varOEMiRev_NO, "00#") & "-" & Format(SHIP_DATE, "mmddyy") & vbCr
    
    'Supplier Name
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0780" + "0980"
    sFormatTx = sFormatTx + "Supplier Name" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0780" + "0730"
    sFormatTx = sFormatTx + ": " + "PEAK" + vbCr
    
    '업체별 주소가져오기
    varCustomCd = Trim(Right(C_CUSTOM_CD, 10))
    'TEST varCustomCD = "100001"
    sql$ = "select custom_cd, CUSTOM_EN  custom_nm, addre_tx from MKT_CUST_1 "
    sql$ = sql$ & " where custom_cd = '" & varCustomCd & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf

    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User

    If SSDATA.RecordCount > 0 Then
        'Address
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0980"
        sFormatTx = sFormatTx + "Address" + vbCr
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0840"
        sFormatTx = sFormatTx + ":  " & SSDATA!custom_nm & vbCr
        If Len(Trim(SSDATA!addre_tx)) > 46 Then '주소길이에 따라 3줄로 표시한다.
            sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0910" + "0980"
            sFormatTx = sFormatTx + " " & Left(SSDATA!addre_tx, 46) & vbCr
            sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0960" + "0980"
            sFormatTx = sFormatTx + " " & Mid(SSDATA!addre_tx, 47, Len(SSDATA!addre_tx)) & vbCr
        Else
            sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0910" + "0980"
            sFormatTx = sFormatTx + " " & SSDATA!addre_tx & vbCr
        End If
    Else
        '업체정보가 존재하지 않음 (공백으로 표시)
        sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0860" + "0980"
        sFormatTx = sFormatTx + "Address : " + vbCr
    End If
    SSDATA.Close
    
    sFormatTx = sFormatTx + "E" + vbCr
    
End Sub

Private Sub TR_TSMC_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2014.04.07
    ' Tray TSMC BarCode를 작성하기 위해 작성된 루틴임 - jwkim
    '======================================================================
    Dim sText1                  As String
    Dim sText2                  As String
    Dim sql                     As String
    Dim charRowAdd(5 To 48)     As Integer
    
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim varCustomName           As String
    Dim varPO                   As String
    Dim varUnitQty              As String   'Quantity
    Dim varMatrBatch            As String   'Lot#
    Dim varPrdDesc              As String
    Dim varDrawingNo            As String
    Dim varSPECNo               As String
    Dim varITEMNo               As String   'Item_no
    Dim varIPartNo              As String   'IPart_no
    Dim varCustSuppCD           As String
    Dim varMFGDT                As String
    Dim varMoldStamp           As String
    Dim varOEMMold_no           As String
    Dim varInvoiceNo    As String
    Dim varQApassDT               As String
    
    varQApassDT = Format(SHIP_DATE, "yyyy.mm.dd")
    
    varInvoiceNo = C_INV_NO
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), C_LOT_NO), "@@@@@@@@") 'Lot 생성일자를 가져옴
    
    varCustomName = Trim(Left(C_CUSTOM_CD, 30))
    varPO = C_PO_NO
    If GLO_COMP_CD = "FU" Then
        'varPrdDesc = Trim(Left(C_PROD_KB, 30)) & " " & GF_TrucateCR(I_DESC_NM)    '제품군 + 제품명
        varPrdDesc = GF_TrucateCR(I_DESC_NM)
    Else
        varPrdDesc = GF_TrucateCR(I_DESC_NM) & " " & I_METERIAL_NM & "-" & I_COLOR_NM
    End If
    varITEMNo = C_ITEM_NO
    varIPartNo =  I_PART_NO
    
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    ' MOLD STAMP by FRC Jim Chen 2020/02/06
    '===========================================================================
    varMoldStamp = ""
    varOEMMold_no = ""
    If GLO_COMP_CD = "FU" Then
        sql$ = " select distinct a.prod_cd, b.prod_gubun, a.mold_no, a.mold_desc, " & vbCrLf
        sql$ = sql$ & " nvl((select mold_stamp from prd_mold_1 f where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and f.mold_no =  b.mold_no), ' ') mold_stamp " & vbCrLf
        sql$ = sql$ & " from PRD_MAST_1_VIEW a, PRD_WORK_11 b " & vbCrLf
        sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
        sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        sql$ = sql$ & " and a.prod_cd = b.prod_cd "
        sql$ = sql$ & " and a.comp_cd = b.comp_cd and a.bsns_cd = b.bsns_cd " & vbCrLf
        If varMoldNo = "ASSEMBLY" Then   'Assembly 제품의 경우 Main 제품의 LOT로 비교한다.
            sql$ = sql$ & " and   b.outlot_no in (select distinct outlot_no from prd_work_31 where lot_no= '" & Replace(Trim(C_LOT_NO), "-", "") & "' and main_kb ='1' ) " & vbCrLf
        Else
            sql$ = sql$ & " and   b.outlot_no = '" & Replace(Trim(C_LOT_NO), "-", "") & "' " & vbCrLf
        End If
        
        Set SSDATA = Nothing
        SSDATA.Open sql$, DB_User
        
        If SSDATA.RecordCount = 0 Then
            varMoldStamp = ""
        Else
            varMoldStamp = Trim(SSDATA!mold_stamp)
        End If
        SSDATA.Close
     
        'Assembly P-Bin Management and system will no write any main part information in Prd_work_31. Thus we need to get main part in Prd_BOM_1
        'FRC Jim Chen 2020/02/06
        If UCase(varMoldNo) = "ASSEMBLY" Then     'FRCUNION은 Add varOEMMold_no by jim
             sql$ = " select distinct a.prod_cd, a.mold_no, a.mold_desc " & vbCrLf
             sql$ = sql$ & " from PRD_MAST_1_VIEW a" & vbCrLf
             sql$ = sql$ & " where a.comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
             sql$ = sql$ & " and a.bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
'             sql$ = sql$ & " and a.prod_cd = (select child_cd from PRD_BOM_1 where parent_cd =(select distinct P_PROD_CD from prd_work_31 where LOT_NO ='" & Trim(C_LOT_NO) & "')" & vbCrLf
'             sql$ = sql$ & "and child_cd like 'TR%')" & vbCrLf
             sql$ = sql$ & " and a.mold_no = (SELECT DISTINCT MASTER_MOLD_NO from PRD_WORK_1 where LOT_NO  ='" & Trim(C_LOT_NO) & "')" & vbCrLf
         
             Set SSDATA = Nothing
             SSDATA.Open sql$, DB_User
             
             If SSDATA.RecordCount = 0 Then
                 varOEMMold_no = ""
             Else
                 varOEMMold_no = str(Int(Right(Trim(SSDATA!mold_desc), 2)))
             End If
             SSDATA.Close
        End If
    End If
    
    '===========================================================================
    ' Angle 정의
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"

    ' Horizontal expansion rate 정의
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

    ' Vertical expansion rate 정의
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

    ' Font 9 일때 Size 정의
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

    ' Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    '초기 정의값
    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
'=========================================================================================================
'    1       234     567     8901     2345     6
'Angle180 + "911" + "003" + "0065" + "0950" + "Description " + vbCr
'1              : 출력각 조절        1 = 0:  2 = 90:   3 = 180:    4 = 270
'2              : Font 결정          0 ~ 8 : System Font     9 : Smooth Font
'3              : Font의 가로비율    1 ~ 9, A ~ O  (1 ~ 24 까지 의미 1~9, 10~24)
'4              : Font의 세로비율    1 ~ 9, A ~ O  (1 ~ 24 까지 의미 1~9, 10~24)
'5, 6, 7        : Font가 Smooth 일때 값 결정(2번째 자리가 9일때)  001(6 Pt) ~ 010(48 Pt)
'8, 9, 10, 11   : Row 의 위치 0000~9999  단위 : 0.01 inch/step
'12, 13, 14, 15 : Col 의 위치 0000~0416  단위 : 0.01 inch/step
'16             : 문자 입력
       
' 외부 Box 그리기           -- reserved --   -Row-    -Col-   -box- -수평길이+수직길이- -- line 폭 --
    'sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0030" + "b" + "0970" + "0570" + "0006" + "0006" + vbCr
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0035" + "l" + "0960" + "0006" + vbCr '맨 윗줄
    sFormatTx = sFormatTx + "1X11" + "000" + "0565" + "0035" + "l" + "0960" + "0006" + vbCr '맨 아랫줄
' Barcode 상단의 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0090" + "0035" + "l" + "0960" + "0006" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0165" + "0035" + "l" + "0960" + "0006" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0005" + "0525" + "l" + "0006" + "0085" + vbCr '세로
    
    Dim varRate     As String
    'varRate = IIf(opt_5to2, "52", "42")
    
    varRate = "52"
    '                      -출력각     --- Font ---    -Row-    -Col-
    sText1 = "From"
    sFormatTx = sFormatTx + Angle180 + "923" + "001" + "0080" + "0980" + sText1 + vbCr
    sText1 = IIf(GLO_COMP_CD = "FU", "PEAK", GLO_COMP_NM)   '2018.05.09 jeenie & Liam 변경 요청 건
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0080" + "0860" + sText1 + vbCr
    
    sText1 = "To"
    sFormatTx = sFormatTx + Angle180 + "923" + "001" + "0080" + "0500" + sText1 + vbCr
    sText1 = varCustomName
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0080" + "0430" + sText1 + vbCr
    
    ' 문자 표시
' Descriptiion
     sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0150" + "0980" + varPrdDesc + vbCr
    
    ' Barcode 표시         -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
' TSMC P/N #
    sText1 = varITEMNo
    sFormatTx = sFormatTx + Angle180 + "a" + varRate + "054" + "0235" + "0975"
    sFormatTx = sFormatTx + sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "002" + "0270" + "0980" + "TSMC P/N : " + sText1 + vbCr

    ' Barcode 표시         -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께]
    
    If GLO_COMP_CD = "FU" Then 'FRCUNION은 Add MOLD STAMP by jim
        'Mold No
        sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0235" + "0400" + "MOLD# : " + IIf(Trim(varMoldStamp) = "", Trim(varOEMMold_no), Trim(varMoldStamp)) + vbCr
    End If
    
' DW Lot #
    sText1 = IIf(GLO_COMP_CD = "FU", varInvoiceNo, Trim(C_LOT_NO))
    sFormatTx = sFormatTx + Angle180 + "a" + varRate + "054" + "0330" + "0975"
    sFormatTx = sFormatTx + sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "002" + "0365" + "0980" + "LOT NO : " + sText1 + vbCr
    
    ' Barcode 표시         -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
' QTY
    sText1 = Trim(varUnitQty)
    sFormatTx = sFormatTx + Angle180 + "a" + varRate + "054" + "0425" + "0975"
    sFormatTx = sFormatTx + sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "002" + "0460" + "0980" + "QTY : " + sText1 + " PCS" + vbCr
    
    ' Barcode 표시         -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
' DOM (=Mfg Date)
    sText1 = Format(Trim(varMFGDT), "@@@@@@@@")
    sFormatTx = sFormatTx + Angle180 + "a" + varRate + "054" + "0520" + "0975"
    sFormatTx = sFormatTx + sText1 + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "002" + "0555" + "0980" + "DOM : " + sText1 + vbCr
    
    'QA ACCEPT STAMP
'                         Rotation                       ROW     Col                           Radius
    'sFormatTx = sFormatTx + "1" + "X" + "11" + "000" + "0400" + "0220" + "C" + "001" + "0001" + "0088" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0400" + "0250"
    'sFormatTx = sFormatTx & "QA" & vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0440" + "0280"
    'sFormatTx = sFormatTx & "ACCEPT" & vbCr
'  QA Pass Date
    sFormatTx = sFormatTx + Angle180 + "911" + "002" + "0555" + "0380" + "( QA PASS : " + varQApassDT + " )" + vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr

End Sub

Private Sub T4_ASEKS_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2019. 11. 08
    ' DEVELOPER : FRC IT Jim Chen
    '======================================================================
    Dim sql                          As String
    Dim sText1                   As String
    Dim sText2                   As String
    Dim varLot_no            As String
    Dim varUnitQty           As String
    Dim varQApass           As String
    Dim subSpace             As String
    Dim varMFGDT           As String
    Dim varEXPDT             As String
    Dim yyyy                      As String
    Dim MM                      As String
    Dim dd                         As String
    
    Dim charRowAdd(5 To 48)     As Integer
        
    'RT
    varLot_no = C_LOT_NO
    
    'MFG DATE
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    varMFGDT = Replace(varMFGDT, "-", "/")
        
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    varLot_no = Replace(C_LOT_NO, "-", "")
    
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    'Barcode    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
                            
    'Batch NO (RT)
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0400" + "1000"
    sFormatTx = sFormatTx + "Batch NO:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0400" + "0610"
    sFormatTx = sFormatTx + Trim(C_LOT_ID) + vbCr
        
    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "1000"
    sFormatTx = sFormatTx + "Part NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0500" + "0610"
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr

    'Verdenor LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0600" + "1000"
    sFormatTx = sFormatTx + "Vendor Lot" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0600" + "0610"
    sFormatTx = sFormatTx & varLot_no & vbCr

    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0700" + "1000"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0700" + "0610"
    sFormatTx = sFormatTx + varUnitQty + vbCr
    
    'DOM(YYYY/MM/dd) 'LOT (MFG Date)
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0800" + "1000"
    sFormatTx = sFormatTx + "Manufacture Date:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0800" + "0610"
    sFormatTx = sFormatTx + varMFGDT + vbCr
    
    'C/T (Box No)
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0900" + "1000"
    sFormatTx = sFormatTx + "Box NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0900" + "0610"
    sFormatTx = sFormatTx & varCartonNo & vbCr
    
    '2D QR Code
    sFormatTx = sFormatTx & "3W1D660000300100002HM, " & Trim(C_LOT_ID) & "|" & C_ITEM_NO & "|" & varLot_no & "|" & varUnitQty & "|" & varMFGDT & Chr$(13)
    
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_ASESH_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2019. 11. 08
    ' DEVELOPER : FRC JIM CHEN
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim charRowAdd(5 To 48)     As Integer
    Dim varDWG_no               As String
    Dim varPrdITEM_no           As String
    Dim varMFGDT                            As String
    Dim varEXPDT                As String
    
    varLot_no = C_LOT_NO
        
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, Trim(varMFGDT), Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    
    'varEXPDT = Replace(DateAdd("yyyy", 5, varMFGDT), "-", "")
    varMFGDT = Replace(varMFGDT, "-", "/")
    varEXPDT = Replace(varEXPDT, "-", "/")

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    '===========================================================================
    ' SMP 정보 확인, Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " (select distinct nvl(item_no, ' ') item_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as prd_item_no, " & vbCrLf
    sql$ = sql$ & " (select distinct nvl(dwg_no, ' ') dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as dwg_no" & vbCrLf
    sql$ = sql$ & " from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    varPrdITEM_no = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
        varDWG_no = SSDATA!dwg_no
        varPrdITEM_no = SSDATA!prd_item_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
   
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
   'CUST'R
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
    
   'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0240" + "0980"
    sFormatTx = sFormatTx + "CUST'R PO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0240" + "0730"
    sFormatTx = sFormatTx + ": " + C_PO_NO + vbCr

    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0310" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0310" + "0730"
    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr
    
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0380" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '  I_PART_NO + vbCr
    
    'PEAK DWG#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0450" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0450" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr

        'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0520" + "0980"
    sFormatTx = sFormatTx + "QTY" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0520" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
        
        'MFG DATE
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0590" + "0980"
    sFormatTx = sFormatTx + "MFG" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0590" + "0730"
    sFormatTx = sFormatTx + ": " + varMFGDT + vbCr
        
        'SHIP'G LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0660" + "0980"
    sFormatTx = sFormatTx + "LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0660" + "0730"
    sFormatTx = sFormatTx + ": " + varLot_no + vbCr
        
        'EXP DATE
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"
    sFormatTx = sFormatTx + "EXP DATE" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0730" + "0730"
    sFormatTx = sFormatTx + ": " + varEXPDT + vbCr
    
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    sFormatTx = sFormatTx + "E" + vbCr
End Sub

Private Sub T4_Micron_SGP_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2021. 06. 11
    ' Customer : Micron SINGAPORE
    ' Author : FRC IT Jim Chen
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varMoldSeq              As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varDWG_no               As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim strLOT_ID               As String
    
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    varMoldSeq = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2)
    
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@@@@@") 'Lot 생성일자를 가져옴
    'varMFGDT = Left(varMFGDT, 6) & "01" '20170918 생산월 01로 고정 표시 ihson , CPL hilee 2020.11.24 실제 생산일자를 표시함 (모든 SPIL SITE 공통적용)
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
      ' 해당월의 출하 횟수를 확인
    ' Display Shipping Date for the C_LOT_NO for Delivery Order for the Lot No
    
    sql$ = "select ipart_no , Prev_No,   'P' || '" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "' || trim(to_char(ROWNUM, '00')) as LOT_ID"
    sql$ = sql$ & " from ( select distinct ipart_no, Prev_No from MKT_OUT_11 a where a.comp_cd = 'SG' and a.bsns_cd = '11' and a.outorder_no ='" & Trim(C_LOT_ID) & "'"
    sql$ = sql$ & ")"
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
            
    strLOT_ID = ""
    If SSDATA.RecordCount > 0 Then
        Do Until SSDATA.EOF
            If Trim(SSDATA!ipart_no) = Trim( I_PART_NO) And Trim(SSDATA!Prev_No) = Trim(I_REV_NO) Then strLOT_ID = Trim(SSDATA!LOT_ID)
            SSDATA.MoveNext
        Loop
    End If
    SSDATA.Close
    If Trim(strLOT_ID) = "" Then strLOT_ID = "P" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "01"

    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no,  nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as dwg_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    If SSDATA.RecordCount > 0 Then
        varDWG_no = SSDATA!dwg_no
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
 
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
 
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' 외부 Box 그리기           -- reserved --   -Row-    -Col-   -box- -수평길이+수직길이- -- line 폭 --
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0030" + "b" + "0980" + "0950" + "0004" + "0004" + vbCr

    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    
    '2D QR Code Request by FRC CPL Liam 2020-06-05
    'If GLO_COMP_CD = "FU" Then
        'sFormatTx = sFormatTx & "3W1D66000042002902HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)
    'End If

    '2D BARCODE(QR CODE)    'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로, b,c = Module Size
    'sFormatTx = sFormatTx & "3W1D33000030502302HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)  'FRC '2D QR Code Request by FRC CPL Liam 2020-06-05
    'sFormatTx = sFormatTx & "3W1D55000063002202HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)



    ' 가로줄 'Vertial Line
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0004" + vbCr '1
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0030" + "l" + "0980" + "0004" + vbCr '2
    If GLO_COMP_CD = "FU" Then
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr '3
    Else
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr
    End If
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0030" + "l" + "0980" + "0004" + vbCr '4----(-40)430-390
    'sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr    test
    sFormatTx = sFormatTx + "1X11" + "000" + "0500" + "0030" + "l" + "0980" + "0004" + vbCr '5--550-510
    
    
'20220720 Originel code  jj
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0030" + "l" + "0980" + "0004" + vbCr '6----
    'sFormatTx = sFormatTx + "1X11" + "000" + "0740" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0810" + "0030" + "l" + "0980" + "0004" + vbCr '7------
    'sFormatTx = sFormatTx + "1X11" + "000" + "0880" + "0030" + "l" + "0980" + "0004" + vbCr
    
 
    
    
    '세로줄

    'If GLO_COMP_CD = "FU" Then
        'If Trim(Right(C_CUSTOM_CD, 10)) = "21311" Then
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0365" + vbCr
        'Else
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0240" + vbCr
        'End If
    'End If
    'Vertial Line #2
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0330" + "l" + "0004" + "0120" + vbCr

    'sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0245" + "l" + "0004" + "0380" + vbCr
    'Vertial Line #1                          strat    v-position              Height
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0760" + "l" + "0004" + "0950" + vbCr
        
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0980"
    sFormatTx = sFormatTx + "Supplier" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0570"
    sFormatTx = sFormatTx + "PEAK" + vbCr
        
   'CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0740"
    If Trim(Right(C_CUSTOM_CD, 10)) = "22206" Then
        sFormatTx = sFormatTx + "MICRON Memory Malasia Sdn Bhd" + vbCr
    Else
        sFormatTx = sFormatTx + "MICRON SEMICONDUCTOR ASIA" + vbCr
    End If
        
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0260" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0265" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & varOEMiPART_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0305" + "0740"
    sFormatTx = sFormatTx & varOEMiPART_NO & vbCr
             
    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0740"
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr
    
    'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0980"
    sFormatTx = sFormatTx + "PO NO#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & C_PO_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0740"
    sFormatTx = sFormatTx & C_PO_NO & vbCr
    
    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0320"
    sFormatTx = sFormatTx + "QTY:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0200"  'Code 128
    sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0200"
    sFormatTx = sFormatTx + varUnitQty + vbCr
    
    'PEAK DWG
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0620" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & I_DESC_NM + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0670" + "0740"
    sFormatTx = sFormatTx & I_DESC_NM & vbCr

    'LOT NO (P+YMMDDXX) : XX = 출하횟수 -> 20170918 실제 LOT 표시 ihson
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0750" + "0980"
    sFormatTx = sFormatTx + "LOT NO:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0755" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & Replace(varLot_no, "-", "") & vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0800" + "0740"
    sFormatTx = sFormatTx + Replace(varLot_no, "-", "") + vbCr
    
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0980"
    sFormatTx = sFormatTx + "C/NO:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0860"
    sFormatTx = sFormatTx & varCartonNo & vbCr
    
    'COO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0740"
    sFormatTx = sFormatTx + "COO: " + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0600"
    sFormatTx = sFormatTx & "MADE IN TAIWAN" + vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr

End Sub
Private Sub T4_Micron_SGP_BoxLabel_NEW_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2021. 06. 11
    ' Customer : Micron SINGAPORE
    ' Author : FRC IT Jim Chen
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim sText3                  As String
    Dim varLot_no               As String
    Dim varMoldSeq              As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varDWG_no               As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim strLOT_ID               As String
    
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    varMoldSeq = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2)
    
    varMFGDT = Format(GF_CHECK_LotDate("TR", varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, varMFGDT, Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    
    varMFGDT = Replace(varMFGDT, "-", "")
    varEXPDT = Format(varEXPDT, "yyyymmdd")
     
    'varEXPDT = DateAdd("m", 36, varMFGDT)
    'varEXPDT = DateAdd("m", 36, 20210729)
    
    'varMFGDT = Left(varMFGDT, 6) & "01" '20170918 생산월 01로 고정 표시 ihson , CPL hilee 2020.11.24 실제 생산일자를 표시함 (모든 SPIL SITE 공통적용)
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
      ' 해당월의 출하 횟수를 확인
    ' Display Shipping Date for the C_LOT_NO for Delivery Order for the Lot No
    
    sql$ = "select ipart_no , Prev_No,   'P' || '" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "' || trim(to_char(ROWNUM, '00')) as LOT_ID"
    sql$ = sql$ & " from ( select distinct ipart_no, Prev_No from MKT_OUT_11 a where a.comp_cd = 'SG' and a.bsns_cd = '11' and a.outorder_no ='" & Trim(C_LOT_ID) & "'"
    sql$ = sql$ & ")"
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
            
    strLOT_ID = ""
    If SSDATA.RecordCount > 0 Then
        Do Until SSDATA.EOF
            If Trim(SSDATA!ipart_no) = Trim( I_PART_NO) And Trim(SSDATA!Prev_No) = Trim(I_REV_NO) Then strLOT_ID = Trim(SSDATA!LOT_ID)
            SSDATA.MoveNext
        Loop
    End If
    SSDATA.Close
    If Trim(strLOT_ID) = "" Then strLOT_ID = "P" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "01"

    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no,  nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as dwg_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    If SSDATA.RecordCount > 0 Then
        varDWG_no = SSDATA!dwg_no
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
 
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
 
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "42"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' 외부 Box 그리기           -- reserved --   -Row-    -Col-   -box- -수평길이+수직길이- -- line 폭 --
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0030" + "b" + "0980" + "0950" + "0004" + "0004" + vbCr

    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    
    '2D DataMatrix Request by FRC CPL Liam 2021-09-07
    If GLO_COMP_CD = "FU" Then
        'sFormatTx = sFormatTx & "3W1D33000080003002HM," & "P" & C_ITEM_NO & "|1P" & varOEMiPART_NO & "|21P" & "FRC" & "|1T" & Replace(varLot_no, "-", "") & "|2D" & varEXPDT & "|S" & varCartonNo & "|7Q" & varUnitQty & "4LTW" & Chr$(13)
        sFormatTx = sFormatTx & "3W1c55" + "000" + "0600" + "0280" + "200" + "0" + "000" + "000" & "P" & C_ITEM_NO & "|1P" & varOEMiPART_NO & "|21P" & "FRC" & "|1T" & Replace(varLot_no, "-", "") & "|2D" & varEXPDT & "|S" & varCartonNo & "|7Q" & varUnitQty & "|4LTW" & Chr$(13)
    End If

    '2D BARCODE(QR CODE)    'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로, b,c = Module Size
    'sFormatTx = sFormatTx & "3W1D33000030502302HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)  'FRC '2D QR Code Request by FRC CPL Liam 2020-06-05
    'sFormatTx = sFormatTx & "3W1D55000063002202HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)

    ' 가로줄 'Vertial Line
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0004" + vbCr
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0200" + "0030" + "l" + "0980" + "0004" + vbCr
    If GLO_COMP_CD = "FU" Then
        sFormatTx = sFormatTx + "1X11" + "000" + "0300" + "0030" + "l" + "0980" + "0004" + vbCr
    Else
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr
    End If
    sFormatTx = sFormatTx + "1X11" + "000" + "0400" + "0030" + "l" + "0980" + "0004" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr    test
    sFormatTx = sFormatTx + "1X11" + "000" + "0500" + "0330" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0600" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0700" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0810" + "0030" + "l" + "0980" + "0004" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0880" + "0030" + "l" + "0980" + "0004" + vbCr
    
    '세로줄

    'If GLO_COMP_CD = "FU" Then
        'If Trim(Right(C_CUSTOM_CD, 10)) = "21311" Then
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0365" + vbCr
        'Else
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0240" + vbCr
        'End If
    'End If
    'Vertial Line #2
    sFormatTx = sFormatTx + "1X11" + "000" + "0400" + "0330" + "l" + "0004" + "0210" + vbCr

    'sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0245" + "l" + "0004" + "0380" + vbCr
    'Vertial Line #1                          strat    v-position              Height
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0760" + "l" + "0004" + "0950" + vbCr
        
   'Title & CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0080" + "0980"
    sFormatTx = sFormatTx + "PEAK" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0080" + "0740"
    sFormatTx = sFormatTx + "MICRON SEMICONDUCTOR ASIA" + vbCr

    'CUSTOMER PO#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "PO NO#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0145" + "0740"
    sFormatTx = sFormatTx & "K" & C_PO_NO & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0190" + "0740"  'Code 128
    sText3 = "K" & C_PO_NO
    sFormatTx = sFormatTx & "B" & sText3 + vbCr
        
   'PKG Slip
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0270" + "0980"
    sFormatTx = sFormatTx + "PKG Slip" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0245" + "0740"
    sFormatTx = sFormatTx + "11K" & Replace(varLot_no, "-", "") & "-" & Format(SHIP_DATE, "yyyymmdd") & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0290" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & "11K" & Replace(varLot_no, "-", "") & "-" & Format(SHIP_DATE, "yyyymmdd") + vbCr
    
    'Manufature
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0370" + "0980"
    sFormatTx = sFormatTx + "Manufature:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0345" + "0740"
    sFormatTx = sFormatTx + "21P" & "FRC" & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0390" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & "21P" & "FRC" & vbCr
    
    'Origin
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0470" + "0980"
    sFormatTx = sFormatTx + "Origin:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0445" + "0740"
    sFormatTx = sFormatTx + "4L" & "TW" & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0490" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & "4L" & "TW" & vbCr
    
    'P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0570" + "0980"
    sFormatTx = sFormatTx + "P/N:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0740"
    sFormatTx = sFormatTx + "P" & C_ITEM_NO & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0590" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & "P" & C_ITEM_NO & vbCr
    
    'LOT NO#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0670" + "0980"
    sFormatTx = sFormatTx + "LOT NO#:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0645" + "0740"
    sFormatTx = sFormatTx + "Z" & Replace(varLot_no, "-", "") & "+" & varEXPDT & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0690" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & "Z" & Replace(varLot_no, "-", "") & "+" & varEXPDT + vbCr
    
    'OTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0770" + "0980"
    sFormatTx = sFormatTx + "OTY:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0745" + "0740"
    sFormatTx = sFormatTx + "7Q" & varUnitQty & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0790" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & "7Q" & varUnitQty & vbCr
    
    
    'S/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0870" + "0980"
    sFormatTx = sFormatTx + "S/N#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0850" + "0740"
    sFormatTx = sFormatTx + "S" & Replace(varLot_no, "-", "") & "-" & varCartonNo & vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "050" + "0895" + "0740"  'Code 128
    sText3 = "S" & Replace(varLot_no, "-", "") & "-" & varCartonNo
    sFormatTx = sFormatTx & "B" & sText3 + vbCr
    
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0940" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0940" + "0740"
    sFormatTx = sFormatTx + varOEMiPART_NO + vbCr
    
    
    'OLD--------
    'PEAK P/N
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0260" + "0980"
    'sFormatTx = sFormatTx + "PEAK P/N:" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0265" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & varOEMiPART_NO + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0305" + "0740"
    'sFormatTx = sFormatTx & varOEMiPART_NO & vbCr
             
    'CUST'R P/N
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    'sFormatTx = sFormatTx + "CUST'R P/N:" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0740"
    'sFormatTx = sFormatTx + C_ITEM_NO + vbCr
    
    'CUST'R PO
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0980"
    'sFormatTx = sFormatTx + "PO NO#" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & C_PO_NO + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0740"
    'sFormatTx = sFormatTx & C_PO_NO & vbCr
    
    'QTY
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0320"
    'sFormatTx = sFormatTx + "QTY:" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0200"  'Code 128
    'sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0200"
    'sFormatTx = sFormatTx + varUnitQty + vbCr
    
    'PEAK DWG
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0620" + "0980"
    'sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & I_DESC_NM + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0670" + "0740"
    'sFormatTx = sFormatTx & I_DESC_NM & vbCr

    'LOT NO (P+YMMDDXX) : XX = 출하횟수 -> 20170918 실제 LOT 표시 ihson
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0750" + "0980"
    'sFormatTx = sFormatTx + "LOT NO:" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0755" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & Replace(varLot_no, "-", "") & vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0800" + "0740"
    'sFormatTx = sFormatTx + Replace(varLot_no, "-", "") + vbCr
    
    'C/NO
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0980"
    'sFormatTx = sFormatTx + "C/NO:" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0860"
    'sFormatTx = sFormatTx & varCartonNo & vbCr
    
    'COO
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0740"
    'sFormatTx = sFormatTx + "COO: " + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0600"
    'sFormatTx = sFormatTx & "MADE IN TAIWAN" + vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr

End Sub
Private Sub T4_Walton_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2021. 06. 18
    '
    '
    ' Auhtor : FRC Jim Chen
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim charRowAdd(5 To 48)     As Integer
    Dim varDWG_no               As String
    Dim varCUSTDWG_no           As String
    Dim varPrdITEM_no           As String
    Dim varMFGDT                As String
    
    varLot_no = C_LOT_NO

    'MFG Date
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@-@@-@@") 'Lot 생성일자를 가져옴
    
    
    'varEXPDT = Replace(DateAdd("yyyy", 5, varMFGDT), "-", "")
    varMFGDT = Replace(varMFGDT, "-", "")
  
    

    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
    
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    '===========================================================================
    ' SMP 정보 확인, Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no, nvl(a.smp_rev_no,' ') as smp_rev_no, nvl(a.cust_dwg_no,' ') as cust_dwg_no," & vbCrLf
    sql$ = sql$ & " (select distinct nvl(item_no, ' ') item_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as prd_item_no, " & vbCrLf
    sql$ = sql$ & " (select distinct nvl(dwg_no, ' ') dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO) as dwg_no" & vbCrLf
    sql$ = sql$ & " from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    varCUSTDWG_no = ""
    varPrdITEM_no = ""
    If SSDATA.RecordCount > 0 Then
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
        varDWG_no = SSDATA!dwg_no
        varPrdITEM_no = SSDATA!prd_item_no
        varCUSTDWG_no = SSDATA!cust_dwg_no
    End If
    SSDATA.Close

    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
    
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0007" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0850" + "0030" + "l" + "0980" + "0007" + vbCr
    
    ' Barcode 표시              -출력각- -type-  --두께--   -높이-   -row-    -col-          * 두께 : 굵은 두께 + 얇은 두께
    
    
    
    '2D BARCODE(QR CODE)    'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로, b,c = Module Size
    sFormatTx = sFormatTx & "3W1D55000065002302HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varMFGDT & Chr$(13)  'FRC '2D QR Code Request by FRC CPL Liam 2020-06-05
   
    'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "006" + "0080" + "0850"
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" & vbCr
        
    'CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0170" + "0730"
    sFormatTx = sFormatTx + ": " + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
    
    'CUSTOMER P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0270" + "0980"
    sFormatTx = sFormatTx + "Part No" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0270" + "0730"
    sFormatTx = sFormatTx + ": " + C_ITEM_NO + vbCr
    
    'Product DWG#
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0370" + "0980"
    sFormatTx = sFormatTx + "Item" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0370" + "0730"
    sFormatTx = sFormatTx + ": " + sText1 + vbCr
    
    'SHIPPG LOT NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0470" + "0980"
    sFormatTx = sFormatTx + "LOT NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0470" + "0730"
    sFormatTx = sFormatTx + ": " + C_LOT_NO + vbCr
    
    'MFG Date
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0570" + "0980"
    sFormatTx = sFormatTx + "MFG Date" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0570" + "0730"
    sFormatTx = sFormatTx + ": " + varMFGDT + vbCr
        
    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0670" + "0980"
    sFormatTx = sFormatTx + "Quantity" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0670" + "0730"
    sFormatTx = sFormatTx + ": " + varUnitQty + vbCr
      
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0770" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0770" + "0730"
    sFormatTx = sFormatTx + ": " + varOEMiPART_NO + vbCr '  I_PART_NO + vbCr

    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0920" + "0980"
    sFormatTx = sFormatTx + "C/NO" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "004" + "0920" + "0730"
    sFormatTx = sFormatTx + ": " & varCartonNo & vbCr
    sFormatTx = sFormatTx + "E" + vbCr
End Sub
Private Sub T4_ASE_SGP_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2021. 06. 11
    ' Customer : Micron SINGAPORE
    ' Author : FRC IT Jim Chen
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varMoldSeq              As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varDWG_no               As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim strLOT_ID               As String
    
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    varMoldSeq = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2)
    
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@@@@@") 'Lot 생성일자를 가져옴
    'varMFGDT = Left(varMFGDT, 6) & "01" '20170918 생산월 01로 고정 표시 ihson , CPL hilee 2020.11.24 실제 생산일자를 표시함 (모든 SPIL SITE 공통적용)
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
      ' 해당월의 출하 횟수를 확인
    ' Display Shipping Date for the C_LOT_NO for Delivery Order for the Lot No
    
    sql$ = "select ipart_no , Prev_No,   'P' || '" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "' || trim(to_char(ROWNUM, '00')) as LOT_ID"
    sql$ = sql$ & " from ( select distinct ipart_no, Prev_No from MKT_OUT_11 a where a.comp_cd = 'SG' and a.bsns_cd = '11' and a.outorder_no ='" & Trim(C_LOT_ID) & "'"
    sql$ = sql$ & ")"
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
            
    strLOT_ID = ""
    If SSDATA.RecordCount > 0 Then
        Do Until SSDATA.EOF
            If Trim(SSDATA!ipart_no) = Trim( I_PART_NO) And Trim(SSDATA!Prev_No) = Trim(I_REV_NO) Then strLOT_ID = Trim(SSDATA!LOT_ID)
            SSDATA.MoveNext
        Loop
    End If
    SSDATA.Close
    If Trim(strLOT_ID) = "" Then strLOT_ID = "P" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "01"

    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no,  nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as dwg_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    If SSDATA.RecordCount > 0 Then
        varDWG_no = SSDATA!dwg_no
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
 
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
 
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' 외부 Box 그리기           -- reserved --   -Row-    -Col-   -box- -수평길이+수직길이- -- line 폭 --
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0030" + "b" + "0980" + "0950" + "0004" + "0004" + vbCr

    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    
    '2D QR Code Request by FRC CPL Liam 2020-06-05
    'If GLO_COMP_CD = "FU" Then
        'sFormatTx = sFormatTx & "3W1D66000042002902HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)
    'End If

    '2D BARCODE(QR CODE)    'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로, b,c = Module Size
    'sFormatTx = sFormatTx & "3W1D33000030502302HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)  'FRC '2D QR Code Request by FRC CPL Liam 2020-06-05
    'sFormatTx = sFormatTx & "3W1D55000063002202HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)

    ' 가로줄 'Vertial Line
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0004" + vbCr
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0030" + "l" + "0980" + "0004" + vbCr
    If GLO_COMP_CD = "FU" Then
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr
    Else
        sFormatTx = sFormatTx + "1X11" + "000" + "0310" + "0030" + "l" + "0980" + "0004" + vbCr
    End If
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0030" + "l" + "0980" + "0004" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr    test
    sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0670" + "0030" + "l" + "0980" + "0004" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0740" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0810" + "0030" + "l" + "0980" + "0004" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0880" + "0030" + "l" + "0980" + "0004" + vbCr
    
    '세로줄

    'If GLO_COMP_CD = "FU" Then
        'If Trim(Right(C_CUSTOM_CD, 10)) = "21311" Then
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0365" + vbCr
        'Else
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0240" + vbCr
        'End If
    'End If
    'Vertial Line #2
    sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0330" + "l" + "0004" + "0120" + vbCr

    'sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0245" + "l" + "0004" + "0380" + vbCr
    'Vertial Line #1                          strat    v-position              Height
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0760" + "l" + "0004" + "0950" + vbCr
        
   'Title
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0080" + "0980"
    sFormatTx = sFormatTx + "SUPPLIER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0740"
    sFormatTx = sFormatTx + "PEAK" + vbCr
   
   'CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0740"
    sFormatTx = sFormatTx + "ASE SINGAPORE PTE. LTD." + vbCr
        
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0260" + "0980"
    sFormatTx = sFormatTx + "PEAK P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0265" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & varOEMiPART_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0305" + "0740"
    sFormatTx = sFormatTx & varOEMiPART_NO & vbCr
             
    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0380" + "0980"
    sFormatTx = sFormatTx + "CUST'R P/N" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0385" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0420" + "0740"
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr
    
    'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0980"
    sFormatTx = sFormatTx + "PO NO#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & C_PO_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0740"
    sFormatTx = sFormatTx & C_PO_NO & vbCr
    
    'QTY
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0320"
    sFormatTx = sFormatTx + "QTY:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0200"  'Code 128
    sFormatTx = sFormatTx & "B" & varUnitQty + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0545" + "0200"
    sFormatTx = sFormatTx + varUnitQty + vbCr
    
    'PEAK DWG
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0620" + "0980"
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & I_DESC_NM + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0670" + "0740"
    sFormatTx = sFormatTx & I_DESC_NM & vbCr

    'LOT NO (P+YMMDDXX) : XX = 출하횟수 -> 20170918 실제 LOT 표시 ihson
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0750" + "0980"
    sFormatTx = sFormatTx + "LOT NO:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0755" + "0740"  'Code 128
    sFormatTx = sFormatTx & "B" & Replace(varLot_no, "-", "") & vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0800" + "0740"
    sFormatTx = sFormatTx + Replace(varLot_no, "-", "") + vbCr
    
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0980"
    sFormatTx = sFormatTx + "C/NO:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0860"
    sFormatTx = sFormatTx & varCartonNo & vbCr
    
    'COO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0740"
    sFormatTx = sFormatTx + "COO: " + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0885" + "0600"
    sFormatTx = sFormatTx & "MADE IN TAIWAN" + vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr

End Sub
Private Sub T4_SIGUARD_BoxLabel_Print(varCartonNo As Integer, varTotalCartonNo As Integer)
    '======================================================================
    ' 2022. 07. 18
    ' Customer : SIGURD 20908?21547?21987
    ' Author : FRC IT Jim Chen
    '======================================================================
    Dim sql                     As String
    Dim sText1                  As String
    Dim sText2                  As String
    Dim varLot_no               As String
    Dim varMoldSeq              As String
    Dim varUnitQty              As String
    Dim varQApass               As String
    Dim subSpace                As String
    Dim varMFGDT                As String
    Dim varEXPDT                As String
    Dim varDWG_no               As String
    Dim varOEMiPART_NO          As String
    Dim varOEMiRev_NO           As String
    Dim strLOT_ID               As String
    
    Dim yyyy                    As String
    Dim MM                      As String
    Dim dd                      As String
    
    Dim charRowAdd(5 To 48)     As Integer
    
    varLot_no = C_LOT_NO
    varMoldSeq = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2)
    varMFGDT = Format(GF_CHECK_LotDate(Trim(D_PROD_KB), varLot_no), "@@@@/@@/@@") 'Lot 생성일자를 가져옴
    
    
    'varMFGDT = Left(varMFGDT, 6) & "01" '20170918 생산월 01로 고정 표시 ihson , CPL hilee 2020.11.24 실제 생산일자를 표시함 (모든 SPIL SITE 공통적용)
    varEXPDT = GF_GET_EXPDATE(GLO_COMP_CD, GLO_BSNS_CD, Trim(SHIP_DATE), Trim(Right(C_CUSTOM_CD, 5)), Trim( I_PART_NO), Trim(I_REV_NO), IIf(C_ITEM_NO = "", " ", C_ITEM_NO), Trim(Right(D_PROD_KB, 10)), , , Format(SHIP_DATE, "yyyy-mm-dd"))
    varEXPDT = Replace(varEXPDT, "-", "/")  'JJ202207
     
    'Check last C/T Qty
    If varCartonNo = varTotalCartonNo Then
        varUnitQty = CStr(Val(I_TOTAL_QT) Mod Val(I_BOXUNIT_NB))
        If varUnitQty = "0" Then varUnitQty = I_BOXUNIT_NB
    Else
        varUnitQty = I_BOXUNIT_NB
    End If
        
    If Val(I_START_CT_NO) = Val(I_END_CT_NO) Then
        If Val(Result_QT) > 0 Then I_BOXUNIT_NB = Result_QT
    End If
    
      ' 해당월의 출하 횟수를 확인
    ' Display Shipping Date for the C_LOT_NO for Delivery Order for the Lot No
    
    sql$ = "select ipart_no , Prev_No,   'P' || '" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "' || trim(to_char(ROWNUM, '00')) as LOT_ID"
    sql$ = sql$ & " from ( select distinct ipart_no, Prev_No from MKT_OUT_11 a where a.comp_cd = 'SG' and a.bsns_cd = '11' and a.outorder_no ='" & Trim(C_LOT_ID) & "'"
    sql$ = sql$ & ")"
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
            
    strLOT_ID = ""
    If SSDATA.RecordCount > 0 Then
        Do Until SSDATA.EOF
            If Trim(SSDATA!ipart_no) = Trim( I_PART_NO) And Trim(SSDATA!Prev_No) = Trim(I_REV_NO) Then strLOT_ID = Trim(SSDATA!LOT_ID)
            SSDATA.MoveNext
        Loop
    End If
    SSDATA.Close
    If Trim(strLOT_ID) = "" Then strLOT_ID = "P" & Right(Format(SHIP_DATE, "yyyymmdd"), 5) & "01"

    '===========================================================================
    ' Drawing No 확인
    '===========================================================================
    sql$ = "select nvl(a.smp_part_no,' ') as smp_part_no,  nvl(a.smp_rev_no,' ') as smp_rev_no, " & vbCrLf
    sql$ = sql$ & " nvl((select distinct dwg_no from PRD_MAST_1_VIEW f  where f.comp_cd = a.comp_cd and f.bsns_cd = a.bsns_cd and a.Prev_No=f.Prev_No and a.iPART_NO=f.iPART_NO ),' ') as dwg_no from MKT_CUST_2 a" & vbCrLf
    sql$ = sql$ & " where a.custom_cd ='" & Trim(Right(C_CUSTOM_CD, 10)) & "'" & vbCrLf
    sql$ = sql$ & " and a.Prev_No='" & I_REV_NO & "' and a.iPART_NO='" &  I_PART_NO & "'"
    sql$ = sql$ & " and comp_cd = '" & GLO_COMP_CD & "' " & vbCrLf
    sql$ = sql$ & " and bsns_cd = '" & GLO_BSNS_CD & "' " & vbCrLf
        
    Set SSDATA = Nothing
    SSDATA.Open sql$, DB_User
    
    varOEMiPART_NO = ""
    varOEMiRev_NO = ""
    varDWG_no = ""
    If SSDATA.RecordCount > 0 Then
        varDWG_no = SSDATA!dwg_no
        varOEMiPART_NO = SSDATA!smp_part_no
        varOEMiRev_NO = SSDATA!smp_rev_no
    End If
    SSDATA.Close
 
    If InStr(I_DESC_NM, vbCrLf) > 0 Then
        sText1 = Trim(Left(I_DESC_NM, InStr(I_DESC_NM, vbCrLf) - 1))
        sText2 = Trim(Right(I_DESC_NM, Len(I_DESC_NM) - Len(sText1) - 2))
        
        If InStr(sText2, "BAKE") > 0 Then
            sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        Else
            sText2 = sText2 & " " & Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
        End If
    Else
        sText1 = I_DESC_NM
        sText2 = Mid(varMoldNo, InStr(varMoldNo, "-") + 1, 2) & varOEMiRev_NO
    End If
 
'============Angle
    Const Angle0 = "1", Angle90 = "2", Angle180 = "3", Angle270 = "4"
    
'============Horizontal expansion rate
    Const Hori1 = "1", Hori2 = "2", Hori3 = "3", Hori4 = "4"
    Const Hori5 = "5", Hori6 = "6", Hori7 = "7", Hori8 = "8"
    Const Hori9 = "9", Hori10 = "A", Hori11 = "B", Hori12 = "C"
    Const Hori13 = "D", Hori14 = "E", Hori15 = "F", Hori16 = "G"
    Const Hori17 = "H", Hori18 = "I", Hori19 = "J", Hori20 = "K"
    Const Hori21 = "L", Hori22 = "M", Hori23 = "N", Hori24 = "O"

'============Vertical expansion rate
    Const Verti1 = "1", Verti2 = "2", Verti3 = "3", Verti4 = "4"
    Const Verti5 = "5", Verti6 = "6", Verti7 = "7", Verti8 = "8"
    Const Verti9 = "9", Verti10 = "A", Verti11 = "B", Verti12 = "C"
    Const Verti13 = "D", Verti14 = "E", Verti15 = "F", Verti16 = "G"
    Const Verti17 = "H", Verti18 = "I", Verti19 = "J", Verti20 = "K"
    Const Verti21 = "L", Verti22 = "M", Verti23 = "N", Verti24 = "O"

'============Font 9 일때 Size
    Const Font0 = "000"
    Const Font6 = "001", Font8 = "002", Font10 = "003", Font12 = "004"
    Const Font14 = "005", Font18 = "006", Font24 = "007", Font30 = "008", Font36 = "009", Font48 = "010"

    Const cFont6 = "101", cFont8 = "102", cFont10 = "103", cFont12 = "104"
    Const cFont14 = "105", cFont18 = "106", cFont24 = "107", cFont30 = "108", cFont36 = "109", cFont48 = "110"

'============Row Add
    charRowAdd(5) = 10:    charRowAdd(6) = 18:    charRowAdd(8) = 20:    charRowAdd(10) = 23
    charRowAdd(12) = 25:   charRowAdd(14) = 28:   charRowAdd(18) = 30:   charRowAdd(24) = 35
    charRowAdd(30) = 43:   charRowAdd(36) = 48:   charRowAdd(48) = 78:
    
    Dim varRate     As String
    varRate = "52"

    sFormatTx = Chr(2) + "m" + vbCr
    sFormatTx = sFormatTx + Chr(2) + "L" + vbCr
    sFormatTx = sFormatTx + "D11" + vbCr
    
    ' 외부 Box 그리기           -- reserved --   -Row-    -Col-   -box- -수평길이+수직길이- -- line 폭 --
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0030" + "b" + "0980" + "0950" + "0004" + "0004" + vbCr

    
    ' Barcode 줄 긋기    -- reserved --   -Row-    -Col-   -line-  -길이 + 두께-
    
    '2D QR Code Request by FRC CPL Liam 2020-06-05
    'If GLO_COMP_CD = "FU" Then
        'sFormatTx = sFormatTx & "3W1D66000042002902HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)
    'End If

    '2D BARCODE(QR CODE)    'aW a=3은 180도회전  ffff 높을수록 아래 gggg 높을수록 왼쪽으로, b,c = Module Size
    'sFormatTx = sFormatTx & "3W1D33000030502302HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)  'FRC '2D QR Code Request by FRC CPL Liam 2020-06-05
    'sFormatTx = sFormatTx & "3W1D55000063002202HM, " & C_ITEM_NO & ";" & Trim(Replace(varLot_no, "-", "")) & ";" & varUnitQty & ";" & varMFGDT & Chr$(13)

    ' 가로줄 'Vertial Line
    sFormatTx = sFormatTx + "1X11" + "000" + "0100" + "0030" + "l" + "0980" + "0004" + vbCr '1
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0030" + "l" + "0980" + "0004" + vbCr '2
    If GLO_COMP_CD = "FU" Then
        sFormatTx = sFormatTx + "1X11" + "000" + "0260" + "0030" + "l" + "0980" + "0004" + vbCr '3  300-260
    Else
        sFormatTx = sFormatTx + "1X11" + "000" + "0260" + "0030" + "l" + "0980" + "0004" + vbCr
    End If
    sFormatTx = sFormatTx + "1X11" + "000" + "0340" + "0030" + "l" + "0980" + "0004" + vbCr '4 0430-390 400-340
    'sFormatTx = sFormatTx + "1X11" + "000" + "0550" + "0030" + "l" + "0980" + "0004" + vbCr    test
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0410" + "0030" + "l" + "0980" + "0004" + vbCr '5 550-510 510-470  470-410
    sFormatTx = sFormatTx + "1X11" + "000" + "0530" + "0030" + "l" + "0980" + "0004" + vbCr '6 670-630  630-590   590-530
    'sFormatTx = sFormatTx + "1X11" + "000" + "0740" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0660" + "0030" + "l" + "0980" + "0004" + vbCr ' 7 810-770 760-720  720-660
    
    sFormatTx = sFormatTx + "1X11" + "000" + "0780" + "0030" + "l" + "0980" + "0004" + vbCr ' 8  '810-780
    
    'sFormatTx = sFormatTx + "1X11" + "000" + "0880" + "0030" + "l" + "0980" + "0004" + vbCr
    sFormatTx = sFormatTx + "1X11" + "000" + "0900" + "0030" + "l" + "0980" + "0004" + vbCr ' 9
    
    '세로줄

    'If GLO_COMP_CD = "FU" Then
        'If Trim(Right(C_CUSTOM_CD, 10)) = "21311" Then
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0365" + vbCr
        'Else
            'sFormatTx = sFormatTx + "1X11" + "000" + "0190" + "0330" + "l" + "0004" + "0240" + vbCr
        'End If
    'End If
 
       'Vertial Line #3
 '  sFormatTx = sFormatTx + "1X11" + "000" + "0760" + "0300" + "l" + "0004" + "0135" + vbCr
       'Vertial Line #3
  '  sFormatTx = sFormatTx + "1X11" + "000" + "0760" + "0480" + "l" + "0004" + "0135" + vbCr
    '                                         Y                                LONG
        
    'Vertial Line #2
    'sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0330" + "l" + "0004" + "0120" + vbCr
    'sFormatTx = sFormatTx + "1X11" + "000" + "0430" + "0245" + "l" + "0004" + "0380" + vbCr
    'Vertial Line #1                          strat    v-position              Height
    sFormatTx = sFormatTx + "1X11" + "000" + "0000" + "0760" + "l" + "0004" + "0950" + vbCr
        
        
   '1 Title
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0980"
    sFormatTx = sFormatTx + "Supplier" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "005" + "0080" + "0670" '570-670
    sFormatTx = sFormatTx + "PEAK SHIPPING LABEL" + vbCr     '20220718 JJ  PEAK ->  PEAK SHIPPING LABEL
   '2 CUSTOMER
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0980"
    sFormatTx = sFormatTx + "CUSTOMER:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0170" + "0740"
    If Trim(Right(C_CUSTOM_CD, 10)) = "22206" Then
        sFormatTx = sFormatTx + "MICRON Memory Malasia Sdn Bhd" + vbCr
    Else
        sFormatTx = sFormatTx + Trim(Left(C_CUSTOM_CD, 30)) + vbCr
    End If
         
    'PEAK P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0250" + "0980" '260-250
    sFormatTx = sFormatTx + "PEAK P/N:" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0265" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & varOEMiPART_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0250" + "0740"   'JJ"0305"->0260 260-250
    sFormatTx = sFormatTx & varOEMiPART_NO & vbCr
 
    'PEAK DWG
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0320" + "0980"  '380-320
    sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & I_DESC_NM + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0320" + "0740"  '---420  ->380(40) 380-320
    sFormatTx = sFormatTx & I_DESC_NM & vbCr
  
    'CUST'R PO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0400" + "0980" '500-460  460-400
    sFormatTx = sFormatTx + "CUST'R PO#:" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0505" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & C_PO_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0400" + "0740" '545-505'505-460  460-400
    sFormatTx = sFormatTx & C_PO_NO & vbCr
 
    'PEAK DWG
  '  sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0620" + "0980"
  '  sFormatTx = sFormatTx + "PEAK DWG#" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0625" + "0740"  'Code 128
    'sFormatTx = sFormatTx & "B" & I_DESC_NM + vbCr
  '  sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0670" + "0740"
  '  sFormatTx = sFormatTx & I_DESC_NM & vbCr
 
    'CUST'R P/N
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0500" + "0980" '620-580 580-500
    sFormatTx = sFormatTx + "CUST'R P/N#:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0485" + "0740"  'Code 128'625-585  585-545 545-485
    sFormatTx = sFormatTx & "B" & C_ITEM_NO + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0530" + "0740" '670-630'  630-590 590-530
    sFormatTx = sFormatTx + C_ITEM_NO + vbCr
 
    'LOT NO (P+YMMDDXX) : XX = 출하횟수 -> 20170918 실제 LOT 표시 ihson
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0610" + "0980" '750-710  710-670 670-610
    sFormatTx = sFormatTx + "LOT NO#" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "070" + "0605" + "0740"  'Code 128'755-705 705-665 665-605
    sFormatTx = sFormatTx & "B" & Replace(varLot_no, "-", "") & "-" & Format(SHIP_DATE, "mm/dd/yy") & vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0650" + "0740" '800-760'760-750 750-710  710-650
    'sFormatTx = sFormatTx + Replace(varLot_no, "-", "") + "-" + Format(SHIP_DATE, "mm/dd/yy") & vbCr
    sFormatTx = sFormatTx + Replace(varLot_no, "-", "") & vbCr  'jj202207
 
    'MFG Date
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0730" + "0980"   'NEED TEST'0320-0980  850-770 770-710 710-730
    sFormatTx = sFormatTx + "MFG DATE:" + vbCr
    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "060" + "0720" + "0740"  'Code 128  varMFGDT  070-060 845-765 765-705 705-715 715-720
    sFormatTx = sFormatTx & "B" & varMFGDT + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0770" + "0740" '--0200-0740 varMFGDT  885-805 805-745 745-765
    sFormatTx = sFormatTx + varMFGDT + vbCr
   
   
   
    'EXP Date
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0870" + "0980"   'NEED TEST'0620-0550  550-450 0450-0980 900-870
    sFormatTx = sFormatTx + "EXP DATE:" + vbCr
    'sFormatTx = sFormatTx + Angle180 + "e" + varRate + "060" + "0845" + "0740"  'Code 128  070-050 900-840
    'sFormatTx = sFormatTx & "B" & varEXPDT + vbCr
    'sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0900" + "0740" '920-880 880-900
    'sFormatTx = sFormatTx + varEXPDT + vbCr


    varEXPDT = DateAdd("yyyy", 3, varMFGDT)   'JJ202207
    varEXPDT = Replace(varEXPDT, "-", "/")    'JJ202207
  

    sFormatTx = sFormatTx + Angle180 + "e" + varRate + "060" + "0845" + "0740"  'Code 128  070-050 900-840
    sFormatTx = sFormatTx & "B" & varEXPDT + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0900" + "0740" '920-880 880-900
    sFormatTx = sFormatTx + varEXPDT + vbCr


 
    
    
'---- DOWN OK
 
    'C/NO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0950" + "0980" '885->955
    sFormatTx = sFormatTx + "C/NO:" + vbCr
    
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0950" + "0860" '885->955
    sFormatTx = sFormatTx & varCartonNo & vbCr
    
    'COO
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0950" + "0740" '885->955  JJ20220720
    sFormatTx = sFormatTx + "COO: " + vbCr
    sFormatTx = sFormatTx + Angle180 + "911" + "003" + "0950" + "0600" '885->955
    sFormatTx = sFormatTx & "MADE IN TAIWAN" + vbCr
    
    sFormatTx = sFormatTx + "E" + vbCr

End Sub
