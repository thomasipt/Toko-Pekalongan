VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form BL001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PEMBELIAN BARANG"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "BL001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "BL001.frx":000C
   ScaleHeight     =   6765
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdtutup 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      TabIndex        =   14
      ToolTipText     =   "Klik untuk keluar tanpa melakukan transaksi"
      Top             =   6255
      Width           =   10590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3150
      OleObjectBlob   =   "BL001.frx":034E
      Top             =   7425
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   4695
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   7770
      Width           =   1965
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1890
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "No faktur pembelian"
      Top             =   630
      Width           =   3345
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   5509
      TabIndex        =   15
      Text            =   "1,000,000.00"
      Top             =   45
      Width           =   5190
   End
   Begin VB.PictureBox CRPT 
      Height          =   480
      Left            =   2535
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   20
      Top             =   9540
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   1845
      TabIndex        =   16
      ToolTipText     =   "Klik untuk edit"
      Top             =   90
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   20774913
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   330
      Left            =   8610
      OleObjectBlob   =   "BL001.frx":0582
      TabIndex        =   18
      Top             =   5865
      Width           =   2085
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   330
      Left            =   5505
      OleObjectBlob   =   "BL001.frx":05F4
      TabIndex        =   19
      Top             =   5865
      Width           =   3210
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4635
      Left            =   105
      TabIndex        =   21
      Top             =   1170
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DAFTAR PEMBELIAN BARANG"
      TabPicture(0)   =   "BL001.frx":067C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text7"
      Tab(0).Control(1)=   "Combo1"
      Tab(0).Control(2)=   "Combo2"
      Tab(0).Control(3)=   "Text6"
      Tab(0).Control(4)=   "cmdBL003"
      Tab(0).Control(5)=   "Command4"
      Tab(0).Control(6)=   "SkinLabel6"
      Tab(0).Control(7)=   "SkinLabel8"
      Tab(0).Control(8)=   "SkinLabel9"
      Tab(0).Control(9)=   "SkinLabel7"
      Tab(0).Control(10)=   "grid"
      Tab(0).Control(11)=   "SkinLabel17"
      Tab(0).Control(12)=   "Frame1"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "PEMBAYARAN"
      TabPicture(1)   =   "BL001.frx":0698
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdsimpan"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67410
         TabIndex        =   5
         Text            =   "Text7"
         ToolTipText     =   "Nominal beli barang"
         Top             =   975
         Width           =   2310
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72518
         TabIndex        =   1
         Text            =   "Combo1"
         ToolTipText     =   "Info barang tekan F1"
         Top             =   450
         Width           =   2310
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -72518
         TabIndex        =   2
         Text            =   "Combo2"
         ToolTipText     =   "Info barang tekan F1"
         Top             =   975
         Width           =   2310
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -67410
         TabIndex        =   4
         Text            =   "Text6"
         ToolTipText     =   "Jumlah barang "
         Top             =   450
         Width           =   2310
      End
      Begin VB.CommandButton cmdBL003 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -65055
         TabIndex        =   6
         Top             =   975
         Width           =   510
      End
      Begin VB.CommandButton Command4 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   -70125
         TabIndex        =   3
         Top             =   450
         Width           =   375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Supplier"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   150
         TabIndex        =   29
         Top             =   2925
         Width           =   5115
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1665
            TabIndex        =   9
            Text            =   "Combo3"
            ToolTipText     =   "Klik untuk edit"
            Top             =   450
            Width           =   3255
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1665
            TabIndex        =   30
            Text            =   "3"
            Top             =   945
            Width           =   3255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   240
            Left            =   135
            OleObjectBlob   =   "BL001.frx":06B4
            TabIndex        =   31
            Top             =   540
            Width           =   1365
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   240
            Left            =   135
            OleObjectBlob   =   "BL001.frx":072A
            TabIndex        =   32
            Top             =   945
            Width           =   1365
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Syarat Bayar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   5415
         TabIndex        =   22
         Top             =   360
         Width           =   4980
         Begin VB.TextBox Text9 
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1710
            TabIndex        =   11
            Text            =   "9"
            Top             =   1980
            Width           =   3165
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
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
            Left            =   1710
            TabIndex        =   10
            Text            =   "5"
            Top             =   465
            Width           =   3165
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0E0FF&
            Height          =   900
            Left            =   1710
            MultiLine       =   -1  'True
            TabIndex        =   12
            Text            =   "BL001.frx":0790
            Top             =   2475
            Width           =   3165
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "BL001.frx":0793
            TabIndex        =   23
            Top             =   495
            Width           =   1410
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "BL001.frx":0805
            TabIndex        =   24
            Top             =   1035
            Width           =   1410
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "BL001.frx":087D
            TabIndex        =   25
            Top             =   1530
            Width           =   1410
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "BL001.frx":08F1
            TabIndex        =   26
            Top             =   2040
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   420
            Left            =   1710
            TabIndex        =   27
            ToolTipText     =   "Klik untuk edit"
            Top             =   945
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   741
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   49152
            CalendarTitleForeColor=   0
            CalendarTrailingForeColor=   16777088
            Format          =   20774913
            CurrentDate     =   39286
            MinDate         =   39083
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   420
            Left            =   1710
            TabIndex        =   28
            ToolTipText     =   "Klik untuk edit"
            Top             =   1440
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   741
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   49152
            CalendarTitleForeColor=   0
            CalendarTrailingForeColor=   16777088
            Format          =   20774913
            CurrentDate     =   39286
            MinDate         =   39083
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "BL001.frx":095B
            TabIndex        =   47
            Top             =   2805
            Width           =   1410
         End
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5415
         TabIndex        =   13
         ToolTipText     =   "Klik untuk simpan transaksi"
         Top             =   4050
         Width           =   4980
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   240
         Left            =   -74153
         OleObjectBlob   =   "BL001.frx":09CD
         TabIndex        =   38
         Top             =   540
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   240
         Left            =   -74153
         OleObjectBlob   =   "BL001.frx":0A41
         TabIndex        =   39
         Top             =   1065
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   240
         Left            =   -69090
         OleObjectBlob   =   "BL001.frx":0AB5
         TabIndex        =   40
         Top             =   1065
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   240
         Left            =   -69090
         OleObjectBlob   =   "BL001.frx":0B2B
         TabIndex        =   41
         Top             =   540
         Width           =   1560
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   2760
         Left            =   -74850
         TabIndex        =   42
         ToolTipText     =   "Daftar pembelian barang"
         Top             =   1500
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   4868
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   16777152
         GridColor       =   0
         TextStyle       =   3
         TextStyleFixed  =   3
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   240
         Left            =   -74850
         OleObjectBlob   =   "BL001.frx":0B95
         TabIndex        =   46
         Top             =   4275
         Width           =   5115
      End
      Begin VB.Frame Frame2 
         Caption         =   "Transaksi Pembayaran"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   150
         TabIndex        =   33
         Top             =   360
         Width           =   5115
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1080
            TabIndex        =   7
            Text            =   "Text14"
            Top             =   495
            Width           =   1860
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   360
            Left            =   1080
            TabIndex        =   8
            Text            =   "Text15"
            Top             =   915
            Width           =   1860
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3060
            TabIndex        =   34
            Text            =   "4"
            Top             =   1665
            Width           =   1950
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   240
            Left            =   90
            OleObjectBlob   =   "BL001.frx":0C29
            TabIndex        =   35
            Top             =   555
            Width           =   1560
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   240
            Left            =   90
            OleObjectBlob   =   "BL001.frx":0C95
            TabIndex        =   36
            Top             =   975
            Width           =   2460
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   240
            Left            =   1305
            OleObjectBlob   =   "BL001.frx":0D03
            TabIndex        =   37
            Top             =   1695
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "DAFTAR BARANG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   -70185
         TabIndex        =   43
         Top             =   405
         Width           =   5550
         Begin VB.CommandButton Command2 
            Caption         =   "BATAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1283
            TabIndex        =   44
            Top             =   2835
            Width           =   2985
         End
         Begin MSFlexGridLib.MSFlexGrid gridDF 
            Height          =   2490
            Left            =   45
            TabIndex        =   45
            Top             =   315
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   4392
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            BackColor       =   16777215
            BackColorFixed  =   65280
            BackColorBkg    =   16777152
            GridColor       =   0
            TextStyle       =   3
            TextStyleFixed  =   3
            Appearance      =   0
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   195
      Left            =   225
      OleObjectBlob   =   "BL001.frx":0D7A
      TabIndex        =   48
      Top             =   750
      Width           =   1365
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   195
      Left            =   225
      OleObjectBlob   =   "BL001.frx":0DDE
      TabIndex        =   49
      Top             =   210
      Width           =   1365
   End
End
Attribute VB_Name = "BL001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RNO, RDEl, RSL, RSLUser, RHarga, RCari, RCari2, RCari3, RCari4, RCari5, RCari6, RCari7, RCari8, RCari9, RCari10, RCari11, RCari12, RCari13, RCari14, RCari15, RCari16, RSave, RSave2, RSave3, RSave4, RSave5, RSave6, RSave7, RSave8, RSave9, RSave10, RSave11, RSave12, REdit As rdoResultset
Private SNO, SDel, SQL, SQLUser, SHarga, SCari, SCari2, SCari3, SCari4, SCari5, SCari6, SCari7, SCari8, SCari9, SCari10, SCari11, SCari12, SCari13, SCari14, SCari15, SCari16, SSave, SSave2, SSave3, SSave4, SSave5, SSave6, SSave7, SSave8, SSave9, SSave10, SSave11, SSave12, SEdit As String

Private RKTG, RSTN, RSPL, RPBR As rdoResultset
Private SKTG, SSTN, SSPL, SPBR As String

Private RKAS As rdoResultset
Private SKAS As String

Private SqlNo As String

Private RBahan As rdoResultset
Private SBahan As String


Private Sub cmdBL003_Click()
Dim Tanya
If Combo1 = "" Or Combo2 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbSystemModal, "KONFIRMASI"
    Combo1.SetFocus
Exit Sub
Else
    Tanya = MsgBox("TAMBAH DATA PEMBELIAN", vbSystemModal, "KONFIRMASI")
    If Tanya = vbOK Then
        Text8 = Format(CCur(Text8) + (CCur(Text6) * CCur(Text7)), "##,###.00")
        Call SimpanBL001
        Call SiapkanGrid
        Call IsiGrid
        Call Auto
        Combo1 = ""
        Combo2 = ""
        Text6 = ""
        Text7 = ""
        Combo1.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub SimpanBL001()

SCari5 = "Select * From B003 where KodeBR = '" + Trim(Combo1) + "'"
Set RCari5 = RDCO.OpenResultset(SCari5, rdOpenDynamic, rdConcurRowVer)
    INDUK = RCari5("KodeInd")
        
SSave = "Select * From BL001"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("No_Trans") = SkinLabel3
    RSave("No_Fak") = Trim(Text1)
    RSave("No_Urut") = Text2
    RSave("Kode_Ind") = INDUK
    RSave("Kode_BR") = Combo1
    RSave("Nama_BR") = Combo2
    RSave("Jml_Beli") = CCur(Text6)
    RSave("Harga_PCS") = CCur(Text7)
    RSave("Jml_Harga") = CCur(Text7) * CCur(Text6)
    RSave("UserCode") = Operator
    RSave("Tanggal") = Date
RSave.Update
RSave.Close
Set RSave = Nothing

RCari5.Close
Set RCari5 = Nothing

End Sub

Private Sub cmdsimpan_Click()
Dim Tanya

If Text1 = "" Then
    MsgBox "NOMOR PEMBELIAN MASIH KOSONG", vbSystemModal, "KONFIRMASI"
    Text1.SetFocus
    Exit Sub
End If

If Text8 = 0 Then
    MsgBox "TIDAK ADA TRANSAKSI PEMBELIAN", vbSystemModal, "KONFIRMASI"
    Combo1.SetFocus
    Frame1.Visible = False
    SSTab1.Tab = 0
    Exit Sub
End If

'If Text4 = 0 Then
'    MsgBox "TRANSAKSI PEMBAYARAN MASIH KOSONG", vbSystemModal, "KONFIRMASI"
'    Text14.SetFocus
'    Exit Sub
'End If

If Text3 = "" Then
    MsgBox "DATA SUPPLIER MASIH KOSONG", vbSystemModal, "KONFIRMASI"
    Combo3.SetFocus
    Exit Sub
End If

If Frame4.Visible = True Then
    If Text5 = "" Or Text9 = "" Or Text10 = "" Then
        MsgBox "SYARAT BAYAR MASIH KOSONG", vbSystemModal, "KONFIRMASI"
        Text5.SetFocus
        Exit Sub
    End If
End If

Tanya = MsgBox("ANDA YAKIN MELAKUKAN TRANSAKSI PEMBELIAN", vbSystemModal, "KONFIRMASI")
If Tanya = vbOK Then
    Call HISBahan
    If Frame4.Visible = True Then
        Call HISHutang
    End If
    Call KOSONG
    grid.Clear
    Call DelBL001
    SSTab1.Tab = 0
    Text1.SetFocus
End If
Unload Me
BL001.Show 1
End Sub

Private Sub HISHutang()
SSave = "Select * From H001"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NO_BUKTI") = SkinLabel3
    RSave("NO_HUTANG") = Trim(Text5)
    RSave("NO_BG") = Trim(Text9)
    RSave("NO_SPL") = Combo3
    RSave("NAMA_SPL") = Trim(Text3)
    RSave("KETERANGAN") = Trim(Text10)
    RSave("PLAFON") = CCur(Text15)
    RSave("TGL_MULAI") = DTPicker2
    RSave("TGL_JATUH") = DTPicker3
    RSave("TGL_BAYAR") = DTPicker2
    RSave("TANGGAL") = DTPicker1
    RSave("USER_CODE") = Operator
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub HISBahan()
SCari12 = "Select * From BL001 where no_Trans= '" + Trim(SkinLabel3) + "'"
Set RCari12 = RDCO.OpenResultset(SCari12, rdOpenKeyset, rdConcurRowVer)
RCari12.MoveFirst
Do While Not RCari12.EOF
    KODEJNS = RCari12("Kode_BR")

    SCari13 = "Select * From B003 where KodeBR = '" + Trim(KODEJNS) + "'"
    Set RCari13 = RDCO.OpenResultset(SCari13, rdOpenKeyset, rdConcurRowVer)
        JMLDBT = CCur(RCari13("JD")) + CCur(RCari12("JML_BELI"))
        JMLAKHIR = CCur(RCari13("JAkhir")) + CCur(RCari12("JML_BELI"))
        MUTASIDBT = CCur(RCari13("mutasid")) + CCur(RCari12("JML_HARGA"))
        SALDOAKHIR = CCur(RCari13("saldo")) + CCur(RCari12("JML_HARGA"))
        
        SCari14 = "Select * From B004"
        Set RCari14 = RDCO.OpenResultset(SCari14, rdOpenKeyset, rdConcurRowVer)
            HJUAL = CCur(RCari12("HARGA_PCS")) * CCur(RCari13("PERSEN") / 100) + CCur(RCari12("HARGA_PCS"))
            
            'INFO HARGA JUAL PCS'
            RCari14.AddNew
            RCari14("NO_TRANS") = Trim(SkinLabel3)
            RCari14("TGL_BELI") = DTPicker1
            RCari14("KODE_JNS") = KODEJNS
            RCari14("NAMA_JNS") = RCari12("NAMA_BR")
            RCari14("JML_SATUAN") = RCari12("JML_BELI")
            RCari14("HBELI_PCS") = RCari12("HARGA_PCS")
            RCari14("HARGA_BELI") = RCari12("JML_HARGA")
            RCari14("JML_SALDO") = RCari12("JML_BELI")
            RCari14("NOM_SALDO") = RCari12("JML_HARGA")
            RCari14("HJUAL_PCS") = CCur(HJUAL)
        
        'UPDATE MUTASI PEMBELIAN TABEL MASTER BAHAN '
        RCari13.Edit
        RCari13("KodeDist") = Trim(Combo3)
        RCari13("JD") = CCur(JMLDBT)
        RCari13("JAKHIR") = CCur(JMLAKHIR)
        RCari13("MUTASID") = CCur(MUTASIDBT)
        RCari13("SALDO") = CCur(SALDOAKHIR)
        
            If CCur(HJUAL) < RCari13("HJUAL") Then
                HHJUAL = RCari13("HJUAL")
            ElseIf CCur(HJUAL) = RCari13("HJUAL") Then
                HHJUAL = RCari13("HJUAL")
            Else
                HHJUAL = CCur(HJUAL)
            End If
        
        RCari13("HJUAL") = CCur(HHJUAL)
            
    RCari13.Update
    RCari13.Close
    Set RCari13 = Nothing
        
        RCari14.Update
        RCari14.Close
        Set RCari14 = Nothing

            'UPDATE HISTORY TRANSAKSI'
            SCari15 = "Select * From B005"
            Set RCari15 = RDCO.OpenResultset(SCari15, rdOpenKeyset, rdConcurRowVer)
            RCari15.AddNew
            RCari15("Status") = 1
            RCari15("KODE_TRANS") = "BL"
            RCari15("KODE_JNS") = RCari12("Kode_BR")
            RCari15("NAMA_JNS") = RCari12("Nama_BR")
            RCari15("NO_FAKTUR") = RCari12("No_Fak")
            RCari15("NO_BUKTI") = RCari12("No_Trans")
            RCari15("KETERANGAN") = "PEMBELIAN NO." + RCari12("No_Trans")
            RCari15("JML_DBT") = RCari12("Jml_Beli")
            RCari15("JML_CRD") = 0
            RCari15("JML_AKHIR") = CCur(JMLAKHIR)
            RCari15("MUTASI_DBT") = RCari12("Jml_Harga")
            RCari15("MUTASI_CRT") = 0
            RCari15("SALDO_AKHIR") = CCur(SALDOAKHIR)
            RCari15("H_POKOK") = RCari12("Harga_PCS")
            RCari15("NOMDISC") = 0
            RCari15("SPCDISC") = 0
            RCari15("LABA") = 0
            RCari15("KAS") = CCur(SALDOAKHIR)
            RCari15("KODE_SPL") = Trim(Combo3)
            If Frame4.Visible = True Then
                RCari15("NO_HUTANG") = Trim(Text5)
                RCari15("LUNAS") = 0
            Else
                RCari15("NO_HUTANG") = "-"
            End If
            RCari15("TGL_S") = Date
            RCari15("TGL_FAK") = DTPicker1
            RCari15.Update
            RCari15.Close
            Set RCari15 = Nothing
        
RCari12.MoveNext
Loop
RCari12.Close
Set RCari12 = Nothing

SBahan = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RBahan = RDCO.OpenResultset(SBahan, rdOpenKeyset, rdConcurRowVer)
If RBahan.RowCount <> 0 Then
    NoFuckU = CCur(RBahan("NoBeli"))

    RBahan.Edit
    RBahan("NoBeli") = NoFuckU + 1
End If
RBahan.Update
RBahan.Close
Set RBahan = Nothing

End Sub

Private Sub cmdtutup_Click()
Unload Me
End Sub


Private Sub Command2_Click()
Frame1.Visible = False
Combo1.SetFocus
End Sub

Private Sub Command4_Click()
Frame1.Visible = True
Frame1.ZOrder
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=KASIR", rdDriverNoPrompt, False, CN)
Frame1.Visible = False

Call DelBL001
Call NoBukti
Call KOSONG
Call IsiCombo
Call SiapkanGrid
Call IsiGrid

Call SiapkanGridDF
Call IsiGridDF
    
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = DateAdd("m", 1, Date)

Call Auto
Text8 = 0
SSTab1.Tab = 0

Call IsiCombo3
Frame4.Visible = False

End Sub

Private Sub SiapkanGridDF()
With gridDF
    .Cols = 2
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2700: .Text = "NAMA": .CellAlignment = 4
End With
End Sub

Private Sub IsiGridDF()
SKTG = "Select * From B003 order by KodeBR Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGridDF
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      gridDF.Rows = B + 1
      gridDF.Row = B
         With gridDF
              .Col = 0: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("NamaBR")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub DelBL001()
SDel = "Delete * From BL001"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing
End Sub

Private Sub Auto()
Dim No As Double
SqlNo = "Select count(*) as No from BL001"
Set RSLNO = RDCO.OpenResultset(SqlNo, rdOpenDynamic, rdConcurRowVer)
No = Val(RSLNO("No")) + 1
Text2 = No
RSLNO.Close
Set RSLNO = Nothing
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 6
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 3000: .Text = "NAMA BARANG": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1000: .Text = "JUMLAH": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 2000: .Text = "HARGA PCS": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 2000: .Text = "SUB TOTAL": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari4 = "Select * From BL001"
Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurReadOnly)
If RCari4.RowCount <> 0 Then
   RCari4.MoveFirst
   B = 1
   Do Until RCari4.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari4("No_Urut"): .CellAlignment = 4
              .Col = 1: .Text = RCari4("Kode_BR"): .CellAlignment = 4
              .Col = 2: .Text = RCari4("Nama_BR")
              .Col = 3: .Text = RCari4("Jml_Beli"): .CellAlignment = 4
              .Col = 4: .Text = Format(RCari4("Harga_PCS"), "##,###.00")
              .Col = 5: .Text = Format(RCari4("Jml_Harga"), "##,###.00")
         End With
      B = B + 1
      RCari4.MoveNext
   Loop
End If
RCari4.Close
Set RCari4 = Nothing
End Sub

Private Sub IsiCombo()
SKTG = "Select * From B003 order by kodebr"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenDynamic, rdOpenKeyset)
RKTG.MoveFirst
Do While Not RKTG.EOF
    Combo1.AddItem RKTG("KodeBR")
RKTG.MoveNext
Loop
RKTG.Close
Set RKTG = Nothing

SSTN = "Select * From B003 order by namabr"
Set RSTN = RDCO.OpenResultset(SSTN, rdOpenDynamic, rdOpenKeyset)
RSTN.MoveFirst
Do While Not RSTN.EOF
    Combo2.AddItem RSTN("NamaBR")
RSTN.MoveNext
Loop
RSTN.Close
Set RSTN = Nothing

End Sub

Private Sub KOSONG()
ClearTextBoxes BL001
Combo1 = ""
Combo2 = ""
Combo3 = ""
End Sub

Private Sub NoBukti()
Dim No As Double
SqlNo = "Select * from C013 where nama = '" + Operator + "'"
Set RSLNO = RDCO.OpenResultset(SqlNo, rdOpenDynamic, rdConcurRowVer)
No = Val(RSLNO("NoBeli")) + 1
NoStr = Digit(7, No)
SkinLabel3 = "1." + NoStr
RSLNO.Close
Set RSLNO = Nothing
End Sub

Private Sub gridDF_dblClick()
Combo1 = (gridDF.TextMatrix(gridDF.Row, 0))
Combo2 = (gridDF.TextMatrix(gridDF.Row, 1))
Frame1.Visible = False
Text6.SetFocus
End Sub


'Private Sub SSTab1_Click(PreviousTab As Integer)
'If SSTab1.Tab = 1 Then Text14.SetFocus
'End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 0
    Combo1.SetFocus
End If
End Sub

Private Sub Text1_Lostfocus()
Text1 = Format(Text1, ">")
Call CekData
End Sub

Private Sub CekData()
If Text1.Text = "" Then Exit Sub

SCari = "Select * From B005 where NO_FAKTUR = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " NO TELAH DIGUNAKAN", vbCritical, "KONFIRMASI"
        Text1 = ""
        Text1.SetFocus
    End If

RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text10_LostFocus()
Text10 = Format(Text10, ">")
End Sub

Private Sub text14_GotFocus()
Frame4.Visible = False
Text14 = ""
Text15 = ""
Text4 = 0
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text14_LostFocus()
If Text14 = "" Then
    Text14 = 0
End If

If CCur(Text14) < CCur(Text8) Then
    Text14 = Format(Text14, "##,###.00")
    Text15 = CCur(Text8) - CCur(Text14)
    Text15 = Format(Text15, "##,###.00")
    Text4 = Format(Text14, "##,###.00")
    Frame4.Visible = True
ElseIf CCur(Text14) > CCur(Text8) Then
    MsgBox "PEMBAYARAN TUNAI MELEBIHI TOTAL PEMBELIAN", vbSystemModal, "KONFIRMASI"
    Text14.SetFocus
    Text14 = ""
Else
    Text15 = 0
    Text14 = Format(Text14, "##,###.00")
    Text4 = Format(Text14, "##,###.00")
    Frame4.Visible = False
    Combo3.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text14.SetFocus
End Select
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text6_LostFocus()
If Text6 = "" Then Exit Sub
Call IsiHarga
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text14.SetFocus
End Select
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If

If KeyAscii = 13 Then
    SendKeys "{TAB}"
    Text7 = Format(Text7, "##,###.00")
End If

End Sub

Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text14.SetFocus
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub combo2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SSTab1.Tab = 1
        Text14.SetFocus
End Select
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub IsiCombo3()
SSPL = "Select * From D001 order by kode"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
If RSPL.RowCount <> 0 Then
    RSPL.MoveFirst
    Do While Not RSPL.EOF
        Combo3.AddItem RSPL("Kode")
    RSPL.MoveNext
    Loop
    Combo3.ListIndex = 0
Else
    MsgBox "DATA SUPPLIER MASIH KOSONG", vbSystemModal, "KONFIRMASI"
End If
    RSPL.Close
    Set RSPL = Nothing
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()

If Combo1 = "" Then Exit Sub
SCari = "Select * From B003 where KodeBR='" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Combo2 = RCari("NamaBR")
    Text6.SetFocus
Else
    MsgBox "KODE BARANG BELUM TERDAFTAR", vbSystemModal, "KONFIRMASI"
    Combo1 = ""
    Combo1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Combo2_LostFocus()

If Combo2 = "" Then Exit Sub
SCari2 = "Select * From B003 where NamaBR='" + Trim(Combo2) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Combo1 = RCari2("KodeBR")
    Text6.SetFocus
Else
    MsgBox "NAMA BARANG BELUM TERDAFTAR", vbSystemModal, "KONFIRMASI"
    Combo2 = ""
    Combo2.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Combo3_LostFocus()
If Combo3 = "" Then Exit Sub
SCari3 = "Select * From D001 where Kode ='" + Trim(Combo3) + "'"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenDynamic, rdConcurRowVer)
If RCari3.RowCount <> 0 Then
    Text3 = RCari3("Nama")
End If
RCari3.Close
Set RCari3 = Nothing
End Sub

Private Sub IsiHarga()
If Combo1 = "" Or Combo2 = "" Then
    Text7.SetFocus
Else
    SHarga = "Select * From B003 where KodeBR ='" + Trim(Combo1) + "'"
    Set RHarga = RDCO.OpenResultset(SHarga, rdOpenDynamic, rdConcurRowVer)
    If RHarga.RowCount <> 0 Then
        Text7 = RHarga("HBeli")
    End If
    RHarga.Close
    Set RHarga = Nothing
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text9_LostFocus()
Text9 = Format(Text9, ">")
End Sub
