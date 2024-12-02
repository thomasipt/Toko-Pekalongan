VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MAINMENU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MENU UTAMA"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   9450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "HUTANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   6885
      TabIndex        =   14
      ToolTipText     =   "Klik untuk masuk ke menu penjualan"
      Top             =   1785
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PIUTANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1350
      TabIndex        =   13
      ToolTipText     =   "Klik untuk masuk ke menu penjualan"
      Top             =   1785
      Width           =   1215
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   5400
      Top             =   315
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PEMBELIAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   8138
      TabIndex        =   12
      ToolTipText     =   "Klik untuk masuk ke menu pembelian"
      Top             =   1785
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PENJUALAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   98
      TabIndex        =   11
      ToolTipText     =   "Klik untuk masuk ke menu penjualan"
      Top             =   1785
      Width           =   1215
   End
   Begin VB.CommandButton cmdCLOSE 
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
      Height          =   555
      Left            =   98
      TabIndex        =   10
      ToolTipText     =   "Klik untuk keluar dari sistem dan kembali ke LOGIN"
      Top             =   7350
      Width           =   9255
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6293
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   6825
      Width           =   3060
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3195
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6825
      Width           =   3060
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   98
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6825
      Width           =   3060
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   98
      TabIndex        =   2
      Top             =   5520
      Width           =   9255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":0000
         TabIndex        =   4
         Top             =   480
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1710
      Left            =   98
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":0070
         TabIndex        =   1
         Top             =   240
         Width           =   8775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":00E0
         TabIndex        =   3
         Top             =   840
         Width           =   8775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":0150
         TabIndex        =   5
         Top             =   1200
         Width           =   8775
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "MAINMENU.frx":01C0
      Top             =   9900
   End
   Begin VB.PictureBox Picture1 
      Height          =   2745
      Left            =   98
      Picture         =   "MAINMENU.frx":03F4
      ScaleHeight     =   179
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   6
      Top             =   2730
      Width           =   9255
   End
   Begin VB.Menu P 
      Caption         =   "PENJUALAN"
      Index           =   1
      Begin VB.Menu PJ 
         Caption         =   "TRANSAKSI PENJUALAN"
         Index           =   11
      End
      Begin VB.Menu PJ 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu PJ 
         Caption         =   "DAFTAR PIUTANG"
         Index           =   13
      End
   End
   Begin VB.Menu B 
      Caption         =   "PEMBELIAN"
      Index           =   2
      Begin VB.Menu PB 
         Caption         =   "TRANSAKSI PEMBELIAN"
         Index           =   21
      End
      Begin VB.Menu PB 
         Caption         =   "-"
         Index           =   22
      End
      Begin VB.Menu PB 
         Caption         =   "HUTANG"
         Index           =   23
      End
   End
   Begin VB.Menu D 
      Caption         =   "DATA"
      Index           =   31
      Begin VB.Menu DS 
         Caption         =   "KODE KATEGORI BARANG"
         Index           =   31
      End
      Begin VB.Menu DS 
         Caption         =   "KODE BARANG"
         Index           =   32
      End
      Begin VB.Menu DS 
         Caption         =   "KODE PELANGGAN"
         Index           =   33
      End
      Begin VB.Menu DS 
         Caption         =   "KODE SUPPLIER"
         Index           =   34
      End
   End
   Begin VB.Menu T 
      Caption         =   "TOOLS"
      Index           =   4
      Begin VB.Menu TS 
         Caption         =   "SETING TOKO"
         Index           =   41
      End
      Begin VB.Menu TS 
         Caption         =   "GANTI PASSWORD"
         Index           =   42
      End
      Begin VB.Menu TS 
         Caption         =   "USER BARU"
         Index           =   43
         Visible         =   0   'False
      End
   End
   Begin VB.Menu L 
      Caption         =   "LAPORAN"
      Index           =   5
      Begin VB.Menu LS 
         Caption         =   "MUTASI BARANG"
         Index           =   501
      End
      Begin VB.Menu LS 
         Caption         =   "STOCK LIMIT"
         Index           =   502
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   503
      End
      Begin VB.Menu LS 
         Caption         =   "LAP TRANSAKSI"
         Index           =   504
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   507
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "LAP PENJUALAN"
         Index           =   508
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   509
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "LABA RUGI"
         Index           =   510
         Visible         =   0   'False
      End
   End
   Begin VB.Menu K 
      Caption         =   "KELUAR"
      Index           =   6
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MAINMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private Sub cmdCLOSE_Click()
Unload Me
LOGIN.Show
End Sub

Private Sub Command1_Click()
JL001.Show 1
End Sub

Private Sub Command2_Click()
BL001.Show 1
End Sub

Private Sub Command3_Click()
H002_HIS.Show 1
End Sub

Private Sub Command4_Click()
H001_HIS.Show 1
End Sub

Private Sub DS_Click(Index As Integer)
Select Case Index
    Case 31
        B001.Show 1
    Case 32
        B003.Show 1
    Case 33
        P001.Show 1
    Case 34
        D001.Show 1
End Select
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

Text1 = "USER : " + Operator
Text2 = TglS
Text3 = "Copyright © 2008 IPT"

SkinLabel1 = NToko
SkinLabel4 = NAlamat
SkinLabel5 = NMOtto
SkinLabel6 = NTelepon
End Sub

Private Sub LS_Click(Index As Integer)
Select Case Index
    Case 501
        Call LapBR
    Case 504
        B005.Show 1
    Case 508
        Call LapTransJual
    Case 509
    Case 510
        LR.Show 1
End Select
End Sub

Private Sub LapBR()
crpt.ReportFileName = App.Path & "\ReportTOKO\LapBR.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub LapTransJual()
crpt.ReportFileName = App.Path & "\ReportTOKO\TransJual.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub PB_Click(Index As Integer)
Select Case Index
    Case 21
        BL001.Show 1
    Case 23
        H001_HIS.Show 1
        'H003.Show 1
End Select
End Sub

Private Sub PJ_Click(Index As Integer)
Select Case Index
    Case 11
        JL001.Show 1
    Case 12
    Case 13
        H002_HIS.Show 1
        'H004.Show 1
End Select
End Sub

Private Sub TS_Click(Index As Integer)
Select Case Index
    Case 41
        NAMA.Show 1
    Case 42
        GPASS.Show 1
    Case 43
        User.Show 1
End Select
End Sub
