VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form B005 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK LAPORAN"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "CETAK PENJUALAN"
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
      Left            =   233
      TabIndex        =   9
      Top             =   1800
      Width           =   5385
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK PEMBELIAN"
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
      Left            =   233
      TabIndex        =   8
      Top             =   1155
      Width           =   5385
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NON TUNAI"
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
      Left            =   3428
      TabIndex        =   7
      Top             =   4050
      Width           =   2190
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TUNAI"
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
      Left            =   233
      TabIndex        =   6
      Top             =   4050
      Width           =   2190
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
      Left            =   233
      TabIndex        =   0
      Top             =   2655
      Width           =   5385
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   420
      Left            =   248
      TabIndex        =   1
      Top             =   135
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
      Format          =   22151169
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   420
      Left            =   3293
      TabIndex        =   2
      Top             =   135
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
      Format          =   22151169
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   3668
      OleObjectBlob   =   "B005.frx":0000
      TabIndex        =   3
      Top             =   660
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   240
      Left            =   623
      OleObjectBlob   =   "B005.frx":0078
      TabIndex        =   4
      Top             =   660
      Width           =   1560
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   1215
      Top             =   6390
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      Height          =   1545
      Left            =   -45
      ScaleHeight     =   1485
      ScaleWidth      =   7335
      TabIndex        =   5
      Top             =   990
      Width           =   7395
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2700
         OleObjectBlob   =   "B005.frx":00EE
         Top             =   3555
      End
   End
End
Attribute VB_Name = "B005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private SqlPass As String
Private tUser As rdoResultset
Private tMasuk As rdoResultset

Private RR, RCari As rdoResultset
Private SR, SCari As String

Private T, M, D, T2, M2, D2

Private Sub Command1_Click()
Call TGL
crpt.ReportFileName = App.Path & "\ReportTOKO\TransBeli.rpt"
crpt.SelectionFormula = "{B005.TGL_S} in date (" & T & "," & M & "," & D & ") to date (" & T2 & "," & M2 & "," & D2 & ")  "
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Command2_Click()
Call TGL
crpt.ReportFileName = App.Path & "\ReportTOKO\TransBeli2.rpt"
crpt.SelectionFormula = "{B005.TGL_S} in date (" & T & "," & M & "," & D & ") to date (" & T2 & "," & M2 & "," & D2 & ")  "
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Command3_Click()
Call TGL
crpt.ReportFileName = App.Path & "\ReportTOKO\TransBeli3.rpt"
crpt.SelectionFormula = "{B005.TGL_S} in date (" & T & "," & M & "," & D & ") to date (" & T2 & "," & M2 & "," & D2 & ")  "
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Command4_Click()
Call TGL
crpt.ReportFileName = App.Path & "\ReportTOKO\TransJual.rpt"
crpt.SelectionFormula = "{B005.TGL_S} in date (" & T & "," & M & "," & D & ") to date (" & T2 & "," & M2 & "," & D2 & ")  "
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=TOKO", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
DTPicker1 = Date
DTPicker2 = Date
End Sub

Private Sub TGL()
T = Year(DTPicker1)
M = Month(DTPicker1)
D = Day(DTPicker1)

T2 = Year(DTPicker2)
M2 = Month(DTPicker2)
D2 = Day(DTPicker2)
End Sub

Private Sub cmdCLOSE_Click()
Unload Me
End Sub


