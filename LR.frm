VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form LR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN LABA RUGI"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crpt 
      Left            =   315
      Top             =   1890
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LAPORAN RUGI LABA KOTOR"
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
      Left            =   112
      TabIndex        =   1
      ToolTipText     =   "Klik untuk lihat laporan"
      Top             =   735
      Width           =   5670
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
      Left            =   1267
      TabIndex        =   0
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   6090
      Width           =   3360
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   525
      OleObjectBlob   =   "LR.frx":0000
      Top             =   7350
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4635
      Left            =   75
      TabIndex        =   2
      ToolTipText     =   "Klik untuk edit"
      Top             =   1365
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   8176
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      GridColor       =   0
      Enabled         =   -1  'True
      TextStyle       =   3
      TextStyleFixed  =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   420
      TabIndex        =   3
      Top             =   210
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   63242241
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   3045
      TabIndex        =   4
      Top             =   210
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   63242241
      CurrentDate     =   39286
      MinDate         =   39083
   End
End
Attribute VB_Name = "LR"
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

Private Sub Command1_Click()
Call TGL
'crpt.ReportFileName = App.Path & "\ReportTOKO\LR.rpt"
'crpt.SelectionFormula = "{LR.TGL_S} in date (" & T & "," & M & "," & D & ") to date (" & T2 & "," & M2 & "," & D2 & ")  "
'crpt.WindowState = crptMaximized
'crpt.WindowMaxButton = False
'crpt.WindowMinButton = False
'crpt.Action = 1
Call IsiGrid
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=TOKO", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
DTPicker1 = Date
DTPicker2 = Date

Call SiapkanGrid

End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 2
    .Col = 0: .ColWidth(0) = 1500: .Text = "TGL": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 3500: .Text = "LABA / RUGI": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * From LR"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari("TGL_S"): .CellAlignment = 4
              .Col = 1: .Text = Format(CCur(RCari("LABA")), "##,###.00")
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub
