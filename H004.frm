VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form H004 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PIUTANG"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "History Piutang"
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
      Left            =   97
      TabIndex        =   10
      Top             =   4320
      Width           =   4980
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1710
         TabIndex        =   14
         Text            =   "9"
         Top             =   1620
         Width           =   3165
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         TabIndex        =   13
         Text            =   "5"
         Top             =   330
         Width           =   3165
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   900
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "H004.frx":0000
         Top             =   2025
         Width           =   3165
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CETAK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3975
         TabIndex        =   11
         ToolTipText     =   "Klik untuk kembali ke menu utama"
         Top             =   3060
         Width           =   900
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H004.frx":0003
         TabIndex        =   15
         Top             =   360
         Width           =   1410
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H004.frx":0077
         TabIndex        =   16
         Top             =   810
         Width           =   1410
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H004.frx":00EF
         TabIndex        =   17
         Top             =   1260
         Width           =   1410
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H004.frx":0163
         TabIndex        =   18
         Top             =   1680
         Width           =   1410
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   420
         Left            =   1710
         TabIndex        =   19
         ToolTipText     =   "Klik untuk edit"
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
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
         TabIndex        =   20
         ToolTipText     =   "Klik untuk edit"
         Top             =   1170
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
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
         OleObjectBlob   =   "H004.frx":01CD
         TabIndex        =   21
         Top             =   2355
         Width           =   1410
      End
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
      Left            =   5317
      TabIndex        =   7
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   7350
      Width           =   6075
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   5452
      TabIndex        =   6
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   6255
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "PROSES PEMBAYARAN PIUTANG"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   6667
      TabIndex        =   2
      Top             =   6255
      Width           =   4590
      Begin VB.CommandButton Command3 
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
         Height          =   420
         Left            =   3330
         TabIndex        =   3
         ToolTipText     =   "Klik untuk kembali ke menu utama"
         Top             =   360
         Width           =   1080
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H004.frx":023F
         TabIndex        =   4
         Top             =   450
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   420
         Left            =   1530
         TabIndex        =   5
         ToolTipText     =   "Klik untuk edit"
         Top             =   360
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
   End
   Begin VB.CommandButton Command4 
      Caption         =   "REKAP LUNAS"
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
      Left            =   8677
      TabIndex        =   1
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   4545
      Width           =   2520
   End
   Begin VB.CommandButton Command5 
      Caption         =   "REKAP BELUM LUNAS"
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
      Left            =   5632
      TabIndex        =   0
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   4545
      Width           =   2520
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   5760
      Top             =   8460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5715
      OleObjectBlob   =   "H004.frx":02AB
      Top             =   8370
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4155
      Left            =   52
      TabIndex        =   8
      ToolTipText     =   "Klik untuk proses"
      Top             =   90
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   7329
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
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   420
      Left            =   6082
      TabIndex        =   22
      Top             =   5175
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
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   420
      Left            =   9127
      TabIndex        =   23
      Top             =   5175
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   9157
      OleObjectBlob   =   "H004.frx":04DF
      TabIndex        =   24
      Top             =   5700
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   240
      Left            =   6112
      OleObjectBlob   =   "H004.frx":0557
      TabIndex        =   25
      Top             =   5700
      Width           =   1560
   End
   Begin VB.Frame Frame3 
      Height          =   1680
      Left            =   5317
      TabIndex        =   26
      Top             =   4320
      Width           =   6075
   End
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   5317
      TabIndex        =   9
      Top             =   6030
      Width           =   6075
   End
End
Attribute VB_Name = "H004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSkin, RToko, RTgl, RHapus, RDEl, RSave2, RSave3, RSave4, RCari, RCari2, RSLNO, rscs3 As rdoResultset
Private SSkin, SToko, STgl, SHapus, SDel, SSave2, SSave3, SSave4, SCari, SCari2, SqlNo, sqlcs3, Kode As String

Private T, M, D, T2, M2, D2

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If Text5 = "" Then Exit Sub

Command1.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command2_Click()
If Text5 = "" Then Exit Sub

crpt.ReportFileName = App.Path & "\ReportTOKO\MutasiPiutang.rpt"
crpt.SelectionFormula = "{H001.NO_PIUTANG} = '" + Trim(Text5) + "'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Command3_Click()
SSave = "Select * From H002 where NO_PIUTANG = '" + Trim(Text5) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
RSave.Edit
    RSave("Tanggal") = DTPicker1
    RSave("Status_Bayar") = 1
RSave.Update
RSave.Close
Set RSave = Nothing
Unload Me
H004.Show 1
End Sub

Private Sub Command4_Click()
Call TGL

crpt.ReportFileName = App.Path & "\ReportTOKO\MutasiPiutang.rpt"
crpt.SelectionFormula = "{B005.TGL_S} in date (" & T & "," & M & "," & D & ") to date (" & T2 & "," & M2 & "," & D2 & ") and {H001.STATUS_BAYAR} = '1'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1

DTPicker4 = Date
DTPicker5 = Date

End Sub

Private Sub Command5_Click()
Call TGL

crpt.ReportFileName = App.Path & "\ReportTOKO\MutasiPiutang.rpt"
crpt.SelectionFormula = "{B005.TGL_S} in date (" & T & "," & M & "," & D & ") to date (" & T2 & "," & M2 & "," & D2 & ") and {H001.STATUS_BAYAR} = '0'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1

DTPicker4 = Date
DTPicker5 = Date

End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=TOKO", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Call SiapkanGrid
Call IsiGrid

Text5 = ""

DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
DTPicker4 = Date
DTPicker5 = Date

Text9 = ""
Text10 = ""

Frame1.Visible = False

Call TGL

End Sub

Private Sub TGL()
T = Year(DTPicker4)
M = Month(DTPicker4)
D = Day(DTPicker4)

T2 = Year(DTPicker5)
M2 = Month(DTPicker5)
D2 = Day(DTPicker5)
End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 7
    .Row = 0
    .RowHeight(0) = 500
    .Col = 0: .ColWidth(0) = 1500: .Text = "NO PIUTANG": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 1500: .Text = "NO BUKTI": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 1500: .Text = "NO BG": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 2500: .Text = "NAMA PELANGGAN": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 1600: .Text = "PLAFON": .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 1100: .Text = "TGL MULAI": .CellAlignment = 4: .CellFontBold = True
    .Col = 6: .ColWidth(6) = 1100: .Text = "JTH TEMPO": .CellAlignment = 4: .CellFontBold = True
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * From H002 where Status_Bayar = '0' order by Tgl_Mulai Asc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
            .Col = 0: .Text = RCari("No_Piutang"): .CellAlignment = 4
            .Col = 1: .Text = RCari("No_Bukti"): .CellAlignment = 4
            .Col = 2: .Text = RCari("No_BG"): .CellAlignment = 4
            .Col = 3: .Text = RCari("Nama_PLG")
            .Col = 4: .Text = Format(RCari("Plafon"), "##,###.00")
            .Col = 5: .Text = RCari("Tgl_Mulai"): .CellAlignment = 4
            If RCari("Tgl_Jatuh") = Date Then
                .Col = 6: .Text = RCari("Tgl_Jatuh"): .CellAlignment = 4: .CellBackColor = &HC0C0FF
            ElseIf RCari("Tgl_Jatuh") < Date Then
                .Col = 6: .Text = RCari("Tgl_Jatuh"): .CellAlignment = 4: .CellBackColor = &HC0FFFF
            ElseIf RCari("Tgl_Jatuh") <> Date Then
                .Col = 6: .Text = RCari("Tgl_Jatuh"): .CellAlignment = 4
            End If
            
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid_dblClick()
Frame1.Visible = False

KB = (grid.TextMatrix(grid.Row, 1))

SCari2 = "Select * From H002 where No_Bukti = '" + Trim(KB) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Text5 = RCari2("No_Piutang")
    DTPicker2 = RCari2("Tgl_Mulai")
    DTPicker3 = RCari2("Tgl_Jatuh")
    Text9 = RCari2("No_BG")
    Text10 = RCari2("Keterangan")
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

