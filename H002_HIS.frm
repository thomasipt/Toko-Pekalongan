VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form H002_HIS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAFTAR PIUTANG"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
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
      Height          =   555
      Left            =   10050
      TabIndex        =   27
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   7042
      Width           =   2730
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PENERIMAAN   CEK / BG"
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
      Left            =   10050
      TabIndex        =   4
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   6367
      Width           =   2730
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
      Left            =   1695
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2662
      Width           =   3525
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
      Left            =   10050
      TabIndex        =   1
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   7687
      Width           =   2730
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6000
      OleObjectBlob   =   "H002_HIS.frx":0000
      Top             =   2497
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2580
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Klik untuk proses"
      Top             =   22
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   4551
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16776960
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   195
      Left            =   180
      OleObjectBlob   =   "H002_HIS.frx":0234
      TabIndex        =   3
      Top             =   2767
      Width           =   1410
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   195
      Left            =   15
      OleObjectBlob   =   "H002_HIS.frx":02A8
      TabIndex        =   6
      Top             =   5782
      Width           =   2805
   End
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   2580
      Left            =   45
      TabIndex        =   16
      ToolTipText     =   "Klik untuk proses"
      Top             =   3172
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   4551
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   8438015
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   195
      Left            =   10020
      OleObjectBlob   =   "H002_HIS.frx":033E
      TabIndex        =   17
      Top             =   5782
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entry Cek / BG"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   45
      TabIndex        =   5
      Top             =   6007
      Width           =   8940
      Begin VB.CommandButton Command5 
         Caption         =   "EDIT"
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
         Left            =   6885
         TabIndex        =   28
         ToolTipText     =   "Klik untuk kembali ke menu utama"
         Top             =   900
         Width           =   1875
      End
      Begin VB.CommandButton Command1 
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
         Height          =   555
         Left            =   6885
         TabIndex        =   26
         ToolTipText     =   "Klik untuk kembali ke menu utama"
         Top             =   270
         Width           =   1875
      End
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
         Height          =   555
         Left            =   6885
         TabIndex        =   25
         ToolTipText     =   "Klik untuk kembali ke menu utama"
         Top             =   1560
         Width           =   1875
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4995
         TabIndex        =   13
         Text            =   "7"
         Top             =   1395
         Width           =   1680
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4995
         TabIndex        =   12
         Text            =   "6"
         Top             =   1035
         Width           =   1680
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   14
         Text            =   "5"
         Top             =   1800
         Width           =   6495
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1485
         TabIndex        =   9
         Text            =   "4"
         Top             =   1170
         Width           =   1680
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1485
         TabIndex        =   8
         Text            =   "3"
         Top             =   810
         Width           =   1680
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1485
         TabIndex        =   7
         Text            =   "2"
         Top             =   450
         Width           =   1680
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H002_HIS.frx":03CA
         TabIndex        =   15
         Top             =   495
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H002_HIS.frx":0430
         TabIndex        =   18
         Top             =   855
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H002_HIS.frx":04A8
         TabIndex        =   19
         Top             =   1215
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "H002_HIS.frx":0512
         TabIndex        =   20
         Top             =   1575
         Width           =   1365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   240
         Left            =   3330
         OleObjectBlob   =   "H002_HIS.frx":0584
         TabIndex        =   21
         Top             =   360
         Width           =   1590
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   240
         Left            =   3330
         OleObjectBlob   =   "H002_HIS.frx":05F4
         TabIndex        =   22
         Top             =   720
         Width           =   1590
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   240
         Left            =   3330
         OleObjectBlob   =   "H002_HIS.frx":066E
         TabIndex        =   23
         Top             =   1080
         Width           =   1725
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   240
         Left            =   3330
         OleObjectBlob   =   "H002_HIS.frx":06E6
         TabIndex        =   24
         Top             =   1440
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   4995
         TabIndex        =   10
         ToolTipText     =   "Klik untuk edit"
         Top             =   315
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
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
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   4995
         TabIndex        =   11
         ToolTipText     =   "Klik untuk edit"
         Top             =   675
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
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
   Begin Crystal.CrystalReport crpt 
      Left            =   15
      Top             =   22
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "H002_HIS"
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
Dim Tanya

If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "DATA MASIH KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Tanya = MsgBox("TRANSAKSI SELESAI", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        Call H002_HIS
    ElseIf Tanya = vbCancel Then
        MsgBox "CANCEL", vbCritical, "KONFIRMASI"
        Text2.SetFocus
    End If
    
IPT = Text1
ClearTextBoxes Me
Text1 = IPT
Frame1.Visible = False
Call IsiGrid2

End Sub

Private Sub H002_HIS_EDIT()
SSave = "Select * From H002_HIS where NO_GIRO='" + Trim(KB) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
If RSave.RowCount <> 0 Then
    RSave.Edit
        RSave("BANK_GIRO") = Trim(Text2)
        RSave("NO_GIRO") = Trim(Text3)
        RSave("PLAFON") = CCur(Text4)
        RSave("TGL_GIRO") = DTPicker1
        RSave("TGL_JTHTEMPO") = DTPicker2
        RSave("KPD_REK") = Trim(Text6)
        RSave("KPD_BANK") = Trim(Text7)
        RSave("KETERANGAN") = Trim(Text5)
    RSave.Update
End If
RSave.Close
Set RSave = Nothing
End Sub

Private Sub H002_HIS()
SSave = "Select * From H002_HIS"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NO_BUKTI") = Trim(Text1)
    RSave("NO_PIUTANG") = Trim(Text1)
    RSave("BANK_GIRO") = Trim(Text2)
    RSave("NO_GIRO") = Trim(Text3)
    RSave("PLAFON") = CCur(Text4)
    RSave("TGL_GIRO") = DTPicker1
    RSave("TGL_JTHTEMPO") = DTPicker2
    RSave("KPD_REK") = Trim(Text6)
    RSave("KPD_BANK") = Trim(Text7)
    RSave("KETERANGAN") = Trim(Text5)
    RSave("STATUS_BAYAR") = "BELUM"
    RSave("TANGGAL") = Date
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
    
    Command1.Visible = True
    Command3.Visible = True
    Command5.Visible = False
    
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
    MsgBox "PILIH NO. PIUTANG", vbCritical, "KONFIRMASI"
Else
    Frame1.Visible = True
    
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    
    Text2.SetFocus
    
    Command1.Visible = True
    Command5.Visible = False
    
End If
End Sub

Private Sub Command4_Click()
If Text1 = "" Then
    MsgBox "PILIH NO. PIUTANG", vbCritical, "KONFIRMASI"
ElseIf IPT = 0 Then
    MsgBox "DATA TIDAK ADA", vbCritical, "KONFIRMASI"
Else
    crpt.ReportFileName = App.Path & "\ReportTOKO\MutasiPiutang2.rpt"
    crpt.SelectionFormula = "{H001.NO_PIUTANG} = '" + Trim(Text1) + "'"
    crpt.WindowState = crptMaximized
    crpt.WindowMaxButton = False
    crpt.WindowMinButton = False
    crpt.Action = 1
End If
End Sub

Private Sub Command5_Click()
Dim Tanya

If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
    MsgBox "DATA MASIH KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Tanya = MsgBox("TRANSAKSI SELESAI", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        Call H002_HIS_EDIT
    ElseIf Tanya = vbCancel Then
        MsgBox "CANCEL", vbCritical, "KONFIRMASI"
        Text2.SetFocus
    End If
    
IPT = Text1
ClearTextBoxes Me
Text1 = IPT
Frame1.Visible = False
Call IsiGrid2
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=TOKO", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

Call SiapkanGrid1
Call SiapkanGrid2
Call IsiGrid

ClearTextBoxes Me
DTPicker1 = Date
DTPicker2 = Date

Frame1.Visible = False

End Sub

Private Sub SiapkanGrid1()
With grid
    .Cols = 7
    .Row = 0
    .RowHeight(0) = 500
    .Col = 0: .ColWidth(0) = 1500: .Text = "NO PIUTANG": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 1500: .Text = "NO BUKTI": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 1500: .Text = "NO BG": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 2500: .Text = "NAMA PELANGGAN": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 1600: .Text = "PLAFON": .CellAlignment = 4: .CellFontBold = True
    '.Col = 5: .ColWidth(5) = 1600: .Text = "BAKI DEBET": .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 1100: .Text = "TGL MULAI": .CellAlignment = 4: .CellFontBold = True
    .Col = 6: .ColWidth(6) = 1100: .Text = "JTH TEMPO": .CellAlignment = 4: .CellFontBold = True
End With
End Sub

Private Sub SiapkanGrid2()
With grid2
    .Cols = 10
    .Row = 0
    .RowHeight(0) = 500
    .Col = 0: .ColWidth(0) = 1000: .Text = "STS": .CellAlignment = 4: .CellFontBold = True
    .Col = 1: .ColWidth(1) = 1500: .Text = "BANK": .CellAlignment = 4: .CellFontBold = True
    .Col = 2: .ColWidth(2) = 1500: .Text = "NO GIRO/CEK": .CellAlignment = 4: .CellFontBold = True
    .Col = 3: .ColWidth(3) = 1500: .Text = "PLAFON": .CellAlignment = 4: .CellFontBold = True
    .Col = 4: .ColWidth(4) = 3000: .Text = "KETERANGAN": .CellAlignment = 4: .CellFontBold = True
    .Col = 5: .ColWidth(5) = 1500: .Text = "TGL.TERIMA": .CellAlignment = 4: .CellFontBold = True
    .Col = 6: .ColWidth(6) = 1500: .Text = "TGL.JTH": .CellAlignment = 4: .CellFontBold = True
    .Col = 7: .ColWidth(7) = 1500: .Text = "REK. TRANS": .CellAlignment = 4: .CellFontBold = True
    .Col = 8: .ColWidth(8) = 1500: .Text = "BANK. TRANS": .CellAlignment = 4: .CellFontBold = True
    .Col = 9: .ColWidth(9) = 500: .Text = "": .CellAlignment = 4: .CellFontBold = True
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
            .Col = 6: .Text = RCari("Tgl_Jatuh"): .CellAlignment = 4
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing

End Sub

Private Sub IsiGrid2()
grid2.Clear
grid2.Refresh

Call SiapkanGrid2

IPT = 0

SCari2 = "Select * From H002_HIS where NO_PIUTANG = '" + Trim(Text1) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurReadOnly)
If RCari2.RowCount <> 0 Then
   RCari2.MoveFirst
   B = 1
   Do Until RCari2.EOF
      grid2.Rows = B + 1
      grid2.Row = B
         With grid2
            .Col = 0: .Text = RCari2("STATUS_BAYAR"): .CellAlignment = 4: .CellFontBold = True
            .Col = 1: .Text = RCari2("BANK_GIRO"): .CellAlignment = 4
            .Col = 2: .Text = RCari2("NO_GIRO"): .CellAlignment = 4
            .Col = 3: .Text = RCari2("PLAFON"): .CellAlignment = 4
            .Col = 4: .Text = RCari2("KETERANGAN"): .CellAlignment = 4
            .Col = 5: .Text = RCari2("TGL_GIRO"): .CellAlignment = 4
            .Col = 6: .Text = RCari2("TGL_JTHTEMPO"): .CellAlignment = 4
            .Col = 7: .Text = RCari2("KPD_REK"): .CellAlignment = 4
            .Col = 8: .Text = RCari2("KPD_BANK"): .CellAlignment = 4
            .Col = 9: .Text = "EDIT": .CellAlignment = 4: .CellFontBold = True
         End With
      B = B + 1
      RCari2.MoveNext
   Loop
   IPT = 1
End If
RCari2.Close
Set RCari2 = Nothing

End Sub

Private Sub grid_dblClick()
grid.Col = 0
KB = ""
Clipboard.SetText (grid.Text)
KB = grid.Text

If KB = "" Then Exit Sub

Text1 = KB

Call IsiGrid2
End Sub

Private Sub grid2_dblClick()
If grid2.Col = 0 Then
    Call Klik_STS
ElseIf grid2.Col = 9 Then
    Call Klik_EDIT
End If
End Sub

Private Sub Klik_STS()
grid2.Col = 2
KB = ""
K = ""

Clipboard.SetText (grid2.Text)
KB = grid2.Text

If KB = "" Then Exit Sub

    SSave = "Select * From H002_HIS where NO_GIRO='" + Trim(KB) + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
    K = RSave("STATUS_BAYAR")
    If K = "BELUM" Then
        RSave.Edit
            RSave("STATUS_BAYAR") = "CAIR"
        RSave.Update
        RSave.Close
        Set RSave = Nothing
    ElseIf K = "CAIR" Then
        RSave.Edit
            RSave("STATUS_BAYAR") = "BATAL"
        RSave.Update
        RSave.Close
        Set RSave = Nothing
    ElseIf K = "BATAL" Then
        RSave.Edit
            RSave("STATUS_BAYAR") = "BELUM"
        RSave.Update
        RSave.Close
        Set RSave = Nothing
    End If
    
Call IsiGrid2
grid2.Refresh
End Sub

Private Sub Klik_EDIT()
Frame1.Visible = True

grid2.Col = 2
KB = ""
K = ""

Clipboard.SetText (grid2.Text)
KB = grid2.Text

If KB = "" Then Exit Sub
    
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    
    Command1.Visible = False
    Command3.Visible = False
    Command5.Visible = True

    SCari = "Select * From H002_HIS where NO_GIRO='" + Trim(KB) + "'"
    Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        Text2 = RCari("BANK_GIRO")
        Text3 = RCari("NO_GIRO")
        Text4 = Format(RCari("PLAFON"), "##,###.00")
        DTPicker1 = RCari("TGL_GIRO")
        DTPicker2 = RCari("TGL_JTHTEMPO")
        Text6 = RCari("KPD_REK")
        Text7 = RCari("KPD_BANK")
        Text5 = RCari("KETERANGAN")
    End If
    RCari.Close
    Set RCari = Nothing
    
Call IsiGrid2
grid2.Refresh
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text3_LostFocus()
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then
    Text4 = 0
End If
Text4 = Format(Text4, "##,###.00")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
Text6 = Format(Text6, ">")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_LostFocus()
Text7 = Format(Text7, ">")
End Sub
