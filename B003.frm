VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form B003 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE BARANG"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2496
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   488
      Width           =   2580
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2496
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1995
      Width           =   1440
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2496
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   2580
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
      Left            =   7251
      TabIndex        =   8
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   2640
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   242
      TabIndex        =   6
      ToolTipText     =   "Klik untuk edit"
      Top             =   2625
      Width           =   1890
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2496
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1590
      Width           =   2580
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2496
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   1215
      Width           =   2580
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2496
      TabIndex        =   0
      Text            =   "Combo4"
      Top             =   90
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2490
      TabIndex        =   7
      Top             =   6810
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   55836673
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   390
      OleObjectBlob   =   "B003.frx":0000
      Top             =   9000
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   210
      Left            =   285
      OleObjectBlob   =   "B003.frx":0234
      TabIndex        =   9
      Top             =   892
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   210
      Left            =   285
      OleObjectBlob   =   "B003.frx":02A8
      TabIndex        =   10
      Top             =   135
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   210
      Left            =   285
      OleObjectBlob   =   "B003.frx":0328
      TabIndex        =   11
      Top             =   1642
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   210
      Left            =   285
      OleObjectBlob   =   "B003.frx":039A
      TabIndex        =   12
      Top             =   2047
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   210
      Left            =   285
      OleObjectBlob   =   "B003.frx":0404
      TabIndex        =   13
      Top             =   6855
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   210
      Left            =   285
      OleObjectBlob   =   "B003.frx":0470
      TabIndex        =   14
      Top             =   1267
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   210
      Left            =   285
      OleObjectBlob   =   "B003.frx":04E2
      TabIndex        =   15
      Top             =   540
      Width           =   2070
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3075
      Left            =   60
      TabIndex        =   16
      ToolTipText     =   "Klik untuk edit"
      Top             =   3465
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   5424
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      AllowUserResizing=   3
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   165
      Left            =   4140
      OleObjectBlob   =   "B003.frx":0556
      TabIndex        =   17
      Top             =   165
      Width           =   5115
   End
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   -135
      ScaleHeight     =   810
      ScaleWidth      =   9945
      TabIndex        =   18
      Top             =   2475
      Width           =   10005
   End
End
Attribute VB_Name = "B003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSLUser, RSave, RSave2, REdit, RKTG, RKTG2, RSTN, RSPL, RPBR, RDATE, RCari, RCari2, RCari3, RCari4, RCari5 As rdoResultset
Private SQLUser, SSave, SSave2, SEdit, SKTG, SKTG2, SSTN, SSPL, SPBR, SDATE, SCari, SCari2, SCari3, SCari4, SCari5 As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdHAPUS_Click()
MsgBox "ANDA AKAN MENGHAPUS DATA OBAT", vbCritical, "KONFIRMASI"

End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Or Text5 = "" Or Text4 = "" Or Combo1 = "" Or Combo1 = "" Or Combo4 = "" Or DTPicker1 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Combo4.SetFocus
Exit Sub
End If

If Text1 = "" Then Exit Sub
SQLTiket = "Select * from B003 where KodeBR = '" + Trim(Text1) + "'"
Set RSLTiket = RDCO.OpenResultset(SQLTiket, rdOpenDynamic, rdConcurRowVer)
If RSLTiket.RowCount <> 0 Then
    MsgBox "KODE BARANG SUDAH ADA", vbCritical, "KONFIRMASI"
    Call KOSONG
    Text1.SetFocus
Else
    Call Simpan
    Call KOSONG
    Call SiapkanGrid
    Call IsiGrid
    'Call CekSatuan
End If
RSLTiket.Close
Set RSLTiket = Nothing
Text1.SetFocus
End Sub

Private Sub CekSatuan()
SSave = "Select * From ST01 where NSATUAN = '" + Trim(Combo1) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
If RSave.RowCount = 0 Then
    RSave.AddNew
        RSave("NSatuan") = Combo1
    RSave.Update
End If
    RSave.Close
    Set RSave = Nothing
End Sub

Private Sub Simpan()
SSave = "Select * From B003"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("KodeInd") = Combo4
        RSave("KodeBR") = Trim(Text1)
        RSave("NamaBR") = Trim(Text2)
        RSave("JAkhir") = 0
        RSave("HBeli") = CCur(Text5)
        RSave("HJual") = CCur(Text4)
        RSave("Persen") = 0
        RSave("Satuan") = Combo1
        RSave("Tanggal") = DTPicker1
        RSave("Status") = 1
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=TOKO", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Call KOSONG

Call SiapkanGrid
Call IsiGrid
grid.Refresh

SSTN = "Select * From ST01"
Set RSTN = RDCO.OpenResultset(SSTN, rdOpenDynamic, rdOpenKeyset)
RSTN.MoveFirst
Do While Not RSTN.EOF
    Combo1.AddItem RSTN("NSatuan")
RSTN.MoveNext
Loop
RSTN.Close
Set RSTN = Nothing

SSPL = "Select Kode From B001 order by KODE Asc"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
If RSPL.RowCount = 0 Then
    MsgBox "KODE INDUK BARANG MASIH KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
Else
    RSPL.MoveFirst
    Do While Not RSPL.EOF
        Combo4.AddItem RSPL("Kode")
    RSPL.MoveNext
    Loop
    RSPL.Close
    Set RSPL = Nothing
End If

End Sub

Private Sub KOSONG()

ClearTextBoxes Me

Combo1 = ""
Combo4 = ""

SkinLabel6 = ""

DTPicker1 = Date

End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 1000: .Text = "INDUK": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2700: .Text = "NAMA": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 750: .Text = "JUMLAH": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1500: .Text = "BELI Rp.": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "JUAL Rp.": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SKTG = "Select * From B003 order by KodeBR Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RKTG("KodeInd"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 2: .Text = RKTG("NamaBR")
              .Col = 3: .Text = RKTG("JAkhir"): .CellAlignment = 4
              .Col = 4: .Text = Format(RKTG("HBeli"), "##,###.00")
              .Col = 5: .Text = Format(RKTG("HJual"), "##,###.00")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text1_Lostfocus()
Text1 = Format(Text1, ">")
Call CekData
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
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text4_GotFocus()
Text4 = ""
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
If Text5 = 0 Then Exit Sub
Text4 = Format(Text4, "##,###.00")
End Sub

Private Sub Text5_GotFocus()
Text5 = ""
Text4 = ""
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text5_LostFocus()
If Text5 = "" Then
    Text5 = 0
End If
Text5 = Format(Text5, "##,###.00")
End Sub

Private Sub CekData()
If Text1.Text = "" Then Exit Sub

SCari = "Select * From B003 where KodeBR = '" + Trim(Text1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " KODE BARANG SUDAH TERDAFTAR", vbCritical, "KONFIRMASI"
        Text1 = ""
        Text1.SetFocus
    Else
       Text2.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
Combo1 = Format(Combo1, ">")
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo4_LostFocus()
If Combo4 = "" Then Exit Sub
SCari2 = "Select * From B001 where Kode = '" + Combo4 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel6 = RCari2("Nama")
Else
    MsgBox "KODE INDUK BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo4.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

'Private Sub grid_Click()
'grid.Row = 0
'BR = ""
'Clipboard.SetText (grid.Text)
'BR = grid.Text
'End Sub

Private Sub grid_dblClick()
grid.Col = 1
BR = ""
Clipboard.SetText (grid.Text)
BR = grid.Text

If BR = "" Then Exit Sub

Unload Me
B003EDIT.Show 1
End Sub




