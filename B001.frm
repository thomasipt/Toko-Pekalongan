VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form B001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KODE KATEGORI BARANG"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "HAPUS"
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
      Left            =   2865
      TabIndex        =   8
      Top             =   987
      Width           =   1890
   End
   Begin VB.CommandButton cmdEDIT 
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
      Left            =   98
      TabIndex        =   7
      Top             =   987
      Width           =   1890
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1433
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   559
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1433
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   139
      Width           =   1275
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   315
      OleObjectBlob   =   "B001.frx":0000
      Top             =   6615
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
      Left            =   5618
      TabIndex        =   3
      Top             =   1002
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
      Left            =   98
      TabIndex        =   2
      Top             =   987
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   278
      OleObjectBlob   =   "B001.frx":0234
      TabIndex        =   4
      Top             =   147
      Width           =   930
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   278
      OleObjectBlob   =   "B001.frx":029A
      TabIndex        =   5
      Top             =   567
      Width           =   930
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2220
      Left            =   98
      TabIndex        =   6
      ToolTipText     =   "Klik untuk edit"
      Top             =   1722
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   3916
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "B001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSkin, RToko, RTgl, RHapus, RDEl, RSave2, RSave3, RSave4, RCari, RCari2, RSLNO, rscs3 As rdoResultset
Private SSkin, SToko, STgl, SHapus, SDel, SSave2, SSave3, SSave4, SCari, SCari2, SqlNo, sqlcs3, Kode As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdEDIT_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "DATA TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Else
    SSave = "Select * From B001 where Kode = '" + Trim(KB) + "'"
    Set RSave = RDCO.OpenResultset(SSave, rdOpenKeyset, rdConcurRowVer)
    RSave.Edit
        RSave("Kode") = Trim(Text1)
        RSave("Nama") = Trim(Text2)
    RSave.Update
    RSave.Close
    Set RSave = Nothing
    Call IsiGrid
    ClearTextBoxes B001
    Text1.SetFocus
    cmdOK.Visible = True
    cmdEDIT.Visible = False
End If
End Sub

Private Sub cmdOK_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "DATA TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Else
    SSave2 = "Select * From B001"
    Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
    RSave2.AddNew
        RSave2("Kode") = Trim(Text1)
        RSave2("Nama") = Format(Text2, ">")
    RSave2.Update
    RSave2.Close
    Set RSave2 = Nothing
    Call IsiGrid
    ClearTextBoxes B001
    Text1.SetFocus
    cmdOK.Visible = True
    cmdEDIT.Visible = False
End If
End Sub

Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Then
    MsgBox "DATA TIDAK BOLEH KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Else
    SDel = "Delete From B001 where Kode = '" + Trim(KB) + "'"
    Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
    RDEl.Close
    Set RDEl = Nothing
    Unload Me
    B001.Show 1
End If
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=TOKO", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
ClearTextBoxes B001
Call SiapkanGrid
Call IsiGrid
cmdOK.Visible = True
cmdEDIT.Visible = False
Command1.Visible = False
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 2
    .Col = 0: .ColWidth(0) = 2000: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 5000: .Text = "NAMA KATEGORI": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * From B001 order by KODE Asc"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari("Kode"): .CellAlignment = 4
              .Col = 1: .Text = RCari("Nama")
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub grid_dblClick()
grid.Col = 0
KB = ""
Clipboard.SetText (grid.Text)
KB = grid.Text

If KB = "" Then Exit Sub

cmdOK.Visible = False
cmdEDIT.Visible = True
Command1.Visible = True

Call IsiText

End Sub

Private Sub IsiText()
SCari2 = "Select * From B001 where Kode = '" + Trim(KB) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Text1 = RCari2("Kode")
    Text2 = RCari2("Nama")
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text1_GotFocus()
Text1.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text1_Lostfocus()
Text1.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = RGB(255, 255, 0)
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = RGB(255, 255, 255)
Text2 = Format(Text2, ">")
End Sub


