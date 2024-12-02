VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form B003EDIT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT BARANG"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2512
      TabIndex        =   0
      Text            =   "Combo4"
      Top             =   135
      Width           =   1440
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2512
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   1260
      Width           =   2580
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2512
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1635
      Width           =   2580
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2512
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   885
      Width           =   2580
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2512
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2040
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2512
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   540
      Width           =   2580
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   8055
      Top             =   1485
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Text3 
      Height          =   765
      Left            =   5805
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   3690
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3390
      OleObjectBlob   =   "B003EDIT.frx":0000
      Top             =   5145
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2535
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   49152
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   16777088
      Format          =   22151169
      CurrentDate     =   39286
      MinDate         =   39083
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   210
      Left            =   345
      OleObjectBlob   =   "B003EDIT.frx":0234
      TabIndex        =   10
      Top             =   4395
      Width           =   945
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
      Left            =   7305
      TabIndex        =   7
      ToolTipText     =   "Klik untuk kembali ke menu utama"
      Top             =   2670
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
      Left            =   296
      TabIndex        =   6
      ToolTipText     =   "Klik untuk edit"
      Top             =   2655
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   1680
      Left            =   -210
      ScaleHeight     =   1620
      ScaleWidth      =   9945
      TabIndex        =   11
      Top             =   2520
      Width           =   10005
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   210
      Left            =   300
      OleObjectBlob   =   "B003EDIT.frx":02A0
      TabIndex        =   12
      Top             =   930
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   210
      Left            =   300
      OleObjectBlob   =   "B003EDIT.frx":0314
      TabIndex        =   13
      Top             =   180
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   210
      Left            =   300
      OleObjectBlob   =   "B003EDIT.frx":0394
      TabIndex        =   14
      Top             =   1680
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   210
      Left            =   300
      OleObjectBlob   =   "B003EDIT.frx":0406
      TabIndex        =   15
      Top             =   2085
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   210
      Left            =   300
      OleObjectBlob   =   "B003EDIT.frx":0470
      TabIndex        =   16
      Top             =   1305
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   210
      Left            =   300
      OleObjectBlob   =   "B003EDIT.frx":04E2
      TabIndex        =   17
      Top             =   585
      Width           =   2070
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   165
      Left            =   4155
      OleObjectBlob   =   "B003EDIT.frx":0556
      TabIndex        =   18
      Top             =   210
      Width           =   5115
   End
End
Attribute VB_Name = "B003EDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RSLUser, RSave, RSave2, REdit, RKTG, RSTN, RSPL, RPBR, RDATE, RCari, RCari2, RCari3, RCari4, RCari5 As rdoResultset
Private SQLUser, SSave, SSave2, SEdit, SKTG, SSTN, SSPL, SPBR, SDATE, SCari, SCari2, SCari3, SCari4, SCari5 As String

Private Sub IPT()
Text1 = BR
Text3 = BR

SCari3 = "Select * From B003 where KodeBR = '" + Trim(Text1) + "'"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurReadOnly)
If RCari3.RowCount <> 0 Then
    Combo4 = RCari3("KodeInd")
    Text1 = RCari3("KodeBR")
    Text2 = RCari3("namaBR")
    Text5 = Format(RCari3("HBeli"), "##,###.00")
    Text4 = Format(RCari3("HJual"), "##,###.00")
    Text6 = 0
    Combo1 = RCari3("Satuan")
    DTPicker1 = RCari3("Tanggal")
End If
RCari3.Close
Set RCari3 = Nothing
End Sub

Private Sub cmdEDIT_Click()
Dim tanya
tanya = MsgBox("YAKIN AKAN MERUBAH DATA", vbOKCancel, "KONFIRMASI")
If tanya = vbOK Then
    SCari4 = "Select * From B003 where KodeBR = '" + Trim(Text3) + "'"
    Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenDynamic, rdConcurRowVer)
    RCari4.Edit
        RCari4("KodeInd") = Combo4
        RCari4("KodeBR") = Text1
        RCari4("namaBR") = Text2
        RCari4("HBeli") = CCur(Text5)
        RCari4("HJual") = CCur(Text4)
        RCari4("Persen") = 0
        RCari4("Satuan") = Combo1
        RCari4("Tanggal") = DTPicker1
        RCari4("Status") = 1
    RCari4.Update
    RCari4.Close
    MsgBox "DATABASE TELAH DI UPDATE", vbCritical, "KONFIRMASI"
End If
Unload Me
B003.Show 1
End Sub

Private Sub cmdCLOSE_Click()
Unload Me
B003.Show 1
End Sub

Private Sub Combo4_LostFocus()
Call Cari2
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=TOKO", rdDriverNoPrompt, False, CN)
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

SSTN = "Select * From ST01"
Set RSTN = RDCO.OpenResultset(SSTN, rdOpenDynamic, rdOpenKeyset)
RSTN.MoveFirst
Do While Not RSTN.EOF
    Combo1.AddItem RSTN("NSatuan")
RSTN.MoveNext
Loop
RSTN.Close
Set RSTN = Nothing

SSPL = "Select Kode From B001"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
RSPL.MoveFirst
Do While Not RSPL.EOF
    Combo4.AddItem RSPL("Kode")
RSPL.MoveNext
Loop
RSPL.Close
Set RSPL = Nothing

Call IPT
Call Cari2
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
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

Private Sub text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        Combo1.SetFocus
        Text4 = Format(Text4, "##,###.00")
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        Text5 = Format(Text5, "##,###.00")
    End If
End Sub

Private Sub Cari2()
SCari2 = "Select * From B001 where Kode = '" + Combo4 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel6 = RCari2("Nama")
Else
    MsgBox "KODE INDUK BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    'Combo4.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub



