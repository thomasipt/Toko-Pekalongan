VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LOGIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   735
      OleObjectBlob   =   "LOGIN.frx":0000
      Top             =   3255
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
      Left            =   2175
      TabIndex        =   3
      ToolTipText     =   "Klik untuk keluar"
      Top             =   1350
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "MASUK"
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
      Left            =   150
      TabIndex        =   2
      ToolTipText     =   "Klik untuk masuk ke sistem"
      Top             =   1350
      Width           =   1890
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      IMEMode         =   3  'DISABLE
      Left            =   1057
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Password user"
      Top             =   765
      Width           =   2370
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   1057
      TabIndex        =   0
      Text            =   "Text1"
      ToolTipText     =   "Nama user"
      Top             =   150
      Width           =   2370
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private SqlPass As String
Private tUser As rdoResultset
Private tMasuk As rdoResultset

Private RTgl, RHapus, RDEl, RSave2, RSave3, RSave4, RCari, RCari2, RSLNO, rscs3 As rdoResultset
Private STgl, SHapus, SDel, SSave2, SSave3, SSave4, SCari, SCari2, SqlNo, sqlcs3, Kode As String

Private Sub cmdCLOSE_Click()
End
End Sub

Private Sub Masuk2()
SCari = "Select * From C013 where UserCode = '" + Text1 + "' and Password = '" + Text2 + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Call Masuk
    Unload Me
Else
    LOGIN.Hide
    MsgBox "ANDA TIDAK BERHAK LOG IN KE SYSTEM", vbCritical, "KONFIRMASI"
    LOGIN.Show
    Text1 = ""
    Text2 = ""
    Text1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Masuk()
SqlPass = "Select * from C013 where UserCode =  '" + Trim(Text1) + "' "
Set tMasuk = RDCO.OpenResultset(SqlPass, rdOpenDynamic, rdConcurRowVer)
If tMasuk.RowCount <> 0 Then
    If tMasuk("MAIN") = "01" Then
        Operator = Trim(tMasuk("Nama"))
        MAINMENU.Show
    ElseIf tMasuk("MAIN") = "02" Then
        Operator = Trim(tMasuk("Nama"))
        MAINKASIR.Show
    End If
End If
tMasuk.Close
Set tMasuk = Nothing
End Sub

Private Sub cmdOK_Click()
Call Masuk2
End Sub

Private Sub TGL()
STgl = "Select * from A001"
Set RTgl = RDCO.OpenResultset(STgl, rdOpenDynamic, rdConcurRowVer)
If RTgl.RowCount <> 0 Then
    TglS = Date
    Skin = RTgl("S")
    NToko = RTgl("Toko")
    NAlamat = RTgl("ALamat")
    NMOtto = RTgl("Motto")
    NTelepon = RTgl("Telepon")
Else
End If
RTgl.Close
Set RTgl = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=TOKO", rdDriverNoPrompt, False, CN)
Text1 = ""
Text2 = ""
Call TGL
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_Lostfocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub



