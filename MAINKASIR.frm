VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MAINKASIR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MENU KASIR"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   2
      Top             =   1320
      Width           =   8355
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
      Left            =   3218
      TabIndex        =   1
      Top             =   120
      Width           =   2115
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1950
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "MAINKASIR.frx":0000
      Top             =   3480
   End
End
Attribute VB_Name = "MAINKASIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
    With StatusBar1.Panels
        .Item(1).Style = sbrText
        .Item(1).Text = "USERCODE : " & Operator
        .Item(1).AutoSize = sbrSpring
        .Item(2).Style = sbrText
        .Item(2).AutoSize = sbrSpring
        .Item(2).Text = "TANGGAL SYSTEM  : " & TglS
        .Item(3).Style = sbrText
        .Item(3).AutoSize = sbrSpring
        .Item(3).Text = "Copyright® EDP IT SOLUTION"
    End With
End Sub

Private Sub K_Click(Index As Integer)
cepat = 1000
While Top - Height < Screen.Height
    DoEvents
    Top = Top + cepat
Wend
Hide
Unload Me
LOGIN.Show 1
End Sub

