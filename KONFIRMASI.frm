VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form KONFIRMASI 
   Caption         =   "KONFIRMASI"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "TIDAK"
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
      Left            =   5332
      TabIndex        =   1
      Top             =   1035
      Width           =   1890
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "YA"
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
      Left            =   232
      TabIndex        =   0
      Top             =   1035
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   630
      OleObjectBlob   =   "KONFIRMASI.frx":0000
      Top             =   3210
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   1222
      OleObjectBlob   =   "KONFIRMASI.frx":0234
      TabIndex        =   2
      Top             =   60
      Width           =   5490
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   1222
      OleObjectBlob   =   "KONFIRMASI.frx":02B6
      TabIndex        =   3
      Top             =   435
      Width           =   5490
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   82
      TabIndex        =   4
      Top             =   810
      Width           =   7290
   End
End
Attribute VB_Name = "KONFIRMASI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

