VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form JL001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PENJUALAN"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   90
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   525
      Width           =   4110
   End
   Begin VB.TextBox Text15 
      Height          =   315
      Left            =   5400
      TabIndex        =   21
      Text            =   "Text15"
      Top             =   8175
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4095
      TabIndex        =   19
      Text            =   "1,000,000.00"
      Top             =   525
      Width           =   5610
   End
   Begin VB.CommandButton cmdbatal 
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
      Height          =   465
      Left            =   221
      TabIndex        =   15
      ToolTipText     =   "Klik untuk membatalkan transaksi"
      Top             =   6795
      Width           =   1725
   End
   Begin VB.CommandButton cmdtutup 
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
      Height          =   465
      Left            =   7848
      TabIndex        =   16
      ToolTipText     =   "Klik untuk keluar tanpa melakukan transaksi"
      Top             =   6795
      Width           =   1725
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1575
      OleObjectBlob   =   "JL001.frx":0000
      Top             =   11625
   End
   Begin VB.PictureBox crpt 
      Height          =   480
      Left            =   375
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   38
      Top             =   11850
      Width           =   1200
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   150
      OleObjectBlob   =   "JL001.frx":0234
      TabIndex        =   18
      Top             =   75
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   285
      Left            =   1800
      OleObjectBlob   =   "JL001.frx":02AA
      TabIndex        =   20
      Top             =   53
      Width           =   3660
   End
   Begin VB.CommandButton Command2 
      Height          =   840
      Left            =   37
      TabIndex        =   22
      Top             =   6600
      Width           =   9720
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   240
      Left            =   105
      OleObjectBlob   =   "JL001.frx":031E
      TabIndex        =   23
      Top             =   6225
      Width           =   3540
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   -300
      TabIndex        =   26
      Top             =   375
      Width           =   10290
   End
   Begin VB.Frame BAYAR 
      Height          =   4815
      Left            =   37
      TabIndex        =   25
      Top             =   1275
      Width           =   9720
      Begin VB.Frame Frame4 
         Caption         =   "Syarat Bayar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   4680
         TabIndex        =   46
         Top             =   1530
         Width           =   4980
         Begin VB.TextBox Text19 
            BackColor       =   &H00C0E0FF&
            Height          =   900
            Left            =   1710
            MultiLine       =   -1  'True
            TabIndex        =   13
            Text            =   "JL001.frx":03B2
            Top             =   2160
            Width           =   3165
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
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
            TabIndex        =   9
            Text            =   "12"
            Top             =   330
            Width           =   3165
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Left            =   1710
            TabIndex        =   12
            Text            =   "11"
            Top             =   1710
            Width           =   3165
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "JL001.frx":03B7
            TabIndex        =   47
            Top             =   360
            Width           =   1410
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "JL001.frx":042B
            TabIndex        =   48
            Top             =   810
            Width           =   1410
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "JL001.frx":04A3
            TabIndex        =   49
            Top             =   1305
            Width           =   1410
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "JL001.frx":0517
            TabIndex        =   50
            Top             =   1770
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   420
            Left            =   1710
            TabIndex        =   10
            ToolTipText     =   "Klik untuk edit"
            Top             =   720
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
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   420
            Left            =   1710
            TabIndex        =   11
            ToolTipText     =   "Klik untuk edit"
            Top             =   1215
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "JL001.frx":0581
            TabIndex        =   51
            Top             =   2490
            Width           =   1410
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Pelanggan"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   4680
         TabIndex        =   42
         Top             =   135
         Width           =   4980
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1665
            TabIndex        =   43
            Text            =   "10"
            Top             =   945
            Width           =   3210
         End
         Begin VB.ComboBox Combo3 
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
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Klik untuk edit"
            Top             =   450
            Width           =   3210
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   240
            Left            =   135
            OleObjectBlob   =   "JL001.frx":05F3
            TabIndex        =   44
            Top             =   540
            Width           =   1365
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   240
            Left            =   135
            OleObjectBlob   =   "JL001.frx":066B
            TabIndex        =   45
            Top             =   945
            Width           =   1365
         End
      End
      Begin VB.CommandButton cmdsimpan 
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
         Height          =   600
         Left            =   90
         TabIndex        =   14
         ToolTipText     =   "Klik untuk simpan transaksi"
         Top             =   4110
         Width           =   4530
      End
      Begin VB.Frame Frame5 
         Caption         =   "Transaksi"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   75
         TabIndex        =   52
         Top             =   135
         Width           =   4530
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            TabIndex        =   7
            Text            =   "Text4"
            ToolTipText     =   "Nominal bayar"
            Top             =   2445
            Width           =   2310
         End
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
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
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "Text16"
            Top             =   1425
            Width           =   2310
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "Text9"
            Top             =   960
            Width           =   2310
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "Text5"
            Top             =   585
            Width           =   2310
         End
         Begin VB.TextBox Text20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            TabIndex        =   53
            Text            =   "Text20"
            ToolTipText     =   "Nominal bayar"
            Top             =   2880
            Width           =   2310
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   195
            Left            =   180
            OleObjectBlob   =   "JL001.frx":06D1
            TabIndex        =   57
            Top             =   645
            Width           =   2295
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "JL001.frx":074D
            TabIndex        =   58
            Top             =   2505
            Width           =   1155
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   195
            Left            =   180
            OleObjectBlob   =   "JL001.frx":07B5
            TabIndex        =   59
            Top             =   1020
            Width           =   2295
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   195
            Left            =   180
            OleObjectBlob   =   "JL001.frx":082B
            TabIndex        =   60
            Top             =   1545
            Width           =   2025
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   240
            Left            =   180
            OleObjectBlob   =   "JL001.frx":0893
            TabIndex        =   61
            Top             =   2940
            Width           =   1650
         End
      End
   End
   Begin VB.Frame PENJUALAN 
      Height          =   4815
      Left            =   37
      TabIndex        =   24
      Top             =   1275
      Width           =   9720
      Begin VB.CommandButton cmdBL003 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9075
         TabIndex        =   6
         Top             =   1290
         Width           =   510
      End
      Begin VB.CommandButton Command4 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   4095
         TabIndex        =   2
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7725
         TabIndex        =   37
         Text            =   "18"
         Top             =   4425
         Width           =   1500
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6225
         TabIndex        =   36
         Text            =   "17"
         Top             =   4425
         Width           =   1500
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3750
         TabIndex        =   35
         Text            =   "8"
         Top             =   4425
         Width           =   465
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6300
         TabIndex        =   4
         Text            =   "Text7"
         ToolTipText     =   "Harga jual barang"
         Top             =   690
         Width           =   2310
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "JL001.frx":0903
         Left            =   1725
         List            =   "JL001.frx":0905
         TabIndex        =   0
         Text            =   "Combo1"
         ToolTipText     =   "Info barang tekan F1"
         Top             =   300
         Width           =   2310
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1725
         TabIndex        =   1
         Text            =   "Combo2"
         ToolTipText     =   "Info barang tekan F1"
         Top             =   930
         Width           =   2310
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6300
         TabIndex        =   3
         Text            =   "Text6"
         ToolTipText     =   "Jumlah barang"
         Top             =   315
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6300
         TabIndex        =   5
         Text            =   "Text1"
         ToolTipText     =   "Subdiskon"
         Top             =   1050
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6300
         TabIndex        =   17
         Text            =   "Text14"
         ToolTipText     =   "Harga setelah diskon"
         Top             =   1425
         Width           =   2310
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   165
         Left            =   75
         OleObjectBlob   =   "JL001.frx":0907
         TabIndex        =   27
         Top             =   375
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   165
         Left            =   75
         OleObjectBlob   =   "JL001.frx":097B
         TabIndex        =   28
         Top             =   1005
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   165
         Left            =   4635
         OleObjectBlob   =   "JL001.frx":09EF
         TabIndex        =   29
         Top             =   750
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   165
         Left            =   4635
         OleObjectBlob   =   "JL001.frx":0A65
         TabIndex        =   30
         Top             =   375
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   240
         Left            =   4635
         OleObjectBlob   =   "JL001.frx":0ACF
         TabIndex        =   31
         Top             =   1065
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   165
         Left            =   4635
         OleObjectBlob   =   "JL001.frx":0B45
         TabIndex        =   32
         Top             =   1440
         Width           =   1560
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   2595
         Left            =   195
         TabIndex        =   33
         ToolTipText     =   "Daftar penjualan barang"
         Top             =   1770
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   4577
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   16777152
         GridColor       =   0
         TextStyle       =   3
         TextStyleFixed  =   3
      End
      Begin VB.Frame Frame2 
         Caption         =   "DAFTAR BARANG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   4095
         TabIndex        =   39
         Top             =   300
         Width           =   5550
         Begin VB.CommandButton Command1 
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
            Height          =   420
            Left            =   1283
            TabIndex        =   40
            Top             =   2835
            Width           =   2985
         End
         Begin MSFlexGridLib.MSFlexGrid gridDF 
            Height          =   2490
            Left            =   45
            TabIndex        =   41
            Top             =   315
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   4392
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            BackColor       =   16777215
            BackColorFixed  =   65280
            BackColorBkg    =   16777152
            GridColor       =   0
            TextStyle       =   3
            TextStyleFixed  =   3
            Appearance      =   0
         End
      End
   End
End
Attribute VB_Name = "JL001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl As rdoResultset
Private SDel As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String
Private TTL

Private Sub cmdbatal_Click()
PENJUALAN.Visible = True
BAYAR.Visible = False
Combo1.SetFocus
Text14 = 0
End Sub

Private Sub cmdsimpan_Click()
Dim Tanya
Tanya = MsgBox("TRANSAKSI SELESAI", vbOKCancel, "KONFIRMASI")
    If Tanya = vbOK Then
        Call NoBukti2
        Call PERSEDIAAN_BAHAN
        If Text20 > 0 Then
            Call HISPiutang
        End If
        'If Frame3.Visible = True Then
        '    Call HISPiutang
        'End If
    ElseIf Tanya = vbCancel Then
        MsgBox "CANCEL", vbCritical, "KONFIRMASI"
    End If
Unload Me
JL001.Show 1
End Sub

Private Sub HISPiutang()
SSave = "Select * From H002"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
    RSave("NO_BUKTI") = SkinLabel16
    RSave("NO_PIUTANG") = Trim(Text12)
    RSave("NO_BG") = Trim(Text11)
    RSave("NO_PLG") = Combo3
    RSave("NAMA_PLG") = Trim(Text10)
    RSave("KETERANGAN") = Trim(Text19)
    RSave("PLAFON") = CCur(Text20)
    RSave("TGL_MULAI") = DTPicker2
    RSave("TGL_JATUH") = DTPicker3
    RSave("TGL_BAYAR") = DTPicker2
    RSave("TANGGAL") = Date
    RSave("USER_CODE") = Operator
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub cmdtutup_Click()
Unload Me
End Sub

Private Sub DelJL001()
SDel = "Delete * From JL001"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo3_LostFocus()
If Combo3 = "" Then Exit Sub
SCari3 = "Select * From P001 where Kode ='" + Trim(Combo3) + "'"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenDynamic, rdConcurRowVer)
If RCari3.RowCount <> 0 Then
    Text10 = RCari3("Nama")
End If
RCari3.Close
Set RCari3 = Nothing
End Sub

Private Sub Command1_Click()
Frame2.Visible = False
Combo1.SetFocus
End Sub

Private Sub Command4_Click()
Frame2.Visible = True
Frame2.ZOrder
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=KASIR", rdDriverNoPrompt, False, CN)
Frame2.Visible = False

PENJUALAN.Visible = True
BAYAR.Visible = False

Call DelJL001
Call KOSONG
Call NoBukti
Call IsiCombo
Call IsiText
Call SiapkanGrid
Call IsiGrid
Text15 = 1

Call SiapkanGridDF
Call IsiGridDF

'Frame3.Visible = False
Frame4.Visible = False

    SSPL = "Select * From P001 order by kode"
    Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
    If RSPL.RowCount <> 0 Then
        RSPL.MoveFirst
        Do While Not RSPL.EOF
            Combo3.AddItem RSPL("Kode")
        RSPL.MoveNext
        Loop
        Combo3.ListIndex = 0
    Else
        MsgBox "DATA PELANGGAN MASIH KOSONG", vbSystemModal, "KONFIRMASI"
    End If
        RSPL.Close
        Set RSPL = Nothing

DTPicker2 = Date
DTPicker3 = DateAdd("m", 1, Date)
Text3 = 0
End Sub

Private Sub SiapkanGridDF()
With gridDF
    .Cols = 2
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2700: .Text = "NAMA": .CellAlignment = 4
End With
End Sub

Private Sub IsiGridDF()
SKTG = "Select * From B003 order by KodeBR Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGridDF
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      gridDF.Rows = B + 1
      gridDF.Row = B
         With gridDF
              .Col = 0: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("NamaBR")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub KOSONG()
ClearTextBoxes JL001
Combo1 = ""
Combo2 = ""
End Sub

Private Sub IsiText()
Text7 = 0
Text1 = 0
Text14 = 0
End Sub

Private Sub NoBukti()
Dim No As Double
SqlNo = "Select * from C013 where nama = '" + Operator + "'"
Set RSLNO = RDCO.OpenResultset(SqlNo, rdOpenDynamic, rdConcurRowVer)
No = Val(RSLNO("NoJual")) + 1
NoStr = Digit(7, No)
SkinLabel16 = "1." + NoStr
RSLNO.Close
Set RSLNO = Nothing
End Sub

Private Sub NoBukti2()
SCari9 = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RCari9 = RDCO.OpenResultset(SCari9, rdOpenKeyset, rdConcurRowVer)
    TOGEL = RCari9("NoJual")
    RCari9.Edit
        RCari9("NoJual") = TOGEL + 1
RCari9.Update
RCari9.Close
Set RCari9 = Nothing
End Sub


Private Sub IsiCombo()
SKTG = "Select * From B003 order by kodebr"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenDynamic, rdOpenKeyset)
RKTG.MoveFirst
Do While Not RKTG.EOF
    Combo1.AddItem RKTG("KodeBR")
RKTG.MoveNext
Loop
RKTG.Close
Set RKTG = Nothing

SSTN = "Select * From B003 order by namabr"
Set RSTN = RDCO.OpenResultset(SSTN, rdOpenDynamic, rdOpenKeyset)
RSTN.MoveFirst
Do While Not RSTN.EOF
    Combo2.AddItem RSTN("NamaBR")
RSTN.MoveNext
Loop
RSTN.Close
Set RSTN = Nothing

End Sub

Private Sub Combo1_GotFocus()
Combo1.BackColor = RGB(255, 255, 0)
End Sub

'Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'    Case vbKeyF5
'        If Text3 = "" Then
'            MsgBox "TRANSAKSI MASIH KOSONG", vbCritical, "KONFIRMASI"
'            PENJUALAN.Visible = True
'            BAYAR.Visible = False
'            Combo1.SetFocus
'        Else
'            PENJUALAN.Visible = False
'            BAYAR.Visible = True
'            Text4.SetFocus
'            Exit Sub
'        End If
'    Case vbKeyF1
'        If Combo1 <> "" Then Combo1 = ""
'        KodeBR = ""
'        NamaBR = ""
'        IB01.Show 1
'        If KodeBR = "" Or NamaBR = "" Then Exit Sub
'        Combo1 = KodeBR
'        Call Combo1_LostFocus
'End Select
'End Sub

Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        If Text3 = 0 Then
            MsgBox "TIDAK ADA TRANSAKSI", vbCritical, "KONFIRMASI"
            Exit Sub
        Else
            PENJUALAN.Visible = False
            BAYAR.Visible = True
            Text4.SetFocus
        End If
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
Combo1.BackColor = RGB(255, 255, 255)

If Combo1 = "" Then Exit Sub
SCari = "Select * From B003 where KodeBR='" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Combo2 = RCari("NamaBR")
    Text7 = Format(RCari("HJUAL"), "##,###.00")
    Text6.SetFocus
Else
    MsgBox "KODE BARANG BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo1 = ""
    Combo2 = ""
    Combo1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Combo2_GotFocus()
Combo2.BackColor = RGB(255, 255, 0)
End Sub

Private Sub combo2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        If Text3 = 0 Then
            MsgBox "TIDAK ADA TRANSAKSI", vbCritical, "KONFIRMASI"
            Exit Sub
        Else
            PENJUALAN.Visible = False
            BAYAR.Visible = True
            Text4.SetFocus
        End If
End Select
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo2_LostFocus()
Combo2.BackColor = RGB(255, 255, 255)

If Combo2 = "" Then Exit Sub
SCari2 = "Select * From B003 where NamaBR='" + Trim(Combo2) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Combo1 = RCari2("KodeBR")
    Text7 = Format(RCari2("HJUAL"), "##,###.00")
    Text6.SetFocus
Else
    MsgBox "NAMA BARANG BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo1 = ""
    Combo2 = ""
    Combo1.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
Combo2 = Format(Combo2, ">")
End Sub

Private Sub gridDF_dblClick()
Combo1 = (gridDF.TextMatrix(gridDF.Row, 0))
Combo2 = (gridDF.TextMatrix(gridDF.Row, 1))
Frame2.Visible = False
Text6.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        If Text3 = 0 Then
            MsgBox "TIDAK ADA TRANSAKSI", vbCritical, "KONFIRMASI"
            Exit Sub
        Else
            PENJUALAN.Visible = False
            BAYAR.Visible = True
            Text4.SetFocus
        End If
End Select
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = RGB(255, 255, 0)
Text1 = ""
End Sub

Private Sub Text1_Lostfocus()
Text1.BackColor = RGB(255, 255, 255)
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Text1 = "" Then
        Text1 = 0
        Text14 = Format((CCur((Text6) * (Text7)) - (CCur((Text6) * (Text7) * (Text1) / 100))), "##,###.00")
    Else
        Text14 = Format((CCur((Text6) * (Text7)) - (CCur((Text6) * (Text7) * (Text1) / 100))), "##,###.00")
        SendKeys "{TAB}"
    End If
End If
End Sub

Private Sub cmdBL003_GotFocus()
Text14.BackColor = RGB(255, 255, 0)
End Sub

Private Sub cmdBL003_LostFocus()
Text14.BackColor = RGB(255, 255, 255)
End Sub

Private Sub cmdBL003_Click()

SJS = "Select * From B003 where KodeBR = '" + Trim(Combo1) + "'"
Set RJS = RDCO.OpenResultset(SJS, rdOpenKeyset, rdConcurRowVer)

TTL = CCur(RJS("JAkhir"))

If CCur(Text6) > RJS("JAkhir") Then
    MsgBox "JUMLAH STOCK BARANG TERSEDIA " + Trim(RJS("JAkhir")) + " PCS", vbCritical, "KONFIRMASI"
    Text6 = ""
    Text6.SetFocus
    Exit Sub

Else

Dim Tanya
    If Combo1 = "" Or Combo2 = "" Or Text6 = "" Or Text7 = "" Or Text1 = "" Or Text14 = ",00" Then
        MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
        Combo1.SetFocus
        Exit Sub
    Else
        Tanya = MsgBox("MASUKAN DATA PENJUALAN", vbCritical, "KONFIRMASI")
        If Tanya = vbOK Then
            
            Call SimpanJL001
                SJual10 = "Select * From JL001A"
                Set RJual10 = RDCO.OpenResultset(SJual10, rdOpenKeyset, rdConcurReadOnly)
                    Text8 = RJual10("SumOfJML_BAHAN")
                    Text17 = Format(RJual10("SumOfTOTAL_JUAL"), "##,###.00")
                    Text18 = Format(RJual10("SumOfTOTAL_DISCOUNT"), "##,###.00")
                RJual10.Close
                Set RJual10 = Nothing
            Call SiapkanGrid
            Call IsiGrid
            Call IsiText
            Call TOTAL
                Text6 = ""
                Combo1 = ""
                Combo2 = ""
                Combo1.SetFocus
            Exit Sub
        End If
    End If
End If

RJS.Close
Set RJS = Nothing

End Sub

Private Sub TOTAL()
SCari3 = "Select * From JL001A"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurReadOnly)
If RCari3.RowCount <> 0 Then
    TOTALP = RCari3("sumofTOTAL_JUAL")
    TOTALD = RCari3("SumOfNOMDISC")
    Text3 = Format(TOTALP, "##,###.00")
Else
    Text3 = 0
End If

RCari3.Close
Set RCari3 = Nothing
Text2.Text = "TOTAL BAYAR"
End Sub

Private Sub SimpanJL001()
SCari6 = "Select * From JL001 where Kode_JNS = '" + Trim(Combo1) + "'"
Set RCari6 = RDCO.OpenResultset(SCari6, rdOpenDynamic, rdConcurRowVer)
If RCari6.RowCount <> 0 Then
    JUMLAH = RCari6("JML_BAHAN") + CCur(Text6)
    HARGA = CCur(Text7) * JUMLAH
    NNOMDISC = CCur(Text1) / 100 * CCur(Text7)
    
'CEK JUMLAH BARANG B003 DENGAN JL001

    If CCur(TTL) < CCur(JUMLAH) Then
        MsgBox "JUMLAH STOCK BARANG TERSEDIA " + Trim(RJS("JAkhir")) + " PCS", vbCritical, "KONFIRMASI"
        Text6.SetFocus
        Exit Sub
    End If
    
    RCari6.Edit
    RCari6("JML_BAHAN") = CCur(JUMLAH)
    RCari6("HJUAL_PCS") = CCur(Text7)
    RCari6("HARGA_JUAL") = CCur(HARGA)
    RCari6("DISCOUNT") = CCur(Text1)
    RCari6("NOMINAL") = 0
    RCari6("NOMDISC") = CCur(NNOMDISC)
    RCari6("TOTAL_JUAL") = CCur(Text7) * CCur(JUMLAH)
    RCari6("TOTAL_DISCOUNT") = CCur(HARGA) - CCur(NNOMDISC) * CCur(JUMLAH)
    RCari6.Update
    RCari6.Close
    Set RCari6 = Nothing

Else

    SCari5 = "Select * From B003 where KodeBR = '" + Trim(Combo1) + "'"
    Set RCari5 = RDCO.OpenResultset(SCari5, rdOpenDynamic, rdConcurRowVer)
        INDUK = RCari5("KodeInd")
        HJUAL = RCari5("HJual")
        
        SSave = "Select * From JL001"
        Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
        RSave.AddNew
            RSave("No_Trans") = SkinLabel16
            RSave("No_Urut") = Text15
            RSave("Kode_Ind") = INDUK
            RSave("Kode_JNS") = Combo1
            RSave("Nama_JNS") = Combo2
            RSave("Jml_BAHAN") = CCur(Text6)
            RSave("HJual_PCS") = CCur(Text7)
            RSave("Harga_JUAL") = HJUAL * CCur(Text6)
            RSave("Nominal") = 0
            RSave("Laba") = 0
            RSave("User_Code") = Operator
            RSave("Discount") = CCur(Text1)
            RSave("NomDisc") = CCur((Text7) * (Text1) / 100)
            RSave("TOTAL_JUAL") = CCur(Text7) * CCur(Text6)
            RSave("TOTAL_DISCOUNT") = (CCur(Text7) - (CCur((Text7) * (Text1) / 100))) * CCur(Text6)
        RSave.Update
        RSave.Close
        Set RSave = Nothing
        Text15 = Text15 + 1
    RCari5.Close
    Set RCari5 = Nothing

End If
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 8
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2000: .Text = "NAMA BARANG": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 500: .Text = "JML": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1500: .Text = "HARGA PCS": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 500: .Text = "%": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 1500: .Text = "JUMLAH HARGA": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 1500: .Text = "HARGA BERSIH": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari4 = "Select * From JL001"
Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurReadOnly)
If RCari4.RowCount <> 0 Then
   RCari4.MoveFirst
   B = 1
   Do Until RCari4.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari4("No_Urut"): .CellAlignment = 4
              .Col = 1: .Text = RCari4("Kode_JNS"): .CellAlignment = 4
              .Col = 2: .Text = RCari4("Nama_JNS")
              .Col = 3: .Text = RCari4("Jml_BAHAN"): .CellAlignment = 4
              .Col = 4: .Text = Format(RCari4("HJual_PCS"), "##,###.00")
              .Col = 5: .Text = RCari4("Discount"): .CellAlignment = 4
              .Col = 6: .Text = Format(RCari4("TOTAL_JUAL"), "##,###.00")
              .Col = 7: .Text = Format((RCari4("HJUAL_PCS") - RCari4("NomDisc")) * RCari4("Jml_BAHAN"), "##,###.00"): .CellFontBold = True
         End With
      B = B + 1
      RCari4.MoveNext
   Loop
End If
RCari4.Close
Set RCari4 = Nothing

End Sub

Private Sub IsiHarga()
If Combo1 = "" Or Combo2 = "" Then
    Text7.SetFocus
Else
    SHarga = "Select * From B003 where KodeBR ='" + Trim(Combo1) + "'"
    Set RHarga = RDCO.OpenResultset(SHarga, rdOpenDynamic, rdConcurRowVer)
    If RHarga.RowCount <> 0 Then
        Text7 = RHarga("HJual")
    End If
    RHarga.Close
    Set RHarga = Nothing
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text11_LostFocus()
Text11 = Format(Text11, ">")
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text12_LostFocus()
Dim Tanya
If Text12 = "" Then Exit Sub
SKode = "Select * From H002 where NO_PIUTANG = '" + Text12 + "'"
Set RKode = RDCO.OpenResultset(SKode, rdOpenDynamic, rdOpenKeyset)
If RKode.RowCount <> 0 Then
    Tanya = MsgBox("NOMOR PIUTANG TELAH TERDAFTAR", vbCritical, "KONFIRMASI")
    Text12 = ""
    Text12.SetFocus
End If
RKode.Close
Set RKode = Nothing
Text12 = Format(Text12, ">")
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsimpan.SetFocus
End If
End Sub

Private Sub Text19_LostFocus()
Text19 = Format(Text19, ">")
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        If Text3 = 0 Then
            MsgBox "TIDAK ADA TRANSAKSI", vbCritical, "KONFIRMASI"
            Exit Sub
        Else
            PENJUALAN.Visible = False
            BAYAR.Visible = True
            Text4.SetFocus
        End If
End Select
End Sub

Private Sub text6_GotFocus()
Text6.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text6_LostFocus()
'Call IsiHarga
Text6.BackColor = RGB(255, 255, 255)
If Not IsNumeric(Text6) Then
    Text6 = 0
    Exit Sub
Else
    Text14 = Format(CCur(Text6) * CCur(Text7), "##,###.00")
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        If Text3 = 0 Then
            MsgBox "TIDAK ADA TRANSAKSI", vbCritical, "KONFIRMASI"
            Exit Sub
        Else
            PENJUALAN.Visible = False
            BAYAR.Visible = True
            Text4.SetFocus
        End If
End Select
End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text7_LostFocus()
Text7.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    Text7 = Format(CCur(Text7), "##,###.00")
    Text14 = Format(CCur(Text6) * CCur(Text7), "##,###.00")
    SendKeys "{TAB}"
End If
End Sub


'Private Sub text4_gotfocus()
'Text3 = Format(CCur(Text16), "##,###.00")
'Text2.Text = "TOTAL BAYAR"
'End Sub

Private Sub Text4_GotFocus()
Text4 = ""
Text20 = 0
'Frame3.Visible = False
Frame4.Visible = False


SSave3 = "Select * From JL001A"
Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenKeyset, rdConcurReadOnly)
    Text5 = Format(RSave3("sumofTOTAL_JUAL"), "##,###.00")
    Text9 = Format(RSave3("SumOfTOTAL_JUAL") - RSave3("SumOfTOTAL_DISCOUNT"), "##,###.00")
    Text10 = Format(CCur(Text5) - CCur(Text9), "##,###.00")
    Text16 = Format(CCur(Text5) - CCur(Text9), "##,###.00")
RSave3.Close
Set RSave3 = Nothing
Text3 = Format(CCur(Text16), "##,###.00")
Text2.Text = "TOTAL BAYAR"
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Text4 = "" Then
        Text4.SetFocus
    Else
        Text4 = Format(CCur(Text4), "##,###.00")
        Combo3.SetFocus
        'cmdsimpan.SetFocus
    End If
End If
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then
    Text4 = 0
    Exit Sub
End If

If CCur(Text4) < CCur(Text16) Then
    Text20 = Format(CCur(Text16) - CCur(Text), "##,###.00")
    'Frame3.Visible = True
    Frame4.Visible = True
    Combo3.SetFocus
    Text10 = "<No. Name>"
ElseIf CCur(Text4) > CCur(Text16) Then
    Text2.Text = "KEMBALI"
    Text20 = 0
    Text3 = Format(CCur(Text4) - CCur(Text16), "##,###.00")
    'Frame3.Visible = False
    Frame4.Visible = False
    cmdsimpan.SetFocus
ElseIf CCur(Text4) = CCur(Text16) Then
    Text20 = 0
    'Frame3.Visible = False
    Frame4.Visible = False
    cmdsimpan.SetFocus
End If
End Sub

Private Sub text13_gotfocus()
If Text4 < Text16 Then
    MsgBox "NOMINAL PEMBAYARAN KURANG", vbCritical, "KONFIRMASI"
    Text4.SetFocus
Else
    cmdsimpan.SetFocus
End If
End Sub

Private Sub PERSEDIAAN_BAHAN()

SJual4 = "Select * From JL001 where NO_TRANS = '" + Trim(SkinLabel16) + "' ORDER BY NO_URUT"
Set RJual4 = RDCO.OpenResultset(SJual4, rdOpenKeyset, rdConcurRowVer)
RJual4.MoveFirst
Do While Not RJual4.EOF
    NOURUT = RJual4("NO_URUT")
    KODEJNS = RJual4("KODE_JNS")
    NAMAJNS = RJual4("NAMA_JNS")
    JMLBAHAN = RJual4("JML_BAHAN")
    HJUALPCS = RJual4("HJUAL_PCS")
    HARGAJUAL = RJual4("HARGA_JUAL")
    NNOMDISC = RJual4("NOMDISC") * RJual4("JML_BAHAN")
    
'EDIT MUTASIPRODUKSI BERDASARKAN METODE STOCK'
    SJual5 = "Select * From B003 where KODEBR = '" + Trim(KODEJNS) + "'"
    Set RJual5 = RDCO.OpenResultset(SJual5, rdOpenKeyset, rdConcurRowVer)
        JMLBAHAN1 = CCur(JMLBAHAN)
        NOMINAL = 0
        NOMINAL1 = 0
        HPOKOK = 0
        SDISC = 0
        MTSTOCK = RJual5("MSTOCK")

        SJual6 = "Select * From B004 where KODE_JNS = '" + Trim(KODEJNS) + "' ORDER BY NO_URUT"
        Set RJual6 = RDCO.OpenResultset(SJual6, rdOpenKeyset, rdConcurRowVer)
        RJual6.MoveFirst
        Do While Not RJual6.EOF
            NO4 = RJual6("NO_URUT")
            HBELIPCS = RJual6("HBELI_PCS")
            JMLSALDO = RJual6("JML_SALDO")
            NOMSALDO = RJual6("NOM_SALDO")
            HPPCS = RJual6("HJUAL_PCS")
            
            If JMLBAHAN1 >= JMLSALDO Then
            JMLBAHAN1 = JMLBAHAN1 - JMLSALDO
            NOMINAL1 = NOMINAL1 + NOMSALDO
            HPOKOK = HPOKOK + SALDOHP
            
                RJual6.Edit
                RJual6("JML_SALDO") = 0
                RJual6("NOM_SALDO") = 0
                RJual6.Update
                
            ElseIf JMLBAHAN1 < JMLSALDO And JMLBAHAN1 <> 0 Then
            JMLSALDO = JMLSALDO - JMLBAHAN1
            NOMSALDO = NOMSALDO - (HBELIPCS * JMLBAHAN1)
            NOMINAL1 = NOMINAL1 + (HBELIPCS * JMLBAHAN1)
            SALDOHP = SALDOHP - (HPPCS * JMLBAHAN1)
            HPOKOK = HPOKOK + (HPPCS * JMLBAHAN1)
                
                RJual6.Edit
                RJual6("JML_SALDO") = CCur(JMLSALDO)
                RJual6("NOM_SALDO") = CCur(NOMSALDO)
                RJual6.Update
                
                JMLBAHAN1 = 0
            End If
            
            SSDISC = CCur(SPCDISC) * CCur(JMLBAHAN)
            
            JMLSTOCK = CCur(RJual5("JAKHIR"))
    
                If JMLBAHAN > JMLSTOCK And JMLSTOCK > 0 Then
                JMLBHNLEBIH = CCur(JMLBAHAN) - CCur(JMLSTOCK)
                NOMLEBIH = CCur(HJUALPCS) * CCur(JMLBHNLEBIH)
                NOMINAL1 = CCur(NOMINAL1) + CCur(NOMLEBIH)
                End If
    
            LABA = RJual4("HARGA_JUAL") - (NOMINAL1 + NNOMDISC + SSDISC)
            RJual4.Edit
            RJual4("NOMINAL") = CCur(NOMINAL1)
            RJual4("LABA") = CCur(LABA)
            RJual4.Update
        
        RJual6.MoveNext
        Loop
        RJual6.Close
        Set RJual6 = Nothing
        
                    SJual7 = "Delete * From B004 where JML_SALDO = 0 AND NOM_SALDO < 1"
                    Set RJual7 = RDCO.OpenResultset(SJual7, rdOpenDynamic, rdConcurRowVer)
                    RJual7.Close
                    Set RJual7 = Nothing
    
    JMLCRD = RJual5("JC")
    JMLAKHIR = RJual5("JAKHIR")
    MUTASICRT = RJual5("MUTASIC")
    SALDOAKHIR = RJual5("SALDO")

                If SALDOAKHIR <= 0 Then
                SDISC = SPCDISC * JMLBAHAN
                HARGAJUAL = RJual5("HJUAL")
                NNOMDISC = RJual4("NOMDISC")
                NOMINAL1 = (HARGAJUAL - NOMDISC) - SDISC
                LABA = 0
                End If
        
        JMLCRD = JMLCRD + JMLBAHAN
        JMLAKHIR = JMLAKHIR - JMLBAHAN
        MUTASICRT = MUTASICRT + NOMINAL1
        SALDOAKHIR = SALDOAKHIR - NOMINAL1
        
    RJual5.Edit
    RJual5("JC") = CCur(JMLCRD)
    RJual5("JAKHIR") = CCur(JMLAKHIR)
    RJual5("MUTASIC") = CCur(MUTASICRT)
    RJual5("SALDO") = CCur(SALDOAKHIR)
        
'UPDATE HISTORY TRANSAKSI BAHAN BAKU'
                    SJual8 = "Select * From B005 ORDER BY NO_URUT"
                    Set RJual8 = RDCO.OpenResultset(SJual8, rdOpenKeyset, rdConcurRowVer)
                    RJual8.AddNew
                        RJual8("Status") = 1
                        RJual8("KODE_TRANS") = "JL"
                        RJual8("KODE_JNS") = KODEJNS
                        RJual8("NAMA_JNS") = NAMAJNS
                        RJual8("NO_FAKTUR") = SkinLabel16
                        RJual8("NO_BUKTI") = SkinLabel16
                        RJual8("KETERANGAN") = "PENJUALAN NO." + RJual4("NO_TRANS")
                        RJual8("JML_DBT") = 0
                        RJual8("JML_CRD") = JMLBAHAN
                        RJual8("JML_AKHIR") = JMLAKHIR
                        RJual8("MUTASI_DBT") = 0
                        RJual8("MUTASI_CRT") = NOMINAL1
                        RJual8("SALDO_AKHIR") = SALDOAKHIR
                        RJual8("H_POKOK") = HARGAJUAL
                        RJual8("NOMDISC") = NNOMDISC
                        RJual8("SPCDISC") = SDISC
                        RJual8("LABA") = LABA
                        RJual8("KAS") = 0
                        RJual8("TGL_S") = TglS
                        RJual8("TGL_FAK") = TglS
                        RJual8("KODE_SPL") = Trim(Combo3)
                    RJual8.Update
                    RJual8.Close
                    Set RJual8 = Nothing
   
    RJual5.Update
    RJual5.Close
    Set RJual5 = Nothing

RJual4.MoveNext
Loop
RJual4.Close
Set RJual4 = Nothing
        
End Sub

