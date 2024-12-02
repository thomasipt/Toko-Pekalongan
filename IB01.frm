VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form IB01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORMASI GUDANG / TOKO (Double Click For Change)"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "IB01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   7515
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Klik untuk input barang"
      Top             =   68
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   13256
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "IB01.frx":000C
      Top             =   6720
   End
End
Attribute VB_Name = "IB01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private RKTG As rdoResultset
Private SKTG As String


Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=KASIR", rdDriverNoPrompt, False, CN)
Call SiapkanGrid
Call IsiGrid
End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 2
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2700: .Text = "NAMA": .CellAlignment = 4
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

Private Sub Grid_dblClick()
KodeBR = grid.TextMatrix(grid.Row, 0)
NamaBR = grid.TextMatrix(grid.Row, 1)
Unload Me
End Sub
