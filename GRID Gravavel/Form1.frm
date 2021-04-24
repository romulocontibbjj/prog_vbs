VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   5400
      TabIndex        =   1
      Top             =   150
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Estou ultilizando o microsoft flexgrid control 6.0

Public Function MontaGrid()
  With grid
    .Cols = 3
    .Rows = 2
    .FixedCols = 0
    .FixedRows = 1
    .TextMatrix(0, 0) = "Código"
    .TextMatrix(0, 1) = "Nome"
    .TextMatrix(0, 2) = "Telefone"
    .ColWidth(0) = 700
    .ColWidth(1) = 2000
    .ColWidth(2) = 1000
  End With
End Function

Private Sub Form_Load()
  MontaGrid
End Sub

Private Sub grid_Click()
  Call grid_EnterCell
End Sub

Private Sub grid_DblClick()
  Call grid_EnterCell
End Sub

Private Sub grid_EnterCell()
  With Text1
    .Visible = True
    .Top = grid.Top + grid.CellTop
    .Left = grid.Left + grid.CellLeft
    .Width = grid.CellWidth
    .Height = grid.CellHeight
    .Text = UCase(grid.Text)
    .Visible = True
    .SetFocus
  End With
End Sub

Private Sub grid_GotFocus()
  Call grid_EnterCell
End Sub

Private Sub grid_LeaveCell()
  With grid
    If .MouseCol <> 7 Then
      Text1.Visible = False
      .TextMatrix(.Row, .Col) = UCase(Text1.Text)
    End If
  End With
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  With grid
    If .Col < .Cols - 1 Then
      If KeyCode = 13 Or KeyCode = 9 Then
        .Col = .Col + 1
        Call grid_EnterCell
      Else
        Exit Sub
      End If
    Else
      If KeyCode = 13 Or KeyCode = 9 Then
        If .Row = .Rows - 1 Then
          .AddItem ""
          .Col = 0
          .Row = .Rows - 1
          Call grid_EnterCell
        Else
          .Row = .Row + 1
          .Col = 0
          Call grid_EnterCell
        End If
      End If
    End If
  End With
End Sub
