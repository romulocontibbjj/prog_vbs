VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoCODSIATA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Códigos IATA"
   ClientHeight    =   4950
   ClientLeft      =   2580
   ClientTop       =   2100
   ClientWidth     =   6615
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoCODSIATA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCAD 
      Caption         =   "Cadastrar Novos Produtos"
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Top             =   105
      Width           =   2175
   End
   Begin VB.TextBox TxtBusca 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin MSFlexGridLib.MSFlexGrid FlexCODS 
      Height          =   4275
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7541
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmEmissaoCODSIATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCAD_Click()
frmCadProdINT.Show 1
Call Form_Load
End Sub

Private Sub CmdCAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub FlexCODS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    frmEmissao.TxtDescrIATA.Text = FlexCODS.TextMatrix(FlexCODS.Row, 1) & " - " & FlexCODS.TextMatrix(FlexCODS.Row, 2)
    Unload Me
    ElseIf KeyAscii = 27 Then
    Unload Me
    End If

End Sub

Private Sub Form_Load()

    If de_informa.rsSel_CadProdINT.State = 1 Then de_informa.rsSel_CadProdINT.Close
    de_informa.Sel_CadProdINT "%"
    
FlexCODS.Clear
FlexCODS.Rows = de_informa.rsSel_CadProdINT.RecordCount + 1
FlexCODS.Cols = 3
FlexCODS.FixedCols = 1
FlexCODS.FixedRows = 1

FlexCODS.TextMatrix(0, 1) = "Cod."
FlexCODS.TextMatrix(0, 2) = "Descrição"

Y = 0

    Do Until de_informa.rsSel_CadProdINT.EOF
    Y = Y + 1
    FlexCODS.TextMatrix(Y, 1) = de_informa.rsSel_CadProdINT.Fields("codigo")
    FlexCODS.TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CadProdINT.Fields("descricao"))
    de_informa.rsSel_CadProdINT.MoveNext
    Loop

FlexCODS.ColWidth(0) = 200
FlexCODS.ColWidth(1) = 500
FlexCODS.ColWidth(2) = 4000
End Sub

Private Sub TxtBusca_Change()
TxtBusca.Text = UCase(TxtBusca.Text)
TxtBusca.SelStart = Len(TxtBusca.Text)
End Sub

Private Sub TxtBusca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
ElseIf KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub TxtBusca_LostFocus()
If Len(Trim(TxtBusca.Text)) > 0 Then
    For Y = 1 To FlexCODS.Rows - 1
        If InStr(1, FlexCODS.TextMatrix(Y, 2), TxtBusca.Text, vbTextCompare) > 0 Then
        FlexCODS.TopRow = Y
        FlexCODS.Row = Y
        FlexCODS.Col = 1
        FlexCODS.SetFocus
        SendKeys "{DOWN}"
        SendKeys "{UP}"
        Y = FlexCODS.Rows + 1
        End If
    Next
End If
End Sub
