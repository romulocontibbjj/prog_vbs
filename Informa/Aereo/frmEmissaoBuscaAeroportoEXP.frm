VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoBuscaAeroportoEXP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Especifique seu Aeroporto"
   ClientHeight    =   2115
   ClientLeft      =   1725
   ClientTop       =   2520
   ClientWidth     =   8355
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoBuscaAeroportoEXP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   1875
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   3307
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmEmissaoBuscaAeroportoEXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xRs As Recordset


Private Sub Form_Load()
Dim X As Integer
    If de_informa.rsSelAeroportoSigla.State = 1 Then
    Set xRs = de_informa.rsSelAeroportoSigla
    ElseIf de_informa.rsSelAeroportoCidade.State = 1 Then
    Set xRs = de_informa.rsSelAeroportoCidade
    End If

Flex.Clear
Flex.Cols = 4
Flex.Rows = xRs.RecordCount + 1
Flex.FixedRows = 1
Flex.FixedCols = 0
Flex.TextMatrix(0, 0) = "Sigla"
Flex.TextMatrix(0, 1) = "UF"
Flex.TextMatrix(0, 2) = "Cidade"
Flex.TextMatrix(0, 3) = "Aeroporto"
Flex.ColWidth(0) = 800
Flex.ColWidth(1) = 500
Flex.ColWidth(2) = 2500
Flex.ColWidth(3) = 4000
X = 1
    Do Until xRs.EOF
    Flex.TextMatrix(X, 0) = xRs.Fields("sigla")
    Flex.TextMatrix(X, 1) = xRs.Fields("uf")
    Flex.TextMatrix(X, 2) = PriMaiuscula(xRs.Fields("localidade"))
    Flex.TextMatrix(X, 3) = PriMaiuscula(xRs.Fields("aeroporto"))
    X = X + 1
    xRs.MoveNext
    Loop

End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)
Dim X As Integer

X = Flex.Row
    If KeyAscii = 27 Then
    Unload Me
    ElseIf KeyAscii = 13 Then
    frmEmissao.TxtSiglaExpedidor = Flex.TextMatrix(X, 0)
    frmEmissao.TxtAeroportoExpedidor = Flex.TextMatrix(X, 2) & " - " & Flex.TextMatrix(X, 1) & " (" & Flex.TextMatrix(X, 3) & ")"
    Unload Me
    End If
End Sub
