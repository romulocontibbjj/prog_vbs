VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoFiltraEXP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Especifique seu Cliente"
   ClientHeight    =   4395
   ClientLeft      =   1725
   ClientTop       =   1185
   ClientWidth     =   9135
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoFiltraEXP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7329
      _Version        =   393216
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmEmissaoFiltraEXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xRs As Recordset

Private Sub Flex_KeyPress(KeyAscii As Integer)
Dim X As Integer

X = Flex.Row
    If KeyAscii = 27 Then
    Unload Me
    ElseIf KeyAscii = 13 Then
        With frmEmissao
        .TxtCGCExpedidor = Flex.TextMatrix(X, 0)
        .TxtNomeExpedidor = Flex.TextMatrix(X, 1)
        .TxtCidadeExpedidor = Flex.TextMatrix(X, 2)
        .TxtUFExpedidor = Flex.TextMatrix(X, 3)
        .TxtEndExpedidor.Text = Flex.TextMatrix(X, 4)
        .TxtCEPExpedidor.Text = Flex.TextMatrix(X, 5)
        .TxtTelExpedidor.Text = Flex.TextMatrix(X, 6)
        .TxtFAXExpedidor.Text = Flex.TextMatrix(X, 7)
        .TxtSeguradoraExpedidor.Text = Flex.TextMatrix(X, 8)
        .TxtApoliceExpedidor.Text = Flex.TextMatrix(X, 9)
        .TxtInscrEstExpedidor.Text = Flex.TextMatrix(X, 10)
        End With
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
Dim X As Integer
    If de_informa.rsSelClienteAPELIDO.State = 1 Then
    Set xRs = de_informa.rsSelClienteAPELIDO
    ElseIf de_informa.rsSelClienteCNPJ.State = 1 Then
    Set xRs = de_informa.rsSelClienteCNPJ
    ElseIf de_informa.rsSelClienteFANTASIA.State = 1 Then
    Set xRs = de_informa.rsSelClienteFANTASIA
    ElseIf de_informa.rsSelClienteNOME.State = 1 Then
    Set xRs = de_informa.rsSelClienteNOME
    End If

Flex.Clear
Flex.Cols = 13
Flex.Rows = xRs.RecordCount + 1
Flex.FixedRows = 1
Flex.FixedCols = 0
Flex.TextMatrix(0, 0) = "CNPJ"
Flex.TextMatrix(0, 1) = "Nome"
Flex.TextMatrix(0, 2) = "Cidade"
Flex.TextMatrix(0, 3) = "UF"
Flex.TextMatrix(0, 4) = "Endereço"
Flex.TextMatrix(0, 5) = "CEP"
Flex.TextMatrix(0, 6) = "Tel."
Flex.TextMatrix(0, 7) = "Fax"
Flex.TextMatrix(0, 8) = "Seguradora"
Flex.TextMatrix(0, 9) = "Apólice"
Flex.TextMatrix(0, 10) = "Inscr. Est."
Flex.TextMatrix(0, 11) = "Nome Fantasia"
Flex.TextMatrix(0, 12) = "Apelido"

Flex.ColWidth(0) = 1500
Flex.ColWidth(1) = 3100
Flex.ColWidth(2) = 2500
Flex.ColWidth(3) = 500
Flex.ColWidth(4) = 4000
Flex.ColWidth(5) = 1500
Flex.ColWidth(6) = 1500
Flex.ColWidth(7) = 1500
Flex.ColWidth(8) = 2500
Flex.ColWidth(9) = 1500
Flex.ColWidth(10) = 1500
Flex.ColWidth(11) = 1500
Flex.ColWidth(12) = 1500

X = 1

    Do Until xRs.EOF
    If IsNull(xRs.Fields("cgc")) = False Then Flex.TextMatrix(X, 0) = xRs.Fields("cgc")
    If IsNull(xRs.Fields("nome")) = False Then Flex.TextMatrix(X, 1) = PriMaiuscula(xRs.Fields("nome"))
    If IsNull(xRs.Fields("cidade")) = False Then Flex.TextMatrix(X, 2) = PriMaiuscula(xRs.Fields("cidade"))
    If IsNull(xRs.Fields("uf")) = False Then Flex.TextMatrix(X, 3) = xRs.Fields("uf")
    If IsNull(xRs.Fields("endereco")) = False Then Flex.TextMatrix(X, 4) = PriMaiuscula(xRs.Fields("endereco"))
    If IsNull(xRs.Fields("cep")) = False Then Flex.TextMatrix(X, 5) = xRs.Fields("cep")
    If IsNull(xRs.Fields("pabx")) = False Then Flex.TextMatrix(X, 6) = xRs.Fields("pabx")
    If IsNull(xRs.Fields("fax")) = False Then Flex.TextMatrix(X, 7) = xRs.Fields("fax")
    If IsNull(xRs.Fields("seguradora")) = False Then Flex.TextMatrix(X, 8) = PriMaiuscula(xRs.Fields("seguradora"))
    If IsNull(xRs.Fields("apolice")) = False Then Flex.TextMatrix(X, 9) = xRs.Fields("apolice")
    If IsNull(xRs.Fields("ie")) = False Then Flex.TextMatrix(X, 10) = xRs.Fields("ie")
    If IsNull(xRs.Fields("fantasia")) = False Then Flex.TextMatrix(X, 11) = PriMaiuscula(xRs.Fields("fantasia"))
    If IsNull(xRs.Fields("apelido")) = False Then Flex.TextMatrix(X, 12) = PriMaiuscula(xRs.Fields("apelido"))
    X = X + 1
    xRs.MoveNext
    Loop
    
End Sub
