VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoFiltraRESPRES 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Especifique seu Representante"
   ClientHeight    =   2115
   ClientLeft      =   1770
   ClientTop       =   3120
   ClientWidth     =   8355
   Icon            =   "frmEmissaoFiltraRESPRES.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
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
Attribute VB_Name = "frmEmissaoFiltraRESPRES"
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
    
    frmEmissao.TxtNomeDestinatario.Text = PriMaiuscula(Flex.TextMatrix(X, 0))
    frmEmissao.TxtCidadeDestinatario.Text = PriMaiuscula(Flex.TextMatrix(X, 1))
    frmEmissao.TxtUFDestinatario.Text = Flex.TextMatrix(X, 2)
    frmEmissao.TxtCGCDestinatario.Text = PriMaiuscula(Flex.TextMatrix(X, 3))
    frmEmissao.TxtInscrEstDestinatario.Text = Flex.TextMatrix(X, 4)
    frmEmissao.TxtEndDestinatario.Text = PriMaiuscula(Flex.TextMatrix(X, 5))
    frmEmissao.TxtBairroDEST.Text = PriMaiuscula(Flex.TextMatrix(X, 6))
    frmEmissao.TxtCEPDestinatario.Text = Flex.TextMatrix(X, 7)
    frmEmissao.TxtTelDestinatario.Text = Flex.TextMatrix(X, 8)
    frmEmissao.TxtFAXDestinatario.Text = Flex.TextMatrix(X, 9)
    
    If de_informa.rsSelAeroportoCidade.State = 1 Then de_informa.rsSelAeroportoCidade.Close
    de_informa.SelAeroportoCidade UCase(Flex.TextMatrix(X, 10))
    frmEmissao.TxtSiglaVIA.Text = de_informa.rsSelAeroportoCidade.Fields("sigla")
    frmEmissao.TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("localidade")) & " - " & de_informa.rsSelAeroportoCidade.Fields("uf") & " (" & PriMaiuscula(de_informa.rsSelAeroportoCidade.Fields("aeroporto")) & ")"
    
    Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim X As Integer

Set xRs = de_informa.rsSelRepres

Flex.Clear
Flex.Cols = 12
Flex.Rows = xRs.RecordCount + 1
Flex.FixedRows = 1
Flex.FixedCols = 0
Flex.TextMatrix(0, 0) = "Nome"
Flex.TextMatrix(0, 1) = "Cidade"
Flex.TextMatrix(0, 2) = "UF"
Flex.TextMatrix(0, 3) = "CNPJ"
Flex.TextMatrix(0, 4) = "Inscr. Est."
Flex.TextMatrix(0, 5) = "Endereço"
Flex.TextMatrix(0, 6) = "Bairro"
Flex.TextMatrix(0, 7) = "CEP"
Flex.TextMatrix(0, 8) = "Tel. Com."
Flex.TextMatrix(0, 9) = "FAX"
Flex.TextMatrix(0, 10) = "Cidade Retira"
Flex.TextMatrix(0, 11) = "UF Retira"
Flex.ColWidth(0) = 3500
Flex.ColWidth(1) = 2500
Flex.ColWidth(2) = 500
Flex.ColWidth(3) = 1500
Flex.ColWidth(4) = 1500
Flex.ColWidth(5) = 3500
Flex.ColWidth(6) = 2500
Flex.ColWidth(7) = 1500
Flex.ColWidth(8) = 1500
Flex.ColWidth(9) = 1500
Flex.ColWidth(10) = 2500
Flex.ColWidth(11) = 500



X = 1
    Do Until xRs.EOF
    Flex.TextMatrix(X, 0) = PriMaiuscula(xRs.Fields("nome"))
    Flex.TextMatrix(X, 1) = PriMaiuscula(xRs.Fields("localidade"))
    Flex.TextMatrix(X, 2) = xRs.Fields("uf")
    Flex.TextMatrix(X, 3) = xRs.Fields("cgc")
    Flex.TextMatrix(X, 4) = xRs.Fields("inscr_est")
    Flex.TextMatrix(X, 5) = xRs.Fields("endereco")
    Flex.TextMatrix(X, 6) = xRs.Fields("bairro")
    Flex.TextMatrix(X, 7) = xRs.Fields("cep")
    Flex.TextMatrix(X, 8) = xRs.Fields("telcom")
    Flex.TextMatrix(X, 9) = xRs.Fields("fax")
    Flex.TextMatrix(X, 10) = PriMaiuscula(xRs.Fields("cidaderetira"))
    Flex.TextMatrix(X, 11) = xRs.Fields("ufretira")
    X = X + 1
    xRs.MoveNext
    Loop
End Sub
