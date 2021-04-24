VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoFiltraClienteExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Especifique seu Cliente"
   ClientHeight    =   2670
   ClientLeft      =   210
   ClientTop       =   2295
   ClientWidth     =   11595
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmEmissaoFiltraCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraFiliais 
      Caption         =   "Escolha seu Tipo de Busca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   2895
      Begin VB.TextBox TxtBusca 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   2655
      End
      Begin VB.CommandButton CmdBusca 
         Caption         =   "Buscar!"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   2040
         Width           =   2655
      End
      Begin VB.OptionButton OptFantasia 
         Caption         =   "Busca por Nome Fantasia"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1000
         Width           =   2655
      End
      Begin VB.OptionButton OptNome 
         Caption         =   "Busca por Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1340
         Width           =   2655
      End
      Begin VB.OptionButton OptCGC 
         Caption         =   "Busca por CNPJ"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2655
      End
      Begin VB.OptionButton OptApelido 
         Caption         =   "Busca por Apelido"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2415
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmEmissaoFiltraClienteExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBusca_Click()
Dim X As Integer

If Len(Trim(TxtBusca.Text)) > 0 Then
    If OptApelido.Value = True Then
        If de_informa.rsSelClienteAPELIDO.State = 1 Then de_informa.rsSelClienteAPELIDO.Close
        de_informa.SelClienteAPELIDO "%" & TxtBusca.Text & "%"
            If de_informa.rsSelClienteAPELIDO.RecordCount > 0 Then
            Flex.Clear
            Flex.Cols = 7
            Flex.Rows = de_informa.rsSelClienteAPELIDO.RecordCount + 1
            Flex.FixedRows = 1
            Flex.FixedCols = 0
            Flex.TextMatrix(0, 0) = "CNPJ"
            Flex.TextMatrix(0, 1) = "Apelido"
            Flex.TextMatrix(0, 2) = "Nome Fantasia"
            Flex.TextMatrix(0, 3) = "Nome"
            Flex.TextMatrix(0, 4) = "Cidade"
            Flex.TextMatrix(0, 5) = "UF"
            Flex.TextMatrix(0, 6) = "Endereço"
            Flex.ColWidth(0) = 1500
            Flex.ColWidth(1) = 1500
            Flex.ColWidth(2) = 4000
            Flex.ColWidth(3) = 4000
            Flex.ColWidth(4) = 3000
            Flex.ColWidth(5) = 500
            Flex.ColWidth(6) = 6000
            X = 1
            Do Until de_informa.rsSelClienteAPELIDO.EOF
            Flex.TextMatrix(X, 0) = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("cgc"))
            Flex.TextMatrix(X, 1) = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("apelido"))
            Flex.TextMatrix(X, 2) = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("fantasia"))
            Flex.TextMatrix(X, 3) = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("nome"))
            Flex.TextMatrix(X, 4) = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("cidade"))
            Flex.TextMatrix(X, 5) = de_informa.rsSelClienteAPELIDO.Fields("uf")
            Flex.TextMatrix(X, 6) = PriMaiuscula(de_informa.rsSelClienteAPELIDO.Fields("endereco"))
            X = X + 1
            de_informa.rsSelClienteAPELIDO.MoveNext
            Loop
            Flex.SetFocus
            Else
            Flex.Clear
            MsgBox "Sua Pesquisa não retornou registro algum!", vbCritical, ""
            End If
    ElseIf OptCGC.Value = True Then
        If de_informa.rsSelClienteCNPJ.State = 1 Then de_informa.rsSelClienteCNPJ.Close
        de_informa.SelClienteCNPJ TxtBusca.Text & "%"
        
            If de_informa.rsSelClienteCNPJ.RecordCount > 0 Then
            Flex.Clear
            Flex.Cols = 7
            Flex.Rows = de_informa.rsSelClienteCNPJ.RecordCount + 1
            Flex.FixedRows = 1
            Flex.FixedCols = 0
            Flex.TextMatrix(0, 0) = "CNPJ"
            Flex.TextMatrix(0, 1) = "Apelido"
            Flex.TextMatrix(0, 2) = "Nome Fantasia"
            Flex.TextMatrix(0, 3) = "Nome"
            Flex.TextMatrix(0, 4) = "Cidade"
            Flex.TextMatrix(0, 5) = "UF"
            Flex.TextMatrix(0, 6) = "Endereço"
            Flex.ColWidth(0) = 1500
            Flex.ColWidth(1) = 1500
            Flex.ColWidth(2) = 4000
            Flex.ColWidth(3) = 4000
            Flex.ColWidth(4) = 3000
            Flex.ColWidth(5) = 500
            Flex.ColWidth(6) = 6000
            X = 1
            Do Until de_informa.rsSelClienteCNPJ.EOF
            Flex.TextMatrix(X, 0) = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("cgc"))
            Flex.TextMatrix(X, 1) = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("apelido"))
            Flex.TextMatrix(X, 2) = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("fantasia"))
            Flex.TextMatrix(X, 3) = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("nome"))
            Flex.TextMatrix(X, 4) = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("cidade"))
            Flex.TextMatrix(X, 5) = de_informa.rsSelClienteCNPJ.Fields("uf")
            Flex.TextMatrix(X, 6) = PriMaiuscula(de_informa.rsSelClienteCNPJ.Fields("endereco"))
            X = X + 1
            de_informa.rsSelClienteCNPJ.MoveNext
            Loop
            Flex.SetFocus
            Else
            Flex.Clear
            MsgBox "Sua Pesquisa não retornou registro algum!", vbCritical, ""
            End If
    ElseIf OptFantasia.Value = True Then
        If de_informa.rsSelClienteFANTASIA.State = 1 Then de_informa.rsSelClienteFANTASIA.Close
        de_informa.SelClientefantasia "%" & TxtBusca.Text & "%"
        
            If de_informa.rsSelClienteFANTASIA.RecordCount > 0 Then
            Flex.Clear
            Flex.Cols = 7
            Flex.Rows = de_informa.rsSelClienteFANTASIA.RecordCount + 1
            Flex.FixedRows = 1
            Flex.FixedCols = 0
            Flex.TextMatrix(0, 0) = "CNPJ"
            Flex.TextMatrix(0, 1) = "Apelido"
            Flex.TextMatrix(0, 2) = "Nome Fantasia"
            Flex.TextMatrix(0, 3) = "Nome"
            Flex.TextMatrix(0, 4) = "Cidade"
            Flex.TextMatrix(0, 5) = "UF"
            Flex.TextMatrix(0, 6) = "Endereço"
            Flex.ColWidth(0) = 1500
            Flex.ColWidth(1) = 1500
            Flex.ColWidth(2) = 4000
            Flex.ColWidth(3) = 4000
            Flex.ColWidth(4) = 3000
            Flex.ColWidth(5) = 500
            Flex.ColWidth(6) = 6000
            X = 1
            Do Until de_informa.rsSelClienteFANTASIA.EOF
            Flex.TextMatrix(X, 0) = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("cgc"))
            Flex.TextMatrix(X, 1) = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("apelido"))
            Flex.TextMatrix(X, 2) = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("fantasia"))
            Flex.TextMatrix(X, 3) = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("nome"))
            Flex.TextMatrix(X, 4) = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("cidade"))
            Flex.TextMatrix(X, 5) = de_informa.rsSelClienteFANTASIA.Fields("uf")
            Flex.TextMatrix(X, 6) = PriMaiuscula(de_informa.rsSelClienteFANTASIA.Fields("endereco"))
            X = X + 1
            de_informa.rsSelClienteFANTASIA.MoveNext
            Loop
            Flex.SetFocus
            Else
            Flex.Clear
            MsgBox "Sua Pesquisa não retornou registro algum!", vbCritical, ""
            End If
    ElseIf OptNome.Value = True Then
        If de_informa.rsSelClienteNOME.State = 1 Then de_informa.rsSelClienteNOME.Close
        de_informa.SelClientenome "%" & TxtBusca.Text & "%"
        
            If de_informa.rsSelClienteNOME.RecordCount > 0 Then
            Flex.Clear
            Flex.Cols = 7
            Flex.Rows = de_informa.rsSelClienteNOME.RecordCount + 1
            Flex.FixedRows = 1
            Flex.FixedCols = 0
            Flex.TextMatrix(0, 0) = "CNPJ"
            Flex.TextMatrix(0, 1) = "Apelido"
            Flex.TextMatrix(0, 2) = "Nome Fantasia"
            Flex.TextMatrix(0, 3) = "Nome"
            Flex.TextMatrix(0, 4) = "Cidade"
            Flex.TextMatrix(0, 5) = "UF"
            Flex.TextMatrix(0, 6) = "Endereço"
            Flex.ColWidth(0) = 1500
            Flex.ColWidth(1) = 1500
            Flex.ColWidth(2) = 4000
            Flex.ColWidth(3) = 4000
            Flex.ColWidth(4) = 3000
            Flex.ColWidth(5) = 500
            Flex.ColWidth(6) = 6000
            X = 1
            Do Until de_informa.rsSelClienteNOME.EOF
            Flex.TextMatrix(X, 0) = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("cgc"))
            Flex.TextMatrix(X, 1) = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("apelido"))
            Flex.TextMatrix(X, 2) = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("fantasia"))
            Flex.TextMatrix(X, 3) = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("nome"))
            Flex.TextMatrix(X, 4) = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("cidade"))
            Flex.TextMatrix(X, 5) = de_informa.rsSelClienteNOME.Fields("uf")
            Flex.TextMatrix(X, 6) = PriMaiuscula(de_informa.rsSelClienteNOME.Fields("endereco"))
            X = X + 1
            de_informa.rsSelClienteNOME.MoveNext
            Loop
            Flex.SetFocus
            Else
            Flex.Clear
            MsgBox "Sua Pesquisa não retornou registro algum!", vbCritical, ""
            End If
    End If
End If
End Sub

Private Sub CmdBusca_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)
Dim X As Integer


X = Flex.Row

    If KeyAscii = 27 Then
    Unload Me
    ElseIf KeyAscii = 13 Then
        If X >= 1 Then
        frmEmissao.TxtCGCExpedidor.Text = Flex.TextMatrix(X, 0)
        frmEmissao.TxtNomeExpedidor.Text = Flex.TextMatrix(X, 3)
        frmEmissao.TxtCidadeExpedidor.Text = Flex.TextMatrix(X, 4)
        frmEmissao.TxtUFExpedidor.Text = Flex.TextMatrix(X, 5)
        Unload Me
        End If
    End If
        
End Sub

Private Sub OptApelido_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Private Sub OptCGC_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Private Sub OptFantasia_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Private Sub OptNome_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub

Private Sub TxtBusca_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub
