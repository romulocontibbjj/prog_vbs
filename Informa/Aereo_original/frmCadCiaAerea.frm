VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadCiaAerea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cia. Aérea"
   ClientHeight    =   5475
   ClientLeft      =   2010
   ClientTop       =   1185
   ClientWidth     =   7275
   ControlBox      =   0   'False
   Icon            =   "frmCadCiaAerea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7275
   Begin MSFlexGridLib.MSFlexGrid FlexCia 
      Height          =   2715
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   4789
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   880
      Width           =   1275
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   1520
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancSair 
      Caption         =   "Canc/Sair"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   240
      Width           =   1275
   End
   Begin VB.Frame fraDados 
      Caption         =   "Dados da Cia. Aérea"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5655
      Begin VB.Frame Frame2 
         Caption         =   "Formulário de AWB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   5415
         Begin VB.TextBox txtProximoNum 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtEstoqueAtual 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   18
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtEstoqueMinimo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkAvisoMinimo 
            Caption         =   "Aviso de Estoque Mínimo Atingido"
            Height          =   195
            Left            =   2520
            TabIndex        =   4
            Top             =   765
            Width           =   2775
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Próximo Número:"
            Height          =   195
            Left            =   2580
            TabIndex        =   17
            Top             =   405
            Width           =   1200
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Estoque Atual:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   765
            Width           =   1200
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Estoque Mínimo:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   405
            Width           =   1200
         End
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   4275
      End
      Begin VB.TextBox txtFantasia 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtCodCia 
         Height          =   285
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   13
         Top             =   180
         Width           =   75
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia:"
         Height          =   195
         Left            =   2880
         TabIndex        =   12
         Top             =   405
         Width           =   1110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Descrição:"
         Height          =   195
         Left            =   420
         TabIndex        =   11
         Top             =   765
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código/Sigla:"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   420
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmCadCiaAerea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAvisoMinimo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmdAlterar_Click()
    fraDados.Enabled = True
    cmdNovo.Enabled = False
    cmdAlterar.Enabled = False
    cmdGravar.Enabled = True
    FlexCia.Enabled = False
    txtEstoqueAtual.Enabled = False
    txtProximoNum.Enabled = False
    txtCodCia.Enabled = False
    txtFantasia.BackColor = &HC0FFFF
    txtDescricao.BackColor = &HC0FFFF
    txtEstoqueMinimo.BackColor = &HC0FFFF
    lblStatus = "ALTERAÇÃO"
    txtFantasia.SetFocus
End Sub

Private Sub cmdCancSair_Click()
    If Len(lblStatus) > 1 Then
        fraDados.Enabled = False
        cmdNovo.Enabled = True
        txtCodCia.Enabled = True
        cmdGravar.Enabled = False
        FlexCia.Enabled = True
        txtCodCia.Text = ""
        txtFantasia.Text = ""
        txtDescricao.Text = ""
        txtEstoqueAtual.Text = ""
        txtEstoqueMinimo.Text = ""
        txtProximoNum.Text = ""
        chkAvisoMinimo.Value = 0
        lblStatus = ""
        txtCodCia.BackColor = &H80000014
        txtFantasia.BackColor = &H80000014
        txtDescricao.BackColor = &H80000014
        txtEstoqueMinimo.BackColor = &H80000014
        cmdCancSair.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub cmdGravar_Click()
    Dim xavisominimo As String
    If Len(Trim$(txtCodCia.Text)) < 2 Then
        MsgBox "Código de Cia. Aérea Inválido !"
        txtCodCia.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtFantasia.Text)) < 2 Then
        MsgBox "Nome Fantasia da Cia. Aérea Inválido !"
        txtFantasia.SetFocus
        Exit Sub
    End If
    If Len(Trim$(txtDescricao.Text)) < 2 Then
        MsgBox "Nome/Descrição da Cia. Aérea Inválido !"
        txtDescricao.SetFocus
        Exit Sub
    End If
    If txtEstoqueMinimo < 0 Then
        MsgBox "Valor de Estoque Mínimo Inválido !"
        txtEstoqueMinimo.SetFocus
        Exit Sub
    End If
    If chkAvisoMinimo = 1 Then
        xavisominimo = "S"
    Else
        xavisominomo = "N"
    End If
    
    'verifica se código e / ou fantasia já estão cadastrados
    
    If lblStatus = "INCLUSÃO" Then
        If de_informa.rsSel_CiaAereaPorCodigo.State = 1 Then de_informa.rsSel_CiaAereaPorCodigo.Close
        de_informa.Sel_CiaAereaPorCodigo Trim$(txtCodCia.Text)
        If de_informa.rsSel_CiaAereaPorCodigo.RecordCount > 0 Then
            MsgBox "Cia Aérea com este Código já Cadastrada !"
            txtCodCia.SetFocus
            Exit Sub
        End If
    
        If de_informa.rsSel_CiaAereaPorFantasia.State = 1 Then de_informa.rsSel_CiaAereaPorFantasia.Close
        de_informa.Sel_CiaAereaPorFantasia Trim$(txtFantasia.Text)
        If de_informa.rsSel_CiaAereaPorFantasia.RecordCount > 0 Then
            MsgBox "Cia. Aérea com este Nome Fantasia já Cadastrada !"
            txtFantasia.SetFocus
            Exit Sub
        End If
        
    
        'cadastra a Cia Aérea
    
        de_informa.Ins_CiaAerea Trim$(txtCodCia.Text), Trim$(txtFantasia.Text), Trim$(txtDescricao.Text), Val(txtEstoqueMinimo.Text), xavisominimo
    
        MsgBox "OK ! Cia Aérea Cadastrada"
    ElseIf lblStatus = "ALTERAÇÃO" Then
        
        'grava as alterações da Cia Aérea
    
        de_informa.Alt_CiaAerea Trim$(txtFantasia.Text), Trim$(txtDescricao.Text), Val(txtEstoqueMinimo.Text), xavisominimo, Trim$(txtCodCia.Text)
    
        MsgBox "OK ! Dados Alterados"
    
    End If
    
    Dim X, Y As Integer

If de_informa.rsSel_CiaAerea.State = 1 Then de_informa.rsSel_CiaAerea.Close
de_informa.Sel_CiaAerea

FlexCia.Clear
FlexCia.Rows = de_informa.rsSel_CiaAerea.RecordCount + 1
FlexCia.Cols = 7
FlexCia.FixedRows = 1
FlexCia.FixedCols = 0

FlexCia.TextMatrix(0, 0) = "Sigla"
FlexCia.TextMatrix(0, 1) = "Fantasia"
FlexCia.TextMatrix(0, 2) = "Descrição"
FlexCia.TextMatrix(0, 3) = "Est. Min."
FlexCia.TextMatrix(0, 4) = "Est. Atual"
FlexCia.TextMatrix(0, 5) = "Prox. Num."
FlexCia.TextMatrix(0, 6) = "Aviso"

FlexCia.ColWidth(0) = 500
FlexCia.ColWidth(1) = 2000
FlexCia.ColWidth(2) = 3500
FlexCia.ColWidth(3) = 820
FlexCia.ColWidth(4) = 820
FlexCia.ColWidth(5) = 820
FlexCia.ColWidth(6) = 820



Y = 1

    Do Until de_informa.rsSel_CiaAerea.EOF
    If IsNull(de_informa.rsSel_CiaAerea.Fields("codcia")) = False Then FlexCia.TextMatrix(Y, 0) = de_informa.rsSel_CiaAerea.Fields("codcia")
    If IsNull(de_informa.rsSel_CiaAerea.Fields("fantasia")) = False Then FlexCia.TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CiaAerea.Fields("fantasia"))
    If IsNull(de_informa.rsSel_CiaAerea.Fields("descricao")) = False Then FlexCia.TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CiaAerea.Fields("descricao"))
    If IsNull(de_informa.rsSel_CiaAerea.Fields("estoqueminimo")) = False Then FlexCia.TextMatrix(Y, 3) = de_informa.rsSel_CiaAerea.Fields("estoqueminimo")
    If IsNull(de_informa.rsSel_CiaAerea.Fields("estoqueatual")) = False Then FlexCia.TextMatrix(Y, 4) = de_informa.rsSel_CiaAerea.Fields("estoqueatual")
    If IsNull(de_informa.rsSel_CiaAerea.Fields("proximonum")) = False Then FlexCia.TextMatrix(Y, 5) = de_informa.rsSel_CiaAerea.Fields("proximonum")
    If IsNull(de_informa.rsSel_CiaAerea.Fields("avisominimo")) = False Then FlexCia.TextMatrix(Y, 6) = de_informa.rsSel_CiaAerea.Fields("avisominimo")
    Y = Y + 1
    de_informa.rsSel_CiaAerea.MoveNext
    Loop

    
    fraDados.Enabled = False
    cmdNovo.Enabled = True
    cmdGravar.Enabled = False
    FlexCia.Enabled = True
    txtCodCia.Enabled = True
    txtCodCia.Text = ""
    txtFantasia.Text = ""
    txtDescricao.Text = ""
    txtEstoqueAtual.Text = ""
    txtEstoqueMinimo.Text = ""
    txtProximoNum.Text = ""
    chkAvisoMinimo.Value = 0
    lblStatus = ""
    txtCodCia.BackColor = &H80000014
    txtFantasia.BackColor = &H80000014
    txtDescricao.BackColor = &H80000014
    txtEstoqueMinimo.BackColor = &H80000014
    cmdCancSair.SetFocus
    
End Sub

Private Sub cmdNovo_Click()
    fraDados.Enabled = True
    cmdNovo.Enabled = False
    cmdAlterar.Enabled = False
    cmdGravar.Enabled = True
    FlexCia.Enabled = False
    txtCodCia.Text = ""
    txtFantasia.Text = ""
    txtDescricao.Text = ""
    txtEstoqueMinimo = 0
    txtEstoqueAtual = ""
    txtProximoNum = ""
    chkAvisoMinimo.Value = 0
    txtCodCia.BackColor = &HC0FFFF
    txtFantasia.BackColor = &HC0FFFF
    txtDescricao.BackColor = &HC0FFFF
    txtEstoqueMinimo.BackColor = &HC0FFFF
    txtEstoqueAtual.Enabled = False
    txtProximoNum.Enabled = False
    lblStatus = "INCLUSÃO"
    txtCodCia.SetFocus
End Sub

Private Sub FlexCia_Click()
    txtCodCia.Text = FlexCia.TextMatrix(FlexCia.Row, 0)
    txtFantasia.Text = FlexCia.TextMatrix(FlexCia.Row, 1)
    txtDescricao.Text = FlexCia.TextMatrix(FlexCia.Row, 2)
    txtEstoqueMinimo = FlexCia.TextMatrix(FlexCia.Row, 3)
    txtEstoqueAtual = FlexCia.TextMatrix(FlexCia.Row, 4)
    txtProximoNum = FlexCia.TextMatrix(FlexCia.Row, 5)
    If FlexCia.TextMatrix(FlexCia.Row, 6) = "S" Then
        chkAvisoMinimo.Value = 1
    Else
        chkAvisoMinimo.Value = 0
    End If
    'lblData = GridCiaAerea.Columns(7)
    cmdAlterar.Enabled = True
End Sub

Private Sub Form_Activate()
    cmdCancSair.SetFocus
End Sub




Private Sub Form_Load()

Dim X, Y As Integer

If de_informa.rsSel_CiaAerea.State = 1 Then de_informa.rsSel_CiaAerea.Close
de_informa.Sel_CiaAerea

FlexCia.Clear
FlexCia.Rows = de_informa.rsSel_CiaAerea.RecordCount + 1
FlexCia.Cols = 7
FlexCia.FixedRows = 1
FlexCia.FixedCols = 0

FlexCia.TextMatrix(0, 0) = "Sigla"
FlexCia.TextMatrix(0, 1) = "Fantasia"
FlexCia.TextMatrix(0, 2) = "Descrição"
FlexCia.TextMatrix(0, 3) = "Est. Min."
FlexCia.TextMatrix(0, 4) = "Est. Atual"
FlexCia.TextMatrix(0, 5) = "Prox. Num."
FlexCia.TextMatrix(0, 6) = "Aviso"

FlexCia.ColWidth(0) = 500
FlexCia.ColWidth(1) = 2000
FlexCia.ColWidth(2) = 3500
FlexCia.ColWidth(3) = 820
FlexCia.ColWidth(4) = 820
FlexCia.ColWidth(5) = 820
FlexCia.ColWidth(6) = 820



Y = 1

    Do Until de_informa.rsSel_CiaAerea.EOF
    If IsNull(de_informa.rsSel_CiaAerea.Fields("codcia")) = False Then FlexCia.TextMatrix(Y, 0) = de_informa.rsSel_CiaAerea.Fields("codcia")
    If IsNull(de_informa.rsSel_CiaAerea.Fields("fantasia")) = False Then FlexCia.TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CiaAerea.Fields("fantasia"))
    If IsNull(de_informa.rsSel_CiaAerea.Fields("descricao")) = False Then FlexCia.TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CiaAerea.Fields("descricao"))
    If IsNull(de_informa.rsSel_CiaAerea.Fields("estoqueminimo")) = False Then FlexCia.TextMatrix(Y, 3) = de_informa.rsSel_CiaAerea.Fields("estoqueminimo")
    If IsNull(de_informa.rsSel_CiaAerea.Fields("estoqueatual")) = False Then FlexCia.TextMatrix(Y, 4) = de_informa.rsSel_CiaAerea.Fields("estoqueatual")
    If IsNull(de_informa.rsSel_CiaAerea.Fields("proximonum")) = False Then FlexCia.TextMatrix(Y, 5) = de_informa.rsSel_CiaAerea.Fields("proximonum")
    If IsNull(de_informa.rsSel_CiaAerea.Fields("avisominimo")) = False Then FlexCia.TextMatrix(Y, 6) = de_informa.rsSel_CiaAerea.Fields("avisominimo")
    Y = Y + 1
    de_informa.rsSel_CiaAerea.MoveNext
    Loop
    



End Sub

Private Sub txtCodCia_GotFocus()
    txtCodCia.SelStart = 0
    txtCodCia.SelLength = 3
End Sub

Private Sub txtCodCia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCodCia_LostFocus()
    If Len(Trim$(txtCodCia)) > 0 Then
        txtCodCia = UCase(Trim$(txtCodCia))
    End If
End Sub

Private Sub txtDescricao_GotFocus()
    txtDescricao.SelStart = 0
    txtDescricao.SelLength = 50
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtDescricao_LostFocus()
    If Len(Trim$(txtDescricao)) > 0 Then
        txtDescricao = UCase(Trim$(txtDescricao))
    End If
End Sub

Private Sub txtEstoqueMinimo_Change()
    sonumero (txtEstoqueMinimo)
End Sub

Private Sub txtEstoqueMinimo_GotFocus()
    txtEstoqueMinimo.SelStart = 0
    txtEstoqueMinimo.SelLength = 6
End Sub

Private Sub txtEstoqueMinimo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFantasia_GotFocus()
    txtFantasia.SelStart = 0
    txtFantasia.SelLength = 20
End Sub

Private Sub txtFantasia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFantasia_LostFocus()
    If Len(Trim$(txtFantasia)) > 0 Then
        txtFantasia = UCase(Trim$(txtFantasia))
    End If
End Sub
