VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadProdINT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Interno de Produtos"
   ClientHeight    =   5655
   ClientLeft      =   1530
   ClientTop       =   1470
   ClientWidth     =   9855
   ControlBox      =   0   'False
   Icon            =   "frmCadProdINT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9855
   Begin VB.Frame Frame2 
      Caption         =   "Produtos Cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Left            =   120
      TabIndex        =   9
      Top             =   1500
      Width           =   4755
      Begin MSFlexGridLib.MSFlexGrid FlexPROD 
         Height          =   3675
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   6482
         _Version        =   393216
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descrições dos Códigos IATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   4980
      TabIndex        =   8
      Top             =   60
      Width           =   4755
      Begin MSFlexGridLib.MSFlexGrid FlexIATA 
         Height          =   5115
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   9022
         _Version        =   393216
         BackColor       =   16777152
         ForeColor       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.CommandButton cmd_Alterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1320
      TabIndex        =   3
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Gravar 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   2520
      TabIndex        =   4
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Canc 
      Caption         =   "Canc/Sair"
      Height          =   435
      Left            =   3720
      TabIndex        =   6
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Nova 
      Caption         =   "Nova"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
   Begin VB.Frame fra_Dados 
      Caption         =   "Descrição do Produto"
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
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4755
      Begin VB.TextBox TxtCod 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   840
         MaxLength       =   25
         TabIndex        =   2
         Top             =   300
         Width           =   3795
      End
   End
   Begin VB.Label lbl_status 
      Caption         =   "INICIO"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmCadProdINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xCodigo As String
Public xDescricao As String
Private Sub cmd_Alterar_Click()
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False
FlexPROD.Enabled = False
fra_Dados.Enabled = True
cmd_Gravar.Enabled = True
lbl_status.Caption = "ALTERACAO"

xCodigo = TxtCod.Text
xDescricao = txtDescricao.Text

TxtCod.BackColor = &HC0FFFF
TxtCod.SetFocus
txtDescricao.BackColor = &HC0FFFF
End Sub

Private Sub cmd_Canc_Click()

    If lbl_status.Caption = "INCLUSAO" Or lbl_status.Caption = "ALTERACAO" Then
    cmd_nova.Enabled = True
    FlexPROD.Enabled = True
    fra_Dados.Enabled = False
    cmd_Gravar.Enabled = False
    
    TxtCod.BackColor = &H80000014
    txtDescricao.BackColor = &H80000014
    
        
    If de_informa.rsSel_CadProdINT.State = 1 Then de_informa.rsSel_CadProdINT.Close
    de_informa.Sel_CadProdINT "%"
    FlexPROD.Clear
    FlexPROD.Rows = de_informa.rsSel_CadProdINT.RecordCount + 1
    FlexPROD.Cols = 2
    FlexPROD.FixedRows = 1
    FlexPROD.FixedCols = 0
    FlexPROD.TextMatrix(0, 0) = "Código"
    FlexPROD.TextMatrix(0, 1) = "Descrição"
    FlexPROD.ColWidth(0) = 420
    FlexPROD.ColWidth(1) = 3000
    
    Y = 0
    
        Do Until de_informa.rsSel_CadProdINT.EOF
        Y = Y + 1
        FlexPROD.TextMatrix(Y, 0) = String(3 - Len(Trim(Str(Val(de_informa.rsSel_CadProdINT.Fields("codigo"))))), "0") & Trim(Str(Val(de_informa.rsSel_CadProdINT.Fields("codigo"))))
        FlexPROD.TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadProdINT.Fields("descricao"))
        de_informa.rsSel_CadProdINT.MoveNext
        Loop
    

   
    
    TxtCod.Text = ""
    txtDescricao.Text = ""

    lbl_status.Caption = "INICIO"
    
    Else
    Unload Me
    End If
End Sub



Private Sub cmd_Gravar_Click()
Dim xACHOU As Boolean

    If Len(Trim(txtDescricao.Text)) = 0 Then
    MsgBox "Não foi informada corretamente a Descricao", vbCritical, ""
    Exit Sub
    ElseIf Len(Trim(TxtCod.Text)) = 0 Then
    MsgBox "Não foi informado corretamente o Código", vbCritical, ""
    Exit Sub
    End If
    
    xACHOU = False
        
        For Y = 1 To FlexPROD.Rows - 1
            If UCase(FlexPROD.TextMatrix(Y, 0)) = UCase(Trim(TxtCod.Text)) Then
            xACHOU = True
            Y = FlexPROD.Rows
            End If
        Next
    
    If xACHOU = False Then
    MsgBox "O código IATA que você está informando, não consta no rol de códigos cadastrados. Reveja os dados informados ou cadastre este novo código IATA.", vbCritical, ""
    Exit Sub
    End If

    If lbl_status.Caption = "INCLUSAO" Then
        For Y = 1 To FlexPROD.Rows - 1
            If UCase(FlexPROD.TextMatrix(Y, 1)) = UCase(Trim(txtDescricao.Text)) Then
            MsgBox "O Produto " & UCase(txtDescricao.Text) & " já está cadastrado.", vbCritical, ""
            Exit Sub
            End If
        Next
    
        For Y = 1 To FlexPROD.Rows - 1
            If InStr(1, FlexPROD.TextMatrix(Y, 1), Trim(txtDescricao.Text), vbTextCompare) > 0 Then
                If MsgBox("O Produto " & UCase(FlexPROD.TextMatrix(Y, 1)) & " já está cadastrado. Deseja cadastrar " & txtDescricao.Text & " também?", vbYesNo + vbExclamation, "") = vbNo Then
                Exit Sub
                End If
            End If
        Next
        
        de_informa.Ins_CadProdINT TxtCod.Text, UCase(Trim(txtDescricao.Text))
    ElseIf lbl_status.Caption = "ALTERACAO" Then
        de_informa.Update_CadProdINT TxtCod.Text, UCase(Trim(txtDescricao.Text)), xCodigo, xDescricao
    End If
    


cmd_nova.Enabled = True
FlexPROD.Enabled = True
fra_Dados.Enabled = False
cmd_Gravar.Enabled = False

TxtCod.BackColor = &H80000014
TxtCod.Text = ""

txtDescricao.BackColor = &H80000014
txtDescricao.Text = ""

Call Form_Load

lbl_status.Caption = "INICIO"
End Sub

Private Sub cmd_Nova_Click()
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False
FlexPROD.Enabled = False
fra_Dados.Enabled = True
cmd_Gravar.Enabled = True
lbl_status.Caption = "INCLUSAO"

TxtCod.BackColor = &HC0FFFF
TxtCod.Text = ""
TxtCod.SetFocus

txtDescricao.BackColor = &HC0FFFF
txtDescricao.Text = ""
End Sub


Private Sub FlexPROD_Click()
TxtCod.Text = FlexPROD.TextMatrix(FlexPROD.Row, 0)
txtDescricao.Text = FlexPROD.TextMatrix(FlexPROD.Row, 1)
cmd_Alterar.Enabled = True
End Sub

Private Sub Form_Load()
    
    If de_informa.rsSel_CadProdINT.State = 1 Then de_informa.rsSel_CadProdINT.Close
    de_informa.Sel_CadProdINT "%"
    FlexPROD.Clear
    FlexPROD.Rows = de_informa.rsSel_CadProdINT.RecordCount + 1
    FlexPROD.Cols = 2
    FlexPROD.FixedRows = 1
    FlexPROD.FixedCols = 0
    FlexPROD.TextMatrix(0, 0) = "Código"
    FlexPROD.TextMatrix(0, 1) = "Descrição"
    FlexPROD.ColWidth(0) = 420
    FlexPROD.ColWidth(1) = 3000
    
    Y = 0
    
        Do Until de_informa.rsSel_CadProdINT.EOF
        Y = Y + 1
        FlexPROD.TextMatrix(Y, 0) = String(3 - Len(Trim(Str(Val(de_informa.rsSel_CadProdINT.Fields("codigo"))))), "0") & Trim(Str(Val(de_informa.rsSel_CadProdINT.Fields("codigo"))))
        FlexPROD.TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadProdINT.Fields("descricao"))
        de_informa.rsSel_CadProdINT.MoveNext
        Loop
    
    
    If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
    de_informa.Sel_Cadiata "%"
    FlexIATA.Clear
    FlexIATA.Rows = de_informa.rsSel_CadIATA.RecordCount + 1
    FlexIATA.Cols = 2
    FlexIATA.FixedRows = 1
    FlexIATA.FixedCols = 0
    FlexIATA.TextMatrix(0, 0) = "Código"
    FlexIATA.TextMatrix(0, 1) = "Descrição"
    FlexIATA.ColWidth(0) = 420
    FlexIATA.ColWidth(1) = 4000
    
    Y = 0
    
        Do Until de_informa.rsSel_CadIATA.EOF
        Y = Y + 1
        FlexIATA.TextMatrix(Y, 0) = String(3 - Len(Trim(Str(Val(de_informa.rsSel_CadIATA.Fields("codigo"))))), "0") & Trim(Str(Val(de_informa.rsSel_CadIATA.Fields("codigo"))))
        FlexIATA.TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadIATA.Fields("descricao"))
        de_informa.rsSel_CadIATA.MoveNext
        Loop
    DoEvents
    
End Sub

Private Sub Teste_Change()

Dim Abc As String
Dim Cont As Integer
Dim Flag As Boolean
Dim TextoAux As String

Abc = "ABCEDFGHIJKLMNOPQRSTUVYXWZ"
TextoAux = ""

    If Len(TxtSigla.Text) > 0 Then
        
        Flag = False
        
        For Cont = 1 To Len(Abc)
            If UCase(Mid(TxtSigla.Text, TxtSigla.SelStart, 1)) = Mid(Abc, Cont, 1) Then
            Flag = True
            End If
        Next
    
        If Flag = False Then
            For Cont = 1 To Len(TxtSigla.Text)
                If Cont <> TxtSigla.SelStart Then
                TextoAux = Mid(TxtSigla, Cont, 1)
                End If
            Next
        Else
        TextoAux = TxtSigla.Text
        End If
        
        TxtSigla.Text = TextoAux
    End If
        
End Sub


Private Sub TxtCod_GotFocus()
TxtCod.SelStart = 0
TxtCod.SelLength = 100
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
    If KeyAscii < 47 Or KeyAscii > 58 Then
        If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtCod_LostFocus()
TxtCod.Text = String(3 - Len(Trim(TxtCod.Text)), "0") & Trim(TxtCod.Text)
End Sub

Private Sub txtDescricao_GotFocus()
txtDescricao.SelStart = 0
txtDescricao.SelLength = 100
End Sub


Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub


Private Sub txtDescricao_LostFocus()
txtDescricao.Text = UCase(txtDescricao.Text)
End Sub

Private Sub TxtSigla_GotFocus()
TxtSigla.SelStart = 0
TxtSigla.SelLength = 10
End Sub


Private Sub TxtSigla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub


Private Sub TxtSigla_LostFocus()
TxtSigla.Text = UCase(TxtSigla.Text)
End Sub

Private Sub txtUF_GotFocus()
TxtUF.SelStart = 0
TxtUF.SelLength = 10
End Sub


Private Sub txtUF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
End If
End Sub


Private Sub txtUF_LostFocus()
TxtUF.Text = UCase(TxtUF.Text)
End Sub


