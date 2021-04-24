VERSION 5.00
Begin VB.Form FRM_FichadeClientes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Clientes"
   ClientHeight    =   5445
   ClientLeft      =   2145
   ClientTop       =   2265
   ClientWidth     =   9195
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9195
   Begin VB.Frame FrameDadosCadastrais 
      Caption         =   " Dados Cadastrais do Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.TextBox txt_obs 
         Height          =   855
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   3600
         Width           =   8895
      End
      Begin VB.TextBox txt_email 
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txt_celular 
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txt_fone 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txt_contato 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txt_cidade 
         Height          =   285
         Left            =   6000
         TabIndex        =   7
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txt_bairro 
         Height          =   285
         Left            =   4200
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txt_endereco 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txtIE 
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtcgc 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtFantasia 
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtRazao 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.Frame Frame_cadastroClientes 
         Height          =   735
         Left            =   6480
         TabIndex        =   15
         Top             =   4560
         Width           =   2535
         Begin VB.CommandButton cmd_altera 
            Caption         =   "&Alterar"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton CbtSair 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   1320
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmb_inclui 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label12 
         Caption         =   "Obs"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Celular"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Telefone"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Contato"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   6000
         TabIndex        =   22
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Endereço"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Inscrição Estadual"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "CGC"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nome Fantasia"
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Razão Social"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRM_FichadeClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CbtIncluir_Click()

Dim cn As New ADODB.Connection
'Dim rs As New ADODB.Recordset
Dim CnString As String
Dim SqlString As String
Dim SqlInser As String

    CnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\CotyBoys\CotyBoys.mdb;Mode=ReadWrite;Persist Security Info=False"
    'SqlString = " Select * from Clientes where RazaoSocial = '" & txtRazao.Text & "' "
    If txtRazao.Text = Empty Then
        MsgBox ("Por favor, informe a Razão Social do Cliente"), vbInformation
        txtRazao.SetFocus
        Exit Sub
    End If
    SqlString = "insert into Clientes(RazaoSocial) "
    SqlString = SqlString & "Values "
    SqlString = SqlString & "(" & "'" & txtRazao.Text & "'" & ")"
    cn.Open (CnString)
        cn.BeginTrans
        cn.Execute (SqlString)
        
        If Err.Number <> 0 Then
        
            cn.RollbackTrans
            MsgBox Err.Description
            cn.Close
        Else
            MsgBox ("Cliente Cadastrado com Sucesso"), vbInformation, "Inclusão"
            cn.CommitTrans
            cn.Close
        End If
    'rs.Open (SqlString), cn





'If rs.EOF Then

    'MsgBox ("NãO FOI ENCONTRADO NADA ")
    
'Else

    
'End If





End Sub

Private Sub CbtExcluir_Click()

End Sub

Private Sub CbtSair_Click()
FRM_CadastrodeClientes.Enabled = True
limpa_tela (1)
Unload Me
End Sub

Private Sub cmb_inclui_Click()

If Trim$(Len(txtRazao.Text)) = 0 Then
    MsgBox " Digite a Razão Social do Cliente.", vbInformation, "CLIENTE"
    txtRazao.SetFocus
    Exit Sub
    
ElseIf Trim$(Len(txtcgc.Text)) < 12 Then
    MsgBox "Corrija o CGC do Cliente.", vbInformation, "CLIENTE"
    txtcgc.SelStart = 0
    txtcgc.SelLength = Len(txtcgc.Text)
    Exit Sub

End If





deb_coty.In_Clientes txtRazao.Text, txtFantasia.Text, txtcgc.Text, txtIE.Text, _
                        txt_endereco.Text, txt_bairro.Text, txt_Cidade.Text, txt_contato.Text, _
                        txt_fone.Text, txt_celular.Text, txt_email.Text, txt_obs.Text
                        
If deb_coty.rsSel_clientes.State = 1 Then deb_coty.rsSel_clientes.Close
    deb_coty.Sel_clientes
    
    FRM_CadastrodeClientes.DTGridClientes.DataMember = "Sel_clientes"
    FRM_CadastrodeClientes.DTGridClientes.Refresh

MsgBox "Cliente: " & txtRazao & Chr$(13) & "Cadastrado com Sucesso", vbInformation, "CLIENTES"

limpa_tela (1)
FRM_CadastrodeClientes.Enabled = True
Unload Me
End Sub

Private Sub cmd_altera_Click()
txtcgc.Locked = True

deb_coty.Alt_Clientes txtRazao.Text, txtFantasia.Text, txtcgc.Text, txtIE.Text, _
                        txt_endereco.Text, txt_bairro.Text, txt_Cidade.Text, txt_contato.Text, _
                        txt_fone.Text, txt_celular.Text, txt_email.Text, txt_obs.Text, txtcgc.Text

MsgBox "Cliente: " & txtRazao.Text & Chr$(13) & "Alterado com Sucesso", vbInformation, "CLIENTE"

limpa_tela (1)

FRM_CadastrodeClientes.Enabled = True

If deb_coty.rsSel_clientes.State = 1 Then deb_coty.rsSel_clientes.Close
    deb_coty.Sel_clientes
    
    FRM_CadastrodeClientes.DTGridClientes.DataMember = "Sel_clientes"
    FRM_CadastrodeClientes.DTGridClientes.Refresh

Unload Me


End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))

End Sub

Private Sub Text6_Change()
Dim X As Integer
Dim Y As String

If Trim$(Len(Text6.Text)) <> 0 Then

    If IsNumeric(Text6.Text) = False Then
        
        MsgBox "Digite o Número do Celular"
        X = Len(Text6.Text)
        Y = Mid(Text6.Text, 1, X - 1)
        Text6.Text = Y
        Text6.SelStart = 0
        Text6.SelLength = X
        Text6.SetFocus
    End If
End If


End Sub

Private Sub txt_email_LostFocus()
txt_email.Text = LCase(txt_email.Text)
End Sub
