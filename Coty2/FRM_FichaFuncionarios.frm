VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FRM_FichadeMotoqueiros 
   AutoRedraw      =   -1  'True
   Caption         =   "Ficha de Motoqueiros"
   ClientHeight    =   5505
   ClientLeft      =   2355
   ClientTop       =   3240
   ClientWidth     =   9180
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535.165
   ScaleMode       =   0  'User
   ScaleWidth      =   13092.79
   Begin VB.Frame FrameDadosCadastrais 
      Caption         =   " Dados Cadastrais de Motoqueiros"
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
      Begin MSMask.MaskEdBox mask_vencimento 
         Height          =   300
         Left            =   5280
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mask_nascimento 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_categoria 
         Height          =   285
         Left            =   6840
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txt_cidade 
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cmb_uf 
         Height          =   315
         ItemData        =   "FRM_FichaFuncionarios.frx":0000
         Left            =   8280
         List            =   "FRM_FichaFuncionarios.frx":0055
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cmb_EstadoCivil 
         Height          =   315
         ItemData        =   "FRM_FichaFuncionarios.frx":00C5
         Left            =   1680
         List            =   "FRM_FichaFuncionarios.frx":00D5
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txt_cnh 
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txt_rg 
         Height          =   285
         Left            =   4200
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txt_cpf 
         Height          =   285
         Left            =   6000
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txt_end 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox txt_bairro 
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txt_contato 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txt_fone 
         Height          =   285
         Left            =   2400
         TabIndex        =   14
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txt_celular 
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txt_email 
         Height          =   285
         Left            =   6000
         TabIndex        =   16
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txt_obs 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   3600
         Width           =   8895
      End
      Begin VB.Frame Frame_cadastroClientes 
         Height          =   735
         Left            =   6480
         TabIndex        =   20
         Top             =   4560
         Width           =   2535
         Begin VB.CommandButton cmd_altera 
            Caption         =   "&Alterar"
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton CbtExcluir 
            Caption         =   "OK"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CbtSair 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label17 
         Caption         =   "Categoria"
         Height          =   255
         Left            =   6840
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "UF"
         Height          =   255
         Left            =   8520
         TabIndex        =   36
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "Data de Vencimento"
         Height          =   255
         Left            =   5280
         TabIndex        =   35
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "CNH"
         Height          =   255
         Left            =   3480
         TabIndex        =   34
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "Data de Nascimento"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Est. Civil"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "RG"
         Height          =   255
         Left            =   4200
         TabIndex        =   30
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "CPF"
         Height          =   255
         Left            =   6000
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Endereço"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   4200
         TabIndex        =   27
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   6000
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Contato"
         Height          =   255
         Left            =   120
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
      Begin VB.Label Label10 
         Caption         =   "Celular"
         Height          =   255
         Left            =   4200
         TabIndex        =   23
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "E-Mail"
         Height          =   255
         Left            =   6000
         TabIndex        =   22
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Obs"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Width           =   495
      End
   End
End
Attribute VB_Name = "FRM_FichadeMotoqueiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CbtExcluir_Click()

If Trim$(Len(txt_nome.Text)) = 0 Then
    MsgBox "Entre com o Nome do Motoqueiro", vbInformation, "CADASTRO"
    txt_nome.SetFocus
    Exit Sub
ElseIf Trim$(Len(txt_cpf.Text)) = 0 Then
    MsgBox "Entre com o CPF do Motoqueiro", vbInformation, "CADASTRO"
    txt_cpf.SetFocus
    Exit Sub
ElseIf Trim$(Len(txt_cnh.Text)) = 0 Then
    MsgBox "Entre com o número da CNH do Motoqueiro", vbInformation, "CADASTRO"
    txt_nome.SetFocus
    Exit Sub
Else
    
    deb_coty.in_Motoqueiro txt_nome, txt_cpf, txt_rg, txt_end, txt_bairro, _
                            txt_Cidade, cmb_Uf.Text, txt_fone, txt_celular, _
                            txt_cnh, CDate(mask_vencimento), txt_categoria, _
                            CDate(mask_vencimento), cmb_EstadoCivil.Text, txt_contato, _
                            LCase(txt_email), txt_obs
    MsgBox "Motoqueiro : " & txt_nome & Chr$(13) & "Cadastrado com Sucesso", vbInformation, "CADASTRO"
    
    limpa_tela (2)
    FRM_CadastrodeMotoqueiros.Enabled = True
        
    If deb_coty.rsSel_Motoqueiros.State = 1 Then deb_coty.rsSel_Motoqueiros.Close
        deb_coty.Sel_Motoqueiros
        
        FRM_CadastrodeMotoqueiros.dt_motoqueiros.DataMember = "Sel_Motoqueiros"
        FRM_CadastrodeMotoqueiros.dt_motoqueiros.Refresh
    
     Unload Me

End If


End Sub

Private Sub CbtSair_Click()
limpa_tela (2)
FRM_CadastrodeMotoqueiros.Enabled = True
Unload Me
End Sub

Private Sub cmd_altera_Click()

deb_coty.Up_Motoqueiro txt_nome, txt_cpf, txt_rg, txt_end, txt_bairro, _
                            txt_Cidade, cmb_Uf.Text, txt_fone, txt_celular, _
                            txt_cnh, CDate(mask_vencimento), txt_categoria, _
                            CDate(mask_vencimento), cmb_EstadoCivil.Text, txt_contato, _
                            LCase(txt_email), txt_obs, txt_cpf.Text
                            
    limpa_tela (2)
    FRM_CadastrodeMotoqueiros.Enabled = True
        
    If deb_coty.rsSel_Motoqueiros.State = 1 Then deb_coty.rsSel_Motoqueiros.Close
        deb_coty.Sel_Motoqueiros
        
        FRM_CadastrodeMotoqueiros.dt_motoqueiros.DataMember = "Sel_Motoqueiros"
        FRM_CadastrodeMotoqueiros.dt_motoqueiros.Refresh
    
     Unload Me
     
    MsgBox "Alteração Feita com Sucesso", vbInformation, "ALTERAÇÃO"

End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))

End Sub

Private Sub txt_celular_Change()
Dim x As Integer
Dim y As String

If Trim$(Len(txt_fone)) > 0 Then
    
    If IsNumeric(txt_celular.Text) = False Then
    
        MsgBox "Digite O Nº Correto do Celular", vbInformation, "CADASTRO"
            
        y = Mid(txt_celular, 1, Len(txt_celular) - 1)
        txt_celular = y
        txt_celular.SelStart = 0
        txt_celular.SelLength = Len(txt_fone)
        txt_celular.SetFocus
    End If
End If

End Sub

Private Sub txt_fone_Change()
Dim x As Integer
Dim y As String

If Trim$(Len(txt_fone)) > 0 Then
    
    If IsNumeric(txt_fone.Text) = False Then
    
        MsgBox "Digite O Nº Correto do Telefone", vbInformation, "CADASTRO"
            
        y = Mid(txt_fone, 1, Len(txt_fone) - 1)
        txt_fone = y
        txt_fone.SelStart = 0
        txt_fone.SelLength = Len(txt_fone)
        txt_fone.SetFocus
    End If

End If


End Sub
