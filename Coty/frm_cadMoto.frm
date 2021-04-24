VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cadMoto 
   Caption         =   "CADASTRO DE MOTOS / MOTOBOYS"
   ClientHeight    =   8175
   ClientLeft      =   1905
   ClientTop       =   1155
   ClientWidth     =   11775
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   4440
         TabIndex        =   30
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Frame fra_moto 
         Caption         =   "CADASTRO DE MOTOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   11535
         Begin VB.CommandButton cmd_AtualizaCods 
            Caption         =   "ATUALIZA"
            Height          =   255
            Left            =   2640
            TabIndex        =   54
            Top             =   360
            Width           =   975
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   300
            Left            =   4440
            TabIndex        =   52
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "99/99"
            PromptChar      =   "_"
         End
         Begin VB.Frame Frame2 
            Height          =   1815
            Left            =   9600
            TabIndex        =   49
            Top             =   1560
            Width           =   1815
            Begin VB.CommandButton cmd_inserir 
               Caption         =   "&INSERIR"
               Height          =   255
               Left            =   240
               TabIndex        =   50
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox txt_ufplaca 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   7920
            TabIndex        =   23
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txt_cidadeplaca 
            Height          =   285
            Left            =   4440
            TabIndex        =   22
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txt_numplaca 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   21
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txt_letraplaca 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   20
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txt_modelo 
            Height          =   285
            Left            =   1560
            TabIndex        =   19
            Top             =   1320
            Width           =   1695
         End
         Begin VB.ComboBox cmb_marcas 
            Height          =   315
            ItemData        =   "frm_cadMoto.frx":0000
            Left            =   1560
            List            =   "frm_cadMoto.frx":0025
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   840
            Width           =   1695
         End
         Begin VB.ComboBox cmb_motoboy 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "ANO:"
            Height          =   255
            Left            =   3720
            TabIndex        =   51
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "UF:"
            Height          =   255
            Left            =   7560
            TabIndex        =   46
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label18 
            Caption         =   "CIDADE:"
            Height          =   255
            Left            =   3720
            TabIndex        =   45
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "-"
            Height          =   255
            Left            =   2280
            TabIndex        =   44
            Top             =   1800
            Width           =   135
         End
         Begin VB.Label Label16 
            Caption         =   "PLACA:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "MODELO:"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "MARCA:"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "COD. MOTOBOY:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1335
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   9600
            Picture         =   "frm_cadMoto.frx":008C
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame fra_motoboy 
         Caption         =   "CADASTRO DE MOTOBOYS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   11535
         Begin VB.CommandButton cma_atualiza 
            Caption         =   "ATUALIZA"
            Height          =   255
            Left            =   7440
            TabIndex        =   53
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txt_celular 
            Height          =   285
            Left            =   3840
            TabIndex        =   11
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox txt_fone 
            Height          =   285
            Left            =   840
            TabIndex        =   10
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox txt_cat 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   14
            Top             =   2760
            Width           =   855
         End
         Begin MSMask.MaskEdBox mask_venc 
            Height          =   300
            Left            =   3840
            TabIndex        =   13
            Top             =   2760
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txt_cnh 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   840
            TabIndex        =   12
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox txt_uf 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   9
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox txt_cidade 
            Height          =   285
            Left            =   3840
            TabIndex        =   8
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txt_bairro 
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Frame Frame4 
            Height          =   1575
            Left            =   9600
            TabIndex        =   33
            Top             =   1560
            Width           =   1815
            Begin VB.CommandButton cmd_procura 
               Caption         =   "&Procurar"
               Height          =   255
               Left            =   240
               TabIndex        =   16
               Top             =   600
               Width           =   1335
            End
            Begin VB.CommandButton cmd_gravar 
               Caption         =   "&Gravar"
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox txt_numero 
            Height          =   285
            Left            =   6360
            TabIndex        =   6
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txt_end 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   1320
            Width           =   4815
         End
         Begin VB.TextBox txt_rg 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            TabIndex        =   4
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txt_cpf 
            Height          =   285
            Left            =   840
            TabIndex        =   3
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txt_cod 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txt_motoboy 
            Height          =   285
            Left            =   840
            TabIndex        =   1
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label Label21 
            Caption         =   "CELULAR:"
            Height          =   255
            Left            =   3000
            TabIndex        =   48
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "FONE:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "CAT.:"
            Height          =   255
            Left            =   5880
            TabIndex        =   39
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "VENC.:"
            Height          =   255
            Left            =   3120
            TabIndex        =   38
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "CNH:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "UF:"
            Height          =   255
            Left            =   5880
            TabIndex        =   36
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "CIDADE:"
            Height          =   255
            Left            =   3000
            TabIndex        =   35
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "BAIRRO:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Nº:"
            Height          =   255
            Left            =   5880
            TabIndex        =   32
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "END.:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "RG:"
            Height          =   255
            Left            =   3120
            TabIndex        =   29
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "CPF:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "COD.:"
            Height          =   255
            Left            =   5760
            TabIndex        =   27
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "NOME:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   9600
            Picture         =   "frm_cadMoto.frx":1653F
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   11640
         Y1              =   3840
         Y2              =   3840
      End
   End
End
Attribute VB_Name = "frm_cadMoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cma_atualiza_Click()

deb_coty.rsSel_codMaxMotoboy.Open

    If IsNull(deb_coty.rsSel_codMaxMotoboy.Fields("COD")) Then

        frm_cadMoto.txt_cod.Text = 1

    Else
    
        frm_cadMoto.txt_cod.Text = deb_coty.rsSel_codMaxMotoboy.Fields("COD")
    
    End If

deb_coty.rsSel_codMaxMotoboy.Close

destrava_tela (1)
cmd_gravar.Enabled = True
txt_motoboy.SetFocus



End Sub

Private Sub cmd_AtualizaCods_Click()
deb_coty.rsSel_CodsMotoboyMotos.Open
If deb_coty.rsSel_CodsMotoboyMotos.RecordCount > 0 Then
    If IsNull(deb_coty.rsSel_CodsMotoboyMotos) = True Then

        frm_cadMoto.cmb_motoboy.Enabled = False
    
    Else
        If deb_coty.rsSel_CodsMotoboyMotos.RecordCount > 0 Then
            deb_coty.rsSel_CodsMotoboyMotos.MoveFirst

            Do Until deb_coty.rsSel_CodsMotoboyMotos.EOF
    
                cmb_motoboy.AddItem deb_coty.rsSel_CodsMotoboyMotos.Fields("COD2")
        
                deb_coty.rsSel_CodsMotoboyMotos.MoveNext
        
            Loop
        End If
    End If

Else

MsgBox "Não Há Registro de MotoBoys", vbInformation, "MOTOBOYS"

End If

deb_coty.rsSel_CodsMotoboyMotos.Close

cmb_motoboy.SetFocus

End Sub

Private Sub cmd_gravar_Click()
Dim xnome As String

If Len(Trim$(txt_motoboy.Text)) = 0 Then
    
    MsgBox "Digite o nome do Motoboy", vbInformation, "MOTOBOY"
    txt_motoboy.SetFocus
    Exit Sub

ElseIf Len(Trim$(txt_cpf.Text)) = 0 Then

    MsgBox "Digite o CPF", vbInformation, "CPF"
    txt_cpf.SetFocus
    Exit Sub
    
ElseIf Len(Trim$(txt_end.Text)) = 0 Then
    
    MsgBox "Digite o Endereço do Motoboy", vbInformation, "ENDEREÇO"
    txt_end.SetFocus
    Exit Sub
    
ElseIf Len(Trim$(txt_cnh.Text)) = 0 Then

    MsgBox "Digite o Número da 'CNH' do Motoboy", vbInformation, "CNH - CARTEIRA NACIONAL DE HABILITAÇÃO"
    txt_cnh.SetFocus
    Exit Sub

End If

deb_coty.in_motoboy UCase(txt_motoboy.Text), txt_cpf.Text, txt_rg.Text, UCase(txt_end.Text) & "," & txt_numero.Text, UCase(txt_cidade.Text), UCase(txt_uf.Text), _
                    txt_fone.Text, txt_celular.Text, txt_cnh.Text, CDate(mask_venc.Text), txt_cat.Text, txt_bairro.Text

xnome = branco(txt_motoboy.Text)


MsgBox "MOTOBOY: " & xnome & Chr$(13) & Chr$(13) & "CÓDIGO:   " & txt_cod.Text & Chr$(13) & Chr$(13) & "CADASTRADO", vbInformation, "CADASTRADO"

limpa_tela (1)

deb_coty.rsSel_codMaxMotoboy.Open

    If IsNull(deb_coty.rsSel_codMaxMotoboy.Fields("COD")) Then

        frm_cadMoto.txt_cod.Text = 1

    Else
    
        frm_cadMoto.txt_cod.Text = deb_coty.rsSel_codMaxMotoboy.Fields("COD")
    
    End If

deb_coty.rsSel_codMaxMotoboy.Close

mask_venc.Mask = ""
mask_venc.Text = ""
mask_venc.Mask = "99/99/9999"

txt_motoboy.SetFocus





End Sub

Private Sub cmd_sair_Click()

MDIForm1.Toolbar1.Enabled = True

Unload Me

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

