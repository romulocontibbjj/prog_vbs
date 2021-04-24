VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_consulta 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   1365
   ClientTop       =   1710
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin TabDlg.SSTab SSTab1 
         Height          =   7815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   13785
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "CLIENTES"
         TabPicture(0)   =   "frm_consulta.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "MOTOBOYS"
         TabPicture(1)   =   "frm_consulta.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "frm_consulta.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.Frame Frame4 
            Height          =   6975
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   11055
            Begin VB.CommandButton cmd_alteradados 
               Caption         =   "&ALTERAR"
               Height          =   375
               Left            =   9360
               TabIndex        =   47
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Frame Frame6 
               Caption         =   "STATUS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   240
               TabIndex        =   43
               Top             =   4320
               Width           =   7455
               Begin VB.Label lab_staus 
                  Alignment       =   2  'Center
                  Height          =   255
                  Left            =   120
                  TabIndex        =   44
                  Top             =   240
                  Width           =   7215
               End
            End
            Begin VB.TextBox txt_cat 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   7080
               MaxLength       =   2
               TabIndex        =   42
               Top             =   3840
               Width           =   495
            End
            Begin VB.TextBox txt_venc 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   4320
               MaxLength       =   10
               TabIndex        =   40
               Top             =   3840
               Width           =   2295
            End
            Begin VB.TextBox txt_cnh 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               TabIndex        =   38
               Top             =   3840
               Width           =   2295
            End
            Begin VB.TextBox txt_celular 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4320
               TabIndex        =   36
               Top             =   3360
               Width           =   2295
            End
            Begin VB.TextBox txt_fone 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               TabIndex        =   34
               Top             =   3360
               Width           =   2295
            End
            Begin VB.CommandButton cmd_novoMotoboy 
               Caption         =   "&INSERIR"
               Height          =   375
               Left            =   9360
               TabIndex        =   32
               Top             =   1920
               Width           =   1455
            End
            Begin VB.CommandButton cmd_refresh_cod 
               Caption         =   "COD"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2280
               TabIndex        =   31
               Top             =   840
               Width           =   615
            End
            Begin VB.CommandButton cmd_refresh_nome 
               Caption         =   "Nomes"
               Height          =   255
               Left            =   7680
               TabIndex        =   30
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txt_uf 
               Height          =   285
               Left            =   7080
               MaxLength       =   2
               TabIndex        =   27
               Top             =   2880
               Width           =   495
            End
            Begin VB.TextBox txt_cidade 
               Height          =   285
               Left            =   4320
               TabIndex        =   25
               Top             =   2880
               Width           =   2295
            End
            Begin VB.TextBox txt_bairro 
               Height          =   285
               Left            =   1200
               TabIndex        =   23
               Top             =   2880
               Width           =   2295
            End
            Begin VB.TextBox txt_endereco 
               Height          =   285
               Left            =   1200
               TabIndex        =   21
               Top             =   2400
               Width           =   6375
            End
            Begin VB.ComboBox cmb_codigos 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txt_rg 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4680
               TabIndex        =   16
               Top             =   1920
               Width           =   2895
            End
            Begin VB.TextBox TXT_CGC 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1200
               TabIndex        =   15
               Top             =   1920
               Width           =   2895
            End
            Begin VB.Frame Frame5 
               Caption         =   "PESQUISA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Left            =   9240
               TabIndex        =   12
               Top             =   240
               Width           =   1695
               Begin VB.OptionButton opt_CodigoMotoboy 
                  Caption         =   "CÓDIGO"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   14
                  Top             =   720
                  Width           =   975
               End
               Begin VB.OptionButton opt_nomeMotoboy 
                  Caption         =   "NOME"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   13
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1095
               End
            End
            Begin VB.ComboBox cmb_motoboy 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   360
               Width           =   6375
            End
            Begin VB.Label Label15 
               Height          =   255
               Left            =   240
               TabIndex        =   46
               Top             =   1560
               Width           =   375
            End
            Begin VB.Label Label14 
               BackColor       =   &H00800000&
               Caption         =   "Label14"
               Height          =   255
               Left            =   1320
               TabIndex        =   45
               Top             =   5280
               Width           =   1695
            End
            Begin VB.Label Label13 
               Caption         =   "CAT:"
               Height          =   255
               Left            =   6720
               TabIndex        =   41
               Top             =   3840
               Width           =   375
            End
            Begin VB.Label Label12 
               Caption         =   "VENC:"
               Height          =   255
               Left            =   3600
               TabIndex        =   39
               Top             =   3840
               Width           =   495
            End
            Begin VB.Label Label11 
               Caption         =   "CNH:"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   3840
               Width           =   975
            End
            Begin VB.Label Label10 
               Caption         =   "CEL.:"
               Height          =   255
               Left            =   3600
               TabIndex        =   35
               Top             =   3360
               Width           =   615
            End
            Begin VB.Label Label9 
               Caption         =   "FONE:"
               Height          =   255
               Left            =   240
               TabIndex        =   33
               Top             =   3360
               Width           =   735
            End
            Begin VB.Label lab_uf 
               Caption         =   "UF:"
               Height          =   255
               Left            =   6720
               TabIndex        =   26
               Top             =   2880
               Width           =   375
            End
            Begin VB.Label Label8 
               Caption         =   "CIDADE:"
               Height          =   255
               Left            =   3600
               TabIndex        =   24
               Top             =   2880
               Width           =   735
            End
            Begin VB.Label Label7 
               Caption         =   "BAIRRO:"
               Height          =   255
               Left            =   240
               TabIndex        =   22
               Top             =   2880
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "ENDEREÇO:"
               Height          =   255
               Left            =   240
               TabIndex        =   20
               Top             =   2400
               Width           =   975
            End
            Begin VB.Line Line1 
               X1              =   240
               X2              =   9120
               Y1              =   1440
               Y2              =   1440
            End
            Begin VB.Label Label5 
               Caption         =   "CÓDIGO:"
               Height          =   255
               Left            =   240
               TabIndex        =   18
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label4 
               Caption         =   "RG:"
               Height          =   255
               Left            =   4200
               TabIndex        =   17
               Top             =   1920
               Width           =   495
            End
            Begin VB.Label Label3 
               Caption         =   "CPF:"
               Height          =   255
               Left            =   240
               TabIndex        =   11
               Top             =   1920
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "MOTOBOY:"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   2
            Top             =   600
            Width           =   11055
            Begin VB.CommandButton cmd_novoCliente 
               Caption         =   "&Novo"
               Height          =   375
               Left            =   9360
               TabIndex        =   29
               Top             =   1920
               Width           =   1455
            End
            Begin VB.CommandButton cmd_sair 
               Caption         =   "&SAIR"
               Height          =   375
               Left            =   9360
               TabIndex        =   28
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Frame Frame3 
               Caption         =   "PESQUISA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1455
               Left            =   9240
               TabIndex        =   5
               Top             =   240
               Width           =   1695
               Begin VB.OptionButton opt_cgc 
                  Caption         =   "CNPJ / CGC"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   7
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.OptionButton opt_nome 
                  Caption         =   "NOME"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   6
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1215
               End
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   1200
               TabIndex        =   4
               Text            =   "Combo1"
               Top             =   360
               Width           =   6375
            End
            Begin VB.Label Label1 
               Caption         =   "CLIENTE:"
               Height          =   255
               Left            =   240
               TabIndex        =   3
               Top             =   360
               Width           =   735
            End
         End
      End
   End
End
Attribute VB_Name = "frm_consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmb_motoboy_Click()

With deb_coty.rsSel_DadosMotoboy

If .State = 1 Then .Close
    deb_coty.Sel_DadosMotoboy cmb_motoboy.Text
    
    If .Fields("ativa") = "S" Then
        lab_staus.ForeColor = &H80&
        lab_staus.Caption = "ATIVO"
    Else
        lab_staus.ForeColor = &H800000
        lab_staus.Caption = "INATIVO"
    End If
    
        
    txt_cat.Text = .Fields("CATEGORIA")
    txt_bairro.Text = .Fields("BAIRRO")
    txt_celular.Text = .Fields("CELULAR")
    TXT_CGC.Text = .Fields("CPF")
    txt_cidade.Text = .Fields("CIDADE")
    txt_cnh.Text = .Fields("CNH")
    txt_endereco.Text = .Fields("ENDERECO")
    txt_fone.Text = .Fields("FONE")
    txt_rg.Text = .Fields("RG")
    txt_uf.Text = .Fields("UF")
    txt_venc.Text = .Fields("VENCIMENTO")
    
    
    
    'cmb_codigos.Text = .Fields("cod_motoboy")
    
    
End With


    
    
    
    
    
    

End Sub

Private Sub cmd_alteradados_Click()

destrava_tela (2)


End Sub

Private Sub cmd_novoCliente_Click()

frm_cadclientes.Show
Unload Me


End Sub

Private Sub cmd_novoMotoboy_Click()

frm_cadMoto.Show
MDIForm1.Toolbar1.Enabled = False
frm_cadMoto.fra_moto.Visible = False
frm_cadMoto.fra_moto.Enabled = False
frm_cadMoto.cmd_gravar.Enabled = False
trava_tela (1)
Unload Me




End Sub

Private Sub cmd_refresh_cod_Click()

deb_coty.rsSel_CodMotoboy.Open
deb_coty.rsSel_CodMotoboy.MoveFirst

    Do Until deb_coty.rsSel_CodMotoboy.EOF
        
        cmb_codigos.AddItem deb_coty.rsSel_CodMotoboy.Fields("COD")
        
    Loop
    
deb_coty.rsSel_CodMotoboy.Close
        


End Sub

Private Sub cmd_refresh_nome_Click()

deb_coty.rsSel_NomeMotoboy.Open
deb_coty.rsSel_NomeMotoboy.MoveFirst
    
    Do Until deb_coty.rsSel_NomeMotoboy.EOF
    
        cmb_motoboy.AddItem deb_coty.rsSel_NomeMotoboy.Fields("nome")
        deb_coty.rsSel_NomeMotoboy.MoveNext
        
    Loop
    
deb_coty.rsSel_NomeMotoboy.Close


End Sub

Private Sub cmd_sair_Click()
MDIForm1.Toolbar1.Enabled = True

Unload Me

End Sub


Private Sub opt_CodigoMotoboy_Click()
cmd_refresh_nome.Enabled = False
cmd_refresh_cod.Enabled = True
cmb_motoboy.Enabled = False
cmb_codigos.Enabled = True

End Sub

Private Sub opt_nomeMotoboy_Click()
cmd_refresh_nome.Enabled = True
cmd_refresh_cod.Enabled = False
cmb_motoboy.Enabled = True
cmb_codigos.Enabled = False

End Sub



