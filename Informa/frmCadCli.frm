VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCadCli 
   Caption         =   "Clientes Cadastrados"
   ClientHeight    =   7455
   ClientLeft      =   1740
   ClientTop       =   1560
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7455
   ScaleWidth      =   9030
   Begin VB.CommandButton cmdajustar 
      Caption         =   "ACERTO"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Emails p/ Atendimento (SAC)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      TabIndex        =   36
      Top             =   1920
      Width           =   2895
      Begin VB.TextBox txtEmail4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtEmail5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame fraEmail 
      Caption         =   "Emails p/ Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6000
      TabIndex        =   28
      Top             =   240
      Width           =   2895
      Begin VB.TextBox txtEmail3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtEmail2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtEmail1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5775
      Begin VB.TextBox txtIe 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   34
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox chkEmailFer 
         Caption         =   "Enviar Email Informando Futuros Feriados"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CheckBox chkEmailOco 
         Caption         =   "Enviar Email de Ocorrências p/ Cliente/SAC."
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtTabPrazo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtUf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   13
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtCidade 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtEnd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtRazaoSoc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inscr.Estadual:"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tab. de Prazos:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   4680
         TabIndex        =   14
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lblCGC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000C&
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CGC:"
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame fraConsCli 
      Caption         =   "Consulta Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   8805
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca"
         Height          =   375
         Left            =   7200
         TabIndex        =   24
         Top             =   2760
         Width           =   1335
      End
      Begin VB.OptionButton optBuscaNomeTd 
         Caption         =   "Busca por Nome. No texto Todo."
         Height          =   195
         Left            =   3840
         TabIndex        =   19
         Top             =   3120
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton optBuscaNomeIn 
         Caption         =   "Busca por Nome. Início do Texto."
         Height          =   195
         Left            =   3840
         TabIndex        =   18
         Top             =   2880
         Width           =   3015
      End
      Begin VB.OptionButton optBuscaCGC 
         Caption         =   "Busca por Número de CGC"
         Height          =   195
         Left            =   3840
         TabIndex        =   17
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtBusca 
         Height          =   285
         Left            =   960
         MaxLength       =   25
         TabIndex        =   1
         Top             =   2760
         Width           =   2265
      End
      Begin MSDataGridLib.DataGrid GridConsCli 
         Bindings        =   "frmCadCli.frx":0000
         Height          =   2175
         Left            =   240
         TabIndex        =   2
         Top             =   315
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   3836
         _Version        =   393216
         BackColor       =   8388608
         ForeColor       =   8454143
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "Sel_CadCli"
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "cgc"
            Caption         =   "cgc"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "nome"
            Caption         =   "nome"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "endereco"
            Caption         =   "endereco"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cidade"
            Caption         =   "cidade"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "uf"
            Caption         =   "uf"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "ie"
            Caption         =   "ie"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "prazo"
            Caption         =   "prazo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "env_emailOco"
            Caption         =   "env_emailOco"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "env_emailFer"
            Caption         =   "env_emailFer"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "email1"
            Caption         =   "email1"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "email2"
            Caption         =   "email2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "email3"
            Caption         =   "email3"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "email4"
            Caption         =   "email4"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "email5"
            Caption         =   "email5"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "Cli_Dest"
            Caption         =   "Cli_Dest"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3449,764
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2564,788
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   329,953
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   659,906
            EndProperty
         EndProperty
      End
      Begin VB.Label lblBusca 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   2760
         Width           =   465
      End
   End
   Begin VB.Frame fraGravar 
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   480
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraComandos 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   20
      Top             =   6480
      Width           =   4815
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar Dados"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdInserir 
         Caption         =   "Inserir Novo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCadCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdajustar_Click()

    If de_informa.rsSel_AjusteBuscaNFS.State = 1 Then de_informa.rsSel_AjusteBuscaNFS.Close
    de_informa.Sel_AjusteBuscaNFS
    
    If de_informa.rsSel_AjusteBuscaNFS.RecordCount > 0 Then
    
        Do Until de_informa.rsSel_AjusteBuscaNFS.EOF
            
            If de_informa.rsSel_AjusteBuscaCtr.State = 1 Then de_informa.rsSel_AjusteBuscaCtr.Close
            de_informa.Sel_AjusteBuscaCtr de_informa.rsSel_AjusteBuscaNFS.Fields("remet_cgc"), _
                                            de_informa.rsSel_AjusteBuscaNFS.Fields("numnfnum")
                                            
            If de_informa.rsSel_AjusteBuscaCtr.RecordCount > 0 Then
            
                'verifica se já existe 00
                
                If de_informa.rsSel_AjusteBuscaOcorr00.State = 1 Then de_informa.rsSel_AjusteBuscaOcorr00.Close
                de_informa.Sel_AjusteBuscaOcorr00 de_informa.rsSel_AjusteBuscaNFS.Fields("filialctc")
                
                If de_informa.rsSel_AjusteBuscaOcorr00.RecordCount < 1 Then
                
                    de_informa.cn_informa.BeginTrans
                    
                    de_informa.ins_ocorr4 de_informa.rsSel_AjusteBuscaNFS.Fields("filialctc"), _
                                            de_informa.rsSel_AjusteBuscaNFS.Fields("data"), _
                                            de_informa.rsSel_AjusteBuscaNFS.Fields("remet_cgc"), _
                                            "99", "OUTROS TIPOS DE OCORRÊNCIAS", de_informa.rsSel_AjusteBuscaNFS.Fields("data"), _
                                            "00:00", xusuario, de_informa.rsSel_AjusteBuscaNFS.Fields("data")
                                            
                    de_informa.alt_obs_ocorr de_informa.rsSel_AjusteBuscaNFS.Fields("filialctc"), _
                                            "99", de_informa.rsSel_AjusteBuscaNFS.Fields("data"), _
                                            "00:00", "CTC EMITIDO SOMENTE PARA COBRANCA"
                                            
                    de_informa.ins_ocorr4cod00 de_informa.rsSel_AjusteBuscaNFS.Fields("filialctc"), _
                                            de_informa.rsSel_AjusteBuscaNFS.Fields("data"), _
                                            de_informa.rsSel_AjusteBuscaNFS.Fields("remet_cgc"), _
                                            "00", "CTC/NF  B A I X A D O", de_informa.rsSel_AjusteBuscaNFS.Fields("data"), _
                                            "00:00", xusuario, de_informa.rsSel_AjusteBuscaNFS.Fields("data")
                                            
                    de_informa.alt_temocorr_sn "0", de_informa.rsSel_AjusteBuscaNFS.Fields("filialctc")
                    
                    de_informa.cn_informa.CommitTrans
                    
                End If
            
            End If
            
            de_informa.rsSel_AjusteBuscaNFS.MoveNext
    
        Loop
    End If
    
    
    MsgBox "FIM"

End Sub

Private Sub cmdAlterar_Click()
    If Mid$(xdireitos, 6, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        fraComandos.Visible = False
        fraGravar.Visible = True
        txtRazaoSoc.Enabled = True
        txtRazaoSoc.BackColor = &HC0FFFF
        txtEnd.Enabled = True
        txtEnd.BackColor = &HC0FFFF
        txtCidade.Enabled = True
        txtCidade.BackColor = &HC0FFFF
        txtUf.Enabled = True
        txtUf.BackColor = &HC0FFFF
        txtTabPrazo.Enabled = True
        txtTabPrazo.BackColor = &HC0FFFF
        chkEmailFer.Enabled = True
        chkEmailOco.Enabled = True
        txtEmail1.Enabled = True
        txtEmail1.BackColor = &HC0FFFF
        txtEmail2.Enabled = True
        txtEmail2.BackColor = &HC0FFFF
        txtEmail3.Enabled = True
        txtEmail3.BackColor = &HC0FFFF
        txtEmail4.Enabled = True
        txtEmail4.BackColor = &HC0FFFF
        txtEmail5.Enabled = True
        txtEmail5.BackColor = &HC0FFFF
        fraConsCli.Enabled = False
    End If
End Sub

Private Sub cmdBusca_Click()
'    If txtBusca.Text = "" Then
'    Else
        If de_informa.rsSel_ConsCadCliNome.State = 1 Then de_informa.rsSel_ConsCadCliNome.Close
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        If optBuscaCGC = True Then
            de_informa.Sel_ConsCadCli Trim(txtBusca) & "%"
            GridConsCli.DataMember = "sel_consCadCli"
            GridConsCli.Refresh
        ElseIf optBuscaNomeIn = True Then
            de_informa.Sel_ConsCadCliNome Trim(txtBusca) & "%"
            GridConsCli.DataMember = "sel_consCadCliNome"
            GridConsCli.Refresh
        ElseIf optBuscaNomeTd = True Then
            de_informa.Sel_ConsCadCliNome "%" & Trim(txtBusca) & "%"
            GridConsCli.DataMember = "sel_consCadCliNome"
            GridConsCli.Refresh
        End If
'    End If
    lblcgc.Caption = ""
    txtRazaoSoc.Text = ""
    txtEnd.Text = ""
    txtCidade.Text = ""
    txtUf.Text = ""
    txtTabPrazo.Text = ""
    cmdAlterar.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    fraComandos.Visible = True
    fraConsCli.Enabled = True
    fraGravar.Visible = False
    txtRazaoSoc.Enabled = False
    txtRazaoSoc.BackColor = &H8000000E
    txtEnd.Enabled = False
    txtEnd.BackColor = &H8000000E
    txtCidade.Enabled = False
    txtCidade.BackColor = &H8000000E
    txtUf.Enabled = False
    txtUf.BackColor = &H8000000E
    txtTabPrazo.Enabled = False
    txtTabPrazo.BackColor = &H8000000E
    txtEmail1.Enabled = False
    txtEmail1.BackColor = &H8000000E
    txtEmail2.Enabled = False
    txtEmail2.BackColor = &H8000000E
    txtEmail3.Enabled = False
    txtEmail3.BackColor = &H8000000E
    txtEmail4.Enabled = False
    txtEmail4.BackColor = &H8000000E
    txtEmail5.Enabled = False
    txtEmail5.BackColor = &H8000000E
    chkEmailOco.Enabled = False
    chkEmailFer.Enabled = False
    lblcgc.Caption = GridConsCli.Columns(0)
    txtRazaoSoc.Text = GridConsCli.Columns(1)
    txtEnd.Text = GridConsCli.Columns(2)
    txtCidade.Text = GridConsCli.Columns(3)
    txtUf.Text = GridConsCli.Columns(4)
    txtIe.Text = GridConsCli.Columns(5)
    txtTabPrazo.Text = GridConsCli.Columns(6)
    If GridConsCli.Columns(7) = "S" Then
        chkEmailOco.Value = 1
    Else
        chkEmailOco.Value = 0
    End If
    If GridConsCli.Columns(8) = "S" Then
        chkEmailFer.Value = 1
    Else
        chkEmailFer.Value = 0
    End If
    txtEmail1.Text = GridConsCli.Columns(9)
    txtEmail2.Text = GridConsCli.Columns(10)
    txtEmail3.Text = GridConsCli.Columns(11)
    txtEmail4.Text = GridConsCli.Columns(12)
    txtEmail5.Text = GridConsCli.Columns(13)
    cmdAlterar.Enabled = True
End Sub

Private Sub cmdGravar_Click()
    Dim xenv_emailOco As String, xenv_emailFer As String
    
    'if txtEnd.Text = "" Then txtEnd.Text = " "
    'If txtUf.Text = "" Then txtUf.Text = " "
    If chkEmailOco.Value = 1 Then
        xenv_emailOco = "S"
    Else
        xenv_emailOco = "N"
    End If
    If chkEmailFer.Value = 1 Then
        xenv_emailFer = "S"
    Else
        xenv_emailFer = "N"
    End If
    de_informa.alt_cadcli lblcgc.Caption, txtRazaoSoc.Text, txtEnd.Text, txtCidade.Text, txtUf.Text, txtIe, txtTabPrazo.Text, xenv_emailOco, xenv_emailFer, txtEmail1, txtEmail2, txtEmail3, txtEmail4, txtEmail5
    'de_informa.alt_cadcliEmail Mid(lblCGC, 1, 8) & "%", xenv_email, txtEmail1, txtEmail2, txtEmail3, txtEmail4, txtEmail5
    fraComandos.Visible = True
    txtRazaoSoc.Enabled = False
    fraConsCli.Enabled = True
    txtRazaoSoc.BackColor = &H8000000E
    txtEnd.Enabled = False
    txtEnd.BackColor = &H8000000E
    txtCidade.Enabled = False
    txtCidade.BackColor = &H8000000E
    txtUf.Enabled = False
    txtUf.BackColor = &H8000000E
    txtTabPrazo.Enabled = False
    txtTabPrazo.BackColor = &H8000000E
    txtEmail1.Enabled = False
    txtEmail1.BackColor = &H8000000E
    txtEmail2.Enabled = False
    txtEmail2.BackColor = &H8000000E
    txtEmail3.Enabled = False
    txtEmail3.BackColor = &H8000000E
    txtEmail4.Enabled = False
    txtEmail4.BackColor = &H8000000E
    txtEmail5.Enabled = False
    txtEmail5.BackColor = &H8000000E
    chkEmailOco.Enabled = False
    chkEmailFer.Enabled = False
    'If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
    'de_informa.Sel_ConsCadCli "%"
    If de_informa.rsSel_ConsCadCliNome.State = 1 Then de_informa.rsSel_ConsCadCliNome.Close
    If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
    If optBuscaCGC = True Then
        de_informa.Sel_ConsCadCli Trim(txtBusca) & "%"
        GridConsCli.DataMember = "sel_consCadCli"
        GridConsCli.Refresh
    ElseIf optBuscaNomeIn = True Then
        de_informa.Sel_ConsCadCliNome Trim(txtBusca) & "%"
        GridConsCli.DataMember = "sel_consCadCliNome"
        GridConsCli.Refresh
    ElseIf optBuscaNomeTd = True Then
        de_informa.Sel_ConsCadCliNome "%" & Trim(txtBusca) & "%"
        GridConsCli.DataMember = "sel_consCadCliNome"
        GridConsCli.Refresh
    End If
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "CAD. DE CLIENTES: " & txtRazaoSoc
    
    DoEvents
    cmdCancelar_Click
    fraComandos.Visible = True
    fraGravar.Visible = False
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdInserir_Click()
    If Mid$(xdireitos, 5, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        MsgBox "Em Desenvolvimento !"
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    
    GridConsCli.DataMember = ""
    GridConsCli.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmCadCli = Nothing
End Sub

Private Sub GridConsCli_Click()
    lblcgc.Caption = GridConsCli.Columns(0)
    txtRazaoSoc.Text = GridConsCli.Columns(1)
    txtEnd.Text = GridConsCli.Columns(2)
    txtCidade.Text = GridConsCli.Columns(3)
    txtUf.Text = GridConsCli.Columns(4)
    txtIe.Text = GridConsCli.Columns(5)
    txtTabPrazo.Text = GridConsCli.Columns(6)
    If GridConsCli.Columns(7) = "S" Then
        chkEmailOco.Value = 1
    Else
        chkEmailOco.Value = 0
    End If
    If GridConsCli.Columns(8) = "S" Then
        chkEmailFer.Value = 1
    Else
        chkEmailFer.Value = 0
    End If
    txtEmail1.Text = GridConsCli.Columns(9)
    txtEmail2.Text = GridConsCli.Columns(10)
    txtEmail3.Text = GridConsCli.Columns(11)
    txtEmail4.Text = GridConsCli.Columns(12)
    txtEmail5.Text = GridConsCli.Columns(13)
    cmdAlterar.Enabled = True
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "CONSULTA", xusuario, "CAD. DE CLIENTES: " & txtRazaoSoc
    
End Sub

Private Sub optBuscaCGC_Click()
    lblBusca.Caption = "CGC:"
End Sub
Private Sub optBuscaNomeIn_Click()
    lblBusca.Caption = "Nome:"
End Sub
Private Sub optBuscaNomeTd_Click()
    lblBusca.Caption = "Nome:"
End Sub
Private Sub txtBuscaNome_Change()
    txtBuscaNome.SelStart = 0
    txtBuscaNome.SelLength = 25
End Sub
Private Sub txtBusca_GotFocus()
    txtBusca.SelStart = 0
    txtBusca.SelLength = 20
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEmail1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEmail1_LostFocus()
    txtEmail1 = LCase(txtEmail1)
End Sub

Private Sub txtEmail2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEmail2_LostFocus()
    txtEmail2 = LCase(txtEmail2)
End Sub

Private Sub txtEmail3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEmail3_LostFocus()
    txtEmail3 = LCase(txtEmail3)
End Sub

Private Sub txtEmail4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEmail4_LostFocus()
    txtEmail4 = LCase(txtEmail4)
End Sub

Private Sub txtEmail5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEmail5_LostFocus()
    txtEmail5 = LCase(txtEmail5)
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtIe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRazaoSoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtTabPrazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
