VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_CadastrodeMotoqueiros 
   Caption         =   "Cadastro de Motoqueiros"
   ClientHeight    =   6150
   ClientLeft      =   1365
   ClientTop       =   2775
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6493.401
   ScaleMode       =   0  'User
   ScaleWidth      =   10005.09
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame_cadastroClientes 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton CbtAlterar 
         Caption         =   "&Alterar"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CbtExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CbtIncluir 
         Caption         =   "&Incluir"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CbtSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid dt_motoqueiros 
      Bindings        =   "FRM_CadastrodeFuncionarios.frx":0000
      Height          =   5295
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9340
      _Version        =   393216
      BackColor       =   8388608
      ForeColor       =   65535
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
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "COD_MOTOBOY"
         Caption         =   "COD_MOTOBOY"
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
         DataField       =   "NOME"
         Caption         =   "NOME"
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
         DataField       =   "CPF"
         Caption         =   "CPF"
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
         DataField       =   "RG"
         Caption         =   "RG"
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
         DataField       =   "NASCIMENTO"
         Caption         =   "NASCIMENTO"
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
         DataField       =   "ENDERECO"
         Caption         =   "ENDERECO"
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
         DataField       =   "BAIRRO"
         Caption         =   "BAIRRO"
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
         DataField       =   "CIDADE"
         Caption         =   "CIDADE"
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
         DataField       =   "UF"
         Caption         =   "UF"
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
         DataField       =   "CNH"
         Caption         =   "CNH"
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
         DataField       =   "VENCIMENTO"
         Caption         =   "VENCIMENTO"
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
         DataField       =   "CATEGORIA"
         Caption         =   "CATEGORIA"
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
         DataField       =   "ESTADO_CIVIL"
         Caption         =   "ESTADO_CIVIL"
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
         DataField       =   "CONTATO"
         Caption         =   "CONTATO"
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
         DataField       =   "FONE"
         Caption         =   "FONE"
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
      BeginProperty Column15 
         DataField       =   "CELULAR"
         Caption         =   "CELULAR"
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
      BeginProperty Column16 
         DataField       =   "EMAIL"
         Caption         =   "EMAIL"
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
            ColumnWidth     =   1323,792
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   541,913
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1008,198
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1293,654
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1745,154
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRM_CadastrodeMotoqueiros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CbtAlterar_Click()
If deb_coty.rsSel_Motoqueiros.State = 1 Then deb_coty.rsSel_Motoqueiros.Close
        deb_coty.Sel_Motoqueiros
    If deb_coty.rsSel_Motoqueiros.RecordCount = 0 Then
        MsgBox "Não Há Registros a serem Alterados", vbExclamation, "MOTOQUEIROS"
        Exit Sub
    End If



FRM_FichadeMotoqueiros.Show
FRM_FichadeMotoqueiros.cmd_altera.Visible = True

If deb_coty.rsSel_PorCodMotoboy.State = 1 Then deb_coty.rsSel_PorCodMotoboy.Close
    deb_coty.Sel_PorCodMotoboy deb_coty.rsSel_Motoqueiros.Fields("COD_MOTOBOY")

With deb_coty.rsSel_PorCodMotoboy
FRM_FichadeMotoqueiros.txt_nome.Text = .Fields("NOME")
FRM_FichadeMotoqueiros.txt_bairro.Text = .Fields("BAIRRO")
FRM_FichadeMotoqueiros.txt_categoria.Text = .Fields("CATEGORIA")
FRM_FichadeMotoqueiros.txt_celular.Text = .Fields("CELULAR")
FRM_FichadeMotoqueiros.txt_cnh.Text = .Fields("CNH")

If IsNull(.Fields("CONTATO")) = False Then

    FRM_FichadeMotoqueiros.txt_contato.Text = .Fields("CONTATO")

End If

FRM_FichadeMotoqueiros.txt_cpf.Text = .Fields("CPF")

If IsNull(.Fields("EMAIL")) = False Then

    FRM_FichadeMotoqueiros.txt_email.Text = .Fields("EMAIL")

End If

FRM_FichadeMotoqueiros.txt_end.Text = .Fields("ENDERECO")
FRM_FichadeMotoqueiros.txt_fone.Text = .Fields("FONE")

If IsNull(.Fields("OBS")) = False Then

    FRM_FichadeMotoqueiros.txt_obs.Text = .Fields("OBS")
End If

FRM_FichadeMotoqueiros.txt_rg.Text = .Fields("RG")
FRM_FichadeMotoqueiros.cmb_EstadoCivil.Text = .Fields("ESTADO_CIVIL")
FRM_FichadeMotoqueiros.cmb_Uf.Text = .Fields("UF")
FRM_FichadeMotoqueiros.mask_nascimento = .Fields("NASCIMENTO")
FRM_FichadeMotoqueiros.mask_vencimento = .Fields("VENCIMENTO")
FRM_FichadeMotoqueiros.txt_Cidade.Text = .Fields("CIDADE")
FRM_FichadeMotoqueiros.txt_bairro.Text = .Fields("BAIRRO")
FRM_FichadeMotoqueiros.txt_cnh.Text = .Fields("CNH")

End With



DoEvents

End Sub

Private Sub CbtExcluir_Click()
Dim xcod As Integer
Dim xnome As String

If deb_coty.rsSel_Motoqueiros.State = 1 Then deb_coty.rsSel_Motoqueiros.Close
        deb_coty.Sel_Motoqueiros
    If deb_coty.rsSel_Motoqueiros.RecordCount = 0 Then
        MsgBox "Não Há Registros a serem Removidos", vbExclamation, "MOTOQUEIROS"
        Exit Sub
    End If
    
    


xnome = deb_coty.rsSel_Motoqueiros.Fields("NOME")
xcod = deb_coty.rsSel_Motoqueiros.Fields("COD_MOTOBOY")
If MsgBox("Deseja Excluir o Motoqueiro:" & xnome, vbQuestion + vbYesNo, "EXLUSÃO") = vbYes Then
    deb_coty.Del_Motoboy xcod
    dt_motoqueiros.Refresh
    If deb_coty.rsSel_Motoqueiros.State = 1 Then deb_coty.rsSel_Motoqueiros.Close
        deb_coty.Sel_Motoqueiros
        dt_motoqueiros.DataMember = "Sel_Motoqueiros"
        dt_motoqueiros.Refresh
End If
End Sub

Private Sub CbtIncluir_Click()
FRM_FichadeMotoqueiros.Show
DoEvents
End Sub

Private Sub CbtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
   If deb_coty.rsSel_Motoqueiros.State = 1 Then deb_coty.rsSel_Motoqueiros.Close
        deb_coty.Sel_Motoqueiros
        FRM_CadastrodeMotoqueiros.dt_motoqueiros.DataMember = "Sel_Motoqueiros"
        FRM_CadastrodeMotoqueiros.dt_motoqueiros.Refresh
End Sub
