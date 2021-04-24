VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_CadastrodeClientes 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   6150
   ClientLeft      =   1410
   ClientTop       =   3795
   ClientWidth     =   9975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6477.602
   ScaleMode       =   0  'User
   ScaleWidth      =   10005.09
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame_cadastroClientes 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton CbtSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CbtIncluir 
         Caption         =   "&Incluir"
         Height          =   375
         Left            =   120
         TabIndex        =   4
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
      Begin VB.CommandButton CbtAlterar 
         Caption         =   "&Alterar"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DTGridClientes 
      Bindings        =   "FRM_CadastrodeClientes.frx":0000
      Height          =   5295
      Left            =   0
      TabIndex        =   0
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "RAZAO_SOCIAL"
         Caption         =   "RAZAO_SOCIAL"
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
         DataField       =   "FANTASIA"
         Caption         =   "FANTASIA"
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
         DataField       =   "CGC"
         Caption         =   "CGC"
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
         DataField       =   "IE"
         Caption         =   "IE"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
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
      BeginProperty Column07 
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
      BeginProperty Column08 
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
      BeginProperty Column09 
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
      BeginProperty Column10 
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
      BeginProperty Column11 
         DataField       =   "OBS"
         Caption         =   "OBS"
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
            ColumnWidth     =   1745,154
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
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1745,154
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1745,154
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRM_CadastrodeClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CbtAlterar_Click()
If deb_coty.rsSel_clientes.State = 1 Then deb_coty.rsSel_clientes.Close
    deb_coty.Sel_clientes
    
    If deb_coty.rsSel_clientes.RecordCount = 0 Then
         MsgBox "Não Há Registros a serem Alterados", vbExclamation, "CLIENTES"
        Exit Sub
    End If



FRM_FichadeClientes.Show

With deb_coty.rsSel_clientes
FRM_FichadeClientes.txtcgc.Text = .Fields("CGC")
FRM_FichadeClientes.txtRazao.Text = .Fields("RAZAO_SOCIAL")
FRM_FichadeClientes.txt_bairro.Text = .Fields("BAIRRO")
FRM_FichadeClientes.txt_endereco.Text = .Fields("ENDERECO")
FRM_FichadeClientes.txt_celular.Text = .Fields("CELULAR")
FRM_FichadeClientes.txt_cidade.Text = .Fields("CIDADE")
FRM_FichadeClientes.txt_contato.Text = .Fields("CONTATO")
FRM_FichadeClientes.txt_email.Text = .Fields("EMAIL")
FRM_FichadeClientes.txt_fone.Text = .Fields("FONE")
FRM_FichadeClientes.txt_obs.Text = .Fields("OBS")
FRM_FichadeClientes.txtFantasia.Text = .Fields("FANTASIA")
FRM_FichadeClientes.txtIE.Text = .Fields("IE")



End With

FRM_FichadeClientes.cmd_altera.Visible = True

End Sub

Private Sub CbtExcluir_Click()
Dim xcgc As String
Dim xnome As String


If deb_coty.rsSel_clientes.State = 1 Then deb_coty.rsSel_clientes.Close
    deb_coty.Sel_clientes
    
    If deb_coty.rsSel_clientes.RecordCount = 0 Then
         MsgBox "Não Há Registros a serem Removidos", vbExclamation, "CLIENTES"
        Exit Sub
    End If

xcgc = deb_coty.rsSel_clientes.Fields("CGC")
xnome = deb_coty.rsSel_clientes.Fields("razao_social")


If MsgBox("Deseja excluir este Cliente: " & xnome, vbYesNo, "CLIENTES") = vbYes Then
    deb_coty.Del_Clientes xcgc
    
    deb_coty.Del_Clientes xcgc
    
    If deb_coty.rsSel_clientes.State = 1 Then deb_coty.rsSel_clientes.Close
        deb_coty.Sel_clientes
    
        DTGridClientes.DataMember = "Sel_clientes"
        DTGridClientes.Refresh
End If


End Sub

Private Sub CbtIncluir_Click()
FRM_FichadeClientes.Show
DoEvents
FRM_FichadeClientes.txtRazao.SetFocus
Me.Enabled = False

End Sub

Private Sub CbtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

If deb_coty.rsSel_clientes.State = 1 Then deb_coty.rsSel_clientes.Close
    deb_coty.Sel_clientes
    
    DTGridClientes.DataMember = "Sel_clientes"
    DTGridClientes.Refresh
    


End Sub
