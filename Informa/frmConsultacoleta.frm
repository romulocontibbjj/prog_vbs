VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConsultaColeta 
   Caption         =   "Consulta de Coletas"
   ClientHeight    =   7695
   ClientLeft      =   570
   ClientTop       =   570
   ClientWidth     =   11115
   Icon            =   "frmConsultacoleta.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7695
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Dados do Remetente / Local de Coleta e Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Left            =   120
      TabIndex        =   5
      Top             =   1500
      Width           =   10875
      Begin VB.CommandButton CmdAlterarObservacao 
         Caption         =   "Alterar Obs."
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   55
         Top             =   1860
         Width           =   1155
      End
      Begin VB.CommandButton CmdGravarObservacao 
         Caption         =   "Gravar Obs."
         Enabled         =   0   'False
         Height          =   375
         Left            =   9600
         TabIndex        =   54
         Top             =   2340
         Width           =   1155
      End
      Begin VB.Frame Frame3 
         Caption         =   "Confirmação / Baixa / Ocorrência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   120
         TabIndex        =   48
         Top             =   2880
         Width           =   10635
         Begin VB.TextBox TxtObsOcorr 
            Enabled         =   0   'False
            Height          =   795
            Left            =   120
            TabIndex        =   49
            Top             =   2100
            Width           =   10335
         End
         Begin MSDataGridLib.DataGrid DataGridOcorr 
            Bindings        =   "frmConsultacoleta.frx":000C
            Height          =   1395
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   10395
            _ExtentX        =   18336
            _ExtentY        =   2461
            _Version        =   393216
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
            DataMember      =   "ColetaSelOcorr"
            ColumnCount     =   10
            BeginProperty Column00 
               DataField       =   "coditem"
               Caption         =   "coditem"
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
               DataField       =   "filialcoleta"
               Caption         =   "filialcoleta"
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
               DataField       =   "emissaocoleta"
               Caption         =   "emissaocoleta"
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
               DataField       =   "cod_ocorr"
               Caption         =   "Código"
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
               DataField       =   "descricao"
               Caption         =   "Descricao"
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
               DataField       =   "data"
               Caption         =   "Data"
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
               DataField       =   "hora"
               Caption         =   "Hora"
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
               DataField       =   "obs"
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
            BeginProperty Column08 
               DataField       =   "usuario"
               Caption         =   "Usuario"
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
               DataField       =   "data_usu"
               Caption         =   "Data_Inc_Sis"
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
               MarqueeStyle    =   3
               BeginProperty Column00 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   734,74
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   3044,977
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Observação / Cometário:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1860
            Width           =   1785
         End
      End
      Begin VB.TextBox TxtValMerc 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   6060
         TabIndex        =   40
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtPeso 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         TabIndex        =   39
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtVolumes 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   9780
         TabIndex        =   38
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtNFs 
         Enabled         =   0   'False
         Height          =   615
         Left            =   5700
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   900
         Width           =   5055
      End
      Begin VB.TextBox TxtOBS 
         Enabled         =   0   'False
         Height          =   975
         Left            =   5340
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox TxtEspecie 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9240
         TabIndex        =   35
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox TxtNatureza 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7140
         TabIndex        =   34
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox TxtTipoFrete 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   32
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox TxtContato 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   2460
         Width           =   1815
      End
      Begin VB.TextBox TxtHoraColeta 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   28
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox TxtDataColeta 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox TxtSolicitante 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Top             =   1860
         Width           =   3615
      End
      Begin VB.TextBox TxtCidadeDes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox TxtCidadeRem 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   1260
         Width           =   3615
      End
      Begin VB.TextBox TxtEnderecoRem 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox TxtRemetente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   645
         Width           =   3615
      End
      Begin VB.TextBox TxtCGCRem 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "OBS:"
         Height          =   195
         Left            =   5340
         TabIndex        =   45
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Merc:"
         Height          =   195
         Left            =   5340
         TabIndex        =   44
         Top             =   645
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Peso:"
         Height          =   195
         Left            =   7440
         TabIndex        =   43
         Top             =   645
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Volumes:"
         Height          =   195
         Left            =   9060
         TabIndex        =   42
         Top             =   645
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "NFs:"
         Height          =   195
         Left            =   5340
         TabIndex        =   41
         Top             =   900
         Width           =   330
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Natureza/Espécie:"
         Height          =   195
         Left            =   5670
         TabIndex        =   33
         Top             =   345
         Width           =   1335
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Frete:"
         Height          =   195
         Left            =   3480
         TabIndex        =   31
         Top             =   2505
         Width           =   765
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
         Height          =   195
         Left            =   825
         TabIndex        =   29
         Top             =   2505
         Width           =   600
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Horário:"
         Height          =   195
         Left            =   3000
         TabIndex        =   27
         Top             =   2205
         Width           =   555
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Data de Coleta:"
         Height          =   195
         Left            =   315
         TabIndex        =   24
         Top             =   2205
         Width           =   1110
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante:"
         Height          =   195
         Left            =   645
         TabIndex        =   23
         Top             =   1905
         Width           =   780
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Destino/UF:"
         Height          =   195
         Left            =   555
         TabIndex        =   21
         Top             =   1605
         Width           =   870
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Cidade/UF:"
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   1305
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   690
         TabIndex        =   17
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Rementente:"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   690
         Width           =   915
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CGC Rementente:"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   405
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Ordem de Coleta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   10875
      Begin VB.CommandButton CmdReimpressao 
         Caption         =   "Re-Imprimir Coleta"
         Enabled         =   0   'False
         Height          =   555
         Left            =   3120
         TabIndex        =   52
         Top             =   600
         Width           =   1275
      End
      Begin VB.Frame Frame4 
         Height          =   915
         Left            =   4500
         TabIndex        =   46
         Top             =   240
         Width           =   3435
         Begin VB.Label LblStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1620
            TabIndex        =   47
            Top             =   360
            Width           =   105
         End
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "Sair"
         Height          =   315
         Left            =   8280
         TabIndex        =   3
         Top             =   780
         Width           =   2235
      End
      Begin VB.TextBox TxtPrioridade 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9060
         TabIndex        =   12
         Top             =   435
         Width           =   1455
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   300
         Width           =   1275
      End
      Begin VB.TextBox TxtEmissor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   900
         Width           =   1635
      End
      Begin VB.TextBox TxtColeta 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1860
         MaxLength       =   8
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
      Begin VB.TextBox TxtEmissao 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox TxtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   0
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   735
         TabIndex        =   53
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Prioridade:"
         Height          =   195
         Left            =   8280
         TabIndex        =   11
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   135
         Left            =   1800
         TabIndex        =   10
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emissor:"
         Height          =   195
         Left            =   780
         TabIndex        =   7
         Top             =   645
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial/Coleta:"
         Height          =   195
         Left            =   495
         TabIndex        =   6
         Top             =   360
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmConsultaColeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAlterarObservacao_Click()
TxtOBS.BackColor = xamarelo1
TxtOBS.Enabled = True
TxtOBS.SetFocus
CmdGravarObservacao.Enabled = True
DoEvents
End Sub

Private Sub CmdBuscar_Click()
    If Len(Trim(TxtFilial.Text)) = 0 Then
    MsgBox "Filial incorreta.", vbCritical, ""
    Exit Sub
    ElseIf Len(Trim(TxtColeta.Text)) = 0 Then
    MsgBox "Coleta incorreta.", vbCritical, ""
    Exit Sub
    End If

xFilialColeta = transctc(TxtFilial.Text, TxtColeta.Text)

If de_informa.rsColetaSel.State = 1 Then de_informa.rsColetaSel.Close
de_informa.ColetaSel xFilialColeta

    If de_informa.rsColetaSel.RecordCount = 0 Then
    MsgBox "Coleta não encontrada!", vbCritical, ""
    Exit Sub
    End If
    
If de_informa.rsColetaSelData.State = 1 Then de_informa.rsColetaSelData.Close
de_informa.ColetaSelData xFilialColeta
    
If de_informa.rsColetaSelOcorr.State = 1 Then de_informa.rsColetaSelOcorr.Close
de_informa.ColetaSelOcorr xFilialColeta

Set DataGridOcorr.DataSource = de_informa
DataGridOcorr.DataMember = "ColetaSelOcorr"
DataGridOcorr.Refresh

TxtCGCRem.Text = de_informa.rsColetaSel.Fields("remetcgc")
TxtRemetente.Text = PriMaiuscula(de_informa.rsColetaSel.Fields("remetnome"))
TxtEnderecoRem.Text = PriMaiuscula(de_informa.rsColetaSel.Fields("remetend"))
TxtCidadeRem.Text = PriMaiuscula(de_informa.rsColetaSel.Fields("remetcidade")) & "/" & de_informa.rsColetaSel.Fields("remetuf")
TxtCidadeDes.Text = PriMaiuscula(de_informa.rsColetaSel.Fields("destcidade")) & "/" & de_informa.rsColetaSel.Fields("destuf")
TxtSolicitante.Text = PriMaiuscula(de_informa.rsColetaSel.Fields("solempnome"))
TxtContato.Text = PriMaiuscula(de_informa.rsColetaSel.Fields("solnome"))
TxtNatureza.Text = PriMaiuscula(de_informa.rsColetaSel.Fields("natureza"))
TxtEspecie.Text = PriMaiuscula(de_informa.rsColetaSel.Fields("especie"))
TxtValMerc.Text = Format(de_informa.rsColetaSel.Fields("valmerc"), "##,##0.00")
TxtPeso.Text = Format(de_informa.rsColetaSel.Fields("peso"), "#0.0")
TxtVolumes.Text = de_informa.rsColetaSel.Fields("volumes")
TxtNFs.Text = de_informa.rsColetaSel.Fields("nfs")
TxtOBS.Text = de_informa.rsColetaSel.Fields("obs")
TxtEmissao = de_informa.rsColetaSel.Fields("dataemissao")
TxtEmissor = de_informa.rsColetaSel.Fields("emissor")
TxtTipoFrete = de_informa.rsColetaSel.Fields("tipofrete")
TxtPrioridade = de_informa.rsColetaSel.Fields("prioridade")
TxtDataColeta.Text = de_informa.rsColetaSel.Fields("datacoleta")
TxtHoraColeta.Text = de_informa.rsColetaSel.Fields("horacoleta")

    If de_informa.rsColetaSel.Fields("tem_ocorr") = "N" Then
    lblStatus.Caption = "Sem Posição"
    ElseIf de_informa.rsColetaSel.Fields("tem_ocorr") = "1" Then
    lblStatus.Caption = "Entregue/Concluída"
    ElseIf de_informa.rsColetaSel.Fields("tem_ocorr") = "2" Then
    lblStatus.Caption = "Em Ocorrência"
    ElseIf de_informa.rsColetaSel.Fields("tem_ocorr") = "0" Then
    lblStatus.Caption = "Coleta Baixada"
    ElseIf de_informa.rsColetaSel.Fields("tem_ocorr") = "C" Then
    lblStatus.Caption = "Coleta Cancelada"
    End If

'LOG DE USUÁRIO
de_informa.ins_LogUsuario "CONSULTA", xusuario, "CONSULTA COLETA: " & xFilialColeta

CmdReimpressao.Enabled = True
CmdAlterarObservacao.Enabled = True


End Sub


Private Sub CmdGravarObservacao_Click()
CmdGravarObservacao.Enabled = False
de_informa.ColetaAltOBS Trim(TxtOBS.Text), transctc(TxtFilial.Text, TxtColeta.Text)
TxtOBS.Enabled = False
TxtOBS.BackColor = xbranco
DoEvents

End Sub

Private Sub CmdReimpressao_Click()
Me.MousePointer = 11
DoEvents
Call ImprimeColeta(transctc(TxtFilial.Text, TxtColeta.Text))
Me.MousePointer = 0
DoEvents
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGridOcorr_Click()
TxtObsOcorr.Text = Trim(UCase(DataGridOcorr.Columns(4)))
End Sub

Private Sub DataGridOcorr_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DataGridOcorr_Click
End Sub

Private Sub Form_Activate()
    If Len(Trim(TxtFilial.Text)) > 0 And Len(Trim(TxtColeta.Text)) > 0 Then
    CmdBuscar_Click
    End If
End Sub

Private Sub TxtColeta_Change()
If Len(TxtColeta.Text) = 8 Then SendKeys "{TAB}"
End Sub

Private Sub TxtColeta_GotFocus()
Dim xTextBox As TextBox
Set xTextBox = TxtColeta
xTextBox.SelStart = 0
xTextBox.SelLength = 500
End Sub

Private Sub txtcoleta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub


Private Sub TxtFilial_Change()
If Len(TxtFilial.Text) = 2 Then SendKeys "{TAB}"
End Sub

Private Sub TxtFilial_GotFocus()
Dim xTextBox As TextBox
Set xTextBox = TxtFilial
xTextBox.SelStart = 0
xTextBox.SelLength = 500
End Sub

Private Sub txtfilial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub


Private Sub TxtOBS_Change()
X = TxtOBS.SelStart
TxtOBS.Text = UCase(TxtOBS.Text)
TxtOBS.SelStart = X
End Sub

Private Sub TxtObsOcorr_Change()
Dim X As Integer
X = TxtObsOcorr.SelStart
TxtObsOcorr.Text = UCase(TxtObsOcorr.Text)
TxtObsOcorr.SelStart = X
End Sub

Private Sub TxtObsOcorr_GotFocus()
Dim xTextBox As TextBox
Set xTextBox = TxtFilial
xTextBox.SelStart = 0
xTextBox.SelLength = 500
End Sub

Private Sub TxtObsOcorr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Public Sub TravaFrame(xtela As Form, xFrame As frame, Tipo As Integer)
    Dim xcontrol As Control
    For Each xcontrol In xtela
        If TypeOf xcontrol Is frame And xcontrol <> xFrame And Tipo = 0 Then
        xcontrol.Enabled = False
        ElseIf TypeOf xcontrol Is frame And xcontrol <> xFrame And Tipo = 1 Then
        xcontrol.Enabled = True
        End If
    Next
End Sub

Public Sub OrdenaFlexGrid(xFlexGrid As MSFlexGrid, xColuna As Integer)

Dim xlinha(0 To 1000) As String
Dim X As Integer
Dim I As Integer
Dim J As Integer

    For I = 1 To xFlexGrid.Rows - 1
        For J = 1 To xFlexGrid.Rows - 2
            If xFlexGrid.TextMatrix(I, xColuna) < xFlexGrid.TextMatrix(J, xColuna) Then
                For X = 0 To xFlexGrid.Cols - 1
                xlinha(X) = xFlexGrid.TextMatrix(J, X)
                Next
                
                For X = 0 To xFlexGrid.Cols - 1
                xFlexGrid.TextMatrix(J, X) = xFlexGrid.TextMatrix(I, X)
                Next
            
                For X = 0 To xFlexGrid.Cols - 1
                xFlexGrid.TextMatrix(I, X) = xlinha(X)
                Next
            End If
        Next
    Next
End Sub


Public Sub ImprimeColeta(XColeta As String)
    Dim xVIAs As Integer
    Dim xLin As Integer
    Dim ximpr_cfg As String, ximpr_inst As Printer
    Dim xlinha As String
    Dim X As Double
    Dim Y As Double
    
    
    'busca impressora para este documento
    If Dir(App.Path & "\coletaimp.cfg") <> "" Then
        
        Open App.Path & "\coletaimp.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "COL" Then
                ximpr_cfg = Trim$(Mid$(xlinha, 5))
                Exit Do
            End If
        Loop
        
        Close #1
        
    Else
    
        MsgBox "Não está Configurado a Impressora para Este Documento: Coleta " & transctc(TxtFilial.Text, TxtColeta.Text)
        Exit Sub
        
    End If
    
    'seta impressora
    
    For Each ximpr_inst In Printers
        If ximpr_inst.DeviceName = ximpr_cfg Then
            Set Printer = ximpr_inst
            DoEvents
            Exit For
        End If
    Next
    
    'BUSCA A MINUTA A SER IMPRESSAO
    
    If de_informa.rsColetaSel.State = 1 Then de_informa.rsColetaSel.Close
    de_informa.ColetaSel XColeta
    
    If de_informa.rsColetaSel.RecordCount < 1 Then
        MsgBox "Coleta Para Impressão Inexistente!"
        Exit Sub
    End If
    
Dim xIMPColeta As String
Dim xIMPDataEmissao As String
Dim xIMPHoraEmissao As String
Dim xIMPEmissor As String
Dim xIMPSolEmpNome As String
Dim xIMPSolContato As String
Dim xIMPSolTel As String
Dim xIMPNomeRem As String
Dim xIMPEndRem As String
Dim xIMPCidadeUfRem As String
Dim xIMPNomeDest As String
Dim xIMPEndDest As String
Dim xIMPCidadeUfDest As String
Dim xIMPLocalColeta As String
Dim xIMPDataColeta As String
Dim xIMPHoraColeta As String
Dim xIMPVolumes As String
Dim xIMPPeso As String
Dim xIMPEspecie As String
Dim xIMPNatureza As String
Dim xIMPTipoFrete As String

Dim xIMPVolumesT As String
Dim xIMPPesoT As String
Dim xIMPEspecieT As String
Dim xIMPNaturezaT As String
Dim xIMPTipoFreteT As String


Dim xIMPNfs As String
Dim xIMPPrioridade As String
Dim xIMPOBServacoes As String
Dim xAux As String

Dim xIMPNfs1 As String
Dim xIMPNfs2 As String
Dim xIMPNfs3 As String
Dim xIMPNfs4 As String
Dim xIMPNfs5 As String

Dim xIMPOBServacoes1 As String
Dim xIMPOBServacoes2 As String
Dim xIMPOBServacoes3 As String
Dim xIMPOBServacoes4 As String
Dim xIMPOBServacoes5 As String

xAux = " "

xIMPColeta = Mid(XColeta, 1, 2) & "-" & Mid(XColeta, 3)
xIMPDataEmissao = de_informa.rsColetaSel.Fields("dataemissao")
xIMPHoraEmissao = de_informa.rsColetaSel.Fields("horaemissao")
xIMPEmissor = de_informa.rsColetaSel.Fields("emissor")
xIMPSolEmpNome = de_informa.rsColetaSel.Fields("SOLempnome")
xIMPSolContato = de_informa.rsColetaSel.Fields("SOLnome")
xIMPSolTel = de_informa.rsColetaSel.Fields("SOLtel")
xIMPNomeRem = de_informa.rsColetaSel.Fields("remetnome")
xIMPEndRem = de_informa.rsColetaSel.Fields("remetend")
xIMPCidadeUfRem = de_informa.rsColetaSel.Fields("remetcidade") & " - " & de_informa.rsColetaSel.Fields("remetuf")
xIMPNomeDest = de_informa.rsColetaSel.Fields("destnome")
xIMPEndDest = de_informa.rsColetaSel.Fields("destend")
xIMPCidadeUfDest = de_informa.rsColetaSel.Fields("destcidade") & " - " & de_informa.rsColetaSel.Fields("destuf")
xIMPLocalColeta = "LOCAL DE COLETA: " & de_informa.rsColetaSel.Fields("coletaend") & " - " & de_informa.rsColetaSel.Fields("coletacidade") & " - " & de_informa.rsColetaSel.Fields("coletauf")
xIMPDataColeta = "DATA A COLETAR: " & de_informa.rsColetaSel.Fields("datacoleta")
xIMPHoraColeta = "HORARIO: " & de_informa.rsColetaSel.Fields("horacoleta")

xIMPVolumesT = "VOLs."
xIMPPesoT = "PESO"
xIMPEspecieT = "ESPECIE"
xIMPNaturezaT = "NATUREZA"
xIMPTipoFreteT = "FRETE"


xIMPVolumes = de_informa.rsColetaSel.Fields("volumes")
xIMPPeso = de_informa.rsColetaSel.Fields("peso")
xIMPEspecie = de_informa.rsColetaSel.Fields("especie")
xIMPNatureza = de_informa.rsColetaSel.Fields("natureza")
xIMPTipoFrete = IIf(de_informa.rsColetaSel.Fields("tipofrete") = "1", "PAGO", "A PAGAR")
xIMPNfs = de_informa.rsColetaSel.Fields("NFS")
xIMPPrioridade = de_informa.rsColetaSel.Fields("prioridade")
xIMPOBServacoes = de_informa.rsColetaSel.Fields("OBS")

xIMPColeta = Mid(xIMPColeta, 1, 12) & String(12 - Len(Mid(xIMPColeta, 1, 12)), xAux)
xIMPDataEmissao = Mid(xIMPDataEmissao, 1, 11) & String(11 - Len(Mid(xIMPDataEmissao, 1, 11)), xAux)
xIMPHoraEmissao = Mid(xIMPHoraEmissao, 1, 11) & String(11 - Len(Mid(xIMPHoraEmissao, 1, 11)), xAux)
xIMPEmissor = Mid(xIMPEmissor, 1, 11) & String(11 - Len(Mid(xIMPEmissor, 1, 11)), xAux)

xIMPSolEmpNome = Mid(xIMPSolEmpNome, 1, 78) & String(78 - Len(Mid(xIMPSolEmpNome, 1, 78)), xAux)
xIMPSolContato = Mid(xIMPSolContato, 1, 78) & String(78 - Len(Mid(xIMPSolContato, 1, 78)), xAux)
xIMPSolTel = Mid(xIMPSolTel, 1, 78) & String(78 - Len(Mid(xIMPSolTel, 1, 78)), xAux)
xIMPNomeRem = Mid(xIMPNomeRem, 1, 78) & String(78 - Len(Mid(xIMPNomeRem, 1, 78)), xAux)
xIMPEndRem = Mid(xIMPEndRem, 1, 78) & String(78 - Len(Mid(xIMPEndRem, 1, 78)), xAux)
xIMPCidadeUfRem = Mid(xIMPCidadeUfRem, 1, 78) & String(78 - Len(Mid(xIMPCidadeUfRem, 1, 78)), xAux)
xIMPNomeDest = Mid(xIMPNomeDest, 1, 78) & String(78 - Len(Mid(xIMPNomeDest, 1, 78)), xAux)
xIMPEndDest = Mid(xIMPEndDest, 1, 78) & String(78 - Len(Mid(xIMPEndDest, 1, 78)), xAux)
xIMPCidadeUfDest = Mid(xIMPCidadeUfDest, 1, 78) & String(78 - Len(Mid(xIMPCidadeUfDest, 1, 78)), xAux)
'xIMPLocalColeta = Mid(xIMPLocalColeta, 1, 92) & String(92 - Len(Mid(xIMPLocalColeta, 1, 92)), xAux)

xIMPLocalColeta = String((92 - Len(Mid(xIMPLocalColeta, 1, 92))) / 2, xAux) & Mid(xIMPLocalColeta, 1, 92) & String((92 - Len(Mid(xIMPLocalColeta, 1, 92))) / 2, xAux)
xIMPDataColeta = String((45 - Len(Mid(xIMPDataColeta, 1, 45))) / 2, xAux) & Mid(xIMPDataColeta, 1, 45) & String((45 - Len(Mid(xIMPDataColeta, 1, 45))) / 2, xAux)
xIMPHoraColeta = String((44 - Len(Mid(xIMPHoraColeta, 1, 44))) / 2, xAux) & Mid(xIMPHoraColeta, 1, 44) & String((44 - Len(Mid(xIMPHoraColeta, 1, 44))) / 2, xAux)
xIMPVolumes = String(15 - Len(Mid(xIMPVolumes, 1, 15)), xAux) & Mid(xIMPVolumes, 1, 15)
xIMPPeso = String(12 - Len(Mid(xIMPPeso, 1, 12)), xAux) & Mid(xIMPPeso, 1, 12)
xIMPEspecie = Mid(xIMPEspecie, 1, 22) & String(22 - Len(Mid(xIMPEspecie, 1, 22)), xAux)
xIMPNatureza = Mid(xIMPNatureza, 1, 22) & String(22 - Len(Mid(xIMPNatureza, 1, 22)), xAux)
xIMPTipoFrete = Mid(xIMPTipoFrete, 1, 9) & String(9 - Len(Mid(xIMPTipoFrete, 1, 9)), xAux)

xIMPVolumesT = String((15 - Len(Mid(xIMPVolumesT, 1, 15))) / 2, xAux) & Mid(xIMPVolumesT, 1, 15) & String((15 - Len(Mid(xIMPVolumesT, 1, 15))) / 2, xAux)
xIMPPesoT = String((12 - Len(Mid(xIMPPesoT, 1, 12))) / 2, xAux) & Mid(xIMPPesoT, 1, 12) & String((12 - Len(Mid(xIMPPesoT, 1, 12))) / 2, xAux)
xIMPEspecieT = String((22 - Len(Mid(xIMPEspecieT, 1, 22))) / 2, xAux) & Mid(xIMPEspecieT, 1, 22) & String((22 - Len(Mid(xIMPEspecieT, 1, 22))) / 2, xAux)
xIMPNaturezaT = String((22 - Len(Mid(xIMPNaturezaT, 1, 22))) / 2, xAux) & Mid(xIMPNaturezaT, 1, 22) & String((22 - Len(Mid(xIMPNaturezaT, 1, 22))) / 2, xAux)
xIMPTipoFreteT = String((9 - Len(Mid(xIMPTipoFreteT, 1, 9))) / 2, xAux) & Mid(xIMPTipoFreteT, 1, 9) & String((9 - Len(Mid(xIMPTipoFreteT, 1, 9))) / 2, xAux)


xIMPNfs = Mid(xIMPNfs, 1, 310) & String(310 - Len(Mid(xIMPNfs, 1, 310)), xAux)
xIMPPrioridade = Mid(xIMPPrioridade, 1, 10) & String(10 - Len(Mid(xIMPPrioridade, 1, 10)), xAux)
xIMPOBServacoes = Mid(xIMPOBServacoes, 1, 310) & String(310 - Len(Mid(xIMPOBServacoes, 1, 310)), xAux)
xIMPNfs1 = Mid(xIMPNfs, 1, 58)
xIMPNfs2 = Mid(xIMPNfs, 59, 63)
xIMPNfs3 = Mid(xIMPNfs, 122, 63)
xIMPNfs4 = Mid(xIMPNfs, 185, 63)
xIMPNfs5 = Mid(xIMPNfs, 248, 63)
xIMPOBServacoes1 = Mid(xIMPOBServacoes, 1, 58)
xIMPOBServacoes2 = Mid(xIMPOBServacoes, 59, 63)
xIMPOBServacoes3 = Mid(xIMPOBServacoes, 122, 63)
xIMPOBServacoes4 = Mid(xIMPOBServacoes, 185, 63)
xIMPOBServacoes5 = Mid(xIMPOBServacoes, 248, 63)

frmLogo.Text1.Text = XColeta
    
    For xVIAs = 1 To 1  'DUAS VIAS
    
        If xVIAs = 1 Then
        xLin = 0
        Printer.DrawStyle = 0
        Printer.ForeColor = &H80000008  'PRETO
        'Printer.DrawWidth = 8
        Printer.DrawMode = 9
        ElseIf xVIAs = 2 Then
        xLin = 149
        Printer.DrawStyle = 0
        Printer.ForeColor = &H80000008  'PRETO
        'Printer.DrawWidth = 8
        Printer.DrawMode = 9
        End If
    Printer.CurrentX = 0
    Printer.CurrentY = 0 + xLin
    
    Printer.FontName = "Courier New"
    Printer.FontSize = 3
    Printer.Print
    Printer.FontSize = 6
    Printer.Print Spc(24); "INTEC-Integração Nacional de Transportes de Encom. e Cargas Ltda"
    Printer.FontSize = 1
    Printer.Print ""
    Printer.FontSize = 6
    Printer.Print Spc(24); "AV. MARG. DIREITA DO RIO TIETÊ, 504 - BARUERI/SP - CEP 06455-050"
    Printer.FontSize = 1
    Printer.Print ""
    Printer.FontSize = 6
    Printer.Print Spc(24); "CNPJ: 52.134.798-0001-68         INSCR.ESTADUAL: 206.182.910.118"
    Printer.FontSize = 1
    Printer.Print ""
    Printer.FontSize = 6
    Printer.Print Spc(24); "TELEFONES: (11) 4689-7575 / 4193-5921           www.intec.com.br"
            
    'GRAFICOS
    Printer.ForeColor = &H80000008   'PRETO
    Printer.ScaleMode = vbMillimeters
    Printer.Line (0, 0 + xLin)-(198, 15 + xLin), , B
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (113, 0 + xLin)-(198, 15 + xLin), , B 'QUADRO MINUTA DE TRANSP, NUMERO DA MINUTA DATA, EMISSOR, ETC
    Printer.FontName = "ARIAL"
    Printer.FontSize = 30
    Printer.FontBold = True
    Printer.CurrentX = 130
    Printer.CurrentY = 2 + xLin
    Printer.Print "COLETA"
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.ForeColor = &HC0C0C0     'CINZA
    Printer.Line (0, 15 + xLin)-(198, 20 + xLin), , BF 'QUADRO MINUTA DE TRANSP, NUMERO DA MINUTA DATA, EMISSOR, ETC
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (0, 15 + xLin)-(198, 20 + xLin), , B 'QUADRO MINUTA DE TRANSP, NUMERO DA MINUTA DATA, EMISSOR, ETC
    Printer.CurrentX = 0
    Printer.CurrentY = 15.5 + xLin
    Printer.FontBold = True
    Printer.Print " Numero: " & xIMPColeta & "   DATA: " & xIMPDataEmissao & "   HORA: " & xIMPHoraEmissao & "   EMISSOR: " & xIMPEmissor & "   " & Trim(Str(xVIAs)) & "ª VIA"
    Printer.FontBold = False
    
    
    Printer.PaintPicture frmLogo.piclogo.Picture, 1, 1 + xLin, frmLogo.piclogo.Picture.Width * 0.0013, frmLogo.piclogo.Picture.Height * 0.0013
    Printer.PaintPicture frmLogo.Picture1, 140, 113 + xLin, frmLogo.Picture1.Picture.Width * 0.0068, frmLogo.Picture1.Picture.Height * 0.008
    
    
    Printer.Line (0, 20 + xLin)-(198, 33 + xLin), , B
    Printer.Line (0, 33 + xLin)-(198, 49 + xLin), , B
    Printer.Line (0, 49 + xLin)-(198, 63 + xLin), , B
    Printer.ForeColor = &HC0C0C0     'CINZA
    Printer.Line (0, 63 + xLin)-(198, 66.5 + xLin), , BF
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (0, 63 + xLin)-(198, 66.5 + xLin), , B
    Printer.ForeColor = &HC0C0C0     'CINZA
    Printer.Line (0, 66.5 + xLin)-(198, 70.5 + xLin), , BF
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (0, 66.5 + xLin)-(198, 70.5 + xLin), , B
    Printer.Line (0, 70.5 + xLin)-(37, 74.5 + xLin), , B
    Printer.Line (37, 70.5 + xLin)-(69, 74.5 + xLin), , B
    Printer.Line (69, 70.5 + xLin)-(121, 74.5 + xLin), , B
    Printer.Line (121, 70.5 + xLin)-(175, 74.5 + xLin), , B
    Printer.Line (175, 70.5 + xLin)-(198, 74.5 + xLin), , B
    
    Printer.Line (0, 74.5 + xLin)-(37, 78.5 + xLin), , B
    Printer.Line (37, 74.5 + xLin)-(69, 78.5 + xLin), , B
    Printer.Line (69, 74.5 + xLin)-(121, 78.5 + xLin), , B
    Printer.Line (121, 74.5 + xLin)-(175, 78.5 + xLin), , B
    Printer.Line (175, 74.5 + xLin)-(198, 78.5 + xLin), , B
    
    Printer.Line (0, 78.5 + xLin)-(138, 97 + xLin), , B
        If UCase(Trim(xIMPPrioridade)) = "URGENTE" Then
        Printer.ForeColor = &HC0C0C0     'CINZA
        Printer.Line (138, 78.5 + xLin)-(198, 97 + xLin), , BF
        Printer.ForeColor = &H80000008   'PRETO
        End If
    Printer.Line (138, 78.5 + xLin)-(198, 97 + xLin), , B
    
    Printer.Line (0, 97 + xLin)-(138, 116 + xLin), , B
    Printer.Line (138, 97 + xLin)-(198, 138 + xLin), , B
    
    Printer.ForeColor = &HC0C0C0     'CINZA
    Printer.Line (0, 116 + xLin)-(138, 120 + xLin), , BF
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (0, 116 + xLin)-(138, 120 + xLin), , B
    
    Printer.Line (0, 120 + xLin)-(138, 138 + xLin), , B
    
    Printer.CurrentX = 0
    Printer.CurrentY = 20.5 + xLin
    
    Printer.Print " SOLICITANTE : "; xIMPSolEmpNome
    Printer.Print " CONTATO     : "; xIMPSolContato
    Printer.Print " TELEFONE    : "; xIMPSolTel
    Printer.Print ""
    Printer.Print " REMETENTE   : "; xIMPNomeRem
    Printer.Print " ENDERECO    : "; xIMPEndRem
    Printer.Print " CIDADE-UF   : "; xIMPCidadeUfRem
    Printer.Print ""
    Printer.Print " DESTINATARIO: "; xIMPNomeDest
    Printer.Print " ENDERECO    : "; xIMPEndDest
    Printer.Print " CIDADE-UF   : "; xIMPCidadeUfDest
    Printer.FontSize = 2
    Printer.Print ""
    Printer.FontSize = 10
    Printer.FontBold = True
    'Printer.Print " LOCAL DE COLETA: "
    Printer.Print " "; xIMPLocalColeta
    Printer.Print " "; xIMPDataColeta & "   " & xIMPHoraColeta
    Printer.Print " "; xIMPVolumesT; "   "; xIMPPesoT; "   "; xIMPEspecieT; "  "; xIMPNaturezaT; "   "; xIMPTipoFreteT
    Printer.FontBold = False
    Printer.Print " "; xIMPVolumes; "   "; xIMPPeso; "   "; xIMPEspecie; "   "; xIMPNatureza; "   "; xIMPTipoFrete
    Printer.Print " "; "NFs: " & xIMPNfs1 '& "  Prioridade:"
    Printer.Print " "; xIMPNfs2
    Printer.Print " "; xIMPNfs3
    Printer.Print " "; xIMPNfs4
    Printer.Print " "; xIMPNfs5
    Printer.Print " "; "OBS: " & xIMPOBServacoes1
    Printer.Print " "; xIMPOBServacoes2
    Printer.Print " "; xIMPOBServacoes3
    Printer.Print " "; xIMPOBServacoes4
    Printer.Print " "; xIMPOBServacoes5
    Printer.FontBold = True
    Printer.Print "                          R E C E B I M E N T O "
    Printer.FontSize = 8
    Printer.Print " NOME:"
    Printer.Print ""
    Printer.Print " Nº RG:"
    Printer.Print "                                                          ____________________"
    Printer.Print " DATA/HORA:                                                    ASSINATURA"
    
    X = Printer.CurrentX
    Y = Printer.CurrentY
    Printer.FontName = "ARIAL"
    Printer.FontSize = 25
    Printer.FontBold = True
    Printer.CurrentX = 140
    Printer.CurrentY = 83 + xLin
    Printer.Print xIMPPrioridade
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = X
    Printer.CurrentY = Y
    
    Next
    
    Printer.EndDoc
    
End Sub

