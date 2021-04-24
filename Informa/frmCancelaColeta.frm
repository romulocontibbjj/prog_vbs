VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCancelaColeta 
   Caption         =   "Cancelamento de Coletas"
   ClientHeight    =   6870
   ClientLeft      =   675
   ClientTop       =   1095
   ClientWidth     =   11115
   ControlBox      =   0   'False
   Icon            =   "frmCancelaColeta.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6870
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   10875
      Begin VB.TextBox TxtFilialColeta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   60
         TabIndex        =   62
         Top             =   900
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   4500
         TabIndex        =   46
         Top             =   120
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
         Top             =   660
         Width           =   2235
      End
      Begin VB.TextBox TxtPrioridade 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9060
         TabIndex        =   12
         Top             =   315
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
         Left            =   2940
         TabIndex        =   9
         Top             =   600
         Width           =   1455
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
         Width           =   1455
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Prioridade:"
         Height          =   195
         Left            =   8280
         TabIndex        =   11
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   135
         Left            =   3960
         TabIndex        =   10
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emissor/Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial/Coleta:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   870
      End
   End
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
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   10875
      Begin VB.Frame FraCancelamento 
         Caption         =   "Cancelamento de Coleta"
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
         Height          =   2775
         Left            =   5460
         TabIndex        =   52
         Top             =   2640
         Width           =   5295
         Begin VB.TextBox TxtUsuario 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   4020
            TabIndex        =   60
            Top             =   300
            Width           =   1155
         End
         Begin VB.TextBox TxtDataCanc 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   540
            TabIndex        =   57
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox TxtHoraCanc 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   2220
            TabIndex        =   56
            Top             =   300
            Width           =   975
         End
         Begin VB.CommandButton CmdCancelaColeta 
            Caption         =   "Confirmar Cancelamento"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   2340
            Width           =   5055
         End
         Begin VB.TextBox TxtMotivoCanc 
            Enabled         =   0   'False
            Height          =   1335
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   53
            Top             =   960
            Width           =   5055
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Usuário:"
            Height          =   195
            Left            =   3360
            TabIndex        =   61
            Top             =   345
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   345
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   1785
            TabIndex        =   58
            Top             =   345
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Observação / Cometário:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   1785
         End
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
         Height          =   2535
         Left            =   120
         TabIndex        =   48
         Top             =   2880
         Width           =   5235
         Begin VB.TextBox TxtObsOcorr 
            Enabled         =   0   'False
            Height          =   555
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   1860
            Width           =   4995
         End
         Begin MSDataGridLib.DataGrid DataGridOcorr 
            Bindings        =   "frmCancelaColeta.frx":000C
            Height          =   1335
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   2355
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
            Top             =   1620
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
         Left            =   5700
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   1560
         Width           =   5055
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
         Left            =   5280
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
End
Attribute VB_Name = "frmCancelaColeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

TxtFilialColeta.Text = de_informa.rsColetaSel.Fields("filialcoleta")
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

TxtMotivoCanc.Text = ""
TxtUsuario.Text = ""
TxtDataCanc.Text = ""
TxtHoraCanc.Text = ""

    If de_informa.rsColetaSel.Fields("tem_ocorr") = "N" Then
    lblStatus.Caption = "Sem Posição"
    xcancela = "S"
    ElseIf de_informa.rsColetaSel.Fields("tem_ocorr") = "1" Then
    lblStatus.Caption = "Entregue/Concluída"
    xcancela = "S"
    ElseIf de_informa.rsColetaSel.Fields("tem_ocorr") = "2" Then
    lblStatus.Caption = "Em Ocorrência"
    xcancela = "S"
    ElseIf de_informa.rsColetaSel.Fields("tem_ocorr") = "0" Then
    lblStatus.Caption = "Coleta Baixada"
    xcancela = "S"
    ElseIf de_informa.rsColetaSel.Fields("tem_ocorr") = "C" Then
    lblStatus.Caption = "Coleta Cancelada"
    xcancela = "N"
    TxtMotivoCanc.Text = de_informa.rsColetaSel.Fields("canc_motivo")
    TxtUsuario.Text = de_informa.rsColetaSel.Fields("canc_usu")
    TxtDataCanc.Text = de_informa.rsColetaSel.Fields("canc_data")
    TxtHoraCanc.Text = de_informa.rsColetaSel.Fields("canc_hora")
    End If

TxtFilial.SetFocus
    
    If xcancela = "S" Then
    FraCancelamento.Enabled = True
    TxtMotivoCanc.Enabled = True
    CmdCancelaColeta.Enabled = True
    TxtMotivoCanc.BackColor = xamarelo1
    TxtMotivoCanc.SetFocus
    End If
End Sub


Private Sub CmdCancelaColeta_Click()
    If Len(Trim(TxtMotivoCanc.Text)) = 0 Then
    Exit Sub
    End If
    
    de_informa.cn_informa.BeginTrans
    de_informa.ColetaCanc "C", datahora("DATA"), datahora("HORA"), Trim(UCase(TxtMotivoCanc.Text)), xusuario, TxtFilialColeta.Text
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "CANCELAMENTO", xusuario, "COLETA: " & TxtFilialColeta.Text

    de_informa.cn_informa.CommitTrans
    
    FraCancelamento.Enabled = False
    TxtMotivoCanc.Enabled = False
    CmdCancelaColeta.Enabled = False
    TxtMotivoCanc.BackColor = xbranco

    CmdBuscar_Click
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

Private Sub TxtMotivoCanc_Change()
X = TxtMotivoCanc.SelStart
TxtMotivoCanc.Text = UCase(TxtMotivoCanc.Text)
TxtMotivoCanc.SelStart = X
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
