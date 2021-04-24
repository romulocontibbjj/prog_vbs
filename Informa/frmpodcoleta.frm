VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPodColeta 
   Caption         =   "POD - Coletas"
   ClientHeight    =   6855
   ClientLeft      =   600
   ClientTop       =   885
   ClientWidth     =   10770
   ControlBox      =   0   'False
   Icon            =   "frmpodcoleta.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6855
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraOcorr 
      Caption         =   "Códigos de Ocorrências"
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
      Left            =   480
      TabIndex        =   59
      Top             =   780
      Visible         =   0   'False
      Width           =   4935
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexCodOcorr 
         Height          =   5115
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   9022
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame FraSombra 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   600
      TabIndex        =   61
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Frame FraPOD 
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
      Height          =   5535
      Left            =   5820
      TabIndex        =   12
      Top             =   1200
      Width           =   4815
      Begin MSMask.MaskEdBox TxtData 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSDataGridLib.DataGrid DataGridOcorr 
         Bindings        =   "frmpodcoleta.frx":000C
         Height          =   2415
         Left            =   120
         TabIndex        =   56
         Top             =   1140
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4260
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
      Begin VB.TextBox TxtDescricao 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton CmdGravar 
         Caption         =   "Gravar Ocorrência de Coleta"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   5100
         Width           =   4515
      End
      Begin VB.TextBox TxtObsOcorr 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   3900
         Width           =   4515
      End
      Begin VB.TextBox TxtHora 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtCodOcorr 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Observação / Cometário:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   3660
         Width           =   1785
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Left            =   2640
         TabIndex        =   21
         Top             =   765
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   765
         Width           =   390
      End
   End
   Begin VB.Frame FraDados 
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
      TabIndex        =   11
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox TxtValMerc 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   50
         Top             =   3060
         Width           =   975
      End
      Begin VB.TextBox TxtPeso 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2700
         TabIndex        =   49
         Top             =   3060
         Width           =   975
      End
      Begin VB.TextBox TxtVolumes 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   48
         Top             =   3060
         Width           =   975
      End
      Begin VB.TextBox TxtNFs 
         Enabled         =   0   'False
         Height          =   615
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   3360
         Width           =   5055
      End
      Begin VB.TextBox TxtOBS 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   4320
         Width           =   5415
      End
      Begin VB.TextBox TxtEspecie 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4020
         TabIndex        =   45
         Top             =   2760
         Width           =   1515
      End
      Begin VB.TextBox TxtNatureza 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   44
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox TxtTipoFrete 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   42
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox TxtContato 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   40
         Top             =   2460
         Width           =   1815
      End
      Begin VB.TextBox TxtHoraColeta 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   38
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox TxtDataColeta 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   36
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox TxtSolicitante 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Top             =   1860
         Width           =   3615
      End
      Begin VB.TextBox TxtCidadeDes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   32
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox TxtCidadeRem 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Top             =   1260
         Width           =   3615
      End
      Begin VB.TextBox TxtEnderecoRem 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox TxtRemetente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   26
         Top             =   660
         Width           =   3615
      End
      Begin VB.TextBox TxtCGCRem 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "OBS:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Merc:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   3105
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Peso:"
         Height          =   195
         Left            =   2220
         TabIndex        =   53
         Top             =   3105
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Volumes:"
         Height          =   195
         Left            =   3840
         TabIndex        =   52
         Top             =   3105
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "NFs:"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   3360
         Width           =   330
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Natureza/Espécie:"
         Height          =   195
         Left            =   450
         TabIndex        =   43
         Top             =   2805
         Width           =   1335
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Frete:"
         Height          =   195
         Left            =   3840
         TabIndex        =   41
         Top             =   2505
         Width           =   765
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
         Height          =   195
         Left            =   1185
         TabIndex        =   39
         Top             =   2505
         Width           =   600
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Horário:"
         Height          =   195
         Left            =   3360
         TabIndex        =   37
         Top             =   2205
         Width           =   555
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Data de Coleta:"
         Height          =   195
         Left            =   675
         TabIndex        =   34
         Top             =   2205
         Width           =   1110
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Solicitante:"
         Height          =   195
         Left            =   1005
         TabIndex        =   33
         Top             =   1905
         Width           =   780
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Destino/UF:"
         Height          =   195
         Left            =   915
         TabIndex        =   31
         Top             =   1605
         Width           =   870
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Cidade/UF:"
         Height          =   195
         Left            =   960
         TabIndex        =   29
         Top             =   1305
         Width           =   825
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   1050
         TabIndex        =   27
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Rementente:"
         Height          =   195
         Left            =   870
         TabIndex        =   25
         Top             =   705
         Width           =   915
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CGC Rementente:"
         Height          =   195
         Left            =   495
         TabIndex        =   23
         Top             =   405
         Width           =   1290
      End
   End
   Begin VB.Frame FraMain 
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
      TabIndex        =   10
      Top             =   60
      Width           =   10515
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   4500
         TabIndex        =   57
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
            TabIndex        =   58
            Top             =   360
            Width           =   105
         End
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "Sair"
         Height          =   315
         Left            =   8160
         TabIndex        =   9
         Top             =   660
         Width           =   2235
      End
      Begin VB.TextBox TxtPrioridade 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8940
         TabIndex        =   19
         Top             =   255
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         Left            =   8160
         TabIndex        =   18
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   135
         Left            =   3960
         TabIndex        =   17
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emissor/Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   645
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial/Coleta:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmPodColeta"
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

txtCodOcorr.Text = ""
txtDescricao.Text = ""
TxtObsOcorr.Text = ""
TxtData.Mask = ""
TxtData.Text = ""
TxtHora.Text = ""
txtCodOcorr.Enabled = True
txtDescricao.Enabled = True
TxtData.Enabled = False
TxtHora.Enabled = False
TxtObsOcorr.Enabled = False
cmdGravar.Enabled = False
txtCodOcorr.BackColor = xamarelo1
txtDescricao.BackColor = xamarelo1
TxtData.BackColor = xbranco
TxtHora.BackColor = xbranco
TxtObsOcorr.BackColor = xbranco
txtCodOcorr.SetFocus
End Sub

Private Sub cmdGravar_Click()

    If TxtData.Text = "" Then
    MsgBox "Data incorreta.", vbCritical
    TxtData.SetFocus
    Exit Sub
    End If
    
    If UCase(Trim(TxtObsOcorr.Text)) = "DATA AGENDADA PRIMORDIALMENTE" Then
    MsgBox "Texto de Observação inválido. Por favor, corrija sua observação.", vbCritical
    TxtObsOcorr.Text = ""
    TxtObsOcorr.SetFocus
    Exit Sub
    End If
    
    If Len(Trim(TxtFilial.Text)) = 0 Then
    MsgBox "Filial incorreta.", vbCritical
    Exit Sub
    ElseIf Len(Trim(TxtColeta.Text)) = 0 Then
    MsgBox "Coleta incorreta.", vbCritical
    Exit Sub
    End If
    
    If txtCodOcorr.Text = "01" And lblStatus.Caption = "Entregue/Concluída" Then
    MsgBox "Esta Coleta já está concluída. Não é possível incluir uma nova data de entrega.", vbCritical
    Exit Sub
    End If
    
    If txtCodOcorr.Text = "00" And lblStatus.Caption = "Coleta Baixada" Then
    MsgBox "Esta Coleta já está Baixada. Não é possível incluir uma nova ocorrência de Baixa.", vbCritical
    Exit Sub
    End If
    
    If txtCodOcorr.Text = "00" And Len(Trim(TxtObsOcorr.Text)) = 0 Then
    MsgBox "É necessário que você insira uma observação justificando o motivo desta Coleta estar sendo baixada.", vbCritical
    Exit Sub
    End If
    
    If txtCodOcorr.Text = "99" And Len(Trim(TxtObsOcorr.Text)) = 0 Then
    MsgBox "É necessário que você insira uma observação descrevendo este outro tipo de ococrrência.", vbCritical
    Exit Sub
    End If
    
    If txtCodOcorr.Text = "01" And lblStatus.Caption = "Coleta Baixada" Then
    MsgBox "Esta Coleta já está Baixada. Não é possível incluir uma nova data de Entrega.", vbCritical
    Exit Sub
    End If
    
    If lblStatus.Caption = "Coleta Cancelada" Then
    MsgBox "Esta Coleta está Cancelada. Não é possível incluir ocorrências para esta Coleta.", vbCritical
    Exit Sub
    End If
    
de_informa.cn_informa.BeginTrans
de_informa.ColetaInsOcorr transctc(TxtFilial.Text, TxtColeta.Text), CDate(TxtEmissao.Text), txtCodOcorr.Text, txtDescricao.Text, UCase(Trim(TxtObsOcorr.Text)), CDate(TxtData.Text), TxtHora.Text, xusuario, datahora("DATAHORA")
    If txtCodOcorr.Text = "01" Then
    de_informa.ColetaUpdateTemOcorr transctc(TxtFilial.Text, TxtColeta.Text), "1"
    ElseIf txtCodOcorr.Text = "00" Then
    de_informa.ColetaUpdateTemOcorr transctc(TxtFilial.Text, TxtColeta.Text), "0"
    Else
        If lblStatus.Caption <> "Entregue/Concluída" And lblStatus.Caption <> "Coleta Baixada" Then
        de_informa.ColetaUpdateTemOcorr transctc(TxtFilial.Text, TxtColeta.Text), "2"
        End If
    End If
'LOG DE USUÁRIO
de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/COLETA:" & transctc(TxtFilial.Text, TxtColeta.Text) & " OCORR:" & txtCodOcorr & "-" & txtDescricao

de_informa.cn_informa.CommitTrans

If de_informa.rsColetaSelOcorr.State = 1 Then de_informa.rsColetaSelOcorr.Close
de_informa.ColetaSelOcorr transctc(TxtFilial.Text, TxtColeta.Text)

Set DataGridOcorr.DataSource = de_informa
DataGridOcorr.DataMember = "ColetaSelOcorr"
DataGridOcorr.Refresh
DoEvents
CmdBuscar_Click
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGridOcorr_Click()
TxtObsOcorr.Text = DataGridOcorr.Columns(7)
End Sub

Private Sub DataGridOcorr_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DataGridOcorr_Click
End Sub

Private Sub FlexCodOcorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtCodOcorr.Text = FlexCodOcorr.TextMatrix(FlexCodOcorr.Row, 0)
    txtDescricao.Text = UCase(FlexCodOcorr.TextMatrix(FlexCodOcorr.Row, 1))
    FraSombra.Visible = False
    FraOcorr.Visible = False
    fraDados.Enabled = True
    FraMain.Enabled = True
    FraPOD.Enabled = True
    DoEvents
    txtDescricao.SetFocus
    ElseIf KeyAscii = 27 Then
    txtCodOcorr.Text = ""
    txtDescricao.Text = ""
    FraSombra.Visible = False
    FraOcorr.Visible = False
    fraDados.Enabled = True
    FraMain.Enabled = True
    FraPOD.Enabled = True
    DoEvents
    txtDescricao.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Len(Trim(TxtFilial.Text)) > 0 And Len(Trim(TxtColeta.Text)) > 0 Then
    CmdBuscar_Click
    End If
End Sub

Private Sub TxtCodOcorr_GotFocus()
Dim xTextBox As TextBox
Set xTextBox = txtCodOcorr
xTextBox.SelStart = 0
xTextBox.SelLength = 500
End Sub

Private Sub TxtCodOcorr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub txtCodOcorr_LostFocus()
If Len(txtCodOcorr.Text) > 0 Then
    If Trim(txtCodOcorr.Text) = "?" Then
    If de_informa.rsColetaSelCodOcorr.State = 1 Then de_informa.rsColetaSelCodOcorr.Close
    de_informa.ColetaSelCodOcorr "%", "%"
    FlexCodOcorr.Clear
    FlexCodOcorr.Rows = de_informa.rsColetaSelCodOcorr.RecordCount + 1
    FlexCodOcorr.Cols = 2
    FlexCodOcorr.FixedRows = 1
    FlexCodOcorr.FixedCols = 0
    FlexCodOcorr.TextMatrix(0, 0) = "Código"
    FlexCodOcorr.TextMatrix(0, 1) = "Descrição"
    FlexCodOcorr.ColWidth(0) = 800
    FlexCodOcorr.ColWidth(1) = 3500
    linha = 0
        Do Until de_informa.rsColetaSelCodOcorr.EOF
        linha = linha + 1
        FlexCodOcorr.TextMatrix(linha, 0) = de_informa.rsColetaSelCodOcorr.Fields("cod_ocorr")
        FlexCodOcorr.TextMatrix(linha, 1) = PriMaiuscula(de_informa.rsColetaSelCodOcorr.Fields("descricao"))
        de_informa.rsColetaSelCodOcorr.MoveNext
        Loop
    FraSombra.Visible = True
    FraOcorr.Visible = True
    FlexCodOcorr.SetFocus
    
    fraDados.Enabled = False
    FraMain.Enabled = False
    FraPOD.Enabled = False
    DoEvents
    Else
    If de_informa.rsColetaSelCodOcorr.State = 1 Then de_informa.rsColetaSelCodOcorr.Close
    de_informa.ColetaSelCodOcorr txtCodOcorr.Text & "%", "%"
        If de_informa.rsColetaSelCodOcorr.RecordCount > 0 Then
        txtCodOcorr.Text = de_informa.rsColetaSelCodOcorr.Fields("cod_ocorr")
        txtDescricao.Text = de_informa.rsColetaSelCodOcorr.Fields("descricao")
        TxtData.Enabled = True
        TxtHora.Enabled = True
        TxtObsOcorr.Enabled = True
        cmdGravar.Enabled = True
        TxtData.BackColor = xamarelo1
        TxtHora.BackColor = xamarelo1
        TxtObsOcorr.BackColor = xamarelo1
        Else
        txtCodOcorr.Text = ""
        txtDescricao.Text = ""
        TxtObsOcorr.Text = ""
        TxtData.Enabled = False
        TxtHora.Enabled = False
        TxtObsOcorr.Enabled = False
        cmdGravar.Enabled = False
        TxtData.BackColor = xbranco
        TxtHora.BackColor = xbranco
        TxtObsOcorr.BackColor = xbranco
        End If
    End If
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

Private Sub TxtData_GotFocus()
Call Date_MskEdBox_GotFocus(TxtData)
End Sub

Private Sub TxtData_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtData_LostFocus()
Call Date_MskEdBox_LostFocus(TxtData)
End Sub

Private Sub TxtDescricao_GotFocus()
Dim xTextBox As TextBox
Set xTextBox = txtDescricao
xTextBox.SelStart = 0
xTextBox.SelLength = 500
End Sub

Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtDescricao_LostFocus()
If Len(txtDescricao.Text) > 0 Then
    If Trim(txtDescricao.Text) = "?" Then
    If de_informa.rsColetaSelCodOcorr.State = 1 Then de_informa.rsColetaSelCodOcorr.Close
    de_informa.ColetaSelCodOcorr "%", "%"
    FlexCodOcorr.Clear
    FlexCodOcorr.Rows = de_informa.rsColetaSelCodOcorr.RecordCount + 1
    FlexCodOcorr.Cols = 2
    FlexCodOcorr.FixedRows = 1
    FlexCodOcorr.FixedCols = 0
    FlexCodOcorr.TextMatrix(0, 0) = "Código"
    FlexCodOcorr.TextMatrix(0, 1) = "Descrição"
    FlexCodOcorr.ColWidth(0) = 800
    FlexCodOcorr.ColWidth(1) = 3500
    linha = 0
        Do Until de_informa.rsColetaSelCodOcorr.EOF
        linha = linha + 1
        FlexCodOcorr.TextMatrix(linha, 0) = de_informa.rsColetaSelCodOcorr.Fields("cod_ocorr")
        FlexCodOcorr.TextMatrix(linha, 1) = PriMaiuscula(de_informa.rsColetaSelCodOcorr.Fields("descricao"))
        de_informa.rsColetaSelCodOcorr.MoveNext
        Loop
    FraSombra.Visible = True
    FraOcorr.Visible = True
    FlexCodOcorr.SetFocus
    fraDados.Enabled = False
    FraMain.Enabled = False
    FraPOD.Enabled = False
    DoEvents
    Else
    If de_informa.rsColetaSelCodOcorr.State = 1 Then de_informa.rsColetaSelCodOcorr.Close
    de_informa.ColetaSelCodOcorr "%", txtDescricao.Text & "%"
        If de_informa.rsColetaSelCodOcorr.RecordCount > 0 Then
        txtCodOcorr.Text = de_informa.rsColetaSelCodOcorr.Fields("cod_ocorr")
        txtDescricao.Text = de_informa.rsColetaSelCodOcorr.Fields("descricao")
        TxtData.Enabled = True
        TxtHora.Enabled = True
        TxtObsOcorr.Enabled = True
        cmdGravar.Enabled = True
        TxtData.BackColor = xamarelo1
        TxtHora.BackColor = xamarelo1
        TxtObsOcorr.BackColor = xamarelo1
        Else
        txtCodOcorr.Text = ""
        txtDescricao.Text = ""
        TxtObsOcorr.Text = ""
        TxtData.Enabled = False
        TxtHora.Enabled = False
        TxtObsOcorr.Enabled = False
        cmdGravar.Enabled = False
        TxtData.BackColor = xbranco
        TxtHora.BackColor = xbranco
        TxtObsOcorr.BackColor = xbranco
        End If
    End If
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

Private Sub txtHora_GotFocus()
    If Len(Trim(TxtHora.Text)) = 0 Then
    TxtHora.Text = "  :  "
    End If
End Sub

Private Sub txtHora_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    ElseIf KeyAscii = 8 Then
    TxtHora.Text = "  :  "
    Else
        If TxtHora.Text = "  :  " Then
            If Chr(KeyAscii) > 2 Then
            KeyAscii = 0
            Else
            TxtHora.Text = Chr(KeyAscii) & " :  "
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHora.Text, 2) = " :  " Then
            If Mid(TxtHora.Text, 1, 1) = "2" Then
                If Chr(KeyAscii) > 3 Then
                KeyAscii = 0
                Else
                TxtHora.Text = Mid(TxtHora.Text, 1, 1) & Chr(KeyAscii) & ":  "
                KeyAscii = 0
                End If
            Else
            TxtHora.Text = Mid(TxtHora.Text, 1, 1) & Chr(KeyAscii) & ":  "
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHora.Text, 3) = ":  " Then
            If Chr(KeyAscii) > 5 Then
            KeyAscii = 0
            Else
            TxtHora.Text = Mid(TxtHora.Text, 1, 3) & Chr(KeyAscii) & " "
            KeyAscii = 0
            End If
        ElseIf Mid(TxtHora.Text, 5) = " " Then
            TxtHora.Text = Mid(TxtHora.Text, 1, 4) & Chr(KeyAscii)
            KeyAscii = 0
        End If
    KeyAscii = 0
    End If
End Sub

Private Sub txtHora_LostFocus()
    If InStr(1, TxtHora.Text, " ", vbTextCompare) > 0 Then
    TxtHora.Text = ""
    End If
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

