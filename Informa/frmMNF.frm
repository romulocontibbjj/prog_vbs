VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMNF 
   Caption         =   "Consulta de Manifestos"
   ClientHeight    =   7965
   ClientLeft      =   1035
   ClientTop       =   960
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin VB.Frame fraAguarde 
      Height          =   855
      Left            =   4440
      TabIndex        =   77
      Top             =   2880
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "A G U A R D E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   300
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   700
      Left            =   11040
      TabIndex        =   8
      Top             =   230
      Width           =   855
   End
   Begin VB.Frame frame4 
      Caption         =   "Veículo"
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
      TabIndex        =   40
      Top             =   960
      Width           =   11775
      Begin VB.Label lblProprietario 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4560
         TabIndex        =   66
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label lblPlaca 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2400
         TabIndex        =   65
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lblCodVeiculo 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   64
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         Height          =   195
         Left            =   1800
         TabIndex        =   59
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Proprietário:"
         Height          =   195
         Left            =   3600
         TabIndex        =   58
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Frota:"
         Height          =   195
         Left            =   3720
         TabIndex        =   57
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblFrotaVeic 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   56
         Top             =   700
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Marca/Mod:"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblMarcaVeic 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   54
         Top             =   720
         Width           =   2580
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   4920
         TabIndex        =   53
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblTipoVeic 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5355
         TabIndex        =   52
         Top             =   700
         Width           =   1260
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Suspensão a Ar:"
         Height          =   195
         Left            =   9960
         TabIndex        =   51
         Top             =   290
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Plataforma Hidráulica:"
         Height          =   195
         Left            =   9600
         TabIndex        =   50
         Top             =   530
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Câmara Fria:"
         Height          =   195
         Left            =   10200
         TabIndex        =   49
         Top             =   770
         Width           =   885
      End
      Begin VB.Label lblSuspVeic 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   11280
         TabIndex        =   48
         Top             =   290
         Width           =   285
      End
      Begin VB.Label lblPlataformaVeic 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   11280
         TabIndex        =   47
         Top             =   530
         Width           =   285
      End
      Begin VB.Label lblCamaraVeic 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   11280
         TabIndex        =   46
         Top             =   770
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Capacidade:"
         Height          =   195
         Left            =   6720
         TabIndex        =   45
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblCapacPesoVeic 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7800
         TabIndex        =   44
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblCapacM3Veic 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8600
         TabIndex        =   43
         Top             =   700
         Width           =   825
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Rastreamento:"
         Height          =   195
         Left            =   6720
         TabIndex        =   42
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblRastreamVeic 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7800
         TabIndex        =   41
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.Frame frame5 
      Caption         =   "Motorista / Ajudantes"
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
      TabIndex        =   29
      Top             =   2160
      Width           =   11775
      Begin VB.Label lblCpfMotorista 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7920
         TabIndex        =   76
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CPF:"
         Height          =   195
         Left            =   7440
         TabIndex        =   75
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblAjudantes 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   68
         Top             =   720
         Width           =   6300
      End
      Begin VB.Label lblMotorista 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   67
         Top             =   360
         Width           =   3300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Motorista:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Função:"
         Height          =   195
         Left            =   4560
         TabIndex        =   38
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblFuncaoMotorista 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5280
         TabIndex        =   37
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "CNH:"
         Height          =   195
         Left            =   7440
         TabIndex        =   36
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblCNH 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7920
         TabIndex        =   35
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Vencto:"
         Height          =   195
         Left            =   9480
         TabIndex        =   34
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblcnhVencto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10320
         TabIndex        =   33
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Admissão:"
         Height          =   195
         Left            =   9480
         TabIndex        =   32
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lblAdmissaoMotorista 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10320
         TabIndex        =   31
         Top             =   700
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Ajudantes:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   700
         Width           =   750
      End
   End
   Begin VB.Frame frame6 
      Caption         =   "Data / Conferente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   11775
      Begin VB.Label lblHoraEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   72
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblObs 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6600
         TabIndex        =   71
         Top             =   360
         Width           =   4980
      End
      Begin VB.Label lblConferente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4320
         TabIndex        =   70
         Top             =   360
         Width           =   1740
      End
      Begin VB.Label lblDataEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   69
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conferente Oper:"
         Height          =   195
         Left            =   3000
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "OBS:"
         Height          =   195
         Left            =   6120
         TabIndex        =   26
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame frame3 
      Caption         =   "Manifesto"
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
      Left            =   5280
      TabIndex        =   22
      Top             =   120
      Width           =   5655
      Begin VB.Label lblLacre 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4200
         TabIndex        =   63
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label lblFilialDest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   62
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblDistribTransf 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lacre de Seg:"
         Height          =   195
         Left            =   3120
         TabIndex        =   24
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Filial Destino:"
         Height          =   195
         Left            =   1680
         TabIndex        =   23
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.Frame frame7 
      Caption         =   "CTRs / CTCs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   11775
      Begin VB.CommandButton cmdGeraArquivo 
         Caption         =   "Gera Arquivo..."
         Enabled         =   0   'False
         Height          =   300
         Left            =   10080
         TabIndex        =   7
         Top             =   330
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid gridDadosMnf 
         Bindings        =   "frmMNF.frx":0000
         Height          =   2775
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4895
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
         DataMember      =   "Sel_ManifestoPorNum"
         ColumnCount     =   24
         BeginProperty Column00 
            DataField       =   "filialmanifesto"
            Caption         =   "Filial-Manifesto"
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
            DataField       =   "motivo"
            Caption         =   "motivo"
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
            DataField       =   "filialdest"
            Caption         =   "filialdest"
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
            DataField       =   "lacre"
            Caption         =   "lacre"
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
            DataField       =   "ajudantes"
            Caption         =   "ajudantes"
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
            DataField       =   "dtemissao"
            Caption         =   "Emissão"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "hsemissao"
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
            DataField       =   "codveiculo"
            Caption         =   "Cod.Veic."
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
            DataField       =   "placaveic"
            Caption         =   "Placa"
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
            DataField       =   "proprietario"
            Caption         =   "Proprietário Veic."
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
            DataField       =   "motorista"
            Caption         =   "Motorista"
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
            DataField       =   "conferente"
            Caption         =   "conferente"
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
            DataField       =   "obs"
            Caption         =   "obs"
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
            DataField       =   "filialctc"
            Caption         =   "Filial-CTC"
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
            DataField       =   "data"
            Caption         =   "Emissão CTC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "remet_nome"
            Caption         =   "Remetente"
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
            DataField       =   "dest_nome"
            Caption         =   "Destinatário"
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
         BeginProperty Column17 
            DataField       =   "cidade_dest"
            Caption         =   "Cidade Destino"
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
         BeginProperty Column18 
            DataField       =   "uf_dest"
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
         BeginProperty Column19 
            DataField       =   "nfs"
            Caption         =   "Notas Fiscais"
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
         BeginProperty Column20 
            DataField       =   "valmerc"
            Caption         =   "Valor de Mercadoria"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column21 
            DataField       =   "peso"
            Caption         =   "Peso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0,0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column22 
            DataField       =   "volumes"
            Caption         =   "Volumes"
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
         BeginProperty Column23 
            DataField       =   "modal"
            Caption         =   "Modal"
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
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
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
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   2369,764
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   2610,142
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   2039,811
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   374,74
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1964,976
            EndProperty
            BeginProperty Column20 
               Alignment       =   1
               ColumnWidth     =   1530,142
            EndProperty
            BeginProperty Column21 
               Alignment       =   1
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column23 
               ColumnWidth     =   1379,906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Mnfs:"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblTotMnf 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   73
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Ctcs/Ctrs:"
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Peso (Kg):"
         Height          =   195
         Left            =   2880
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Volumes:"
         Height          =   195
         Left            =   5040
         TabIndex        =   19
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Valor Mercadoria R$:"
         Height          =   195
         Left            =   6720
         TabIndex        =   18
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblTotCtcs 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblTotPeso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblTotVol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   15
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lblTotValMerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8280
         TabIndex        =   14
         Top             =   360
         Width           =   1620
      End
   End
   Begin VB.Frame frame2 
      Caption         =   "Filial+Número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1920
      TabIndex        =   11
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmbProcurar 
         Caption         =   ">>"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtManifesto 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   4
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   360
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   330
      End
      Begin VB.TextBox txtPlaca 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton cmbSair 
      Caption         =   "Sair"
      Height          =   555
      Left            =   6360
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame frame1 
      Caption         =   "Procura Por ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1680
      Begin VB.OptionButton optPlacaData 
         Caption         =   "Placa + Data"
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optMnf 
         Caption         =   "Núm. Manifesto"
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProcurar_Click()
    Dim xrs As Recordset
    Dim xTotMnf As Integer, xTotCtc As Integer, xTotPeso As Currency
    Dim xTotVol As Long, xTotValmerc As Currency, xMnfAnterior As String
    
    If optMnf.Value = True Then
        Frame1.Enabled = False
        Frame2.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        frame5.Enabled = False
        frame6.Enabled = False
        frame7.Enabled = False
        CmdSair.Enabled = False
        fraAguarde.Visible = True
        If de_informa.rsSel_ManifestoPorNum.State = 1 Then de_informa.rsSel_ManifestoPorNum.Close
        de_informa.Sel_ManifestoPorNum transmanif(TxtFilial, txtManifesto)
        gridDadosMnf.DataMember = "sel_manifestopornum"
        gridDadosMnf.Refresh
        Set xrs = de_informa.rsSel_ManifestoPorNum
    ElseIf optPlacaData.Value = True Then
        If Not IsDate(mskData) Then
            MsgBox "Data Inválida !", vbCritical
            mskData.SetFocus
            Exit Sub
        End If
        Frame1.Enabled = False
        Frame2.Enabled = False
        Frame3.Enabled = False
        Frame4.Enabled = False
        frame5.Enabled = False
        frame6.Enabled = False
        frame7.Enabled = False
        CmdSair.Enabled = False
        fraAguarde.Visible = True
        If de_informa.rsSel_ManifestoPorDtPlaca.State = 1 Then de_informa.rsSel_ManifestoPorDtPlaca.Close
        de_informa.Sel_ManifestoPorDtPlaca CDate(mskData), Trim$(txtPlaca) & "%"
        gridDadosMnf.DataMember = "Sel_ManifestoPorDtPlaca"
        gridDadosMnf.Refresh
        Set xrs = de_informa.rsSel_ManifestoPorDtPlaca
    End If
    
    If xrs.RecordCount > 0 Then

        lblDistribTransf = xrs.Fields("motivo")
        lblFilialDest = xrs.Fields("filialdest")
        lblLacre = xrs.Fields("lacre")
        lblCodVeiculo = xrs.Fields("codveiculo")
        lblPlaca = xrs.Fields("placaveic")
        lblProprietario = xrs.Fields("proprietario")
        lblMotorista = xrs.Fields("motorista")
        lblAjudantes = xrs.Fields("ajudantes")
        lblDataEmissao = xrs.Fields("dtemissao")
        lblHoraEmissao = xrs.Fields("hsemissao")
        lblConferente = xrs.Fields("conferente")
        lblObs = xrs.Fields("obs")
        
        'dados do veículo
        If de_informa.rsSel_Veiculo.State = 1 Then de_informa.rsSel_Veiculo.Close
        de_informa.Sel_Veiculo xrs.Fields("placaveic")
        
        If de_informa.rsSel_Veiculo.RecordCount > 0 Then
            lblRastreamVeic = de_informa.rsSel_Veiculo.Fields("rastreamento")
            lblSuspVeic = de_informa.rsSel_Veiculo.Fields("suspensaoar")
            lblMarcaVeic = de_informa.rsSel_Veiculo.Fields("marca")
            lblFrotaVeic = de_informa.rsSel_Veiculo.Fields("frota")
            lblTipoVeic = de_informa.rsSel_Veiculo.Fields("tipo")
            lblCapacM3Veic = de_informa.rsSel_Veiculo.Fields("capacidadem3")
            lblCapacPesoVeic = de_informa.rsSel_Veiculo.Fields("capacidadepeso")
            lblPlataformaVeic = de_informa.rsSel_Veiculo.Fields("plataformahidr")
            lblCamaraVeic = de_informa.rsSel_Veiculo.Fields("camarafria")
        Else
            lblRastreamVeic = ""
            lblSuspVeic = ""
            lblMarcaVeic = ""
            lblFrotaVeic = ""
            lblTipoVeic = ""
            lblCapacM3Veic = ""
            lblCapacPesoVeic = ""
            lblPlataformaVeic = ""
            lblCamaraVeic = ""
        End If
        
        'dados do motorista
        
        If de_informa.rsSel_Motorista.State = 1 Then de_informa.rsSel_Motorista.Close
        de_informa.Sel_Motorista Trim$(xrs.Fields("motorista"))
        
        If de_informa.rsSel_Motorista.RecordCount > 0 Then
            lblFuncaoMotorista = de_informa.rsSel_Motorista.Fields("funcao")
            lblCNH = de_informa.rsSel_Motorista.Fields("cnh")
            lblCpfMotorista = de_informa.rsSel_Motorista.Fields("cpf")
            lblcnhVencto = de_informa.rsSel_Motorista.Fields("cnhvencto")
            lblAdmissaoMotorista = de_informa.rsSel_Motorista.Fields("admissao")
        Else
            lblFuncaoMotorista = ""
            lblCNH = ""
            lblCpfMotorista = ""
            lblcnhVencto = ""
            lblAdmissaoMotorista = ""
        End If
        
        'dados totalizados
        
        xrs.MoveFirst
        xTotCtc = 0
        xTotMnf = 0
        xTotPeso = 0
        xTotValmerc = 0
        xTotVol = 0
        xMnfAnterior = ""
        
        xTotCtc = xrs.RecordCount
        
        Do Until xrs.EOF
            If xMnfAnterior <> xrs.Fields("filialmanifesto") Then xTotMnf = xTotMnf + 1
            xMnfAnterior = xrs.Fields("filialmanifesto")
            xTotPeso = xTotPeso + xrs.Fields("peso")
            xTotVol = xTotVol + xrs.Fields("volumes")
            xTotValmerc = xTotValmerc + xrs.Fields("valmerc")
            xrs.MoveNext
        Loop
        
        xrs.MoveFirst
        
        lblTotMnf = xTotMnf
        lblTotCtcs = xTotCtc
        lblTotPeso = Format(xTotPeso, "###,##0.0")
        lblTotVol = xTotVol
        lblTotValMerc = Format(xTotValmerc, "##,###,##0.00")
        cmdGeraArquivo.Enabled = True
    Else
        MsgBox "Não Foi Encontrado Manifesto(s) Com Estes Dados !", vbCritical, "OPS"
        cmdGeraArquivo.Enabled = False
    End If
    
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = True
    frame5.Enabled = True
    frame6.Enabled = True
    frame7.Enabled = True
    CmdSair.Enabled = True
    fraAguarde.Visible = False

    On Error Resume Next
    txtPlaca.SetFocus
    TxtFilial.SetFocus


End Sub
Private Sub cmdGeraArquivo_Click()
    Dim xrs As Recordset, xfile As String

    If optMnf.Value = True Then
        Set xrs = de_informa.rsSel_ManifestoPorNum
        xfile = "C:\MNF_" & transmanif(TxtFilial, txtManifesto) & ".TXT"
    ElseIf optPlacaData.Value = True Then
        Set xrs = de_informa.rsSel_ManifestoPorDtPlaca
        xfile = "C:\MNF_" & Trim$(txtPlaca) & "_" & Mid$(mskData, 1, 2) & Mid$(mskData, 4, 2) & ".TXT"
    End If
    
    If xrs.RecordCount > 0 Then
        
        xrs.MoveFirst
        
        
        Open xfile For Output As #1
        
        Print #1, "Filial-Manifesto#Emissao#Hora#Cod.Veiculo#Placa#Proprietario#Motorista#Filial-CTC#Data CTC#Remetente#Destinatario#Cidade#UF#Notas Fiscais#Valor Merc#Peso#Volumes#Modal"
        
        Do Until xrs.EOF
        
            xlinha = xrs.Fields("filialmanifesto") & "#" & xrs.Fields("dtemissao") & "#" & xrs.Fields("hsemissao") & "#" & xrs.Fields("codveiculo") & "#" & _
                     xrs.Fields("placaveic") & "#" & xrs.Fields("proprietario") & "#" & xrs.Fields("motorista") & "#" & xrs.Fields("filialctc") & "#" & _
                     xrs.Fields("data") & "#" & xrs.Fields("remet_nome") & "#" & xrs.Fields("dest_nome") & "#" & xrs.Fields("cidade_dest") & "#" & _
                     xrs.Fields("uf_dest") & "#" & xrs.Fields("nfs") & "#" & (SoNumeros(Format(xrs.Fields("valmerc"), "##,###,##0.00")) / 100) & "#" & _
                     (SoNumeros(Format(xrs.Fields("peso"), "##,##0.0")) / 10) & "#" & xrs.Fields("volumes") & "#" & xrs.Fields("modal") & "#"
            
            Print #1, xlinha
            
            xrs.MoveNext
        
        Loop
        
        xrs.MoveFirst
        
        Close #1
        
        MsgBox "Arquivo TXT Gerado ! Para Importá-lo para o Excel, Abra-o como Arquivo Texto e Escolha como Delimitador o caracter '#' ." + Chr(13) + Chr(10) + Chr(13) + Chr(10) + xfile, vbInformation, "Arquivo Gerado"
        
    Else
        MsgBox "Não Há Dados para Geração de Arquivo !", vbCritical
    End If

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.StatusBar1.Visible = False
    gridDadosMnf.DataMember = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Visible = True
    mdiInforma.StatusBar1.Visible = True
End Sub
Private Sub mskData_LostFocus()
    If mskData.Text <> "__/__/____" Then
        mskData.Text = century(mskData.Text)
        If IsDate(mskData.Text) = False Or Mid(mskData.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskData.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub optMnf_Click()
    If optMnf.Value = True Then
        TxtFilial.Visible = True
        txtManifesto.Visible = True
        txtPlaca.Visible = False
        mskData.Visible = False
        TxtFilial.SetFocus
    ElseIf optPlacaData.Value = True Then
        TxtFilial.Visible = False
        txtManifesto.Visible = False
        txtPlaca.Visible = True
        mskData.Visible = True
        txtPlaca.SetFocus
    End If
End Sub

Private Sub optMnf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optPlacaData_Click()
    If optMnf.Value = True Then
        TxtFilial.Visible = True
        txtManifesto.Visible = True
        txtPlaca.Visible = False
        mskData.Visible = False
        TxtFilial.SetFocus
    ElseIf optPlacaData.Value = True Then
        TxtFilial.Visible = False
        txtManifesto.Visible = False
        txtPlaca.Visible = True
        mskData.Visible = True
        txtPlaca.SetFocus
    End If

End Sub

Private Sub optPlacaData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub TxtFilial_Change()
    On Error Resume Next
    If Len(TxtFilial.Text) >= 2 Then txtManifesto.SetFocus
End Sub
Private Sub TxtFilial_GotFocus()
    TxtFilial.SelStart = 0
    TxtFilial.SelLength = 2
End Sub
Private Sub txtfilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtFilial_LostFocus()
    If TxtFilial.Text <> "" Then
        If Not IsNumeric(TxtFilial.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            TxtFilial.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtManifesto_GotFocus()
    txtManifesto.SelStart = 0
    txtManifesto.SelLength = 6
End Sub
Private Sub txtManifesto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtManifesto_LostFocus()
    If txtManifesto.Text <> "" Then
        If Not IsNumeric(txtManifesto.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            txtManifesto.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtPlaca_GotFocus()
    txtPlaca.SelStart = 0
    txtPlaca.SelLength = 7
End Sub
Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub mskData_GotFocus()
    mskData.SelStart = 0
    mskData.SelLength = 10
End Sub
Private Sub mskData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtPlaca_LostFocus()
    txtPlaca = UCase(txtPlaca)
End Sub
