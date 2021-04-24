VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPodNF 
   Caption         =   "Lançamento Manual de Ocorrências e PODs"
   ClientHeight    =   8205
   ClientLeft      =   -2190
   ClientTop       =   3945
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame xt 
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
      Height          =   735
      Left            =   3960
      TabIndex        =   43
      Top             =   120
      Width           =   3360
      Begin VB.OptionButton optNf 
         Caption         =   "Por Núm de NF"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optCTC 
         Caption         =   "Por Núm. de CTC"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNumNf 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox txtCtc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "STATUS DO CTC"
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
      TabIndex        =   12
      Top             =   120
      Width           =   3855
      Begin VB.Label lblEntregueSN 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3585
      End
   End
   Begin VB.Frame fraNfsOcorr 
      Caption         =   "Total de Notas Fiscais deste CTC: 0 NF(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   35
      Top             =   5400
      Width           =   6615
      Begin VB.ListBox lstNfOcorr 
         Height          =   1815
         ItemData        =   "frmPodNF.frx":0000
         Left            =   240
         List            =   "frmPodNF.frx":0002
         TabIndex        =   42
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox lstNfOcorrNAO 
         Height          =   1815
         ItemData        =   "frmPodNF.frx":0004
         Left            =   5280
         List            =   "frmPodNF.frx":0006
         TabIndex        =   41
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   5280
         X2              =   2040
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   4440
         X2              =   1320
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblOcorrPorNFNao 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 Nfs NÃO tiveram a Ocorrência Abaixo"
         Height          =   195
         Left            =   2160
         TabIndex        =   40
         Top             =   960
         Width           =   2835
      End
      Begin VB.Label lblOcorrSelec 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   525
         Left            =   1920
         TabIndex        =   39
         Top             =   1440
         Width           =   2805
      End
      Begin VB.Label lblOcorrPorNF 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0 Nfs tiveram a Ocorrência Abaixo"
         Height          =   195
         Left            =   1800
         TabIndex        =   38
         Top             =   480
         Width           =   2445
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ocorrência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   6840
      TabIndex        =   27
      Top             =   960
      Width           =   5055
      Begin VB.Frame Frame1 
         Caption         =   "Para Ocorrência 01 - ENTREGA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   4815
         Begin VB.CommandButton Command2 
            Caption         =   "Comentários/ Observações de Entrega ..."
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   4680
            Width           =   4335
         End
         Begin VB.Frame fraPreBaixa 
            Caption         =   "Dados da Pré Baixa (Emails, Telefone, etc)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   120
            TabIndex        =   64
            Top             =   840
            Width           =   4575
            Begin VB.TextBox txtRecPreBx 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   1320
               MaxLength       =   25
               TabIndex        =   66
               Top             =   720
               Width           =   3135
            End
            Begin VB.CommandButton cmdExclPreBx 
               Caption         =   "EXCLUIR esta Pré-Baixa"
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               TabIndex        =   65
               Top             =   1200
               Width           =   3135
            End
            Begin VB.Label lblDtPreBx 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   1320
               TabIndex        =   71
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
               Height          =   195
               Left            =   3120
               TabIndex        =   70
               Top             =   360
               Width           =   390
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Data Entrega:"
               Height          =   195
               Left            =   120
               TabIndex        =   69
               Top             =   360
               Width           =   990
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Recebedor...:"
               Height          =   195
               Left            =   120
               TabIndex        =   68
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblHsPreBx 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3600
               TabIndex        =   67
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame fraBaixaFinal 
            Caption         =   "Dados da Baixa Física (Com o CTC Físico)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   120
            TabIndex        =   55
            Top             =   2640
            Width           =   4575
            Begin VB.CheckBox chkCanhoto 
               Caption         =   "Possui o Canhoto da Nota Fiscal do Cliente ?"
               Height          =   195
               Left            =   120
               TabIndex        =   58
               Top             =   1080
               Width           =   3495
            End
            Begin VB.TextBox txtRecBx 
               BackColor       =   &H8000000E&
               Height          =   285
               Left            =   1320
               MaxLength       =   25
               TabIndex        =   57
               Top             =   720
               Width           =   3135
            End
            Begin VB.CommandButton Command4 
               Caption         =   "EXCLUIR esta Baixa Física"
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               TabIndex        =   56
               Top             =   1440
               Width           =   3135
            End
            Begin VB.Label lblHsBx 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3600
               TabIndex        =   63
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
               Height          =   195
               Left            =   3120
               TabIndex        =   62
               Top             =   360
               Width           =   390
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Data Entrega:"
               Height          =   195
               Left            =   120
               TabIndex        =   61
               Top             =   360
               Width           =   990
            End
            Begin VB.Label lblDtBx 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1320
               TabIndex        =   60
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Recebedor...:"
               Height          =   195
               Left            =   120
               TabIndex        =   59
               Top             =   600
               Width           =   975
            End
         End
         Begin VB.OptionButton optBaixaFinal 
            Caption         =   "Baixa Física ou Ambas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   54
            ToolTipText     =   "Considerado Data de Entrega Independente da data de Pré-Baixa"
            Top             =   360
            Width           =   2295
         End
         Begin VB.OptionButton optPreBaixa 
            Caption         =   "Pré Baixa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   53
            ToolTipText     =   "Considerado com Data de Entrega na ausência de Baixa Física"
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCodOcorr 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Left            =   2400
         TabIndex        =   10
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hs:"
         Height          =   195
         Left            =   2160
         TabIndex        =   31
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cód. de Ocorrência:"
         Height          =   195
         Left            =   3120
         TabIndex        =   30
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label lblDescOcorr 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados do CTC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   6615
      Begin VB.Frame Frame3 
         Caption         =   "Origem / Destino"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   6375
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Dest."
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblDestUf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5880
            TabIndex        =   47
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblDestCid 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4080
            TabIndex        =   46
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblDest 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   600
            TabIndex        =   45
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Rem."
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblRemet 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   600
            TabIndex        =   21
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label lblRemetCid 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4080
            TabIndex        =   20
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblRemetUf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5880
            TabIndex        =   19
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Label lblMensagem 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Mensagem:"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   6375
      End
      Begin VB.Label lblfilialctc 
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
         Left            =   720
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Filialctc:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emissão/Hs:"
         Height          =   195
         Left            =   2040
         TabIndex        =   26
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblDtEmiss 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblHsEmiss 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modal:"
         Height          =   195
         Left            =   4800
         TabIndex        =   23
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblModal 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5400
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame frame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7320
      TabIndex        =   16
      Top             =   120
      Width           =   3720
      Begin VB.CommandButton cmbSair 
         Caption         =   "Canc/Sair"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton cmbGravar 
         Caption         =   "Gravar a Ocorr."
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "Procurar..."
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame fraOcorrencias 
      Caption         =   "Atuais Ocorrências deste CTC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "Comentários/Observações de Ocorrência ..."
         Height          =   255
         Left            =   2880
         TabIndex        =   51
         Top             =   2040
         Width           =   3615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Excluir a Ocorrência"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2040
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid gridOcorr 
         Bindings        =   "frmPodNF.frx":0008
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483634
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
         DataMember      =   "Sel_ConsOcorr"
         ColumnCount     =   10
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "cod_ocorr"
            Caption         =   "Cd."
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
            DataField       =   "descr_ocorr"
            Caption         =   "Ocorrência"
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
            DataField       =   "usu_ocorr"
            Caption         =   "usu_ocorr"
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
            DataField       =   "usu_dataocorr"
            Caption         =   "usu_dataocorr"
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
            DataField       =   "obs_ocorr"
            Caption         =   "obs_ocorr"
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
            DataField       =   "rel_arq_data"
            Caption         =   "rel_arq_data"
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
            DataField       =   "rel_arq_num"
            Caption         =   "rel_arq_num"
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
            DataField       =   "codigo"
            Caption         =   "codigo"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   524,976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   315,213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   4589,858
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1289,764
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   615
      Left            =   11160
      Picture         =   "frmPodNF.frx":0021
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblbxfinalSim 
      Height          =   255
      Left            =   4800
      TabIndex        =   34
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblcontroletela 
      AutoSize        =   -1  'True
      Caption         =   "normal"
      Height          =   195
      Left            =   3960
      TabIndex        =   33
      Top             =   7920
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "frmPodNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Sub chkBaixa_Click()
    'definições das escolhas de prébaixa, baixa final ou ambas
 '   If chkBaixa.Value = 1 Then
  '      optBaixaFinal.Enabled = False
   '     optPreBaixa.Enabled = False
    '    fraBaixaFinal.Enabled = False
     '   fraPreBaixa.Enabled = True
      '   fraPreBaixa.Caption = "Dados da Baixa (Pré e Final)"
'        txtRecPreBx.BackColor = &HC0FFFF      'AMARELO
'        txtRecBx.BackColor = &H8000000E   'branco
    'ElseIf chkBaixa.Value = 0 Then
'        optBaixaFinal.Enabled = True
'        optPreBaixa.Enabled = True
'        fraBaixaFinal.Enabled = True
'        fraPreBaixa.Enabled = True
'         fraPreBaixa.Caption = "Dados da Pré Baixa"
'        If optPreBaixa.Value = True Then
'            optPreBaixa_Click
'        Else
'            optBaixaFinal_Click
'        End If
'    End If
'End Sub

Private Sub chkObsEmiss_Click()

End Sub

Private Sub chkCanhoto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub cmbGravar_Click()

frmPodNF.MousePointer = 11

    Dim xcanhoto As String
    Dim xcontnfscanhoto As Long
    Dim xdatabaixa As Variant
    Dim xcont As Long
    Dim xdataservidor As Date
    Dim xtemocorr As String
    Dim xtemocorrCTC As String
    
    If optCTC = True Then
        xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
    ElseIf optNf = True Then
        xtemocorr = de_informa.rsSel_NFsdoCTC.Fields("tem_ocorr")
        xtemocorrCTC = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
        If xtemocorr <> xtemocorrCTC Then
            lblMensagem = " Mensagem: O CTC desta NF Está Com Status Diferente. Para Inform. Consulte o CTC."
        End If
    End If
    
    
    xcanhoto = ""
                 
'TRATAMENTO DE  B A I X A S
    
    If txtCodOcorr.Text = "01" Then  'se for "01" (entrega realizada/baixa)
        
        If optCTC.Value = True Then  'a baixa é de um CTC
            'verifica se o CTC já tem ocorrência fechada cadastrada. Caso tenha não possibilita a baixa
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "0" Then
                frmPodNF.MousePointer = 0
                MsgBox "Este CTC está Baixado, indicando que esta entrega não ocorreria. Caso deseje baixar como ENTREGA, você deve primeiro excluir a Ocorrência  C T C   B A I X A D O", vbOKCancel
                txtCodOcorr.SetFocus
                Exit Sub
            'verifica se o ctc já possui ocorrência e se a data que se quer baixar não é menor que a data de alguma ocorrência
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
                de_informa.rsSel_ConsOcorrCTCGrid.MoveFirst
                Do Until de_informa.rsSel_ConsOcorrCTCGrid.EOF
                    If CDate(mskData.Text) < de_informa.rsSel_ConsOcorrCTCGrid.Fields("data") Then
                        frmPodNF.MousePointer = 0
                        If MsgBox("Você está tentando baixar um CTC com data Menor que uma Ocorrência Cadastrada. Você tem certeza que quer baixar este CTC como entrega nesta data ?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then
                            mskData.SetFocus
                            Exit Sub
                        Else
                            Exit Do
                        End If
                    End If
                    de_informa.rsSel_ConsOcorrCTCGrid.MoveNext
                Loop
            End If
        ElseIf optNf.Value = True Then  'a baixa de por NF
            de_informa.rsSel_NFsdoCTC.MoveFirst
            Do Until de_informa.rsSel_NFsdoCTC.Fields("numnfnum") = Val(txtNumNf)
                de_informa.rsSel_NFsdoCTC.MoveNext
            Loop
            
            'verifica se a NF já tem ocorrência fechada cadastrada. Caso tenha não possibilita a baixa
            If de_informa.rsSel_NFsdoCTC.Fields("tem_ocorr") = "0" Then
                frmPodNF.MousePointer = 0
                MsgBox "Esta NF deste CTC está Baixado, indicando que a entrega DESTA NF não ocorreria. Caso deseje baixar como ENTREGA, você deve primeiro excluir a Ocorrência  CTC/NF  B A I X A D O", vbOKCancel
                txtCodOcorr.SetFocus
                Exit Sub
            'verifica se o ctc já possui ocorrência e se a data que se quer baixar não é menor que a data de alguma ocorrência
            ElseIf de_informa.rsSel_NFsdoCTC.Fields("tem_ocorr") = "2" Then
                de_informa.rsSel_ConsOcorrNF.MoveFirst
                Do Until de_informa.rsSel_ConsOcorrNF.EOF
                    If CDate(mskData.Text) < de_informa.rsSel_ConsOcorrNF.Fields("data") Then
                        frmPodNF.MousePointer = 0
                        If MsgBox("Você está tentando baixar uma NF de um CTC com data Menor que uma Ocorrência Cadastrada. Você tem certeza que quer baixar este CTC como entrega nesta data ?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then
                            mskData.SetFocus
                            Exit Sub
                        Else
                            Exit Do
                        End If
                    End If
                    de_informa.rsSel_ConsOcorrNF.MoveNext
                Loop
            End If
        End If
            
        'busca dados de data e hora do servidor

        If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
        de_informa.Sel_DataServidor
        xdataservidor = de_informa.rsSel_DataServidor.Fields("agora")

        If optCTC = True Then
        
            'preenche e traz as NFs deste CTC para confirmação de sua baixa (form de NFS)
        
            frmContrNFOcorr.lstGravarNFS.Clear
            frmContrNFOcorr.lstNAOGravarNFS.Clear
            frmContrNFOcorr.lblfilialctc = lblfilialctc
            frmContrNFOcorr.optFilialctc = True
            frmContrNFOcorr.lblPorFilialctc = lblfilialctc
        
            de_informa.rsSel_NFsdoCTC.MoveFirst
            Do Until de_informa.rsSel_NFsdoCTC.EOF
                frmContrNFOcorr.lstGravarNFS.AddItem de_informa.rsSel_NFsdoCTC.Fields("numnf")
                de_informa.rsSel_NFsdoCTC.MoveNext
            Loop
        
            frmContrNFOcorr.Show 1
        
            If lblcontroletela.Caption = "cancelar" Then
                lblcontroletela.Caption = "normal"
                Unload frmContrNFOcorr
                Me.MousePointer = 0
                cmdProcurar_Click
                Exit Sub
            End If
        End If
        
        'início do processo de baixa.

        If optPreBaixa.Value = True Then    'se for uma pré baixa
        
            'inicia a transação do processo de pré-baixa
            de_informa.cn_informa.BeginTrans
                
            If optCTC = True Then
                'faz a baixa em cada NF/CTC escolhida no LST Gravar
                For xcont = 0 To frmContrNFOcorr.lstGravarNFS.ListCount - 1
                        
                    'busca no BD o Filialctc e a NF com a ocorr de entrega
                    If de_informa.rsSel_ConsOcorrNF.State = 1 Then de_informa.rsSel_ConsOcorrNF.Close
                    de_informa.Sel_ConsOcorrNF lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), "01"
                            
                    If de_informa.rsSel_ConsOcorrNF.RecordCount > 0 Then 'encontrou NF com baixa
                        MsgBox "A Nota Fiscal Número " & Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)) & " do CTC " & lblfilialctc & " já está baixada ! Para baixá-la com outra data você deve primeiro excluir a baixa atual !"
                    Else '
                        'grava os dados de ocorr/entrega desta nf/CTC
                        de_informa.ins_ocorr01pre lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), _
                        de_informa.rsSel_Ctc_SAC.Fields("data"), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                        txtCodOcorr.Text, lblDescOcorr.Caption, CDate(mskData), mskHora, CDate(mskData), mskHora, _
                        txtRecPreBx, xusuario, xdataservidor
                        'atualiza o tem_ocorr da NF
                        de_informa.alt_temocorr_nf "1", lblfilialctc, frmContrNFOcorr.lstGravarNFS.List(xcont)
                        'atualiza o tem_ocorr do CTC
                        de_informa.alt_temocorr_sn "1", lblfilialctc
                        'atualiza o AT_Cliente
                        de_informa.Alt_AtClienteNFBranco lblfilialctc, frmContrNFOcorr.lstGravarNFS.List(xcont)
                        'atualiza o LOG de usuario
                        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC/NF: " & lblfilialctc & "/" & frmContrNFOcorr.lstGravarNFS.List(xcont) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " PRÉ-BAIXA"
                    End If
                Next
            ElseIf optNf = True Then
                
                'busca no BD o Filialctc e a NF com a ocorr de entrega
                If de_informa.rsSel_ConsOcorrNF.State = 1 Then de_informa.rsSel_ConsOcorrNF.Close
                de_informa.Sel_ConsOcorrNF lblfilialctc, Trim$(txtNumNf), "01"
                            
                If de_informa.rsSel_ConsOcorrNF.RecordCount > 0 Then 'encontrou NF com baixa
                    MsgBox "A Nota Fiscal Número " & Trim$(txtNumNf) & " do CTC " & lblfilialctc & " já está baixada ! Para baixá-la com outra data você deve primeiro excluir a baixa atual !"
                Else '
                    'grava os dados de ocorr/entrega desta nf/CTC
                    de_informa.ins_ocorr01pre lblfilialctc, Trim$(txtNumNf), _
                    de_informa.rsSel_Ctc_SAC.Fields("data"), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                    txtCodOcorr.Text, lblDescOcorr.Caption, CDate(mskData), mskHora, CDate(mskData), mskHora, _
                    txtRecPreBx, xusuario, xdataservidor
                    'atualiza o tem_ocorr da NF
                    de_informa.alt_temocorr_nf "1", lblfilialctc, Trim$(txtNumNf)
                    'atualiza o tem_ocorr do CTC
                    de_informa.alt_temocorr_sn "1", lblfilialctc
                    'atualiza o AT_Cliente
                    de_informa.Alt_AtClienteNFBranco lblfilialctc, Trim$(txtNumNf)
                    'atualiza o LOG de usuario
                    de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC/NF: " & lblfilialctc & "/" & Trim$(txtNumNf) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " PRÉ-BAIXA"
                End If
            End If
                
            'finaliza a transação do processo de pré-baixa
            de_informa.cn_informa.CommitTrans

            If optCTC = True Then
                'fecha o form de controle das NFs que sofrerão ocorrência
                Unload frmContrNFOcorr
            End If

        ElseIf optBaixaFinal.Value = True Then   'se for uma baixa física
                    
            'CONTROLE DOS CANHOTOS DAS NFS
            If chkCanhoto.Value = 1 Then
                frmContrCanhotos.lstPresentes.Clear
                frmContrCanhotos.lstFaltantes.Clear
                xcanhoto = "S"
                If chkCanhoto.Enabled = True Then
                    If optCTC = True Then
                        'mudar recordset para as NFS dos LST Gravar, SE FOR POR CTC
                        For xcont = 0 To frmContrNFOcorr.lstGravarNFS.ListCount - 1
                            frmContrCanhotos.lstPresentes.AddItem frmContrNFOcorr.lstGravarNFS.List(xcont)
                        Next
                    ElseIf optNf = True Then
                            frmContrCanhotos.lstPresentes.AddItem Trim$(txtNumNf)
                    End If
                    frmContrCanhotos.lblfilialctc = lblfilialctc
                    frmContrCanhotos.fraPresentes.Caption = frmContrCanhotos.lstPresentes.ListCount & " Canhotos"
                    frmContrCanhotos.Show 1
                    If lblcontroletela.Caption = "cancelar" Then
                        lblcontroletela.Caption = "normal"
                        Unload frmContrCanhotos
                        Unload frmContrNFOcorr
                        Me.MousePointer = 0
                        cmdProcurar_Click
                        Exit Sub
                    End If
                Else
                    xcanhoto = "N"
                End If
            Else
                xcanhoto = "N"
            End If
            
            'inicia a transação do processo de baixa-física
            de_informa.cn_informa.BeginTrans
            
            If optCTC = True Then
                
                'faz a baixa em cada NF/CTC escolhida no LST Gravar
                For xcont = 0 To frmContrNFOcorr.lstGravarNFS.ListCount - 1
                        
                    'busca no BD o Filialctc e a NF com a ocorr de entrega
                    If de_informa.rsSel_ConsOcorrNF.State = 1 Then de_informa.rsSel_ConsOcorrNF.Close
                    de_informa.Sel_ConsOcorrNF lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), "01"
                            
                    If de_informa.rsSel_ConsOcorrNF.RecordCount > 0 Then 'encontrou NF com baixa
                        'verifica se a baixa contida no bd é física. Se for, não deixar baixar novamente
                        'se não for, é porque já está baixado como pré. Somente atualizar com os dados da
                        'baixa física
                        If de_informa.rsSel_ConsOcorrNF.Fields("baixadofinal") = "S" Then
                            MsgBox "A Nota Fiscal Número " & Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)) & " do CTC " & lblfilialctc & " já está baixada ! Para baixá-la com outra data você deve primeiro excluir a baixa atual !"
                        Else
                            'atualiza com os dados de baixa física desta nf/CTC
                            de_informa.alt_ocorr01fisico lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), _
                            mskData, mskHora, mskData, mskHora, txtRecBx, xusuario, xdataservidor, "S", xdataservidor, xcanhoto
                            'Se a data de Entrega Bx Física é diferente da pré baixa, atualiza o cliente
                            If de_informa.rsSel_ConsOcorrNF.Fields("dtbaixapre") <> CDate(mskData) Then
                                'atualiza o AT_Cliente
                                de_informa.Alt_AtClienteNFBranco lblfilialctc, frmContrNFOcorr.lstGravarNFS.List(xcont)
                            End If
                            'atualiza o LOG de usuario
                            de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC/NF: " & lblfilialctc & "/" & frmContrNFOcorr.lstGravarNFS.List(xcont) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " BAIXA FÍSICA"
                        End If
                    Else  'este CTC/NF não contém Baixa, então INCLUE uma baixa física
                        'grava os dados de ocorr/entrega desta nf/CTC
                        de_informa.ins_ocorr01fis lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), _
                        de_informa.rsSel_Ctc_SAC.Fields("data"), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                        txtCodOcorr.Text, lblDescOcorr.Caption, CDate(mskData), mskHora, CDate(mskData), mskHora, _
                        txtRecBx, xusuario, xdataservidor, xcanhoto
                        'atualiza o tem_ocorr da NF
                        de_informa.alt_temocorr_nf "1", lblfilialctc, frmContrNFOcorr.lstGravarNFS.List(xcont)
                        'atualiza o tem_ocorr do CTC
                        de_informa.alt_temocorr_sn "1", lblfilialctc
                        'atualiza o AT_Cliente
                        de_informa.Alt_AtClienteNFBranco lblfilialctc, frmContrNFOcorr.lstGravarNFS.List(xcont)
                        'atualiza o LOG de usuario
                        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC/NF: " & lblfilialctc & "/" & frmContrNFOcorr.lstGravarNFS.List(xcont) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " BAIXA FÍSICA"
                    End If
                Next
                
            ElseIf optNf = True Then
                'busca no BD o Filialctc e a NF com a ocorr de entrega
                If de_informa.rsSel_ConsOcorrNF.State = 1 Then de_informa.rsSel_ConsOcorrNF.Close
                de_informa.Sel_ConsOcorrNF lblfilialctc, Trim$(txtNumNf), "01"
                            
                If de_informa.rsSel_ConsOcorrNF.RecordCount > 0 Then 'encontrou NF com baixa
                    'verifica se a baixa contida no bd é física. Se for, não deixar baixar novamente
                    'se não for, é porque já está baixado como pré. Somente atualizar com os dados da
                    'baixa física
                    If de_informa.rsSel_ConsOcorrNF.Fields("baixadofinal") = "S" Then
                        MsgBox "A Nota Fiscal Número " & Trim$(txtNumNf) & " do CTC " & lblfilialctc & " já está baixada ! Para baixá-la com outra data você deve primeiro excluir a baixa atual !"
                    Else
                        'atualiza com os dados de baixa física desta nf/CTC
                        de_informa.alt_ocorr01fisico lblfilialctc, Trim$(txtNumNf), mskData, mskHora, mskData, mskHora, _
                        txtRecBx, xusuario, xdataservidor, "S", xdataservidor, xcanhoto
                        'Se a data de Entrega Bx Física é diferente da pré baixa, atualiza o cliente
                        If de_informa.rsSel_ConsOcorrNF.Fields("dtbaixapre") <> CDate(mskData) Then
                            'atualiza o AT_Cliente
                            de_informa.Alt_AtClienteNFBranco lblfilialctc, Trim$(txtNumNf)
                        End If
                        'atualiza o LOG de usuario
                        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC/NF: " & lblfilialctc & "/" & Trim$(txtNumNf) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " BAIXA FÍSICA"
                    End If
                Else  'este CTC/NF não contém Baixa, então INCLUE uma baixa física
                        'grava os dados de ocorr/entrega desta nf/CTC
                    de_informa.ins_ocorr01fis lblfilialctc, Trim$(txtNumNf), _
                    de_informa.rsSel_Ctc_SAC.Fields("data"), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                    txtCodOcorr.Text, lblDescOcorr.Caption, CDate(mskData), mskHora, CDate(mskData), mskHora, _
                    txtRecBx, xusuario, xdataservidor, xcanhoto
                    'atualiza o tem_ocorr da NF
                    de_informa.alt_temocorr_nf "1", lblfilialctc, Trim$(txtNumNf)
                    'atualiza o tem_ocorr do CTC
                    de_informa.alt_temocorr_sn "1", lblfilialctc
                    'atualiza o AT_Cliente
                    de_informa.Alt_AtClienteNFBranco lblfilialctc, Trim$(txtNumNf)
                    'atualiza o LOG de usuario
                    de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC/NF: " & lblfilialctc & "/" & Trim$(txtNumNf) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " BAIXA FÍSICA"
                End If
            End If
                
            'atualiza as NFs que contém o canhoto
                        
            For xcontnfscanhoto = 1 To frmContrCanhotos.lstPresentes.ListCount
                de_informa.Alt_CanhotoNFSN "S", lblfilialctc, frmContrCanhotos.lstPresentes.List(xcontnfscanhoto - 1)
            Next
                        
            'atualiza as NFs que NÃO contém o canhoto
                        
            For xcontnfscanhoto = 1 To frmContrCanhotos.lstFaltantes.ListCount
                de_informa.Alt_CanhotoNFSN "N", lblfilialctc, frmContrCanhotos.lstFaltantes.List(xcontnfscanhoto - 1)
            Next
                    
            'finaliza a transação do processo de pré-baixa
            de_informa.cn_informa.CommitTrans
                
            'fecha o form de controle das NFs que sofrerão ocorrência
            If optCTC = True Then
                Unload frmContrNFOcorr
            End If
            Unload frmContrCanhotos
        
        End If
            
'TRATAMENTO DE  O C O R R Ê N C I A S
        
    Else   'se nao for baixa (ocorr # 01) então é somente ocorrência
        
        If txtCodOcorr.Text = "00" Then   'se for ocorr 00
            If xtemocorr = "1" Then
                frmPodNF.MousePointer = 0
                MsgBox "CTC/NF já Baixado como Entregue. Não é Possível informar Ocorrência  C T C   B A I X A D O"
                txtCodOcorr.SetFocus
                Exit Sub
            ElseIf xtemocorr = "0" Then
                frmPodNF.MousePointer = 0
                MsgBox "CTC/NF já possui Ocorrência  C T C   B A I X A D O"
                txtCodOcorr.SetFocus
                Exit Sub
            ElseIf xtemocorr = "N" Then
                frmPodNF.MousePointer = 0
                MsgBox "Você só pode informar Ocorrência  C T C   B A I X A D O, se o CTC já tiver alguma ocorrência que a explique o motivo."
                txtCodOcorr.SetFocus
                Exit Sub
            End If
        Else   'se não for é ocorrência normal
            If xtemocorr = "1" Then
                If IsDate(lblDtBx.Caption) Then
                    If CDate(mskData.Text) > CDate(lblDtBx.Caption) Then
                        frmPodNF.MousePointer = 0
                        If MsgBox("Você está tentando incluir uma Ocorrência com data Posterior à sua Data de Entrega. Você tem certeza que deseja informar esta ocorrência com esta data ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                            mskData.SetFocus
                            Exit Sub
                        End If
                    End If
                ElseIf IsDate(lblDtPreBx.Caption) Then
                    If CDate(mskData.Text) > CDate(lblDtPreBx.Caption) Then
                        frmPodNF.MousePointer = 0
                        If MsgBox("Você está tentando incluir uma Ocorrência com data Posterior à Data de Entrega. Você tem certeza que deseja informar esta ocorrência com esta data ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                            mskData.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        
        'BUSCA A DATA/HORA DO SERVIDOR
        
        If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
        de_informa.Sel_DataServidor
        xdataservidor = de_informa.rsSel_DataServidor.Fields("agora")

        If optCTC = True Then
        
            'preenche e traz as NFs deste CTC para confirmação da ocorrência (form de NFS)
        
            frmContrNFOcorr.lstGravarNFS.Clear
            frmContrNFOcorr.lstNAOGravarNFS.Clear
            frmContrNFOcorr.lblfilialctc = lblfilialctc
            frmContrNFOcorr.optFilialctc = True
            frmContrNFOcorr.lblPorFilialctc = lblfilialctc
        
            de_informa.rsSel_NFsdoCTC.MoveFirst
            Do Until de_informa.rsSel_NFsdoCTC.EOF
                frmContrNFOcorr.lstGravarNFS.AddItem de_informa.rsSel_NFsdoCTC.Fields("numnf")
                de_informa.rsSel_NFsdoCTC.MoveNext
            Loop
        
            frmContrNFOcorr.Show 1
        
            If lblcontroletela.Caption = "cancelar" Then
                lblcontroletela.Caption = "normal"
                Unload frmContrNFOcorr
                Me.MousePointer = 0
                cmdProcurar_Click
                Exit Sub
            End If
        End If
        
        'ATUALIZA BD COM OS DADOS DA OCORRÊNCIA
        
        
'**************************   P A R E I    A Q U I   ********************
        
'FICAR ATENDO PARA O ALT_TEMOCORR POR NF E CTC, QUANDO FOR "00" E SE O CTC JÁ ESTIVER ENTREGE
        
        
        'INICIA A TRANSAÇÃO
        de_informa.cn_informa.BeginTrans
        
        If optCTC = True Then
            
            If txtCodOcorr.Text = "00" Then   'se for ocorr 00
                de_informa.ins_ocorrencia00 lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), CDate(frmPodNF.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, xusuario, xdataservidor
                '0 = IDENTIFICA COMO CTC COM OCORRÊNCIA FECHADA
                
                'atualiza o tem_ocorr da NF
                de_informa.alt_temocorr_nf "0", lblfilialctc, frmContrNFOcorr.lstGravarNFS.List(xcont)
                'atualiza o LOG de usuario
                de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "POD/OCORR - CTC/NF: " & lblfilialctc & "/" & frmContrNFOcorr.lstGravarNFS.List(xcont) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr
            
            Else  'se for outro tipo de ocorrência
                de_informa.ins_ocorrencia lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), _
                CDate(frmPodNF.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, xusuario, xdataservidor
                '2 = IDENTIFICA COMO CTC COM OCORRÊNCIAS PENDENTE
                de_informa.Alt_AtClienteNFBranco lblfilialctc, ""
                If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "N" Then
                    de_informa.alt_temocorr_nf "2", lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont))   'atualiza arquivo de CTC com tem_ocorr = 2
                End If
            End If
            
        ElseIf optNf = True Then
            
            If txtCodOcorr.Text = "00" Then   'se for ocorr 00
                de_informa.ins_ocorrencia00 lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), CDate(frmPodNF.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, xusuario, xdataservidor
                '0 = IDENTIFICA COMO CTC COM OCORRÊNCIA FECHADA
                If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
                    de_informa.alt_temocorr_nf "0", lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont))  'atualiza arquivo de CTC com tem_ocorr = 0
                End If
            Else  'se for outro tipo de ocorrência
                de_informa.ins_ocorrencia lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont)), _
                CDate(frmPodNF.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, xusuario, xdataservidor
                '2 = IDENTIFICA COMO CTC COM OCORRÊNCIAS PENDENTE
                de_informa.Alt_AtClienteNFBranco lblfilialctc, ""
                If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "N" Then
                    de_informa.alt_temocorr_nf "2", lblfilialctc, Trim$(frmContrNFOcorr.lstGravarNFS.List(xcont))   'atualiza arquivo de CTC com tem_ocorr = 2
                End If
            End If
            
        End If
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "INCLUSAO", xusuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr
        
        de_informa.cn_informa.CommitTrans
        
    End If
    
    mskData.Mask = ""
    mskData.Text = ""
    mskData.Mask = "##/##/####"
    mskHora.Mask = ""
    mskHora.Text = ""
    mskHora.Mask = "##:##"
    
'tratamento de acerto aws (data de emissão)------------------------------------------------

   ' mskEmissaoNova.Mask = ""
   ' mskEmissaoNova.Text = ""
   ' mskEmissaoNova.Mask = "##/##/####"

'------------------------------------------------------------------------------------------
    
    
    frmPodNF.MousePointer = 0
    txtCodOcorr = ""
    lblDescOcorr.Caption = ""
    txtRecPreBx.BackColor = &H8000000E   'branco
    txtRecBx.BackColor = &H8000000E   'branco
    cmdProcurar_Click
    If optCTC = True Then
        txtFilial.SetFocus
    ElseIf optNf = True Then
        txtNumNf.SetFocus
    End If
End Sub
Private Sub cmbSair_Click()
    If mskData.Text <> "__/__/____" Then   'quer dizer que é CANCELAR
        mskData.Mask = ""
        mskData.Text = ""
        mskData.Mask = "##/##/####"
        mskHora.Mask = ""
        mskHora.Text = ""
        mskHora.Mask = "##:##"
        txtCodOcorr = ""
        lblDescOcorr.Caption = ""
        txtRecPreBx.Text = ""
        txtRecBx.Text = ""
        txtRecPreBx.BackColor = &H8000000E   'branco
        txtRecBx.BackColor = &H8000000E   'branco
        cmdProcurar_Click
        txtFilial.SetFocus
    Else                            'caso contrário é SAIR
        frmAtualPrazos.Show 1
        If de_informa.rsSel_ConsOcorrNF.State = 1 Then de_informa.rsSel_ConsOcorrNF.Close
        If de_informa.rsSel_ConsOcorrCTCGrid.State = 1 Then de_informa.rsSel_ConsOcorrCTCGrid.Close
        If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
        If de_informa.rsSel_NfsComOcorr.State = 1 Then de_informa.rsSel_NfsComOcorr.Close
        If de_informa.rsSel_NfsNaoOcorr.State = 1 Then de_informa.rsSel_NfsNaoOcorr.Close
        gridOcorr.DataMember = "Sel_ConsOcorrCTC"
        gridOcorr.Refresh
        Set frmPodNF = Nothing
        Unload Me
    End If
End Sub


Private Sub cmdCalendario_Click()

End Sub

Private Sub cmdExclBx_Click()
    If Mid$(xdireitos, 23, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        'Alteração no Registro de Baixa, (UPDATE) para BAIXAFINAL = N
        'Atualização do Campo DATA do TB_OCORR para a Data da Pré-Baixa
        If MsgBox("Confirma Exclusão dos Dados de BAIXA FÍSICA ? ", vbYesNo, "Atenção") = vbYes Then
        
            de_informa.cn_informa.BeginTrans
            
            de_informa.Alt_ExclBaixaFisica transctc(txtFilial, txtCtc)
            
            de_informa.alt_ExclCanhotoNF transctc(txtFilial, txtCtc)
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " BAIXA FÍSICA"
            
            de_informa.cn_informa.CommitTrans
            
            cmdProcurar_Click
        End If
    End If
End Sub

Private Sub cmdExclOcorr_Click()
    If Mid$(xdireitos, 22, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        Dim xcodocorr As String
        If MsgBox("Confirma a Exclusão da Ocorrência Selecionada ?", vbQuestion + vbYesNo, "Exclusão") = vbYes Then
            xcodocorr = gridOcorr.Columns(2)
            
            de_informa.cn_informa.BeginTrans
            
            de_informa.excl_ocorr gridOcorr.Columns(9)
            If xcodocorr = "00" Then 'se for "00" altera o temocorr para 02
                de_informa.alt_temocorr_sn "2", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr
            End If
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & gridOcorr.Columns(2) & "-" & gridOcorr.Columns(3)
            
            'busca as ocorrências e atualiza o grid de ocorrências
            
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 transctc(txtFilial, txtCtc), "01"
            Set gridOcorr.DataSource = de_informa
            gridOcorr.DataMember = "Sel_ConsOcorr2"
            gridOcorr.Refresh
            
         'verifica se é a última ocorrência baixada e se é ocorr "00"
        'se for exclui ela também (pois o processo não está finalizado) e atualiza o grid novamente
 
            If de_informa.rsSel_ConsOcorr2.RecordCount = 1 Then
                If de_informa.rsSel_ConsOcorr2.Fields("cod_ocorr") = "00" Then
                    de_informa.excl_ocorr gridOcorr.Columns(9)
                    If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                    de_informa.Sel_ConsOcorr2 transctc(txtFilial, txtCtc), "01"
                    Set gridOcorr.DataSource = de_informa
                    gridOcorr.DataMember = "Sel_ConsOcorr2"
                    gridOcorr.Refresh
                End If
            End If
            
        'se o grid estiver vazio e se não estiver baixa o CTC atualiza o temocorr para "N" (sem posição)
            If de_informa.rsSel_ConsOcorr2.RecordCount = 0 And de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") <> "1" Then  'verifica se não há mais ocorrência e se não está baixado
                de_informa.alt_temocorr_sn "N", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr
            End If
                
            de_informa.cn_informa.CommitTrans
                
            cmdProcurar_Click
        End If
    End If
End Sub



Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    
End Sub
Private Sub cmdProcurar_Click()
    
'versão sac novo

Dim xagora As Date, xtemocorrCTC As String, xtemocorr As String
    
'BUSCA DATA DO SERVIDOR
If de_informa.rsSel_DataServidor.State = 1 Then de_informa.rsSel_DataServidor.Close
de_informa.Sel_DataServidor
xagora = de_informa.rsSel_DataServidor.Fields("agora")

'grava nas variaveis globais os ctcs / nfs consultados
    
xultimofilial = txtFilial
xultimoctc = txtCtc
xultimonf = txtNumNf

'Limpa os grids do form
gridOcorr.DataMember = ""
gridOcorr.Refresh

'limpa os Label/Outros Objetos do Form

Call limpatela(Me)

lstNfOcorr.Clear
lstNfOcorrNAO.Clear
lblEntregueSN = ""
lblOcorrPorNF = "0 Nfs tiveram a Ocorr. Abaixo"
lblOcorrPorNFNao = "0 Nfs NÃO tiveram a Ocorr. Abaixo"
fraNfsOcorr = "Total de Notas Fiscais do CTC: 0 NF(s)"
lblMensagem = " Mensagem:"

'LIMPA MSK
mskData.Mask = ""
mskData.Text = ""
mskData.Mask = "##/##/####"
mskHora.Mask = ""
mskHora.Text = ""
mskHora.Mask = "##:##"

'trata text box
mskData.Enabled = False
mskData.BackColor = &HFFFFFF
mskHora.Enabled = False
mskHora.BackColor = &HFFFFFF
txtCodOcorr.Enabled = False
txtCodOcorr.BackColor = &HFFFFFF
txtObs_Ocorr.Enabled = False
txtObs_Ocorr.BackColor = &HFFFFFF

'volta os text filial/ctc/nf para os valores globais

txtFilial = xultimofilial
txtCtc = xultimoctc
txtNumNf = xultimonf

If optCTC.Value = True Then   'Se for procura por Filial / CTC
        
    lblfilialctc = transctc(txtFilial, txtCtc)
    
    'verifica consistência dos dados
        
    If txtFilial.Text = "" Or txtCtc.Text = "" Then
        MsgBox "Filial / CTC Inválidos !", vbCritical, "Erro"
        txtFilial.SetFocus
        Exit Sub
    End If
    If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
    de_informa.Sel_Ctc_SAC lblfilialctc 'Procura na Tabela a Filial/CTC
    If de_informa.rsSel_Ctc_SAC.RecordCount = 0 Then
        MsgBox "Número de Filial/CTC Não Encontrados !", vbCritical + vbOKOnly, "Erro"
        txtFilial.SetFocus
        Exit Sub
    End If

    'BUSCA AS NOTAS FISCAIS DESTE CTC
    If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
    de_informa.Sel_NFsdoCTC lblfilialctc
    
ElseIf optNf.Value = True Then   'Se for procura por Nota Fiscal

    'verifica consistência dos dados
        
    If txtNumNf.Text = "" Then
        MsgBox "Número de Nota Fiscal Inválida !", vbCritical, "Erro"
        txtNumNf.SetFocus
        Exit Sub
    End If
    If de_informa.rsSel_NF_SAC.State = 1 Then de_informa.rsSel_NF_SAC.Close
    de_informa.Sel_NF_SAC Val(txtNumNf)   'Procura a NF na Tabela
    If de_informa.rsSel_NF_SAC.RecordCount = 0 Then
        MsgBox "Número de NF Não Encontrado !", vbCritical + vbOKOnly, "Erro"
        txtNumNf.SetFocus
        Exit Sub
    ElseIf de_informa.rsSel_NF_SAC.RecordCount > 1 Then  'achou mais de uma NF com o mesmo número
        frmDuplNF.Caption = "POD - Número de NFs Duplicadas"
        frmDuplNF.Show 1  'direciona para o form que trata casos de NF duplicadas
    Else  'Caso seja encontrada somente uma NF
        lblfilialctc = de_informa.rsSel_NF_SAC.Fields("filialctc")
    End If
    
    'PROCURA O CTC NA TABELA
    If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
    de_informa.Sel_Ctc_SAC lblfilialctc  'Procura na Tabela a Filial/CTC

    If de_informa.rsSel_Ctc_SAC.RecordCount = 0 Then
        MsgBox "Erro de Consistência. Chame Suporte Técnico ! ", vbCritical + vbOKOnly, "Erro" 'erro de consistência
        txtNumNf.SetFocus
        Exit Sub
    End If
    
    'posiciona o registro na NF que está sendo consultada
    If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
    de_informa.Sel_NFsdoCTC lblfilialctc
    Do Until de_informa.rsSel_NFsdoCTC.Fields("numnfnum") = Val(txtNumNf)
        de_informa.rsSel_NFsdoCTC.MoveNext
    Loop
    
End If

'DADOS DO CTC

lblDtEmiss = de_informa.rsSel_Ctc_SAC.Fields("data")
lblHsEmiss = de_informa.rsSel_Ctc_SAC.Fields("hora")
lblModal = de_informa.rsSel_Ctc_SAC.Fields("modal")
    
If optCTC = True Then
    xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
ElseIf optNf = True Then
    xtemocorr = de_informa.rsSel_NFsdoCTC.Fields("tem_ocorr")
    xtemocorrCTC = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
    If xtemocorr <> xtemocorrCTC Then
        lblMensagem = " Mensagem: O CTC desta NF Está Com Status Diferente. Para Inform. Consulte o CTC."
    End If
End If
        
'Origem / Remetente
        
If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
de_informa.Sel_ConsCadCli de_informa.rsSel_Ctc_SAC.Fields("remet_cgc")
        
lblRemet = de_informa.rsSel_Ctc_SAC.Fields("remet_nome")
lblRemetCid = de_informa.rsSel_Ctc_SAC.Fields("cidade_orig")
If Not IsNull(de_informa.rsSel_ConsCadCli.Fields("uf")) Then
    lblRemetUf = de_informa.rsSel_ConsCadCli.Fields("uf")
End If
        
'Destino / Destinatário
        
If de_informa.rsSel_ConsCadCliDest.State = 1 Then de_informa.rsSel_ConsCadCliDest.Close
de_informa.Sel_ConsCadCliDest de_informa.rsSel_Ctc_SAC.Fields("dest_cgc")
        
lblDest = de_informa.rsSel_Ctc_SAC.Fields("dest_nome")
lblDestCid = de_informa.rsSel_Ctc_SAC.Fields("cidade_dest")
lblDestUf = de_informa.rsSel_Ctc_SAC.Fields("uf_dest")
        
'Ocorrências (GRID)
        
If optCTC = True Then
    If de_informa.rsSel_ConsOcorrCTC.State = 1 Then de_informa.rsSel_ConsOcorrCTC.Close
    de_informa.Sel_ConsOcorrCTC lblfilialctc, "%"
    
    fraOcorrencias = "Ocorrências deste CTC"
    gridOcorr.DataMember = "sel_consocorrctc"
    gridOcorr.Refresh
    
    If de_informa.rsSel_ConsOcorrCTC.RecordCount > 0 Then
        'observação de ocorrência
        de_informa.rsSel_ConsOcorrCTC.MoveFirst
        If Not IsNull(de_informa.rsSel_ConsOcorrCTC.Fields("obs_ocorr")) Then
            txtObs_Ocorr = de_informa.rsSel_ConsOcorrCTC.Fields("obs_ocorr")
        End If
    End If
ElseIf optNf = True Then

    If de_informa.rsSel_ConsOcorrNF.State = 1 Then de_informa.rsSel_ConsOcorrNF.Close
    de_informa.Sel_ConsOcorrNF lblfilialctc, txtNumNf, "%"
    
    fraOcorrencias = "Ocorrências desta Nota Fiscal"
    gridOcorr.DataMember = "sel_consocorrnf"
    gridOcorr.Refresh
    
    If de_informa.rsSel_ConsOcorrNF.RecordCount > 0 Then
        'observação de ocorrência
        de_informa.rsSel_ConsOcorrNF.MoveFirst
        If Not IsNull(de_informa.rsSel_ConsOcorrNF.Fields("obs_ocorr")) Then
            txtObs_Ocorr = de_informa.rsSel_ConsOcorrNF.Fields("obs_ocorr")
        End If
    End If
End If
        
fraNfsOcorr = "Total de Notas Fiscais do CTC: " & Trim$(Str(de_informa.rsSel_NFsdoCTC.RecordCount)) & " NF(s)"

If xtemocorr <> "N" And xtemocorr <> "C" Then
    'Total de Notas Fiscais deste CTC:
    If de_informa.rsSel_NfsComOcorr.State = 1 Then de_informa.rsSel_NfsComOcorr.Close
    de_informa.Sel_NfsComOcorr lblfilialctc, gridOcorr.Columns(2)
            
    If de_informa.rsSel_NfsNaoOcorr.State = 1 Then de_informa.rsSel_NfsNaoOcorr.Close
    de_informa.Sel_NfsNaoOcorr lblfilialctc, gridOcorr.Columns(2)
            
    Do Until de_informa.rsSel_NfsComOcorr.EOF  'preenche a LST Nf com Ocorr
        lstNfOcorr.AddItem de_informa.rsSel_NfsComOcorr.Fields("numnf")
        de_informa.rsSel_NfsComOcorr.MoveNext
    Loop
            
    lblOcorrPorNF = Trim$(Str(de_informa.rsSel_NfsComOcorr.RecordCount)) & _
                    " Nfs tiveram a Ocorr. Abaixo"
            
    Do Until de_informa.rsSel_NfsNaoOcorr.EOF  'preenche a LST Nf com Ocorr
        lstNfOcorrNAO.AddItem de_informa.rsSel_NfsNaoOcorr.Fields("numnf")
        de_informa.rsSel_NfsNaoOcorr.MoveNext
    Loop
            
    lblOcorrPorNFNao = Trim$(Str(de_informa.rsSel_NfsNaoOcorr.RecordCount)) & _
                    " Nfs NÃO tiveram a Ocorr. Abaixo"
            
    lblOcorrSelec = gridOcorr.Columns(3)
            
End If
        
'S T A T U S   D O   C T C / N F
        
If optCTC = True Then
    fraStatus = "S T A T U S   D O   C T C"
ElseIf optNf = True Then
    fraStatus = "S T A T U S   D A   N F"
End If

lblEntregueSN.ToolTipText = ""
If xtemocorr = "0" Then
    lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
    lblEntregueSN.Caption = "OCORR/Baixado"
ElseIf xtemocorr = "1" Then
    lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
    lblEntregueSN.Caption = "OK. ENTREGUE"
ElseIf xtemocorr = "2" Then
    lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
    lblEntregueSN.Caption = "OCORR/Pendente"
ElseIf xtemocorr = "N" Then
    If de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") >= xagora Then
        lblEntregueSN.ForeColor = &HC00000             'LABEL NA COR AZUL
        lblEntregueSN.Caption = "EM TRÂNSITO"
        lblEntregueSN.ToolTipText = "EM TRÂNSITO = Até a Previsão de Entrega"
    Else
        lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
        lblEntregueSN.Caption = "SEM POSIÇÃO há " & Trim$(Str(Val(xagora - de_informa.rsSel_Ctc_SAC.Fields("prev_entrega")))) & " dia(s)"
        lblEntregueSN.ToolTipText = "SEM POSIÇÃO = Após a Previsão de Entrega"
    End If
ElseIf xtemocorr = "C" Then
    lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
    lblEntregueSN.Caption = "CTC CANCELADO"
    lblEntregueSN.ToolTipText = "Cancelado em:" & de_informa.rsSel_Ctc_SAC.Fields("canc_data") & _
                                "  Usuário:" & de_informa.rsSel_Ctc_SAC.Fields("canc_usu") & _
                                "  Motivo:" & de_informa.rsSel_Ctc_SAC.Fields("canc_obs")
End If

If xtemocorr = "1" Then
        
    If optCTC = True Then
        If de_informa.rsSel_ConsEntregaCTC.State = 1 Then de_informa.rsSel_ConsEntregaCTC.Close
        de_informa.Sel_ConsEntregaCTC lblfilialctc
    
        'Pré Baixa (emails, relatórios, etc)
        lblDtPreBx = de_informa.rsSel_ConsEntregaCTC.Fields("dtbaixapre")
        lblHsPreBx = de_informa.rsSel_ConsEntregaCTC.Fields("hsbaixapre")
        txtRecPreBx = de_informa.rsSel_ConsEntregaCTC.Fields("recebpre")
        
        If de_informa.rsSel_ConsEntregaCTC.Fields("baixadofinal") = "S" Then
                
            'Baixa Física (pelo CTC Físico)
            lblDtBx = de_informa.rsSel_ConsEntregaCTC.Fields("dtbaixa")
            lblHsBx = de_informa.rsSel_ConsEntregaCTC.Fields("hsbaixa")
            txtRecBx = de_informa.rsSel_ConsEntregaCTC.Fields("receb")
            chkCanhoto.Enabled = False
            If de_informa.rsSel_ConsEntregaCTC.Fields("canhotonf") = "S" Then
                chkCanhoto.Value = 1
            Else
                chkCanhoto.Value = 0
            End If
        Else
            chkCanhoto.Value = 0
        End If
        
    ElseIf optNf = True Then
    
        If de_informa.rsSel_ConsEntregaNF.State = 1 Then de_informa.rsSel_ConsEntregaNF.Close
        de_informa.Sel_ConsEntregaNF lblfilialctc, txtNumNf
    
        'Pré Baixa (emails, relatórios, etc)
        lblDtPreBx = de_informa.rsSel_ConsEntregaNF.Fields("dtbaixapre")
        lblHsPreBx = de_informa.rsSel_ConsEntregaNF.Fields("hsbaixapre")
        txtRecPreBx = de_informa.rsSel_ConsEntregaNF.Fields("recebpre")
        
        If de_informa.rsSel_ConsEntregaNF.Fields("baixadofinal") = "S" Then
                
            'Baixa Física (pelo CTC Físico)
            lblDtBx = de_informa.rsSel_ConsEntregaNF.Fields("dtbaixa")
            lblHsBx = de_informa.rsSel_ConsEntregaNF.Fields("hsbaixa")
            txtRecBx = de_informa.rsSel_ConsEntregaNF.Fields("receb")
            chkCanhoto.Enabled = False
            If de_informa.rsSel_ConsEntregaNF.Fields("canhotonf") = "S" Then
                chkCanhoto.Value = 1
            Else
                chkCanhoto.Value = 0
            End If
        Else
            chkCanhoto.Value = 0
        End If
        
    End If
End If
     
'Caso o CTC esteja SEM POSICAO ou CANCELADO preenche com todas as NF o lstNfOcorr

If xtemocorr = "N" Or xtemocorr = "C" Then
    de_informa.rsSel_NFsdoCTC.MoveFirst
    Do Until de_informa.rsSel_NFsdoCTC.EOF
        lstNfOcorr.AddItem de_informa.rsSel_NFsdoCTC.Fields("numnf")
        de_informa.rsSel_NFsdoCTC.MoveNext
    Loop
    lblOcorrPorNF = "Todas as NFs do CTC"
    lblOcorrPorNFNao = ""
End If

'volta o foco para a Filial / NF
If optCTC = True Then
    txtFilial.SetFocus
ElseIf optNf = True Then
    txtNumNf.SetFocus
End If
        
'LOG DE USUÁRIO
de_informa.ins_LogUsuario "CONSULTA", xusuario, "CONSULTA NO POD - CONSULTA CTC: " & lblfilialctc

mskData.Enabled = True
mskData.BackColor = &HC0FFFF
mskHora.Enabled = True
mskHora.BackColor = &HC0FFFF
txtCodOcorr.Enabled = True
txtCodOcorr.BackColor = &HC0FFFF
txtObs_Ocorr.Enabled = True
txtObs_Ocorr.BackColor = &HC0FFFF

mskData.SetFocus

End Sub

Private Sub cmdExclPreBx_Click()
    If Mid$(xdireitos, 23, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        'Exclusão do Registro de Baixa, pois se o CTC tiver baixa física também,
        'a mesma é excluida pois não é possivel CTC baixado Físico sem Pré-Baixa.
        'Não é Possível Excluir Pré-Baixa e deixar a baixa física !
        If Len(lblDtBx) > 0 Then  'tem baixa física
            If MsgBox("Confirma Exclusão dos Dados de BAIXA (Pré-Baixa e Baixa Física) ? ", vbYesNo, "Atenção") = vbYes Then
            
                de_informa.cn_informa.BeginTrans
            
                de_informa.excl_BaixaPOD transctc(txtFilial, txtCtc)
                
                'exclui informação sobre canhoto
                de_informa.alt_ExclCanhotoNF transctc(txtFilial, txtCtc)
                
                'se houver ocorrência, tem_ocorr = '2', caso contrário tem_ocorr = 'N'
                If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                    de_informa.alt_temocorr_sn "2", transctc(txtFilial, txtCtc)
                Else
                    de_informa.alt_temocorr_sn "N", transctc(txtFilial, txtCtc)
                End If
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " PRÉ-BAIXA/FÍSICA"
                
                de_informa.cn_informa.CommitTrans
                
                cmdProcurar_Click
            End If
        Else  'é só Pré-Baixa
            If MsgBox("Confirma Exclusão dos Dados de PRÉ-BAIXA ? ", vbYesNo, "Atenção") = vbYes Then
            
                de_informa.cn_informa.BeginTrans
            
                de_informa.excl_BaixaPOD transctc(txtFilial, txtCtc)
                'se houver ocorrência, tem_ocorr = '2', caso contrário tem_ocorr = 'N'
                If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
                    de_informa.alt_temocorr_sn "2", transctc(txtFilial, txtCtc)
                Else
                    de_informa.alt_temocorr_sn "N", transctc(txtFilial, txtCtc)
                End If
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "EXCLUSÃO", xusuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " PRÉ-BAIXA"
                
                de_informa.cn_informa.CommitTrans
                
                cmdProcurar_Click
            End If
        End If
    End If
End Sub



Private Sub Form_Activate()
    txtFilial.Text = xultimofilial
    txtCtc.Text = xultimoctc
    txtNumNf.Text = xultimonf
    If optCTC = True Then
        txtFilial.SetFocus
    Else
        txtNumNf.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    'CONFIGURA OS OPTIONS, FRAMES E CHECKS
        optBaixaFinal.Enabled = False
        optPreBaixa.Enabled = False
        fraPreBaixa.Enabled = False
        fraBaixaFinal.Enabled = False
        gridOcorr.DataMember = ""
        gridOcorr.Refresh

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdiInforma.Toolbar1.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmPodNF = Nothing
End Sub

Private Sub gridOcorr_Click()
    If de_informa.rsSel_NfsComOcorr.State = 1 Then de_informa.rsSel_NfsComOcorr.Close
    de_informa.Sel_NfsComOcorr lblfilialctc, gridOcorr.Columns(2)
    'Set gridNFsComOcorr.DataSource = de_informa
    gridNFsComOcorr.DataMember = "sel_nfscomocorr"
    gridNFsComOcorr.Refresh
    lblOcorrSelec = gridOcorr.Columns(3)
    lblOcorrPorNF = "Do Total de " & Trim$(Str(de_informa.rsSel_NFsdoCTC.RecordCount)) & " NFs, " & _
    Trim$(Str(de_informa.rsSel_NfsComOcorr.RecordCount)) & " tiveram a Ocorrência Abaixo:"
End Sub

Private Sub mskData_Change()
    If Len(txtCodOcorr) >= 2 And IsDate(mskData) Then
        cmbGravar.Enabled = True
    Else
        cmbGravar.Enabled = False
    End If
End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskData_LostFocus()

    If mskData.Text <> "__/__/____" Then
        mskData.Text = century(mskData.Text)
        If Not IsDate(mskData.Text) Or Mid(mskData.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskData.SetFocus
            Exit Sub
        End If
        'tratamento acerto aws ---------------------------------------------
        If CDate(mskData.Text) < CDate(lblDtEmiss) Then
            MsgBox "Erro ! Data anterior à emissão.", vbCritical, "Erro"
            mskData.SetFocus
            Exit Sub
        End If
        '------------------------------------------------------------------
        If CDate(mskData.Text) > Date Then
            MsgBox "Erro ! Data posterior à hoje.", vbCritical, "Erro"
            mskData.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub mskHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskHora_LostFocus()
    If mskHora.Text <> "__:__" Then
        If Mid(mskHora.Text, 1, 2) > 23 Or Mid(mskHora.Text, 4, 2) > 59 Then
            MsgBox "Hora Inválida !", vbCritical, "Erro"
            mskHora.SetFocus
            Exit Sub
        End If
    Else
        mskHora.Text = "00:00"
    End If
End Sub

Private Sub optCTC_Click()
    txtFilial.Visible = True
    txtCtc.Visible = True
    txtNumNf.Visible = False
    txtFilial.SetFocus
End Sub

Private Sub optNf_Click()
    On Error Resume Next
    txtFilial.Visible = False
    txtCtc.Visible = False
    txtNumNf.Visible = True
    txtNumNf.SetFocus
End Sub

Private Sub txtCodOcorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCtc_Change()
    If Len(txtCtc.Text) >= 8 Then cmdProcurar.SetFocus
End Sub

Private Sub txtCTC_GotFocus()
   'RECEBER FOCO SELECIONADO
    txtCtc.SelStart = 0
    txtCtc.SelLength = 8
End Sub

Private Sub mskData_GotFocus()
    mskData.SelStart = 0
    mskData.SelLength = 10
End Sub

Private Sub txtCTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCtc_LostFocus()
    If txtCtc.Text <> "" Then
        If Not IsNumeric(txtCtc.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            txtCtc.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtfilial_Change()
    On Error Resume Next
    If Len(txtFilial.Text) >= 2 Then txtCtc.SetFocus
End Sub

Private Sub txtfilial_gotfocus()
   'RECEBER FOCO SELECIONADO
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub

Private Sub mskHora_GotFocus()
    mskHora.SelStart = 0
    mskHora.SelLength = 5
End Sub

Private Sub optBaixaFinal_Click()
    If lblDtPreBx.Caption = "" Then
       fraPreBaixa.Enabled = True
    Else
       fraPreBaixa.Enabled = False
    End If
       fraBaixaFinal.Enabled = True
       txtRecBx.BackColor = &HC0FFFF      'amarelo
       txtRecPreBx.BackColor = &H8000000E       'BRANCO
       txtRecBx.Enabled = True
       txtRecBx.SetFocus
       chkCanhoto.Value = 1
End Sub
Private Sub optPreBaixa_Click()
       fraPreBaixa.Enabled = True
       fraBaixaFinal.Enabled = False
       txtRecPreBx.BackColor = &HC0FFFF      'AMARELO
       txtRecBx.BackColor = &H8000000E       'BRANCO
       txtRecBx.Enabled = False
       txtRecPreBx.SetFocus
       chkCanhoto.Value = 0
End Sub
Private Sub txtCodOcorr_Change()
    If txtCodOcorr.Text = "01" Then
        optBaixaFinal.Enabled = True
        optPreBaixa.Enabled = True
        If optPreBaixa.Value = True Then
            optPreBaixa_Click
        ElseIf optBaixaFinal.Value = True Then
            optBaixaFinal_Click
        End If
    Else
        optBaixaFinal.Enabled = False
        optPreBaixa.Enabled = False
        fraPreBaixa.Enabled = False
        fraBaixaFinal.Enabled = False
        txtRecPreBx.BackColor = &H8000000E       'BRANCO
        txtRecBx.BackColor = &H8000000E       'BRANCO
    End If
    If Len(txtCodOcorr) >= 2 And IsDate(mskData) Then
        cmbGravar.Enabled = True
    Else
        cmbGravar.Enabled = False
    End If
End Sub
Private Sub txtCodOcorr_GotFocus()
    'RECEBER FOCO SELECIONADO
    txtCodOcorr.SelStart = 0
    txtCodOcorr.SelLength = 65000
End Sub
Private Sub txtCodOcorr_LostFocus()
    If txtCodOcorr.Text = "" Then
        Exit Sub
    Else
    'VERIFICA O CÓDIGO DE OCORRÊNCIA QUANDO DIGITADO E ATUALIZA A LABEL DE DESCRICAO DE OCORR
        If de_informa.rsSel_ConsCadOcor.State = 1 Then de_informa.rsSel_ConsCadOcor.Close
        de_informa.Sel_ConsCadOcor txtCodOcorr
        If de_informa.rsSel_ConsCadOcor.RecordCount > 0 Then
            lblDescOcorr.Caption = de_informa.rsSel_ConsCadOcor.Fields("descricao")
        Else
            MsgBox "Código de Ocorrência Inválido !", vbOKOnly + vbCritical, "Erro"
            txtCodOcorr.SetFocus
        End If
    End If
End Sub

Private Sub txtfilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        If txtFilial.Text = "" Then
            KeyAscii = 0
            optNf.Value = True
            optNf_Click
        Else
            KeyAscii = 0
            SendKeys "{TAB}"  'ENVIA UM TAB
        End If
    End If
End Sub

Private Sub txtFilial_LostFocus()
    If txtFilial.Text <> "" Then
        If Not IsNumeric(txtFilial.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            txtFilial.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtNumNf_GotFocus()
    txtNumNf.SelStart = 0
    txtNumNf.SelLength = 12
End Sub

Private Sub txtNumNf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        If txtNumNf.Text = "" Then
            KeyAscii = 0
            optCTC.Value = True
            optCTC_Click
        Else
            KeyAscii = 0
            SendKeys "{TAB}"  'ENVIA UM TAB
        End If
    End If
End Sub

Private Sub txtObs_Ocorr_Change()
    lblChar = zeros(Len(txtObs_Ocorr.Text), 3)
    If Val(lblChar) = 70 Then
        MsgBox "ATENÇÃO. Você atingiu a quantidade de 70 caracteres no texto desta observação. Os demais dados digitados a partir deste ponto, não serão enviados para os clientes que operam com Sistema de EDI. Para tanto, tente manter as informações mais importantes dentro da faixa de até 70 caracteres."
    End If
End Sub

Private Sub txtObs_Ocorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRecBx_Change()
    If optBaixaFinal.Value = True And lblDtPreBx = "" Then
        txtRecPreBx.Text = txtRecBx.Text
    End If
    If txtRecPreBx = "" And lblDtPreBx <> "" Then
        txtRecPreBx = "."
    End If
End Sub

Private Sub txtRecBx_GotFocus()
    txtRecBx.SelStart = 0
    txtRecBx.SelLength = 25
    chkCanhoto.Value = 1
End Sub

Private Sub txtRecBx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRecBx_LostFocus()
    txtRecBx.Text = UCase(txtRecBx.Text)
End Sub

Private Sub txtRecPreBx_GotFocus()
    txtRecPreBx.SelStart = 0
    txtRecPreBx.SelLength = 25
End Sub

Private Sub txtRecPreBx_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRecPreBx_LostFocus()
    txtRecPreBx.Text = UCase(txtRecPreBx.Text)
End Sub


