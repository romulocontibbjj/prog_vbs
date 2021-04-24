VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAIRPod 
   Caption         =   "Informação de Entregas e Ocorrências"
   ClientHeight    =   8190
   ClientLeft      =   -1425
   ClientTop       =   1170
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   12015
   WindowState     =   2  'Maximized
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
      Height          =   6855
      Left            =   6840
      TabIndex        =   47
      Top             =   960
      Width           =   5055
      Begin VB.TextBox txtCodOcorr 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Frame Frame6 
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
         Height          =   5535
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   4815
         Begin VB.Frame Frame10 
            Caption         =   "BAIXA FÍSICA"
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
            TabIndex        =   56
            Top             =   2880
            Width           =   4575
            Begin VB.CommandButton cmdExclBx 
               Caption         =   "EXCLUIR esta Baixa Física"
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               TabIndex        =   70
               Top             =   1680
               Width           =   3135
            End
            Begin VB.Frame fraBaixaFinal 
               Caption         =   "Dados da Baixa Física (Com o CTC Físico)"
               Height          =   1335
               Left            =   120
               TabIndex        =   64
               Top             =   240
               Width           =   4335
               Begin VB.CheckBox chkCanhoto 
                  Caption         =   "Possui o Canhoto da Nota Fiscal do Cliente ?"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   12
                  Top             =   1080
                  Width           =   3495
               End
               Begin VB.TextBox txtRecBx 
                  BackColor       =   &H8000000E&
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   25
                  TabIndex        =   11
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.Label lblHsBx 
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   3480
                  TabIndex        =   69
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
                  Caption         =   "Hora:"
                  Height          =   195
                  Left            =   3000
                  TabIndex        =   68
                  Top             =   360
                  Width           =   390
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  Caption         =   "Data Entrega:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   67
                  Top             =   360
                  Width           =   990
               End
               Begin VB.Label lblDtBx 
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   66
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  Caption         =   "Recebedor...:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   65
                  Top             =   600
                  Width           =   975
               End
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "PRÉ-BAIXA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   120
            TabIndex        =   55
            Top             =   960
            Width           =   4575
            Begin VB.CommandButton cmdExclPreBx 
               Caption         =   "EXCLUIR esta Pré-Baixa"
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               TabIndex        =   63
               Top             =   1440
               Width           =   3135
            End
            Begin VB.Frame fraPreBaixa 
               Caption         =   "Dados da Pré Baixa (Emails, Telefone, etc)"
               Height          =   1095
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Width           =   4335
               Begin VB.TextBox txtRecPreBx 
                  BackColor       =   &H8000000E&
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   25
                  TabIndex        =   10
                  Top             =   720
                  Width           =   2895
               End
               Begin VB.Label lblDtPreBx 
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H8000000D&
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   62
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Hora:"
                  Height          =   195
                  Left            =   3000
                  TabIndex        =   61
                  Top             =   360
                  Width           =   390
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Data Entrega:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   60
                  Top             =   360
                  Width           =   990
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "Recebedor...:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   59
                  Top             =   720
                  Width           =   975
               End
               Begin VB.Label lblHsPreBx 
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   3480
                  TabIndex        =   58
                  Top             =   360
                  Width           =   735
               End
            End
         End
         Begin VB.CheckBox chkObsEntr 
            Caption         =   "Comentários/ Observações de Entrega ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   5040
            Width           =   4575
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
            TabIndex        =   8
            ToolTipText     =   "Considerado com Data de Entrega na ausência de Baixa Física"
            Top             =   480
            Width           =   1215
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
            TabIndex        =   9
            ToolTipText     =   "Considerado Data de Entrega Independente da data de Pré-Baixa"
            Top             =   480
            Width           =   2295
         End
      End
      Begin MSMask.MaskEdBox mskHora 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
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
         TabIndex        =   5
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblDescOcorr 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   52
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cód. de Ocorrência:"
         Height          =   195
         Left            =   3120
         TabIndex        =   51
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hs:"
         Height          =   195
         Left            =   2160
         TabIndex        =   50
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   615
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdCalendario 
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   240
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Caption         =   "S T A T U S"
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
      TabIndex        =   40
      Top             =   120
      Width           =   2535
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
         TabIndex        =   41
         Top             =   360
         Width           =   2265
      End
   End
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
      Left            =   2760
      TabIndex        =   37
      Top             =   120
      Width           =   1800
      Begin VB.OptionButton optCTC 
         Caption         =   "Por Núm. de CTC"
         Height          =   255
         Left            =   105
         TabIndex        =   39
         Top             =   210
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optNf 
         Caption         =   "Por Núm de NF"
         Height          =   270
         Left            =   105
         TabIndex        =   38
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
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
      Height          =   2535
      Left            =   120
      TabIndex        =   30
      Top             =   960
      Width           =   6615
      Begin VB.CheckBox chkObsOcorr 
         Caption         =   "Comentários de Ocorrência ..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CommandButton cmdExclOcorr 
         Caption         =   "Excluir a Ocorrência Selecionada"
         Enabled         =   0   'False
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   3015
      End
      Begin MSDataGridLib.DataGrid gridOcorr 
         Height          =   1575
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2778
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         DataMember      =   "Sel_ConsOcorr2"
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
      Left            =   6840
      TabIndex        =   32
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdProcurar 
         Caption         =   "Procurar..."
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmbGravar 
         Caption         =   "Gravar a Ocorr."
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmbSair 
         Caption         =   "Canc/Sair"
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   945
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
      Height          =   4215
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   6615
      Begin VB.Frame Frame7 
         Caption         =   "Notas Fiscais"
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   6375
         Begin VB.Label lblNfs 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   6135
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Destino"
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   6375
         Begin VB.Label lblDestUf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5880
            TabIndex        =   27
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblDestCid 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3840
            TabIndex        =   26
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblDest 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Origem"
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   6375
         Begin VB.Label lblRemetUf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5880
            TabIndex        =   24
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblRemetCid 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3840
            TabIndex        =   23
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblRemet 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Label lblPrioridade 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NORMAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   46
         Top             =   240
         Width           =   960
      End
      Begin VB.Label lblModal 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4800
         TabIndex        =   36
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modal:"
         Height          =   195
         Left            =   4200
         TabIndex        =   35
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblHsEmiss 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3240
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblDtEmiss 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Left            =   2760
         TabIndex        =   19
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1020
      End
   End
   Begin VB.Frame fraProcura 
      Caption         =   "Núm. da  Filial e CTC"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtCtc 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   8
         TabIndex        =   2
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   435
      End
      Begin VB.TextBox txtNumNf 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1965
      End
   End
   Begin VB.Label lblcontroletela 
      AutoSize        =   -1  'True
      Caption         =   "normal"
      Height          =   195
      Left            =   6120
      TabIndex        =   45
      Top             =   7920
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblbxfinalSim 
      Height          =   255
      Left            =   9960
      TabIndex        =   44
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmAIRPod"
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

Private Sub chkObsEntr_Click()
    If chkObsEntr.Value = 1 Then
        frmObsOcorr.Caption = "Observação de Entrega"
        frmObsOcorr.Show 1
        chkObsEntr.Value = 0
        chkObsOcorr.Value = 0
        cmdProcurar_Click
    End If
End Sub

Private Sub chkObsOcorr_Click()
    If chkObsOcorr.Value = 1 Then
        frmObsOcorr.Show 1
        chkObsEntr.Value = 0
        chkObsOcorr.Value = 0
        cmdProcurar_Click
    End If
End Sub

Private Sub cmbGravar_Click()

frmPod.MousePointer = 11
'tratamento de acerto aws (data de emissão)------------------------------------------------

'If mskEmissaoNova.Text <> "__/__/____" Then
'    If IsDate(mskEmissaoNova.Text) Then
'        If CDate(mskEmissaoNova) <> CDate(lblDtEmiss) Then
'            'alterar data de emissão deste CTC no tb_ctc_esp e tb_ocorr
'            de_informa.Acerto_AltDataCTC CDate(mskEmissaoNova.Text), transctc(txtFilial, txtCtc)
'            de_informa.Acerto_AltDataOcorr CDate(mskEmissaoNova.Text), transctc(txtFilial, txtCtc)
'            lblDtEmiss.Caption = mskEmissaoNova.Text
'            lblEmissao2.Caption = mskEmissaoNova.Text
'            MsgBox "OK ! Data de Emissão Alterada !"
'        End If
'    End If
'End If

'------------------------------------------------------------------------------------------

    Dim xcanhoto As String
    Dim xcontnfscanhoto As Integer
    Dim xabonodias As Long
    xcanhoto = ""
'STORE PROCEDURES ocorr1 = Dados de Pré baixa  (ALT E INS)
                 'ocorr2 = Dados de Baixa Final  (ALT E INS)
                 'ocorr3 = Dados de Pré e Baixa Final (ambas) (ALT E INS)
                 'ocorr4 = Dados de Ocorrência  (INS)
                 
'INDICAÇÕES DO tem_ocorr   (N, 0, 1 ou 2)
                 'N = indica que não há ocorrência nem baixa OU em Trânsito
                 '0 = indica processo com ocorrência, mas fechado
                 '1 = indica ctc já entregue
                 '2 = indica ctc com ocorrência, mas NÃO fechado (pendente)
                 
'TRATAMENTO DE  B A I X A S

    If Not IsDate(mskData.Text) Then
        frmPod.MousePointer = 0
        MsgBox "Data Inválida !", vbCritical, "Erro"
        mskData.SetFocus
        Exit Sub
    End If
    
    
    'tratamento acerto aws ----------------------------------------------------------------
   ' If mskEmissaoNova.Text <> "__/__/____" Then
   '     If CDate(mskData.Text) < CDate(mskEmissaoNova) Then
   '         MsgBox "Erro ! Data da Ocorrência/Entrega anterior à emissão.", vbCritical, "Erro"
   '         mskData.SetFocus
   '         Exit Sub
   '     End If
   ' End If
    '--------------------------------------------------------------------------------------


    If txtCodOcorr.Text = "01" Then  'se for "01" (entrega realizada/baixa)
        'verifica se campos estão digitados
        If optBaixaFinal = False And optPreBaixa = False Then
            frmPod.MousePointer = 0
            MsgBox "Escolha a forma de Baixa: Pré-Baixa ou Baixa Física (Final) !"
            Exit Sub
        End If
        If mskData.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Data", vbOKOnly + vbCritical, "ERRO"
            mskData.SetFocus
            Exit Sub
        ElseIf mskHora.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Hora", vbOKOnly + vbCritical, "ERRO"
            mskHora.SetFocus
            Exit Sub
        ElseIf txtCodOcorr.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Cod. Ocorrência", vbOKOnly + vbCritical, "ERRO"
            txtCodOcorr.SetFocus
            Exit Sub
        ElseIf optBaixaFinal = True Then
            If txtRecBx.Text = "" Then
                frmPod.MousePointer = 0
                MsgBox "Dados Inválidos ! Campo: Recebedor (Baixa Final)", vbOKOnly + vbCritical, "ERRO"
                txtRecBx.SetFocus
                Exit Sub
            End If
        ElseIf optPreBaixa = True Then
            If txtRecPreBx.Text = "" Then
                frmPod.MousePointer = 0
                MsgBox "Dados Inválidos ! Campo: Recebedor (Pré Baixa)", vbOKOnly + vbCritical, "ERRO"
                txtRecPreBx.SetFocus
                Exit Sub
            End If
        End If
        'verifica se o CTC já tem ocorrência fechada cadastrada. Caso tenha não possibilita a baixa
        If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "0" Then
            frmPod.MousePointer = 0
            MsgBox "Este CTC está Baixado, indicando que esta entrega não ocorreria. Caso deseje baixar como ENTREGA, você deve primeiro excluir a Ocorrência  C T C   B A I X A D O", vbOKCancel
            txtCodOcorr.SetFocus
            Exit Sub
        'verifica se o ctc já possui ocorrência e se a data que se quer baixar não é menor que a data de alguma ocorrência
        ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
            de_informa.rsSel_ConsOcorr2.MoveFirst
            Do Until de_informa.rsSel_ConsOcorr2.EOF
                If CDate(mskData.Text) < de_informa.rsSel_ConsOcorr2.Fields("data") Then
                    frmPod.MousePointer = 0
                    If MsgBox("Você está tentando baixar um CTC com data Menor que uma Ocorrência Cadastrada. Você tem certeza que quer baixar este CTC como entrega nesta data ?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then
                        mskData.SetFocus
                        Exit Sub
                    Else
                        Exit Do
                    End If
                End If
                de_informa.rsSel_ConsOcorr2.MoveNext
            Loop
        End If
            
'início do processo de baixa.

        If optPreBaixa.Value = True Then         'se for uma pré baixa
            'procura se o CTC já está baixado pré
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr transctc(txtFilial.Text, txtCtc.Text), "01"
            If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
               'Este CTC já Contém Pré Baixa.
                frmPod.MousePointer = 0
                MsgBox "Este CTC já está baixado com Pré-Baixa. Caso esteja tentando alterar a data que já está cadastrada, você deve antes excluir esta Pré-Baixa e lançar novamente com a data correta. Exclusão de Entrega/Ocorrências só pode ser realizada por usuário que possui este direito de acesso. Esta informação não será gravada no sistema.", vbOKOnly + vbExclamation
                mskData.SetFocus
                Exit Sub
                
                    'frmPod.MousePointer = 11
                    'de_informa.cn_informa.BeginTrans
                    'If de_informa.rsSel_ConsOcorr.Fields("baixadofinal") = "S" Then
                    '   de_informa.alt_ocorr1ow transctc(txtFilial.Text, txtCtc.Text), mskData.Text, mskHora.Text, RTrim(txtRecPreBx.Text), xusuario, CVar(Date) & " " & CVar(Time())
                    '   de_informa.alt_temocorr_sn "1", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                    'Else
                    '   'atual_sitla=S  =>  atualizar o sistema SITLA
                    '   de_informa.alt_ocorr1 transctc(txtFilial.Text, txtCtc.Text), mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, RTrim(txtRecPreBx.Text), xusuario, CVar(Date) & " " & CVar(Time()), "S", Date
                    '   de_informa.Alt_AtClienteNFBranco transctc(txtFilial.Text, txtCtc.Text)
                    '   de_informa.alt_temocorr_sn "1", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                    'End If
                    
                    'LOG DE USUÁRIO
                    'de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " PRÉ-BAIXA"
                    'de_informa.cn_informa.CommitTrans

            Else 'SE NÃO HOUVER NENHUMA BAIXA, INCLUI ...  (INSERT "01")
                de_informa.cn_informa.BeginTrans
                de_informa.ins_ocorr1 transctc(txtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, RTrim(txtRecPreBx.Text), xUsuario, DataHora("datahora"), "S", DataHora("data")
                de_informa.alt_temocorr_sn "1", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                de_informa.Alt_AtClienteNFBranco transctc(txtFilial.Text, txtCtc.Text)
                    
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "INCLUSÃO", xUsuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " PRÉ-BAIXA"
                de_informa.cn_informa.CommitTrans
                
                'atualiza os prazos
                frmAtualPrazos.lblFilialctc = transctc(txtFilial.Text, txtCtc.Text)
                frmAtualPrazos.Show 1
            End If
        ElseIf optBaixaFinal.Value = True Then   'se for uma baixa final ou ambas
            If txtRecBx.Text = "" Then
                frmPod.MousePointer = 0
               MsgBox "Dados Inválidos ! Campo: Recebedor", vbOKOnly + vbCritical, "ERRO"
               txtRecBx.SetFocus
               Exit Sub
            End If
            If chkCanhoto.Value = 1 Then
                frmContrCanhotos.lstPresentes.Clear
                frmContrCanhotos.lstFaltantes.Clear
                xcanhoto = "S"
                If chkCanhoto.Enabled = True Then
                    If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
                    de_informa.Sel_NFsdoCTC transctc(txtFilial.Text, txtCtc.Text)
                    If de_informa.rsSel_NFsdoCTC.RecordCount > 0 Then
                        Do Until de_informa.rsSel_NFsdoCTC.EOF
                            frmContrCanhotos.lstPresentes.AddItem de_informa.rsSel_NFsdoCTC.Fields("numnf")
                            de_informa.rsSel_NFsdoCTC.MoveNext
                        Loop
                        frmContrCanhotos.lblFilialctc = transctc(txtFilial.Text, txtCtc.Text)
                        frmContrCanhotos.fraPresentes.Caption = frmContrCanhotos.lstPresentes.ListCount & " Canhotos"
                        frmContrCanhotos.Show 1
                        If lblcontroletela.Caption = "cancelar" Then
                            lblcontroletela.Caption = "normal"
                            Unload frmContrCanhotos
                            Me.MousePointer = 0
                            cmdProcurar_Click
                            Exit Sub
                        End If
                    Else
                        xcanhoto = "N"  'pois não há NFS
                    End If
                End If
            Else
                xcanhoto = "N"
            End If
            'procura se o CTC já está baixado final
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr transctc(txtFilial.Text, txtCtc.Text), "01"
            If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
                If de_informa.rsSel_ConsOcorr.Fields("baixadofinal") = "S" Then
                    'Este CTC já está baixado (ambos ou final)
                    frmPod.MousePointer = 0
                    MsgBox "Este CTC já está baixado com Baixa-Física. Caso esteja tentando alterar a data que já está cadastrada, você deve antes excluir esta Baixa-Física e lançar novamente com a data correta. Exclusão de Entrega/Ocorrências só pode ser realizada por usuário que possui este direito de acesso. Esta informação não será gravada no sistema.", vbOKOnly + vbExclamation
                    Exit Sub
                        
                        'frmPod.MousePointer = 11
                        'inicia a transação
                        'de_informa.cn_informa.BeginTrans
                        
                        'atualiza com os dados de baixa física
                        'de_informa.alt_ocorr2 transctc(txtfilial.Text, txtCTC.Text), mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, RTrim(txtRecBx.Text), xusuario, CVar(Date) & " " & CVar(Time()), "S", Date, xcanhoto
                        'de_informa.alt_temocorr_sn "1", transctc(txtfilial.Text, txtCTC.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                        
                        'atualiza as NFs que contém o canhoto
                        
                        'For xcontnfscanhoto = 1 To frmContrCanhotos.lstPresentes.ListCount
                        '    de_informa.Alt_CanhotoNFSN "S", transctc(txtfilial.Text, txtCTC.Text), frmContrCanhotos.lstPresentes.List(xcontnfscanhoto - 1)
                        'Next
                        
                        'atualiza as NFs que NÃO contém o canhoto
                        
                        'For xcontnfscanhoto = 1 To frmContrCanhotos.lstFaltantes.ListCount
                        '    de_informa.Alt_CanhotoNFSN "N", transctc(txtfilial.Text, txtCTC.Text), frmContrCanhotos.lstFaltantes.List(xcontnfscanhoto - 1)
                        'Next
                        
                        'atualiza status de envio de informação para o cliente
                        'de_informa.Alt_AtClienteNFBranco transctc(txtfilial.Text, txtCTC.Text)
                        
                        'lblbxfinalSim = "SIM" 'identifica label invisível como SIM para controle se executa pergunta de relatório do Protocolo para Setor de Arquivo
                        
                        'LOG DE USUÁRIO
                        'de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "POD/OCORR - CTC:" & transctc(txtfilial.Text, txtCTC.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " BAIXA FINAL (FÍSICA)"
                        
                        'finaliza transação
                        'de_informa.cn_informa.CommitTrans
                        
                Else  'SE NÃO HOUVER BAIXA FINAL, INCLUI NO REGISTRO DE BAIXA (UPDATE "01" BAIXA FINAL)
                
                    'Se Não há baixa Física é porque há baixa final, pois o mesmo já está baixado !
                    
                    If CDate(mskData) <> CDate(lblDtPreBx) Then
                    
                        Me.MousePointer = 0
                        
                        If MsgBox("ATENÇÃO ! A Data desta Baixa Física que você está querendo cadastrar é diferente da data da Pré-Baixa que já está cadastrada para este CTC/NF." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Você tem certeza que esta informação está correta e deseja realmente gravar esta Baixa-Física com a data diferente da Pré-Baixa ?", vbYesNo + vbQuestion + vbCritical, "ATENÇÃO") = vbNo Then
                            MsgBox "OK ! Operação Cancelada. Esta informação não será gravada no sistema."
                            txtRecBx.Text = ""
                            txtCodOcorr.Text = ""
                            mskData.SetFocus
                            Exit Sub
                        End If
                        
                    End If
                        
                    Me.MousePointer = 11
                    'inicia a transação
                    de_informa.cn_informa.BeginTrans
                                                
                    'atualiza os dados de entrega
                    de_informa.alt_ocorr2 transctc(txtFilial.Text, txtCtc.Text), CDate(lblDtPreBx), lblHsPreBx, mskData.Text, mskHora.Text, RTrim(txtRecBx.Text), xUsuario, DataHora("datahora"), "N", DataHora("data"), xcanhoto
                    de_informa.alt_temocorr_sn "1", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                    
                    'atualiza as NFs que contém o canhoto
                        
                    For xcontnfscanhoto = 1 To frmContrCanhotos.lstPresentes.ListCount
                        de_informa.Alt_CanhotoNFSN "S", transctc(txtFilial.Text, txtCtc.Text), frmContrCanhotos.lstPresentes.List(xcontnfscanhoto - 1)
                    Next
                        
                    'atualiza as NFs que NÃO contém o canhoto
                        
                    For xcontnfscanhoto = 1 To frmContrCanhotos.lstFaltantes.ListCount
                        de_informa.Alt_CanhotoNFSN "N", transctc(txtFilial.Text, txtCtc.Text), frmContrCanhotos.lstFaltantes.List(xcontnfscanhoto - 1)
                    Next
                    
                    Unload frmContrCanhotos
                    
                    'atualiza o campo de informação para o cliente
                    de_informa.Alt_AtClienteNFBranco transctc(txtFilial.Text, txtCtc.Text)
                    
                    lblbxfinalSim = "SIM" 'identifica label invisível como SIM para controle se executa pergunta de relatório do Protocolo para Setor de Arquivo
                        
                    'LOG DE USUÁRIO
                    de_informa.ins_LogUsuario "INCLUSAO", xUsuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " BAIXA FINAL/FÍSICA (JÁ HAVIA PRÉ-BAIXA)"
                    
                    'finaliza Transação
                    de_informa.cn_informa.CommitTrans
                    
                    'atualiza os prazos
                    frmAtualPrazos.lblFilialctc = transctc(txtFilial.Text, txtCtc.Text)
                    frmAtualPrazos.Show 1
                            
                End If
                
            Else  'SE NÃO HOUVER NENHUMA BAIXA, INCLUI AMBOS (PRE E FINAL)
            
                'inicia a transação
                de_informa.cn_informa.BeginTrans
                
                'atualiza os dados de entrega
                de_informa.ins_ocorr3 transctc(txtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
                txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, RTrim(txtRecPreBx), mskData.Text, mskHora.Text, RTrim(txtRecBx.Text), xUsuario, DataHora("datahora"), "S", DataHora("data"), xcanhoto 'insere aS baixas ambas
                
                'atualiza o status do CTC
                de_informa.alt_temocorr_sn "1", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr = 1
                
                'atualiza as NFs que contém o canhoto
                        
                For xcontnfscanhoto = 1 To frmContrCanhotos.lstPresentes.ListCount
                    de_informa.Alt_CanhotoNFSN "S", transctc(txtFilial.Text, txtCtc.Text), frmContrCanhotos.lstPresentes.List(xcontnfscanhoto - 1)
                Next
                        
                'atualiza as NFs que NÃO contém o canhoto
                        
                For xcontnfscanhoto = 1 To frmContrCanhotos.lstFaltantes.ListCount
                    de_informa.Alt_CanhotoNFSN "N", transctc(txtFilial.Text, txtCtc.Text), frmContrCanhotos.lstFaltantes.List(xcontnfscanhoto - 1)
                Next
                
                Unload frmContrCanhotos
                
                lblbxfinalSim = "SIM" 'identifica label invisível como SIM para controle se executa pergunta de relatório do Protocolo para Setor de Arquivo
                
                'atualiza campo para informação para o cliente
                de_informa.Alt_AtClienteNFBranco transctc(txtFilial.Text, txtCtc.Text)
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "INCLUSAO", xUsuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr & " PRÉ-BAIXA + BAIXA FINAL/FÍSICA"
                
                'finaliza transação
                de_informa.cn_informa.CommitTrans
                
                'atualiza os prazos
                frmAtualPrazos.lblFilialctc = transctc(txtFilial.Text, txtCtc.Text)
                frmAtualPrazos.Show 1
                
            End If
        End If
        
'TRATAMENTO DE  O C O R R Ê N C I A S
        
    Else   'se nao for baixa (ocorr # 01) então é somente ocorrência
        'verifica se campos estão digitados
        If mskData.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Data", vbOKOnly + vbCritical, "ERRO"
            mskData.SetFocus
            Exit Sub
        ElseIf mskHora.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Hora", vbOKOnly + vbCritical, "ERRO"
            mskHora.SetFocus
            Exit Sub
        ElseIf txtCodOcorr.Text = "" Then
            frmPod.MousePointer = 0
            MsgBox "Dados Inválidos ! Campo: Cod. Ocorrência", vbOKOnly + vbCritical, "ERRO"
            txtCodOcorr.SetFocus
            Exit Sub
        End If
        
        If txtCodOcorr.Text = "00" Then   'se for ocorr 00
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "1" Then
                frmPod.MousePointer = 0
                MsgBox "CTC já Baixado como Entregue. Não é Possível informar Ocorrência  C T C   B A I X A D O"
                txtCodOcorr.SetFocus
                Exit Sub
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "0" Then
                frmPod.MousePointer = 0
                MsgBox "CTC já possui Ocorrência  C T C   B A I X A D O"
                txtCodOcorr.SetFocus
                Exit Sub
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "N" Then
                frmPod.MousePointer = 0
                MsgBox "Você só pode informar Ocorrência  C T C   B A I X A D O, se o CTC já tiver alguma ocorrência que a explique o motivo."
                txtCodOcorr.SetFocus
                Exit Sub
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
                de_informa.rsSel_ConsOcorr2.MoveFirst
                Do Until de_informa.rsSel_ConsOcorr2.EOF
                   If CDate(mskData.Text) < de_informa.rsSel_ConsOcorr2.Fields("data") Then
                       MsgBox "A Data da Baixa Deve ser maior ou igual a última ocorrência cadastrada.", vbOKOnly, "Erro"
                       mskData.SetFocus
                       Exit Sub
                    End If
                    de_informa.rsSel_ConsOcorr2.MoveNext
                Loop
            End If
        Else   'se não for é ocorrência normal
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "1" Then
                If IsDate(lblDtBx.Caption) Then
                    If CDate(mskData.Text) > CDate(lblDtBx.Caption) Then
                        frmPod.MousePointer = 0
                        If MsgBox("Você está tentando incluir uma Ocorrência com data Posterior à sua Data de Entrega. Você tem certeza que deseja informar esta ocorrência com esta data ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                            mskData.SetFocus
                            Exit Sub
                        End If
                    End If
                ElseIf IsDate(lblDtPreBx.Caption) Then
                    If CDate(mskData.Text) > CDate(lblDtPreBx.Caption) Then
                        frmPod.MousePointer = 0
                        If MsgBox("Você está tentando incluir uma Ocorrência com data Posterior à Data de Entrega. Você tem certeza que deseja informar esta ocorrência com esta data ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                            mskData.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            ElseIf de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "0" Then
                de_informa.rsSel_ConsOcorr2.MoveFirst
                Do Until de_informa.rsSel_ConsOcorr2.EOF
                    If de_informa.rsSel_ConsOcorr2.Fields("cod_ocorr") = "00" Then
                        If CDate(mskData.Text) > de_informa.rsSel_ConsOcorr2.Fields("data") Then
                            frmPod.MousePointer = 0
                            If MsgBox("Você está tentando lançar uma ocorrência com data posterior a uma ocorrência  C T C   B A I X A D O", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                                mskData.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    de_informa.rsSel_ConsOcorr2.MoveNext
                Loop
            End If
        End If
        
        'ATUALIZA BD COM OS DADOS DA OCORRÊNCIA
        
        de_informa.cn_informa.BeginTrans
        
        If txtCodOcorr.Text = "00" Then   'se for ocorr 00
            de_informa.ins_ocorr4cod00 transctc(txtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
            txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, xUsuario, DataHora("datahora")
            '0 = IDENTIFICA COMO CTC COM OCORRÊNCIA FECHADA
            'If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "2" Then
            de_informa.alt_temocorr_sn "0", transctc(txtFilial.Text, txtCtc.Text)   'atualiza arquivo de CTC com tem_ocorr = 0
            'End If
        Else  'se for outro tipo de ocorrência
            de_informa.ins_ocorr4 transctc(txtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), _
            txtCodOcorr.Text, lblDescOcorr.Caption, mskData.Text, mskHora.Text, xUsuario, DataHora("datahora")
            '2 = IDENTIFICA COMO CTC COM OCORRÊNCIAS PENDENTE
            de_informa.Alt_AtClienteNFBranco transctc(txtFilial.Text, txtCtc.Text)
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") <> "1" And de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") <> "0" Then
                de_informa.alt_temocorr_sn "2", transctc(txtFilial.Text, txtCtc.Text)   'atualiza arquivo de CTC com tem_ocorr = 2
                If txtCodOcorr = "39" Or txtCodOcorr = "84" Then  'pre-baixa automática por ser CTC/NF Retido para COnferência
                    If MsgBox("Você está lançando uma ocorrência de retenção de Doctos. para conferência, onde provavelmente a entrega foi realizada nesta data." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "O Informa pode incluir automaticamente uma Pré-Baixa para este CTC nesta data. Você Confirma ?", vbYesNo, "Pré-Baixa Automática") = vbYes Then
                        de_informa.ins_ocorr1 transctc(txtFilial.Text, txtCtc.Text), CDate(frmPod.lblDtEmiss), de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), "01", "ENTREGA REALIZADA", _
                        mskData.Text, mskHora.Text, mskData.Text, mskHora.Text, ".", "AUTO-PREBX", DataHora("datahora"), "S", DataHora("data")
                        de_informa.alt_temocorr_sn "1", transctc(txtFilial.Text, txtCtc.Text)
                        MsgBox "Pré-Baixa Automática Gravada !"
                    End If
                End If
            End If
            If de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") = "1" Then
                If txtCodOcorr = "26" Or txtCodOcorr = "85" Then  'abono automático para atraso
                    If de_informa.rsSel_CTCEntrega.State = 1 Then de_informa.rsSel_CTCEntrega.Close
                    de_informa.Sel_CTCEntrega transctc(txtFilial.Text, txtCtc.Text)
                    If de_informa.rsSel_CTCEntrega.RecordCount > 0 Then
                        If de_informa.rsSel_CTCEntrega.Fields("diasuteis") - _
                           de_informa.rsSel_CTCEntrega.Fields("abonodias") > _
                           de_informa.rsSel_CTCEntrega.Fields("prazoentr") Then
                           'está em atraso, lançar abono automático
                           xabonodias = de_informa.rsSel_CTCEntrega.Fields("diasuteis") - de_informa.rsSel_CTCEntrega.Fields("prazoentr")
                           de_informa.Alt_AbonoAtraso xabonodias, "AUTOMATIC", DataHora("DATAHORA"), "Abono Automático Devido Ocorrência", transctc(txtFilial.Text, txtCtc.Text)
                        End If
                    End If
                End If
            End If
        End If
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "INCLUSAO", xUsuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & txtCodOcorr & "-" & lblDescOcorr
        
        de_informa.cn_informa.CommitTrans
        
    End If
    
    mskData.Mask = ""
    mskData.Text = ""
    mskData.Mask = "##/##/####"
    'mskData.Enabled = False
    'mskData.BackColor = &H8000000E   'branco
    mskHora.Mask = ""
    mskHora.Text = ""
    mskHora.Mask = "##:##"
    'mskHora.Enabled = False
    'mskHora.BackColor = &H8000000E   'branco
    

    txtCodOcorr = ""
    'txtCodOcorr.Enabled = False
    'txtCodOcorr.BackColor = &H8000000E   'branco
    lblDescOcorr.Caption = ""
    txtRecPreBx.BackColor = &H8000000E   'branco
    txtRecBx.BackColor = &H8000000E   'branco
    cmbGravar.Enabled = False
    frmPod.MousePointer = 0
    cmdProcurar_Click
    MsgBox "OK ! OCORRÊNCIA REGISTRADA.", vbOKOnly + vbExclamation
    txtFilial.SetFocus
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
        'frmAtualPrazos.Show 1
        'If lblbxfinalSim = "SIM" Then
            'If MsgBox("Deseja Imprimir o Relatório de CTCs Físicos Baixados, para Envio dos Documentos para o Arquivo ? (PROTOCOLO)", vbQuestion + vbYesNo, "Confirmação de Relatório") = vbYes Then
            '    mdiInforma.StatusBar1.Panels.Item(2).Text = "AGUARDE IMPRESSÃO DO RELATÓRIO ..."
            '    DoEvents
            '    Call rel_arquivo
            '    mdiInforma.StatusBar1.Panels.Item(2).Text = ""
            '    DoEvents
            'End If
        'End If
        Set frmPod = Nothing
        Unload Me
    End If
End Sub

Private Sub cmdComentario_Click()
    frmObsOcorr.Show 1
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
            de_informa.ins_LogUsuario "EXCLUSÃO", xUsuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " BAIXA FÍSICA"
            
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
            xcodocorr = GridOcorr.Columns(2)
            
            de_informa.cn_informa.BeginTrans
            
            de_informa.excl_ocorr GridOcorr.Columns(9)
            If xcodocorr = "00" Then 'se for "00" altera o temocorr para 02
                de_informa.alt_temocorr_sn "2", transctc(txtFilial.Text, txtCtc.Text)  'atualiza arquivo de CTC com tem_ocorr
            End If
            
            'LOG DE USUÁRIO
            de_informa.ins_LogUsuario "EXCLUSÃO", xUsuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " OCORR:" & GridOcorr.Columns(2) & "-" & GridOcorr.Columns(3)
            
            'SE HOUVER HOUVER ENTREGA, ABONODIAS = 0 DEVIDO EXCLUSÃO DE OCORRÊNCIA
            
            If de_informa.rsSel_CTCEntrega.State = 1 Then de_informa.rsSel_CTCEntrega.Close
            de_informa.Sel_CTCEntrega transctc(txtFilial.Text, txtCtc.Text)
            
            If de_informa.rsSel_CTCEntrega.RecordCount > 0 Then
                de_informa.Alt_ExclAbono transctc(txtFilial.Text, txtCtc.Text)
                MsgBox "Caso Este CTC tenha algum Abono de Atraso, este abono foi excluido."
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "EXCLUSÃO", xUsuario, "ABONO ATRASO:" & transctc(txtFilial.Text, txtCtc.Text) & " DEVIDO EXCLUSÃO DE OCORRENCIA."
            End If
            
            'busca as ocorrências e atualiza o grid de ocorrências
            
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 transctc(txtFilial, txtCtc), "01"
            Set GridOcorr.DataSource = de_informa
            GridOcorr.DataMember = "Sel_ConsOcorr2"
            GridOcorr.Refresh
            
         'verifica se é a última ocorrência baixada e se é ocorr "00"
        'se for exclui ela também (pois o processo não está finalizado) e atualiza o grid novamente
 
            If de_informa.rsSel_ConsOcorr2.RecordCount = 1 Then
                If de_informa.rsSel_ConsOcorr2.Fields("cod_ocorr") = "00" Then
                    de_informa.excl_ocorr GridOcorr.Columns(9)
                    If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                    de_informa.Sel_ConsOcorr2 transctc(txtFilial, txtCtc), "01"
                    Set GridOcorr.DataSource = de_informa
                    GridOcorr.DataMember = "Sel_ConsOcorr2"
                    GridOcorr.Refresh
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

Private Sub CmdObsEntr_Click()
    frmObsOcorr.Show 1
End Sub


Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    
End Sub
Private Sub cmdProcurar_Click()
    If optNf.Value = True Then  'Se a procura for por NF
        If txtNumNf.Text = "" Then
            MsgBox "Número de Nota Fiscal Inválida !", vbCritical, "Erro"
        End If
            If de_informa.rsSel_NF_SAC.State = 1 Then de_informa.rsSel_NF_SAC.Close
            de_informa.Sel_NF_SAC Val(txtNumNf)   'Procura a NF na Tabela
            If de_informa.rsSel_NF_SAC.RecordCount = 0 Then
                MsgBox "Número de NF Não Encontrado !", vbCritical + vbOKOnly, "Erro"
                txtNumNf.SetFocus
                Exit Sub
            ElseIf de_informa.rsSel_NF_SAC.RecordCount > 1 Then
                frmDuplNF.Caption = "POD - Número de NFs Duplicadas"
                DoEvents
                frmDuplNF.Show 1  'direciona para o form que trata casos de NF duplicadas
                Exit Sub
            Else  'Caso seja encontrada somente uma NF
                optCTC_Click
                txtFilial.Text = Mid(de_informa.rsSel_NF_SAC.Fields("filialctc"), 1, 2)
                txtCtc.Text = Mid(de_informa.rsSel_NF_SAC.Fields("filialctc"), 3, 8) 'Busca a Filial e o CTC com base na NF
            End If
    End If
    optCTC.Value = True
        Dim xtemocorr As String
        If txtFilial.Text = "" Or txtCtc.Text = "" Then
            MsgBox "Filial / CTC Inválidos !", vbCritical, "Erro"
            Exit Sub
        End If
        If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
        de_informa.Sel_Ctc_SAC transctc(txtFilial, txtCtc)  'Procura na Tabela a Filial/CTC
        If de_informa.rsSel_Ctc_SAC.RecordCount = 0 Then
            MsgBox "Número de Filial/CTC Não Encontrados !", vbCritical + vbOKOnly, "Erro"
            txtFilial.SetFocus
            Exit Sub
        End If
'REGISTRA VARIÁVEIS GLOBAIS DE FILIAL E CTC PARA UTILIZAÇÃO EM OUTROS FORMS
        xultimofilial = txtFilial.Text
        xultimoctc = txtCtc.Text
'ATUALIZA DADOS DO CTC NO FORM
        lblDtEmiss.Caption = de_informa.rsSel_Ctc_SAC.Fields("data")
        
        
        'tratamento de data de emissão (acerto aws)--------------------------------------
        'mskEmissaoNova.Text = de_informa.rsSel_Ctc_SAC.Fields("data")
        'lblEmissao2 = de_informa.rsSel_Ctc_SAC.Fields("data")
        '--------------------------------------------------------------------------------
        
        
        lblHsEmiss.Caption = de_informa.rsSel_Ctc_SAC.Fields("hora")
        lblRemet.Caption = de_informa.rsSel_Ctc_SAC.Fields("remet_nome")
        lblRemetCid.Caption = de_informa.rsSel_Ctc_SAC.Fields("cidade_orig")
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli de_informa.rsSel_Ctc_SAC.Fields("remet_cgc")
        lblRemetUf = de_informa.rsSel_ConsCadCli.Fields("uf")
        lblDest.Caption = de_informa.rsSel_Ctc_SAC.Fields("dest_nome")
        lblDestCid.Caption = de_informa.rsSel_Ctc_SAC.Fields("cidade_dest")
        lblDestUf.Caption = de_informa.rsSel_Ctc_SAC.Fields("uf_dest")
        lblNfs.Caption = de_informa.rsSel_Ctc_SAC.Fields("nfs")
        lblModal.Caption = de_informa.rsSel_Ctc_SAC.Fields("modal")
        
        If de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "URGÊNCIA" Or _
            de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "PRIORIDADE" Then
            lblPrioridade.ForeColor = &HC0&
        Else
            lblPrioridade.ForeColor = &H80000012
        End If
        lblPrioridade = de_informa.rsSel_Ctc_SAC.Fields("prioridade")
        
        'LIMPA OS DADOS DO FORM
        mskData.Mask = ""
        mskData.Text = ""
        mskData.Mask = "##/##/####"
        mskHora.Mask = ""
        mskHora.Text = ""
        mskHora.Mask = "##:##"
        txtCodOcorr = ""
        lblDescOcorr.Caption = ""
        xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr") 'verifica se tem Ocorrência
        cmbGravar.Enabled = True
        mskData.Enabled = True
        mskHora.Enabled = True
        txtCodOcorr.Enabled = True
        
        lblEntregueSN.ToolTipText = ""
        If xtemocorr = "0" Then
           lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
           lblEntregueSN.Caption = "OCORR/Baixado"
        ElseIf xtemocorr = "1" Then
           lblEntregueSN.ForeColor = &HC00000   'LABEL NA COR AZUL
           lblEntregueSN.Caption = "OK. ENTREGUE"
        ElseIf xtemocorr = "2" Then
           lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
           lblEntregueSN.Caption = "OCORR/Pendente"
        ElseIf xtemocorr = "N" Then
            If de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") >= DataHora("data") Then
                lblEntregueSN.ForeColor = &HC00000             'LABEL NA COR AZUL
                lblEntregueSN.Caption = "EM TRÂNSITO"
                lblEntregueSN.ToolTipText = "EM TRÂNSITO = Até a Previsão de Entrega"
            Else
                lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
                lblEntregueSN.Caption = "SEM POSIÇÃO"
                lblEntregueSN.ToolTipText = "SEM POSIÇÃO = Após a Previsão de Entrega"
            End If
        ElseIf xtemocorr = "C" Then
            cmbGravar.Enabled = False
            mskData.Enabled = False
            mskHora.Enabled = False
            txtCodOcorr.Enabled = False
            lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
            lblEntregueSN.Caption = "CTC CANCELADO"
            lblEntregueSN.ToolTipText = "Cancelado em:" & de_informa.rsSel_Ctc_SAC.Fields("canc_data") & _
                                        "  Usuário:" & de_informa.rsSel_Ctc_SAC.Fields("canc_usu") & _
                                        "  Motivo:" & de_informa.rsSel_Ctc_SAC.Fields("canc_obs")
        End If

        'se tiver busca as ocorrências e atualiza o grid de ocorrências
        
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 transctc(txtFilial, txtCtc), "01"
            Set GridOcorr.DataSource = de_informa
            GridOcorr.DataMember = "Sel_ConsOcorr2"
            GridOcorr.Refresh
            If de_informa.rsSel_ConsOcorr2.RecordCount = 0 Then
                cmdExclOcorr.Enabled = False
                chkObsOcorr.Enabled = False
            Else
                cmdExclOcorr.Enabled = True
                chkObsOcorr.Enabled = True
            End If

        'se houver baixa atualiza campos de baixa. ocorrência = 01
        
        If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
        de_informa.Sel_ConsOcorr transctc(txtFilial, txtCtc), "01"
        If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr.Fields("baixadopre") = "S" Then
              'SE HOUVER PRÉ-BAIXA ATUALIZA CAMPOS DE PRÉ-BAIXA NO FORM
                lblDtPreBx.Caption = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
                lblHsPreBx.Caption = de_informa.rsSel_ConsOcorr.Fields("hsbaixapre")
                txtRecPreBx.Text = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                cmdExclPreBx.Enabled = True
            Else
              'SE NÃO HOUVER PRÉ-BAIXA ATUALIZA CAMPOS COM BRANCOS ("")
                lblDtPreBx.Caption = ""
                lblHsPreBx.Caption = ""
                txtRecPreBx.Text = ""
                cmdExclPreBx.Enabled = False
            End If
            If de_informa.rsSel_ConsOcorr.Fields("baixadofinal") = "S" Then
               'SE HOUVER BAIXA FINAL ATUALIZA CAMPOS DE BAIXA FINAL NO FORM
                chkCanhoto.Enabled = False
                If de_informa.rsSel_ConsOcorr.Fields("canhotonf") = "S" Then
                    chkCanhoto.Value = 1
                Else
                    chkCanhoto.Value = 0
                End If
                lblDtBx.Caption = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                lblHsBx.Caption = de_informa.rsSel_ConsOcorr.Fields("hsbaixa")
                txtRecBx.Text = de_informa.rsSel_ConsOcorr.Fields("receb")
                cmdExclBx.Enabled = True
                'If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("canhotonf")) Then
                '    If de_informa.rsSel_ConsOcorr.Fields("canhotonf") = "S" Then
                '        chkCanhoto.Value = 1
                '    Else
                '        chkCanhoto.Value = 0
                '    End If
                'Else
                '    chkCanhoto.Value = 0
                'End If
            Else
                chkCanhoto.Enabled = True
               'SE NÃO HOUVER BAIXA FINAL ATUALIZA CAMPOS COM BRANCOS ("")
                lblDtBx.Caption = ""
                lblHsBx.Caption = ""
                txtRecBx.Text = ""
                chkCanhoto.Value = 0
                cmdExclBx.Enabled = False
            End If
            chkObsEntr.Enabled = True
        Else
                chkCanhoto.Enabled = True
                lblDtBx.Caption = ""
                lblHsBx.Caption = ""
                txtRecBx.Text = ""
                lblDtPreBx.Caption = ""
                lblHsPreBx.Caption = ""
                txtRecPreBx.Text = ""
                chkCanhoto.Value = 0
                chkObsEntr.Enabled = False
                cmdExclBx.Enabled = False
                cmdExclPreBx.Enabled = False
        End If
        mskData.BackColor = &HC0FFFF      'AMARELO
        mskHora.BackColor = &HC0FFFF      'AMARELO
        txtCodOcorr.BackColor = &HC0FFFF      'AMARELO
        mskData.Enabled = True
        mskHora.Enabled = True
        txtCodOcorr.Enabled = True
        mskData.SetFocus
        cmbGravar.Enabled = True
        If xtemocorr = "C" Then
            txtCtc.SetFocus
            cmbGravar.Enabled = False
            mskData.Enabled = False
            mskHora.Enabled = False
            txtCodOcorr.Enabled = False
        End If
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
                de_informa.ins_LogUsuario "EXCLUSÃO", xUsuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " PRÉ-BAIXA/FÍSICA"
                
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
                de_informa.ins_LogUsuario "EXCLUSÃO", xUsuario, "POD/OCORR - CTC:" & transctc(txtFilial.Text, txtCtc.Text) & " PRÉ-BAIXA"
                
                de_informa.cn_informa.CommitTrans
                
                cmdProcurar_Click
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    If xultimofilial <> "" Then
        txtFilial.Text = xultimofilial
        txtCtc.Text = xultimoctc
    End If
        txtFilial.SetFocus
End Sub

Private Sub Form_Load()
        
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.StatusBar1.Visible = False
    'CONFIGURA OS OPTIONS, FRAMES E CHECKS
        optBaixaFinal.Enabled = False
        optPreBaixa.Enabled = False
        fraPreBaixa.Enabled = False
        fraBaixaFinal.Enabled = False
        GridOcorr.DataMember = ""
        GridOcorr.Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Visible = True
    mdiInforma.StatusBar1.Visible = True
    Set frmPod = Nothing
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
        If CDate(mskData.Text) > DataHora("data") Then
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

Private Sub optBaixaFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optCTC_Click()
    On Error Resume Next
    fraProcura.Caption = "Núm. da Filial e CTC"
    txtFilial.Visible = True
    txtCtc.Visible = True
    txtNumNf.Visible = False
    txtFilial.SetFocus
End Sub

Private Sub optNf_Click()
    On Error Resume Next
    fraProcura.Caption = "Núm. da NF"
    txtFilial.Visible = False
    txtCtc.Visible = False
    txtNumNf.Visible = True
    txtNumNf.SetFocus
End Sub

Private Sub optPreBaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtCodOcorr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtCodOcorr)) = 2 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    ElseIf KeyAscii = 13 And Len(Trim$(txtCodOcorr)) = 0 Then   'TECLA ENTER
        KeyAscii = 0
        frmBuscaOcorrencias.Show 1
        If Len(Trim$(txtCodOcorr)) = 2 Then
            SendKeys "{TAB}"  'ENVIA UM TAB
        End If
    End If
End Sub
Private Sub txtCtc_Change()
    On Error Resume Next
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
       fraPreBaixa.Enabled = True
       fraBaixaFinal.Enabled = False
       txtRecPreBx.BackColor = &HC0FFFF      'AMARELO
       txtRecBx.BackColor = &H8000000E       'BRANCO
       txtRecBx.Enabled = False
       'txtRecPreBx.SetFocus
       chkCanhoto.Value = 0


'    If lblDtPreBx.Caption = "" Then
'       fraPreBaixa.Enabled = True
'    Else
       fraPreBaixa.Enabled = False
'    End If
       fraBaixaFinal.Enabled = True
       txtRecBx.BackColor = &HC0FFFF      'amarelo
       txtRecPreBx.BackColor = &H8000000E       'BRANCO
       txtRecBx.Enabled = True
       'txtRecBx.SetFocus
       chkCanhoto.Value = 1
End Sub
Private Sub optPreBaixa_Click()
       fraPreBaixa.Enabled = True
       fraBaixaFinal.Enabled = False
       txtRecPreBx.BackColor = &HC0FFFF      'AMARELO
       txtRecBx.BackColor = &H8000000E       'BRANCO
       txtRecBx.Enabled = False
       'txtRecPreBx.SetFocus
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
