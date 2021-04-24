VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_diversos 
   Caption         =   "VARIOS SERVIÇOS MEUS!!!!!!"
   ClientHeight    =   10095
   ClientLeft      =   300
   ClientTop       =   465
   ClientWidth     =   14025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   14025
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      Begin VB.Frame Frame7 
         Height          =   4575
         Left            =   7080
         TabIndex        =   46
         Top             =   5160
         Width           =   5175
         Begin VB.Frame Frame10 
            Caption         =   "SALATIEL"
            Height          =   1215
            Left            =   3360
            TabIndex        =   60
            Top             =   240
            Width           =   1695
            Begin VB.CommandButton cmd_gerar_salatiel 
               Caption         =   "GERAR ARQUIVO"
               Height          =   495
               Left            =   120
               TabIndex        =   63
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox txt_mes_salatiel 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   720
               TabIndex        =   62
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label14 
               Caption         =   "Meses:"
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "ROCHE - PAULO"
            Height          =   1215
            Left            =   120
            TabIndex        =   59
            Top             =   1800
            Width           =   3135
         End
         Begin VB.Frame Frame8 
            Caption         =   "MEDLEY - MARIA"
            Height          =   1575
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   3135
            Begin MSComctlLib.ProgressBar prg_medley 
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   1200
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   1
            End
            Begin MSMask.MaskEdBox mas_medley2 
               Height          =   300
               Left            =   1560
               TabIndex        =   52
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mas_medley1 
               Height          =   300
               Left            =   120
               TabIndex        =   51
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.CommandButton cmd_medley 
               Caption         =   "GERA ARQUIVO"
               Height          =   255
               Left            =   720
               TabIndex        =   50
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label lab_medley 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   2400
               TabIndex        =   53
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Caption         =   "FINAL:"
               Height          =   255
               Left            =   1560
               TabIndex        =   49
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Caption         =   "INCIO:"
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   1095
            End
         End
      End
      Begin VB.Frame fra_diaria 
         Caption         =   "Qtd. Diárias"
         Height          =   4575
         Left            =   120
         TabIndex        =   29
         Top             =   5160
         Width           =   6855
         Begin VB.CommandButton cmd_gera_bomi 
            Caption         =   "GERAR ARQUIVO"
            Height          =   255
            Left            =   4680
            TabIndex        =   45
            Top             =   2880
            Width           =   1815
         End
         Begin MSDataGridLib.DataGrid grd_brasil_farma_nf 
            Bindings        =   "frm_diversos.frx":0000
            Height          =   1335
            Left            =   120
            TabIndex        =   44
            Top             =   2760
            Width           =   4335
            _ExtentX        =   7646
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
            DataMember      =   "sel_brasil_farma_nf"
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "DATA"
               Caption         =   "DATA"
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
               DataField       =   "DOC"
               Caption         =   "DOC"
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
               DataField       =   "FILIAL"
               Caption         =   "FILIAL"
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
               DataField       =   "QTD_NFS"
               Caption         =   "QTD_NFS"
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
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   615,118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   540,284
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   915,024
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmd_cgc 
            Caption         =   "CGC"
            Height          =   375
            Left            =   2400
            TabIndex        =   39
            Top             =   360
            Width           =   615
         End
         Begin VB.Frame fra_per_bomi 
            Caption         =   "Periodo"
            Height          =   975
            Left            =   3840
            TabIndex        =   38
            Top             =   120
            Width           =   2895
            Begin MSMask.MaskEdBox mas_b2 
               Height          =   300
               Left            =   1560
               TabIndex        =   41
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mas_b1 
               Height          =   300
               Left            =   240
               TabIndex        =   40
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   "FIM:"
               Height          =   255
               Left            =   1560
               TabIndex        =   43
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "INICIO:"
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmd_diario 
            Caption         =   "BUSCAR"
            Height          =   255
            Left            =   720
            TabIndex        =   37
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txt_cgc_bomi 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   36
            Top             =   360
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid grd_Brasil_farma 
            Bindings        =   "frm_diversos.frx":0017
            Height          =   1335
            Left            =   120
            TabIndex        =   34
            Top             =   1200
            Width           =   6615
            _ExtentX        =   11668
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
            DataMember      =   "Sel_brasil_Farma"
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "DATA"
               Caption         =   "DATA"
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
               DataField       =   "DOC"
               Caption         =   "DOC"
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
               DataField       =   "FILIAL"
               Caption         =   "FILIAL"
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
               DataField       =   "QTD_CTC"
               Caption         =   "QTD_CTC"
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
               DataField       =   "VOLUMES"
               Caption         =   "VOLUMES"
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
               DataField       =   "PESO"
               Caption         =   "PESO"
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
               DataField       =   "MERC"
               Caption         =   "MERC"
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
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   615,118
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   540,284
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txt_farma 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "04019475000180"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txt_Brasil 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "02426290000165"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Line Line1 
            X1              =   6720
            X2              =   120
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label8 
            Caption         =   "CGC:"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Bomi Farma:"
            Height          =   255
            Left            =   4200
            TabIndex        =   31
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Bomi Brasil:"
            Height          =   255
            Left            =   4200
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tabela Fox"
         Height          =   4935
         Left            =   7080
         TabIndex        =   23
         Top             =   240
         Width           =   5175
         Begin VB.CommandButton cmd_ver_tabela 
            Caption         =   "VER NF BASE_CLI"
            Height          =   495
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmd_exclui_fox 
            Caption         =   "DELETE BASE_CLI"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1560
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid grd_base_cli 
            Bindings        =   "frm_diversos.frx":002E
            Height          =   3375
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   5953
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "NF"
               Caption         =   "NF"
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
               DataField       =   "SERIE"
               Caption         =   "SERIE"
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
               DataField       =   "CLIENTENF"
               Caption         =   "CLIENTENF"
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
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   555,024
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label Label6 
            Caption         =   "Notas Localizadas:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   4440
            Width           =   1575
         End
         Begin VB.Label lab_qtd_nfs 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1680
            TabIndex        =   27
            Top             =   4440
            Width           =   615
         End
      End
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&SAINDO!"
         Height          =   375
         Left            =   12360
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "PROTOCOLOS"
         Height          =   4935
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6855
         Begin VB.Frame Frame5 
            Height          =   735
            Left            =   120
            TabIndex        =   14
            Top             =   3480
            Width           =   4335
            Begin VB.TextBox txt_mes 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   600
               MaxLength       =   2
               TabIndex        =   2
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txt_ano 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2520
               MaxLength       =   4
               TabIndex        =   3
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label3 
               Caption         =   "Ano:"
               Height          =   255
               Left            =   2160
               TabIndex        =   16
               Top             =   240
               Width           =   375
            End
            Begin VB.Label Label2 
               Caption         =   "MÊS:"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   3255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   4335
            Begin VB.Frame fra_pesq_proto 
               Caption         =   "Pesquisar Protocolos"
               Height          =   1815
               Left            =   120
               TabIndex        =   55
               Top             =   1320
               Width           =   4095
               Begin VB.CommandButton cmd_busca_ctc_proto 
                  Caption         =   "OK"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   58
                  Top             =   240
                  Width           =   495
               End
               Begin VB.TextBox txt_pesq_proto 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   960
                  TabIndex        =   57
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label Label13 
                  Caption         =   "Protocolo:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   56
                  Top             =   240
                  Width           =   735
               End
            End
            Begin VB.CommandButton cmd_busca_prot 
               Caption         =   "Command1"
               Height          =   255
               Left            =   120
               TabIndex        =   54
               Top             =   600
               Width           =   375
            End
            Begin VB.CommandButton cmd_protocolos 
               Caption         =   "Protocolos"
               Height          =   255
               Left            =   2640
               TabIndex        =   5
               Top             =   600
               Visible         =   0   'False
               Width           =   1455
            End
            Begin MSDataGridLib.DataGrid grd_protocolos 
               Bindings        =   "frm_diversos.frx":0045
               Height          =   1815
               Left            =   120
               TabIndex        =   12
               Top             =   1320
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   3201
               _Version        =   393216
               BackColor       =   16777215
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
               DataMember      =   "sel_busca_protocolos"
               ColumnCount     =   1
               BeginProperty Column00 
                  DataField       =   "PROTOCOLO"
                  Caption         =   "PROTOCOLO"
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
                     ColumnWidth     =   1739,906
                  EndProperty
               EndProperty
            End
            Begin VB.CommandButton cmd_pesq 
               Caption         =   "&Buscar"
               Height          =   255
               Left            =   720
               TabIndex        =   4
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox txt_cliente 
               Height          =   285
               Left            =   720
               TabIndex        =   1
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lab_cgc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H00000040&
               Height          =   285
               Left            =   2640
               TabIndex        =   18
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lab_nome 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   960
               Width           =   4095
            End
            Begin VB.Label Label1 
               Caption         =   "Cliente:"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   4200
            Width           =   4335
            Begin VB.CommandButton cmd_gerar 
               Caption         =   "&Gerar Arquivo"
               Height          =   255
               Left            =   600
               TabIndex        =   8
               Top             =   240
               Width           =   2895
            End
         End
         Begin MSDataGridLib.DataGrid grd_ctc 
            Bindings        =   "frm_diversos.frx":005C
            Height          =   3615
            Left            =   4560
            TabIndex        =   10
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6376
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
            DataMember      =   "sel_busca_ctc"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "FILIALCTC"
               Caption         =   "FILIALCTC"
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
               DataField       =   "NF"
               Caption         =   "NF"
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
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label lab_tot_gerados 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Left            =   5760
            TabIndex        =   22
            Top             =   4440
            Width           =   855
         End
         Begin VB.Label lab_txt_gerados 
            Caption         =   "Gerados:"
            Height          =   255
            Left            =   4680
            TabIndex        =   21
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label lab_qtd_ctc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   5760
            TabIndex        =   20
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Qtd. CTCS:"
            Height          =   255
            Left            =   4680
            TabIndex        =   19
            Top             =   4080
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frm_diversos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_busca_ctc_proto_Click()
Dim protocolo As String

protocolo = txt_pesq_proto.Text

If deb_bona.rssel_busca_ctc.State = 1 Then deb_bona.rssel_busca_ctc.Close
    deb_bona.sel_busca_ctc protocolo
    
    If deb_bona.rssel_busca_ctc.RecordCount < 1 Then
        MsgBox "Sem Registros", vbInformation, "FATURA: " & protocolo
    Else
        grd_ctc.DataMember = "sel_busca_ctc"
        grd_ctc.Refresh
    End If
    
    


End Sub

Private Sub cmd_busca_prot_Click()
If fra_pesq_proto.Visible = False Then
    fra_pesq_proto.Visible = True
Else
    fra_pesq_proto.Visible = False
End If

    


End Sub

Private Sub cmd_cgc_Click()
If fra_per_bomi.Visible = True Then
    fra_per_bomi.Visible = False
Else
    fra_per_bomi.Visible = True
End If



End Sub

Private Sub cmd_diario_Click()

If deb_bona.rsSel_brasil_Farma.State = 1 Then deb_bona.rsSel_brasil_Farma.Close
    deb_bona.Sel_brasil_Farma mas_b1, mas_b2, txt_cgc_bomi
    
    grd_Brasil_farma.DataMember = "sel_brasil_farma"
    grd_Brasil_farma.Refresh
    
If deb_bona.rssel_brasil_farma_nf.State = 1 Then deb_bona.rssel_brasil_farma_nf.Close
    deb_bona.sel_brasil_farma_nf mas_b1, mas_b2, txt_cgc_bomi
    
    grd_brasil_farma_nf.DataMember = "sel_brasil_farma_nf"
    grd_brasil_farma_nf.Refresh
    
    

End Sub

Private Sub cmd_exclui_fox_Click()
deb_bona.del_base_cli

End Sub

Private Sub cmd_gera_bomi_Click()
Dim xarquivo As String
Dim xarquivonf As String

If deb_bona.rsSel_brasil_Farma.State = 0 Then deb_bona.Sel_brasil_Farma mas_b1, mas_b2, txt_cgc_bomi

If txt_cgc_bomi.Text = "02426290000165" Then
    xarquivo = "C:\BOMI_BRASIL_CTC"
    xarquivonf = "C:\BOMI_BRASIL_NF"
ElseIf txt_cgc_bomi.Text = "04019475000180" Then
    xarquivo = "C:\BOMI_FARMA_CTC"
    xarquivonf = "C:\BOMI_FARMA_NF"
End If

Open xarquivo & Mid(Date, 1, 2) & Mid(Date, 4, 2) & ".TXT" For Output As #1
xdata = "DATA"
XDOC = "DOC"
xfilial = "FILIAL"
xqtd = "QTD_CTC"
xvolumes = "VOLUMES"
xpeso = "PESO"
xmerc = "MERCADORIA"
xlinha = xdata & "#" & XDOC & "#" & xfilial & "#" & xqtd & "#" & xvolumes & "#" & xpeso & "#" & xmerc
Print #1, xlinha
deb_bona.rsSel_brasil_Farma.MoveFirst
Do Until deb_bona.rsSel_brasil_Farma.EOF
    
    With deb_bona.rsSel_brasil_Farma
    
    xdata = .Fields("DATA")
    XDOC = .Fields("DOC")
    xfilial = .Fields("FILIAL")
    xqtd = .Fields("QTD_CTC")
    xvolumes = .Fields("VOLUMES")
    xpeso = .Fields("PESO")
    xmerc = .Fields("MERC")
    xlinha = xdata & "#" & XDOC & "#" & xfilial & "#" & xqtd & "#" & xvolumes & "#" & xpeso & "#" & xmerc
    Print #1, xlinha
    .MoveNext
    End With
    Loop
    
    Close #1
    
    If deb_bona.rssel_brasil_farma_nf.State = 0 Then deb_bona.sel_brasil_farma_nf mas_b1, mas_b2, txt_cgc_bomi
    
    
    Open xarquivonf & Mid(Date, 1, 2) & Mid(Date, 4, 2) & ".TXT" For Output As #1
    xdata = "DATA"
    XDOC = "DOC"
    xfilial = "FILIAL"
    xqtd = "QTD_NF"
    xlinha = xdata & "#" & XDOC & "#" & xfilial & "#" & xqtd
    Print #1, xlinha
    deb_bona.rssel_brasil_farma_nf.MoveFirst
    With deb_bona.rssel_brasil_farma_nf
    Do Until deb_bona.rssel_brasil_farma_nf.EOF
    xdata = .Fields("DATA")
    XDOC = .Fields("DOC")
    xfilial = .Fields("FILIAL")
    xqtd = .Fields("QTD_NFS")
    xlinha = xdata & "#" & XDOC & "#" & xfilial & "#" & xqtd
    Print #1, xlinha
    .MoveNext
    Loop
    End With
    Close #1
    
        
        
    MsgBox "Arquivo Gerado", vbInformation, xarquivo
    
    
    
    

    



End Sub

Private Sub CMD_GERAR_Click()
Dim xarqproto As String
Dim gerados As Integer

If fra_pesq_proto.Visible = False Then

    xarqproto = "C:\PROTO" & Trim(deb_bona.rssel_busca_protocolos.Fields("PROTOCOLO"))
    
Else
    xarqproto = "C:\PROTO" & txt_pesq_proto.Text
End If


If deb_bona.rssel_busca_ctc.State = 0 Then deb_bona.sel_busca_ctc deb_bona.rssel_busca_protocolos.Fields("PROTOCOLO")

Open xarqproto & ".TXT" For Output As #1
xdata = "DATA"
xhora = "HORA"
xprotocolo = "PROTOCOLO"
xqtd = "QTD_DE_CTCS"
xcliente = "CLIENTE"
xctc = "FILIAL_CTC"
xnf = "NF"
xlinha = xctc & "#" & xnf
Print #1, xdata & "#" & Date
Print #1, xhora & "#" & Time
Print #1, xcliente & "#" & lab_nome
If fra_pesq_proto.Visible = False Then
    Print #1, xprotocolo & "#" & deb_bona.rssel_busca_protocolos.Fields("PROTOCOLO")
Else
    Print #1, xprotocolo & "#" & txt_pesq_proto.Text
End If

Print #1, xqtd & "#" & lab_qtd_ctc.Caption
Print #1, ""
Print #1, xlinha
Print #1, ""
deb_bona.rssel_busca_ctc.MoveFirst
Do Until deb_bona.rssel_busca_ctc.EOF
    xctc = "'" & deb_bona.rssel_busca_ctc.Fields("FILIALCTC")
    xnf = deb_bona.rssel_busca_ctc.Fields("NF")
    xlinha = xctc & "#" & xnf
    Print #1, xlinha
    deb_bona.rssel_busca_ctc.MoveNext
    gerados = gerados + 1
    lab_tot_gerados.Caption = gerados
    Loop
Close #1


MsgBox "Arquivo " & xarqproto & " Gerado com Sucesso.", vbInformation, "OK"

    



End Sub

Private Sub cmd_gerar_salatiel_Click()
Dim xmes As String
Dim xcidade As String
Dim xmerc As String
Dim xvolumes As String
Dim xpeso As String
Dim xfrete As String
Dim xqtd_ctc As String

If deb_bona.rssel_salatiel_ctc.State = 1 Then deb_bona.rssel_salatiel_ctc.Close
    deb_bona.sel_salatiel_ctc txt_mes_salatiel.Text
    
    If deb_bona.rssel_salatiel_ctc.RecordCount < 1 Then
        MsgBox "SEM DADOS", vbInformation, "SALATIEL"
    Else
        Open "C:\SALATIEL_CTC.TXT" For Output As #1
            xmes = "MES"
            xcidade = "CIDADE"
            xmerc = "MERCADORIA"
            xvolumes = "VOLUMES"
            xpeso = "PESO"
            xfrete = "FRETE"
            xqtd_ctc = "QTD_CTC"
            xlinha = xmes & "#" & xcidade & "#" & xmerc & "#" & xvolumes & "#" & xpeso & "#" & xfrete & "#" & xqtd_ctc
            Print #1, xlinha
            With deb_bona.rssel_salatiel_ctc
            .MoveFirst
            Do Until .EOF
                xmes = .Fields("MES")
                xcidade = .Fields("CIDADE")
                xmerc = .Fields("MERCADORIA")
                xvolumes = .Fields("VOLUMES")
                xpeso = .Fields("PESO")
                xfrete = .Fields("FRETE")
                xqtd_ctc = .Fields("QTD_CTC")
                xlinha = xmes & "#" & xcidade & "#" & xmerc & "#" & xvolumes & "#" & xpeso & "#" & xfrete & "#" & xqtd_ctc
                Print #1, xlinha
                .MoveNext
                Loop
                End With
                
                Close #1
                
                'MsgBox "C:\SALATIEL_CTC.TXT GERADO", vbInformation, "SALATIEL"
                
            End If
            
If deb_bona.rssel_salatiel_nf.State = 1 Then deb_bona.rssel_salatiel_nf.Close
    deb_bona.sel_salatiel_nf txt_mes_salatiel.Text
    
    If deb_bona.rssel_salatiel_nf.RecordCount < 1 Then
        MsgBox "SEM NFS", vbInformation, "SALATIEL"
    Else
    
    Open "C:\SALATIEL_NF.TXT" For Output As #1
        xdata = "DATA"
        xcidade = "CIDADE"
        xqtd_ctc = "QTD_NFS"
        xlinha = xdata & "#" & xcidade & "#" & xqtd_ctc
        Print #1, xlinha
        With deb_bona.rssel_salatiel_nf
        .MoveFirst
        Do Until .EOF
            xdata = .Fields("DATA")
            xcidade = .Fields("CIDADE")
            xqtd_ctc = .Fields("QTD_NFS")
            xlinha = xdata & "#" & xcidade & "#" & xqtd_ctc
            Print #1, xlinha
            .MoveNext
            Loop
            End With
    Close #1
    
        MsgBox "ARQUIVOS GERADOS", vbInformation, "SALATIEL"
        
    End If
    
                


End Sub

Private Sub cmd_medley_Click()
Dim xdata As String
Dim xnf As String
Dim xpeso As String
Dim xvalor As String
Dim xvolumes As String
Dim xlinha As String
Dim date1 As Date
Dim date2 As Date
Dim xprg As Integer


date1 = mas_medley1
date2 = mas_medley2



If deb_bona.rssel_medley.State = 1 Then deb_bona.rssel_medley.Close
    deb_bona.sel_medley date1, date2
    
    lab_medley.Caption = deb_bona.rssel_medley.RecordCount
    Frame7.Refresh
    prg_medley.Min = xprg
    prg_medley.Max = deb_bona.rssel_medley.RecordCount
    prg_medley.Value = xprg
    
    
        Open "C:\MEDLEY.TXT" For Output As #1
        xdata = "DATA"
        xnf = "NF"
        xpeso = "PESO"
        xvalor = "VALOR"
        xvolumes = "VOLUMES"
        xlinha = xdata & "#" & xnf & "#" & xpeso & "#" & xvalor & "#" & xvolumes
        Print #1, xlinha
        deb_bona.rssel_medley.MoveFirst
        With deb_bona.rssel_medley
        Do Until .EOF
        xdata = .Fields("DATA")
        xnf = .Fields("NF")
        xpeso = .Fields("PESO")
        xvalor = .Fields("VALOR")
        xvolumes = .Fields("VOLUMES")
        xlinha = xdata & "#" & xnf & "#" & xpeso & "#" & xvalor & "#" & xvolumes
        Print #1, xlinha
        .MoveNext
        
        xprg = xprg + 1
        
        prg_medley.Value = xprg
        
        
        Loop
        End With
        Close #1
        
        
        MsgBox "OK"
        


End Sub

Private Sub cmd_pesq_Click()
Dim cliente As String
Dim cgc As String


cliente = txt_cliente.Text


If txt_ano.Text = Empty And txt_mes.Text = Empty Then
    MsgBox "Selecione 'Mês' e 'Ano'", vbInformation, "DATAS"
    
    If txt_mes.Text = Empty Then
        txt_mes.SetFocus
    Else
        txt_ano.SetFocus
    End If
    
    
    
Else
    If deb_bona.rssel_busca_cli.State = 1 Then deb_bona.rssel_busca_cli.Close
        deb_bona.sel_busca_cli txt_ano, txt_mes, "%" & txt_cliente.Text & "%"
           
        If deb_bona.rssel_busca_cli.RecordCount < 1 Then
            MsgBox "Cliente inexistente", vbInformation, "CLIENTES"
        Else
            frm_clientes.grd_cliente.DataMember = "sel_busca_cli"
            frm_clientes.grd_cliente.Refresh
            frm_clientes.Show
            
            DoEvents
            
End If
End If



End Sub

Private Sub cmd_protocolos_Click()

If deb_bona.rssel_busca_protocolos.State = 1 Then deb_bona.rssel_busca_protocolos.Close
    deb_bona.sel_busca_protocolos lab_cgc, txt_ano, txt_mes
    
    If deb_bona.rssel_busca_protocolos.RecordCount < 1 Then
        MsgBox "Sem Protocolos", vbInformation, "PROTOCOLOS"
    Else
        grd_protocolos.DataMember = "sel_busca_protocolos"
        grd_protocolos.Refresh
    
    End If
    
    
    

End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub



Private Sub cmd_ver_tabela_Click()
deb_bona.rssel_base_cli.Open
lab_qtd_nfs.Caption = deb_bona.rssel_base_cli.RecordCount

If deb_bona.rssel_base_cli.RecordCount < 1 Then
    MsgBox "NFS não localizadas", vbInformation, "BASECLI"
    deb_bona.rssel_base_cli.Close
    Exit Sub
Else
grd_base_cli.DataMember = "sel_base_cli"
grd_base_cli.Refresh

lab_qtd_nfs.Caption = deb_bona.rssel_base_cli.RecordCount
cmd_exclui_fox.Enabled = True
End If


deb_bona.rssel_base_cli.Close


End Sub

Private Sub Command1_Click()
Open "\\192.9.205.84\hp_frank" For Output As #1
   Print #1, "ROMULO  GATAO"
   Close #1
End Sub

Private Sub grd_protocolos_Click()
Dim proto As String

proto = deb_bona.rssel_busca_protocolos.Fields("PROTOCOLO")

If deb_bona.rssel_busca_ctc.State = 1 Then deb_bona.rssel_busca_ctc.Close
    deb_bona.sel_busca_ctc proto
    
    grd_ctc.DataMember = "sel_busca_ctc"
    grd_ctc.Refresh
    lab_qtd_ctc.Caption = deb_bona.rssel_busca_ctc.RecordCount
    

End Sub

Private Sub txt_Brasil_DblClick()
txt_cgc_bomi.Text = txt_Brasil.Text

End Sub

Private Sub txt_farma_DblClick()
txt_cgc_bomi.Text = txt_farma.Text
End Sub
