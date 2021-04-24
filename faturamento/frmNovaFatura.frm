VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNovaFatura 
   Caption         =   "Nova Pré-Fatura"
   ClientHeight    =   7995
   ClientLeft      =   1320
   ClientTop       =   1470
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCliente 
      Caption         =   "Cliente / Consignatário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      TabIndex        =   43
      Top             =   80
      Width           =   5655
      Begin VB.CommandButton cmdBuscaPrefat 
         Caption         =   ">>"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtPreFilial 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtPreFatura 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtFilialFatura 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtCnpj 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   14
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtNumBanco 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1800
         Width           =   495
      End
      Begin MSMask.MaskEdBox mskVencimento 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "PréFatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ ou ?"
         Height          =   195
         Left            =   4320
         TabIndex        =   60
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   1560
         TabIndex        =   50
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   49
         Top             =   1080
         Width           =   4560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vencto:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label lblNomeBanco 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   46
         Top             =   1800
         Width           =   1890
      End
      Begin VB.Label lblContaBanco 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4320
         TabIndex        =   45
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         Height          =   195
         Left            =   3600
         TabIndex        =   44
         Top             =   1800
         Width           =   465
      End
   End
   Begin VB.Frame fraDadosCob 
      Caption         =   "Endereço de Cobrança"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   5880
      TabIndex        =   28
      Top             =   80
      Width           =   5895
      Begin VB.TextBox txtContatoCob 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   34
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdEndPadraoCob 
         Caption         =   "End. Padrão"
         Enabled         =   0   'False
         Height          =   465
         Left            =   1800
         TabIndex        =   40
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtEndCob 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   29
         Top             =   360
         Width           =   4695
      End
      Begin VB.TextBox txtCidCob 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         MaxLength       =   35
         TabIndex        =   31
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtUFCob 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   32
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtCepCob 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   30
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtFoneCob 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   33
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdAlterarCob 
         Caption         =   "Alterar"
         Enabled         =   0   'False
         Height          =   465
         Left            =   120
         TabIndex        =   42
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdGravarCob 
         Caption         =   "Gravar"
         Enabled         =   0   'False
         Height          =   465
         Left            =   3480
         TabIndex        =   37
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancCob 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   465
         Left            =   4680
         TabIndex        =   39
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cep/Cidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
         Height          =   195
         Left            =   2880
         TabIndex        =   35
         Top             =   1080
         Width           =   600
      End
   End
   Begin VB.Frame fraFatura 
      Height          =   5580
      Left            =   120
      TabIndex        =   6
      Top             =   2300
      Width           =   11655
      Begin VB.Frame Frame1 
         Height          =   3540
         Left            =   120
         TabIndex        =   55
         Top             =   1950
         Width           =   11415
         Begin VB.CommandButton cmdNovaPreFat 
            Caption         =   "Nova ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   89
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdGerarFat 
            Caption         =   "Gerar Fatura ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   7320
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluirTudo 
            Caption         =   "Excluir Tudo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   88
            Top             =   240
            Width           =   1095
         End
         Begin MSDataGridLib.DataGrid gridItens 
            Bindings        =   "frmNovaFatura.frx":0000
            Height          =   2655
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   4683
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
            DataMember      =   "Sel_PreFatura"
            ColumnCount     =   31
            BeginProperty Column00 
               DataField       =   "tipodoc"
               Caption         =   "Doc"
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
            BeginProperty Column02 
               DataField       =   "data"
               Caption         =   "Data CTC"
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
               DataField       =   "frete"
               Caption         =   "Frete Liq."
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
            BeginProperty Column04 
               DataField       =   "fretebruto"
               Caption         =   "Frete Bruto"
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
            BeginProperty Column05 
               DataField       =   "remet_cgc"
               Caption         =   "Remet_CGC"
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
               DataField       =   "remet_nome"
               Caption         =   "Remet_Nome"
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
               Caption         =   "Observações"
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
               DataField       =   "filialprefatura"
               Caption         =   "filialprefatura"
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
               DataField       =   "digitacao"
               Caption         =   "digitacao"
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
               DataField       =   "emissor"
               Caption         =   "emissor"
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
               DataField       =   "vencimento"
               Caption         =   "vencimento"
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
               DataField       =   "cliente_cgc"
               Caption         =   "cliente_cgc"
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
               DataField       =   "cliente_nome"
               Caption         =   "cliente_nome"
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
               DataField       =   "cliente_ie"
               Caption         =   "cliente_ie"
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
               DataField       =   "cliente_end"
               Caption         =   "cliente_end"
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
               DataField       =   "cliente_cidade"
               Caption         =   "cliente_cidade"
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
               DataField       =   "cliente_uf"
               Caption         =   "cliente_uf"
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
               DataField       =   "cliente_cep"
               Caption         =   "cliente_cep"
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
               DataField       =   "endcob"
               Caption         =   "endcob"
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
               DataField       =   "cidadecob"
               Caption         =   "cidadecob"
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
            BeginProperty Column21 
               DataField       =   "ufcob"
               Caption         =   "ufcob"
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
            BeginProperty Column22 
               DataField       =   "cepcob"
               Caption         =   "cepcob"
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
               DataField       =   "telefonecob"
               Caption         =   "telefonecob"
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
            BeginProperty Column24 
               DataField       =   "contatocob"
               Caption         =   "contatocob"
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
            BeginProperty Column25 
               DataField       =   "banco"
               Caption         =   "banco"
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
            BeginProperty Column26 
               DataField       =   "banconome"
               Caption         =   "banconome"
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
            BeginProperty Column27 
               DataField       =   "conta"
               Caption         =   "conta"
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
            BeginProperty Column28 
               DataField       =   "avulsa"
               Caption         =   "avulsa"
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
            BeginProperty Column29 
               DataField       =   "valorbrutoicms"
               Caption         =   "valorbrutoicms"
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
            BeginProperty Column30 
               DataField       =   "valorbruto"
               Caption         =   "valorbruto"
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
                  ColumnWidth     =   434,835
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1019,906
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1110,047
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1124,787
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1395,213
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   2940,095
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   6164,788
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column13 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column14 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column15 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column16 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column17 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column18 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column19 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column20 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column21 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   540,284
               EndProperty
               BeginProperty Column22 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column23 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column24 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column25 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column26 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column27 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column28 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   540,284
               EndProperty
               BeginProperty Column29 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column30 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdSair 
            Caption         =   "S A I R"
            Height          =   375
            Left            =   10200
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdExcluirCTC 
            Caption         =   "Excluir CTC"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTotalFaturaICMS 
            AutoSize        =   -1  'True
            Caption         =   "COMICMS"
            Height          =   195
            Left            =   2880
            TabIndex        =   63
            Top             =   480
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Qtde CTCs:"
            Height          =   195
            Left            =   5400
            TabIndex        =   59
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lblQtdeCtc 
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
            Left            =   6360
            TabIndex        =   58
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblTotalFatura 
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
            Left            =   4080
            TabIndex        =   57
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total da Fatura:"
            Height          =   195
            Left            =   2520
            TabIndex        =   56
            Top             =   240
            Width           =   1530
         End
      End
      Begin TabDlg.SSTab tabTipoFatura 
         Height          =   1695
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   2990
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Manual (CTC por CTC)"
         TabPicture(0)   =   "frmNovaFatura.frx":0019
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label15"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label16"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblRemetente"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblDestinatario"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label23"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label24"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblEmissao"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblModal"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label18"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lblConsignatario"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label21"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lblFrete"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "lblCartaCorr"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtCTC"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "cmdBuscaCTC"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "optCTC"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "optNFS"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtFilial"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "cmdIncluirCTC"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "chkLeitor"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Timer1"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).ControlCount=   22
         TabCaption(1)   =   "Por Intervalo CTCs"
         TabPicture(1)   =   "frmNovaFatura.frx":0035
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label10"
         Tab(1).Control(1)=   "Label37"
         Tab(1).Control(2)=   "lblNomeRemetInt"
         Tab(1).Control(3)=   "Label35"
         Tab(1).Control(4)=   "Label34"
         Tab(1).Control(5)=   "Label33"
         Tab(1).Control(6)=   "Command6"
         Tab(1).Control(7)=   "Text2"
         Tab(1).Control(8)=   "txtFilialInt"
         Tab(1).Control(9)=   "txtCtc1Int"
         Tab(1).Control(10)=   "txtCtc2Int"
         Tab(1).Control(11)=   "txtRemetInt"
         Tab(1).Control(12)=   "cmdBuscaRemetInt"
         Tab(1).Control(13)=   "cmdIncluirCTCsIntervalo"
         Tab(1).Control(14)=   "chkNormalInt"
         Tab(1).Control(15)=   "chkUrgenciaInt"
         Tab(1).Control(16)=   "chkEntregaInt"
         Tab(1).Control(17)=   "chkReentregaInt"
         Tab(1).Control(18)=   "chkDevolucaoInt"
         Tab(1).Control(19)=   "chkTransferenciaInt"
         Tab(1).Control(20)=   "chkRodoviarioInt"
         Tab(1).Control(21)=   "chkAereoInt"
         Tab(1).ControlCount=   22
         TabCaption(2)   =   "Por Pré-Fatura"
         TabPicture(2)   =   "frmNovaFatura.frx":0051
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label20"
         Tab(2).Control(1)=   "Label22"
         Tab(2).Control(2)=   "lblTotalCtcsPreFat"
         Tab(2).Control(3)=   "optPFVideolar"
         Tab(2).Control(4)=   "optPFProceda"
         Tab(2).Control(5)=   "txtPFArquivo"
         Tab(2).Control(6)=   "cmdProcessarArqPreFatCliente"
         Tab(2).Control(7)=   "txtFilialPreFat"
         Tab(2).Control(8)=   "bar1"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Fatura Avulsa"
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdFecharFaturaAvulsa"
         Tab(3).Control(1)=   "txtValorFatAvulsa"
         Tab(3).Control(2)=   "txtAvulsaDescr"
         Tab(3).Control(3)=   "Label38"
         Tab(3).Control(4)=   "Label12"
         Tab(3).Control(5)=   "Label14"
         Tab(3).ControlCount=   6
         Begin MSComctlLib.ProgressBar bar1 
            Height          =   255
            Left            =   -70800
            TabIndex        =   103
            Top             =   480
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.TextBox txtFilialPreFat 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   -67320
            MaxLength       =   2
            TabIndex        =   100
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdProcessarArqPreFatCliente 
            Caption         =   "Processar Pré-Fatura"
            Height          =   495
            Left            =   -66480
            TabIndex        =   101
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtPFArquivo 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   -73320
            MaxLength       =   50
            TabIndex        =   99
            Text            =   "C:\INFORMA\INTPFT0000.TXT"
            Top             =   1080
            Width           =   3855
         End
         Begin VB.OptionButton optPFProceda 
            Caption         =   "Padrão Proceda"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -72960
            TabIndex        =   97
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optPFVideolar 
            Caption         =   "Modelo Videolar"
            Height          =   195
            Left            =   -74760
            TabIndex        =   96
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.CommandButton cmdFecharFaturaAvulsa 
            Caption         =   "Gravar Fatura Avulsa..."
            Height          =   375
            Left            =   -66120
            TabIndex        =   92
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txtValorFatAvulsa 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   -73440
            MaxLength       =   15
            TabIndex        =   91
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtAvulsaDescr 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   -70920
            MaxLength       =   30
            TabIndex        =   90
            Top             =   600
            Width           =   4575
         End
         Begin VB.CheckBox chkAereoInt 
            Caption         =   "Aéreo"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -67320
            TabIndex        =   81
            Top             =   720
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkRodoviarioInt 
            Caption         =   "Rodoviário"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -67320
            TabIndex        =   80
            Top             =   480
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkTransferenciaInt 
            Caption         =   "Transferência"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -68880
            TabIndex        =   79
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox chkDevolucaoInt 
            Caption         =   "Devoluções"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -68880
            TabIndex        =   78
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox chkReentregaInt 
            Caption         =   "Reentrega"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -68880
            TabIndex        =   77
            Top             =   720
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkEntregaInt 
            Caption         =   "Entrega"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -68880
            TabIndex        =   76
            Top             =   480
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkUrgenciaInt 
            Caption         =   "Urgência"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -67320
            TabIndex        =   75
            Top             =   1200
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkNormalInt 
            Caption         =   "Normal/Prioridade"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -67320
            TabIndex        =   74
            Top             =   960
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CommandButton cmdIncluirCTCsIntervalo 
            Caption         =   "Incluir CTCs..."
            Height          =   495
            Left            =   -65400
            TabIndex        =   73
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdBuscaRemetInt 
            Caption         =   "?"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -72120
            TabIndex        =   72
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtRemetInt 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   -73800
            MaxLength       =   14
            TabIndex        =   71
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtCtc2Int 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   -71880
            MaxLength       =   8
            TabIndex        =   70
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtCtc1Int 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   -73200
            MaxLength       =   8
            TabIndex        =   69
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtFilialInt 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   -74400
            MaxLength       =   2
            TabIndex        =   68
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   285
            Left            =   -73800
            MaxLength       =   14
            TabIndex        =   67
            Top             =   1200
            Width           =   2775
         End
         Begin VB.CommandButton Command6 
            Caption         =   "?"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -70920
            TabIndex        =   66
            Top             =   1200
            Width           =   375
         End
         Begin VB.Timer Timer1 
            Interval        =   150
            Left            =   240
            Top             =   120
         End
         Begin VB.CheckBox chkLeitor 
            Caption         =   "Leitor (Cod.Barras)"
            Height          =   195
            Left            =   2400
            TabIndex        =   18
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton cmdIncluirCTC 
            Caption         =   "Incluir CTC/NFS"
            Enabled         =   0   'False
            Height          =   555
            Left            =   3000
            TabIndex        =   13
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtFilial 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   840
            MaxLength       =   2
            TabIndex        =   10
            Top             =   960
            Width           =   375
         End
         Begin VB.OptionButton optNFS 
            Caption         =   "NF Serviço"
            Height          =   195
            Left            =   960
            TabIndex        =   9
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optCTC 
            Caption         =   "CTC"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton cmdBuscaCTC 
            Caption         =   ">>"
            Height          =   320
            Left            =   2400
            TabIndex        =   12
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtCTC 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   8
            TabIndex        =   11
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lblTotalCtcsPreFat 
            AutoSize        =   -1  'True
            Caption         =   "00000/00000"
            Height          =   195
            Left            =   -65040
            TabIndex        =   104
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Filial Padrão Para os CTCs:"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   -69360
            TabIndex        =   102
            Top             =   1080
            Width           =   1920
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Local do Arquivo: "
            Height          =   195
            Left            =   -74760
            TabIndex        =   98
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Valor da Fatura:"
            Height          =   195
            Left            =   -74760
            TabIndex        =   95
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "** ATENÇÃO ! Para Fatura Avulsa, Não Há Pré-Fatura **"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -71520
            TabIndex        =   94
            Top             =   1320
            Width           =   4815
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   -71760
            TabIndex        =   93
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Filial:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   87
            Top             =   480
            Width           =   345
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   -72120
            TabIndex        =   86
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Remetente:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   85
            Top             =   840
            Width           =   825
         End
         Begin VB.Label lblNomeRemetInt 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   -71640
            TabIndex        =   84
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "CTC de:"
            Height          =   195
            Left            =   -73920
            TabIndex        =   83
            Top             =   480
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nat. Produto:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   82
            Top             =   1200
            Width           =   945
         End
         Begin VB.Label lblCartaCorr 
            AutoSize        =   -1  'True
            Caption         =   "OBSCC"
            Height          =   195
            Left            =   9120
            TabIndex        =   61
            Top             =   1440
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblFrete 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9840
            TabIndex        =   54
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Frete:"
            Height          =   195
            Left            =   9120
            TabIndex        =   53
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label lblConsignatario 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5880
            TabIndex        =   52
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Consignatário:"
            Height          =   195
            Left            =   4800
            TabIndex        =   51
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label lblModal 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9840
            TabIndex        =   27
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblEmissao 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9840
            TabIndex        =   26
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   9120
            TabIndex        =   25
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   9120
            TabIndex        =   24
            Top             =   480
            Width           =   630
         End
         Begin VB.Label lblDestinatario 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5880
            TabIndex        =   23
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblRemetente 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5880
            TabIndex        =   22
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Remetente:"
            Height          =   195
            Left            =   4800
            TabIndex        =   21
            Top             =   480
            Width           =   825
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Destinatário:"
            Height          =   195
            Left            =   4800
            TabIndex        =   20
            Top             =   840
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "FilialCTC:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   660
         End
      End
   End
End
Attribute VB_Name = "frmNovaFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterarCob_Click()
    fraCliente.Enabled = False
    fraFatura.Enabled = False
    cmdAlterarCob.Enabled = False
    cmdGravarCob.Enabled = True
    cmdCancCob.Enabled = True
    cmdEndPadraoCob.Enabled = True
    txtEndCob.Enabled = True
    txtEndCob.BackColor = xamarelo1
    txtCepCob.Enabled = True
    txtCepCob.BackColor = xamarelo1
    txtCidCob.Enabled = True
    txtCidCob.BackColor = xamarelo1
    txtUFCob.Enabled = True
    txtUFCob.BackColor = xamarelo1
    txtFoneCob.Enabled = True
    txtFoneCob.BackColor = xamarelo1
    txtContatoCob.Enabled = True
    txtContatoCob.BackColor = xamarelo1
    txtEndCob.SetFocus
End Sub

Private Sub cmdAlteraVencPrefat_Click()

End Sub

Private Sub cmdBuscaCTC_Click()
    
    If Len(Trim$(txtCnpj)) < 11 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: CLIENTE !"
        txtCnpj.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mskVencimento) Then
        MsgBox "Data de Vencimento Inválida !"
        mskVencimento.SetFocus
        Exit Sub
    End If

    If Len(Trim$(txtNumBanco)) < 1 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: NÚMERO DO BANCO !"
        txtNumBanco.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtEndCob)) < 5 Or Len(Trim$(txtCidCob)) < 3 Or Len(Trim$(txtCepCob)) < 5 Or Len(Trim$(txtFoneCob)) < 5 Then
        MsgBox "Dados de Endereço de Cobrança Inválidos !"
        cmdAlterarCob.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtFilialFatura)) < 2 Then
        MsgBox "Filial da Fatura Inválida !"
        txtFilialFatura.SetFocus
        Exit Sub
    End If
    
    cmdIncluirCTC.Enabled = False
    lblRemetente = ""
    lblDestinatario = ""
    lblConsignatario = ""
    lblEmissao = ""
    lblModal = ""
    lblFrete = ""
    lblCartaCorr = ""
    
    If chkLeitor.Value = 1 Then
        If Len(Trim$(txtCTC)) = 7 Then
            txtCTC = Mid$(txtCTC, 1, 6)
        End If
    End If
    
    If optCTC.Value = True Then
    
        If de_informa.rsSel_CTC.State = 1 Then de_informa.rsSel_CTC.Close
        de_informa.Sel_CTC transctc(txtFilial, txtCTC)
        
        If de_informa.rsSel_CTC.RecordCount < 1 Then
            'não encontrou
            MsgBox "Não Encontrado CTC com este Número !", vbCritical, "Não Encontrado"
            txtCTC.SetFocus
            Exit Sub
        Else
            lblRemetente = de_informa.rsSel_CTC.Fields("remet_nome")
            lblDestinatario = de_informa.rsSel_CTC.Fields("dest_nome")
            lblConsignatario = de_informa.rsSel_CTC.Fields("respons_nome")
            lblEmissao = de_informa.rsSel_CTC.Fields("data")
            lblModal = de_informa.rsSel_CTC.Fields("modal")
            If de_informa.rsSel_CTC.Fields("tipodoc") = "CTC" Or de_informa.rsSel_CTC.Fields("tipodoc") = "COB" Then
                'encontrou CTC
                If de_informa.rsSel_CTC.Fields("tem_ocorr") = "C" Then
                    'ctc cancelado
                    MsgBox "Este CTC Encontra-se Cancelado !!", vbCritical, "CTC Cancelado"
                    txtCTC.SetFocus
                    Exit Sub
                Else
                    If Len(Trim$(de_informa.rsSel_CTC.Fields("faturanum"))) > 1 Then
                        'CTC Encontra-se Faturado
                        MsgBox "Este CTC Encontra-se Faturado !! Fatura: " & de_informa.rsSel_CTC.Fields("faturanum"), vbCritical, "CTC Faturado"
                        txtCTC.SetFocus
                        Exit Sub
                    End If
                    If de_informa.rsSel_CTC.Fields("respons_cgc") <> zeros2(txtCnpj, 14) Then
                        'O Consignatario não é o Mesmo
                        Do While lblCartaCorr = ""
                            If MsgBox("Atenção! O Consignatário deste CTC é Diferente do Cliente Desta Fatura. Pode Ocorrer de o CTC estar com o Consignatário Errado, Necessitando a Emissão de uma Carta de Correção. Deseja, Mesmo Assim Incluir Este CTC Nesta Fatura ? (você terá que informar o número da Carta de Correção).", vbCritical + vbYesNo, "Consignatário Diferente") = vbNo Then
                                txtCTC.SetFocus
                                Exit Sub
                            Else
                                lblCartaCorr = InputBox("Entre com o Número da Carta de Correção que Retifica o Consignatário deste CTC:", "Número da Carta de Correção")
                                If lblCartaCorr = "" Then
                                Else
                                    lblCartaCorr = "Carta de Correção (Consignatário): " & lblCartaCorr
                                End If
                            End If
                        Loop
                    End If
                    
                    If IsNull(de_informa.rsSel_CTC.Fields("subtrib")) Then
                        'faturando CTCs do SITLA
                        If de_informa.rsSel_CTC.Fields("fretetotal") <> de_informa.rsSel_CTC.Fields("fretetotalbruto") Then
                            If MsgBox("(SITLA) Este CTC Caracteriza a Regra de Substituição Tributária ? (Confira no CTC. Se Positivo Será Cobrado Pelo Valor Líquido.)", vbQuestion + vbYesNo, "CTC do SITLA") = vbYes Then
                                'subst.tributarui = S ; Pega valor pelo liquido
                                lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotal"), "##,###,##0.00")
                                If de_informa.rsSel_CTC.Fields("fretetotal") < 0 Then
                                    MsgBox "Valor de Frete deste CTC Inválido !"
                                    txtCTC.SetFocus
                                    Exit Sub
                                Else
                                    cmdIncluirCTC.Enabled = True
                                End If
                            Else
                                'subst.tributarui = N ; Pega valor tributado
                                lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotalbruto"), "##,###,##0.00")
                                If de_informa.rsSel_CTC.Fields("fretetotalbruto") < 0 Then
                                    MsgBox "Valor de Frete deste CTC Inválido !"
                                    txtCTC.SetFocus
                                    Exit Sub
                                Else
                                    cmdIncluirCTC.Enabled = True
                                End If
                            End If
                        Else
                            'frete liq e bruto iguais
                            lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotal"), "##,###,##0.00")
                            If de_informa.rsSel_CTC.Fields("fretetotal") < 0 Then
                                MsgBox "Valor de Frete deste CTC Inválido !"
                                txtCTC.SetFocus
                                Exit Sub
                            Else
                                cmdIncluirCTC.Enabled = True
                            End If
                        End If
                    Else
                        'CTC do Informa
                        If de_informa.rsSel_CTC.Fields("subtrib") = "S" Then
                            lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotal"), "##,###,##0.00")
                            If de_informa.rsSel_CTC.Fields("fretetotal") < 0 Then
                                MsgBox "Valor de Frete deste CTC Inválido !"
                                txtCTC.SetFocus
                                Exit Sub
                            Else
                                cmdIncluirCTC.Enabled = True
                                cmdIncluirCTC.SetFocus
                            End If
                        Else
                            lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotalbruto"), "##,###,##0.00")
                            If de_informa.rsSel_CTC.Fields("fretetotalbruto") < 0 Then
                                MsgBox "Valor de Frete deste CTC Inválido !", vbInformation, "Valor de Frete"
                                txtCTC.SetFocus
                                Exit Sub
                            Else
                                cmdIncluirCTC.Enabled = True
                                cmdIncluirCTC.SetFocus
                            End If
                        End If
                    End If
                End If
                
            Else
                'encontrou documento não fiscal
                MsgBox "Este Documento não é um CTC !! Este Documento não pode ser Faturado: " & de_informa.rsSel_CTC.Fields("tipodoc"), vbCritical, "Documento Não Fiscal"
                txtCTC.SetFocus
                Exit Sub
            End If
        End If
    ElseIf optNFS.Value = True Then
        
        If de_informa.rsSel_NFS.State = 1 Then de_informa.rsSel_NFS.Close
        de_informa.Sel_nfs TransFatur(txtFilial, txtCTC)
        
        If de_informa.rsSel_NFS.RecordCount < 1 Then
            'não encontrou
            MsgBox "Não Encontrado NF Serviço com este Número !", vbCritical, "Não Encontrado"
            txtCTC.SetFocus
            Exit Sub
        Else
            lblRemetente = de_informa.rsSel_NFS.Fields("cliente_nome")
            lblDestinatario = ""
            lblConsignatario = de_informa.rsSel_NFS.Fields("cliente_nome")
            lblEmissao = de_informa.rsSel_NFS.Fields("data")
            lblModal = ""
            If de_informa.rsSel_NFS.Fields("status") = "C" Then
                'ctc cancelado
                MsgBox "Esta NF Serviço Encontra-se Cancelado !!", vbCritical, "NFS Cancelado"
                txtCTC.SetFocus
                Exit Sub
            Else
                If Len(Trim$(de_informa.rsSel_NFS.Fields("faturanum"))) > 1 Then
                    'nfs Encontra-se Faturado
                    MsgBox "Esta NF Serviço Já Encontra-se Faturado !! Fatura: " & de_informa.rsSel_NFS.Fields("faturanum"), vbCritical, "NFS Faturado"
                    txtCTC.SetFocus
                    Exit Sub
                End If
                If de_informa.rsSel_NFS.Fields("cliente_cgc") <> zeros2(txtCnpj, 14) Then
                    'O Consignatario não é o Mesmo
                    Do While lblCartaCorr = ""
                        If MsgBox("Atenção! O Cliente desta NF Serviço é Diferente do Cliente Desta Fatura. Deseja, Mesmo Assim Incluir Esta NF Serviço Nesta Fatura ? (você terá que informar o número da Carta de Correção).", vbCritical + vbYesNo, "Cliente Diferente") = vbNo Then
                            txtCTC.SetFocus
                            Exit Sub
                        Else
                            lblCartaCorr = InputBox("Entre com o Número da Carta de Correção que Retifica o Cliente desta NFS:", "Número da Carta de Correção")
                            If lblCartaCorr = "" Then
                            Else
                                lblCartaCorr = "Carta de Correção (Cliente): " & lblCartaCorr
                            End If
                        End If
                    Loop
                End If
                'frete liq e bruto iguais
                lblFrete = Format(de_informa.rsSel_NFS.Fields("valornfsliquido"), "##,###,##0.00")
                
                'If de_informa.rsSel_NFS.Fields("valornfsliquido") < 1 Then
                '    MsgBox "O Valor desta NF Serviço é Inválido !"
                '    txtCTC.SetFocus
                '    Exit Sub
                'Else
                    cmdIncluirCTC.Enabled = True
                    cmdIncluirCTC.SetFocus
                'End If
                
            End If
        End If
    End If
End Sub
Private Sub cmdBuscaPreFat_Click()
    Dim xPreFat As String
    
    If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
    de_informa.Sel_PreFatura TransFatur(txtPreFilial, txtPreFatura)
    
    gridItens.DataMember = "sel_prefatura"
    gridItens.Refresh
    
    If de_informa.rsSel_PreFatura.RecordCount < 1 Then
        MsgBox "Número de Pré-Fatura Inexistente !", vbInformation, "Ops"
        Exit Sub
    End If
    
    If de_informa.rsSel_PreFatura.Fields("tipodoc") = "NFS" Then
        optNFS.Value = True
        optCTC.Enabled = False
    Else
        optCTC.Value = True
        optNFS.Enabled = False
    End If
    

    fraCliente.Enabled = False
    
    xPreFat = TransFatur(txtPreFilial, txtPreFatura)
    
    txtFilialFatura = Mid$(de_informa.rsSel_PreFatura.Fields("filialprefatura"), 1, 2)
    txtCnpj = de_informa.rsSel_PreFatura.Fields("cliente_cgc")
    lblCliente = de_informa.rsSel_PreFatura.Fields("cliente_nome")
    mskVencimento.Mask = ""
    mskVencimento.Text = de_informa.rsSel_PreFatura.Fields("vencimento")
    mskVencimento.Mask = "##/##/####"
    txtNumBanco = de_informa.rsSel_PreFatura.Fields("banco")
    lblNomeBanco = de_informa.rsSel_PreFatura.Fields("banconome")
    lblContaBanco = de_informa.rsSel_PreFatura.Fields("conta")
    txtEndCob = de_informa.rsSel_PreFatura.Fields("endcob")
    txtCepCob = de_informa.rsSel_PreFatura.Fields("cepcob")
    txtCidCob = de_informa.rsSel_PreFatura.Fields("cidadecob")
    txtUFCob = de_informa.rsSel_PreFatura.Fields("ufcob")
    txtFoneCob = de_informa.rsSel_PreFatura.Fields("telefonecob")
    txtContatoCob = de_informa.rsSel_PreFatura.Fields("contatocob")
    
    If de_informa.rsSel_PreFaturaTotais.State = 1 Then de_informa.rsSel_PreFaturaTotais.Close
    de_informa.Sel_PreFaturaTotais TransFatur(txtPreFilial, txtPreFatura)
    
    If Not IsNull(de_informa.rsSel_PreFaturaTotais.Fields("fretetot")) Then
        lblTotalFatura.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
        lblTotalFaturaICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot"), "##,###,##0.00")
        lblQtdeCtc = Format(de_informa.rsSel_PreFaturaTotais.Fields("qtde"), "#,##0")
    Else
        lblTotalFatura.Caption = "0,00"
        lblTotalFaturaICMS.Caption = "0,00"
        lblQtdeCtc = "0"
    End If
    
    If de_informa.rsSel_PreFaturaTotais.Fields("qtde") > 0 Then
        cmdExcluirCTC.Enabled = True
    Else
        cmdExcluirCTC.Enabled = False
    End If
    
    cmdIncluirCTC.Enabled = False
    txtFilial = ""
    txtCTC = ""
    lblRemetente = ""
    lblDestinatario = ""
    lblConsignatario = ""
    lblEmissao = ""
    lblModal = ""
    lblFrete = ""
    lblCartaCorr = ""
    DoEvents
    txtFilial.SetFocus
    DoEvents

End Sub

Private Sub cmdCancCob_Click()
    fraCliente.Enabled = True
    fraFatura.Enabled = True
    cmdAlterarCob.Enabled = True
    cmdGravarCob.Enabled = False
    cmdCancCob.Enabled = False
    cmdEndPadraoCob.Enabled = False
    txtEndCob.Enabled = False
    txtEndCob.BackColor = xbranco
    txtCepCob.Enabled = False
    txtCepCob.BackColor = xbranco
    txtCidCob.Enabled = False
    txtCidCob.BackColor = xbranco
    txtUFCob.Enabled = False
    txtUFCob.BackColor = xbranco
    txtFoneCob.Enabled = False
    txtFoneCob.BackColor = xbranco
    txtContatoCob.Enabled = False
    txtContatoCob.BackColor = xbranco
    txtCnpj_LostFocus
End Sub

Private Sub cmdEndPadraoCob_Click()
    txtEndCob = de_informa.rsSel_CadCliCGC.Fields("endereco")
    txtCepCob = de_informa.rsSel_CadCliCGC.Fields("cep")
    txtCidCob = de_informa.rsSel_CadCliCGC.Fields("cidade")
    txtUFCob = de_informa.rsSel_CadCliCGC.Fields("uf")
End Sub
Private Sub cmdExcluirCTC_Click()

    If MsgBox("Confirma a Exclusão do CTC/NF: " & Mid$(gridItens.Columns(1), 1, 2) & "-" & Mid$(gridItens.Columns(1), 3, 8) & " Desta Pré-Fatura ?", vbQuestion + vbYesNo, "Confirma Exclusão") = vbYes Then
    
        de_informa.cn_informa.BeginTrans
        
        de_informa.Excl_PreFatCTC TransFatur(txtPreFilial, txtPreFatura), gridItens.Columns(1)
        
        If (optCTC.Value = True And tabTipoFatura.Tab = 0) Or tabTipoFatura.Tab = 1 Or tabTipoFatura.Tab = 2 Then
            de_informa.Alt_CTCFaturado "", gridItens.Columns(1)
        ElseIf optNFS.Value = True And tabTipoFatura.Tab = 0 Then
            de_informa.Alt_NFSFaturado "", gridItens.Columns(1)
        End If
        
        de_informa.cn_informa.CommitTrans
            
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura TransFatur(txtPreFilial, txtPreFatura)
        
        gridItens.DataMember = "sel_prefatura"
        gridItens.Refresh
        
        If de_informa.rsSel_PreFatura.RecordCount > 0 And de_informa.rsSel_PreFatura.RecordCount <= 5 Then
            cmdGerarFat.Enabled = True
        Else
            cmdGerarFat.Enabled = False
        End If
        
        
        If de_informa.rsSel_PreFaturaTotais.State = 1 Then de_informa.rsSel_PreFaturaTotais.Close
        de_informa.Sel_PreFaturaTotais TransFatur(txtPreFilial, txtPreFatura)
        
        If Not IsNull(de_informa.rsSel_PreFaturaTotais.Fields("fretetot")) Then
            lblTotalFatura.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            lblTotalFaturaICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot"), "##,###,##0.00")
            lblQtdeCtc = Format(de_informa.rsSel_PreFaturaTotais.Fields("qtde"), "#,##0")
        Else
            lblTotalFatura.Caption = "0,00"
            lblTotalFaturaICMS.Caption = "0,00"
            lblQtdeCtc = "0"
        End If
        
        If de_informa.rsSel_PreFaturaTotais.Fields("qtde") > 0 Then
            cmdExcluirCTC.Enabled = True
            cmdExcluirTudo.Enabled = True
        Else
            cmdExcluirCTC.Enabled = False
            cmdExcluirTudo.Enabled = False
        End If
    
    End If
    
End Sub

Private Sub cmdExcluirTudo_Click()
    
    If MsgBox("Confirma a Exclusão desta Pré-Fatura (TODOS os CTCs desta Pré-Fatura) ?", vbQuestion + vbYesNo, "Confirma Exclusão") = vbYes Then
    
        de_informa.cn_informa.BeginTrans
        
        de_informa.Excl_PreFatTudo TransFatur(txtPreFilial, txtPreFatura)
        
        If (optCTC.Value = True And tabTipoFatura.Tab = 0) Or tabTipoFatura.Tab = 1 Or tabTipoFatura.Tab = 2 Then
            de_informa.Alt_LimpaFaturaCTC TransFatur(txtPreFilial, txtPreFatura)
        ElseIf optNFS.Value = True And tabTipoFatura.Tab = 0 Then
            de_informa.Alt_LimpaFaturaNFS TransFatur(txtPreFilial, txtPreFatura)
        End If
        
        de_informa.cn_informa.CommitTrans
        
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura TransFatur(txtPreFilial, txtPreFatura)
        
        gridItens.DataMember = "sel_prefatura"
        gridItens.Refresh
        
        If de_informa.rsSel_PreFatura.RecordCount > 0 And de_informa.rsSel_PreFatura.RecordCount <= 5 Then
            cmdGerarFat.Enabled = True
        Else
            cmdGerarFat.Enabled = False
        End If
        
        If de_informa.rsSel_PreFaturaTotais.State = 1 Then de_informa.rsSel_PreFaturaTotais.Close
        de_informa.Sel_PreFaturaTotais TransFatur(txtPreFilial, txtPreFatura)
        
        If Not IsNull(de_informa.rsSel_PreFaturaTotais.Fields("fretetot")) Then
            lblTotalFatura.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            lblTotalFaturaICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot"), "##,###,##0.00")
            lblQtdeCtc = Format(de_informa.rsSel_PreFaturaTotais.Fields("qtde"), "#,##0")
        Else
            lblTotalFatura.Caption = "0,00"
            lblTotalFaturaICMS.Caption = "0,00"
            lblQtdeCtc = "0"
        End If
        
        If de_informa.rsSel_PreFaturaTotais.Fields("qtde") > 0 Then
            cmdExcluirCTC.Enabled = True
            cmdExcluirTudo.Enabled = True
        Else
            cmdExcluirCTC.Enabled = False
            cmdExcluirTudo.Enabled = False
        End If
    
    End If

End Sub

Private Sub cmdFecharFaturaAvulsa_Click()
    tabTipoFatura.TabEnabled(0) = True
    tabTipoFatura.TabEnabled(1) = True

    If Len(Trim$(txtCnpj)) < 11 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: CLIENTE !"
        txtCnpj.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mskVencimento) Then
        MsgBox "Data de Vencimento Inválida !"
        mskVencimento.SetFocus
        Exit Sub
    End If

    If Len(Trim$(txtNumBanco)) < 1 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: NÚMERO DO BANCO !"
        txtNumBanco.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtAvulsaDescr)) < 5 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: DESCRIÇÃO !"
        txtAvulsaDescr.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtEndCob)) < 5 Or Len(Trim$(txtCidCob)) < 3 Or Len(Trim$(txtCepCob)) < 5 Or Len(Trim$(txtFoneCob)) < 5 Then
        MsgBox "Dados de Endereço de Cobrança Inválidos !"
        cmdAlterarCob.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtFilialFatura)) < 2 Then
        MsgBox "Filial da Fatura Inválida !"
        txtFilialFatura.SetFocus
        Exit Sub
    End If

    If IsNumeric(SoNumeros(txtValorFatAvulsa)) Then
        If CDbl(SoNumeros(txtValorFatAvulsa)) < 1 Then
            MsgBox "Valor para Fatura Avulsa Inválido !"
            txtValorFatAvulsa.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Valor para Fatura Avulsa Inválido !"
        txtValorFatAvulsa.SetFocus
        Exit Sub
    End If
        
    frmGravaNovaFatura.lblPrefatura = "AVULSA"
    frmGravaNovaFatura.lblAvulsaDesc = txtAvulsaDescr
    
    frmGravaNovaFatura.lblClienteCNPJ.Caption = zeros2(txtCnpj, 14)
    frmGravaNovaFatura.lblCliente.Caption = lblCliente
    frmGravaNovaFatura.lblVencto.Caption = mskVencimento
    frmGravaNovaFatura.lblEmissor.Caption = xusuario
    frmGravaNovaFatura.lblBanco.Caption = zeros(txtNumBanco, 4) & "-" & lblNomeBanco
    frmGravaNovaFatura.lblConta.Caption = lblcontadoraBanco
    
    
    frmGravaNovaFatura.lblValorFaturaBrutoICMS.Caption = txtValorFatAvulsa
    frmGravaNovaFatura.lblValorFaturaBruto.Caption = txtValorFatAvulsa
    frmGravaNovaFatura.lblValorICMS.Caption = "0,00"
    frmGravaNovaFatura.lblValorFatura.Caption = txtValorFatAvulsa
    
    frmGravaNovaFatura.mskEmissao.Mask = ""
    frmGravaNovaFatura.mskEmissao.Text = datahora("DATA")
    frmGravaNovaFatura.mskEmissao.Mask = "##/##/####"
    
    frmGravaNovaFatura.Show 1
    
End Sub

Private Sub cmdGerarFat_Click()
    If MsgBox("Você Confirma a Geração de FATURA para a Pré-Fatura Número " & TransFatur(Trim$(txtPreFilial), Trim$(txtPreFatura)) & " ?", vbYesNo + vbQuestion, "Confirma") = vbYes Then
            
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura TransFatur(Trim$(txtPreFilial), Trim$(txtPreFatura))
        
        If de_informa.rsSel_PreFatura.RecordCount < 1 Then
            MsgBox "Esta Pré-Fatura não Consta mais no Banco de Dados. Provavelmente Alguém já a Faturou ou Excluiu !"
            Exit Sub
        Else
            
            frmGravaNovaFatura.lblPrefatura = TransFatur(Trim$(txtPreFilial), Trim$(txtPreFatura))
            frmGravaNovaFatura.lblClienteCNPJ.Caption = de_informa.rsSel_PreFatura.Fields("cliente_cgc")
            frmGravaNovaFatura.lblCliente.Caption = de_informa.rsSel_PreFatura.Fields("cliente_nome")
            frmGravaNovaFatura.lblVencto.Caption = de_informa.rsSel_PreFatura.Fields("vencimento")
            frmGravaNovaFatura.lblEmissor.Caption = de_informa.rsSel_PreFatura.Fields("emissor")
            frmGravaNovaFatura.lblBanco.Caption = zeros(de_informa.rsSel_PreFatura.Fields("banco"), 4) & "-" & de_informa.rsSel_PreFatura.Fields("banconome")
            frmGravaNovaFatura.lblConta.Caption = de_informa.rsSel_PreFatura.Fields("conta")
            
            If de_informa.rsSel_PreFaturaTotais.State = 1 Then de_informa.rsSel_PreFaturaTotais.Close
            de_informa.Sel_PreFaturaTotais TransFatur(Trim$(txtPreFilial), Trim$(txtPreFatura))
            
            frmGravaNovaFatura.lblValorFaturaBrutoICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot"), "##,###,##0.00")
            frmGravaNovaFatura.lblValorFaturaBruto.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            frmGravaNovaFatura.lblValorICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot") - de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            frmGravaNovaFatura.lblValorFatura.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            
            frmGravaNovaFatura.mskEmissao.Mask = ""
            frmGravaNovaFatura.mskEmissao.Text = datahora("DATA")
            frmGravaNovaFatura.mskEmissao.Mask = "##/##/####"
            
            frmGravaNovaFatura.Show 1
            
            If cmdGerarFat.Caption = "FATURADO" Then
                cmdNovaPreFat_Click
            End If
            
        End If
        
    End If

End Sub

Private Sub cmdGravarCob_Click()
    
    de_informa.Alt_CadCliDadosCob Trim$(txtEndCob), Trim$(txtCepCob), Trim$(txtCidCob), Trim$(txtUFCob), _
                                  Trim$(txtFoneCob), Trim$(txtContatoCob), zeros2(txtCnpj, 14)
    
    MsgBox "Dados Gravados !", vbInformation, "Gravação"

    fraCliente.Enabled = True
    fraFatura.Enabled = True
    cmdAlterarCob.Enabled = True
    cmdGravarCob.Enabled = False
    cmdCancCob.Enabled = False
    cmdEndPadraoCob.Enabled = False
    txtEndCob.Enabled = False
    txtEndCob.BackColor = xbranco
    txtCepCob.Enabled = False
    txtCepCob.BackColor = xbranco
    txtCidCob.Enabled = False
    txtCidCob.BackColor = xbranco
    txtUFCob.Enabled = False
    txtUFCob.BackColor = xbranco
    txtFoneCob.Enabled = False
    txtFoneCob.BackColor = xbranco
    txtContatoCob.Enabled = False
    txtContatoCob.BackColor = xbranco
    txtCnpj_LostFocus
End Sub
Private Sub cmdIncluirCTC_Click()
    Dim xcontador As Integer, xFreteTot As Currency, xfreteTotICMS As Currency, xPreFat As String, xDocto As String
    Dim xCon As New ADODB.Connection
    Dim xrs As New ADODB.Recordset

    fraCliente.Enabled = False
    Frame1.Enabled = False
    tabTipoFatura.Enabled = False
    
    If Len(Trim$(txtPreFatura.Text)) = 0 Then
        xCon.ConnectionString = xstrcon2
        xCon.ConnectionTimeout = 30
        xCon.Open
        xrs.Open "exec sp_prefat '" & frmNovaFatura.txtFilialFatura & "'", xCon, adOpenStatic, adLockBatchOptimistic
        xPreFat = xrs.Fields(0)
        xrs.Close
        xCon.Close
        
        txtPreFilial = Mid$(txtFilialFatura, 1, 2)
        txtPreFatura = zeros2(xPreFat, 6)
        
    End If
    
    xPreFat = TransFatur(txtPreFilial, txtPreFatura)
    
    If optCTC.Value = True Then
        xDocto = transctc(txtFilial, txtCTC)
    Else
        xDocto = TransFatur(txtFilial, txtCTC)
    End If
    
    If de_informa.rsSel_PreFaturaCTC.State = 1 Then de_informa.rsSel_PreFaturaCTC.Close
    de_informa.Sel_PrefaturaCTC TransFatur(txtPreFilial, txtPreFatura), xDocto
    
    If de_informa.rsSel_PreFaturaCTC.RecordCount > 0 Then
        MsgBox "Este CTC/NFS Já Consta na Tabela Abaixo !", vbInformation, "Ops"
        Frame1.Enabled = True
        tabTipoFatura.Enabled = True
        tabTipoFatura.TabEnabled(0) = True
        tabTipoFatura.TabEnabled(1) = False
        tabTipoFatura.TabEnabled(2) = False
        tabTipoFatura.TabEnabled(3) = False
        Exit Sub
    End If
    
    If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
    de_informa.Sel_CadCliCGC zeros2(txtCnpj, 14)
    
    de_informa.cn_informa.BeginTrans
    
    If optCTC.Value = True Then
    
        de_informa.Ins_PreFatura de_informa.rsSel_CTC.Fields("tipodoc"), de_informa.rsSel_CTC.Fields("filialctc"), _
                                 de_informa.rsSel_CTC.Fields("data"), CDbl(SoNumeros(lblFrete)) / 100, de_informa.rsSel_CTC.Fields("fretetotalbruto"), _
                                 de_informa.rsSel_CTC.Fields("remet_cgc"), de_informa.rsSel_CTC.Fields("remet_nome"), lblCartaCorr, _
                                 xPreFat, xusuario, CDate(mskVencimento), zeros2(txtCnpj, 14), lblCliente, de_informa.rsSel_CadCliCGC.Fields("ie"), _
                                 de_informa.rsSel_CadCliCGC.Fields("endereco"), de_informa.rsSel_CadCliCGC.Fields("cidade"), _
                                 de_informa.rsSel_CadCliCGC.Fields("uf"), de_informa.rsSel_CadCliCGC.Fields("cep"), _
                                 de_informa.rsSel_CadCliCGC.Fields("endcob"), de_informa.rsSel_CadCliCGC.Fields("cidadecob"), _
                                 de_informa.rsSel_CadCliCGC.Fields("ufcob"), de_informa.rsSel_CadCliCGC.Fields("cepcob"), _
                                 de_informa.rsSel_CadCliCGC.Fields("telefonecob"), de_informa.rsSel_CadCliCGC.Fields("contatocob"), _
                                 CDbl(txtNumBanco), Trim$(lblNomeBanco), Trim$(lblContaBanco), "N"
                                 
        de_informa.Alt_CTCFaturado xPreFat, de_informa.rsSel_CTC.Fields("filialctc")
        
        optNFS.Enabled = False
        
    
    ElseIf optNFS.Value = True Then
    
        de_informa.Ins_PreFatura "NFS", de_informa.rsSel_NFS.Fields("filialnfs"), _
                                 de_informa.rsSel_NFS.Fields("data"), CDbl(SoNumeros(lblFrete)) / 100, de_informa.rsSel_NFS.Fields("valornfs"), _
                                 de_informa.rsSel_NFS.Fields("cliente_cgc"), de_informa.rsSel_NFS.Fields("cliente_nome"), lblCartaCorr, _
                                 xPreFat, xusuario, CDate(mskVencimento), zeros2(txtCnpj, 14), lblCliente, de_informa.rsSel_CadCliCGC.Fields("ie"), _
                                 de_informa.rsSel_CadCliCGC.Fields("endereco"), de_informa.rsSel_CadCliCGC.Fields("cidade"), _
                                 de_informa.rsSel_CadCliCGC.Fields("uf"), de_informa.rsSel_CadCliCGC.Fields("cep"), _
                                 de_informa.rsSel_CadCliCGC.Fields("endcob"), de_informa.rsSel_CadCliCGC.Fields("cidadecob"), _
                                 de_informa.rsSel_CadCliCGC.Fields("ufcob"), de_informa.rsSel_CadCliCGC.Fields("cepcob"), _
                                 de_informa.rsSel_CadCliCGC.Fields("telefonecob"), de_informa.rsSel_CadCliCGC.Fields("contatocob"), _
                                 CDbl(txtNumBanco), Trim$(lblNomeBanco), Trim$(lblContaBanco), "N"
                                 
        de_informa.Alt_NFSFaturado xPreFat, de_informa.rsSel_NFS.Fields("filialnfs")
        
        optCTC.Enabled = False
        
    End If
        
    de_informa.cn_informa.CommitTrans
    
    If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
    de_informa.Sel_PreFatura TransFatur(txtPreFilial, txtPreFatura)
    
    gridItens.DataMember = "sel_prefatura"
    gridItens.Refresh
    
    If de_informa.rsSel_PreFatura.RecordCount > 0 And de_informa.rsSel_PreFatura.RecordCount <= 5 Then
        cmdGerarFat.Enabled = True
    Else
        cmdGerarFat.Enabled = False
    End If
    
    If de_informa.rsSel_PreFaturaTotais.State = 1 Then de_informa.rsSel_PreFaturaTotais.Close
    de_informa.Sel_PreFaturaTotais TransFatur(txtPreFilial, txtPreFatura)
    
    If Not IsNull(de_informa.rsSel_PreFaturaTotais.Fields("fretetot")) Then
        lblTotalFatura.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
        lblTotalFaturaICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot"), "##,###,##0.00")
        lblQtdeCtc = Format(de_informa.rsSel_PreFaturaTotais.Fields("qtde"), "#,##0")
    Else
        lblTotalFatura.Caption = "0,00"
        lblTotalFaturaICMS.Caption = "0,00"
        lblQtdeCtc = "0"
    End If
    
    If de_informa.rsSel_PreFaturaTotais.Fields("qtde") > 0 Then
        cmdExcluirCTC.Enabled = True
        cmdExcluirTudo.Enabled = True
    Else
        cmdExcluirCTC.Enabled = False
        cmdExcluirTudo.Enabled = False
    End If
    
    cmdIncluirCTC.Enabled = False
    txtFilial = ""
    txtCTC = ""
    lblRemetente = ""
    lblDestinatario = ""
    lblConsignatario = ""
    lblEmissao = ""
    lblModal = ""
    lblFrete = ""
    lblCartaCorr = ""
    Frame1.Enabled = True
    tabTipoFatura.Enabled = True
    tabTipoFatura.TabEnabled(0) = True
    tabTipoFatura.TabEnabled(1) = False
    tabTipoFatura.TabEnabled(2) = False
    tabTipoFatura.TabEnabled(3) = False
    DoEvents
    
    If txtFilial.Enabled = True Then
         txtFilial.SetFocus
    End If
    
End Sub
Private Sub cmdIncluirCTCsIntervalo_Click()
    Dim xcontador As Integer, xFreteTot As Currency, xfreteTotICMS As Currency, xuf_orig As String
    Dim xFrete As Currency, xPreFat As String
    Dim xCon As New ADODB.Connection
    Dim xrs As New ADODB.Recordset
    
    If Len(Trim$(txtCnpj)) < 11 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: CLIENTE !"
        txtCnpj.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mskVencimento) Then
        MsgBox "Data de Vencimento Inválida !"
        mskVencimento.SetFocus
        Exit Sub
    End If

    If Len(Trim$(txtNumBanco)) < 1 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: NÚMERO DO BANCO !"
        txtNumBanco.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtEndCob)) < 5 Or Len(Trim$(txtCidCob)) < 3 Or Len(Trim$(txtCepCob)) < 5 Or Len(Trim$(txtFoneCob)) < 5 Then
        MsgBox "Dados de Endereço de Cobrança Inválidos !"
        cmdAlterarCob.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtFilialFatura)) < 2 Then
        MsgBox "Filial da Fatura Inválida !"
        txtFilialFatura.SetFocus
        Exit Sub
    End If
    
    If de_informa.rsSel_CTCsIntervalo.State = 1 Then de_informa.rsSel_CTCsIntervalo.Close
    de_informa.Sel_CTCsIntervalo txtCnpj, Trim$(txtRemetInt) & "%", txtFilialInt, CDbl(txtCtc1Int), CDbl(txtCtc2Int)
    
    If de_informa.rsSel_CTCsIntervalo.RecordCount < 1 Then
    
        MsgBox "Não Há CTCs Não Faturados Para as Opções Solicitadas !", vbCritical, "Erro"
        txtFilialInt.SetFocus
        Exit Sub
        
    Else
    
        Frame1.Enabled = False
        fraCliente.Enabled = False
        tabTipoFatura.Enabled = False
        tabTipoFatura.TabEnabled(0) = False
        tabTipoFatura.TabEnabled(1) = False
        tabTipoFatura.TabEnabled(2) = False
        tabTipoFatura.TabEnabled(3) = False
        
        If Len(Trim$(txtPreFatura)) = 0 Then
            
            xCon.ConnectionString = xstrcon2
            xCon.ConnectionTimeout = 30
            xCon.Open
            xrs.Open "exec sp_prefat '" & frmNovaFatura.txtFilialFatura & "'", xCon, adOpenStatic, adLockBatchOptimistic
            xPreFat = xrs.Fields(0)
            xrs.Close
            xCon.Close
            txtPreFilial = Mid$(txtFilialFatura, 1, 2)
            txtPreFatura = zeros2(xPreFat, 6)
    
        End If
    
        xPreFat = TransFatur(txtPreFilial, txtPreFatura)
        
        de_informa.cn_informa.BeginTrans
    
        Do Until de_informa.rsSel_CTCsIntervalo.EOF
            
            If IsNull(de_informa.rsSel_CTCsIntervalo.Fields("subtrib")) Then
                'faturando CTCs do SITLA
                If de_informa.rsSel_CTCsIntervalo.Fields("fretetotal") <> de_informa.rsSel_CTCsIntervalo.Fields("fretetotalbruto") Then
                    If MsgBox("(SITLA) Este CTC, " & de_informa.rsSel_CTCsIntervalo.Fields("filialctc") & ", Caracteriza a Regra de Substituição Tributária ? (Confira no CTC. Se Positivo Será Cobrado Pelo Valor Líquido.)", vbQuestion + vbYesNo, "CTC do SITLA") = vbYes Then
                        'subst.tributarui = S ; Pega valor pelo liquido
                        xFrete = de_informa.rsSel_CTCsIntervalo.Fields("fretetotal")
                        If de_informa.rsSel_CTCsIntervalo.Fields("fretetotal") < 0 Then
                            MsgBox "Valor de Frete deste CTC Inválido: " & de_informa.rsSel_CTCsIntervalo.Fields("filialctc") & " !", vbCritical
                            Frame1.Enabled = True
                            tabTipoFatura.Enabled = True
                            tabTipoFatura.TabEnabled(0) = False
                            tabTipoFatura.TabEnabled(1) = True
                            tabTipoFatura.TabEnabled(2) = False
                            tabTipoFatura.TabEnabled(3) = False
                            txtCtc1Int.SetFocus
                            Exit Sub
                        Else
                            cmdIncluirCTC.Enabled = True
                            'cmdIncluirCTC.SetFocus
                        End If
                    Else
                        'subst.tributarui = N ; Pega valor tributado
                        xFrete = de_informa.rsSel_CTCsIntervalo.Fields("fretetotalbruto")
                        If de_informa.rsSel_CTCsIntervalo.Fields("fretetotalbruto") < 0 Then
                            MsgBox "Valor de Frete deste CTC Inválido: " & de_informa.rsSel_CTCsIntervalo.Fields("filialctc") & " !", vbCritical
                            Frame1.Enabled = True
                            tabTipoFatura.Enabled = True
                            tabTipoFatura.TabEnabled(0) = False
                            tabTipoFatura.TabEnabled(1) = True
                            tabTipoFatura.TabEnabled(2) = False
                            tabTipoFatura.TabEnabled(3) = False
                            txtCtc1Int.SetFocus
                            Exit Sub
                        Else
                            cmdIncluirCTC.Enabled = True
                        End If
                    End If
                Else
                    'frete liq e bruto iguais
                    xFrete = de_informa.rsSel_CTCsIntervalo.Fields("fretetotal")
                    If de_informa.rsSel_CTCsIntervalo.Fields("fretetotal") < 0 Then
                        MsgBox "Valor de Frete deste CTC Inválido: " & de_informa.rsSel_CTCsIntervalo.Fields("filialctc") & " !", vbCritical
                        Frame1.Enabled = True
                        tabTipoFatura.Enabled = True
                        tabTipoFatura.TabEnabled(0) = False
                        tabTipoFatura.TabEnabled(1) = True
                        tabTipoFatura.TabEnabled(2) = False
                        tabTipoFatura.TabEnabled(3) = False
                        txtCtc1Int.SetFocus
                        Exit Sub
                    Else
                        cmdIncluirCTC.Enabled = True
                    End If
                End If
            Else
                'CTC do Informa
                If de_informa.rsSel_CTCsIntervalo.Fields("subtrib") = "S" Then
                    xFrete = de_informa.rsSel_CTCsIntervalo.Fields("fretetotal")
                    If de_informa.rsSel_CTCsIntervalo.Fields("fretetotal") < 0 Then
                        MsgBox "Valor de Frete deste CTC Inválido: " & de_informa.rsSel_CTCsIntervalo.Fields("filialctc") & " !", vbCritical
                        Frame1.Enabled = True
                        tabTipoFatura.Enabled = True
                        tabTipoFatura.TabEnabled(0) = False
                        tabTipoFatura.TabEnabled(1) = True
                        tabTipoFatura.TabEnabled(2) = False
                        tabTipoFatura.TabEnabled(3) = False
                        txtCtc1Int.SetFocus
                        Exit Sub
                    Else
                        cmdIncluirCTC.Enabled = True
                    End If
                Else
                    xFrete = de_informa.rsSel_CTCsIntervalo.Fields("fretetotalbruto")
                    If de_informa.rsSel_CTCsIntervalo.Fields("fretetotalbruto") < 0 Then
                        MsgBox "Valor de Frete deste CTC Inválido: " & de_informa.rsSel_CTCsIntervalo.Fields("filialctc") & " !", vbCritical
                        Frame1.Enabled = True
                        tabTipoFatura.Enabled = True
                        tabTipoFatura.TabEnabled(0) = False
                        tabTipoFatura.TabEnabled(1) = True
                        tabTipoFatura.TabEnabled(2) = False
                        tabTipoFatura.TabEnabled(3) = False
                        txtCtc1Int.SetFocus
                        Exit Sub
                    Else
                        cmdIncluirCTC.Enabled = True
                    End If
                End If
            End If
            
            If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
            de_informa.Sel_CadCliCGC zeros2(txtCnpj, 14)
            
            de_informa.Ins_PreFatura de_informa.rsSel_CTCsIntervalo.Fields("tipodoc"), de_informa.rsSel_CTCsIntervalo.Fields("filialctc"), _
                                     de_informa.rsSel_CTCsIntervalo.Fields("data"), xFrete, de_informa.rsSel_CTCsIntervalo.Fields("fretetotalbruto"), _
                                     de_informa.rsSel_CTCsIntervalo.Fields("remet_cgc"), de_informa.rsSel_CTCsIntervalo.Fields("remet_nome"), "", _
                                     xPreFat, xusuario, CDate(mskVencimento), zeros2(txtCnpj, 14), lblCliente, de_informa.rsSel_CadCliCGC.Fields("ie"), _
                                     de_informa.rsSel_CadCliCGC.Fields("endereco"), de_informa.rsSel_CadCliCGC.Fields("cidade"), _
                                     de_informa.rsSel_CadCliCGC.Fields("uf"), de_informa.rsSel_CadCliCGC.Fields("cep"), _
                                     de_informa.rsSel_CadCliCGC.Fields("endcob"), de_informa.rsSel_CadCliCGC.Fields("cidadecob"), _
                                     de_informa.rsSel_CadCliCGC.Fields("ufcob"), de_informa.rsSel_CadCliCGC.Fields("cepcob"), _
                                     de_informa.rsSel_CadCliCGC.Fields("telefonecob"), de_informa.rsSel_CadCliCGC.Fields("contatocob"), _
                                     CDbl(txtNumBanco), Trim$(lblNomeBanco), Trim$(lblContaBanco), "N"
                             
            de_informa.Alt_CTCFaturado xPreFat, de_informa.rsSel_CTCsIntervalo.Fields("filialctc")
            
            de_informa.rsSel_CTCsIntervalo.MoveNext

        Loop
        
        de_informa.cn_informa.CommitTrans
    
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura TransFatur(txtPreFilial, txtPreFatura)
        
        gridItens.DataMember = "sel_prefatura"
        gridItens.Refresh
        
        If de_informa.rsSel_PreFaturaTotais.State = 1 Then de_informa.rsSel_PreFaturaTotais.Close
        de_informa.Sel_PreFaturaTotais TransFatur(txtPreFilial, txtPreFatura)
        
        If Not IsNull(de_informa.rsSel_PreFaturaTotais.Fields("fretetot")) Then
            lblTotalFatura.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            lblTotalFaturaICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot"), "##,###,##0.00")
            lblQtdeCtc = Format(de_informa.rsSel_PreFaturaTotais.Fields("qtde"), "#,##0")
        Else
            lblTotalFatura.Caption = "0,00"
            lblTotalFaturaICMS.Caption = "0,00"
            lblQtdeCtc = "0"
        End If
        
        If de_informa.rsSel_PreFaturaTotais.Fields("qtde") > 0 Then
            cmdExcluirCTC.Enabled = True
            cmdExcluirTudo.Enabled = True
        Else
            cmdExcluirCTC.Enabled = False
            cmdExcluirTudo.Enabled = False
        End If
        
        Frame1.Enabled = True
        tabTipoFatura.Enabled = True
        tabTipoFatura.TabEnabled(0) = False
        tabTipoFatura.TabEnabled(1) = True
        tabTipoFatura.TabEnabled(2) = False
        tabTipoFatura.TabEnabled(3) = False
        
        cmdIncluirCTC.Enabled = False
        txtFilial = ""
        txtCTC = ""
        lblRemetente = ""
        lblDestinatario = ""
        lblConsignatario = ""
        lblEmissao = ""
        lblModal = ""
        lblFrete = ""
        lblCartaCorr = ""
        txtFilial.SetFocus
    
    End If
    
End Sub
Private Sub cmdProcessarArqPreFatCliente_Click()
    Dim xcontador As Integer, xFreteTot As Currency, xfreteTotICMS As Currency, xuf_orig As String
    Dim xFrete As Currency, xPreFat As String, xFilialCtc As String, xFretePreFat As Variant, xlin As Long
    Dim xCon As New ADODB.Connection
    Dim xrs As New ADODB.Recordset
    
    If Len(Trim$(txtCnpj)) < 11 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: CLIENTE !"
        txtCnpj.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mskVencimento) Then
        MsgBox "Data de Vencimento Inválida !"
        mskVencimento.SetFocus
        Exit Sub
    End If

    If Len(Trim$(txtNumBanco)) < 1 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: NÚMERO DO BANCO !"
        txtNumBanco.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtEndCob)) < 5 Or Len(Trim$(txtCidCob)) < 3 Or Len(Trim$(txtCepCob)) < 5 Or Len(Trim$(txtFoneCob)) < 5 Then
        MsgBox "Dados de Endereço de Cobrança Inválidos !"
        cmdAlterarCob.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtFilialFatura)) < 2 Then
        MsgBox "Filial da Fatura Inválida !"
        txtFilialFatura.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtFilialPreFat)) < 2 Then
        MsgBox "Filial dos CTCs da Pré-Fatura Inválida !"
        txtFilialPreFat.SetFocus
        Exit Sub
    End If
    
    'abrir o arquivo TXT (pré-fatura)
    If optPFVideolar.Value = True Then
    
        If Dir(txtPFArquivo.Text) = "" Then
            MsgBox "Erro ! Arquivo Não Encontrado !", vbCritical
            txtPFArquivo.SetFocus
            Exit Sub
        End If

        Open txtPFArquivo.Text For Input As #1
        Line Input #1, xlinha
        
        'checar se este arquivo é mesmo da Videolar
        'checando se o código da transportadora confere
        If Mid$(xlinha, 1, 1) = "1" Then
            If Mid$(xlinha, 22, 10) <> "3000000164" Then
                MsgBox "Erro ! Identificação de Transportadora Errado (diferente de 3000000164). Procure Suporte Técnico !", vbCritical
                txtPFArquivo.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Erro ! Arquivo Não Identificado. Procure Suporte Técnico !", vbCritical
            txtPFArquivo.SetFocus
            Exit Sub
        End If
        
        Do Until EOF(1)
            Line Input #1, xlinha
            xlin = xlin + 1
        Loop
        
        bar1.Max = xlin
        
        xlin = 0
        
        'verificar os botões a serem travados
        
        Frame1.Enabled = False
        fraCliente.Enabled = False
        tabTipoFatura.Enabled = False
        tabTipoFatura.TabEnabled(0) = False
        tabTipoFatura.TabEnabled(1) = False
        tabTipoFatura.TabEnabled(2) = False
        tabTipoFatura.TabEnabled(3) = False
        
        If Len(Trim$(txtPreFatura)) = 0 Then
            
            xCon.ConnectionString = xstrcon2
            xCon.ConnectionTimeout = 30
            xCon.Open
            xrs.Open "exec sp_prefat '" & frmNovaFatura.txtFilialFatura & "'", xCon, adOpenStatic, adLockBatchOptimistic
            xPreFat = xrs.Fields(0)
            xrs.Close
            xCon.Close
            txtPreFilial = Mid$(txtFilialFatura, 1, 2)
            txtPreFatura = zeros2(xPreFat, 6)
    
        End If
    
        xPreFat = TransFatur(txtPreFilial, txtPreFatura)
        Close #1
        Open txtPFArquivo.Text For Input As #1
        
        de_informa.cn_informa.BeginTrans
     
            Do Until EOF(1)
            
                xlin = xlin + 1
                bar1.Value = xlin - 1
                lblTotalCtcsPreFat = zeros(xlin - 1, 5) & "/" & zeros(bar1.Max, 5)
                DoEvents
                Line Input #1, xlinha
            
                If xlin = 1 And Mid$(xlinha, 1, 1) = "1" Then
                
                    xnum_prefatura = Mid(xlinha, 82, 10)
                    xvalor_prefatura = CDbl(Mid(xlinha, 92, 17)) / 100
                    
                ElseIf xlin > 1 And Mid$(xlinha, 1, 1) = "2" Then
                
                    'identificar o número do ctc e transformá-lo em xfilialctc
                    xFilialCtc = Trim$(Mid$(xlinha, 2, 10))
                    xFilialCtc = transctc(txtFilialPreFat, xFilialCtc)
                    xFretePreFat = CDbl(Mid(xlinha, 17, 17)) / 100
                
                    If de_informa.rsSel_CTC.State = 1 Then de_informa.rsSel_CTC.Close
                    de_informa.Sel_CTC xFilialCtc
                    
                    Do While True
                        If de_informa.rsSel_CTC.RecordCount < 1 Then
                            If MsgBox("ATENÇÃO ! O CTC Número " & xFilialCtc & " não foi encontrado no Sistema. Foi atribuido a filial número " & txtFilialPreFat & " para este CTC. Deseja informar ou outro número de filial para tentar uma nova busca ?", vbYesNo, "CTC Não Encontrado") = vbYes Then
                                xfilial = InputBox("Entre com o Número da Filial: ", "CTC Não Encontrado")
                                If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
                                de_informa.Sel_CadFilial xfilial
                                If de_informa.rsSel_CadFilial.RecordCount < 1 Then
                                    MsgBox "Número de Filial Inválida !", vbCritical
                                Else
                                    txtFilialPreFat = xfilial
                                    xFilialCtc = transctc(txtFilialPreFat, Mid$(xFilialCtc, 3))
                                    If de_informa.rsSel_CTC.State = 1 Then de_informa.rsSel_CTC.Close
                                    de_informa.Sel_CTC xFilialCtc
                                End If
                            Else
                                de_informa.cn_informa.RollbackTrans
                                Close #1
                                MsgBox "Processo Cancelado !", vbCritical
                                Frame1.Enabled = True
                                tabTipoFatura.Enabled = True
                                tabTipoFatura.TabEnabled(0) = False
                                tabTipoFatura.TabEnabled(1) = False
                                tabTipoFatura.TabEnabled(2) = True
                                tabTipoFatura.TabEnabled(3) = False
                                
                                txtPFArquivo.SetFocus
                                Exit Sub
                            End If
                        ElseIf de_informa.rsSel_CTC.RecordCount > 0 And Mid$(de_informa.rsSel_CTC.Fields("respons_cgc"), 1, 8) <> "04229761" Then
                            If MsgBox("ATENÇÃO ! O CTC Número " & xFilialCtc & " Não está com responsável pelo frete a Videolar. Foi atribuido a filial número " & txtFilialPreFat & " para este CTC. Deseja informar ou outro número de filial para tentar uma nova busca ?", vbYesNo, "Responsável Diferente") = vbYes Then
                                xfilial = InputBox("Entre com o Número da Filial: ", "Responsável Diferente")
                                If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
                                de_informa.Sel_CadFilial xfilial
                                If de_informa.rsSel_CadFilial.RecordCount < 1 Then
                                    MsgBox "Número de Filial Inválida !", vbCritical
                                Else
                                    txtFilialPreFat = xfilial
                                    xFilialCtc = transctc(txtFilialPreFat, Mid(xFilialCtc, 3))
                                    If de_informa.rsSel_CTC.State = 1 Then de_informa.rsSel_CTC.Close
                                    de_informa.Sel_CTC xFilialCtc
                                End If
                            Else
                                de_informa.cn_informa.RollbackTrans
                                Close #1
                                MsgBox "Processo Cancelado !", vbCritical
                                Frame1.Enabled = True
                                tabTipoFatura.Enabled = True
                                tabTipoFatura.TabEnabled(0) = False
                                tabTipoFatura.TabEnabled(1) = False
                                tabTipoFatura.TabEnabled(2) = True
                                tabTipoFatura.TabEnabled(3) = False
                                txtPFArquivo.SetFocus
                                Exit Sub
                            End If
                        ElseIf de_informa.rsSel_CTC.RecordCount > 0 And Mid$(de_informa.rsSel_CTC.Fields("respons_cgc"), 1, 8) = "04229761" And _
                               Abs(de_informa.rsSel_CTC.Fields("fretefinal") - xFretePreFat) > 0.1 Then
                            If MsgBox("ATENÇÃO ! No CTC Número " & xFilialCtc & " o valor de frete não está conferindo com o que está no sistema. Foi atribuido a filial número " & txtFilialPreFat & " para este CTC. Deseja informar ou outro número de filial para tentar uma nova busca ?", vbYesNo, "Frete Diferente") = vbYes Then
                                xfilial = InputBox("Entre com o Número da Filial: ", "Frete Diferente")
                                If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
                                de_informa.Sel_CadFilial xfilial
                                If de_informa.rsSel_CadFilial.RecordCount < 1 Then
                                    MsgBox "Número de Filial Inválida !", vbCritical
                                Else
                                    txtFilialPreFat = xfilial
                                    xFilialCtc = transctc(txtFilialPreFat, Mid(xFilialCtc, 3))
                                    If de_informa.rsSel_CTC.State = 1 Then de_informa.rsSel_CTC.Close
                                    de_informa.Sel_CTC xFilialCtc
                                End If
                            Else
                                de_informa.cn_informa.RollbackTrans
                                Close #1
                                MsgBox "Processo Cancelado !", vbCritical
                                Frame1.Enabled = True
                                tabTipoFatura.Enabled = True
                                tabTipoFatura.TabEnabled(0) = False
                                tabTipoFatura.TabEnabled(1) = False
                                tabTipoFatura.TabEnabled(2) = True
                                tabTipoFatura.TabEnabled(3) = False
                                txtPFArquivo.SetFocus
                                Exit Sub
                            End If
                        ElseIf Len(Trim$(de_informa.rsSel_CTC.Fields("faturanum"))) > 0 Then
                            If MsgBox("ATENÇÃO ! No CTC Número " & xFilialCtc & " já consta como faturado. Pode ser que o número da Filial Esteja Incorreto. Confira este CTC e verifique o que aconteceu. Caso seja o número da filial que está errado, você pode ajustar este número. Deseja Ajustar o número da filial deste CTC ? ", vbYesNo, "Frete Diferente") = vbYes Then
                                xfilial = InputBox("Entre com o Número da Filial: ", "Filial Diferente Diferente")
                                If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
                                de_informa.Sel_CadFilial xfilial
                                If de_informa.rsSel_CadFilial.RecordCount < 1 Then
                                    MsgBox "Número de Filial Inválida !", vbCritical
                                Else
                                    txtFilialPreFat = xfilial
                                    xFilialCtc = transctc(txtFilialPreFat, Mid(xFilialCtc, 3))
                                    If de_informa.rsSel_CTC.State = 1 Then de_informa.rsSel_CTC.Close
                                    de_informa.Sel_CTC xFilialCtc
                                End If
                            Else
                                de_informa.cn_informa.RollbackTrans
                                Close #1
                                MsgBox "Processo Cancelado !", vbCritical
                                Frame1.Enabled = True
                                tabTipoFatura.Enabled = True
                                tabTipoFatura.TabEnabled(0) = False
                                tabTipoFatura.TabEnabled(1) = False
                                tabTipoFatura.TabEnabled(2) = True
                                tabTipoFatura.TabEnabled(3) = False
                                txtPFArquivo.SetFocus
                                Exit Sub
                            End If
                        
                            
'                            MsgBox "ATENÇÃO ! No CTC Número " & xFilialCtc & " encontra-se Faturado ou em Pré-Fatura da Intec (ou então com instrução de Não-Faturável). Fatura Num: " & de_informa.rsSel_CTC.Fields("faturanum") & ".", vbCritical
'                            de_informa.cn_informa.RollbackTrans
'                            Close #1
'                            MsgBox "Processo Cancelado !", vbCritical
'                            Frame1.Enabled = True
'                            tabTipoFatura.Enabled = True
'                            tabTipoFatura.TabEnabled(0) = False
 '                           tabTipoFatura.TabEnabled(1) = False
 '                           tabTipoFatura.TabEnabled(2) = True
 '                           tabTipoFatura.TabEnabled(3) = False
 '                           txtPFArquivo.SetFocus
 '                           Exit Sub
 
                        Else
                        
                            Exit Do
                            
                        End If
                        
                    Loop
                    
                    If de_informa.rsSel_CTC.Fields("subtrib") = "S" Then
                        xFrete = de_informa.rsSel_CTC.Fields("fretetotal")
                    Else
                        xFrete = de_informa.rsSel_CTC.Fields("fretetotalbruto")
                    End If
                    
                    If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
                    de_informa.Sel_CadCliCGC zeros2(txtCnpj, 14)
                    
                    If de_informa.rsSel_CTC.Fields("tem_ocorr") = "C" Then
                        MsgBox "ATENÇÃO ! O CTC " & de_informa.rsSel_CTC.Fields("filialctc") & " - (Frete: R$ " & Format(de_informa.rsSel_CTC.Fields("fretefinal"), "##,###,##0.00") & " Não poderá ser Incluso Nesta Pré-Fatura Pois o Mesmo Encontra-se CANCELADO !! Favor Anotar e Informar o Cliente que a Fatura não Constará Este CTC."
                    Else
                            de_informa.Ins_PreFatura de_informa.rsSel_CTC.Fields("tipodoc"), de_informa.rsSel_CTC.Fields("filialctc"), _
                                             de_informa.rsSel_CTC.Fields("data"), xFrete, de_informa.rsSel_CTC.Fields("fretetotalbruto"), _
                                             de_informa.rsSel_CTC.Fields("remet_cgc"), de_informa.rsSel_CTC.Fields("remet_nome"), "", _
                                             xPreFat, xusuario, CDate(mskVencimento), zeros2(txtCnpj, 14), lblCliente, de_informa.rsSel_CadCliCGC.Fields("ie"), _
                                             de_informa.rsSel_CadCliCGC.Fields("endereco"), de_informa.rsSel_CadCliCGC.Fields("cidade"), _
                                             de_informa.rsSel_CadCliCGC.Fields("uf"), de_informa.rsSel_CadCliCGC.Fields("cep"), _
                                             de_informa.rsSel_CadCliCGC.Fields("endcob"), de_informa.rsSel_CadCliCGC.Fields("cidadecob"), _
                                             de_informa.rsSel_CadCliCGC.Fields("ufcob"), de_informa.rsSel_CadCliCGC.Fields("cepcob"), _
                                             de_informa.rsSel_CadCliCGC.Fields("telefonecob"), de_informa.rsSel_CadCliCGC.Fields("contatocob"), _
                                             CDbl(txtNumBanco), Trim$(lblNomeBanco), Trim$(lblContaBanco), "N"
                           de_informa.Alt_CTCFaturado xPreFat, de_informa.rsSel_CTC.Fields("filialctc")
                    End If
                Else
                    de_informa.cn_informa.RollbackTrans
                    Close #1
                    MsgBox "Erro Com os Registros do Arquivo. Processo Cancelado !", vbCritical
                    Frame1.Enabled = True
                    tabTipoFatura.Enabled = True
                    tabTipoFatura.TabEnabled(0) = False
                    tabTipoFatura.TabEnabled(1) = False
                    tabTipoFatura.TabEnabled(2) = True
                    tabTipoFatura.TabEnabled(3) = False
                    txtPFArquivo.SetFocus
                    Exit Sub
                End If
            
             Loop
         
         de_informa.cn_informa.CommitTrans
     
         Close #1
            
        If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
        de_informa.Sel_PreFatura TransFatur(txtPreFilial, txtPreFatura)
        
        gridItens.DataMember = "sel_prefatura"
        gridItens.Refresh
        
        If de_informa.rsSel_PreFaturaTotais.State = 1 Then de_informa.rsSel_PreFaturaTotais.Close
        de_informa.Sel_PreFaturaTotais TransFatur(txtPreFilial, txtPreFatura)
        
        If Not IsNull(de_informa.rsSel_PreFaturaTotais.Fields("fretetot")) Then
            lblTotalFatura.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretetot"), "##,###,##0.00")
            lblTotalFaturaICMS.Caption = Format(de_informa.rsSel_PreFaturaTotais.Fields("fretebrtot"), "##,###,##0.00")
            lblQtdeCtc = Format(de_informa.rsSel_PreFaturaTotais.Fields("qtde"), "#,##0")
        Else
            lblTotalFatura.Caption = "0,00"
            lblTotalFaturaICMS.Caption = "0,00"
            lblQtdeCtc = "0"
        End If
        
        If de_informa.rsSel_PreFaturaTotais.Fields("qtde") > 0 Then
            cmdExcluirCTC.Enabled = True
            cmdExcluirTudo.Enabled = True
        Else
            cmdExcluirCTC.Enabled = False
            cmdExcluirTudo.Enabled = False
        End If
        
        If Abs(xvalor_prefatura - (CDbl(SoNumeros(lblTotalFatura)) / 100)) > 0.1 Then
            MsgBox "ATENÇÃO ! O Arquivo de Pré-Fatura está com Total de R$ " & Trim$(Str(xvalor_prefatura)) & " enquanto os dados processados geraram uma fatura de R$ " & lblTotalFatura & " . Antes de fechar a fatura faça uma conferência dos valores.", vbInformation
        End If
        
        Frame1.Enabled = True
        tabTipoFatura.Enabled = True
        tabTipoFatura.TabEnabled(0) = False
        tabTipoFatura.TabEnabled(1) = False
        tabTipoFatura.TabEnabled(2) = True
        tabTipoFatura.TabEnabled(3) = False
        
        txtPFArquivo.SetFocus
        
    ElseIf optPFProceda.Value = True Then
    
     
    End If
        
    

End Sub

Private Sub cmdSair_Click()
    If Len(Trim$(lblTotalFatura)) > 0 Then
        If CDbl(SoNumeros(lblTotalFatura)) / 100 > 0 Then
            If MsgBox("Você Confirma que Deseja Sair Desta Pré-Fatura ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                MsgBox "OK ! Anote o Número da Pré-Fatura. Você Pode continuá-la no Menu Faturamento/Pré-Fatura Consulta/Altera.", vbInformation
                Unload Me
            End If
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub


Private Sub flexCTCs_Click()

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
                        
    contador = 51
    
    '52
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdAlterarCob.Visible = False
    End If
                    
    '53
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdIncluirCTC.Enabled = False
    End If
    
    '54
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdIncluirCTCsIntervalo.Enabled = False
    End If
    
    '55
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdExcluirCTC.Enabled = False
    End If
                    
    '56
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdExcluirTudo.Enabled = False
    End If
    
    '57
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdFecharFaturaAvulsa.Enabled = False
    End If
    
    '58
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdGerarFat.Enabled = False
    End If
        
    'O 59 esta no load do form mdiFatura
    
    contador = 60
    
    '60
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdAlterarCob.Enabled = False
    End If
    
    '61
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdIncluirCTC.Enabled = False
    End If
    
    '62
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdExcluirCTC.Enabled = False
    End If
    
    '63
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdExcluirTudo.Enabled = False
    End If
                    
    '64
    contador = contador + 1
    If Mid$(xdireitos, contador, 1) = "0" Then
        frmNovaFatura.cmdGerarFat.Enabled = False
    End If
    
    
    '73 ao 76 serão validados no Form frmConsultaFatura
    
    
    mdiFatura.ToolFaturamento.Visible = True
    If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
    gridItens.DataMember = "sel_prefatura"
    gridItens.Refresh
End Sub
Private Sub cmdNovaPreFat_Click()
    If Len(Trim$(txtPreFatura)) = 0 Then
        If MsgBox("Confirma Fazer uma Nova Pré-fatura ?", vbQuestion + vbYesNo, "Nova Pré-Fatura") <> vbYes Then
            Exit Sub
        End If
    End If
    lblCliente.Caption = ""
    txtEndCob.Text = ""
    txtCepCob.Text = ""
    txtCidCob.Text = ""
    txtUFCob.Text = ""
    txtFoneCob.Text = ""
    txtContatoCob.Text = ""
    cmdGerarFat.Caption = "Gerar Fatura ..."
    txtPreFilial.Text = ""
    txtPreFatura.Text = ""
    txtFilialFatura.Text = ""
    txtCnpj.Text = ""
    mskVencimento.Mask = ""
    mskVencimento.Text = ""
    mskVencimento.Mask = "##/##/####"
    txtNumBanco.Text = ""
    lblNomeBanco.Caption = ""
    lblContaBanco.Caption = ""
    txtFilial.Text = ""
    txtCTC.Text = ""
    lblRemetente.Caption = ""
    lblDestinatario.Caption = ""
    lblConsignatario.Caption = ""
    lblEmissao.Caption = ""
    lblModal.Caption = ""
    lblFrete.Caption = ""
    lblCartaCorr.Caption = "OBSCC"
    lblTotalFatura.Caption = ""
    lblTotalFaturaICMS.Caption = "COMICMS"
    lblQtdeCtc.Caption = ""
    txtFilialInt = ""
    txtCtc1Int = ""
    txtCtc2Int = ""
    txtRemetInt = ""
    txtValorFatAvulsa = ""
    fraCliente.Enabled = True
    fraDadosCob.Enabled = True
    If de_informa.rsSel_PreFatura.State = 1 Then de_informa.rsSel_PreFatura.Close
    gridItens.DataMember = "Sel_PreFatura"
    gridItens.Refresh
    tabTipoFatura.Tab = 0
    txtFilialFatura.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFatura.ToolFaturamento.Visible = True
End Sub

Private Sub mskVencimento_GotFocus()
    mskVencimento.SelStart = 0
    mskVencimento.SelLength = 10
End Sub

Private Sub mskVencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(mskVencimento)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskVencimento_LostFocus()
    If mskVencimento.Text <> "__/__/____" Then
        mskVencimento.Text = century(mskVencimento.Text)
        If IsDate(mskVencimento.Text) = False Or Mid(mskVencimento.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskVencimento.SetFocus
            Exit Sub
        End If
        If CDate(mskVencimento.Text) < datahora("data") Then
            MsgBox "ERRO ! Data de Vencimento Menor que Hoje ???", vbCritical, "Erro"
            mskVencimento.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub optCTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then    'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optNFS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then    'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 0
    If txtPreFilial.Enabled = True Then
        If Len(Trim$(txtPreFilial)) > 0 Then
            cmdBuscaPreFat_Click
        Else
            txtPreFilial.SetFocus
        End If
    End If
End Sub

Private Sub txtCepCob_Change()
    If Len(Trim$(txtCepCob)) > 0 Then
        If Not IsNumeric(txtCepCob) Or Mid$(txtCepCob, Len(txtCepCob), 1) = "," Or Mid$(txtCepCob, Len(txtCepCob), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
End Sub

Private Sub txtCepCob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtCepCob)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCidCob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtCidCob)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCidCob_LostFocus()
    txtCidCob = UCase(txtCidCob)
End Sub

Private Sub txtCnpj_Change()
    If Len(Trim$(txtCnpj)) > 0 Then
        If Not IsNumeric(txtCnpj) Or Mid$(txtCnpj, Len(txtCnpj), 1) = "," Or Mid$(txtCnpj, Len(txtCnpj), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
End Sub

Private Sub txtCnpj_GotFocus()
    txtCnpj.SelStart = 0
    txtCnpj.SelLength = 14
End Sub

Private Sub txtCnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtCnpj)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtCnpj_LostFocus()
    If Len(Trim$(txtCnpj)) > 0 Then
    
        If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
        de_informa.Sel_CadCliCGC zeros2(txtCnpj, 14)
        
        If de_informa.rsSel_CadCliCGC.RecordCount < 1 Then
            MsgBox "CNPJ Não Encontrado na Base de Dados !"
            cmdAlterarCob.Enabled = False
            txtCnpj.SetFocus
            Exit Sub
        Else
            lblCliente = de_informa.rsSel_CadCliCGC.Fields("nome")
            cmdAlterarCob.Enabled = True
            If IsNull(de_informa.rsSel_CadCliCGC.Fields("endcob")) Then
                txtEndCob = ""
                txtCepCob = ""
                txtCidCob = ""
                txtUFCob = ""
                txtFoneCob = ""
                txtContatoCob = ""
            Else
                txtEndCob = de_informa.rsSel_CadCliCGC.Fields("endcob")
                txtCepCob = de_informa.rsSel_CadCliCGC.Fields("cepcob")
                txtCidCob = de_informa.rsSel_CadCliCGC.Fields("cidadecob")
                txtUFCob = de_informa.rsSel_CadCliCGC.Fields("ufcob")
                txtFoneCob = de_informa.rsSel_CadCliCGC.Fields("telefonecob")
                txtContatoCob = de_informa.rsSel_CadCliCGC.Fields("contatocob")
            End If
        End If
    Else
        lblCliente = ""
        txtEndCob = ""
        txtCepCob = ""
        txtCidCob = ""
        txtUFCob = ""
        txtFoneCob = ""
        txtContatoCob = ""
    End If
End Sub

Private Sub txtContatoCob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtContatoCob)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtContatoCob_LostFocus()
    txtContatoCob = UCase(txtContatoCob)
End Sub

Private Sub txtCTC_Change()
    If Len(Trim$(txtCTC)) > 0 Then
        If Not IsNumeric(txtCTC) Or Mid$(txtCTC, Len(txtCTC), 1) = "," Or Mid$(txtCTC, Len(txtCTC), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
End Sub

Private Sub txtCTC_GotFocus()
    txtCTC.SelStart = 0
    txtCTC.SelLength = 8
End Sub

Private Sub txtCTC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtCTC)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCTC_LostFocus()
    If Len(Trim$(txtFilial)) > 0 And Len(Trim$(txtCTC)) > 0 Then    'procurar
        If chkLeitor.Value = 1 Then
            cmdBuscaCTC_Click
        End If
    End If
End Sub

Private Sub txtCtc1Int_Change()
    If Len(Trim$(txtCtc1Int)) > 0 Then
        If Not IsNumeric(txtCtc1Int) Or Mid$(txtCtc1Int, Len(txtCtc1Int), 1) = "," Or Mid$(txtCtc1Int, Len(txtCtc1Int), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
End Sub

Private Sub txtCtc1Int_GotFocus()
    txtCtc1Int.SelStart = 0
    txtCtc1Int.SelLength = 8
End Sub

Private Sub txtCtc1Int_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtCtc1Int)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtCtc2Int_Change()
    If Len(Trim$(txtCtc2Int)) > 0 Then
        If Not IsNumeric(txtCtc2Int) Or Mid$(txtCtc2Int, Len(txtCtc2Int), 1) = "," Or Mid$(txtCtc2Int, Len(txtCtc2Int), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
End Sub

Private Sub txtCtc2Int_GotFocus()
    txtCtc2Int.SelStart = 0
    txtCtc2Int.SelLength = 8
End Sub

Private Sub txtCtc2Int_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtCtc2Int)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEndCob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtEndCob)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtEndCob_LostFocus()
    txtEndCob = UCase(txtEndCob)
End Sub

Private Sub txtFilial_Change()
    If Len(Trim$(txtFilial)) > 0 Then
        If Not IsNumeric(txtFilial) Or Mid$(txtFilial, Len(txtFilial), 1) = "," Or Mid$(txtFilial, Len(txtFilial), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
    If Len(Trim$(txtFilial)) = 2 Then
        txtCTC.SetFocus
    End If
End Sub

Private Sub txtFilial_GotFocus()
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub

Private Sub txtFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFilial)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFilial_LostFocus()
    If Len(Trim$(txtFilial)) = 1 Then
        txtFilial = "0" & Trim$(txtFilial)
    End If
End Sub

Private Sub txtFilialFatura_Change()
    If Not IsNumeric(txtFilialFatura) Then
        SendKeys "{BACKSPACE}"
    End If
    If Len(Trim$(txtFilialFatura)) = 2 Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtFilialFatura_GotFocus()
    txtFilialFatura.SelStart = 0
    txtFilialFatura.SelLength = 2
End Sub

Private Sub txtFilialFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFilial)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtFilialFatura_LostFocus()
    If Len(Trim$(txtFilialFatura)) > 0 Then
        txtFilialFatura = zeros(txtFilialFatura.Text, 2)
        If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
        de_informa.Sel_CadFilial Trim$(txtFilialFatura)
        If de_informa.rsSel_CadFilial.RecordCount < 1 Then
            MsgBox "Filial Não Encontrada !", vbInformation
            txtFilialFatura.SetFocus
        End If
    End If
End Sub

Private Sub txtFilialInt_Change()
    If Len(Trim$(txtFilialInt)) > 0 Then
        If Not IsNumeric(txtFilialInt) Or Mid$(txtFilialInt, Len(txtFilialInt), 1) = "," Or Mid$(txtFilialInt, Len(txtFilialInt), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
    If Len(Trim$(txtFilialInt)) = 2 Then
        txtCtc1Int.SetFocus
    End If
End Sub

Private Sub txtFilialInt_GotFocus()
    txtFilialInt.SelStart = 0
    txtFilialInt.SelLength = 2
End Sub

Private Sub txtFilialInt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFilialInt)) > 0 Then    'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFilialInt_LostFocus()
    If Len(Trim$(txtFilialInt)) = 1 Then
        txtFilialInt = "0" & Trim$(txtFilialInt)
    End If
End Sub

Private Sub txtFilialPreFat_Change()
    If Not IsNumeric(txtFilialPreFat) Then
        SendKeys "{BACKSPACE}"
    End If
    If Len(Trim$(txtFilialPreFat)) = 2 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFilialPreFat_GotFocus()
    txtFilialPreFat.SelStart = 0
    txtFilialPreFat.SelLength = 2
End Sub

Private Sub txtFilialPreFat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFilialPreFat)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFilialPreFat_LostFocus()
    If Len(Trim$(txtFilialPreFat)) > 0 Then
        txtFilialPreFat = zeros(txtFilialPreFat.Text, 2)
        If de_informa.rsSel_CadFilial.State = 1 Then de_informa.rsSel_CadFilial.Close
        de_informa.Sel_CadFilial Trim$(txtFilialPreFat)
        If de_informa.rsSel_CadFilial.RecordCount < 1 Then
            MsgBox "Filial Não Encontrada !", vbInformation
            txtFilialPreFat.SetFocus
        End If
    End If
End Sub

Private Sub txtFoneCob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFoneCob)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFoneCob_LostFocus()
    txtFoneCob = UCase(txtFoneCob)
End Sub

Private Sub txtNumBanco_Change()
    If Len(Trim$(txtNumBanco)) > 0 Then
        If Not IsNumeric(txtNumBanco) Or Mid$(txtNumBanco, Len(txtNumBanco), 1) = "," Or Mid$(txtNumBanco, Len(txtNumBanco), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If

End Sub

Private Sub txtNumBanco_GotFocus()
    txtNumBanco.SelStart = 0
    txtNumBanco.SelLength = 4
End Sub

Private Sub txtNumBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtNumBanco)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtNumBanco_LostFocus()
    If Len(Trim$(txtNumBanco)) > 0 Then
        If de_informa.rsSel_BancoCob.State = 1 Then de_informa.rsSel_BancoCob.Close
        de_informa.Sel_BancoCob Int(txtNumBanco), txtFilialFatura
        
        If de_informa.rsSel_BancoCob.RecordCount < 1 Then
            MsgBox "Número de Banco não Encontrado na Base de Dados !"
            lblNomeBanco = ""
            lblcontadoraBanco = ""
            txtNumBanco.SetFocus
            Exit Sub
        Else
            lblNomeBanco = de_informa.rsSel_BancoCob.Fields("nome")
            lblContaBanco = de_informa.rsSel_BancoCob.Fields("conta")
        End If
    End If
End Sub

Private Sub txtPFArquivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtPFArquivo)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPreFatura_Change()
    If Not IsNumeric(txtPreFatura) Then
        SendKeys "{BACKSPACE}"
    End If
End Sub

Private Sub txtPreFatura_GotFocus()
    txtPreFatura.SelStart = 0
    txtPreFatura.SelLength = 2
End Sub

Private Sub txtPreFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtPreFatura)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPreFilial_Change()
    If Not IsNumeric(txtPreFilial) Then
        SendKeys "{BACKSPACE}"
    End If
    If Len(Trim$(txtPreFilial)) = 2 Then
        SendKeys "{TAB}"
    End If
End Sub
Private Sub txtPreFilial_GotFocus()
    txtPreFilial.SelStart = 0
    txtPreFilial.SelLength = 2
End Sub

Private Sub txtPreFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtPreFilial)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRemetInt_Change()
    If Len(Trim$(txtRemetInt)) > 0 Then
        If Not IsNumeric(txtRemetInt) Or Mid$(txtRemetInt, Len(txtRemetInt), 1) = "," Or Mid$(txtRemetInt, Len(txtRemetInt), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
End Sub

Private Sub txtRemetInt_GotFocus()
    txtRemetInt.SelStart = 0
    txtRemetInt.SelLength = 14
End Sub

Private Sub txtRemetInt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtRemetInt)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtRemetInt_LostFocus()
    
    If Len(Trim$(txtRemetInt)) > 0 Then
    
        If de_informa.rsSel_CadCliLike.State = 1 Then de_informa.rsSel_CadCliLike.Close
        de_informa.sel_CadCliLike txtRemetInt & "%"
        
        If de_informa.rsSel_CadCliLike.RecordCount < 1 Then
            MsgBox "CNPJ Não Encontrado na Base de Dados !"
            txtRemetInt.SetFocus
        Else
            lblNomeRemetInt = de_informa.rsSel_CadCliLike.Fields("nome")
        End If
        
    Else
        lblNomeRemetInt = ""
    End If
    
    
End Sub

Private Sub txtUFCob_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtUFCob)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtUFCob_LostFocus()
    txtUFCob = UCase(txtUFCob)
End Sub
Private Sub txtValorFatAvulsa_Change()
    If Not IsNumeric(txtValorFatAvulsa) Then
        SendKeys "{BACKSPACE}"
        Exit Sub
    End If
    Call TextMoneyBox_Change(txtValorFatAvulsa)
    DoEvents
End Sub

Private Sub txtValorFatAvulsa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub



Private Sub backup()
    
    If Len(Trim$(txtCnpj)) < 11 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: CLIENTE !"
        txtCnpj.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mskVencimento) Then
        MsgBox "Data de Vencimento Inválida !"
        mskVencimento.SetFocus
        Exit Sub
    End If

    If Len(Trim$(txtNumBanco)) < 1 Then
        MsgBox "Falta Dados para a Geração de Fatura Avulsa: NÚMERO DO BANCO !"
        txtNumBanco.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtEndCob)) < 5 Or Len(Trim$(txtCidCob)) < 3 Or Len(Trim$(txtCepCob)) < 5 Or Len(Trim$(txtFoneCob)) < 5 Then
        MsgBox "Dados de Endereço de Cobrança Inválidos !"
        cmdAlterarCob.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(txtFilialFatura)) < 2 Then
        MsgBox "Filial da Fatura Inválida !"
        txtFilialFatura.SetFocus
        Exit Sub
    End If
    
    cmdIncluirCTC.Enabled = False
    lblRemetente = ""
    lblDestinatario = ""
    lblConsignatario = ""
    lblEmissao = ""
    lblModal = ""
    lblFrete = ""
    lblCartaCorr = ""
    
    If chkLeitor.Value = 1 Then
        If Len(Trim$(txtCTC)) = 7 Then
            txtCTC = Mid$(txtCTC, 1, 6)
        End If
    End If
    
    If de_informa.rsSel_CTC.State = 1 Then de_informa.rsSel_CTC.Close
    de_informa.Sel_CTC transctc(txtFilial, txtCTC)
    
    If de_informa.rsSel_CTC.RecordCount < 1 Then
        'não encontrou
        MsgBox "Não Encontrado CTC com este Número !", vbCritical, "Não Encontadorrado"
        txtCTC.SetFocus
        Exit Sub
    Else
        lblRemetente = de_informa.rsSel_CTC.Fields("remet_nome")
        lblDestinatario = de_informa.rsSel_CTC.Fields("dest_nome")
        lblConsignatario = de_informa.rsSel_CTC.Fields("respons_nome")
        lblEmissao = de_informa.rsSel_CTC.Fields("data")
        lblModal = de_informa.rsSel_CTC.Fields("modal")
        If de_informa.rsSel_CTC.Fields("tipodoc") = "CTC" Or de_informa.rsSel_CTC.Fields("tipodoc") = "COB" Then
            'encontrou CTC
            If de_informa.rsSel_CTC.Fields("tem_ocorr") = "C" Then
                'ctc cancelado
                MsgBox "Este CTC Encontra-se Cancelado !!", vbCritical, "CTC Cancelado"
                txtCTC.SetFocus
                Exit Sub
            Else
                If Len(Trim$(de_informa.rsSel_CTC.Fields("faturanum"))) > 1 Then
                    'CTC Encontra-se Faturado
                    MsgBox "Este CTC Encontra-se Faturado !! Fatura: " & de_informa.rsSel_CTC.Fields("faturanum"), vbCritical, "CTC Cancelado"
                    txtCTC.SetFocus
                    Exit Sub
                End If
                If de_informa.rsSel_CTC.Fields("respons_cgc") <> zeros2(txtCnpj, 14) Then
                    'O Consignatario não é o Mesmo
                    Do While lblCartaCorr = ""
                        If MsgBox("Atenção! O Consignatário deste CTC é Diferente do Cliente Desta Fatura. Pode Ocorrer de o CTC estar com o Consignatário Errado, Necessitando a Emissão de uma Carta de Correção. Deseja, Mesmo Assim Incluir Este CTC Nesta Fatura ? (você terá que informar o número da Carta de Correção).", vbCritical + vbYesNo, "Consignatário Diferente") = vbNo Then
                            txtCTC.SetFocus
                            Exit Sub
                        Else
                            lblCartaCorr = InputBox("Entre com o Número da Carta de Correção que Retifica o Consignatário deste CTC:", "Número da Carta de Correção")
                            If lblCartaCorr = "" Then
                            Else
                                lblCartaCorr = "Carta de Correção (Consignatário): " & lblCartaCorr
                            End If
                        End If
                    Loop
                End If
                
                If IsNull(de_informa.rsSel_CTC.Fields("subtrib")) Then
                    'faturando CTCs do SITLA
                    If de_informa.rsSel_CTC.Fields("fretetotal") <> de_informa.rsSel_CTC.Fields("fretetotalbruto") Then
                        'checando se é substituicao tributária CTC do SITLA
                        'busca dados no cadastro UF do remetente
                        If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
                        de_informa.Sel_CadCliCGC de_informa.rsSel_CTC.Fields("remet_cgc")
                        xuf_orig = de_informa.rsSel_CadCliCGC.Fields("uf")
                        'busca dados no cadastro do cliente
                        If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
                        de_informa.Sel_CadCliCGC txtCnpj
                        'busca dados do UF Fiscal
                        If de_informa.rsSel_UFFiscal.State = 1 Then de_informa.rsSel_UFFiscal.Close
                        de_informa.Sel_UFFiscal xuf_orig, de_informa.rsSel_CTC.Fields("uf_dest")
                        'verifica inscricao estadual
                        If IsNumeric(SoNumeros(de_informa.rsSel_CadCliCGC.Fields("ie"))) Then
                            If de_informa.rsSel_UFFiscal.Fields("subtribut") = "1" Then
                                'check condicoes da regra da substr.tribut.
                                'dúvida: verificação do UF de coleta/origem ou do remetente ???
                                If ((txtCnpj = de_informa.rsSel_CTC.Fields("dest_cgc") Or txtCnpj = de_informa.rsSel_CTC.Fields("remet_cgc")) Or _
                                    (Mid$(txtCnpj, 1, 8) = "02426290")) And _
                                    CDbl(SoNumeros(de_informa.rsSel_CadCliCGC.Fields("ie"))) > 10000 Then
                                    
                                    lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotal"), "##,###,##0.00")
                                    If de_informa.rsSel_CTC.Fields("fretetotal") < 0.1 Then
                                        MsgBox "Valor de Frete deste CTC Inválido !"
                                        txtCTC.SetFocus
                                        Exit Sub
                                    Else
                                        cmdIncluirCTC.Enabled = True
                                        'cmdIncluirCTC.SetFocus
                                    End If

                                Else
                                    lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotalbruto"), "##,###,##0.00")
                                    If de_informa.rsSel_CTC.Fields("fretetotalbruto") < 0.1 Then
                                        MsgBox "Valor de Frete deste CTC Inválido !"
                                        txtCTC.SetFocus
                                        Exit Sub
                                    Else
                                        cmdIncluirCTC.Enabled = True
                                        'cmdIncluirCTC.SetFocus
                                    End If
                                    
                                End If
                            Else
                                lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotalbruto"), "##,###,##0.00")
                                If de_informa.rsSel_CTC.Fields("fretetotalbruto") < 0.1 Then
                                    MsgBox "Valor de Frete deste CTC Inválido !"
                                    txtCTC.SetFocus
                                    Exit Sub
                                Else
                                    cmdIncluirCTC.Enabled = True
                                    'cmdIncluirCTC.SetFocus
                                End If
                            End If
                        Else
                            lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotalbruto"), "##,###,##0.00")
                            If de_informa.rsSel_CTC.Fields("fretetotalbruto") < 0.1 Then
                                MsgBox "Valor de Frete deste CTC Inválido !"
                                txtCTC.SetFocus
                                Exit Sub
                            Else
                                cmdIncluirCTC.Enabled = True
                                'cmdIncluirCTC.SetFocus
                            End If
                        End If
                    
                    Else
                        lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotal"), "##,###,##0.00")
                        If de_informa.rsSel_CTC.Fields("fretetotal") < 0.1 Then
                            MsgBox "Valor de Frete deste CTC Inválido !"
                            txtCTC.SetFocus
                            Exit Sub
                        Else
                            cmdIncluirCTC.Enabled = True
                            'cmdIncluirCTC.SetFocus
                        End If
                    End If
                Else
                    If de_informa.rsSel_CTC.Fields("subtrib") = "S" Then
                        lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotal"), "##,###,##0.00")
                        If de_informa.rsSel_CTC.Fields("fretetotal") < 0.1 Then
                            MsgBox "Valor de Frete deste CTC Inválido !"
                            txtCTC.SetFocus
                            Exit Sub
                        Else
                            cmdIncluirCTC.Enabled = True
                            'cmdIncluirCTC.SetFocus
                        End If
                    Else
                        lblFrete = Format(de_informa.rsSel_CTC.Fields("fretetotalbruto"), "##,###,##0.00")
                        If de_informa.rsSel_CTC.Fields("fretetotalbruto") < 0.1 Then
                            MsgBox "Valor de Frete deste CTC Inválido !", vbInformation, "Valor de Frete"
                            txtCTC.SetFocus
                            Exit Sub
                        Else
                            cmdIncluirCTC.Enabled = True
                            'cmdIncluirCTC.SetFocus
                        End If
                    End If
                End If
            End If
        Else
            'encontrou documento não fiscal
            MsgBox "Este Documento não é um CTC !! Este Documento não pode ser Faturado: " & de_informa.rsSel_CTC.Fields("tipodoc"), vbCritical, "Documento Não Fiscal"
            txtCTC.SetFocus
            Exit Sub
        End If
    End If



End Sub
