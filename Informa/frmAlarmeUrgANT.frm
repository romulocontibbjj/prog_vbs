VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAlarmeUrg 
   Caption         =   "Alarme Informa - PENDENTES (Sem Posição)"
   ClientHeight    =   7800
   ClientLeft      =   225
   ClientTop       =   795
   ClientWidth     =   11415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11415
   Begin VB.Timer tm_atualiza 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "URGÊNCIAS"
      TabPicture(0)   =   "frmAlarmeUrg.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraUrgencias"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdImprListUrg"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSairUrg"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdImprTelaUrg"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCopiarUrg"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "PRIORIDADES"
      TabPicture(1)   =   "frmAlarmeUrg.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCopiarPri"
      Tab(1).Control(1)=   "cmdImprTelaPri"
      Tab(1).Control(2)=   "fraPrioridades"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(4)=   "cmdSairPri"
      Tab(1).Control(5)=   "cmdImprListPri"
      Tab(1).Control(6)=   "Label5"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "RESUMO"
      TabPicture(2)   =   "frmAlarmeUrg.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdProcessar"
      Tab(2).Control(1)=   "fraResNor"
      Tab(2).Control(2)=   "fraResUrg"
      Tab(2).Control(3)=   "fraResPri"
      Tab(2).Control(4)=   "cmdSairRes"
      Tab(2).Control(5)=   "cmdImprTelaRes"
      Tab(2).Control(6)=   "Label10"
      Tab(2).Control(7)=   "Label7"
      Tab(2).Control(8)=   "Label6"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "GERENCIAL"
      TabPicture(3)   =   "frmAlarmeUrg.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab2"
      Tab(3).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   6975
         Left            =   -74880
         TabIndex        =   74
         Top             =   480
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   12303
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   -2147483630
         TabCaption(0)   =   "Movimento Geral"
         TabPicture(0)   =   "frmAlarmeUrg.frx":0070
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblLegendaAereo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblLegendaTotalMesAnt"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblLegendaRodo"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblLegendaTotal"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lblMensagem"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblTexto"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "flexVarPercent"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "flexAereoGerTot"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "flexGeralGerMesAnt"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "flexGeralGerTotMesAnt"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "flexAereoGer"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "cmdProcessarGerGeral"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "comboMesAnoGeral"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "cmdSairGer"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "cmdDetalhe"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "flexRodoGer"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "flexGeralGer"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "flexRodoGerTot"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "flexGeralGerTot"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "cmdImprGer1"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).ControlCount=   20
         TabCaption(1)   =   "Movimento por Cliente"
         TabPicture(1)   =   "frmAlarmeUrg.frx":008C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdImprGer2"
         Tab(1).Control(1)=   "fraClientes"
         Tab(1).Control(2)=   "fraPerfRodo"
         Tab(1).Control(3)=   "fraPerfAir"
         Tab(1).Control(4)=   "fraPorModal"
         Tab(1).Control(5)=   "cmdSairGerencial"
         Tab(1).Control(6)=   "comboMesAnoCliente"
         Tab(1).Control(7)=   "lblAguarde"
         Tab(1).ControlCount=   8
         Begin VB.CommandButton cmdImprGer2 
            Caption         =   "Impr.Tela"
            Enabled         =   0   'False
            Height          =   375
            Left            =   -66960
            TabIndex        =   110
            Top             =   6360
            Width           =   1215
         End
         Begin VB.CommandButton cmdImprGer1 
            Caption         =   "Impr.Tela"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9000
            TabIndex        =   109
            Top             =   6480
            Width           =   975
         End
         Begin VB.Frame fraClientes 
            Caption         =   "Duplo-Clique Sobre o Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5895
            Left            =   -74880
            TabIndex        =   102
            Top             =   360
            Width           =   4095
            Begin MSDataGridLib.DataGrid gridClientes 
               Bindings        =   "frmAlarmeUrg.frx":00A8
               Height          =   5055
               Left            =   120
               TabIndex        =   103
               Top             =   720
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   8916
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
               DataMember      =   "Sel_ClientesAlarmMov"
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   "cgcbase"
                  Caption         =   "CGC Base"
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
                  DataField       =   "nome"
                  Caption         =   "Cliente"
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
                     ColumnWidth     =   900,284
                  EndProperty
                  BeginProperty Column01 
                     ColumnWidth     =   2399,811
                  EndProperty
               EndProperty
            End
            Begin VB.Label lblCliente 
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
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   360
               Width           =   3855
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGeralGerTot 
            Height          =   375
            Left            =   360
            TabIndex        =   101
            Top             =   6000
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexRodoGerTot 
            Height          =   375
            Left            =   360
            TabIndex        =   100
            Top             =   3960
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGeralGer 
            Height          =   1455
            Left            =   360
            TabIndex        =   94
            Top             =   4440
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   2566
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexRodoGer 
            Height          =   1455
            Left            =   360
            TabIndex        =   93
            Top             =   2400
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   2566
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.CommandButton cmdDetalhe 
            Caption         =   "Comparação Mês Atual/Anterior"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6480
            TabIndex        =   86
            Top             =   6480
            Width           =   2415
         End
         Begin VB.Frame fraPerfRodo 
            Caption         =   "Performance Rodoviário"
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
            Left            =   -67440
            TabIndex        =   81
            Top             =   2040
            Width           =   3255
            Begin VB.CommandButton cmdPerfUFRodo 
               Caption         =   "Performance Rodoviário por UF..."
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   91
               Top             =   1150
               Width           =   3000
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPerfRodo 
               Height          =   855
               Left            =   120
               TabIndex        =   85
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1508
               _Version        =   393216
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPerfUfRodo 
               Height          =   2535
               Left            =   120
               TabIndex        =   88
               Top             =   1560
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4471
               _Version        =   393216
               Rows            =   28
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.Frame fraPerfAir 
            Caption         =   "Performance Aéreo"
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
            Left            =   -70680
            TabIndex        =   82
            Top             =   2040
            Width           =   3255
            Begin VB.CommandButton cmdPerfUFAir 
               Caption         =   "Performance Aéreo por UF..."
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   90
               Top             =   1150
               Width           =   3000
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPerfAir 
               Height          =   855
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1508
               _Version        =   393216
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPerfUfAir 
               Height          =   2535
               Left            =   120
               TabIndex        =   87
               Top             =   1560
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4471
               _Version        =   393216
               Rows            =   28
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.Frame fraPorModal 
            Caption         =   "Por Modal "
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
            Left            =   -70680
            TabIndex        =   80
            Top             =   360
            Width           =   6495
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexPorModal 
               Height          =   1335
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   2355
               _Version        =   393216
               Rows            =   4
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.CommandButton cmdSairGerencial 
            Caption         =   "S A I R"
            Height          =   375
            Left            =   -65520
            TabIndex        =   79
            Top             =   6360
            Width           =   1215
         End
         Begin VB.ComboBox comboMesAnoCliente 
            Height          =   315
            Left            =   -74880
            TabIndex        =   78
            Text            =   "Mes/Ano"
            Top             =   6480
            Width           =   2295
         End
         Begin VB.CommandButton cmdSairGer 
            Caption         =   "S A I R"
            Height          =   375
            Left            =   10080
            TabIndex        =   77
            Top             =   6480
            Width           =   735
         End
         Begin VB.ComboBox comboMesAnoGeral 
            Height          =   315
            ItemData        =   "frmAlarmeUrg.frx":00C1
            Left            =   360
            List            =   "frmAlarmeUrg.frx":00C3
            TabIndex        =   76
            Text            =   "Mes/Ano"
            Top             =   6480
            Width           =   2295
         End
         Begin VB.CommandButton cmdProcessarGerGeral 
            Caption         =   "Processar"
            Height          =   375
            Left            =   5280
            TabIndex        =   75
            Top             =   6480
            Width           =   1095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexAereoGer 
            Height          =   1455
            Left            =   360
            TabIndex        =   92
            Top             =   360
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   2566
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGeralGerTotMesAnt 
            Height          =   375
            Left            =   360
            TabIndex        =   104
            Top             =   1920
            Visible         =   0   'False
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexGeralGerMesAnt 
            Height          =   1455
            Left            =   360
            TabIndex        =   105
            Top             =   360
            Visible         =   0   'False
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   2566
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexAereoGerTot 
            Height          =   375
            Left            =   360
            TabIndex        =   99
            Top             =   1920
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexVarPercent 
            Height          =   375
            Left            =   360
            TabIndex        =   107
            Top             =   4440
            Visible         =   0   'False
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblTexto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   360
            TabIndex        =   108
            Top             =   5040
            Visible         =   0   'False
            Width           =   10455
         End
         Begin VB.Label lblMensagem 
            Alignment       =   2  'Center
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
            Left            =   2760
            TabIndex        =   98
            Top             =   6480
            Width           =   2415
         End
         Begin VB.Label lblLegendaTotal 
            Caption         =   "T O T A L"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   97
            Top             =   4440
            Width           =   255
         End
         Begin VB.Label lblLegendaRodo 
            Caption         =   "R O D O"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   96
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label lblAguarde 
            AutoSize        =   -1  'True
            Caption         =   "Processando. Aguarde ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   -70440
            TabIndex        =   89
            Top             =   6480
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Label lblLegendaTotalMesAnt 
            Caption         =   "T O T A L"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   106
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblLegendaAereo 
            Caption         =   "A É R E O"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   95
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdCopiarPri 
         Caption         =   "Copiar..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70200
         TabIndex        =   72
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCopiarUrg 
         Caption         =   "Copiar..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   71
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar Resumo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68880
         TabIndex        =   70
         Top             =   6960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame fraResNor 
         Caption         =   "Movimento Normal (Tratando Transit-Time)"
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
         Height          =   1935
         Left            =   -74880
         TabIndex        =   59
         Top             =   4920
         Width           =   10935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexNorAnoMes 
            Height          =   1575
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2778
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexNorRegiao 
            Height          =   1575
            Left            =   3360
            TabIndex        =   69
            Top             =   240
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   2778
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraResUrg 
         Caption         =   "Urgências (Não Tratando Transit-Time)"
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
         Height          =   1935
         Left            =   -74880
         TabIndex        =   58
         Top             =   840
         Width           =   10935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexUrgAnoMes 
            Height          =   1575
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2778
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexUrgRegiao 
            Height          =   1575
            Left            =   3360
            TabIndex        =   65
            Top             =   240
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   2778
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraResPri 
         Caption         =   "Prioridades (Tratando Transit-Time)"
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
         Height          =   1935
         Left            =   -74880
         TabIndex        =   57
         Top             =   2880
         Width           =   10935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexPriAnoMes 
            Height          =   1575
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2778
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexPriRegiao 
            Height          =   1575
            Left            =   3360
            TabIndex        =   67
            Top             =   240
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   2778
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.CommandButton cmdSairRes 
         Caption         =   "SAIR"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -65280
         TabIndex        =   56
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprTelaRes 
         Caption         =   "Imprimir Tela..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67080
         TabIndex        =   55
         Top             =   6960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprTelaPri 
         Caption         =   "Imprimir Tela..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68880
         TabIndex        =   54
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprTelaUrg 
         Caption         =   "Imprimir Tela..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   6120
         TabIndex        =   53
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Frame fraPrioridades 
         Caption         =   "CTCs com Prioridade (Tratando Transit-Time)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74880
         TabIndex        =   49
         Top             =   4200
         Width           =   10935
         Begin MSDataGridLib.DataGrid gridPrioridades 
            Bindings        =   "frmAlarmeUrg.frx":00C5
            Height          =   2295
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4048
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
            DataMember      =   "Sel_Urgencias"
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "filialctc"
               Caption         =   "filialctc"
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
               DataField       =   "data"
               Caption         =   "data"
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
               DataField       =   "hora"
               Caption         =   "hora"
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
               DataField       =   "modal"
               Caption         =   "modal"
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
               DataField       =   "prioridade"
               Caption         =   "prioridade"
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
               DataField       =   "remet_nome"
               Caption         =   "remet_nome"
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
               DataField       =   "cidade_orig"
               Caption         =   "cidade_orig"
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
               DataField       =   "dest_nome"
               Caption         =   "dest_nome"
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
               DataField       =   "cidade_dest"
               Caption         =   "cidade_dest"
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
               DataField       =   "uf_dest"
               Caption         =   "uf_dest"
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
               DataField       =   "nfs"
               Caption         =   "nfs"
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
               DataField       =   "obs_emissao"
               Caption         =   "obs_emissao"
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
               DataField       =   "tem_ocorr"
               Caption         =   "tem_ocorr"
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
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1154,835
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   3135,118
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   3254,74
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   2594,835
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   420,095
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   780,095
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   3615
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame11 
            Caption         =   "Observação de Emissão do CTC"
            Height          =   855
            Left            =   120
            TabIndex        =   40
            Top             =   2640
            Width           =   10695
            Begin VB.Label lblObsEmissPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   10455
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Origem"
            Height          =   975
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   5295
            Begin VB.Label lblRemetPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   5055
            End
            Begin VB.Label lblCidadeOrigPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   38
               Top             =   600
               Width           =   3255
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Destino"
            Height          =   975
            Left            =   5520
            TabIndex        =   33
            Top             =   840
            Width           =   5295
            Begin VB.Label lblDestPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   5055
            End
            Begin VB.Label lblCidadeDestPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   35
               Top             =   600
               Width           =   3255
            End
            Begin VB.Label lblUfDestPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3600
               TabIndex        =   34
               Top             =   600
               Width           =   375
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Notas Fiscais"
            Height          =   615
            Left            =   120
            TabIndex        =   31
            Top             =   1920
            Width           =   10695
            Begin VB.Label lblNfsPri 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   10455
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "STATUS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8880
            TabIndex        =   29
            Top             =   120
            Width           =   1935
            Begin VB.Label lblStatusPri 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   945
               TabIndex        =   30
               Top             =   240
               Width           =   90
            End
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   4560
            TabIndex        =   63
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblDataEmiPri 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3000
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblHoraEmiPri 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5040
            TabIndex        =   47
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblFilialCtcPri 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   46
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "CTC: "
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   2280
            TabIndex        =   44
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   6240
            TabIndex        =   43
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblModalPri 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6840
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdSairPri 
         Caption         =   "SAIR"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -65280
         TabIndex        =   25
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprListPri 
         Caption         =   "Imprimir Listagem..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67080
         TabIndex        =   24
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CommandButton cmdSairUrg 
         Caption         =   "SAIR"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9720
         TabIndex        =   23
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdImprListUrg 
         Caption         =   "Imprimir Listagem..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   7920
         TabIndex        =   22
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Frame fraUrgencias 
         Caption         =   "CTCs com Urgência (Não tratando Transit-Time)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   20
         Top             =   4200
         Width           =   10935
         Begin MSDataGridLib.DataGrid gridUrgentes 
            Bindings        =   "frmAlarmeUrg.frx":00DE
            Height          =   2295
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
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
            DataMember      =   "Sel_Urgencias"
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "filialctc"
               Caption         =   "filialctc"
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
               DataField       =   "data"
               Caption         =   "data"
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
               DataField       =   "hora"
               Caption         =   "hora"
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
               DataField       =   "modal"
               Caption         =   "modal"
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
               DataField       =   "prioridade"
               Caption         =   "prioridade"
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
               DataField       =   "remet_nome"
               Caption         =   "remet_nome"
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
               DataField       =   "cidade_orig"
               Caption         =   "cidade_orig"
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
               DataField       =   "dest_nome"
               Caption         =   "dest_nome"
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
               DataField       =   "cidade_dest"
               Caption         =   "cidade_dest"
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
               DataField       =   "uf_dest"
               Caption         =   "uf_dest"
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
               DataField       =   "nfs"
               Caption         =   "nfs"
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
               DataField       =   "obs_emissao"
               Caption         =   "obs_emissao"
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
               DataField       =   "tem_ocorr"
               Caption         =   "tem_ocorr"
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
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               BeginProperty Column00 
                  ColumnWidth     =   1110,047
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1035,213
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1154,835
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1140,095
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   3135,118
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   3254,74
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   2594,835
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   420,095
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   780,095
               EndProperty
            EndProperty
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
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame13 
            Caption         =   "STATUS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8880
            TabIndex        =   26
            Top             =   120
            Width           =   1935
            Begin VB.Label lblStatusUrg 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   945
               TabIndex        =   27
               Top             =   240
               Width           =   90
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Notas Fiscais"
            Height          =   615
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   10695
            Begin VB.Label lblNfsUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   10455
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Destino"
            Height          =   975
            Left            =   5520
            TabIndex        =   7
            Top             =   840
            Width           =   5295
            Begin VB.Label lblUfDestUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   3600
               TabIndex        =   10
               Top             =   600
               Width           =   375
            End
            Begin VB.Label lblCidadeDestUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   9
               Top             =   600
               Width           =   3255
            End
            Begin VB.Label lblDestUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   5055
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Origem"
            Height          =   975
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   5295
            Begin VB.Label lblCidadeOrigUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   6
               Top             =   600
               Width           =   3255
            End
            Begin VB.Label lblRemetUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Width           =   5055
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Observação de Emissão do CTC"
            Height          =   855
            Left            =   120
            TabIndex        =   2
            Top             =   2640
            Width           =   10695
            Begin VB.Label lblObsEmissUrg 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H000000C0&
               Height          =   495
               Left            =   120
               TabIndex        =   3
               Top             =   240
               Width           =   10455
            End
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   4560
            TabIndex        =   62
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lblModalUrg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6840
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   6240
            TabIndex        =   18
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   2280
            TabIndex        =   17
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "CTC: "
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   405
         End
         Begin VB.Label lblFilialctcUrg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblHoraEmiUrg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5040
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblDataEmiUrg 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3000
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Pend. = SEM POSIÇÃO"
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
         Left            =   -73800
         TabIndex        =   73
         Top             =   7080
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Totais por Ano/Mês"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74040
         TabIndex        =   61
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Detalhe Por Regiao de Atendimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -69000
         TabIndex        =   60
         Top             =   480
         Width           =   3060
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "PRIORIDADES PENDENTES"
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
         Height          =   360
         Left            =   -74640
         TabIndex        =   52
         Top             =   6960
         Width           =   4020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "URGÊNCIAS PENDENTES"
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
         Height          =   360
         Left            =   480
         TabIndex        =   51
         Top             =   6960
         Width           =   3750
      End
   End
End
Attribute VB_Name = "frmAlarmeUrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub cmdCopiarPri_Click()
    Dim atcopiar As CClipboard, xlinha As String
    
    xlinha = "CTC: " & gridPrioridades.Columns(0) & "  Data: " & gridPrioridades.Columns(1) & "  Modal: " & gridPrioridades.Columns(3) & "  " & gridPrioridades.Columns(4) & Chr(13) & Chr(10) & _
             "Remetente: " & gridPrioridades.Columns(5) & Chr(13) & Chr(10) & _
             "Cidade - UF: " & Trim$(gridPrioridades.Columns(8)) & " - " & gridPrioridades.Columns(9) & Chr(13) & Chr(10) & _
             "Obs: " & Trim$(gridPrioridades.Columns(11)) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Set atcopiar = New CClipboard
    
    atcopiar.Clear
    atcopiar.SetText xlinha
    
    'MsgBox xlinha

End Sub

Private Sub cmdCopiarUrg_Click()
    Dim atcopiar As CClipboard, xlinha As String
    
    xlinha = "CTC: " & gridUrgentes.Columns(0) & "  Data: " & gridUrgentes.Columns(1) & "  Modal: " & gridUrgentes.Columns(3) & "  " & gridUrgentes.Columns(4) & Chr(13) & Chr(10) & _
             "Remetente: " & gridUrgentes.Columns(5) & Chr(13) & Chr(10) & _
             "Cidade - UF: " & Trim$(gridUrgentes.Columns(8)) & " - " & gridUrgentes.Columns(9) & Chr(13) & Chr(10) & _
             "Obs: " & Trim$(gridUrgentes.Columns(11)) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    
    Set atcopiar = New CClipboard
    
    atcopiar.Clear
    atcopiar.SetText xlinha
    
    'MsgBox xlinha
End Sub

Private Sub cmdDetalhe_Click()
    If cmdDetalhe.Caption = "Comparação Mês Atual/Anterior" Then
        lblLegendaAereo.Visible = False
        lblLegendaRodo.Visible = False
        lblLegendaTotalMesAnt.Visible = True
        
        flexAereoGer.Visible = False
        flexRodoGer.Visible = False
        flexGeralGerMesAnt.Visible = True
        
        flexAereoGerTot.Visible = False
        flexRodoGerTot.Visible = False
        flexGeralGerTotMesAnt.Visible = True
        
        lblLegendaTotal.Top = 2400  'para voltar 4440
        flexGeralGer.Top = 2400
        flexGeralGerTot.Top = 3960  'para voltar 6000
        
        flexVarPercent.Visible = True
        lblTexto.Visible = True
        
        cmdDetalhe.Caption = "Voltar"
        comboMesAnoGeral.Enabled = False
        cmdProcessarGerGeral.Enabled = False
    Else
        lblLegendaAereo.Visible = True
        lblLegendaRodo.Visible = True
        lblLegendaTotalMesAnt.Visible = False
        
        flexAereoGer.Visible = True
        flexRodoGer.Visible = True
        flexGeralGerMesAnt.Visible = False
        
        flexAereoGerTot.Visible = True
        flexRodoGerTot.Visible = True
        flexGeralGerTotMesAnt.Visible = False
        
        lblLegendaTotal.Top = 4440
        flexGeralGer.Top = 4440
        flexGeralGerTot.Top = 6000
        
        flexVarPercent.Visible = False
        lblTexto.Visible = False
        
        cmdDetalhe.Caption = "Comparação Mês Atual/Anterior"
        comboMesAnoGeral.Enabled = True
        cmdProcessarGerGeral.Enabled = True
    End If
    
    DoEvents
    
    
End Sub

Private Sub cmdImprGer1_Click()
    
    Me.MousePointer = 11
    comboMesAnoGeral.Enabled = False
    cmdProcessarGerGeral.Enabled = False
    cmdDetalhe.Enabled = False
    cmdSairGer.Enabled = False
    SSTab1.Enabled = False
    flexAereoGer.Enabled = False
    flexRodoGer.Enabled = False
    flexGeralGer.Enabled = False
    cmdImprGer1.Enabled = False
    DoEvents
    
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    
    MsgBox "Imagem da Tela Enviado para a Impressora !"
    
    Me.MousePointer = 0
    comboMesAnoGeral.Enabled = True
    cmdProcessarGerGeral.Enabled = True
    cmdSairGer.Enabled = True
    SSTab1.Enabled = True
    flexAereoGer.Enabled = True
    flexRodoGer.Enabled = True
    flexGeralGer.Enabled = True
    DoEvents
    lblMensagem.Caption = comboMesAnoGeral.Text
    comboMesAnoGeral.Enabled = True
    cmdProcessarGerGeral.Enabled = True
    cmdImprGer1.Enabled = True
    If comboMesAnoGeral.ListIndex = 0 Then
        cmdDetalhe.Enabled = True
    Else
        cmdDetalhe.Enabled = False
    End If

End Sub

Private Sub cmdImprGer2_Click()
    Dim xenableAir As Boolean, xenableRodo As Boolean
    frmAlarmeUrg.MousePointer = 11
    xenableAir = cmdPerfUFAir.Enabled
    xenableRodo = cmdPerfUFRodo.Enabled
    lblaguarde.Visible = True
    fraClientes.Enabled = False
    fraPorModal.Enabled = False
    fraPerfRodo.Enabled = False
    fraPerfAir.Enabled = False
    cmdPerfUFRodo.Enabled = False
    cmdPerfUFAir.Enabled = False
    cmdSairGerencial.Enabled = False
    comboMesAnoCliente.Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    cmdImprGer2.Enabled = False
    DoEvents
    
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    
    MsgBox "Imagem da Tela Enviado para a Impressora !"
    
    frmAlarmeUrg.MousePointer = 0
    lblaguarde.Visible = False
    fraClientes.Enabled = True
    fraPorModal.Enabled = True
    fraPerfRodo.Enabled = True
    fraPerfAir.Enabled = True
    cmdSairGerencial.Enabled = True
    comboMesAnoCliente.Enabled = True
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    cmdImprGer2.Enabled = True
    cmdPerfUFRodo.Enabled = xenableRodo
    cmdPerfUFAir.Enabled = xenableAir
    DoEvents

End Sub

Private Sub cmdImprListPri_Click()

Dim xColuna As Single, xlinha As Single, xremetente As String, xdestinatario As String, xcidadeuf As String
    
cmdImprTelaUrg.Enabled = False
cmdImprListUrg.Enabled = False
cmdSairUrg.Enabled = False
cmdImprTelaPri.Enabled = False
cmdImprListPri.Enabled = False
cmdSairPri.Enabled = False
    
If Printer.Orientation = vbPRORLandscape Then Printer.Orientation = vbPRORPortrait
    
If de_informa.rsSel_Prioridades.RecordCount > 0 Then
    
    de_informa.rsSel_Prioridades.MoveFirst
    xColuna = 1
    xlinha = 0
    Do Until de_informa.rsSel_Prioridades.EOF
        If xlinha = 0 And xColuna = 1 Then   'identifica inicio da página/cabeçário
            Printer.FontName = "Courier New"
            Printer.Print
            Printer.Print
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontUnderline = True
            Printer.Print Spc(5); "INTEC TRANSPORTES"
            Printer.FontUnderline = False
            Printer.Print
            Printer.Print Spc(5); "RELATÓRIO DE PRIORIDADES PENDENTES (Sem Posição)"
            Printer.Print Spc(5); "USUÁRIO: " & xusuario
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontStrikethru = False
            Printer.Print Spc(9); "Filial-CTC"; Spc(5); "Data"; Spc(7); "Modal"; Spc(5); "Remetente"; Spc(14); "Destinatário"; Spc(11); "Cidade/UF"; Spc(10); "Status"
            Printer.FontSize = 12
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontStrikethru = False
            Printer.FontBold = False
            Printer.FontUnderline = False
        End If
        
        xremetente = Trim$(Mid$(de_informa.rsSel_Prioridades.Fields("remet_nome"), 1, 20))
        If Len(xremetente) < 20 Then
            xremetente = xremetente & String(20 - Len(xremetente), " ")
        End If
        xdestinatario = Trim$(Mid$(de_informa.rsSel_Prioridades.Fields("dest_nome"), 1, 20))
        If Len(xdestinatario) < 20 Then
            xdestinatario = xdestinatario & String(20 - Len(xdestinatario), " ")
        End If
        xcidadeuf = Trim$(Mid$(de_informa.rsSel_Prioridades.Fields("Cidade_dest"), 1, 15)) & "-" & de_informa.rsSel_Prioridades.Fields("uf_dest")
        If Len(xcidadeuf) < 18 Then
            xcidadeuf = xcidadeuf & String(18 - Len(xcidadeuf), " ")
        End If

        Printer.Print Spc(9); de_informa.rsSel_Prioridades.Fields("filialctc"); Spc(2); _
                                zeros(Day(de_informa.rsSel_Prioridades.Fields("data")), 2) & "/" & _
                                zeros(Month(de_informa.rsSel_Prioridades.Fields("data")), 2) & "/" & _
                                zeros(Year(de_informa.rsSel_Prioridades.Fields("data")), 4); Spc(2); _
                                Trim$(de_informa.rsSel_Prioridades.Fields("modal")); Spc(2 + (10 - Len(Trim$(de_informa.rsSel_Prioridades.Fields("modal"))))); _
                                xremetente; Spc(3); xdestinatario; Spc(3); xcidadeuf; Spc(3); _
                                de_informa.rsSel_Prioridades.Fields("tem_ocorr")
        xlinha = xlinha + 1
        If xlinha = 70 Then
            xlinha = 0
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontBold = False
            Printer.FontStrikethru = False
            Printer.Print
            Printer.NewPage
        End If
        de_informa.rsSel_Prioridades.MoveNext
    Loop
    
    de_informa.rsSel_Prioridades.MoveFirst
            
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontStrikethru = True
    Printer.Print Spc(5); String(132, " ")
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.FontStrikethru = False
    Printer.Print
    Printer.NewPage
    Printer.EndDoc   'finaliza spool da impressão
    DoEvents
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "IMPRESSÃO", xusuario, "RELATÓRIO DE PRIORIDADES"
        
    MsgBox "RELATÓRIO ENVIADO PARA IMPRESSÃO !"
    
Else
    MsgBox "Não Há Dados a Serem Impressos !"
End If

cmdImprTelaUrg.Enabled = True
cmdImprListUrg.Enabled = True
cmdSairUrg.Enabled = True
cmdImprTelaPri.Enabled = True
cmdImprListPri.Enabled = True
cmdSairPri.Enabled = True

End Sub

Private Sub cmdImprListUrg_Click()

Dim xColuna As Single, xlinha As Single, xremetente As String, xdestinatario As String, xcidadeuf As String

cmdImprTelaUrg.Enabled = False
cmdImprListUrg.Enabled = False
cmdSairUrg.Enabled = False
cmdImprTelaPri.Enabled = False
cmdImprListPri.Enabled = False
cmdSairPri.Enabled = False
    
If Printer.Orientation = vbPRORLandscape Then Printer.Orientation = vbPRORPortrait

If de_informa.rsSel_Urgencias.RecordCount > 0 Then
    de_informa.rsSel_Urgencias.MoveFirst
    xColuna = 1
    xlinha = 0
    Do Until de_informa.rsSel_Urgencias.EOF
        If xlinha = 0 And xColuna = 1 Then   'identifica inicio da página/cabeçário
            Printer.FontName = "Courier New"
            Printer.Print
            Printer.Print
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontUnderline = True
            Printer.Print Spc(5); "INTEC TRANSPORTES"
            Printer.FontUnderline = False
            Printer.Print
            Printer.Print Spc(5); "RELATÓRIO DE URGÊNCIAS PENDENTES (Sem Posição)"
            Printer.Print Spc(5); "USUÁRIO: " & xusuario
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontStrikethru = False
            Printer.Print Spc(9); "Filial-CTC"; Spc(5); "Data"; Spc(7); "Modal"; Spc(5); "Remetente"; Spc(14); "Destinatário"; Spc(11); "Cidade/UF"; Spc(10); "Status"
            Printer.FontSize = 12
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontStrikethru = False
            Printer.FontBold = False
            Printer.FontUnderline = False
        End If
        
        xremetente = Trim$(Mid$(de_informa.rsSel_Urgencias.Fields("remet_nome"), 1, 20))
        If Len(xremetente) < 20 Then
            xremetente = xremetente & String(20 - Len(xremetente), " ")
        End If
        xdestinatario = Trim$(Mid$(de_informa.rsSel_Urgencias.Fields("dest_nome"), 1, 20))
        If Len(xdestinatario) < 20 Then
            xdestinatario = xdestinatario & String(20 - Len(xdestinatario), " ")
        End If
        xcidadeuf = Trim$(Mid$(de_informa.rsSel_Urgencias.Fields("Cidade_dest"), 1, 15)) & "-" & de_informa.rsSel_Urgencias.Fields("uf_dest")
        If Len(xcidadeuf) < 18 Then
            xcidadeuf = xcidadeuf & String(18 - Len(xcidadeuf), " ")
        End If

        Printer.Print Spc(9); de_informa.rsSel_Urgencias.Fields("filialctc"); Spc(2); _
                                zeros(Day(de_informa.rsSel_Urgencias.Fields("data")), 2) & "/" & _
                                zeros(Month(de_informa.rsSel_Urgencias.Fields("data")), 2) & "/" & _
                                zeros(Year(de_informa.rsSel_Urgencias.Fields("data")), 4); Spc(2); _
                                Trim$(de_informa.rsSel_Urgencias.Fields("modal")); Spc(2 + (10 - Len(Trim$(de_informa.rsSel_Urgencias.Fields("modal"))))); _
                                xremetente; Spc(3); xdestinatario; Spc(3); xcidadeuf; Spc(3); _
                                de_informa.rsSel_Urgencias.Fields("tem_ocorr")
        xlinha = xlinha + 1
        If xlinha = 70 Then
            xlinha = 0
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 8
            Printer.FontBold = False
            Printer.FontStrikethru = False
            Printer.Print
            Printer.NewPage
        End If
        de_informa.rsSel_Urgencias.MoveNext
    Loop
    
    de_informa.rsSel_Urgencias.MoveFirst
            
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontStrikethru = True
    Printer.Print Spc(5); String(132, " ")
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.FontStrikethru = False
    Printer.Print
    Printer.NewPage
    Printer.EndDoc   'finaliza spool da impressão
    DoEvents
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "IMPRESSÃO", xusuario, "RELATÓRIO DE URGÊNCIA"
        
    MsgBox "RELATÓRIO ENVIADO PARA IMPRESSÃO !"
Else
    MsgBox "Não Há Dados a Serem Impressos !"
End If
    
cmdImprTelaUrg.Enabled = True
cmdImprListUrg.Enabled = True
cmdSairUrg.Enabled = True
cmdImprTelaPri.Enabled = True
cmdImprListPri.Enabled = True
cmdSairPri.Enabled = True

End Sub

Private Sub cmdImprTelaPri_Click()
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    MsgBox "Imagem da Tela Enviado para a Impressora !"
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
End Sub

Private Sub cmdImprTelaRes_Click()
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    MsgBox "Imagem da Tela Enviado para a Impressora !"
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True

End Sub

Private Sub cmdImprTelaUrg_Click()
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
    MsgBox "Imagem da Tela Enviado para a Impressora !"
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
    
End Sub
Private Sub cmdPerfUFAir_Click()
    
    frmAlarmeUrg.MousePointer = 11
    lblaguarde.Visible = True
    fraClientes.Enabled = False
    fraPorModal.Enabled = False
    fraPerfRodo.Enabled = False
    cmdPerfUFRodo.Enabled = False
    cmdPerfUFAir.Enabled = False
    cmdSairGerencial.Enabled = False
    comboMesAnoCliente.Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    cmdImprGer2.Enabled = False
    DoEvents
    
    For xcont = 1 To 27
        FlexPerfUfAir.TextMatrix(xcont, 1) = ""
        FlexPerfUfAir.TextMatrix(xcont, 2) = ""
        FlexPerfUfAir.TextMatrix(xcont, 3) = ""
        FlexPerfUfAir.TextMatrix(xcont, 4) = ""
    Next

    FlexPerfUfAir.SetFocus
    SendKeys "^{HOME}"
    
    'AC
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "AC"
                                
    FlexPerfUfAir.TextMatrix(1, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(1, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(1, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(1, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'AL
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "AL"
                                
    FlexPerfUfAir.TextMatrix(2, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(2, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(2, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(2, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'AM
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "AM"
                                
    FlexPerfUfAir.TextMatrix(3, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(3, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(3, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(3, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'AP
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "AP"
                                
    FlexPerfUfAir.TextMatrix(4, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(4, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(4, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(4, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'BA
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "BA"
                                
    FlexPerfUfAir.TextMatrix(5, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(5, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(5, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(5, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'CE
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "CE"
                                
    FlexPerfUfAir.TextMatrix(6, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(6, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(6, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(6, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'DF
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "DF"
                                
    FlexPerfUfAir.TextMatrix(7, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(7, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(7, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(7, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'ES
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "ES"
                                
    FlexPerfUfAir.TextMatrix(8, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(8, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(8, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(8, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'GO
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "GO"
                                
    FlexPerfUfAir.TextMatrix(9, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(9, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(9, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(9, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'MA
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "MA"
                                
    FlexPerfUfAir.TextMatrix(10, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(10, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(10, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(10, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'MG
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "MG"
                                
    FlexPerfUfAir.TextMatrix(11, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(11, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(11, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(11, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'MS
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "MS"
                                
    FlexPerfUfAir.TextMatrix(12, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(12, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(12, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(12, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'MT
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "MT"
                                
    FlexPerfUfAir.TextMatrix(13, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(13, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(13, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(13, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PA
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "PA"
                                
    FlexPerfUfAir.TextMatrix(14, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(14, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(14, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(14, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PB
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "PB"
                                
    FlexPerfUfAir.TextMatrix(15, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(15, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(15, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(15, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PE
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "PE"
                                
    FlexPerfUfAir.TextMatrix(16, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(16, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(16, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(16, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PI
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "PI"
                                
    FlexPerfUfAir.TextMatrix(17, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(17, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(17, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(17, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PR
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "PR"
                                
    FlexPerfUfAir.TextMatrix(18, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(18, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(18, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(18, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RJ
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "RJ"
                                
    FlexPerfUfAir.TextMatrix(19, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(19, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(19, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(19, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RN
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "RN"
                                
    FlexPerfUfAir.TextMatrix(20, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(20, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(20, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(20, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RO
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "RO"
                                
    FlexPerfUfAir.TextMatrix(21, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(21, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(21, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(21, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RR
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "RR"
                                
    FlexPerfUfAir.TextMatrix(22, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(22, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(22, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(22, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RS
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "RS"
                                
    FlexPerfUfAir.TextMatrix(23, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(23, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(23, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(23, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'SC
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "SC"
                                
    FlexPerfUfAir.TextMatrix(24, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(24, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(24, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(24, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'SE
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "SE"
                                
    FlexPerfUfAir.TextMatrix(25, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(25, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(25, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(25, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'SP
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "SP"
                                
    FlexPerfUfAir.TextMatrix(26, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(26, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(26, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(26, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'TO
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "TO"
                                
    FlexPerfUfAir.TextMatrix(27, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfAir.TextMatrix(27, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfAir.TextMatrix(27, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfAir.TextMatrix(27, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "^{HOME}"
    DoEvents
    
    frmAlarmeUrg.MousePointer = 0
    lblaguarde.Visible = False
    fraClientes.Enabled = True
    fraPorModal.Enabled = True
    fraPerfRodo.Enabled = True
    cmdPerfUFRodo.Enabled = True
    cmdPerfUFAir.Enabled = True
    cmdSairGerencial.Enabled = True
    comboMesAnoCliente.Enabled = True
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    cmdImprGer2.Enabled = True
    DoEvents
    
End Sub

Private Sub cmdPerfUFRodo_Click()

    frmAlarmeUrg.MousePointer = 11
    lblaguarde.Visible = True
    fraClientes.Enabled = False
    fraPorModal.Enabled = False
    fraPerfAir.Enabled = False
    cmdPerfUFAir.Enabled = False
    cmdPerfUFRodo.Enabled = False
    cmdSairGerencial.Enabled = False
    comboMesAnoCliente.Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    cmdImprGer2.Enabled = False
    DoEvents
    
    For xcont = 1 To 27
        FlexPerfUfRodo.TextMatrix(xcont, 1) = ""
        FlexPerfUfRodo.TextMatrix(xcont, 2) = ""
        FlexPerfUfRodo.TextMatrix(xcont, 3) = ""
        FlexPerfUfRodo.TextMatrix(xcont, 4) = ""
    Next
    
    FlexPerfUfRodo.SetFocus
    SendKeys "^{HOME}"
    
    'AC
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "AC"
                                
    FlexPerfUfRodo.TextMatrix(1, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(1, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(1, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(1, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    
    'AL
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "AL"
                                
    FlexPerfUfRodo.TextMatrix(2, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(2, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(2, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(2, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'AM
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "AM"
                                
    FlexPerfUfRodo.TextMatrix(3, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(3, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(3, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(3, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'AP
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "AP"
                                
    FlexPerfUfRodo.TextMatrix(4, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(4, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(4, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(4, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'BA
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "BA"
                                
    FlexPerfUfRodo.TextMatrix(5, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(5, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(5, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(5, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'CE
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "CE"
                                
    FlexPerfUfRodo.TextMatrix(6, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(6, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(6, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(6, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'DF
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "DF"
                                
    FlexPerfUfRodo.TextMatrix(7, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(7, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(7, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(7, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'ES
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "ES"
                                
    FlexPerfUfRodo.TextMatrix(8, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(8, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(8, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(8, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'GO
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "GO"
                                
    FlexPerfUfRodo.TextMatrix(9, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(9, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(9, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(9, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'MA
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "MA"
                                
    FlexPerfUfRodo.TextMatrix(10, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(10, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(10, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(10, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'MG
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "MG"
                                
    FlexPerfUfRodo.TextMatrix(11, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(11, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(11, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(11, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'MS
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "MS"
                                
    FlexPerfUfRodo.TextMatrix(12, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(12, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(12, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(12, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'MT
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "MT"
                                
    FlexPerfUfRodo.TextMatrix(13, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(13, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(13, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(13, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PA
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "PA"
                                
    FlexPerfUfRodo.TextMatrix(14, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(14, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(14, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(14, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PB
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "PB"
                                
    FlexPerfUfRodo.TextMatrix(15, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(15, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(15, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(15, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PE
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "PE"
                                
    FlexPerfUfRodo.TextMatrix(16, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(16, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(16, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(16, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PI
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "PI"
                                
    FlexPerfUfRodo.TextMatrix(17, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(17, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(17, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(17, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'PR
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "PR"
                                
    FlexPerfUfRodo.TextMatrix(18, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(18, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(18, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(18, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RJ
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "RJ"
                                
    FlexPerfUfRodo.TextMatrix(19, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(19, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(19, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(19, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RN
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "RN"
                                
    FlexPerfUfRodo.TextMatrix(20, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(20, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(20, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(20, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RO
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "RO"
                                
    FlexPerfUfRodo.TextMatrix(21, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(21, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(21, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(21, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RR
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "RR"
                                
    FlexPerfUfRodo.TextMatrix(22, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(22, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(22, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(22, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'RS
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "RS"
                                
    FlexPerfUfRodo.TextMatrix(23, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(23, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(23, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(23, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'SC
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "SC"
                                
    FlexPerfUfRodo.TextMatrix(24, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(24, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(24, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(24, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'SE
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "SE"
                                
    FlexPerfUfRodo.TextMatrix(25, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(25, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(25, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(25, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'SP
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "SP"
                                
    FlexPerfUfRodo.TextMatrix(26, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(26, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(26, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(26, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "{DOWN}"
    DoEvents
    
    'TO
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "TO"
                                
    FlexPerfUfRodo.TextMatrix(27, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfUfRodo.TextMatrix(27, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfUfRodo.TextMatrix(27, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfUfRodo.TextMatrix(27, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    SendKeys "^{HOME}"
    DoEvents

    frmAlarmeUrg.MousePointer = 0
    lblaguarde.Visible = False
    fraClientes.Enabled = True
    fraPorModal.Enabled = True
    fraPerfAir.Enabled = True
    cmdPerfUFAir.Enabled = True
    cmdPerfUFRodo.Enabled = True
    cmdSairGerencial.Enabled = True
    comboMesAnoCliente.Enabled = True
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    cmdImprGer2.Enabled = True
    DoEvents

End Sub

Private Sub CmdProcessar_Click()
Dim xLin As Integer
    
    'URGÊNCIA ANO/MES
    
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    CmdProcessar.Enabled = False
    fraResUrg.Enabled = False
    fraResPri.Enabled = False
    fraResNor.Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    
    Me.MousePointer = 11
    
    'Configura as Flex
    
    flexUrgAnoMes.Clear
    flexPriAnoMes.Clear
    flexNorAnoMes.Clear
    flexUrgRegiao.Clear
    flexPriRegiao.Clear
    flexNorRegiao.Clear
    
    'Urgencias - Ano/mes
    
    flexUrgAnoMes.Cols = 7
    flexUrgAnoMes.ColWidth(0) = 200
    flexUrgAnoMes.ColWidth(1) = 830
    flexUrgAnoMes.ColWidth(2) = 670
    flexUrgAnoMes.ColWidth(3) = 550
    flexUrgAnoMes.ColWidth(4) = 580
    flexUrgAnoMes.ColWidth(5) = 1
    flexUrgAnoMes.ColWidth(6) = 1
    
    flexUrgAnoMes.TextMatrix(0, 1) = "Mês/Ano"
    flexUrgAnoMes.TextMatrix(0, 2) = " CTCs"
    flexUrgAnoMes.TextMatrix(0, 3) = "Pend."
    flexUrgAnoMes.TextMatrix(0, 4) = "  %"
    
    'Prioridades - Ano/mes
    
    flexPriAnoMes.Cols = 7
    flexPriAnoMes.ColWidth(0) = 200
    flexPriAnoMes.ColWidth(1) = 830
    flexPriAnoMes.ColWidth(2) = 670
    flexPriAnoMes.ColWidth(3) = 550
    flexPriAnoMes.ColWidth(4) = 580
    flexPriAnoMes.ColWidth(5) = 1
    flexPriAnoMes.ColWidth(6) = 1
    
    
    flexPriAnoMes.TextMatrix(0, 1) = "Mês/Ano"
    flexPriAnoMes.TextMatrix(0, 2) = " CTCs"
    flexPriAnoMes.TextMatrix(0, 3) = "Pend."
    flexPriAnoMes.TextMatrix(0, 4) = "  %"
    
    'Movto Normal - Ano/mes
    
    flexNorAnoMes.Cols = 7
    flexNorAnoMes.ColWidth(0) = 200
    flexNorAnoMes.ColWidth(1) = 830
    flexNorAnoMes.ColWidth(2) = 670
    flexNorAnoMes.ColWidth(3) = 550
    flexNorAnoMes.ColWidth(4) = 580
    flexNorAnoMes.ColWidth(5) = 1
    flexNorAnoMes.ColWidth(6) = 1
    
    flexNorAnoMes.TextMatrix(0, 1) = "Mês/Ano"
    flexNorAnoMes.TextMatrix(0, 2) = " CTCs"
    flexNorAnoMes.TextMatrix(0, 3) = "Pend."
    flexNorAnoMes.TextMatrix(0, 4) = "  %"
    
    'Urgências - Por Região
    
    flexUrgRegiao.Cols = 7
    flexUrgRegiao.ColWidth(0) = 200
    flexUrgRegiao.ColWidth(1) = 400
    flexUrgRegiao.ColWidth(2) = 1000
    flexUrgRegiao.ColWidth(3) = 3740
    flexUrgRegiao.ColWidth(4) = 670
    flexUrgRegiao.ColWidth(5) = 550
    flexUrgRegiao.ColWidth(6) = 580
    
    flexUrgRegiao.TextMatrix(0, 1) = "Cód."
    flexUrgRegiao.TextMatrix(0, 2) = "  Respons."
    flexUrgRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexUrgRegiao.TextMatrix(0, 4) = " CTCs"
    flexUrgRegiao.TextMatrix(0, 5) = "Pend."
    flexUrgRegiao.TextMatrix(0, 6) = "  %"
    
    'Prioridades - Por Região
    
    flexPriRegiao.Cols = 7
    flexPriRegiao.ColWidth(0) = 200
    flexPriRegiao.ColWidth(1) = 400
    flexPriRegiao.ColWidth(2) = 1000
    flexPriRegiao.ColWidth(3) = 3740
    flexPriRegiao.ColWidth(4) = 670
    flexPriRegiao.ColWidth(5) = 550
    flexPriRegiao.ColWidth(6) = 580
    
    flexPriRegiao.TextMatrix(0, 1) = "Cód."
    flexPriRegiao.TextMatrix(0, 2) = "  Respons."
    flexPriRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexPriRegiao.TextMatrix(0, 4) = " CTCs"
    flexPriRegiao.TextMatrix(0, 5) = "Pend."
    flexPriRegiao.TextMatrix(0, 6) = "  %"
    
    'Movto Normal - Por Região
    
    flexNorRegiao.Cols = 7
    flexNorRegiao.ColWidth(0) = 200
    flexNorRegiao.ColWidth(1) = 400
    flexNorRegiao.ColWidth(2) = 1000
    flexNorRegiao.ColWidth(3) = 3740
    flexNorRegiao.ColWidth(4) = 670
    flexNorRegiao.ColWidth(5) = 550
    flexNorRegiao.ColWidth(6) = 580
    
    flexNorRegiao.TextMatrix(0, 1) = "Cód."
    flexNorRegiao.TextMatrix(0, 2) = "  Respons."
    flexNorRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexNorRegiao.TextMatrix(0, 4) = " CTCs"
    flexNorRegiao.TextMatrix(0, 5) = "Pend."
    flexNorRegiao.TextMatrix(0, 6) = "  %"
    
    DoEvents
    
    If de_informa.rsSel_AlarmResAnoMes1.State = 1 Then de_informa.rsSel_AlarmResAnoMes1.Close
    de_informa.Sel_AlarmResAnoMes1 "URGÊNCIA"
    
    If de_informa.rsSel_AlarmResAnoMes1.RecordCount > 0 Then
        flexUrgAnoMes.Rows = de_informa.rsSel_AlarmResAnoMes1.RecordCount + 1
    End If
    
    xLin = 0
    Do Until de_informa.rsSel_AlarmResAnoMes1.EOF
        xLin = xLin + 1
        'busca as pendencias deste ano/mes
        If de_informa.rsSel_AlarmResPend1.State = 1 Then de_informa.rsSel_AlarmResPend1.Close
        de_informa.Sel_AlarmResPend1 "URGÊNCIA", de_informa.rsSel_AlarmResAnoMes1.Fields("ano"), _
                                     de_informa.rsSel_AlarmResAnoMes1.Fields("mes"), CDate(datahora("data")) + 100
        flexUrgAnoMes.TextMatrix(xLin, 1) = MesAno(de_informa.rsSel_AlarmResAnoMes1.Fields("mes"), de_informa.rsSel_AlarmResAnoMes1.Fields("ano"))
        flexUrgAnoMes.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmResAnoMes1.Fields("qtd")
        flexUrgAnoMes.TextMatrix(xLin, 3) = de_informa.rsSel_AlarmResPend1.Fields("qtd")
        flexUrgAnoMes.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmResPend1.Fields("qtd") / de_informa.rsSel_AlarmResAnoMes1.Fields("qtd"), "##0.0%")
        flexUrgAnoMes.TextMatrix(xLin, 5) = de_informa.rsSel_AlarmResAnoMes1.Fields("mes")
        flexUrgAnoMes.TextMatrix(xLin, 6) = de_informa.rsSel_AlarmResAnoMes1.Fields("ano")
        de_informa.rsSel_AlarmResAnoMes1.MoveNext
        DoEvents
    Loop
    
    'PRIORIDADE ANO/MES
    
    If de_informa.rsSel_AlarmResAnoMes1.State = 1 Then de_informa.rsSel_AlarmResAnoMes1.Close
    de_informa.Sel_AlarmResAnoMes1 "PRIORIDADE"
    
    If de_informa.rsSel_AlarmResAnoMes1.RecordCount > 0 Then
        flexPriAnoMes.Rows = de_informa.rsSel_AlarmResAnoMes1.RecordCount + 1
    End If
    
    xLin = 0
    Do Until de_informa.rsSel_AlarmResAnoMes1.EOF
        xLin = xLin + 1
        'busca as pendencias deste ano/mes
        If de_informa.rsSel_AlarmResPend1.State = 1 Then de_informa.rsSel_AlarmResPend1.Close
        de_informa.Sel_AlarmResPend1 "PRIORIDADE", de_informa.rsSel_AlarmResAnoMes1.Fields("ano"), _
                                     de_informa.rsSel_AlarmResAnoMes1.Fields("mes"), datahora("data")
        flexPriAnoMes.TextMatrix(xLin, 1) = MesAno(de_informa.rsSel_AlarmResAnoMes1.Fields("mes"), de_informa.rsSel_AlarmResAnoMes1.Fields("ano"))
        flexPriAnoMes.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmResAnoMes1.Fields("qtd")
        flexPriAnoMes.TextMatrix(xLin, 3) = de_informa.rsSel_AlarmResPend1.Fields("qtd")
        flexPriAnoMes.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmResPend1.Fields("qtd") / de_informa.rsSel_AlarmResAnoMes1.Fields("qtd"), "##0.0%")
        flexPriAnoMes.TextMatrix(xLin, 5) = de_informa.rsSel_AlarmResAnoMes1.Fields("mes")
        flexPriAnoMes.TextMatrix(xLin, 6) = de_informa.rsSel_AlarmResAnoMes1.Fields("ano")
        de_informa.rsSel_AlarmResAnoMes1.MoveNext
        DoEvents
    Loop
    
    'MOVIMENTO NORMAL ANO/MES
    
    If de_informa.rsSel_AlarmResAnoMes1.State = 1 Then de_informa.rsSel_AlarmResAnoMes1.Close
    de_informa.Sel_AlarmResAnoMes1 "NORMAL"
    
    If de_informa.rsSel_AlarmResAnoMes1.RecordCount > 0 Then
        flexNorAnoMes.Rows = de_informa.rsSel_AlarmResAnoMes1.RecordCount + 1
    End If
    
    xLin = 0
    Do Until de_informa.rsSel_AlarmResAnoMes1.EOF
        xLin = xLin + 1
        'busca as pendencias deste ano/mes
        If de_informa.rsSel_AlarmResPend1.State = 1 Then de_informa.rsSel_AlarmResPend1.Close
        de_informa.Sel_AlarmResPend1 "NORMAL", de_informa.rsSel_AlarmResAnoMes1.Fields("ano"), _
                                     de_informa.rsSel_AlarmResAnoMes1.Fields("mes"), datahora("data")
        flexNorAnoMes.TextMatrix(xLin, 1) = MesAno(de_informa.rsSel_AlarmResAnoMes1.Fields("mes"), de_informa.rsSel_AlarmResAnoMes1.Fields("ano"))
        flexNorAnoMes.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmResAnoMes1.Fields("qtd")
        flexNorAnoMes.TextMatrix(xLin, 3) = de_informa.rsSel_AlarmResPend1.Fields("qtd")
        flexNorAnoMes.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmResPend1.Fields("qtd") / de_informa.rsSel_AlarmResAnoMes1.Fields("qtd"), "##0.0%")
        flexNorAnoMes.TextMatrix(xLin, 5) = de_informa.rsSel_AlarmResAnoMes1.Fields("mes")
        flexNorAnoMes.TextMatrix(xLin, 6) = de_informa.rsSel_AlarmResAnoMes1.Fields("ano")
        de_informa.rsSel_AlarmResAnoMes1.MoveNext
        DoEvents
    Loop
    
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
    CmdProcessar.Enabled = True
    fraResUrg.Enabled = True
    fraResPri.Enabled = True
    fraResNor.Enabled = True
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    
    If Mid$(xdireitos, 30, 1) = "0" Then
        SSTab1.TabEnabled(3) = False
    Else
        SSTab1.TabEnabled(3) = True
    End If
    
    Me.MousePointer = 0

End Sub
Private Sub cmdProcessarGerGeral_Click()
    Dim xTotValmerc As Currency, xTotFrete As Currency, xTotPeso As Currency
    Dim xTotVol As Long, xTotCtc As Long, xTotNf As Long, xdataper1 As Date, xdataper2 As Date
    Dim xTotValMercMesAnt As Currency, xTotFreteMesAnt As Currency, xTotPesoMesAnt As Currency
    Dim xTotVolMesAnt As Long, xTotCtcMesAnt As Long, xTotNfMesAnt As Long
    
    Me.MousePointer = 11
    comboMesAnoGeral.Enabled = False
    cmdProcessarGerGeral.Enabled = False
    cmdDetalhe.Enabled = False
    cmdSairGer.Enabled = False
    SSTab1.Enabled = False
    flexAereoGer.Enabled = False
    flexRodoGer.Enabled = False
    flexGeralGer.Enabled = False
    cmdImprGer1.Enabled = False
    
    flexAereoGer.Rows = 1
    flexAereoGer.Rows = 2
    flexAereoGer.FixedRows = 1
    flexRodoGer.Rows = 1
    flexRodoGer.Rows = 2
    flexRodoGer.FixedRows = 1
    flexGeralGer.Rows = 1
    flexGeralGer.Rows = 2
    flexGeralGer.FixedRows = 1
    flexGeralGerMesAnt.Rows = 1
    flexGeralGerMesAnt.Rows = 2
    flexGeralGerMesAnt.FixedRows = 1
    flexAereoGerTot.Rows = 0
    flexAereoGerTot.Rows = 1
    flexRodoGerTot.Rows = 0
    flexRodoGerTot.Rows = 1
    flexGeralGerTot.Rows = 0
    flexGeralGerTot.Rows = 1
    flexGeralGerTotMesAnt.Rows = 0
    flexGeralGerTotMesAnt.Rows = 1
    
    lblMensagem.Caption = "Processando. Aguarde..."
    
    DoEvents
    
'MONTANDO O AÉREO
    
    xdataper1 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 1, 4) & "/" & _
                Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 5, 2) & "/" & "01")

    If comboMesAnoGeral.ListIndex = 0 Then  'se for o mês atual, pega somente até a data de hoje, pelo dia
        xdataper2 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 1, 4) & "/" & _
                    Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 5, 2) & "/" & Day(datahora("DATA")))
    Else 'senão, pega até o último dia do mês
        xdataper2 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 1, 4) & "/" & _
                    Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 5, 2) & "/" & UltDiaMes(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 5, 2), Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex), 1, 4)))
    End If
    
    If de_informa.rsSel_AlarmMovGer.State = 1 Then de_informa.rsSel_AlarmMovGer.Close
    de_informa.Sel_AlarmMovGer xdataper1, xdataper2, "AEREO"
        
    If de_informa.rsSel_AlarmMovGer.RecordCount > 0 Then
        flexAereoGer.Rows = de_informa.rsSel_AlarmMovGer.RecordCount + 1
        xTotValmerc = 0
        xTotFrete = 0
        xTotPeso = 0
        xTotVol = 0
        xTotCtc = 0
        xTotNf = 0
        For xLin = 1 To de_informa.rsSel_AlarmMovGer.RecordCount
            flexAereoGer.TextMatrix(xLin, 1) = de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexAereoGer.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmMovGer.Fields("nomefilial")
            flexAereoGer.TextMatrix(xLin, 3) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "###,###,###,##0.00")
            flexAereoGer.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete"), "###,###,##0.00")
            flexAereoGer.TextMatrix(xLin, 5) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete") / de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "##0.000%")
            flexAereoGer.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmMovGer.Fields("tpeso"), "###,###,##0.0")
            flexAereoGer.TextMatrix(xLin, 7) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvol"), "###,###,##0")
            flexAereoGer.TextMatrix(xLin, 8) = Format(de_informa.rsSel_AlarmMovGer.Fields("qtd"), "###,###,##0")
            If de_informa.rsSel_AlarmMovGerNFS.State = 1 Then de_informa.rsSel_AlarmMovGerNFS.Close
            'busca qtde de nfs de cada filial
            de_informa.Sel_AlarmMovGerNFS xdataper1, xdataper2, "AEREO", de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexAereoGer.TextMatrix(xLin, 9) = Format(de_informa.rsSel_AlarmMovGerNFS.Fields("qtd"), "###,###,##0")
            
            xTotValmerc = xTotValmerc + de_informa.rsSel_AlarmMovGer.Fields("tvalmerc")
            xTotFrete = xTotFrete + de_informa.rsSel_AlarmMovGer.Fields("tfrete")
            xTotPeso = xTotPeso + de_informa.rsSel_AlarmMovGer.Fields("tpeso")
            xTotVol = xTotVol + de_informa.rsSel_AlarmMovGer.Fields("tvol")
            xTotCtc = xTotCtc + de_informa.rsSel_AlarmMovGer.Fields("qtd")
            xTotNf = xTotNf + de_informa.rsSel_AlarmMovGerNFS.Fields("qtd")
            
            de_informa.rsSel_AlarmMovGer.MoveNext
            DoEvents
        Next
        
        flexAereoGerTot.TextMatrix(0, 2) = "SUB-TOTAL ............"
        flexAereoGerTot.TextMatrix(0, 3) = Format(xTotValmerc, "###,###,###,##0.00")
        flexAereoGerTot.TextMatrix(0, 4) = Format(xTotFrete, "###,###,##0.00")
        flexAereoGerTot.TextMatrix(0, 5) = Format(xTotFrete / xTotValmerc, "##0.000%")
        flexAereoGerTot.TextMatrix(0, 6) = Format(xTotPeso, "###,###,##0.0")
        flexAereoGerTot.TextMatrix(0, 7) = Format(xTotVol, "###,###,##0")
        flexAereoGerTot.TextMatrix(0, 8) = Format(xTotCtc, "###,###,##0")
        flexAereoGerTot.TextMatrix(0, 9) = Format(xTotNf, "###,###,##0")
        
        DoEvents
    End If

'MONTANDO O RODO
    
    If de_informa.rsSel_AlarmMovGer.State = 1 Then de_informa.rsSel_AlarmMovGer.Close
    de_informa.Sel_AlarmMovGer xdataper1, xdataper2, "RODOVIARIO"
    
    
    
    If de_informa.rsSel_AlarmMovGer.RecordCount > 0 Then
        flexRodoGer.Rows = de_informa.rsSel_AlarmMovGer.RecordCount + 1
        xTotValmerc = 0
        xTotFrete = 0
        xTotPeso = 0
        xTotVol = 0
        xTotCtc = 0
        xTotNf = 0
        For xLin = 1 To de_informa.rsSel_AlarmMovGer.RecordCount
            flexRodoGer.TextMatrix(xLin, 1) = de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexRodoGer.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmMovGer.Fields("nomefilial")
            flexRodoGer.TextMatrix(xLin, 3) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "###,###,###,##0.00")
            flexRodoGer.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete"), "###,###,##0.00")
            flexRodoGer.TextMatrix(xLin, 5) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete") / de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "##0.000%")
            flexRodoGer.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmMovGer.Fields("tpeso"), "###,###,##0.0")
            flexRodoGer.TextMatrix(xLin, 7) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvol"), "###,###,##0")
            flexRodoGer.TextMatrix(xLin, 8) = Format(de_informa.rsSel_AlarmMovGer.Fields("qtd"), "###,###,##0")
            If de_informa.rsSel_AlarmMovGerNFS.State = 1 Then de_informa.rsSel_AlarmMovGerNFS.Close
            'busca qtde de nfs de cada filial
            de_informa.Sel_AlarmMovGerNFS xdataper1, xdataper2, "RODOVIARIO", de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexRodoGer.TextMatrix(xLin, 9) = Format(de_informa.rsSel_AlarmMovGerNFS.Fields("qtd"), "###,###,##0")
            xTotValmerc = xTotValmerc + de_informa.rsSel_AlarmMovGer.Fields("tvalmerc")
            xTotFrete = xTotFrete + de_informa.rsSel_AlarmMovGer.Fields("tfrete")
            xTotPeso = xTotPeso + de_informa.rsSel_AlarmMovGer.Fields("tpeso")
            xTotVol = xTotVol + de_informa.rsSel_AlarmMovGer.Fields("tvol")
            xTotCtc = xTotCtc + de_informa.rsSel_AlarmMovGer.Fields("qtd")
            xTotNf = xTotNf + de_informa.rsSel_AlarmMovGerNFS.Fields("qtd")
            
            de_informa.rsSel_AlarmMovGer.MoveNext
            DoEvents
        Next
        
        flexRodoGerTot.TextMatrix(0, 2) = "SUB-TOTAL ............"
        flexRodoGerTot.TextMatrix(0, 3) = Format(xTotValmerc, "###,###,###,##0.00")
        flexRodoGerTot.TextMatrix(0, 4) = Format(xTotFrete, "###,###,##0.00")
        flexRodoGerTot.TextMatrix(0, 5) = Format(xTotFrete / xTotValmerc, "##0.000%")
        flexRodoGerTot.TextMatrix(0, 6) = Format(xTotPeso, "###,###,##0.0")
        flexRodoGerTot.TextMatrix(0, 7) = Format(xTotVol, "###,###,##0")
        flexRodoGerTot.TextMatrix(0, 8) = Format(xTotCtc, "###,###,##0")
        flexRodoGerTot.TextMatrix(0, 9) = Format(xTotNf, "###,###,##0")
        DoEvents
    End If

'MONTANDO O GERAL (RODO + AEREO)
    
    xTotValmerc = 0
    xTotFrete = 0
    xTotPeso = 0
    xTotVol = 0
    xTotCtc = 0
    xTotNf = 0
    
    If de_informa.rsSel_AlarmMovGer.State = 1 Then de_informa.rsSel_AlarmMovGer.Close
    de_informa.Sel_AlarmMovGer xdataper1, xdataper2, "%"
    
    If de_informa.rsSel_AlarmMovGer.RecordCount > 0 Then
        flexGeralGer.Rows = de_informa.rsSel_AlarmMovGer.RecordCount + 1
        For xLin = 1 To de_informa.rsSel_AlarmMovGer.RecordCount
            flexGeralGer.TextMatrix(xLin, 1) = de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexGeralGer.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmMovGer.Fields("nomefilial")
            flexGeralGer.TextMatrix(xLin, 3) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "###,###,###,##0.00")
            flexGeralGer.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete"), "###,###,##0.00")
            flexGeralGer.TextMatrix(xLin, 5) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete") / de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "##0.000%")
            flexGeralGer.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmMovGer.Fields("tpeso"), "###,###,##0.0")
            flexGeralGer.TextMatrix(xLin, 7) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvol"), "###,###,##0")
            flexGeralGer.TextMatrix(xLin, 8) = Format(de_informa.rsSel_AlarmMovGer.Fields("qtd"), "###,###,##0")
            If de_informa.rsSel_AlarmMovGerNFS.State = 1 Then de_informa.rsSel_AlarmMovGerNFS.Close
            'busca qtde de nfs de cada filial
            de_informa.Sel_AlarmMovGerNFS xdataper1, xdataper2, "%", de_informa.rsSel_AlarmMovGer.Fields("filial")
            flexGeralGer.TextMatrix(xLin, 9) = Format(de_informa.rsSel_AlarmMovGerNFS.Fields("qtd"), "###,###,##0")
            
            xTotValmerc = xTotValmerc + de_informa.rsSel_AlarmMovGer.Fields("tvalmerc")
            xTotFrete = xTotFrete + de_informa.rsSel_AlarmMovGer.Fields("tfrete")
            xTotPeso = xTotPeso + de_informa.rsSel_AlarmMovGer.Fields("tpeso")
            xTotVol = xTotVol + de_informa.rsSel_AlarmMovGer.Fields("tvol")
            xTotCtc = xTotCtc + de_informa.rsSel_AlarmMovGer.Fields("qtd")
            xTotNf = xTotNf + de_informa.rsSel_AlarmMovGerNFS.Fields("qtd")
            
            de_informa.rsSel_AlarmMovGer.MoveNext
            DoEvents
        Next

        flexGeralGerTot.TextMatrix(0, 2) = "TOTAL " & comboMesAnoGeral.Text & ".............."
        flexGeralGerTot.TextMatrix(0, 3) = Format(xTotValmerc, "###,###,###,##0.00")
        flexGeralGerTot.TextMatrix(0, 4) = Format(xTotFrete, "###,###,##0.00")
        flexGeralGerTot.TextMatrix(0, 5) = Format(xTotFrete / xTotValmerc, "##0.000%")
        flexGeralGerTot.TextMatrix(0, 6) = Format(xTotPeso, "###,###,##0.0")
        flexGeralGerTot.TextMatrix(0, 7) = Format(xTotVol, "###,###,##0")
        flexGeralGerTot.TextMatrix(0, 8) = Format(xTotCtc, "###,###,##0")
        flexGeralGerTot.TextMatrix(0, 9) = Format(xTotNf, "###,###,##0")
        DoEvents
    End If

    If comboMesAnoGeral.ListIndex = 0 Then

        'MONTANDO O GERAL (RODO + AEREO) / DO MÊS ANTERIOR
        
        xdataper1 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4) & "/" & _
                    Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2) & "/" & "01")
                    
        If IsDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4) & "/" & _
                  Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2) & "/" & _
                  Day(datahora("DATA"))) Then
            xdataper2 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4) & "/" & _
                        Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2) & "/" & Day(datahora("DATA")))
                  
        Else
            xdataper2 = CDate(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4) & "/" & _
                        Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2) & "/" & _
                        UltDiaMes(Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 5, 2), _
                                  Mid$(comboMesAnoGeral.ItemData(comboMesAnoGeral.ListIndex + 1), 1, 4)))
        End If
        
        'tratar se este mes já estiver no dia 31 e no mês anterior não tiver esta data
        
        If de_informa.rsSel_AlarmMovGer.State = 1 Then de_informa.rsSel_AlarmMovGer.Close
        de_informa.Sel_AlarmMovGer xdataper1, xdataper2, "%"
        
        If de_informa.rsSel_AlarmMovGer.RecordCount > 0 Then
            flexGeralGerMesAnt.Rows = de_informa.rsSel_AlarmMovGer.RecordCount + 1
            xTotValMercMesAnt = 0
            xTotFreteMesAnt = 0
            xTotPesoMesAnt = 0
            xTotVolMesAnt = 0
            xTotCtcMesAnt = 0
            xTotNfMesAnt = 0
            For xLin = 1 To de_informa.rsSel_AlarmMovGer.RecordCount
                flexGeralGerMesAnt.TextMatrix(xLin, 1) = de_informa.rsSel_AlarmMovGer.Fields("filial")
                flexGeralGerMesAnt.TextMatrix(xLin, 2) = de_informa.rsSel_AlarmMovGer.Fields("nomefilial")
                flexGeralGerMesAnt.TextMatrix(xLin, 3) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "###,###,###,##0.00")
                flexGeralGerMesAnt.TextMatrix(xLin, 4) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete"), "###,###,##0.00")
                flexGeralGerMesAnt.TextMatrix(xLin, 5) = Format(de_informa.rsSel_AlarmMovGer.Fields("tfrete") / de_informa.rsSel_AlarmMovGer.Fields("tvalmerc"), "##0.000%")
                flexGeralGerMesAnt.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmMovGer.Fields("tpeso"), "###,###,##0.0")
                flexGeralGerMesAnt.TextMatrix(xLin, 7) = Format(de_informa.rsSel_AlarmMovGer.Fields("tvol"), "###,###,##0")
                flexGeralGerMesAnt.TextMatrix(xLin, 8) = Format(de_informa.rsSel_AlarmMovGer.Fields("qtd"), "###,###,##0")
                If de_informa.rsSel_AlarmMovGerNFS.State = 1 Then de_informa.rsSel_AlarmMovGerNFS.Close
                'busca qtde de nfs de cada filial
                de_informa.Sel_AlarmMovGerNFS xdataper1, xdataper2, "%", de_informa.rsSel_AlarmMovGer.Fields("filial")
                flexGeralGerMesAnt.TextMatrix(xLin, 9) = Format(de_informa.rsSel_AlarmMovGerNFS.Fields("qtd"), "###,###,##0")
                
                xTotValMercMesAnt = xTotValMercMesAnt + de_informa.rsSel_AlarmMovGer.Fields("tvalmerc")
                xTotFreteMesAnt = xTotFreteMesAnt + de_informa.rsSel_AlarmMovGer.Fields("tfrete")
                xTotPesoMesAnt = xTotPesoMesAnt + de_informa.rsSel_AlarmMovGer.Fields("tpeso")
                xTotVolMesAnt = xTotVolMesAnt + de_informa.rsSel_AlarmMovGer.Fields("tvol")
                xTotCtcMesAnt = xTotCtcMesAnt + de_informa.rsSel_AlarmMovGer.Fields("qtd")
                xTotNfMesAnt = xTotNfMesAnt + de_informa.rsSel_AlarmMovGerNFS.Fields("qtd")
                
                de_informa.rsSel_AlarmMovGer.MoveNext
                DoEvents
                DoEvents
            Next
            
            flexGeralGerTotMesAnt.TextMatrix(0, 2) = "TOTAL " & comboMesAnoGeral.List(1) & ".............."
            flexGeralGerTotMesAnt.TextMatrix(0, 3) = Format(xTotValMercMesAnt, "###,###,###,##0.00")
            flexGeralGerTotMesAnt.TextMatrix(0, 4) = Format(xTotFreteMesAnt, "###,###,##0.00")
            flexGeralGerTotMesAnt.TextMatrix(0, 5) = Format(xTotFreteMesAnt / xTotValMercMesAnt, "##0.000%")
            flexGeralGerTotMesAnt.TextMatrix(0, 6) = Format(xTotPesoMesAnt, "###,###,##0.0")
            flexGeralGerTotMesAnt.TextMatrix(0, 7) = Format(xTotVolMesAnt, "###,###,##0")
            flexGeralGerTotMesAnt.TextMatrix(0, 8) = Format(xTotCtcMesAnt, "###,###,##0")
            flexGeralGerTotMesAnt.TextMatrix(0, 9) = Format(xTotNfMesAnt, "###,###,##0")
            DoEvents
        End If
        
        flexVarPercent.TextMatrix(0, 2) = "VARIAÇÃO (%).........."
        flexVarPercent.TextMatrix(0, 3) = Format((xTotValmerc - xTotValMercMesAnt) / xTotValMercMesAnt, "##0.00%")
        flexVarPercent.TextMatrix(0, 4) = Format((xTotFrete - xTotFreteMesAnt) / xTotFreteMesAnt, "##0.00%")
        flexVarPercent.TextMatrix(0, 6) = Format((xTotPeso - xTotPesoMesAnt) / xTotPesoMesAnt, "##0.00%")
        flexVarPercent.TextMatrix(0, 7) = Format((xTotVol - xTotVolMesAnt) / xTotVolMesAnt, "##0.00%")
        flexVarPercent.TextMatrix(0, 8) = Format((xTotCtc - xTotCtcMesAnt) / xTotCtcMesAnt, "##0.00%")
        flexVarPercent.TextMatrix(0, 9) = Format((xTotNf - xTotNfMesAnt) / xTotNfMesAnt, "##0.00%")
        
        lblTexto.Caption = "Comparação do Mês Atual (" & comboMesAnoGeral.List(0) & ") Com o Mês Anterior (" & comboMesAnoGeral.List(1) & ")  no Mesmo Período."
        cmdDetalhe.Enabled = True
        
    Else
    
        cmdDetalhe.Enabled = False
        
    End If
    
    Me.MousePointer = 0
    comboMesAnoGeral.Enabled = True
    cmdProcessarGerGeral.Enabled = True
    cmdSairGer.Enabled = True
    SSTab1.Enabled = True
    flexAereoGer.Enabled = True
    flexRodoGer.Enabled = True
    flexGeralGer.Enabled = True
    lblMensagem.Caption = comboMesAnoGeral.Text
    comboMesAnoGeral.Enabled = True
    cmdProcessarGerGeral.Enabled = True
    cmdImprGer1.Enabled = True

End Sub

Private Sub cmdSairGer_Click()
    de_informa.ins_LogUsuario "SAIR", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    Unload Me
    
End Sub

Private Sub cmdSairGerencial_Click()
    de_informa.ins_LogUsuario "SAIR", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    Unload Me
End Sub

Private Sub cmdSairPri_Click()
    de_informa.ins_LogUsuario "SAIR", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    Unload Me
End Sub

Private Sub cmdSairRes_Click()
    de_informa.ins_LogUsuario "SAIR", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    Unload Me
End Sub

Private Sub cmdSairUrg_Click()
    de_informa.ins_LogUsuario "SAIR", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    Unload Me
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command9_Click()

End Sub

Private Sub DataGrid3_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub flexNorAnoMes_Click()
Dim xano As Integer, xmes As Integer, xLin As Integer, xuf As String
    
    'Movto Normal - Por Região
    
    Me.MousePointer = 11
    
    flexNorRegiao.Clear
    
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    CmdProcessar.Enabled = False
    fraResUrg.Enabled = False
    fraResPri.Enabled = False
    fraResNor.Enabled = False
    
    flexNorRegiao.Cols = 7
    flexNorRegiao.ColWidth(0) = 200
    flexNorRegiao.ColWidth(1) = 400
    flexNorRegiao.ColWidth(2) = 1000
    flexNorRegiao.ColWidth(3) = 3740
    flexNorRegiao.ColWidth(4) = 670
    flexNorRegiao.ColWidth(5) = 550
    flexNorRegiao.ColWidth(6) = 580
    
    flexNorRegiao.TextMatrix(0, 1) = "Cód."
    flexNorRegiao.TextMatrix(0, 2) = "  Respons."
    flexNorRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexNorRegiao.TextMatrix(0, 4) = " CTCs"
    flexNorRegiao.TextMatrix(0, 5) = "Pend."
    flexNorRegiao.TextMatrix(0, 6) = "  %"
    
    DoEvents
    
    If de_informa.rsSel_AcompInfRegioes.State = 1 Then de_informa.rsSel_AcompInfRegioes.Close
    de_informa.Sel_AcompInfRegioes
    
    If de_informa.rsSel_AcompInfRegioes.RecordCount > 0 Then
        flexNorRegiao.Rows = de_informa.rsSel_AcompInfRegioes.RecordCount + 1
    Else
        MsgBox "Não Foi Possivel Trazer Dados por Região !"
        Exit Sub
    End If
    
    flexNorAnoMes.Col = 5
    xmes = flexNorAnoMes.Text
    flexNorAnoMes.Col = 6
    xano = flexNorAnoMes.Text
    
    xLin = 1
    
    Do Until de_informa.rsSel_AcompInfRegioes.EOF
        If de_informa.rsSel_AcompInfUFs.State = 1 Then de_informa.rsSel_AcompInfUFs.Close
        de_informa.Sel_AcompInfUFs de_informa.rsSel_AcompInfRegioes.Fields("regiaosac")
        
        xuf = ""
        Do Until de_informa.rsSel_AcompInfUFs.EOF
            If Len(xuf) > 1 Then
                xuf = xuf & ","
            End If
            xuf = xuf & de_informa.rsSel_AcompInfUFs.Fields("uf")
            de_informa.rsSel_AcompInfUFs.MoveNext
        Loop
        
        If de_informa.rsSel_AlarmResRegCtcs.State = 1 Then de_informa.rsSel_AlarmResRegCtcs.Close
        de_informa.Sel_AlarmResRegCtcs "NORMAL", de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xano, xmes
        
        If de_informa.rsSel_AlarmResRegPend.State = 1 Then de_informa.rsSel_AlarmResRegPend.Close
        de_informa.Sel_AlarmResRegPend "NORMAL", de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xano, xmes, datahora("data")
        
        flexNorRegiao.TextMatrix(xLin, 1) = de_informa.rsSel_AcompInfRegioes.Fields("regiaosac")
        flexNorRegiao.TextMatrix(xLin, 2) = de_informa.rsSel_AcompInfRegioes.Fields("atendsac")
        flexNorRegiao.TextMatrix(xLin, 3) = xuf
        flexNorRegiao.TextMatrix(xLin, 4) = de_informa.rsSel_AlarmResRegCtcs.Fields("qtd")
        flexNorRegiao.TextMatrix(xLin, 5) = de_informa.rsSel_AlarmResRegPend.Fields("qtd")
        If de_informa.rsSel_AlarmResRegCtcs.Fields("qtd") > 0 Then
            flexNorRegiao.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmResRegPend.Fields("qtd") / de_informa.rsSel_AlarmResRegCtcs.Fields("qtd"), "##0.0%")
        Else
            flexNorRegiao.TextMatrix(xLin, 6) = Format(0, "##0.0%")
        End If
        
        de_informa.rsSel_AcompInfRegioes.MoveNext
        xLin = xLin + 1
        
        DoEvents
        
    Loop
    
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
    CmdProcessar.Enabled = True
    fraResUrg.Enabled = True
    fraResPri.Enabled = True
    fraResNor.Enabled = True
    
    Me.MousePointer = 0

End Sub

Private Sub flexPriAnoMes_Click()
Dim xano As Integer, xmes As Integer, xLin As Integer, xuf As String

    'Prioridades - Por Região
    
    Me.MousePointer = 11
    
    flexPriRegiao.Clear
    
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    CmdProcessar.Enabled = False
    fraResUrg.Enabled = False
    fraResPri.Enabled = False
    fraResNor.Enabled = False

    flexPriRegiao.Cols = 7
    flexPriRegiao.ColWidth(0) = 200
    flexPriRegiao.ColWidth(1) = 400
    flexPriRegiao.ColWidth(2) = 1000
    flexPriRegiao.ColWidth(3) = 3740
    flexPriRegiao.ColWidth(4) = 670
    flexPriRegiao.ColWidth(5) = 550
    flexPriRegiao.ColWidth(6) = 580
    
    flexPriRegiao.TextMatrix(0, 1) = "Cód."
    flexPriRegiao.TextMatrix(0, 2) = "  Respons."
    flexPriRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexPriRegiao.TextMatrix(0, 4) = " CTCs"
    flexPriRegiao.TextMatrix(0, 5) = "Pend."
    flexPriRegiao.TextMatrix(0, 6) = "  %"

    DoEvents

    If de_informa.rsSel_AcompInfRegioes.State = 1 Then de_informa.rsSel_AcompInfRegioes.Close
    de_informa.Sel_AcompInfRegioes
    
    If de_informa.rsSel_AcompInfRegioes.RecordCount > 0 Then
        flexPriRegiao.Rows = de_informa.rsSel_AcompInfRegioes.RecordCount + 1
    Else
        MsgBox "Não Foi Possivel Trazer Dados por Região !"
        Exit Sub
    End If
    
    flexPriAnoMes.Col = 5
    xmes = flexPriAnoMes.Text
    flexPriAnoMes.Col = 6
    xano = flexPriAnoMes.Text
    
    xLin = 1
    
    Do Until de_informa.rsSel_AcompInfRegioes.EOF
        If de_informa.rsSel_AcompInfUFs.State = 1 Then de_informa.rsSel_AcompInfUFs.Close
        de_informa.Sel_AcompInfUFs de_informa.rsSel_AcompInfRegioes.Fields("regiaosac")
        
        xuf = ""
        Do Until de_informa.rsSel_AcompInfUFs.EOF
            If Len(xuf) > 1 Then
                xuf = xuf & ","
            End If
            xuf = xuf & de_informa.rsSel_AcompInfUFs.Fields("uf")
            de_informa.rsSel_AcompInfUFs.MoveNext
        Loop
        
        If de_informa.rsSel_AlarmResRegCtcs.State = 1 Then de_informa.rsSel_AlarmResRegCtcs.Close
        de_informa.Sel_AlarmResRegCtcs "PRIORIDADE", de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xano, xmes
        
        If de_informa.rsSel_AlarmResRegPend.State = 1 Then de_informa.rsSel_AlarmResRegPend.Close
        de_informa.Sel_AlarmResRegPend "PRIORIDADE", de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xano, xmes, datahora("data")
        
        flexPriRegiao.TextMatrix(xLin, 1) = de_informa.rsSel_AcompInfRegioes.Fields("regiaosac")
        flexPriRegiao.TextMatrix(xLin, 2) = de_informa.rsSel_AcompInfRegioes.Fields("atendsac")
        flexPriRegiao.TextMatrix(xLin, 3) = xuf
        flexPriRegiao.TextMatrix(xLin, 4) = de_informa.rsSel_AlarmResRegCtcs.Fields("qtd")
        flexPriRegiao.TextMatrix(xLin, 5) = de_informa.rsSel_AlarmResRegPend.Fields("qtd")
        If de_informa.rsSel_AlarmResRegCtcs.Fields("qtd") > 0 Then
            flexPriRegiao.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmResRegPend.Fields("qtd") / de_informa.rsSel_AlarmResRegCtcs.Fields("qtd"), "##0.0%")
        Else
            flexPriRegiao.TextMatrix(xLin, 6) = Format(0, "##0.0%")
        End If
        
        de_informa.rsSel_AcompInfRegioes.MoveNext
        xLin = xLin + 1
        
        DoEvents
        
    Loop
    
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
    CmdProcessar.Enabled = True
    fraResUrg.Enabled = True
    fraResPri.Enabled = True
    fraResNor.Enabled = True
    
    Me.MousePointer = 0

End Sub

Private Sub flexUrgAnoMes_Click()
Dim xano As Integer, xmes As Integer, xLin As Integer, xuf As String

    'Urgências - Por Região
    
    Me.MousePointer = 11
    
    flexUrgRegiao.Clear
    
    cmdImprTelaUrg.Enabled = False
    cmdImprListUrg.Enabled = False
    cmdSairUrg.Enabled = False
    cmdImprTelaPri.Enabled = False
    cmdImprListPri.Enabled = False
    cmdSairPri.Enabled = False
    cmdImprTelaRes.Enabled = False
    cmdSairRes.Enabled = False
    CmdProcessar.Enabled = False
    fraResUrg.Enabled = False
    fraResPri.Enabled = False
    fraResNor.Enabled = False
    
    flexUrgRegiao.Cols = 7
    flexUrgRegiao.ColWidth(0) = 200
    flexUrgRegiao.ColWidth(1) = 400
    flexUrgRegiao.ColWidth(2) = 1000
    flexUrgRegiao.ColWidth(3) = 3740
    flexUrgRegiao.ColWidth(4) = 670
    flexUrgRegiao.ColWidth(5) = 550
    flexUrgRegiao.ColWidth(6) = 580
    
    flexUrgRegiao.TextMatrix(0, 1) = "Cód."
    flexUrgRegiao.TextMatrix(0, 2) = "  Respons."
    flexUrgRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexUrgRegiao.TextMatrix(0, 4) = " CTCs"
    flexUrgRegiao.TextMatrix(0, 5) = "Pend."
    flexUrgRegiao.TextMatrix(0, 6) = "  %"

    DoEvents
    
    If de_informa.rsSel_AcompInfRegioes.State = 1 Then de_informa.rsSel_AcompInfRegioes.Close
    de_informa.Sel_AcompInfRegioes
    
    If de_informa.rsSel_AcompInfRegioes.RecordCount > 0 Then
        flexUrgRegiao.Rows = de_informa.rsSel_AcompInfRegioes.RecordCount + 1
    Else
        MsgBox "Não Foi Possivel Trazer Dados por Região !"
        Exit Sub
    End If
    
    flexUrgAnoMes.Col = 5
    xmes = flexUrgAnoMes.Text
    flexUrgAnoMes.Col = 6
    xano = flexUrgAnoMes.Text
    
    xLin = 1
    
    Do Until de_informa.rsSel_AcompInfRegioes.EOF
        If de_informa.rsSel_AcompInfUFs.State = 1 Then de_informa.rsSel_AcompInfUFs.Close
        de_informa.Sel_AcompInfUFs de_informa.rsSel_AcompInfRegioes.Fields("regiaosac")
        
        xuf = ""
        Do Until de_informa.rsSel_AcompInfUFs.EOF
            If Len(xuf) > 1 Then
                xuf = xuf & ","
            End If
            xuf = xuf & de_informa.rsSel_AcompInfUFs.Fields("uf")
            de_informa.rsSel_AcompInfUFs.MoveNext
        Loop
        
        If de_informa.rsSel_AlarmResRegCtcs.State = 1 Then de_informa.rsSel_AlarmResRegCtcs.Close
        de_informa.Sel_AlarmResRegCtcs "URGÊNCIA", de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xano, xmes
        
        If de_informa.rsSel_AlarmResRegPend.State = 1 Then de_informa.rsSel_AlarmResRegPend.Close
        de_informa.Sel_AlarmResRegPend "URGÊNCIA", de_informa.rsSel_AcompInfRegioes.Fields("regiaosac"), xano, xmes, CDate(datahora("data")) + 100
        
        flexUrgRegiao.TextMatrix(xLin, 1) = de_informa.rsSel_AcompInfRegioes.Fields("regiaosac")
        flexUrgRegiao.TextMatrix(xLin, 2) = de_informa.rsSel_AcompInfRegioes.Fields("atendsac")
        flexUrgRegiao.TextMatrix(xLin, 3) = xuf
        flexUrgRegiao.TextMatrix(xLin, 4) = de_informa.rsSel_AlarmResRegCtcs.Fields("qtd")
        flexUrgRegiao.TextMatrix(xLin, 5) = de_informa.rsSel_AlarmResRegPend.Fields("qtd")
        If de_informa.rsSel_AlarmResRegCtcs.Fields("qtd") > 0 Then
            flexUrgRegiao.TextMatrix(xLin, 6) = Format(de_informa.rsSel_AlarmResRegPend.Fields("qtd") / de_informa.rsSel_AlarmResRegCtcs.Fields("qtd"), "##0.0%")
        Else
            flexUrgRegiao.TextMatrix(xLin, 6) = Format(0, "##0.0%")
        End If
        
        de_informa.rsSel_AcompInfRegioes.MoveNext
        xLin = xLin + 1
        
        DoEvents
        
    Loop
    
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
    CmdProcessar.Enabled = True
    fraResUrg.Enabled = True
    fraResPri.Enabled = True
    fraResNor.Enabled = True

    Me.MousePointer = 0
    
End Sub

Private Sub Form_Load()
    Dim PosWin As Long
    PosWin = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    'Configura as Flex
    
    'Urgencias - Ano/mes
    
    flexUrgAnoMes.Cols = 7
    flexUrgAnoMes.ColWidth(0) = 200
    flexUrgAnoMes.ColWidth(1) = 830
    flexUrgAnoMes.ColWidth(2) = 670
    flexUrgAnoMes.ColWidth(3) = 550
    flexUrgAnoMes.ColWidth(4) = 580
    flexUrgAnoMes.ColWidth(5) = 1
    flexUrgAnoMes.ColWidth(6) = 1
    
    flexUrgAnoMes.TextMatrix(0, 1) = "Mês/Ano"
    flexUrgAnoMes.TextMatrix(0, 2) = " CTCs"
    flexUrgAnoMes.TextMatrix(0, 3) = "Pend."
    flexUrgAnoMes.TextMatrix(0, 4) = "  %"
    
    'Prioridades - Ano/mes
    
    flexPriAnoMes.Cols = 7
    flexPriAnoMes.ColWidth(0) = 200
    flexPriAnoMes.ColWidth(1) = 830
    flexPriAnoMes.ColWidth(2) = 670
    flexPriAnoMes.ColWidth(3) = 550
    flexPriAnoMes.ColWidth(4) = 580
    flexPriAnoMes.ColWidth(5) = 1
    flexPriAnoMes.ColWidth(6) = 1
    
    
    flexPriAnoMes.TextMatrix(0, 1) = "Mês/Ano"
    flexPriAnoMes.TextMatrix(0, 2) = " CTCs"
    flexPriAnoMes.TextMatrix(0, 3) = "Pend."
    flexPriAnoMes.TextMatrix(0, 4) = "  %"
    
    'Movto Normal - Ano/mes
    
    flexNorAnoMes.Cols = 7
    flexNorAnoMes.ColWidth(0) = 200
    flexNorAnoMes.ColWidth(1) = 830
    flexNorAnoMes.ColWidth(2) = 670
    flexNorAnoMes.ColWidth(3) = 550
    flexNorAnoMes.ColWidth(4) = 580
    flexNorAnoMes.ColWidth(5) = 1
    flexNorAnoMes.ColWidth(6) = 1
    
    flexNorAnoMes.TextMatrix(0, 1) = "Mês/Ano"
    flexNorAnoMes.TextMatrix(0, 2) = " CTCs"
    flexNorAnoMes.TextMatrix(0, 3) = "Pend."
    flexNorAnoMes.TextMatrix(0, 4) = "  %"
    
    'Urgências - Por Região
    
    flexUrgRegiao.Cols = 7
    flexUrgRegiao.ColWidth(0) = 200
    flexUrgRegiao.ColWidth(1) = 400
    flexUrgRegiao.ColWidth(2) = 1000
    flexUrgRegiao.ColWidth(3) = 3740
    flexUrgRegiao.ColWidth(4) = 670
    flexUrgRegiao.ColWidth(5) = 550
    flexUrgRegiao.ColWidth(6) = 580
    
    flexUrgRegiao.TextMatrix(0, 1) = "Cód."
    flexUrgRegiao.TextMatrix(0, 2) = "  Respons."
    flexUrgRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexUrgRegiao.TextMatrix(0, 4) = " CTCs"
    flexUrgRegiao.TextMatrix(0, 5) = "Pend."
    flexUrgRegiao.TextMatrix(0, 6) = "  %"
    
    'Prioridades - Por Região
    
    flexPriRegiao.Cols = 7
    flexPriRegiao.ColWidth(0) = 200
    flexPriRegiao.ColWidth(1) = 400
    flexPriRegiao.ColWidth(2) = 1000
    flexPriRegiao.ColWidth(3) = 3740
    flexPriRegiao.ColWidth(4) = 670
    flexPriRegiao.ColWidth(5) = 550
    flexPriRegiao.ColWidth(6) = 580
    
    flexPriRegiao.TextMatrix(0, 1) = "Cód."
    flexPriRegiao.TextMatrix(0, 2) = "  Respons."
    flexPriRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexPriRegiao.TextMatrix(0, 4) = " CTCs"
    flexPriRegiao.TextMatrix(0, 5) = "Pend."
    flexPriRegiao.TextMatrix(0, 6) = "  %"
    
    'Movto Normal - Por Região
    
    flexNorRegiao.Cols = 7
    flexNorRegiao.ColWidth(0) = 200
    flexNorRegiao.ColWidth(1) = 400
    flexNorRegiao.ColWidth(2) = 1000
    flexNorRegiao.ColWidth(3) = 3740
    flexNorRegiao.ColWidth(4) = 670
    flexNorRegiao.ColWidth(5) = 550
    flexNorRegiao.ColWidth(6) = 580
    
    flexNorRegiao.TextMatrix(0, 1) = "Cód."
    flexNorRegiao.TextMatrix(0, 2) = "  Respons."
    flexNorRegiao.TextMatrix(0, 3) = "                            Região/UFs"
    flexNorRegiao.TextMatrix(0, 4) = " CTCs"
    flexNorRegiao.TextMatrix(0, 5) = "Pend."
    flexNorRegiao.TextMatrix(0, 6) = "  %"

    'gerencial - movto geral - aereo
    flexAereoGer.Cols = 10
    flexAereoGer.ColWidth(0) = 150
    flexAereoGer.ColWidth(1) = 400
    flexAereoGer.ColWidth(2) = 1600
    flexAereoGer.ColWidth(3) = 1700
    flexAereoGer.ColWidth(4) = 1400
    flexAereoGer.ColWidth(5) = 970
    flexAereoGer.ColWidth(6) = 1200
    flexAereoGer.ColWidth(7) = 900
    flexAereoGer.ColWidth(8) = 900
    flexAereoGer.ColWidth(9) = 900
    
    flexAereoGer.TextMatrix(0, 1) = "Filial"
    flexAereoGer.TextMatrix(0, 2) = "Nome Filial"
    flexAereoGer.TextMatrix(0, 3) = "Vlr. Mercadoria"
    flexAereoGer.TextMatrix(0, 4) = "Frete Líquido"
    flexAereoGer.TextMatrix(0, 5) = "Frete/Valor"
    flexAereoGer.TextMatrix(0, 6) = "Peso"
    flexAereoGer.TextMatrix(0, 7) = "Volumes"
    flexAereoGer.TextMatrix(0, 8) = "CTCs"
    flexAereoGer.TextMatrix(0, 9) = "NFs"
    
    flexAereoGerTot.Cols = 10
    flexAereoGerTot.ColWidth(0) = 150
    flexAereoGerTot.ColWidth(1) = 400
    flexAereoGerTot.ColWidth(2) = 1600
    flexAereoGerTot.ColWidth(3) = 1700
    flexAereoGerTot.ColWidth(4) = 1400
    flexAereoGerTot.ColWidth(5) = 970
    flexAereoGerTot.ColWidth(6) = 1200
    flexAereoGerTot.ColWidth(7) = 900
    flexAereoGerTot.ColWidth(8) = 900
    flexAereoGerTot.ColWidth(9) = 900

    'gerencial - movto geral - rodo
    flexRodoGer.Cols = 10
    flexRodoGer.ColWidth(0) = 150
    flexRodoGer.ColWidth(1) = 400
    flexRodoGer.ColWidth(2) = 1600
    flexRodoGer.ColWidth(3) = 1700
    flexRodoGer.ColWidth(4) = 1400
    flexRodoGer.ColWidth(5) = 970
    flexRodoGer.ColWidth(6) = 1200
    flexRodoGer.ColWidth(7) = 900
    flexRodoGer.ColWidth(8) = 900
    flexRodoGer.ColWidth(9) = 900
    
    flexRodoGer.TextMatrix(0, 1) = "Filial"
    flexRodoGer.TextMatrix(0, 2) = "Nome Filial"
    flexRodoGer.TextMatrix(0, 3) = "Vlr. Mercadoria"
    flexRodoGer.TextMatrix(0, 4) = "Frete Líquido"
    flexRodoGer.TextMatrix(0, 5) = "Frete/Valor"
    flexRodoGer.TextMatrix(0, 6) = "Peso"
    flexRodoGer.TextMatrix(0, 7) = "Volumes"
    flexRodoGer.TextMatrix(0, 8) = "CTCs"
    flexRodoGer.TextMatrix(0, 9) = "NFs"

    flexRodoGerTot.Cols = 10
    flexRodoGerTot.ColWidth(0) = 150
    flexRodoGerTot.ColWidth(1) = 400
    flexRodoGerTot.ColWidth(2) = 1600
    flexRodoGerTot.ColWidth(3) = 1700
    flexRodoGerTot.ColWidth(4) = 1400
    flexRodoGerTot.ColWidth(5) = 970
    flexRodoGerTot.ColWidth(6) = 1200
    flexRodoGerTot.ColWidth(7) = 900
    flexRodoGerTot.ColWidth(8) = 900
    flexRodoGerTot.ColWidth(9) = 900

    'gerencial - movto geral - rodo
    flexGeralGer.Cols = 10
    flexGeralGer.ColWidth(0) = 150
    flexGeralGer.ColWidth(1) = 400
    flexGeralGer.ColWidth(2) = 1600
    flexGeralGer.ColWidth(3) = 1700
    flexGeralGer.ColWidth(4) = 1400
    flexGeralGer.ColWidth(5) = 970
    flexGeralGer.ColWidth(6) = 1200
    flexGeralGer.ColWidth(7) = 900
    flexGeralGer.ColWidth(8) = 900
    flexGeralGer.ColWidth(9) = 900
    
    flexGeralGer.TextMatrix(0, 1) = "Filial"
    flexGeralGer.TextMatrix(0, 2) = "Nome Filial"
    flexGeralGer.TextMatrix(0, 3) = "Vlr. Mercadoria"
    flexGeralGer.TextMatrix(0, 4) = "Frete Líquido"
    flexGeralGer.TextMatrix(0, 5) = "Frete/Valor"
    flexGeralGer.TextMatrix(0, 6) = "Peso"
    flexGeralGer.TextMatrix(0, 7) = "Volumes"
    flexGeralGer.TextMatrix(0, 8) = "CTCs"
    flexGeralGer.TextMatrix(0, 9) = "NFs"
    
    flexGeralGerTot.Cols = 10
    flexGeralGerTot.ColWidth(0) = 150
    flexGeralGerTot.ColWidth(1) = 400
    flexGeralGerTot.ColWidth(2) = 1600
    flexGeralGerTot.ColWidth(3) = 1700
    flexGeralGerTot.ColWidth(4) = 1400
    flexGeralGerTot.ColWidth(5) = 970
    flexGeralGerTot.ColWidth(6) = 1200
    flexGeralGerTot.ColWidth(7) = 900
    flexGeralGerTot.ColWidth(8) = 900
    flexGeralGerTot.ColWidth(9) = 900
    
    'gerencial - movto geral - rodo / mes anterior
    flexGeralGerMesAnt.Cols = 10
    flexGeralGerMesAnt.ColWidth(0) = 150
    flexGeralGerMesAnt.ColWidth(1) = 400
    flexGeralGerMesAnt.ColWidth(2) = 1600
    flexGeralGerMesAnt.ColWidth(3) = 1700
    flexGeralGerMesAnt.ColWidth(4) = 1400
    flexGeralGerMesAnt.ColWidth(5) = 970
    flexGeralGerMesAnt.ColWidth(6) = 1200
    flexGeralGerMesAnt.ColWidth(7) = 900
    flexGeralGerMesAnt.ColWidth(8) = 900
    flexGeralGerMesAnt.ColWidth(9) = 900
    
    flexGeralGerMesAnt.TextMatrix(0, 1) = "Filial"
    flexGeralGerMesAnt.TextMatrix(0, 2) = "Nome Filial"
    flexGeralGerMesAnt.TextMatrix(0, 3) = "Vlr. Mercadoria"
    flexGeralGerMesAnt.TextMatrix(0, 4) = "Frete Líquido"
    flexGeralGerMesAnt.TextMatrix(0, 5) = "Frete/Valor"
    flexGeralGerMesAnt.TextMatrix(0, 6) = "Peso"
    flexGeralGerMesAnt.TextMatrix(0, 7) = "Volumes"
    flexGeralGerMesAnt.TextMatrix(0, 8) = "CTCs"
    flexGeralGerMesAnt.TextMatrix(0, 9) = "NFs"
    
    flexGeralGerTotMesAnt.Cols = 10
    flexGeralGerTotMesAnt.ColWidth(0) = 150
    flexGeralGerTotMesAnt.ColWidth(1) = 400
    flexGeralGerTotMesAnt.ColWidth(2) = 1600
    flexGeralGerTotMesAnt.ColWidth(3) = 1700
    flexGeralGerTotMesAnt.ColWidth(4) = 1400
    flexGeralGerTotMesAnt.ColWidth(5) = 970
    flexGeralGerTotMesAnt.ColWidth(6) = 1200
    flexGeralGerTotMesAnt.ColWidth(7) = 900
    flexGeralGerTotMesAnt.ColWidth(8) = 900
    flexGeralGerTotMesAnt.ColWidth(9) = 900
    
    flexVarPercent.Cols = 1
    flexVarPercent.Cols = 10
    flexVarPercent.ColWidth(0) = 150
    flexVarPercent.ColWidth(1) = 400
    flexVarPercent.ColWidth(2) = 1600
    flexVarPercent.ColWidth(3) = 1700
    flexVarPercent.ColWidth(4) = 1400
    flexVarPercent.ColWidth(5) = 970
    flexVarPercent.ColWidth(6) = 1200
    flexVarPercent.ColWidth(7) = 900
    flexVarPercent.ColWidth(8) = 900
    flexVarPercent.ColWidth(9) = 900
   
    'gerencial - movto cliente - por modal
    FlexPorModal.Cols = 8
    FlexPorModal.ColWidth(0) = 600
    FlexPorModal.ColWidth(1) = 1400
    FlexPorModal.ColWidth(2) = 1400
    FlexPorModal.ColWidth(3) = 1000
    FlexPorModal.ColWidth(4) = 1100
    FlexPorModal.ColWidth(5) = 900
    FlexPorModal.ColWidth(6) = 900
    FlexPorModal.ColWidth(7) = 900

    FlexPorModal.TextMatrix(0, 1) = "Vlr. Mercadoria"
    FlexPorModal.TextMatrix(0, 2) = "Frete Líquido"
    FlexPorModal.TextMatrix(0, 3) = "Frete/Valor"
    FlexPorModal.TextMatrix(0, 4) = "Peso"
    FlexPorModal.TextMatrix(0, 5) = "Volumes"
    FlexPorModal.TextMatrix(0, 6) = "CTCs"
    FlexPorModal.TextMatrix(0, 7) = "NFs"
    FlexPorModal.TextMatrix(1, 0) = "Aéreo"
    FlexPorModal.TextMatrix(2, 0) = "Rodo"
    FlexPorModal.TextMatrix(3, 0) = "Total"
    
    'performance aéreo e rodoviario
    
    FlexPerfAir.Cols = 5
    FlexPerfAir.ColWidth(0) = 350
    FlexPerfAir.ColWidth(1) = 750
    FlexPerfAir.ColWidth(2) = 750
    FlexPerfAir.ColWidth(3) = 750
    FlexPerfAir.ColWidth(4) = 800
    
    FlexPerfAir.TextMatrix(0, 1) = "     %"
    FlexPerfAir.TextMatrix(0, 2) = "No Prazo"
    FlexPerfAir.TextMatrix(0, 3) = "  Atraso"
    FlexPerfAir.TextMatrix(0, 4) = "Total Entr"
    FlexPerfAir.TextMatrix(1, 0) = "BR"
    
    FlexPerfRodo.Cols = 5
    FlexPerfRodo.ColWidth(0) = 350
    FlexPerfRodo.ColWidth(1) = 750
    FlexPerfRodo.ColWidth(2) = 750
    FlexPerfRodo.ColWidth(3) = 750
    FlexPerfRodo.ColWidth(4) = 800
    
    FlexPerfRodo.TextMatrix(0, 1) = "     %"
    FlexPerfRodo.TextMatrix(0, 2) = "No Prazo"
    FlexPerfRodo.TextMatrix(0, 3) = "  Atraso"
    FlexPerfRodo.TextMatrix(0, 4) = "Total Entr"
    FlexPerfRodo.TextMatrix(1, 0) = "BR"
    
    'PERFORMANCE / por uf
    'AIR
    FlexPerfUfAir.Cols = 5
    FlexPerfUfAir.ColWidth(0) = 350
    FlexPerfUfAir.ColWidth(1) = 750
    FlexPerfUfAir.ColWidth(2) = 750
    FlexPerfUfAir.ColWidth(3) = 750
    FlexPerfUfAir.ColWidth(4) = 800
    
    FlexPerfUfAir.TextMatrix(0, 1) = "     %"
    FlexPerfUfAir.TextMatrix(0, 2) = "No Prazo"
    FlexPerfUfAir.TextMatrix(0, 3) = "  Atraso"
    FlexPerfUfAir.TextMatrix(0, 4) = "Total Entr"
    
    FlexPerfUfAir.TextMatrix(0, 0) = "UF"
    FlexPerfUfAir.TextMatrix(1, 0) = "AC"
    FlexPerfUfAir.TextMatrix(2, 0) = "AL"
    FlexPerfUfAir.TextMatrix(3, 0) = "AM"
    FlexPerfUfAir.TextMatrix(4, 0) = "AP"
    FlexPerfUfAir.TextMatrix(5, 0) = "BA"
    FlexPerfUfAir.TextMatrix(6, 0) = "CE"
    FlexPerfUfAir.TextMatrix(7, 0) = "DF"
    FlexPerfUfAir.TextMatrix(8, 0) = "ES"
    FlexPerfUfAir.TextMatrix(9, 0) = "GO"
    FlexPerfUfAir.TextMatrix(10, 0) = "MA"
    FlexPerfUfAir.TextMatrix(11, 0) = "MG"
    FlexPerfUfAir.TextMatrix(12, 0) = "MS"
    FlexPerfUfAir.TextMatrix(13, 0) = "MT"
    FlexPerfUfAir.TextMatrix(14, 0) = "PA"
    FlexPerfUfAir.TextMatrix(15, 0) = "PB"
    FlexPerfUfAir.TextMatrix(16, 0) = "PE"
    FlexPerfUfAir.TextMatrix(17, 0) = "PI"
    FlexPerfUfAir.TextMatrix(18, 0) = "PR"
    FlexPerfUfAir.TextMatrix(19, 0) = "RJ"
    FlexPerfUfAir.TextMatrix(20, 0) = "RN"
    FlexPerfUfAir.TextMatrix(21, 0) = "RO"
    FlexPerfUfAir.TextMatrix(22, 0) = "RR"
    FlexPerfUfAir.TextMatrix(23, 0) = "RS"
    FlexPerfUfAir.TextMatrix(24, 0) = "SC"
    FlexPerfUfAir.TextMatrix(25, 0) = "SE"
    FlexPerfUfAir.TextMatrix(26, 0) = "SP"
    FlexPerfUfAir.TextMatrix(27, 0) = "TO"
    
    'rodo
    FlexPerfUfRodo.Cols = 5
    FlexPerfUfRodo.ColWidth(0) = 350
    FlexPerfUfRodo.ColWidth(1) = 750
    FlexPerfUfRodo.ColWidth(2) = 750
    FlexPerfUfRodo.ColWidth(3) = 750
    FlexPerfUfRodo.ColWidth(4) = 800
    
    FlexPerfUfRodo.TextMatrix(0, 1) = "     %"
    FlexPerfUfRodo.TextMatrix(0, 2) = "No Prazo"
    FlexPerfUfRodo.TextMatrix(0, 3) = "  Atraso"
    FlexPerfUfRodo.TextMatrix(0, 4) = "Total Entr"
    
    FlexPerfUfRodo.TextMatrix(0, 0) = "UF"
    FlexPerfUfRodo.TextMatrix(1, 0) = "AC"
    FlexPerfUfRodo.TextMatrix(2, 0) = "AL"
    FlexPerfUfRodo.TextMatrix(3, 0) = "AM"
    FlexPerfUfRodo.TextMatrix(4, 0) = "AP"
    FlexPerfUfRodo.TextMatrix(5, 0) = "BA"
    FlexPerfUfRodo.TextMatrix(6, 0) = "CE"
    FlexPerfUfRodo.TextMatrix(7, 0) = "DF"
    FlexPerfUfRodo.TextMatrix(8, 0) = "ES"
    FlexPerfUfRodo.TextMatrix(9, 0) = "GO"
    FlexPerfUfRodo.TextMatrix(10, 0) = "MA"
    FlexPerfUfRodo.TextMatrix(11, 0) = "MG"
    FlexPerfUfRodo.TextMatrix(12, 0) = "MS"
    FlexPerfUfRodo.TextMatrix(13, 0) = "MT"
    FlexPerfUfRodo.TextMatrix(14, 0) = "PA"
    FlexPerfUfRodo.TextMatrix(15, 0) = "PB"
    FlexPerfUfRodo.TextMatrix(16, 0) = "PE"
    FlexPerfUfRodo.TextMatrix(17, 0) = "PI"
    FlexPerfUfRodo.TextMatrix(18, 0) = "PR"
    FlexPerfUfRodo.TextMatrix(19, 0) = "RJ"
    FlexPerfUfRodo.TextMatrix(20, 0) = "RN"
    FlexPerfUfRodo.TextMatrix(21, 0) = "RO"
    FlexPerfUfRodo.TextMatrix(22, 0) = "RR"
    FlexPerfUfRodo.TextMatrix(23, 0) = "RS"
    FlexPerfUfRodo.TextMatrix(24, 0) = "SC"
    FlexPerfUfRodo.TextMatrix(25, 0) = "SE"
    FlexPerfUfRodo.TextMatrix(26, 0) = "SP"
    FlexPerfUfRodo.TextMatrix(27, 0) = "TO"
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAlarmeUrg = Nothing
End Sub

Private Sub gridClientes_DblClick()
    
    frmAlarmeUrg.MousePointer = 11
    lblaguarde.Visible = True
    fraClientes.Enabled = False
    fraPorModal.Enabled = False
    fraPerfRodo.Enabled = False
    fraPerfAir.Enabled = False
    cmdPerfUFRodo.Enabled = False
    cmdPerfUFAir.Enabled = False
    cmdSairGerencial.Enabled = False
    comboMesAnoCliente.Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    cmdImprGer2.Enabled = False
    DoEvents
    lblCliente.Caption = gridClientes.Columns(1)
    
    FlexPorModal.Rows = 1
    FlexPorModal.Rows = 4
    FlexPorModal.FixedRows = 1
    FlexPorModal.TextMatrix(1, 0) = "Aéreo"
    FlexPorModal.TextMatrix(2, 0) = "Rodo"
    FlexPorModal.TextMatrix(3, 0) = "Total"
    
    FlexPerfAir.Rows = 1
    FlexPerfAir.Rows = 2
    FlexPerfAir.FixedRows = 1
    FlexPerfAir.TextMatrix(1, 0) = "BR"
    FlexPerfRodo.Rows = 1
    FlexPerfRodo.Rows = 2
    FlexPerfRodo.FixedRows = 1
    FlexPerfRodo.TextMatrix(1, 0) = "BR"
    
    For xcont = 1 To 27
        FlexPerfUfAir.TextMatrix(xcont, 1) = ""
        FlexPerfUfAir.TextMatrix(xcont, 2) = ""
        FlexPerfUfAir.TextMatrix(xcont, 3) = ""
        FlexPerfUfAir.TextMatrix(xcont, 4) = ""
        FlexPerfUfRodo.TextMatrix(xcont, 1) = ""
        FlexPerfUfRodo.TextMatrix(xcont, 2) = ""
        FlexPerfUfRodo.TextMatrix(xcont, 3) = ""
        FlexPerfUfRodo.TextMatrix(xcont, 4) = ""
    Next

    fraPorModal.Caption = "Por Modal - " & comboMesAnoCliente.Text
    DoEvents
    
'MONTANDO O AÉREO - cliente
    
    If de_informa.rsSel_AlarmMovCli.State = 1 Then de_informa.rsSel_AlarmMovCli.Close
    de_informa.Sel_AlarmMovCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%"
    
    
    If de_informa.rsSel_AlarmMovCli.RecordCount > 0 Then
        FlexPorModal.TextMatrix(1, 1) = Format(de_informa.rsSel_AlarmMovCli.Fields("tvalmerc"), "###,###,###,##0.00")
        FlexPorModal.TextMatrix(1, 2) = Format(de_informa.rsSel_AlarmMovCli.Fields("tfrete"), "###,###,##0.00")
        FlexPorModal.TextMatrix(1, 3) = Format(de_informa.rsSel_AlarmMovCli.Fields("tfrete") / de_informa.rsSel_AlarmMovCli.Fields("tvalmerc"), "##0.00%")
        FlexPorModal.TextMatrix(1, 4) = Format(de_informa.rsSel_AlarmMovCli.Fields("tpeso"), "###,###,##0.0")
        FlexPorModal.TextMatrix(1, 5) = Format(de_informa.rsSel_AlarmMovCli.Fields("tvol"), "###,###,##0")
        FlexPorModal.TextMatrix(1, 6) = Format(de_informa.rsSel_AlarmMovCli.Fields("qtd"), "###,###,##0")
        If de_informa.rsSel_AlarmMovCliNFS.State = 1 Then de_informa.rsSel_AlarmMovCliNFS.Close
        'busca qtde de nfs de cada filial
        de_informa.Sel_AlarmMovCliNFS Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                            Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                            "AEREO", Trim$(gridClientes.Columns(0)) & "%"
        FlexPorModal.TextMatrix(1, 7) = Format(de_informa.rsSel_AlarmMovCliNFS.Fields("qtd"), "###,###,##0")
            
        DoEvents
    Else
        FlexPorModal.TextMatrix(1, 1) = Format(0, "###,###,###,##0.00")
        FlexPorModal.TextMatrix(1, 2) = Format(0, "###,###,##0.00")
        FlexPorModal.TextMatrix(1, 3) = Format(0, "##0.00%")
        FlexPorModal.TextMatrix(1, 4) = Format(0, "###,###,##0.0")
        FlexPorModal.TextMatrix(1, 5) = Format(0, "###,###,##0")
        FlexPorModal.TextMatrix(1, 6) = Format(0, "###,###,##0")
        FlexPorModal.TextMatrix(1, 7) = Format(0, "###,###,##0")
        DoEvents
    End If
    
    
'MONTANDO O RODOVIÁRIO - Cliente
    
    If de_informa.rsSel_AlarmMovCli.State = 1 Then de_informa.rsSel_AlarmMovCli.Close
    de_informa.Sel_AlarmMovCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%"
    
    If de_informa.rsSel_AlarmMovCli.RecordCount > 0 Then
        FlexPorModal.TextMatrix(2, 1) = Format(de_informa.rsSel_AlarmMovCli.Fields("tvalmerc"), "###,###,###,##0.00")
        FlexPorModal.TextMatrix(2, 2) = Format(de_informa.rsSel_AlarmMovCli.Fields("tfrete"), "###,###,##0.00")
        FlexPorModal.TextMatrix(2, 3) = Format(de_informa.rsSel_AlarmMovCli.Fields("tfrete") / de_informa.rsSel_AlarmMovCli.Fields("tvalmerc"), "##0.00%")
        FlexPorModal.TextMatrix(2, 4) = Format(de_informa.rsSel_AlarmMovCli.Fields("tpeso"), "###,###,##0.0")
        FlexPorModal.TextMatrix(2, 5) = Format(de_informa.rsSel_AlarmMovCli.Fields("tvol"), "###,###,##0")
        FlexPorModal.TextMatrix(2, 6) = Format(de_informa.rsSel_AlarmMovCli.Fields("qtd"), "###,###,##0")
        If de_informa.rsSel_AlarmMovCliNFS.State = 1 Then de_informa.rsSel_AlarmMovCliNFS.Close
        'busca qtde de nfs de cada filial
        de_informa.Sel_AlarmMovCliNFS Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                            Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                            "RODOVIARIO", Trim$(gridClientes.Columns(0)) & "%"
        FlexPorModal.TextMatrix(2, 7) = Format(de_informa.rsSel_AlarmMovCliNFS.Fields("qtd"), "###,###,##0")
            
        DoEvents
    Else
        FlexPorModal.TextMatrix(2, 1) = Format(0, "###,###,###,##0.00")
        FlexPorModal.TextMatrix(2, 2) = Format(0, "###,###,##0.00")
        FlexPorModal.TextMatrix(2, 3) = Format(0, "##0.00%")
        FlexPorModal.TextMatrix(2, 4) = Format(0, "###,###,##0.0")
        FlexPorModal.TextMatrix(2, 5) = Format(0, "###,###,##0")
        FlexPorModal.TextMatrix(2, 6) = Format(0, "###,###,##0")
        FlexPorModal.TextMatrix(2, 7) = Format(0, "###,###,##0")
        DoEvents
    End If
    
'MONTANDO O TOTAL  - Cliente
    
    If de_informa.rsSel_AlarmMovCli.State = 1 Then de_informa.rsSel_AlarmMovCli.Close
    de_informa.Sel_AlarmMovCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "%", gridClientes.Columns(0) & "%"
    
    If de_informa.rsSel_AlarmMovCli.RecordCount > 0 Then
        FlexPorModal.TextMatrix(3, 1) = Format(de_informa.rsSel_AlarmMovCli.Fields("tvalmerc"), "###,###,###,##0.00")
        FlexPorModal.TextMatrix(3, 2) = Format(de_informa.rsSel_AlarmMovCli.Fields("tfrete"), "###,###,##0.00")
        FlexPorModal.TextMatrix(3, 3) = Format(de_informa.rsSel_AlarmMovCli.Fields("tfrete") / de_informa.rsSel_AlarmMovCli.Fields("tvalmerc"), "##0.00%")
        FlexPorModal.TextMatrix(3, 4) = Format(de_informa.rsSel_AlarmMovCli.Fields("tpeso"), "###,###,##0.0")
        FlexPorModal.TextMatrix(3, 5) = Format(de_informa.rsSel_AlarmMovCli.Fields("tvol"), "###,###,##0")
        FlexPorModal.TextMatrix(3, 6) = Format(de_informa.rsSel_AlarmMovCli.Fields("qtd"), "###,###,##0")
        If de_informa.rsSel_AlarmMovCliNFS.State = 1 Then de_informa.rsSel_AlarmMovCliNFS.Close
        'busca qtde de nfs de cada filial
        de_informa.Sel_AlarmMovCliNFS Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                            Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                            "%", Trim$(gridClientes.Columns(0)) & "%"
        FlexPorModal.TextMatrix(3, 7) = Format(de_informa.rsSel_AlarmMovCliNFS.Fields("qtd"), "###,###,##0")
            
        DoEvents
    Else
        FlexPorModal.TextMatrix(3, 1) = Format(0, "###,###,###,##0.00")
        FlexPorModal.TextMatrix(3, 2) = Format(0, "###,###,##0.00")
        FlexPorModal.TextMatrix(3, 3) = Format(0, "##0.00%")
        FlexPorModal.TextMatrix(3, 4) = Format(0, "###,###,##0.0")
        FlexPorModal.TextMatrix(3, 5) = Format(0, "###,###,##0")
        FlexPorModal.TextMatrix(3, 6) = Format(0, "###,###,##0")
        FlexPorModal.TextMatrix(3, 7) = Format(0, "###,###,##0")
        DoEvents
    End If
    
    'monta a performance - aéreo / rodo (total BR)
    
    'BR
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "AEREO", gridClientes.Columns(0) & "%", "%"
                                
    FlexPerfAir.TextMatrix(1, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfAir.TextMatrix(1, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfAir.TextMatrix(1, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfAir.TextMatrix(1, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    DoEvents
    
    If de_informa.rsSel_AlarmPrazosCli.Fields("ctcs") > 0 Then
        cmdPerfUFAir.Enabled = True
    Else
        cmdPerfUFAir.Enabled = False
    End If
    
    If de_informa.rsSel_AlarmPrazosCli.State = 1 Then de_informa.rsSel_AlarmPrazosCli.Close
    de_informa.Sel_AlarmPrazosCli Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 1, 4), _
                                Mid$(comboMesAnoCliente.ItemData(comboMesAnoCliente.ListIndex), 5), _
                                "RODOVIARIO", gridClientes.Columns(0) & "%", "%"
                                
    FlexPerfRodo.TextMatrix(1, 1) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("perc"), "##0.00%")
    FlexPerfRodo.TextMatrix(1, 2) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("noprazo"), "##,##0")
    FlexPerfRodo.TextMatrix(1, 3) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("atraso"), "##,##0")
    FlexPerfRodo.TextMatrix(1, 4) = Format(de_informa.rsSel_AlarmPrazosCli.Fields("ctcs"), "##,##0")
    
    If de_informa.rsSel_AlarmPrazosCli.Fields("ctcs") > 0 Then
        cmdPerfUFRodo.Enabled = True
    Else
        cmdPerfUFRodo.Enabled = False
    End If
    
    DoEvents
    
    frmAlarmeUrg.MousePointer = 0
    lblaguarde.Visible = False
    fraClientes.Enabled = True
    fraPorModal.Enabled = True
    fraPerfRodo.Enabled = True
    fraPerfAir.Enabled = True
    cmdSairGerencial.Enabled = True
    comboMesAnoCliente.Enabled = True
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    cmdImprGer2.Enabled = True
    DoEvents
    

End Sub

Private Sub gridPrioridades_Click()
    cmdCopiarPri.Enabled = True
End Sub

Private Sub gridPrioridades_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
        If cmdImprListPri.Enabled = True Then
            lblFilialCtcPri = gridPrioridades.Columns(0)
            lblDataEmiPri = gridPrioridades.Columns(1)
            lblHoraEmiPri = gridPrioridades.Columns(2)
            lblModalPri = gridPrioridades.Columns(3)
            lblRemetPri = gridPrioridades.Columns(5)
            lblCidadeOrigPri = gridPrioridades.Columns(6)
            lblDestPri = gridPrioridades.Columns(7)
            lblCidadeDestPri = gridPrioridades.Columns(8)
            lblUfDestPri = gridPrioridades.Columns(9)
            lblNfsPri = gridPrioridades.Columns(10)
            lblObsEmissPri = gridPrioridades.Columns(11)
            If gridPrioridades.Columns(12) = "N" Then
                lblStatusPri = "SEM POSIÇÃO"
            ElseIf gridPrioridades.Columns(12) = "2" Then
                lblStatusPri = "EM OCORRÊNCIA"
            End If
            cmdCopiarPri.Enabled = True
        End If

End Sub

Private Sub gridUrgentes_Click()
        cmdCopiarUrg.Enabled = True
End Sub

Private Sub gridUrgentes_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If cmdImprListUrg.Enabled = True Then
        lblFilialctcUrg = gridUrgentes.Columns(0)
        lblDataEmiUrg = gridUrgentes.Columns(1)
        lblHoraEmiUrg = gridUrgentes.Columns(2)
        lblModalUrg = gridUrgentes.Columns(3)
        lblRemetUrg = gridUrgentes.Columns(5)
        lblCidadeOrigUrg = gridUrgentes.Columns(6)
        lblDestUrg = gridUrgentes.Columns(7)
        lblCidadeDestUrg = gridUrgentes.Columns(8)
        lblUfDestUrg = gridUrgentes.Columns(9)
        lblNfsUrg = gridUrgentes.Columns(10)
        lblObsEmissUrg = gridUrgentes.Columns(11)
        If gridUrgentes.Columns(12) = "N" Then
            lblStatusUrg = "SEM POSIÇÃO"
        ElseIf gridUrgentes.Columns(12) = "2" Then
            lblStatusUrg = "EM OCORRÊNCIA"
        End If
        cmdCopiarUrg.Enabled = True
    End If
End Sub

Private Sub tm_atualiza_Timer()
    
    tm_atualiza.Interval = 0
    
    If de_informa.rsSel_Urgencias.State = 1 Then de_informa.rsSel_Urgencias.Close
    de_informa.Sel_Urgencias
    
    If de_informa.rsSel_Prioridades.State = 1 Then de_informa.rsSel_Prioridades.Close
    de_informa.Sel_Prioridades datahora("data")
    
    If Mid$(xdireitos, 30, 1) = "0" Then
        SSTab1.TabEnabled(3) = False
    Else
        SSTab1.TabEnabled(3) = True
    End If
    
    If (de_informa.rsSel_Urgencias.RecordCount + de_informa.rsSel_Prioridades.RecordCount) > 0 Then
        
        de_informa.ins_LogUsuario "CONSULTA", xusuario, "TELA DE URGÊNCIAS/PRIORIDADE"
    
        If de_informa.rsSel_Urgencias.RecordCount > 0 Then   'URGENCIAS
            gridUrgentes.DataMember = "sel_urgencias"
            gridUrgentes.Refresh
            lblFilialctcUrg = gridUrgentes.Columns(0)
            lblDataEmiUrg = gridUrgentes.Columns(1)
            lblHoraEmiUrg = gridUrgentes.Columns(2)
            lblModalUrg = gridUrgentes.Columns(3)
            lblRemetUrg = gridUrgentes.Columns(5)
            lblCidadeOrigUrg = gridUrgentes.Columns(6)
            lblDestUrg = gridUrgentes.Columns(7)
            lblCidadeDestUrg = gridUrgentes.Columns(8)
            lblUfDestUrg = gridUrgentes.Columns(9)
            lblNfsUrg = gridUrgentes.Columns(10)
            lblObsEmissUrg = gridUrgentes.Columns(11)
            If gridUrgentes.Columns(12) = "N" Then
                lblStatusUrg = "SEM POSIÇÃO"
            ElseIf gridUrgentes.Columns(12) = "2" Then
                lblStatusUrg = "EM OCORRÊNCIA"
            End If
            fraUrgencias = "CTCs com Urgência (Não tratando Transit-Time): " & Trim$(Str(de_informa.rsSel_Urgencias.RecordCount))
            gridUrgentes.SetFocus
            DoEvents
            Beep
        Else
            MsgBox "Não Há URGÊNCIAS Pendentes mas há PRIORIDADES Pendentes. Clique na Aba PRIORIDADES PENDENTES."
        End If
        
        If de_informa.rsSel_Prioridades.RecordCount > 0 Then   'PRIORIDADES
            gridPrioridades.DataMember = "sel_prioridades"
            gridPrioridades.Refresh
            
            lblFilialCtcPri = gridPrioridades.Columns(0)
            lblDataEmiPri = gridPrioridades.Columns(1)
            lblHoraEmiPri = gridPrioridades.Columns(2)
            lblModalPri = gridPrioridades.Columns(3)
            lblRemetPri = gridPrioridades.Columns(5)
            lblCidadeOrigPri = gridPrioridades.Columns(6)
            lblDestPri = gridPrioridades.Columns(7)
            lblCidadeDestPri = gridPrioridades.Columns(8)
            lblUfDestPri = gridPrioridades.Columns(9)
            lblNfsPri = gridPrioridades.Columns(10)
            lblObsEmissPri = gridPrioridades.Columns(11)
            If gridPrioridades.Columns(12) = "N" Then
                lblStatusPri = "SEM POSIÇÃO"
            ElseIf gridPrioridades.Columns(12) = "2" Then
                lblStatusPri = "EM OCORRÊNCIA"
            End If
            fraPrioridades = "CTCs com Prioridade (Tratando Transit-Time): " & Trim$(Str(de_informa.rsSel_Prioridades.RecordCount))
            DoEvents
            Beep
        End If
    Else
        MsgBox "Não há Urgências / Prioridades Pendentes !"
    End If
    
    cmdImprTelaUrg.Enabled = True
    cmdImprListUrg.Enabled = True
    cmdSairUrg.Enabled = True
    cmdImprTelaPri.Enabled = True
    cmdImprListPri.Enabled = True
    cmdSairPri.Enabled = True
    cmdImprTelaRes.Enabled = True
    cmdSairRes.Enabled = True
    CmdProcessar.Enabled = True
    

    'preencher combos mes/ano - gerencial
    
    Call combomesano(comboMesAnoGeral)
    Call combomesano(comboMesAnoCliente)
    
    comboMesAnoCliente.ListIndex = 0
    comboMesAnoGeral.ListIndex = 0
    
    'preenche clientes que podem ser analisados
    
    If de_informa.rsSel_ClientesAlarmMov.State = 1 Then de_informa.rsSel_ClientesAlarmMov.Close
    de_informa.Sel_ClientesAlarmMov
    
    gridClientes.DataMember = "sel_clientesalarmmov"
    gridClientes.Refresh
    
End Sub
