VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAnEstat2 
   Caption         =   "Análise Estatística 2"
   ClientHeight    =   8325
   ClientLeft      =   210
   ClientTop       =   525
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   12000
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmAnEstat2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmAnEstat2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid7"
      Tab(1).Control(1)=   "MSHFlexGrid8"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame9 
         Caption         =   "Por UF / Região"
         Height          =   6495
         Left            =   5160
         TabIndex        =   20
         Top             =   480
         Width           =   6495
         Begin VB.TextBox Text10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   5760
            Width           =   495
         End
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   5400
            Width           =   495
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   5400
            Width           =   1215
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   5880
            Width           =   495
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   6120
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   5880
            Width           =   1215
         End
         Begin VB.Frame Frame8 
            Height          =   4455
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   3135
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexValmerNONE 
               Height          =   4095
               Left            =   1320
               TabIndex        =   48
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   7223
               _Version        =   393216
               Rows            =   17
               FixedCols       =   0
               ScrollBars      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexUF1 
               Height          =   4095
               Left            =   840
               TabIndex        =   49
               Top             =   240
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   7223
               _Version        =   393216
               Rows            =   17
               FixedCols       =   0
               ScrollBars      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000010&
               X1              =   1200
               X2              =   120
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000010&
               X1              =   720
               X2              =   120
               Y1              =   2160
               Y2              =   2160
            End
            Begin VB.Line Line1 
               BorderColor     =   &H80000010&
               X1              =   120
               X2              =   1200
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label12 
               Caption         =   "Região Norte"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   51
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Região Nordeste"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   50
               Top             =   3000
               Width           =   735
            End
         End
         Begin VB.Frame Frame12 
            Height          =   3975
            Left            =   3240
            TabIndex        =   39
            Top             =   1320
            Width           =   3135
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexExpedSDSU 
               Height          =   2895
               Left            =   1320
               TabIndex        =   40
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   5106
               _Version        =   393216
               Rows            =   12
               FixedCols       =   0
               ScrollBars      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexUF2 
               Height          =   2895
               Left            =   840
               TabIndex        =   41
               Top             =   240
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   5106
               _Version        =   393216
               Rows            =   12
               FixedCols       =   0
               ScrollBars      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexValMerTot 
               Height          =   375
               Left            =   1320
               TabIndex        =   42
               Top             =   3360
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               Rows            =   1
               FixedRows       =   0
               FixedCols       =   0
               ScrollBars      =   0
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Line Line7 
               BorderColor     =   &H80000010&
               X1              =   720
               X2              =   120
               Y1              =   2280
               Y2              =   2280
            End
            Begin VB.Label Label31 
               Caption         =   "Região Sul"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   46
               Top             =   1680
               Width           =   615
            End
            Begin VB.Line Line6 
               BorderColor     =   &H80000010&
               X1              =   1200
               X2              =   120
               Y1              =   3120
               Y2              =   3120
            End
            Begin VB.Line Line5 
               BorderColor     =   &H80000010&
               X1              =   720
               X2              =   120
               Y1              =   1440
               Y2              =   1440
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000010&
               X1              =   120
               X2              =   1200
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Região Centro Oeste"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   120
               TabIndex        =   45
               Top             =   2400
               Width           =   615
            End
            Begin VB.Label Label21 
               Caption         =   "Região Sudeste"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   44
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               Caption         =   "Total Todas as Regiões"
               Height          =   495
               Left            =   240
               TabIndex        =   43
               Top             =   3360
               Width           =   1095
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Expedições"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Notas Fiscais"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Val. NF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Frete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Peso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Mes 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   840
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Mes 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Mes 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   6120
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4560
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   5760
            Width           =   1215
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Região Norte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   5880
            Width           =   945
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Região Nordeste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   6120
            Width           =   1200
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Região Sudeste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   54
            Top             =   5400
            Width           =   1140
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Região Sul"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   53
            Top             =   5760
            Width           =   780
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Região C.Oeste"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   52
            Top             =   6120
            Width           =   1125
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "% Mes 1 x Mes 3"
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   6240
         Width           =   4935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "% Mes 2 x Mes 3"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   5520
         Width           =   4935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mes 1"
         Height          =   1455
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   4935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   1095
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1931
            _Version        =   393216
            Rows            =   4
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mes 3"
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   4935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   1095
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1931
            _Version        =   393216
            Rows            =   4
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "% Mes 1 x Mes 2"
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   4935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   661
            _Version        =   393216
            Rows            =   1
            FixedRows       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mes 2"
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   4935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1095
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1931
            _Version        =   393216
            Rows            =   4
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
         Height          =   6735
         Left            =   -74760
         TabIndex        =   57
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   11880
         _Version        =   393216
         Rows            =   28
         FixedCols       =   0
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
         Height          =   6735
         Left            =   -74160
         TabIndex        =   58
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   11880
         _Version        =   393216
         Rows            =   28
         FixedCols       =   0
         ScrollBars      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Dados Selecionados"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5775
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cliente / Remetente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   3465
      End
   End
   Begin VB.CommandButton cmdNova 
      Caption         =   "Nova ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdImprTela 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      Picture         =   "frmAnEstat2.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdGeraXls 
      Caption         =   "Gerar no EXCEL ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmAnEstat2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
