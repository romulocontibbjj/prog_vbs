VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVLSac 
   Caption         =   "Consultas / Informação - Videolar"
   ClientHeight    =   7935
   ClientLeft      =   750
   ClientTop       =   1875
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame15 
      Caption         =   " S T A T U S "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7920
      TabIndex        =   28
      Top             =   0
      Width           =   3975
      Begin VB.Label lblEntregueSN 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   75
         TabIndex        =   29
         Top             =   360
         Width           =   3810
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "CTC/CTR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6120
      TabIndex        =   70
      Top             =   0
      Width           =   1815
      Begin VB.TextBox txtFilial 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   72
         Top             =   240
         Width           =   330
      End
      Begin VB.TextBox txtCtc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         MaxLength       =   8
         TabIndex        =   71
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "N.Fiscal:"
         Height          =   195
         Left            =   120
         TabIndex        =   81
         Top             =   525
         Width           =   615
      End
      Begin VB.Label lblNumero 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   740
         TabIndex        =   80
         Top             =   525
         Width           =   970
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Origem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   5775
      Begin VB.Label lblUF_Orig 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5235
         TabIndex        =   27
         Top             =   930
         Width           =   420
      End
      Begin VB.Label lblCidade_orig 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   930
         Width           =   4050
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Localidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblRemet_CGC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label lblEndRem 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   675
         Width           =   4575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblRemet_Nome 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   420
         Width           =   4575
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   6120
      TabIndex        =   10
      Top             =   840
      Width           =   5775
      Begin VB.Label lblUf_Dest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5235
         TabIndex        =   18
         Top             =   930
         Width           =   420
      End
      Begin VB.Label lblCidade_Dest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   930
         Width           =   4050
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Destinatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Localidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   930
         Width           =   825
      End
      Begin VB.Label lblDest_CGC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label lblEndDest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   675
         Width           =   4575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   675
         Width           =   690
      End
      Begin VB.Label lblDest_Nome 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   420
         Width           =   4575
      End
   End
   Begin VB.Frame fraProcura 
      Caption         =   "Série e NF/Pedido"
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
      TabIndex        =   8
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtNumNf 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox txtNumPedido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   4
         Top             =   480
         Width           =   1365
      End
      Begin VB.OptionButton optPorPedido 
         Caption         =   "Por Pedido.."
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton optPorNF 
         Caption         =   "Por NF........"
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "1"
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.CommandButton cmbProcurar 
      Caption         =   "Procurar"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdImprTela 
      Height          =   495
      Left            =   4440
      Picture         =   "frmVLSac.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmbSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "frmVLSac.frx":0772
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label17"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label19"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTotitens"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblTotitens2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label21"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "gridManifesto"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "gridVideolarEspeciaisItem"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "gridNFsEspeciais"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "gridCheckEspeciais"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "gridEmbarqueIntec"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "gridVideolarEspeciais"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "gridClienteEspeciais"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Detalhe Dados de Embarque (CTC/CTR)"
      TabPicture(1)   =   "frmVLSac.frx":078E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame24"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame22"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame19"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame1 
         Caption         =   "Características da Carga"
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
         Left            =   -66240
         TabIndex        =   52
         Top             =   1780
         Width           =   2895
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   2760
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   2760
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Peso Tax:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label lblValmerc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   960
            TabIndex        =   59
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   405
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Valor NFs:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Volumes:"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label lblPesotax 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   960
            TabIndex        =   55
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblPeso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   960
            TabIndex        =   54
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblVolumes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   960
            TabIndex        =   53
            Top             =   2040
            Width           =   735
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Dados da Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74880
         TabIndex        =   45
         Top             =   4320
         Width           =   2900
         Begin VB.Label lblRecebPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   51
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label lblHsBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   49
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblDtBaixaPre 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   50
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Receb:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   525
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   390
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Hora:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   540
            Width           =   390
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Observação de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -71880
         TabIndex        =   43
         Top             =   4320
         Width           =   8535
         Begin VB.Label lblObsEntr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   885
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   8280
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Ocorrências"
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
         Left            =   -74880
         TabIndex        =   40
         Top             =   1780
         Width           =   8535
         Begin MSDataGridLib.DataGrid GridConsOcorr 
            Bindings        =   "frmVLSac.frx":07AA
            Height          =   1215
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2143
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            ForeColor       =   0
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
            ColumnCount     =   7
            BeginProperty Column00 
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
            BeginProperty Column01 
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
            BeginProperty Column02 
               DataField       =   "cod_ocorr"
               Caption         =   "cd"
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
               Caption         =   "ocorrência / descrição"
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
               Caption         =   "usuário"
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
               Caption         =   "data inclusão"
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
            SplitCount      =   1
            BeginProperty Split0 
               AllowFocus      =   0   'False
               AllowRowSizing  =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   480,189
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   269,858
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   3509,858
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1005,165
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1560,189
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   14,74
               EndProperty
            EndProperty
         End
         Begin VB.Label lblObs_Ocorr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   840
            Left            =   120
            TabIndex        =   42
            Top             =   1560
            Width           =   8295
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados do Documento de Embarque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Width           =   11535
         Begin VB.Label lblPrevEntrega 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3720
            TabIndex        =   84
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Prev.Entrega:"
            Height          =   195
            Left            =   3720
            TabIndex        =   83
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Prioridade:"
            Height          =   195
            Left            =   2280
            TabIndex        =   82
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "OBS:"
            Height          =   195
            Left            =   120
            TabIndex        =   74
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblObs_Emissao 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   495
            Left            =   600
            TabIndex        =   73
            Top             =   840
            Width           =   10815
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Modal:"
            Height          =   195
            Left            =   5280
            TabIndex        =   39
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblModal 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5280
            TabIndex        =   38
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblTranspsubRedesp 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8640
            TabIndex        =   37
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Redespacho:"
            Height          =   195
            Left            =   7080
            TabIndex        =   36
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblTranspSub 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   7080
            TabIndex        =   35
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblData 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label lblHora 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emissão (Data/Hora):"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1530
         End
         Begin VB.Label lblPrioridade 
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
            Left            =   2280
            TabIndex        =   31
            Top             =   480
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid gridClienteEspeciais 
         Bindings        =   "frmVLSac.frx":07C3
         Height          =   1575
         Left            =   2160
         TabIndex        =   64
         Top             =   600
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2778
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
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
         DataMember      =   "Sel_NFBasecli"
         ColumnCount     =   20
         BeginProperty Column00 
            DataField       =   "clientecgc"
            Caption         =   "clientecgc"
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
            DataField       =   "clientenome"
            Caption         =   "clientenome"
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
            DataField       =   "ordvenda"
            Caption         =   "Ord.Venda"
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
            DataField       =   "item"
            Caption         =   "Ítem"
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
            DataField       =   "pedido"
            Caption         =   "Pedido"
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
            DataField       =   "codclinf"
            Caption         =   "Cod.Cliente"
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
            DataField       =   "clientenf"
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
         BeginProperty Column07 
            DataField       =   "cidadenf"
            Caption         =   "Cidade Dest."
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
            DataField       =   "ufnf"
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
         BeginProperty Column09 
            DataField       =   "codmaterial"
            Caption         =   "Cod.Material"
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
            DataField       =   "material"
            Caption         =   "Material"
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
            DataField       =   "numnf"
            Caption         =   "numnf"
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
            DataField       =   "serie"
            Caption         =   "serie"
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
            DataField       =   "datanf"
            Caption         =   "Data da NF"
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
            DataField       =   "coletadata"
            Caption         =   "coletadata"
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
            DataField       =   "entr_solic"
            Caption         =   "Solic.Entr."
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
            DataField       =   "acaofox"
            Caption         =   "acaofox"
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
            DataField       =   "pacote"
            Caption         =   "Pacote"
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
            DataField       =   "qtdeitem"
            Caption         =   "Qtde."
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
            DataField       =   "dataarq"
            Caption         =   "Data Arquivo"
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
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   420,095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2745,071
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1980,284
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   315,213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   2550,047
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column12 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column14 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column16 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   2355,024
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1335,118
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridVideolarEspeciais 
         Bindings        =   "frmVLSac.frx":07DC
         Height          =   1455
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2566
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
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
         DataMember      =   "Sel_NFVideolar"
         ColumnCount     =   29
         BeginProperty Column00 
            DataField       =   "id_local"
            Caption         =   "Estudio"
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
            DataField       =   "remet_cgc"
            Caption         =   "remet_cgc"
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
         BeginProperty Column03 
            DataField       =   "dest_cgc"
            Caption         =   "dest_cgc"
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
         BeginProperty Column05 
            DataField       =   "dest_ie"
            Caption         =   "dest_ie"
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
            DataField       =   "dest_end"
            Caption         =   "dest_end"
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
            DataField       =   "dest_bairro"
            Caption         =   "dest_bairro"
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
            DataField       =   "dest_cidade"
            Caption         =   "Cidade"
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
            DataField       =   "dest_uf"
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
         BeginProperty Column10 
            DataField       =   "dest_cep"
            Caption         =   "dest_cep"
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
            DataField       =   "tipocarga"
            Caption         =   "tipocarga"
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
            DataField       =   "tipofrete"
            Caption         =   "tipofrete"
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
            DataField       =   "numnf"
            Caption         =   "numnf"
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
            DataField       =   "numnfnum"
            Caption         =   "numnfnum"
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
            DataField       =   "serie"
            Caption         =   "serie"
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
            DataField       =   "emissaonf"
            Caption         =   "Emissão NF"
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
            DataField       =   "natureza"
            Caption         =   "natureza"
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
            DataField       =   "especie"
            Caption         =   "especie"
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
            DataField       =   "volumes"
            Caption         =   "Vols"
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
            DataField       =   "qtdeitem"
            Caption         =   "qtdeitem"
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
            DataField       =   "valmerc"
            Caption         =   "Valor Merc."
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
            DataField       =   "peso"
            Caption         =   "Peso"
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
            DataField       =   "pesocub"
            Caption         =   "pesocub"
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
            DataField       =   "datainterface"
            Caption         =   "datainterface"
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
            DataField       =   "dataimp"
            Caption         =   "dataimp"
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
            DataField       =   "emitido_auto"
            Caption         =   "emitido_auto"
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
            DataField       =   "emit_data"
            Caption         =   "emit_data"
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
            DataField       =   "id_local"
            Caption         =   "id_local"
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
               ColumnAllowSizing=   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2145,26
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
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
               ColumnWidth     =   1950,236
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   329,953
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column12 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column13 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column14 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column15 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column17 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column18 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   599,811
            EndProperty
            BeginProperty Column20 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   705,26
            EndProperty
            BeginProperty Column23 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column24 
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column25 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column26 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column27 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column28 
               ColumnWidth     =   30,047
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridEmbarqueIntec 
         Bindings        =   "frmVLSac.frx":07F5
         Height          =   1095
         Left            =   3120
         TabIndex        =   66
         Top             =   4440
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1931
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
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
         DataMember      =   "Sel_VLNFSerie_Sac"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "filialctc"
            Caption         =   "Ctc / Ctr"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "transp_sub"
            Caption         =   "Redespacho"
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
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   480,189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1785,26
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridCheckEspeciais 
         Bindings        =   "frmVLSac.frx":080E
         Height          =   1095
         Left            =   120
         TabIndex        =   68
         Top             =   4440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1931
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
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
         DataMember      =   "Sel_NFCheckReceb"
         ColumnCount     =   8
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "embarcador"
            Caption         =   "embarcador"
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
            DataField       =   "serie"
            Caption         =   "serie"
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
            DataField       =   "numnf"
            Caption         =   "numnf"
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
            DataField       =   "numnfnum"
            Caption         =   "numnfnum"
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
            Caption         =   "Data / Hora"
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
            DataField       =   "usuario"
            Caption         =   "usuario"
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
            Caption         =   "Obs"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   615,118
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2340,284
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridNFsEspeciais 
         Bindings        =   "frmVLSac.frx":0827
         Height          =   1695
         Left            =   120
         TabIndex        =   69
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
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
         DataMember      =   "Sel_NFsdoCTC"
         ColumnCount     =   21
         BeginProperty Column00 
            DataField       =   "idcodigo"
            Caption         =   "idcodigo"
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
         BeginProperty Column02 
            DataField       =   "numnfnum"
            Caption         =   "numnfnum"
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
            DataField       =   "numnf"
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
         BeginProperty Column04 
            DataField       =   "serie"
            Caption         =   "S"
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
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "emissao_nf"
            Caption         =   "emissao_nf"
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
            DataField       =   "numpedido"
            Caption         =   "numpedido"
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
            DataField       =   "dtpedido"
            Caption         =   "dtpedido"
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
            DataField       =   "valornf"
            Caption         =   "valornf"
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
            DataField       =   "pesonf"
            Caption         =   "pesonf"
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
            DataField       =   "volumesnf"
            Caption         =   "volumesnf"
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
            DataField       =   "data_interface"
            Caption         =   "data_interface"
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
            DataField       =   "hora_interface"
            Caption         =   "hora_interface"
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
            DataField       =   "at_cliente"
            Caption         =   "at_cliente"
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
            DataField       =   "canhotonf"
            Caption         =   "canhotonf"
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
            DataField       =   "canhotonfprot"
            Caption         =   "canhotonfprot"
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
            DataField       =   "canhotonfdata"
            Caption         =   "canhotonfdata"
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
            DataField       =   "tem_ocorrnf"
            Caption         =   "tem_ocorrnf"
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
            DataField       =   "ordem"
            Caption         =   "ordem"
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
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   269,858
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1440
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
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column13 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column15 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   705,26
            EndProperty
            BeginProperty Column16 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column17 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column18 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column19 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column20 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridVideolarEspeciaisItem 
         Bindings        =   "frmVLSac.frx":0840
         Height          =   1215
         Left            =   7320
         TabIndex        =   75
         Top             =   2640
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2143
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
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
         DataMember      =   "Sel_NFVideolarItens"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "id_notfis"
            Caption         =   "id_notfis"
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
            DataField       =   "remet_cgc"
            Caption         =   "remet_cgc"
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
         BeginProperty Column03 
            DataField       =   "numnf"
            Caption         =   "numnf"
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
            DataField       =   "numnfnum"
            Caption         =   "numnfnum"
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
            DataField       =   "serie"
            Caption         =   "serie"
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
            DataField       =   "posicao"
            Caption         =   "posicao"
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
            DataField       =   "codigoitem"
            Caption         =   "codigoitem"
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
            DataField       =   "descricaoitem"
            Caption         =   "Descrição"
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
            DataField       =   "qtdeitem"
            Caption         =   "Qtde"
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
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   615,118
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
               ColumnWidth     =   3240
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   540,284
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridManifesto 
         Bindings        =   "frmVLSac.frx":0859
         Height          =   1095
         Left            =   8760
         TabIndex        =   77
         Top             =   4440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1931
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
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
         DataMember      =   "Sel_ManifestoPorCTC"
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "idcodigo"
            Caption         =   "idcodigo"
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
         BeginProperty Column02 
            DataField       =   "filialmanifesto"
            Caption         =   "Manifesto"
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
            DataField       =   "filial"
            Caption         =   "filial"
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
            DataField       =   "manifesto"
            Caption         =   "manifesto"
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
            DataField       =   "embarcador"
            Caption         =   "embarcador"
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
            DataField       =   "dtemissao"
            Caption         =   "Data Manif."
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
         BeginProperty Column08 
            DataField       =   "dtsaida"
            Caption         =   "Data Saída"
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
            DataField       =   "hssaida"
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
         BeginProperty Column13 
            DataField       =   "at_manif_cif"
            Caption         =   "at_manif_cif"
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
            DataField       =   "at_manif_cif_data"
            Caption         =   "at_manif_cif_data"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   -1  'True
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1665,071
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Total Ítens:"
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
         Left            =   9720
         TabIndex        =   88
         Top             =   3880
         Width           =   990
      End
      Begin VB.Label lblTotitens2 
         Alignment       =   1  'Right Justify
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
         Left            =   10800
         TabIndex        =   87
         Top             =   3885
         Width           =   735
      End
      Begin VB.Label lblTotitens 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         TabIndex        =   86
         Top             =   2220
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Total de Ítens desta NF (Base Estúdio):"
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
         Left            =   3000
         TabIndex        =   85
         Top             =   2220
         Width           =   3375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Manifesto de Saída"
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
         Left            =   8760
         TabIndex        =   79
         Top             =   4200
         Width           =   1680
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dados da VideoLar (EDI) - Ítens"
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
         Left            =   7320
         TabIndex        =   76
         Top             =   2400
         Width           =   2730
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "NFs do CTR"
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
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Informação de Embarque (CTC/CTR)"
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
         Left            =   3120
         TabIndex        =   63
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Entrada/Receb. (LUFT)"
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
         Left            =   120
         TabIndex        =   62
         Top             =   4200
         Width           =   2025
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dados da VideoLar (EDI) - NF"
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
         Left            =   120
         TabIndex        =   61
         Top             =   2400
         Width           =   2550
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Dados Arquivo Cliente Estúdio"
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
         Left            =   2160
         TabIndex        =   60
         Top             =   360
         Width           =   2595
      End
   End
End
Attribute VB_Name = "frmVLSac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public xtemocorr As String
Private Sub cmbProcurar_Click()
Dim xserie As String, xnumnf As String, xnumpedido As String, xtot As Integer, xtot2 As Integer

    If (optPorNF.Value = True And (txtNumNf.Text = "" Or txtSerie.Text = "")) Or _
       (optPorPedido.Value = True And (txtNumPedido.Text = "" Or txtSerie.Text = "")) Then
        If optPorNF.Value = True Then
            MsgBox "Número de Nota Fiscal / Série Inválida !", vbCritical, "Erro"
            txtNumNf.SetFocus
            Exit Sub
        End If
        If optPorPedido.Value = True Then
            MsgBox "Número de Pedido / Série Inválida !", vbCritical, "Erro"
            txtNumPedido.SetFocus
            Exit Sub
        End If
    Else
        cmbProcurar.Enabled = False
        cmdImprTela.Enabled = False
        cmbSair.Enabled = False
        SSTab1.Enabled = False
        cmbProcurar.Caption = "Aguarde..."
        Me.MousePointer = 11
        DoEvents
        
        gridNFsEspeciais.DataMember = ""
        gridNFsEspeciais.Refresh
        gridClienteEspeciais.DataMember = ""
        gridClienteEspeciais.Refresh
        gridVideolarEspeciais.DataMember = ""
        gridVideolarEspeciais.Refresh
        gridVideolarEspeciaisItem.DataMember = ""
        gridVideolarEspeciaisItem.Refresh
        gridCheckEspeciais.DataMember = ""
        gridCheckEspeciais.Refresh
        gridEmbarqueIntec.DataMember = ""
        gridEmbarqueIntec.Refresh
        GridConsOcorr.DataMember = ""
        GridConsOcorr.Refresh
        gridManifesto.DataMember = ""
        gridManifesto.Refresh
    
        xserie = txtSerie
        xnumnf = txtNumNf
        xnumpedido = txtNumPedido
            Call limpatela(Me)
        txtSerie = xserie
        txtNumNf = xnumnf
        txtNumPedido = xnumpedido
        DoEvents
        
        lblEntregueSN.Caption = ""
        
        If optPorPedido.Value = True Then
            If de_informa.rsSel_PedidoBasecli.State = 1 Then de_informa.rsSel_PedidoBasecli.Close
            de_informa.Sel_PedidoBasecli Trim$(txtNumPedido), txtSerie
            
            If de_informa.rsSel_PedidoBasecli.RecordCount = 0 Then
                MsgBox "Número de Pedido / Série Não Encontrada !", vbExclamation, "Erro"
                cmbProcurar.Enabled = True
                cmdImprTela.Enabled = True
                cmbSair.Enabled = True
                SSTab1.Enabled = True
                cmbProcurar.Caption = "Procurar"
                Me.MousePointer = 0
                DoEvents
                txtNumPedido.SetFocus
                Exit Sub
            Else
                lblNumero = de_informa.rsSel_PedidoBasecli.Fields("numnf")
            End If
            
        Else
            lblNumero = Trim$(txtNumNf)
        End If
        
        If de_informa.rsSel_NFBasecli.State = 1 Then de_informa.rsSel_NFBasecli.Close
        de_informa.Sel_NFBasecli "04229761000413", Trim$(lblNumero), Trim$(txtSerie)
        gridClienteEspeciais.DataMember = "Sel_NFBasecli"
        gridClienteEspeciais.Refresh
        
        If de_informa.rsSel_NFBasecli.RecordCount > 0 Then
            xtot = 0
            Do Until de_informa.rsSel_NFBasecli.EOF
                xtot = xtot + de_informa.rsSel_NFBasecli.Fields("qtdeitem")
                de_informa.rsSel_NFBasecli.MoveNext
            Loop
            de_informa.rsSel_NFBasecli.MoveFirst
            lblTotitens = xtot
        Else
            lblTotitens = ""
            
            If xusuario = "FOXFILM" Then
                MsgBox "Nota Fiscal / Série Não Encontrada !", vbExclamation, "Erro"
                cmbProcurar.Enabled = True
                cmdImprTela.Enabled = True
                cmbSair.Enabled = True
                SSTab1.Enabled = True
                cmbProcurar.Caption = "Procurar"
                Me.MousePointer = 0
                DoEvents
                If optPorPedido.Value = True Then
                    txtNumPedido.SetFocus
                Else
                    txtNumNf.SetFocus
                End If
                Exit Sub
            End If
        End If
        
        If de_informa.rsSel_VLNFSerie_Sac.State = 1 Then de_informa.rsSel_VLNFSerie_Sac.Close
        de_informa.Sel_VLNFSerie_Sac CDbl(lblNumero), Trim$(Str(Val(txtSerie)))       'Procura a NF na Tabela
        
        If de_informa.rsSel_VLNFSerie_Sac.RecordCount > 0 Then
            gridEmbarqueIntec.Enabled = True
        Else
            gridEmbarqueIntec.Enabled = False
        End If
        
        gridEmbarqueIntec.DataMember = "Sel_VLNFSerie_Sac"
        gridEmbarqueIntec.Refresh
        
        DoEvents
        
        If de_informa.rsSel_NFVideolar.State = 1 Then de_informa.rsSel_NFVideolar.Close
        de_informa.Sel_NFVideolar "04229761000413", CDbl(Trim$(lblNumero)), Trim$(txtSerie)
        gridVideolarEspeciais.DataMember = "Sel_NFVideolar"
        gridVideolarEspeciais.Refresh
        
        DoEvents
        
        If de_informa.rsSel_NFVideolarItens.State = 1 Then de_informa.rsSel_NFVideolarItens.Close
        de_informa.Sel_NFVideolarItens "04229761000413", CDbl(Trim$(lblNumero)), Trim$(txtSerie)
        gridVideolarEspeciaisItem.DataMember = "Sel_NFVideolarItens"
        gridVideolarEspeciaisItem.Refresh
        
        If de_informa.rsSel_NFVideolarItens.RecordCount > 0 Then
            xtot2 = 0
            Do Until de_informa.rsSel_NFVideolarItens.EOF
                xtot2 = xtot2 + de_informa.rsSel_NFVideolarItens.Fields("qtdeitem")
                de_informa.rsSel_NFVideolarItens.MoveNext
            Loop
            de_informa.rsSel_NFVideolarItens.MoveFirst
            lblTotitens2 = xtot2
        Else
            lblTotitens2 = ""
        End If
        
        DoEvents
        
        If de_informa.rsSel_NFCheckReceb.State = 1 Then de_informa.rsSel_NFCheckReceb.Close
        de_informa.Sel_NFCheckReceb "04229761000413", CDbl(Trim$(lblNumero)), Trim$(txtSerie)
        gridCheckEspeciais.DataMember = "Sel_NFCheckReceb"
        gridCheckEspeciais.Refresh
        
        DoEvents
    
        If de_informa.rsSel_VLNFSerie_Sac.RecordCount = 0 Then
            If de_informa.rsSel_NFCheckReceb.RecordCount = 0 And _
               de_informa.rsSel_NFBasecli.RecordCount = 0 And _
               de_informa.rsSel_NFVideolar.RecordCount = 0 Then
                MsgBox "Nota Fiscal / Série Não Encontrada !", vbExclamation, "Erro"
            End If
            cmbProcurar.Enabled = True
            cmdImprTela.Enabled = True
            cmbSair.Enabled = True
            SSTab1.Enabled = True
            cmbProcurar.Caption = "Procurar"
            Me.MousePointer = 0
            DoEvents
            If optPorPedido.Value = True Then
                txtNumPedido.SetFocus
            Else
                txtNumNf.SetFocus
            End If
            Exit Sub
        Else
            If de_informa.rsSel_VLNFSerie_Sac.RecordCount > 0 Then
                TxtFilial.Text = Mid(de_informa.rsSel_VLNFSerie_Sac.Fields("filialctc"), 1, 2)
                txtCtc.Text = Mid(de_informa.rsSel_VLNFSerie_Sac.Fields("filialctc"), 3, 8) 'Busca a Filial e o CTC com base na NF
            End If
            If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
            de_informa.Sel_Ctc_SAC transctc(TxtFilial, txtCtc)
        End If
        
        DoEvents
        
    End If
    
'registra as variáveis da tela com os dados buscados no recorset
    lblData = de_informa.rsSel_Ctc_SAC.Fields("data")
    lblHora = de_informa.rsSel_Ctc_SAC.Fields("hora")
    lblRemet_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), "@@.@@@.@@@/@@@@-@@")
    lblRemet_Nome = de_informa.rsSel_Ctc_SAC.Fields("remet_nome")
    lblEndRem = de_informa.rsSel_Ctc_SAC.Fields("remet_end")
    lblCidade_orig = de_informa.rsSel_Ctc_SAC.Fields("remet_cidade")
    lblUF_Orig = de_informa.rsSel_Ctc_SAC.Fields("remet_uf")
    lblDest_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("dest_cgc"), "@@.@@@.@@@/@@@@-@@")
    lblDest_Nome = de_informa.rsSel_Ctc_SAC.Fields("dest_nome")
    lblEndDest = de_informa.rsSel_Ctc_SAC.Fields("dest_end")
    lblCidade_Dest = de_informa.rsSel_Ctc_SAC.Fields("dest_cidade")
    lblUf_Dest = de_informa.rsSel_Ctc_SAC.Fields("dest_uf")
    lblValmerc = Format(de_informa.rsSel_Ctc_SAC.Fields("valmerc"), "##,###,##0.00")
    lblPeso = Format(de_informa.rsSel_Ctc_SAC.Fields("peso"), "##,##0.0")
    lblVolumes = Format(de_informa.rsSel_Ctc_SAC.Fields("volumes"), "##,##0")
    lblPrevEntrega = de_informa.rsSel_Ctc_SAC.Fields("prev_entrega")
    If de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "URGÊNCIA" Or _
        de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "PRIORIDADE" Then
        LblPrioridade.ForeColor = &HC0&
    Else
        LblPrioridade.ForeColor = &H80000012
    End If
    LblPrioridade = de_informa.rsSel_Ctc_SAC.Fields("prioridade")
    
    lblModal = de_informa.rsSel_Ctc_SAC.Fields("modal")
    lblObs_Emissao = de_informa.rsSel_Ctc_SAC.Fields("obs_emissao")
    xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
    lblTranspSub = de_informa.rsSel_Ctc_SAC.Fields("transp_sub")
    lblTranspsubRedesp = de_informa.rsSel_Ctc_SAC.Fields("redesp_nome")
    
    
    DoEvents

'ATUALIZA COM DADOS DE ENTREGA E OCORRÊNCIA

    'consulta que traz os campos que são dados de ocorrência
    If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
    de_informa.Sel_ConsOcorr2 transctc(TxtFilial.Text, txtCtc.Text), "01"
    Set GridConsOcorr.DataSource = de_informa
    GridConsOcorr.DataMember = "Sel_ConsOcorr2"
    GridConsOcorr.Refresh
    If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
        If de_informa.rsSel_ConsOcorr2.Fields("obs_ocorr") <> "" Then
            lblObs_Ocorr = de_informa.rsSel_ConsOcorr2.Fields("obs_ocorr")
        Else
            lblObs_Ocorr = ""
        End If
    Else
        lblObs_Ocorr = ""
    End If
    
    DoEvents
    
    'consulta que traz os campos = 01 que é dado de entrega (ENTREGA REALIZADA)
    If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
    de_informa.Sel_ConsOcorr transctc(TxtFilial.Text, txtCtc.Text), "01"
    If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
    'atualiza os campos referente a dados de entrega
        lblDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("data")
        lblHsBaixaPre = de_informa.rsSel_ConsOcorr.Fields("hora")
        lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
        If IsNull(de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")) = False Then
            lblObsEntr = de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")
        Else
            lblObsEntr = ""
        End If
    End If
    
    DoEvents

'dados de status do ctc
    
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
        If de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") >= datahora("data") Then
            lblEntregueSN.ForeColor = &HC00000             'LABEL NA COR AZUL
            lblEntregueSN.Caption = "EM TRÂNSITO"
            lblEntregueSN.ToolTipText = "EM TRÂNSITO = Até a Previsão de Entrega"
        Else
            lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
            lblEntregueSN.Caption = "SEM POSIÇÃO"
            lblEntregueSN.ToolTipText = "SEM POSIÇÃO = Após a Previsão de Entrega"
        End If
    ElseIf xtemocorr = "C" Then
        lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
        lblEntregueSN.Caption = "CTC CANCELADO"
        lblEntregueSN.ToolTipText = "Cancelado em:" & de_informa.rsSel_Ctc_SAC.Fields("canc_data") & _
                                    "  Usuário:" & de_informa.rsSel_Ctc_SAC.Fields("canc_usu") & _
                                    "  Motivo:" & de_informa.rsSel_Ctc_SAC.Fields("canc_obs")
    End If

    DoEvents

'Notas Fiscais do CTC
    
    If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
    de_informa.Sel_NFsdoCTC transctc(TxtFilial, txtCtc)
    
    gridNFsEspeciais.DataMember = "sel_nfsdoctc"
    gridNFsEspeciais.Refresh
    
'ATUALIZA GRID DE MANIFESTO
    If de_informa.rsSel_ManifestoPorCTC.State = 1 Then de_informa.rsSel_ManifestoPorCTC.Close
    de_informa.Sel_ManifestoPorCTC transctc(TxtFilial, txtCtc)
    gridManifesto.DataMember = "sel_manifestoporctc"
    gridManifesto.Refresh
    
    If optPorPedido.Value = True Then
        txtNumPedido.SetFocus
    Else
        txtNumNf.SetFocus
    End If
    
    DoEvents
    
'ÚLTIMAS CONSULTAS DO CTC
    de_informa.ins_ultconssac transctc(TxtFilial, txtCtc), xusuario, datahora("datahora")
        
'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "CONSULTA", xusuario, "INFORMAÇÃO SAC - CONSULTA CTC: " & transctc(TxtFilial, txtCtc)
    
    cmbProcurar.Enabled = True
    cmdImprTela.Enabled = True
    cmbSair.Enabled = True
    SSTab1.Enabled = True
    cmbProcurar.Caption = "Procurar"
    
    Me.MousePointer = 0
    DoEvents
        
End Sub
Private Sub cmbSair_Click()
    Unload Me
End Sub
Private Sub cmdImprTela_Click()
    Printer.KillDoc
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub
Private Sub Form_Load()
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.StatusBar1.Visible = False
    gridNFsEspeciais.DataMember = ""
    gridNFsEspeciais.Refresh
    gridClienteEspeciais.DataMember = ""
    gridClienteEspeciais.Refresh
    gridVideolarEspeciais.DataMember = ""
    gridVideolarEspeciais.Refresh
    gridVideolarEspeciaisItem.DataMember = ""
    gridVideolarEspeciaisItem.Refresh
    gridCheckEspeciais.DataMember = ""
    gridCheckEspeciais.Refresh
    gridEmbarqueIntec.DataMember = ""
    gridEmbarqueIntec.Refresh
    GridConsOcorr.DataMember = ""
    GridConsOcorr.Refresh
    gridManifesto.DataMember = ""
    gridManifesto.Refresh
    cmbProcurar.Enabled = True
    cmbSair.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmSac = Nothing
    mdiInforma.Toolbar1.Visible = True
    mdiInforma.StatusBar1.Visible = True
    GridConsOcorr.DataMember = ""
    gridManifesto.DataMember = ""
    
    If de_informa.rsSel_NFBasecli.State = 1 Then de_informa.rsSel_NFBasecli.Close
    gridClienteEspeciais.DataMember = "Sel_NFBasecli"
    gridClienteEspeciais.Refresh
    
    If de_informa.rsSel_NFVideolar.State = 1 Then de_informa.rsSel_NFVideolar.Close
    gridVideolarEspeciais.DataMember = "Sel_NFVideolar"
    gridVideolarEspeciais.Refresh
    
    If de_informa.rsSel_NFVideolarItens.State = 1 Then de_informa.rsSel_NFVideolarItens.Close
    gridVideolarEspeciaisItem.DataMember = "Sel_NFVideolarItens"
    gridVideolarEspeciaisItem.Refresh
    
    If de_informa.rsSel_CheckReceb.State = 1 Then de_informa.rsSel_CheckReceb.Close
    gridCheckEspeciais.DataMember = "Sel_CheckReceb"
    gridCheckEspeciais.Refresh
    
End Sub

Private Sub GridConsOcorr_Click()
    If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
    'atualiza o campo de obs de ocorrência quando clicado no grid
        lblObs_Ocorr.Caption = GridConsOcorr.Columns(6)
    End If
End Sub
Private Sub gridEmbarqueIntec_Click()
    'registra as variáveis da tela com os dados buscados no recorset
    
    If gridEmbarqueIntec.Enabled = True Then
    
        If de_informa.rsSel_Ctc_SAC.State = 1 Then de_informa.rsSel_Ctc_SAC.Close
        de_informa.Sel_Ctc_SAC gridEmbarqueIntec.Columns(0)
        
        TxtFilial = Mid$(gridEmbarqueIntec.Columns(0), 1, 2)
        txtCtc = Mid$(gridEmbarqueIntec.Columns(0), 3)
        
        lblData = de_informa.rsSel_Ctc_SAC.Fields("data")
        lblHora = de_informa.rsSel_Ctc_SAC.Fields("hora")
        lblRemet_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("remet_cgc"), "@@.@@@.@@@/@@@@-@@")
        lblRemet_Nome = de_informa.rsSel_Ctc_SAC.Fields("remet_nome")
        lblEndRem = de_informa.rsSel_Ctc_SAC.Fields("remet_end")
        lblCidade_orig = de_informa.rsSel_Ctc_SAC.Fields("remet_cidade")
        lblUF_Orig = de_informa.rsSel_Ctc_SAC.Fields("remet_uf")
        lblDest_CGC = Format(de_informa.rsSel_Ctc_SAC.Fields("dest_cgc"), "@@.@@@.@@@/@@@@-@@")
        lblDest_Nome = de_informa.rsSel_Ctc_SAC.Fields("dest_nome")
        lblEndDest = de_informa.rsSel_Ctc_SAC.Fields("dest_end")
        lblCidade_Dest = de_informa.rsSel_Ctc_SAC.Fields("dest_cidade")
        lblUf_Dest = de_informa.rsSel_Ctc_SAC.Fields("dest_uf")
        lblValmerc = Format(de_informa.rsSel_Ctc_SAC.Fields("valmerc"), "##,###,##0.00")
        lblPeso = Format(de_informa.rsSel_Ctc_SAC.Fields("peso"), "##,##0.0")
        lblVolumes = Format(de_informa.rsSel_Ctc_SAC.Fields("volumes"), "##,##0")
        If de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "URGÊNCIA" Or _
            de_informa.rsSel_Ctc_SAC.Fields("prioridade") = "PRIORIDADE" Then
            LblPrioridade.ForeColor = &HC0&
        Else
            LblPrioridade.ForeColor = &H80000012
        End If
        LblPrioridade = de_informa.rsSel_Ctc_SAC.Fields("prioridade")
        
        lblModal = de_informa.rsSel_Ctc_SAC.Fields("modal")
        lblObs_Emissao = de_informa.rsSel_Ctc_SAC.Fields("obs_emissao")
        xtemocorr = de_informa.rsSel_Ctc_SAC.Fields("tem_ocorr")
    
    'ATUALIZA COM DADOS DE ENTREGA E OCORRÊNCIA
    
        'consulta que traz os campos que são dados de ocorrência
        If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
        de_informa.Sel_ConsOcorr2 gridEmbarqueIntec.Columns(0), "01"
        Set GridConsOcorr.DataSource = de_informa
        GridConsOcorr.DataMember = "Sel_ConsOcorr2"
        GridConsOcorr.Refresh
        If de_informa.rsSel_ConsOcorr2.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.Fields("obs_ocorr") <> "" Then
                lblObs_Ocorr = de_informa.rsSel_ConsOcorr2.Fields("obs_ocorr")
            Else
                lblObs_Ocorr = ""
            End If
        Else
            lblObs_Ocorr = ""
        End If
        
        'consulta que traz os campos = 01 que é dado de entrega (ENTREGA REALIZADA)
        If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
        de_informa.Sel_ConsOcorr gridEmbarqueIntec.Columns(0), "01"
        If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
        'atualiza os campos referente a dados de entrega
            lblDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("data")
            lblHsBaixaPre = de_informa.rsSel_ConsOcorr.Fields("hora")
            lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")) = False Then
                lblObsEntr = de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")
            Else
                lblObsEntr = ""
            End If
        End If
    
    'dados de status do ctc
        
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
            If de_informa.rsSel_Ctc_SAC.Fields("prev_entrega") >= datahora("data") Then
                lblEntregueSN.ForeColor = &HC00000             'LABEL NA COR AZUL
                lblEntregueSN.Caption = "EM TRÂNSITO"
                lblEntregueSN.ToolTipText = "EM TRÂNSITO = Até a Previsão de Entrega"
            Else
                lblEntregueSN.ForeColor = &HC0&               'LABEL NA COR VERMELHO
                lblEntregueSN.Caption = "SEM POSIÇÃO"
                lblEntregueSN.ToolTipText = "SEM POSIÇÃO = Após a Previsão de Entrega"
            End If
        ElseIf xtemocorr = "C" Then
            lblEntregueSN.ForeColor = &HC0&              'LABEL NA COR VERMELHO
            lblEntregueSN.Caption = "CTC CANCELADO"
            lblEntregueSN.ToolTipText = "Cancelado em:" & de_informa.rsSel_Ctc_SAC.Fields("canc_data") & _
                                        "  Usuário:" & de_informa.rsSel_Ctc_SAC.Fields("canc_usu") & _
                                        "  Motivo:" & de_informa.rsSel_Ctc_SAC.Fields("canc_obs")
        End If
    
        DoEvents
    
    'Notas Fiscais do CTC
        
        If de_informa.rsSel_NFsdoCTC.State = 1 Then de_informa.rsSel_NFsdoCTC.Close
        de_informa.Sel_NFsdoCTC gridEmbarqueIntec.Columns(0)
        
        gridNFsEspeciais.DataMember = "sel_nfsdoctc"
        gridNFsEspeciais.Refresh
        
    'ATUALIZA GRID DE MANIFESTO
        If de_informa.rsSel_ManifestoPorCTC.State = 1 Then de_informa.rsSel_ManifestoPorCTC.Close
        de_informa.Sel_ManifestoPorCTC gridEmbarqueIntec.Columns(0)
        gridManifesto.DataMember = "sel_manifestoporctc"
        gridManifesto.Refresh
        
        txtNumNf.SetFocus
        
        DoEvents
        
    'ÚLTIMAS CONSULTAS DO CTC
        de_informa.ins_ultconssac gridEmbarqueIntec.Columns(0), xusuario, datahora("datahora")
            
    'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "CONSULTA", xusuario, "INFORMAÇÃO SAC - CONSULTA CTC: " & gridEmbarqueIntec.Columns(0)
        
    End If
            

End Sub

Private Sub gridEmbarqueIntec_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    gridEmbarqueIntec_Click
End Sub

Private Sub gridNFsEspeciais_Click()
    If de_informa.rsSel_NFBasecli.State = 1 Then de_informa.rsSel_NFBasecli.Close
    de_informa.Sel_NFBasecli SoNumeros(lblRemet_CGC), gridNFsEspeciais.Columns(3), gridNFsEspeciais.Columns(4)
    gridClienteEspeciais.DataMember = "Sel_NFBasecli"
    gridClienteEspeciais.Refresh
    
    If de_informa.rsSel_NFVideolar.State = 1 Then de_informa.rsSel_NFVideolar.Close
    de_informa.Sel_NFVideolar SoNumeros(lblRemet_CGC), CDbl(gridNFsEspeciais.Columns(3)), gridNFsEspeciais.Columns(4)
    gridVideolarEspeciais.DataMember = "Sel_NFVideolar"
    gridVideolarEspeciais.Refresh
    
    If de_informa.rsSel_NFVideolarItens.State = 1 Then de_informa.rsSel_NFVideolarItens.Close
    de_informa.Sel_NFVideolarItens SoNumeros(lblRemet_CGC), CDbl(gridNFsEspeciais.Columns(3)), gridNFsEspeciais.Columns(4)
    gridVideolarEspeciaisItem.DataMember = "Sel_NFVideolarItens"
    gridVideolarEspeciaisItem.Refresh
    
    If de_informa.rsSel_NFCheckReceb.State = 1 Then de_informa.rsSel_NFCheckReceb.Close
    de_informa.Sel_NFCheckReceb SoNumeros(lblRemet_CGC), CDbl(gridNFsEspeciais.Columns(3)), gridNFsEspeciais.Columns(4)
    gridCheckEspeciais.DataMember = "Sel_NFCheckReceb"
    gridCheckEspeciais.Refresh
    
    If de_informa.rsSel_VLNFSerie_Sac.State = 1 Then de_informa.rsSel_VLNFSerie_Sac.Close
    de_informa.Sel_VLNFSerie_Sac CDbl(gridNFsEspeciais.Columns(3)), gridNFsEspeciais.Columns(4)
    gridEmbarqueIntec.DataMember = "Sel_VLNFSerie_Sac"
    gridEmbarqueIntec.Refresh
    
End Sub

Private Sub optPorNF_Click()
    If optPorNF.Value = True Then
        txtNumNf.Enabled = True
        txtNumNf.BackColor = &HC0FFFF
        txtNumPedido.Enabled = False
        txtNumPedido.BackColor = &H80000005
        txtNumNf.SetFocus
        txtNumPedido.Text = ""
    Else
        txtNumNf.Enabled = False
        txtNumNf.BackColor = &H80000005
        txtNumPedido.Enabled = True
        txtNumPedido.BackColor = &HC0FFFF
        txtNumPedido.SetFocus
        txtNumNf.Text = ""
    End If
    DoEvents
End Sub

Private Sub optPorNF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optPorPedido_Click()
    If optPorNF.Value = True Then
        txtNumNf.Enabled = True
        txtNumNf.BackColor = &HC0FFFF
        txtNumPedido.Enabled = False
        txtNumPedido.BackColor = &H80000005
        txtNumNf.SetFocus
        txtNumPedido.Text = ""
    Else
        txtNumNf.Enabled = False
        txtNumNf.BackColor = &H80000005
        txtNumPedido.Enabled = True
        txtNumPedido.BackColor = &HC0FFFF
        txtNumPedido.SetFocus
        txtNumNf.Text = ""
    End If
    DoEvents
End Sub

Private Sub optPorPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
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
            If optPorNF.Value = True Then
                optPorPedido.Value = True
            Else
                optPorNF.Value = True
            End If
        Else
            KeyAscii = 0
            SendKeys "{TAB}"  'ENVIA UM TAB
        End If
    End If
End Sub
Private Sub txtNumNf_LostFocus()
    If txtNumNf.Text <> "" Then
        If Not IsNumeric(txtNumNf.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            txtNumNf.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtNumPedido_GotFocus()
    txtNumPedido.SelStart = 0
    txtNumPedido.SelLength = 20
End Sub

Private Sub txtNumPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        If txtNumPedido.Text = "" Then
            KeyAscii = 0
            If optPorNF.Value = True Then
                optPorPedido.Value = True
            Else
                optPorNF.Value = True
            End If
        Else
            KeyAscii = 0
            SendKeys "{TAB}"  'ENVIA UM TAB
        End If
    End If
End Sub
Private Sub txtNumPedido_LostFocus()
    txtNumPedido.Text = UCase(txtNumPedido)
End Sub

Private Sub txtSerie_gotfocus()
    txtSerie.SelStart = 0
    txtSerie.SelLength = 3
End Sub
Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtSerie_LostFocus()
    If txtSerie.Text <> "" Then
        If Not IsNumeric(txtSerie.Text) Then
            MsgBox "Dado Inválido !", vbCritical, "Erro"
            txtSerie.SetFocus
            Exit Sub
        End If
    End If
End Sub

