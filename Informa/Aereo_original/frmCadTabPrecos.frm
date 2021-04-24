VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCadTabPrecos 
   Caption         =   "Tabela de Preços Cia. Aérea"
   ClientHeight    =   8340
   ClientLeft      =   105
   ClientTop       =   990
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   12030
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAtualiza 
      Caption         =   "A T U A L I Z A R"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4800
      TabIndex        =   16
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Alterar Preços  Tab. TE TC ..."
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Alterar Preços  Tab. Geral..."
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Incluir Localidade Nesta Tabela..."
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "S A I R"
      Height          =   495
      Left            =   10440
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir Tabela..."
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reajuste de Preços..."
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdInclTabelaNova 
      Caption         =   "Incluir Nova Tabela..."
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalhe Tabela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Width           =   6615
      Begin TabDlg.SSTab tabTabelas 
         Height          =   5535
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   9763
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Tabela Geral"
         TabPicture(0)   =   "frmCadTabPrecos.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "gridTabGeral"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tabela T.E.  T.C."
         TabPicture(1)   =   "frmCadTabPrecos.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "gridTabTETC"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin MSDataGridLib.DataGrid gridTabGeral 
            Bindings        =   "frmCadTabPrecos.frx":0038
            Height          =   4815
            Left            =   -74880
            TabIndex        =   6
            Top             =   600
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   8493
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
            DataMember      =   "Sel_TabPrecoGeral"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               AllowRowSizing  =   -1  'True
               AllowSizing     =   -1  'True
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid gridTabTETC 
            Bindings        =   "frmCadTabPrecos.frx":0051
            Height          =   4695
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   8281
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
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
            DataMember      =   "Sel_TabPrecoTETC"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               AllowRowSizing  =   -1  'True
               AllowSizing     =   -1  'True
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Descrição Tabela Específica / Tarifa Charter"
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
            TabIndex        =   12
            Top             =   480
            Width           =   6135
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tabelas"
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
      TabIndex        =   2
      Top             =   2640
      Width           =   5055
      Begin MSDataGridLib.DataGrid gridTabelas 
         Bindings        =   "frmCadTabPrecos.frx":006A
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5530
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         DataMember      =   "Sel_TabPreco"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            AllowRowSizing  =   -1  'True
            AllowSizing     =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cia. Aérea"
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
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   2400
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid gridCia 
         Bindings        =   "frmCadTabPrecos.frx":0083
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3201
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         DataMember      =   "Sel_CiaAerea"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            AllowRowSizing  =   -1  'True
            AllowSizing     =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCadTabPrecos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualiza_Click()
    If de_informa.rsSel_CiaAerea.State = 1 Then de_informa.rsSel_CiaAerea.Close
    de_informa.Sel_CiaAerea
    gridCia.DataMember = "sel_ciaaerea"
    gridCia.Refresh

    If de_informa.rsSel_TabPreco.State = 1 Then de_informa.rsSel_TabPreco.Close
    de_informa.Sel_TabPreco gridCia.Columns(0)
    gridTabelas.DataMember = "sel_tabpreco"
    gridTabelas.Refresh
    
    If de_informa.rsSel_TabPreco.RecordCount > 0 Then
        If de_informa.rsSel_TabPrecoGeral.State = 1 Then de_informa.rsSel_TabPrecoGeral.Close
        de_informa.Sel_TabPrecoGeral gridTabelas.Columns(0)
        gridTabGeral.DataMember = "sel_tabprecogeral"
        gridTabGeral.Refresh
        
        If de_informa.rsSel_TabPrecoTETC.State = 1 Then de_informa.rsSel_TabPrecoTETC.Close
        de_informa.Sel_TabPrecoTETC gridTabelas.Columns(0)
        gridTabTETC.DataMember = "sel_tabprecotetc"
        gridTabTETC.Refresh
        
        gridTabelas.Enabled = True
    Else
        If de_informa.rsSel_TabPrecoGeral.State = 1 Then de_informa.rsSel_TabPrecoGeral.Close
        de_informa.Sel_TabPrecoGeral -1000
        gridTabGeral.DataMember = "sel_tabprecogeral"
        gridTabGeral.Refresh
        
        If de_informa.rsSel_TabPrecoTETC.State = 1 Then de_informa.rsSel_TabPrecoTETC.Close
        de_informa.Sel_TabPrecoTETC -1000
        gridTabTETC.DataMember = "sel_tabprecotetc"
        gridTabTETC.Refresh
        
        gridTabelas.Enabled = False
    End If

End Sub

Private Sub cmdInclTabelaNova_Click()
    frmIncluiTabAerea.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    cmdAtualiza_Click
End Sub

Private Sub gridCia_Click()
    
    If de_informa.rsSel_TabPreco.State = 1 Then de_informa.rsSel_TabPreco.Close
    de_informa.Sel_TabPreco gridCia.Columns(0)
    gridTabelas.DataMember = "sel_tabpreco"
    gridTabelas.Refresh
    
    If de_informa.rsSel_TabPreco.RecordCount > 0 Then
        If de_informa.rsSel_TabPrecoGeral.State = 1 Then de_informa.rsSel_TabPrecoGeral.Close
        de_informa.Sel_TabPrecoGeral gridTabelas.Columns(0)
        gridTabGeral.DataMember = "sel_tabprecogeral"
        gridTabGeral.Refresh
        
        If de_informa.rsSel_TabPrecoTETC.State = 1 Then de_informa.rsSel_TabPrecoTETC.Close
        de_informa.Sel_TabPrecoTETC gridTabelas.Columns(0)
        gridTabTETC.DataMember = "sel_tabprecotetc"
        gridTabTETC.Refresh
        
        gridTabelas.Enabled = True
    Else
        If de_informa.rsSel_TabPrecoGeral.State = 1 Then de_informa.rsSel_TabPrecoGeral.Close
        de_informa.Sel_TabPrecoGeral -1000
        gridTabGeral.DataMember = "sel_tabprecogeral"
        gridTabGeral.Refresh
        
        If de_informa.rsSel_TabPrecoTETC.State = 1 Then de_informa.rsSel_TabPrecoTETC.Close
        de_informa.Sel_TabPrecoTETC -1000
        gridTabTETC.DataMember = "sel_tabprecotetc"
        gridTabTETC.Refresh
        
        gridTabelas.Enabled = False
    End If
    
End Sub

Private Sub gridTabelas_Click()
    
    If de_informa.rsSel_TabPrecoGeral.State = 1 Then de_informa.rsSel_TabPrecoGeral.Close
    de_informa.Sel_TabPrecoGeral gridTabelas.Columns(0)
    gridTabGeral.DataMember = "sel_tabprecogeral"
    gridTabGeral.Refresh
    
    If de_informa.rsSel_TabPrecoTETC.State = 1 Then de_informa.rsSel_TabPrecoTETC.Close
    de_informa.Sel_TabPrecoTETC gridTabelas.Columns(0)
    gridTabTETC.DataMember = "sel_tabprecotetc"
    gridTabTETC.Refresh
End Sub
