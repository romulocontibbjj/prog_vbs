VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   1185
   ClientTop       =   1470
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   11895
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Resumo"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "CTCs Fiscais"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Notas Fiscais de Serviço"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "CTRs (Minutas)"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Faturamento"
      TabPicture(4)   =   "Form1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Frame Frame4 
         Caption         =   "Não Faturado"
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
         TabIndex        =   7
         Top             =   4680
         Width           =   11415
         Begin VB.CommandButton Command8 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   43
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Imprimir/Arquivo"
            Height          =   375
            Left            =   8760
            TabIndex        =   13
            Top             =   1440
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid gridNaoFat 
            Bindings        =   "Form1.frx":008C
            Height          =   1575
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsNaoFat"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "origem"
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
               DataField       =   "fretefinal"
               Caption         =   "fretefinal"
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
               DataField       =   "valmerc"
               Caption         =   "valmerc"
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
                  ColumnWidth     =   1289,764
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label lblTotFreteNaoFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   31
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblTotValMercNaoFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   30
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotCtcNaoFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   29
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   28
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Total Frete:"
            Height          =   195
            Left            =   7440
            TabIndex        =   27
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   26
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Faturado"
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
         TabIndex        =   6
         Top             =   2520
         Width           =   11415
         Begin VB.CommandButton Command7 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   42
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Imprimir/Arquivo"
            Height          =   375
            Left            =   8760
            TabIndex        =   12
            Top             =   1440
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid gridFaturado 
            Bindings        =   "Form1.frx":00A5
            Height          =   1575
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsFaturados"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "origem"
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
               DataField       =   "fretefinal"
               Caption         =   "fretefinal"
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
               DataField       =   "valmerc"
               Caption         =   "valmerc"
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
                  ColumnWidth     =   1289,764
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label lblTotFreteFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   25
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblTotValMercFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   24
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotCtcFat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   23
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   22
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Total Frete:"
            Height          =   195
            Left            =   7440
            TabIndex        =   21
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   20
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total Emitido"
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
         TabIndex        =   5
         Top             =   480
         Width           =   11415
         Begin VB.CommandButton Command6 
            Caption         =   "Imprimir Tela"
            Height          =   375
            Left            =   7200
            TabIndex        =   41
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Imprimir/Arquivo"
            Height          =   375
            Left            =   8760
            TabIndex        =   11
            Top             =   1440
            Width           =   2535
         End
         Begin MSDataGridLib.DataGrid gridEmitido 
            Bindings        =   "Form1.frx":00BE
            Height          =   1575
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2778
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
            DataMember      =   "Sel_GerCtcsEmitidos"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "origem"
               Caption         =   "origem"
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
               DataField       =   "fretefinal"
               Caption         =   "fretefinal"
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
               DataField       =   "valmerc"
               Caption         =   "valmerc"
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
                  ColumnWidth     =   1289,764
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label lblTotFreteEmit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   19
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblTotValMercEmit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   18
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblTotCtcEmit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8880
            TabIndex        =   17
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total Valor Merc:"
            Height          =   195
            Left            =   7440
            TabIndex        =   16
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Total Frete:"
            Height          =   195
            Left            =   7440
            TabIndex        =   15
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total de CTCs:"
            Height          =   195
            Left            =   7440
            TabIndex        =   14
            Top             =   240
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.OptionButton Option3 
         Caption         =   "Aereo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2520
         TabIndex        =   46
         Top             =   680
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Rodo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1440
         TabIndex        =   45
         Top             =   680
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rodo+Aereo"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   680
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Sair"
         Height          =   615
         Left            =   10800
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar"
         Height          =   615
         Left            =   9720
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   34
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPer1 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   38
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Responsável:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6600
         TabIndex        =   36
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   35
         Top             =   540
         Width           =   3015
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Responsável:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3480
         TabIndex        =   33
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Período:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   90
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()

End Sub

Private Sub cmdProcessar_Click()
    If de_informa.rsSel_GerCtcsEmitidos.State = 1 Then de_informa.rsSel_GerCtcsEmitidos.Close
    de_informa.Sel_GerCtcsEmitidos mskPer1, mskPer2, "%", "%", "%"
    
    If de_informa.rsSel_GerCtcsFaturados.State = 1 Then de_informa.rsSel_GerCtcsFaturados.Close
    de_informa.Sel_GerCtcsFaturados mskPer1, mskPer2, "%", "%", "%"
    
    If de_informa.rsSel_GerCtcsNaoFat.State = 1 Then de_informa.rsSel_GerCtcsNaoFat.Close
    de_informa.Sel_GerCtcsNaoFat mskPer1, mskPer2, "%", "%", "%"
    
    gridEmitido.DataMember = "Sel_GerCtcsEmitidos"
    gridEmitido.Refresh
    
    gridFaturado.DataMember = "Sel_GerCtcsFaturados"
    gridFaturado.Refresh
    
    gridNaoFat.DataMember = "Sel_GerCtcsNaoFat "
    gridNaoFat.Refresh
    
    
    de_informa.rsSel_GerCtcsEmitidos.MoveFirst
    Do Until de_informa.rsSel_GerCtcsEmitidos.EOF
        xTotCtcEmit = xTotCtcEmit + de_informa.rsSel_GerCtcsEmitidos.Fields("qtde")
        xTotValMercEmit = xTotValMercEmit + de_informa.rsSel_GerCtcsEmitidos.Fields("valmerc")
        xTotFreteEmit = xTotFreteEmit + de_informa.rsSel_GerCtcsEmitidos.Fields("fretefinal")
        de_informa.rsSel_GerCtcsEmitidos.MoveNext
    Loop
    lblTotCtcEmit = Format(xTotCtcEmit, "##,###,##0.00")
    lblTotValMercEmit = Format(xTotValMercEmit, "##,###,##0.00")
    lblTotFreteEmit = Format(xTotFreteEmit, "##,###,##0.00")
    
    
    
    
    
    
    
    
End Sub

Private Sub lblTotValMercNaoFat_Click()

End Sub
