VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEscCliPer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolher Cliente e Período"
   ClientHeight    =   6630
   ClientLeft      =   930
   ClientTop       =   1485
   ClientWidth     =   7725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7725
   Begin VB.Frame fraDados 
      Caption         =   "Seleção dos Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtFilial 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   3
         Top             =   1200
         Width           =   495
      End
      Begin VB.Frame fraAnalise 
         Caption         =   "Analise"
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
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
         Begin VB.CheckBox chkAnalise 
            Caption         =   "Tratar atrasos analisados como entregue No Prazo. (Ocorrências)"
            Height          =   615
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Value           =   1  'Checked
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Modal"
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
         Left            =   2640
         TabIndex        =   29
         Top             =   1680
         Width           =   2415
         Begin VB.CheckBox chkModal 
            Caption         =   "Todos Modais"
            Enabled         =   0   'False
            Height          =   195
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.OptionButton optRodo 
            Caption         =   "Rodoviário"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1200
            TabIndex        =   6
            Top             =   600
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optAir 
            Caption         =   "Aéreo"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   450
         Left            =   5160
         TabIndex        =   28
         Top             =   2160
         Width           =   2115
      End
      Begin VB.CheckBox chkTodosCli 
         Caption         =   "Todos os Clientes Remetentes"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtCGCCli 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   0
         Top             =   720
         Width           =   1590
      End
      Begin VB.CheckBox chkTodosEstab 
         Caption         =   "Todos os Estabelecimentos"
         Height          =   225
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CommandButton cmdProcessa 
         Caption         =   "Processa..."
         Enabled         =   0   'False
         Height          =   450
         Left            =   5160
         TabIndex        =   7
         Top             =   1560
         Width           =   2115
      End
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   2760
         TabIndex        =   2
         Top             =   1200
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
         Left            =   1200
         TabIndex        =   1
         Top             =   1200
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
      Begin VB.Label lblFilial 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4080
         TabIndex        =   32
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CGC do Cliente/Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Período:  De"
         Height          =   195
         Left            =   105
         TabIndex        =   20
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Top             =   1260
         Width           =   90
      End
      Begin VB.Label lblNomeCli 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3960
         TabIndex        =   18
         Top             =   720
         Width           =   3375
      End
   End
   Begin VB.Frame fraConsCli 
      Caption         =   "Consulta Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   7485
      Begin VB.OptionButton optBuscaTodo 
         Caption         =   "Busca no Texto Todo"
         Height          =   195
         Left            =   3675
         TabIndex        =   10
         Top             =   3150
         Width           =   2220
      End
      Begin VB.OptionButton optBuscaInic 
         Caption         =   "Busca no Início do Texto"
         Height          =   195
         Left            =   3675
         TabIndex        =   9
         Top             =   2940
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca"
         Height          =   330
         Left            =   6195
         TabIndex        =   11
         Top             =   3045
         Width           =   1065
      End
      Begin VB.TextBox txtBuscaNome 
         Height          =   285
         Left            =   1575
         MaxLength       =   25
         TabIndex        =   8
         Top             =   3045
         Width           =   1905
      End
      Begin MSDataGridLib.DataGrid GridConsCli 
         Bindings        =   "frmEscCliPer.frx":0000
         Height          =   2535
         Left            =   120
         TabIndex        =   13
         Top             =   315
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   4471
         _Version        =   393216
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cgc"
            Caption         =   "cgc"
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
            Caption         =   "nome"
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
            DataField       =   "cidade"
            Caption         =   "cidade"
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
               ColumnWidth     =   1470,047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3750,236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1590,236
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Busca por Nome:"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   3045
         Width           =   1230
      End
   End
   Begin VB.Frame fraMensagem 
      Caption         =   "Mensagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   7485
      Begin MSComCtl2.Animation Animation1 
         Height          =   645
         Left            =   2760
         TabIndex        =   23
         Top             =   240
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1138
         _Version        =   393216
         AutoPlay        =   -1  'True
         Center          =   -1  'True
         FullWidth       =   295
         FullHeight      =   43
      End
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   960
         Width           =   6375
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   945
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Processando dados selecionados. Por favor aguarde ..."
         Height          =   480
         Left            =   105
         TabIndex        =   24
         Top             =   315
         Width           =   2565
      End
   End
   Begin VB.Label lblAguarde 
      AutoSize        =   -1  'True
      Caption         =   "Dados em Processamento. Por favor aguarde ..."
      Height          =   195
      Left            =   2100
      TabIndex        =   15
      Top             =   3150
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   120
      X2              =   7560
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "frmEscCliPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub proc_dados()
    'Dimensiona as Matrizes para as variáveis de cálculo
    Dim xNONE1(16) As Currency, xNONE2(16) As Currency, xNONE3(16) As Currency, xNONE4(16) As Currency, xNONE5(16) As Currency, xNONE6(16) As Currency, xNONE7(16) As Single, xNONE8(16) As Single
    Dim xSDSU1(12) As Currency, xSDSU2(12) As Currency, xSDSU3(12) As Currency, xSDSU4(12) As Currency, xSDSU5(12) As Currency, xSDSU6(12) As Currency, xSDSU7(12) As Single, xSDSU8(12) As Single
    Dim xPeso1(10) As Currency, xPeso2(10) As Currency, xPeso3(10) As Currency, xPeso4(10) As Currency, xPeso5(10) As Currency, xPeso6(10) As Currency
    Dim xNF1(9) As Currency, xNF2(9) As Currency, xNF3(9) As Currency, xNF4(9) As Currency, xNF5(9) As Currency, xNF6(9) As Currency
    Dim xy As Integer
    de_informa.rsSel_AnOper.MoveFirst
    'Inicia cálculo registro/registro do recordset da seleção
    Do Until de_informa.rsSel_AnOper.EOF
        If de_informa.rsSel_AnOper.Fields("uf_dest") = "AC" Then
            xNONE1(1) = xNONE1(1) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(1) = xNONE2(1) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(1) = xNONE3(1) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(1) = xNONE8(1) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(1) = xNONE4(1) + 1
            If xNONE4(1) > 0 Then xNONE5(1) = xNONE3(1) / xNONE4(1)
            If xNONE3(1) > 0 Then xNONE6(1) = xNONE2(1) / xNONE3(1)
            If xNONE1(1) > 0 Then xNONE7(1) = xNONE2(1) / xNONE1(1)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "AM" Then
            xNONE1(2) = xNONE1(2) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(2) = xNONE2(2) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(2) = xNONE3(2) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(2) = xNONE8(2) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(2) = xNONE4(2) + 1
            If xNONE4(2) > 0 Then xNONE5(2) = xNONE3(2) / xNONE4(2)
            If xNONE3(2) > 0 Then xNONE6(2) = xNONE2(2) / xNONE3(2)
            If xNONE1(2) > 0 Then xNONE7(2) = xNONE2(2) / xNONE1(2)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "AP" Then
            xNONE1(3) = xNONE1(3) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(3) = xNONE2(3) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(3) = xNONE3(3) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(3) = xNONE8(3) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(3) = xNONE4(3) + 1
            If xNONE4(3) > 0 Then xNONE5(3) = xNONE3(3) / xNONE4(3)
            If xNONE3(3) > 0 Then xNONE6(3) = xNONE2(3) / xNONE3(3)
            If xNONE1(3) > 0 Then xNONE7(3) = xNONE2(3) / xNONE1(3)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "PA" Then
            xNONE1(4) = xNONE1(4) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(4) = xNONE2(4) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(4) = xNONE3(4) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(4) = xNONE8(4) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(4) = xNONE4(4) + 1
            If xNONE4(4) > 0 Then xNONE5(4) = xNONE3(4) / xNONE4(4)
            If xNONE3(4) > 0 Then xNONE6(4) = xNONE2(4) / xNONE3(4)
            If xNONE1(4) > 0 Then xNONE7(4) = xNONE2(4) / xNONE1(4)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "RO" Then
            xNONE1(5) = xNONE1(5) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(5) = xNONE2(5) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(5) = xNONE3(5) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(5) = xNONE8(5) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(5) = xNONE4(5) + 1
            If xNONE4(5) > 0 Then xNONE5(5) = xNONE3(5) / xNONE4(5)
            If xNONE3(5) > 0 Then xNONE6(5) = xNONE2(5) / xNONE3(5)
            If xNONE1(5) > 0 Then xNONE7(5) = xNONE2(5) / xNONE1(5)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "RR" Then
            xNONE1(6) = xNONE1(6) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(6) = xNONE2(6) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(6) = xNONE3(6) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(6) = xNONE8(6) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(6) = xNONE4(6) + 1
            If xNONE4(6) > 0 Then xNONE5(6) = xNONE3(6) / xNONE4(6)
            If xNONE3(6) > 0 Then xNONE6(6) = xNONE2(6) / xNONE3(6)
            If xNONE1(6) > 0 Then xNONE7(6) = xNONE2(6) / xNONE1(6)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "TO" Then
            xNONE1(7) = xNONE1(7) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(7) = xNONE2(7) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(7) = xNONE3(7) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(7) = xNONE8(7) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(7) = xNONE4(7) + 1
            If xNONE4(7) > 0 Then xNONE5(7) = xNONE3(7) / xNONE4(7)
            If xNONE3(7) > 0 Then xNONE6(7) = xNONE2(7) / xNONE3(7)
            If xNONE1(7) > 0 Then xNONE7(7) = xNONE2(7) / xNONE1(7)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "AL" Then
            xNONE1(8) = xNONE1(8) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(8) = xNONE2(8) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(8) = xNONE3(8) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(8) = xNONE8(8) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(8) = xNONE4(8) + 1
            If xNONE4(8) > 0 Then xNONE5(8) = xNONE3(8) / xNONE4(8)
            If xNONE3(8) > 0 Then xNONE6(8) = xNONE2(8) / xNONE3(8)
            If xNONE1(8) > 0 Then xNONE7(8) = xNONE2(8) / xNONE1(8)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "BA" Then
            xNONE1(9) = xNONE1(9) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(9) = xNONE2(9) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(9) = xNONE3(9) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(9) = xNONE8(9) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(9) = xNONE4(9) + 1
            If xNONE4(9) > 0 Then xNONE5(9) = xNONE3(9) / xNONE4(9)
            If xNONE3(9) > 0 Then xNONE6(9) = xNONE2(9) / xNONE3(9)
            If xNONE1(9) > 0 Then xNONE7(9) = xNONE2(9) / xNONE1(9)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "SE" Then
            xNONE1(10) = xNONE1(10) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(10) = xNONE2(10) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(10) = xNONE3(10) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(10) = xNONE8(10) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(10) = xNONE4(10) + 1
            If xNONE4(10) > 0 Then xNONE5(10) = xNONE3(10) / xNONE4(10)
            If xNONE3(10) > 0 Then xNONE6(10) = xNONE2(10) / xNONE3(10)
            If xNONE1(10) > 0 Then xNONE7(10) = xNONE2(10) / xNONE1(10)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "PE" Then
            xNONE1(11) = xNONE1(11) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(11) = xNONE2(11) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(11) = xNONE3(11) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(11) = xNONE8(11) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(11) = xNONE4(11) + 1
            If xNONE4(11) > 0 Then xNONE5(11) = xNONE3(11) / xNONE4(11)
            If xNONE3(11) > 0 Then xNONE6(11) = xNONE2(11) / xNONE3(11)
            If xNONE1(11) > 0 Then xNONE7(11) = xNONE2(11) / xNONE1(11)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "PB" Then
            xNONE1(12) = xNONE1(12) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(12) = xNONE2(12) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(12) = xNONE3(12) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(12) = xNONE8(12) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(12) = xNONE4(12) + 1
            If xNONE4(12) > 0 Then xNONE5(12) = xNONE3(12) / xNONE4(12)
            If xNONE3(12) > 0 Then xNONE6(12) = xNONE2(12) / xNONE3(12)
            If xNONE1(12) > 0 Then xNONE7(12) = xNONE2(12) / xNONE1(12)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "RN" Then
            xNONE1(13) = xNONE1(13) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(13) = xNONE2(13) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(13) = xNONE3(13) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(13) = xNONE8(13) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(13) = xNONE4(13) + 1
            If xNONE4(13) > 0 Then xNONE5(13) = xNONE3(13) / xNONE4(13)
            If xNONE3(13) > 0 Then xNONE6(13) = xNONE2(13) / xNONE3(13)
            If xNONE1(13) > 0 Then xNONE7(13) = xNONE2(13) / xNONE1(13)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "CE" Then
            xNONE1(14) = xNONE1(14) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(14) = xNONE2(14) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(14) = xNONE3(14) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(14) = xNONE8(14) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(14) = xNONE4(14) + 1
            If xNONE4(14) > 0 Then xNONE5(14) = xNONE3(14) / xNONE4(14)
            If xNONE3(14) > 0 Then xNONE6(14) = xNONE2(14) / xNONE3(14)
            If xNONE1(14) > 0 Then xNONE7(14) = xNONE2(14) / xNONE1(14)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "PI" Then
            xNONE1(15) = xNONE1(15) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(15) = xNONE2(15) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(15) = xNONE3(15) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(15) = xNONE8(15) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(15) = xNONE4(15) + 1
            If xNONE4(15) > 0 Then xNONE5(15) = xNONE3(15) / xNONE4(15)
            If xNONE3(15) > 0 Then xNONE6(15) = xNONE2(15) / xNONE3(15)
            If xNONE1(15) > 0 Then xNONE7(15) = xNONE2(15) / xNONE1(15)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "MA" Then
            xNONE1(16) = xNONE1(16) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNONE2(16) = xNONE2(16) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNONE3(16) = xNONE3(16) + de_informa.rsSel_AnOper.Fields("peso")
            xNONE8(16) = xNONE8(16) + de_informa.rsSel_AnOper.Fields("volumes")
            xNONE4(16) = xNONE4(16) + 1
            If xNONE4(16) > 0 Then xNONE5(16) = xNONE3(16) / xNONE4(16)
            If xNONE3(16) > 0 Then xNONE6(16) = xNONE2(16) / xNONE3(16)
            If xNONE1(16) > 0 Then xNONE7(16) = xNONE2(16) / xNONE1(16)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "ES" Then
            xSDSU1(1) = xSDSU1(1) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(1) = xSDSU2(1) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(1) = xSDSU3(1) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(1) = xSDSU8(1) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(1) = xSDSU4(1) + 1
            If xSDSU4(1) > 0 Then xSDSU5(1) = xSDSU3(1) / xSDSU4(1)
            If xSDSU3(1) > 0 Then xSDSU6(1) = xSDSU2(1) / xSDSU3(1)
            If xSDSU1(1) > 0 Then xSDSU7(1) = xSDSU2(1) / xSDSU1(1)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "MG" Then
            xSDSU1(2) = xSDSU1(2) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(2) = xSDSU2(2) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(2) = xSDSU3(2) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(2) = xSDSU8(2) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(2) = xSDSU4(2) + 1
            If xSDSU4(2) > 0 Then xSDSU5(2) = xSDSU3(2) / xSDSU4(2)
            If xSDSU3(2) > 0 Then xSDSU6(2) = xSDSU2(2) / xSDSU3(2)
            If xSDSU1(2) > 0 Then xSDSU7(2) = xSDSU2(2) / xSDSU1(2)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "RJ" Then
            xSDSU1(3) = xSDSU1(3) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(3) = xSDSU2(3) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(3) = xSDSU3(3) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(3) = xSDSU8(3) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(3) = xSDSU4(3) + 1
            If xSDSU4(3) > 0 Then xSDSU5(3) = xSDSU3(3) / xSDSU4(3)
            If xSDSU3(3) > 0 Then xSDSU6(3) = xSDSU2(3) / xSDSU3(3)
            If xSDSU1(3) > 0 Then xSDSU7(3) = xSDSU2(3) / xSDSU1(3)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "SP" Then
            xSDSU1(4) = xSDSU1(4) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(4) = xSDSU2(4) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(4) = xSDSU3(4) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(4) = xSDSU8(4) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(4) = xSDSU4(4) + 1
            If xSDSU4(4) > 0 Then xSDSU5(4) = xSDSU3(4) / xSDSU4(4)
            If xSDSU3(4) > 0 Then xSDSU6(4) = xSDSU2(4) / xSDSU3(4)
            If xSDSU1(4) > 0 Then xSDSU7(4) = xSDSU2(4) / xSDSU1(4)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "PR" Then
            xSDSU1(5) = xSDSU1(5) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(5) = xSDSU2(5) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(5) = xSDSU3(5) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(5) = xSDSU8(5) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(5) = xSDSU4(5) + 1
            If xSDSU4(5) > 0 Then xSDSU5(5) = xSDSU3(5) / xSDSU4(5)
            If xSDSU3(5) > 0 Then xSDSU6(5) = xSDSU2(5) / xSDSU3(5)
            If xSDSU1(5) > 0 Then xSDSU7(5) = xSDSU2(5) / xSDSU1(5)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "RS" Then
            xSDSU1(6) = xSDSU1(6) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(6) = xSDSU2(6) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(6) = xSDSU3(6) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(6) = xSDSU8(6) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(6) = xSDSU4(6) + 1
            If xSDSU4(6) > 0 Then xSDSU5(6) = xSDSU3(6) / xSDSU4(6)
            If xSDSU3(6) > 0 Then xSDSU6(6) = xSDSU2(6) / xSDSU3(6)
            If xSDSU1(6) > 0 Then xSDSU7(6) = xSDSU2(6) / xSDSU1(6)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "SC" Then
            xSDSU1(7) = xSDSU1(7) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(7) = xSDSU2(7) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(7) = xSDSU3(7) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(7) = xSDSU8(7) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(7) = xSDSU4(7) + 1
            If xSDSU4(7) > 0 Then xSDSU5(7) = xSDSU3(7) / xSDSU4(7)
            If xSDSU3(7) > 0 Then xSDSU6(7) = xSDSU2(7) / xSDSU3(7)
            If xSDSU1(7) > 0 Then xSDSU7(7) = xSDSU2(7) / xSDSU1(7)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "DF" Then
            xSDSU1(8) = xSDSU1(8) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(8) = xSDSU2(8) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(8) = xSDSU3(8) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(8) = xSDSU8(8) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(8) = xSDSU4(8) + 1
            If xSDSU4(8) > 0 Then xSDSU5(8) = xSDSU3(8) / xSDSU4(8)
            If xSDSU3(8) > 0 Then xSDSU6(8) = xSDSU2(8) / xSDSU3(8)
            If xSDSU1(8) > 0 Then xSDSU7(8) = xSDSU2(8) / xSDSU1(8)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "GO" Then
            xSDSU1(9) = xSDSU1(9) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(9) = xSDSU2(9) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(9) = xSDSU3(9) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(9) = xSDSU8(9) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(9) = xSDSU4(9) + 1
            If xSDSU4(9) > 0 Then xSDSU5(9) = xSDSU3(9) / xSDSU4(9)
            If xSDSU3(9) > 0 Then xSDSU6(9) = xSDSU2(9) / xSDSU3(9)
            If xSDSU1(9) > 0 Then xSDSU7(9) = xSDSU2(9) / xSDSU1(9)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "MS" Then
            xSDSU1(10) = xSDSU1(10) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(10) = xSDSU2(10) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(10) = xSDSU3(10) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(10) = xSDSU8(10) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(10) = xSDSU4(10) + 1
            If xSDSU4(10) > 0 Then xSDSU5(10) = xSDSU3(10) / xSDSU4(10)
            If xSDSU3(10) > 0 Then xSDSU6(10) = xSDSU2(10) / xSDSU3(10)
            If xSDSU1(10) > 0 Then xSDSU7(10) = xSDSU2(10) / xSDSU1(10)
        ElseIf de_informa.rsSel_AnOper.Fields("uf_dest") = "MT" Then
            xSDSU1(11) = xSDSU1(11) + de_informa.rsSel_AnOper.Fields("valmerc")
            xSDSU2(11) = xSDSU2(11) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xSDSU3(11) = xSDSU3(11) + de_informa.rsSel_AnOper.Fields("peso")
            xSDSU8(11) = xSDSU8(11) + de_informa.rsSel_AnOper.Fields("volumes")
            xSDSU4(11) = xSDSU4(11) + 1
            If xSDSU4(11) > 0 Then xSDSU5(11) = xSDSU3(11) / xSDSU4(11)
            If xSDSU3(11) > 0 Then xSDSU6(11) = xSDSU2(11) / xSDSU3(11)
            If xSDSU1(11) > 0 Then xSDSU7(11) = xSDSU2(11) / xSDSU1(11)
        End If
        
        If de_informa.rsSel_AnOper.Fields("peso") <= 20 Then
            xPeso1(1) = xPeso1(1) + 1
            xPeso3(1) = xPeso3(1) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(1) = xPeso4(1) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(1) = xPeso5(1) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("peso") > 20 And de_informa.rsSel_AnOper.Fields("peso") <= 50 Then
            xPeso1(2) = xPeso1(2) + 1
            xPeso3(2) = xPeso3(2) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(2) = xPeso4(2) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(2) = xPeso5(2) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("peso") > 50 And de_informa.rsSel_AnOper.Fields("peso") <= 100 Then
            xPeso1(3) = xPeso1(3) + 1
            xPeso3(3) = xPeso3(3) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(3) = xPeso4(3) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(3) = xPeso5(3) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("peso") > 100 And de_informa.rsSel_AnOper.Fields("peso") <= 200 Then
            xPeso1(4) = xPeso1(4) + 1
            xPeso3(4) = xPeso3(4) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(4) = xPeso4(4) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(4) = xPeso5(4) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("peso") > 200 And de_informa.rsSel_AnOper.Fields("peso") <= 500 Then
            xPeso1(5) = xPeso1(5) + 1
            xPeso3(5) = xPeso3(5) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(5) = xPeso4(5) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(5) = xPeso5(5) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("peso") > 500 And de_informa.rsSel_AnOper.Fields("peso") <= 1000 Then
            xPeso1(6) = xPeso1(6) + 1
            xPeso3(6) = xPeso3(6) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(6) = xPeso4(6) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(6) = xPeso5(6) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("peso") > 1000 And de_informa.rsSel_AnOper.Fields("peso") <= 2500 Then
            xPeso1(7) = xPeso1(7) + 1
            xPeso3(7) = xPeso3(7) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(7) = xPeso4(7) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(7) = xPeso5(7) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("peso") > 2500 And de_informa.rsSel_AnOper.Fields("peso") <= 5000 Then
            xPeso1(8) = xPeso1(8) + 1
            xPeso3(8) = xPeso3(8) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(8) = xPeso4(8) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(8) = xPeso5(8) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("peso") > 5000 Then
            xPeso1(9) = xPeso1(9) + 1
            xPeso3(9) = xPeso3(9) + de_informa.rsSel_AnOper.Fields("valmerc")
            xPeso4(9) = xPeso4(9) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xPeso5(9) = xPeso5(9) + de_informa.rsSel_AnOper.Fields("peso")
        End If
        
        If de_informa.rsSel_AnOper.Fields("valmerc") <= 100 Then
            xNF1(1) = xNF1(1) + 1
            xNF3(1) = xNF3(1) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNF4(1) = xNF4(1) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNF5(1) = xNF5(1) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("valmerc") > 100 And de_informa.rsSel_AnOper.Fields("valmerc") <= 500 Then
            xNF1(2) = xNF1(2) + 1
            xNF3(2) = xNF3(2) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNF4(2) = xNF4(2) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNF5(2) = xNF5(2) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("valmerc") > 500 And de_informa.rsSel_AnOper.Fields("valmerc") <= 1000 Then
            xNF1(3) = xNF1(3) + 1
            xNF3(3) = xNF3(3) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNF4(3) = xNF4(3) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNF5(3) = xNF5(3) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("valmerc") > 1000 And de_informa.rsSel_AnOper.Fields("valmerc") <= 2500 Then
            xNF1(4) = xNF1(4) + 1
            xNF3(4) = xNF3(4) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNF4(4) = xNF4(4) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNF5(4) = xNF5(4) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("valmerc") > 2500 And de_informa.rsSel_AnOper.Fields("valmerc") <= 5000 Then
            xNF1(5) = xNF1(5) + 1
            xNF3(5) = xNF3(5) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNF4(5) = xNF4(5) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNF5(5) = xNF5(5) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("valmerc") > 5000 And de_informa.rsSel_AnOper.Fields("valmerc") <= 10000 Then
            xNF1(6) = xNF1(6) + 1
            xNF3(6) = xNF3(6) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNF4(6) = xNF4(6) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNF5(6) = xNF5(6) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("valmerc") > 10000 And de_informa.rsSel_AnOper.Fields("valmerc") <= 20000 Then
            xNF1(7) = xNF1(7) + 1
            xNF3(7) = xNF3(7) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNF4(7) = xNF4(7) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNF5(7) = xNF5(7) + de_informa.rsSel_AnOper.Fields("peso")
        ElseIf de_informa.rsSel_AnOper.Fields("valmerc") > 20000 Then
            xNF1(8) = xNF1(8) + 1
            xNF3(8) = xNF3(8) + de_informa.rsSel_AnOper.Fields("valmerc")
            xNF4(8) = xNF4(8) + de_informa.rsSel_AnOper.Fields("fretetotal")
            xNF5(8) = xNF5(8) + de_informa.rsSel_AnOper.Fields("peso")
        End If
        de_informa.rsSel_AnOper.MoveNext
    Loop
    
    'Fim dos cálculos do Recordset
    
    'Somar as variáveis das matrizes e atualizar a variável da matriz referente aos totais
    
    For xy = 1 To 16  'TOTAIS NORTE/NORDESTE
        xSDSU1(12) = xSDSU1(12) + xNONE1(xy)
        xSDSU2(12) = xSDSU2(12) + xNONE2(xy)
        xSDSU3(12) = xSDSU3(12) + xNONE3(xy)
        xSDSU4(12) = xSDSU4(12) + xNONE4(xy)
        xSDSU8(12) = xSDSU8(12) + xNONE8(xy)
    Next xy
    For xy = 1 To 11  'TOTAIS SUL/SUDESTE
        xSDSU1(12) = xSDSU1(12) + xSDSU1(xy)
        xSDSU2(12) = xSDSU2(12) + xSDSU2(xy)
        xSDSU3(12) = xSDSU3(12) + xSDSU3(xy)
        xSDSU4(12) = xSDSU4(12) + xSDSU4(xy)
        xSDSU8(12) = xSDSU8(12) + xSDSU8(xy)
    Next xy
    
    'ÍNDICES TOTAIS
    If xSDSU4(12) > 0 Then xSDSU5(12) = xSDSU3(12) / xSDSU4(12)
    If xSDSU3(12) > 0 Then xSDSU6(12) = xSDSU2(12) / xSDSU3(12)
    If xSDSU1(12) > 0 Then xSDSU7(12) = xSDSU2(12) / xSDSU1(12)
        
    'TOTAIS ESTATÍSTICA POR PESO
    For xy = 1 To 9
        xPeso1(10) = xPeso1(10) + xPeso1(xy)
        xPeso3(10) = xPeso3(10) + xPeso3(xy)
        xPeso4(10) = xPeso4(10) + xPeso4(xy)
        xPeso5(10) = xPeso5(10) + xPeso5(xy)
    Next xy
            
    'PERCENTUAIS CTCS ESTAT. PESO
    For xy = 1 To 10
        If xPeso1(10) > 10 Then xPeso2(xy) = xPeso1(xy) / xPeso1(10)
    Next xy
    
    'TOTAIS ESTATÍSTICA VAL. MERC.
    For xy = 1 To 8
        xNF1(9) = xNF1(9) + xNF1(xy)
        xNF3(9) = xNF3(9) + xNF3(xy)
        xNF4(9) = xNF4(9) + xNF4(xy)
        xNF5(9) = xNF5(9) + xNF5(xy)
    Next xy
            
    'PERCENTUAIS CTCS ESTAT. MERC.
    For xy = 1 To 9
        If xNF1(9) > 0 Then xNF2(xy) = xNF1(xy) / xNF1(9)
    Next xy
    
    'mostrar dados das variáveis das matrizes nos campos do form
    
    'NORTE / NORDESTE
    
    For xy = 1 To 16
        frmAnEstat.FlexValmerNONE.Row = xy
        frmAnEstat.FlexValmerNONE = Format(xNONE1(xy), "###,###,##0.00")
        frmAnEstat.FlexFreteNONE.Row = xy
        frmAnEstat.FlexFreteNONE = Format(xNONE2(xy), "#,###,##0.00")
        frmAnEstat.FlexPesoNONE.Row = xy
        frmAnEstat.FlexPesoNONE = Format(xNONE3(xy), "###,##0.0")
        frmAnEstat.FlexVolNONE.Row = xy
        frmAnEstat.FlexVolNONE = Format(xNONE8(xy), "###,##0")
        frmAnEstat.FlexExpedNONE.Row = xy
        frmAnEstat.FlexExpedNONE = Format(xNONE4(xy), "###,##0")
        frmAnEstat.FlexPercentNONE.Row = xy
        frmAnEstat.FlexPercentNONE.Col = 0
        frmAnEstat.FlexPercentNONE = Format(xNONE5(xy), "##,##0.0")
        frmAnEstat.FlexPercentNONE.Col = 1
        frmAnEstat.FlexPercentNONE = Format(xNONE6(xy), "##,##0.0")
        frmAnEstat.FlexPercentNONE.Col = 2
        frmAnEstat.FlexPercentNONE = Format(xNONE7(xy), "##0.000%")
    Next xy
    
    'SUDESTE/SUL/C.OESTE
    
    For xy = 1 To 11
        frmAnEstat.FlexValMerSDSU.Row = xy
        frmAnEstat.FlexValMerSDSU = Format(xSDSU1(xy), "###,###,##0.00")
        frmAnEstat.FlexFreteSDSU.Row = xy
        frmAnEstat.FlexFreteSDSU = Format(xSDSU2(xy), "#,###,##0.00")
        frmAnEstat.FlexPesoSDSU.Row = xy
        frmAnEstat.FlexPesoSDSU = Format(xSDSU3(xy), "###,##0.0")
        frmAnEstat.FlexVolSDSU.Row = xy
        frmAnEstat.FlexVolSDSU = Format(xSDSU8(xy), "###,##0")
        frmAnEstat.FlexExpedSDSU.Row = xy
        frmAnEstat.FlexExpedSDSU = Format(xSDSU4(xy), "###,##0")
        frmAnEstat.FlexPercentSDSU.Row = xy
        frmAnEstat.FlexPercentSDSU.Col = 0
        frmAnEstat.FlexPercentSDSU = Format(xSDSU5(xy), "##,##0.0")
        frmAnEstat.FlexPercentSDSU.Col = 1
        frmAnEstat.FlexPercentSDSU = Format(xSDSU6(xy), "##,##0.0")
        frmAnEstat.FlexPercentSDSU.Col = 2
        frmAnEstat.FlexPercentSDSU = Format(xSDSU7(xy), "##0.000%")
    Next xy
    
    'TOTAIS POR UF
    
        frmAnEstat.FlexValMerTot.Text = Format(xSDSU1(12), "###,###,##0.00")
        frmAnEstat.FlexFreteTot.Text = Format(xSDSU2(12), "#,###,##0.00")
        frmAnEstat.FlexPesoTot.Text = Format(xSDSU3(12), "###,##0.0")
        frmAnEstat.FlexVolTot.Text = Format(xSDSU8(12), "###,##0")
        frmAnEstat.FlexExpedTot.Text = Format(xSDSU4(12), "###,##0")
        frmAnEstat.FlexPercenttot.Col = 0
        frmAnEstat.FlexPercenttot.Text = Format(xSDSU5(12), "##,##0.0")
        frmAnEstat.FlexPercenttot.Col = 1
        frmAnEstat.FlexPercenttot.Text = Format(xSDSU6(12), "##,##0.0")
        frmAnEstat.FlexPercenttot.Col = 2
        frmAnEstat.FlexPercenttot.Text = Format(xSDSU7(12), "##0.000%")
    
    'ESTAT. POR PESO
    
    'QTD. DE CTCs

        frmAnEstat.lblPeso1A = Format(xPeso1(1), "###,##0")
        frmAnEstat.lblPeso1B = Format(xPeso1(2), "###,##0")
        frmAnEstat.lblPeso1C = Format(xPeso1(3), "###,##0")
        frmAnEstat.lblPeso1D = Format(xPeso1(4), "###,##0")
        frmAnEstat.lblPeso1E = Format(xPeso1(5), "###,##0")
        frmAnEstat.lblPeso1F = Format(xPeso1(6), "###,##0")
        frmAnEstat.lblPeso1G = Format(xPeso1(7), "###,##0")
        frmAnEstat.lblPeso1H = Format(xPeso1(8), "###,##0")
        frmAnEstat.lblPeso1I = Format(xPeso1(9), "###,##0")
        frmAnEstat.lblPeso1Total = Format(xPeso1(10), "###,##0")
    'QTD CTC. PERCENT %
        frmAnEstat.lblPer1A = Format(xPeso2(1), "##0.0%")
        frmAnEstat.lblper1B = Format(xPeso2(2), "##0.0%")
        frmAnEstat.lblper1C = Format(xPeso2(3), "##0.0%")
        frmAnEstat.lblper1D = Format(xPeso2(4), "##0.0%")
        frmAnEstat.lblper1E = Format(xPeso2(5), "##0.0%")
        frmAnEstat.lblper1F = Format(xPeso2(6), "##0.0%")
        frmAnEstat.lblper1G = Format(xPeso2(7), "##0.0%")
        frmAnEstat.lblper1H = Format(xPeso2(8), "##0.0%")
        frmAnEstat.lblper1I = Format(xPeso2(9), "##0.0%")
        frmAnEstat.lblPercTotal = Format(xPeso2(10), "##0.0%")
    'VALOR DE MERCADORIA
        frmAnEstat.lblPeso2A = Format(xPeso3(1), "###,###,##0.00")
        frmAnEstat.lblPeso2B = Format(xPeso3(2), "###,###,##0.00")
        frmAnEstat.lblPeso2C = Format(xPeso3(3), "###,###,##0.00")
        frmAnEstat.lblPeso2D = Format(xPeso3(4), "###,###,##0.00")
        frmAnEstat.lblPeso2E = Format(xPeso3(5), "###,###,##0.00")
        frmAnEstat.lblPeso2F = Format(xPeso3(6), "###,###,##0.00")
        frmAnEstat.lblPeso2G = Format(xPeso3(7), "###,###,##0.00")
        frmAnEstat.lblPeso2H = Format(xPeso3(8), "###,###,##0.00")
        frmAnEstat.lblPeso2I = Format(xPeso3(9), "###,###,##0.00")
        frmAnEstat.lblPeso2Total = Format(xPeso3(10), "###,###,##0.00")
    'VALOR DE FRETE
        frmAnEstat.lblPeso3A = Format(xPeso4(1), "#,###,##0.00")
        frmAnEstat.lblPeso3B = Format(xPeso4(2), "#,###,##0.00")
        frmAnEstat.lblPeso3C = Format(xPeso4(3), "#,###,##0.00")
        frmAnEstat.lblPeso3D = Format(xPeso4(4), "#,###,##0.00")
        frmAnEstat.lblPeso3E = Format(xPeso4(5), "#,###,##0.00")
        frmAnEstat.lblPeso3F = Format(xPeso4(6), "#,###,##0.00")
        frmAnEstat.lblPeso3G = Format(xPeso4(7), "#,###,##0.00")
        frmAnEstat.lblPeso3H = Format(xPeso4(8), "#,###,##0.00")
        frmAnEstat.lblPeso3I = Format(xPeso4(9), "#,###,##0.00")
        frmAnEstat.lblPeso3Total = Format(xPeso4(10), "#,###,##0.00")
    'PESO
        frmAnEstat.lblPeso4A = Format(xPeso5(1), "###,##0.0")
        frmAnEstat.lblPeso4B = Format(xPeso5(2), "###,##0.0")
        frmAnEstat.lblPeso4C = Format(xPeso5(3), "###,##0.0")
        frmAnEstat.lblPeso4D = Format(xPeso5(4), "###,##0.0")
        frmAnEstat.lblPeso4E = Format(xPeso5(5), "###,##0.0")
        frmAnEstat.lblPeso4F = Format(xPeso5(6), "###,##0.0")
        frmAnEstat.lblPeso4G = Format(xPeso5(7), "###,##0.0")
        frmAnEstat.lblPeso4H = Format(xPeso5(8), "###,##0.0")
        frmAnEstat.lblPeso4I = Format(xPeso5(9), "###,##0.0")
        frmAnEstat.lblPeso4Total = Format(xPeso5(10), "###,##0.0")
        
        
    'ESTAT. POR VALOR MERC.
    
    'QTD. DE CTCs

        frmAnEstat.lblNF1A = Format(xNF1(1), "###,##0")
        frmAnEstat.lblNF1B = Format(xNF1(2), "###,##0")
        frmAnEstat.lblNF1C = Format(xNF1(3), "###,##0")
        frmAnEstat.lblNF1D = Format(xNF1(4), "###,##0")
        frmAnEstat.lblNF1E = Format(xNF1(5), "###,##0")
        frmAnEstat.lblNF1F = Format(xNF1(6), "###,##0")
        frmAnEstat.lblNF1G = Format(xNF1(7), "###,##0")
        frmAnEstat.lblNF1H = Format(xNF1(8), "###,##0")
        frmAnEstat.lblNF1Total = Format(xNF1(9), "###,##0")
    'QTD CTC. PERCENT %
        
        frmAnEstat.lblNFPerA = Format(xNF2(1), "##0.0%")
        frmAnEstat.lblNFPerB = Format(xNF2(2), "##0.0%")
        frmAnEstat.lblNFPerC = Format(xNF2(3), "##0.0%")
        frmAnEstat.lblNFPerD = Format(xNF2(4), "##0.0%")
        frmAnEstat.lblNFPerE = Format(xNF2(5), "##0.0%")
        frmAnEstat.lblNFPerF = Format(xNF2(6), "##0.0%")
        frmAnEstat.lblNFPerG = Format(xNF2(7), "##0.0%")
        frmAnEstat.lblNFPerH = Format(xNF2(8), "##0.0%")
        frmAnEstat.lblNFPerTotal = Format(xNF2(9), "##0.0%")
    'VALOR DE MERCADORIA
        frmAnEstat.lblNF2A = Format(xNF3(1), "###,###,##0.00")
        frmAnEstat.lblNF2B = Format(xNF3(2), "###,###,##0.00")
        frmAnEstat.lblNF2C = Format(xNF3(3), "###,###,##0.00")
        frmAnEstat.lblNF2D = Format(xNF3(4), "###,###,##0.00")
        frmAnEstat.lblNF2E = Format(xNF3(5), "###,###,##0.00")
        frmAnEstat.lblNF2F = Format(xNF3(6), "###,###,##0.00")
        frmAnEstat.lblNF2G = Format(xNF3(7), "###,###,##0.00")
        frmAnEstat.lblNF2H = Format(xNF3(8), "###,###,##0.00")
        frmAnEstat.lblNF2Total = Format(xNF3(9), "###,###,##0.00")
    'VALOR DE FRETE
        frmAnEstat.lblNF3A = Format(xNF4(1), "#,###,##0.00")
        frmAnEstat.lblNF3B = Format(xNF4(2), "#,###,##0.00")
        frmAnEstat.lblNF3C = Format(xNF4(3), "#,###,##0.00")
        frmAnEstat.lblNF3D = Format(xNF4(4), "#,###,##0.00")
        frmAnEstat.lblNF3E = Format(xNF4(5), "#,###,##0.00")
        frmAnEstat.lblNF3F = Format(xNF4(6), "#,###,##0.00")
        frmAnEstat.lblNF3G = Format(xNF4(7), "#,###,##0.00")
        frmAnEstat.lblNF3H = Format(xNF4(8), "#,###,##0.00")
        frmAnEstat.lblNF3Total = Format(xNF4(9), "#,###,##0.00")
    'PESO
        frmAnEstat.lblNF4A = Format(xNF5(1), "###,##0.0")
        frmAnEstat.lblNF4B = Format(xNF5(2), "###,##0.0")
        frmAnEstat.lblNF4C = Format(xNF5(3), "###,##0.0")
        frmAnEstat.lblNF4D = Format(xNF5(4), "###,##0.0")
        frmAnEstat.lblNF4E = Format(xNF5(5), "###,##0.0")
        frmAnEstat.lblNF4F = Format(xNF5(6), "###,##0.0")
        frmAnEstat.lblNF4G = Format(xNF5(7), "###,##0.0")
        frmAnEstat.lblNF4H = Format(xNF5(8), "###,##0.0")
        frmAnEstat.lblNF4Total = Format(xNF5(9), "###,##0.0")
End Sub
Private Sub gridprazos()
Dim xcgc1 As String, xy As Long, xnoprazo As Long, xforaprz As Long, xmodal As String, xabonodias As Integer

'DADOS DOS PRAZOS NAS CAPITAIS
    
    'PREENCHENDO AS COLUNAS DE PRAZOS
    
    If chkTodosCli.Value = 0 Then  'escolha de um CGC
        xcgc1 = Trim(txtCGCCli.Text) & "%"
    ElseIf chkTodosCli.Value = 1 Then  'todos os clientes
        xcgc1 = "%"
    End If
    
'    If chkTodosCli.Value = 1 Then
'        frmAnEntregas.lblPrazo = "TAB000"
'    Else
'        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
'        de_informa.Sel_ConsCadCli xcgc1
'        If de_informa.rsSel_ConsCadCli.RecordCount = 0 Then
'           MsgBox "CGC do cliente não encontrado. Erro de Consistência. Chame Suporte Técnico"
'           Exit Sub
'        Else
'           frmAnEntregas.lblPrazo = de_informa.rsSel_ConsCadCli.Fields("prazo")
'        End If
'    End If
'    If de_informa.rsSel_CadPrazo.State = 1 Then de_informa.rsSel_CadPrazo.Close
'    If chkModal.Value = 1 Then
        'SE FOR TODOS MODAIS BUSCAR PRAZO AÉREO
'        de_informa.Sel_CadPrazo frmAnEntregas.lblPrazo, "A"
'    Else
        'caso contrário busca prazo conforme modal
        If optAir.Value = True Then
            xmodal = "AEREO"
        Else
            xmodal = "RODOVIARIO"
        End If
'        de_informa.Sel_CadPrazo frmAnEntregas.lblPrazo, Mid(xmodal, 1, 1)
'    End If
    
'    If de_informa.rsSel_CadPrazo.RecordCount = 0 Then
'       MsgBox "Tabela de Prazos do Cliente Não Encontrada. Procure Suporte Técnico"
'       Exit Sub
'    End If
'    de_informa.rsSel_CadPrazo.MoveFirst
'    frmAnEntregas.FlexCapitais1.Col = 0
'    frmAnEntregas.FlexCapitais2.Col = 0
'    frmAnEntregas.FlexInterior1.Col = 0
'    frmAnEntregas.FlexInterior2.Col = 0
'    Do Until de_informa.rsSel_CadPrazo.EOF
'       If de_informa.rsSel_CadPrazo.Fields("uf") = "AC" Then
'            frmAnEntregas.FlexCapitais1.Row = 1
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 1
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AM" Then
'            frmAnEntregas.FlexCapitais1.Row = 2
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 2
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AP" Then
'            frmAnEntregas.FlexCapitais1.Row = 3
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 3
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PA" Then
'            frmAnEntregas.FlexCapitais1.Row = 4
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 4
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RO" Then
'            frmAnEntregas.FlexCapitais1.Row = 5
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 5
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RR" Then
'            frmAnEntregas.FlexCapitais1.Row = 6
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 6
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "TO" Then
'            frmAnEntregas.FlexCapitais1.Row = 7
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 7
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "AL" Then
'            frmAnEntregas.FlexCapitais1.Row = 8
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 8
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "BA" Then
'            frmAnEntregas.FlexCapitais1.Row = 9
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 9
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SE" Then
'            frmAnEntregas.FlexCapitais1.Row = 10
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 10
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PE" Then
'            frmAnEntregas.FlexCapitais1.Row = 11
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 11
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PB" Then
'            frmAnEntregas.FlexCapitais1.Row = 12
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 12
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RN" Then
'            frmAnEntregas.FlexCapitais1.Row = 13
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 13
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "CE" Then
'            frmAnEntregas.FlexCapitais1.Row = 14
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 14
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PI" Then
'            frmAnEntregas.FlexCapitais1.Row = 15
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 15
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MA" Then
'            frmAnEntregas.FlexCapitais1.Row = 16
'            frmAnEntregas.FlexCapitais1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior1.Row = 16
'            frmAnEntregas.FlexInterior1.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "ES" Then
'            frmAnEntregas.FlexCapitais2.Row = 1
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 1
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MG" Then
'            frmAnEntregas.FlexCapitais2.Row = 2
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 2
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RJ" Then
'            frmAnEntregas.FlexCapitais2.Row = 3
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 3
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SP" Then
'            frmAnEntregas.FlexCapitais2.Row = 4
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 4
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "PR" Then
'            frmAnEntregas.FlexCapitais2.Row = 5
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 5
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "RS" Then
'            frmAnEntregas.FlexCapitais2.Row = 6
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 6
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "SC" Then
'            frmAnEntregas.FlexCapitais2.Row = 7
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 7
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "DF" Then
'            frmAnEntregas.FlexCapitais2.Row = 8
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 8
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "GO" Then
'            frmAnEntregas.FlexCapitais2.Row = 9
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 9
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MS" Then
'            frmAnEntregas.FlexCapitais2.Row = 10
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 10
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       ElseIf de_informa.rsSel_CadPrazo.Fields("uf") = "MT" Then
'            frmAnEntregas.FlexCapitais2.Row = 11
'            frmAnEntregas.FlexCapitais2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_cap")
'            frmAnEntregas.FlexInterior2.Row = 11
'            frmAnEntregas.FlexInterior2.Text = de_informa.rsSel_CadPrazo.Fields("prazo_int")
'       End If
'       de_informa.rsSel_CadPrazo.MoveNext
'       DoEvents
'    Loop
    
'TRAZ QUANTIDADES DE CTC NO PRAZO E FORA DO PRAZO

    If de_informa.rsSel_EntrPrazo.State = 1 Then de_informa.rsSel_EntrPrazo.Close
    de_informa.Sel_EntrPrazo CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc1, xmodal
    If de_informa.rsSel_EntrPrazo.RecordCount = 0 Then
        MsgBox "Não há dados há serem Exibidos. Procure Ajuda Junto ao Gestor do Sistema"
        Exit Sub
    End If
    de_informa.rsSel_EntrPrazo.MoveFirst
    Do Until de_informa.rsSel_EntrPrazo.EOF
        If de_informa.rsSel_UfCidade.State = 1 Then de_informa.rsSel_UfCidade.Close
        de_informa.Sel_UfCidade de_informa.rsSel_EntrPrazo.Fields("uf_dest"), de_informa.rsSel_EntrPrazo.Fields("cidade_dest")
        If de_informa.rsSel_UfCidade.RecordCount < 1 Then  'se for interior
            If chkAnalise.Value = 1 Then
                xabonodias = de_informa.rsSel_EntrPrazo.Fields("abonodias")
            Else
                xabonodias = 0
            End If
            If ((de_informa.rsSel_EntrPrazo.Fields("diasuteis") - xabonodias) - de_informa.rsSel_EntrPrazo.Fields("prazoentr")) > 0 Then  'fora do prazo
                frmAnEntregas.FlexInterior1.Col = 2
                frmAnEntregas.FlexInterior2.Col = 2
                If de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AC" Then
                    frmAnEntregas.FlexInterior1.Row = 1
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AM" Then
                    frmAnEntregas.FlexInterior1.Row = 2
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AP" Then
                    frmAnEntregas.FlexInterior1.Row = 3
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PA" Then
                    frmAnEntregas.FlexInterior1.Row = 4
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RO" Then
                    frmAnEntregas.FlexInterior1.Row = 5
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RR" Then
                    frmAnEntregas.FlexInterior1.Row = 6
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "TO" Then
                    frmAnEntregas.FlexInterior1.Row = 7
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AL" Then
                    frmAnEntregas.FlexInterior1.Row = 8
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "BA" Then
                    frmAnEntregas.FlexInterior1.Row = 9
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SE" Then
                    frmAnEntregas.FlexInterior1.Row = 10
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PE" Then
                    frmAnEntregas.FlexInterior1.Row = 11
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PB" Then
                    frmAnEntregas.FlexInterior1.Row = 12
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RN" Then
                    frmAnEntregas.FlexInterior1.Row = 13
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "CE" Then
                    frmAnEntregas.FlexInterior1.Row = 14
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PI" Then
                    frmAnEntregas.FlexInterior1.Row = 15
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MA" Then
                    frmAnEntregas.FlexInterior1.Row = 16
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "ES" Then
                    frmAnEntregas.FlexInterior2.Row = 1
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MG" Then
                    frmAnEntregas.FlexInterior2.Row = 2
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RJ" Then
                    frmAnEntregas.FlexInterior2.Row = 3
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SP" Then
                    frmAnEntregas.FlexInterior2.Row = 4
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PR" Then
                    frmAnEntregas.FlexInterior2.Row = 5
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RS" Then
                    frmAnEntregas.FlexInterior2.Row = 6
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SC" Then
                    frmAnEntregas.FlexInterior2.Row = 7
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "DF" Then
                    frmAnEntregas.FlexInterior2.Row = 8
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "GO" Then
                    frmAnEntregas.FlexInterior2.Row = 9
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MS" Then
                    frmAnEntregas.FlexInterior2.Row = 10
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MT" Then
                    frmAnEntregas.FlexInterior2.Row = 11
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                End If
            Else  'se for no prazo
                frmAnEntregas.FlexInterior1.Col = 1
                frmAnEntregas.FlexInterior2.Col = 1
                If de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AC" Then
                    frmAnEntregas.FlexInterior1.Row = 1
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AM" Then
                    frmAnEntregas.FlexInterior1.Row = 2
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AP" Then
                    frmAnEntregas.FlexInterior1.Row = 3
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PA" Then
                    frmAnEntregas.FlexInterior1.Row = 4
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RO" Then
                    frmAnEntregas.FlexInterior1.Row = 5
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RR" Then
                    frmAnEntregas.FlexInterior1.Row = 6
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "TO" Then
                    frmAnEntregas.FlexInterior1.Row = 7
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AL" Then
                    frmAnEntregas.FlexInterior1.Row = 8
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "BA" Then
                    frmAnEntregas.FlexInterior1.Row = 9
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SE" Then
                    frmAnEntregas.FlexInterior1.Row = 10
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PE" Then
                    frmAnEntregas.FlexInterior1.Row = 11
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PB" Then
                    frmAnEntregas.FlexInterior1.Row = 12
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RN" Then
                    frmAnEntregas.FlexInterior1.Row = 13
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "CE" Then
                    frmAnEntregas.FlexInterior1.Row = 14
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PI" Then
                    frmAnEntregas.FlexInterior1.Row = 15
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MA" Then
                    frmAnEntregas.FlexInterior1.Row = 16
                    frmAnEntregas.FlexInterior1.Text = CDbl(frmAnEntregas.FlexInterior1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "ES" Then
                    frmAnEntregas.FlexInterior2.Row = 1
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MG" Then
                    frmAnEntregas.FlexInterior2.Row = 2
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RJ" Then
                    frmAnEntregas.FlexInterior2.Row = 3
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SP" Then
                    frmAnEntregas.FlexInterior2.Row = 4
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PR" Then
                    frmAnEntregas.FlexInterior2.Row = 5
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RS" Then
                    frmAnEntregas.FlexInterior2.Row = 6
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SC" Then
                    frmAnEntregas.FlexInterior2.Row = 7
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "DF" Then
                    frmAnEntregas.FlexInterior2.Row = 8
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "GO" Then
                    frmAnEntregas.FlexInterior2.Row = 9
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MS" Then
                    frmAnEntregas.FlexInterior2.Row = 10
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MT" Then
                    frmAnEntregas.FlexInterior2.Row = 11
                    frmAnEntregas.FlexInterior2.Text = CDbl(frmAnEntregas.FlexInterior2.Text) + 1
                End If
            End If
        Else 'se for capital
            If chkAnalise.Value = 1 Then
                xabonodias = de_informa.rsSel_EntrPrazo.Fields("abonodias")
            Else
                xabonodias = 0
            End If
            If ((de_informa.rsSel_EntrPrazo.Fields("diasuteis") - xabonodias) - de_informa.rsSel_EntrPrazo.Fields("prazoentr")) > 0 Then  'fora do prazo
                frmAnEntregas.FlexCapitais1.Col = 2
                frmAnEntregas.FlexCapitais2.Col = 2
                If de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AC" Then
                    frmAnEntregas.FlexCapitais1.Row = 1
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AM" Then
                    frmAnEntregas.FlexCapitais1.Row = 2
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AP" Then
                    frmAnEntregas.FlexCapitais1.Row = 3
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PA" Then
                    frmAnEntregas.FlexCapitais1.Row = 4
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RO" Then
                    frmAnEntregas.FlexCapitais1.Row = 5
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RR" Then
                    frmAnEntregas.FlexCapitais1.Row = 6
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "TO" Then
                    frmAnEntregas.FlexCapitais1.Row = 7
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AL" Then
                    frmAnEntregas.FlexCapitais1.Row = 8
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "BA" Then
                    frmAnEntregas.FlexCapitais1.Row = 9
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SE" Then
                    frmAnEntregas.FlexCapitais1.Row = 10
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PE" Then
                    frmAnEntregas.FlexCapitais1.Row = 11
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PB" Then
                    frmAnEntregas.FlexCapitais1.Row = 12
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RN" Then
                    frmAnEntregas.FlexCapitais1.Row = 13
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "CE" Then
                    frmAnEntregas.FlexCapitais1.Row = 14
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PI" Then
                    frmAnEntregas.FlexCapitais1.Row = 15
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MA" Then
                    frmAnEntregas.FlexCapitais1.Row = 16
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "ES" Then
                    frmAnEntregas.FlexCapitais2.Row = 1
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MG" Then
                    frmAnEntregas.FlexCapitais2.Row = 2
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RJ" Then
                    frmAnEntregas.FlexCapitais2.Row = 3
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SP" Then
                    frmAnEntregas.FlexCapitais2.Row = 4
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PR" Then
                    frmAnEntregas.FlexCapitais2.Row = 5
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RS" Then
                    frmAnEntregas.FlexCapitais2.Row = 6
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SC" Then
                    frmAnEntregas.FlexCapitais2.Row = 7
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "DF" Then
                    frmAnEntregas.FlexCapitais2.Row = 8
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "GO" Then
                    frmAnEntregas.FlexCapitais2.Row = 9
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MS" Then
                    frmAnEntregas.FlexCapitais2.Row = 10
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MT" Then
                    frmAnEntregas.FlexCapitais2.Row = 11
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                End If
            Else  'se for no prazo
                frmAnEntregas.FlexCapitais1.Col = 1
                frmAnEntregas.FlexCapitais2.Col = 1
                If de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AC" Then
                    frmAnEntregas.FlexCapitais1.Row = 1
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AM" Then
                    frmAnEntregas.FlexCapitais1.Row = 2
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AP" Then
                    frmAnEntregas.FlexCapitais1.Row = 3
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PA" Then
                    frmAnEntregas.FlexCapitais1.Row = 4
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RO" Then
                    frmAnEntregas.FlexCapitais1.Row = 5
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RR" Then
                    frmAnEntregas.FlexCapitais1.Row = 6
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "TO" Then
                    frmAnEntregas.FlexCapitais1.Row = 7
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "AL" Then
                    frmAnEntregas.FlexCapitais1.Row = 8
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "BA" Then
                    frmAnEntregas.FlexCapitais1.Row = 9
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SE" Then
                    frmAnEntregas.FlexCapitais1.Row = 10
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PE" Then
                    frmAnEntregas.FlexCapitais1.Row = 11
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PB" Then
                    frmAnEntregas.FlexCapitais1.Row = 12
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RN" Then
                    frmAnEntregas.FlexCapitais1.Row = 13
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "CE" Then
                    frmAnEntregas.FlexCapitais1.Row = 14
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PI" Then
                    frmAnEntregas.FlexCapitais1.Row = 15
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MA" Then
                    frmAnEntregas.FlexCapitais1.Row = 16
                    frmAnEntregas.FlexCapitais1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "ES" Then
                    frmAnEntregas.FlexCapitais2.Row = 1
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MG" Then
                    frmAnEntregas.FlexCapitais2.Row = 2
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RJ" Then
                    frmAnEntregas.FlexCapitais2.Row = 3
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SP" Then
                    frmAnEntregas.FlexCapitais2.Row = 4
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "PR" Then
                    frmAnEntregas.FlexCapitais2.Row = 5
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "RS" Then
                    frmAnEntregas.FlexCapitais2.Row = 6
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "SC" Then
                    frmAnEntregas.FlexCapitais2.Row = 7
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "DF" Then
                    frmAnEntregas.FlexCapitais2.Row = 8
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "GO" Then
                    frmAnEntregas.FlexCapitais2.Row = 9
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MS" Then
                    frmAnEntregas.FlexCapitais2.Row = 10
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                ElseIf de_informa.rsSel_EntrPrazo.Fields("uf_dest") = "MT" Then
                    frmAnEntregas.FlexCapitais2.Row = 11
                    frmAnEntregas.FlexCapitais2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + 1
                End If
            End If
        End If
        de_informa.rsSel_EntrPrazo.MoveNext
        DoEvents
    Loop
        
'ATUALIZA O FLEXGRID DE TOTAIS.

'total On Time no Prazo
    
    frmAnEntregas.FlexTotUf1.Col = 0
    frmAnEntregas.FlexTotUf2.Col = 0
    frmAnEntregas.FlexCapitais1.Col = 1
    frmAnEntregas.FlexCapitais2.Col = 1
    frmAnEntregas.FlexInterior1.Col = 1
    frmAnEntregas.FlexInterior2.Col = 1
    
    For xy = 1 To 16
        frmAnEntregas.FlexTotUf1.Row = xy
        frmAnEntregas.FlexCapitais1.Row = xy
        frmAnEntregas.FlexInterior1.Row = xy
        frmAnEntregas.FlexTotUf1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + CDbl(frmAnEntregas.FlexInterior1.Text)
    Next xy
    For xy = 1 To 11
        frmAnEntregas.FlexTotUf2.Row = xy
        frmAnEntregas.FlexCapitais2.Row = xy
        frmAnEntregas.FlexInterior2.Row = xy
        frmAnEntregas.FlexTotUf2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + CDbl(frmAnEntregas.FlexInterior2.Text)
    Next xy

'total Delay Fora do Prazo
    
    frmAnEntregas.FlexTotUf1.Col = 1
    frmAnEntregas.FlexTotUf2.Col = 1
    frmAnEntregas.FlexCapitais1.Col = 2
    frmAnEntregas.FlexCapitais2.Col = 2
    frmAnEntregas.FlexInterior1.Col = 2
    frmAnEntregas.FlexInterior2.Col = 2
   
    For xy = 1 To 16
        frmAnEntregas.FlexTotUf1.Row = xy
        frmAnEntregas.FlexCapitais1.Row = xy
        frmAnEntregas.FlexInterior1.Row = xy
        frmAnEntregas.FlexTotUf1.Text = CDbl(frmAnEntregas.FlexCapitais1.Text) + CDbl(frmAnEntregas.FlexInterior1.Text)
    Next xy
    For xy = 1 To 11
        frmAnEntregas.FlexTotUf2.Row = xy
        frmAnEntregas.FlexCapitais2.Row = xy
        frmAnEntregas.FlexInterior2.Row = xy
        frmAnEntregas.FlexTotUf2.Text = CDbl(frmAnEntregas.FlexCapitais2.Text) + CDbl(frmAnEntregas.FlexInterior2.Text)
    Next xy
    
'Totais FlexCapitais , FlexInterior, FlexTotalGeral
    
'total capitais: No Prazo
    
    frmAnEntregas.FlexCapitais1.Col = 1
    frmAnEntregas.FlexTotCapitais.Col = 1
    For xy = 1 To 16
        frmAnEntregas.FlexCapitais1.Row = xy
        frmAnEntregas.FlexTotCapitais.Text = CDbl(frmAnEntregas.FlexTotCapitais.Text) + CDbl(frmAnEntregas.FlexCapitais1.Text)
    Next xy
    frmAnEntregas.FlexCapitais2.Col = 1
    For xy = 1 To 11
        frmAnEntregas.FlexCapitais2.Row = xy
        frmAnEntregas.FlexTotCapitais.Text = CDbl(frmAnEntregas.FlexTotCapitais.Text) + CDbl(frmAnEntregas.FlexCapitais2.Text)
    Next xy
    
'total interior: No Prazo
    
    frmAnEntregas.FlexInterior1.Col = 1
    frmAnEntregas.FlexTotInterior.Col = 1
    For xy = 1 To 16
        frmAnEntregas.FlexInterior1.Row = xy
        frmAnEntregas.FlexTotInterior.Text = CDbl(frmAnEntregas.FlexTotInterior.Text) + CDbl(frmAnEntregas.FlexInterior1.Text)
    Next xy
    frmAnEntregas.FlexInterior2.Col = 1
    For xy = 1 To 11
        frmAnEntregas.FlexInterior2.Row = xy
        frmAnEntregas.FlexTotInterior.Text = CDbl(frmAnEntregas.FlexTotInterior.Text) + CDbl(frmAnEntregas.FlexInterior2.Text)
    Next xy
    
'total Geral UF: No Prazo
    
    frmAnEntregas.FlexTotUf1.Col = 0
    frmAnEntregas.FlexTotTotuf.Col = 0
    For xy = 1 To 16
        frmAnEntregas.FlexTotUf1.Row = xy
        frmAnEntregas.FlexTotTotuf.Text = CDbl(frmAnEntregas.FlexTotTotuf.Text) + CDbl(frmAnEntregas.FlexTotUf1.Text)
    Next xy
    frmAnEntregas.FlexTotUf2.Col = 0
    For xy = 1 To 11
        frmAnEntregas.FlexTotUf2.Row = xy
        frmAnEntregas.FlexTotTotuf.Text = CDbl(frmAnEntregas.FlexTotTotuf.Text) + CDbl(frmAnEntregas.FlexTotUf2.Text)
    Next xy
    
'total capitais: Fora Prazo
    
    frmAnEntregas.FlexCapitais1.Col = 2
    frmAnEntregas.FlexTotCapitais.Col = 2
    For xy = 1 To 16
        frmAnEntregas.FlexCapitais1.Row = xy
        frmAnEntregas.FlexTotCapitais.Text = CDbl(frmAnEntregas.FlexTotCapitais.Text) + CDbl(frmAnEntregas.FlexCapitais1.Text)
    Next xy
    frmAnEntregas.FlexCapitais2.Col = 2
    For xy = 1 To 11
        frmAnEntregas.FlexCapitais2.Row = xy
        frmAnEntregas.FlexTotCapitais.Text = CDbl(frmAnEntregas.FlexTotCapitais.Text) + CDbl(frmAnEntregas.FlexCapitais2.Text)
    Next xy
    
'total Interior: Fora Prazo
    
    frmAnEntregas.FlexInterior1.Col = 2
    frmAnEntregas.FlexTotInterior.Col = 2
    For xy = 1 To 16
        frmAnEntregas.FlexInterior1.Row = xy
        frmAnEntregas.FlexTotInterior.Text = CDbl(frmAnEntregas.FlexTotInterior.Text) + CDbl(frmAnEntregas.FlexInterior1.Text)
    Next xy
    frmAnEntregas.FlexInterior2.Col = 2
    For xy = 1 To 11
        frmAnEntregas.FlexInterior2.Row = xy
        frmAnEntregas.FlexTotInterior.Text = CDbl(frmAnEntregas.FlexTotInterior.Text) + CDbl(frmAnEntregas.FlexInterior2.Text)
    Next xy
    
'total Geral UF: No Prazo
    
    frmAnEntregas.FlexTotUf1.Col = 1
    frmAnEntregas.FlexTotTotuf.Col = 1
    For xy = 1 To 16
        frmAnEntregas.FlexTotUf1.Row = xy
        frmAnEntregas.FlexTotTotuf.Text = CDbl(frmAnEntregas.FlexTotTotuf.Text) + CDbl(frmAnEntregas.FlexTotUf1.Text)
    Next xy
    frmAnEntregas.FlexTotUf2.Col = 1
    For xy = 1 To 11
        frmAnEntregas.FlexTotUf2.Row = xy
        frmAnEntregas.FlexTotTotuf.Text = CDbl(frmAnEntregas.FlexTotTotuf.Text) + CDbl(frmAnEntregas.FlexTotUf2.Text)
    Next xy
    
'Cálculos dos Percentuais dos Prazos
        
'Capitais - tab1
    For xy = 1 To 16
        frmAnEntregas.FlexCapitais1.Row = xy
        frmAnEntregas.FlexCapitais1.Col = 1
        xnoprazo = CDbl(frmAnEntregas.FlexCapitais1.Text)
        frmAnEntregas.FlexCapitais1.Col = 2
        xforaprz = CDbl(frmAnEntregas.FlexCapitais1.Text)
        frmAnEntregas.FlexCapitais1.Col = 3
        If (xforaprz + xnoprazo) > 0 Then
            frmAnEntregas.FlexCapitais1.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
        End If
    Next xy

'Interior - tab1
    For xy = 1 To 16
        frmAnEntregas.FlexInterior1.Row = xy
        frmAnEntregas.FlexInterior1.Col = 1
        xnoprazo = CDbl(frmAnEntregas.FlexInterior1.Text)
        frmAnEntregas.FlexInterior1.Col = 2
        xforaprz = CDbl(frmAnEntregas.FlexInterior1.Text)
        frmAnEntregas.FlexInterior1.Col = 3
        If (xforaprz + xnoprazo) > 0 Then
            frmAnEntregas.FlexInterior1.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
        End If
    Next xy

'Total UF - tab1
    For xy = 1 To 16
        frmAnEntregas.FlexTotUf1.Row = xy
        frmAnEntregas.FlexTotUf1.Col = 0
        xnoprazo = CDbl(frmAnEntregas.FlexTotUf1.Text)
        frmAnEntregas.FlexTotUf1.Col = 1
        xforaprz = CDbl(frmAnEntregas.FlexTotUf1.Text)
        frmAnEntregas.FlexTotUf1.Col = 2
        If (xforaprz + xnoprazo) > 0 Then
            frmAnEntregas.FlexTotUf1.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
        End If
    Next xy
    
'Capitais - tab2
    For xy = 1 To 11
        frmAnEntregas.FlexCapitais2.Row = xy
        frmAnEntregas.FlexCapitais2.Col = 1
        xnoprazo = CDbl(frmAnEntregas.FlexCapitais2.Text)
        frmAnEntregas.FlexCapitais2.Col = 2
        xforaprz = CDbl(frmAnEntregas.FlexCapitais2.Text)
        frmAnEntregas.FlexCapitais2.Col = 3
        If (xforaprz + xnoprazo) > 0 Then
            frmAnEntregas.FlexCapitais2.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
        End If
    Next xy
    
'Interior - tab2
    For xy = 1 To 11
        frmAnEntregas.FlexInterior2.Row = xy
        frmAnEntregas.FlexInterior2.Col = 1
        xnoprazo = CDbl(frmAnEntregas.FlexInterior2.Text)
        frmAnEntregas.FlexInterior2.Col = 2
        xforaprz = CDbl(frmAnEntregas.FlexInterior2.Text)
        frmAnEntregas.FlexInterior2.Col = 3
        If (xforaprz + xnoprazo) > 0 Then
            frmAnEntregas.FlexInterior2.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
        End If
    Next xy
    
'Total UF - tab2
    For xy = 1 To 11
        frmAnEntregas.FlexTotUf2.Row = xy
        frmAnEntregas.FlexTotUf2.Col = 0
        xnoprazo = CDbl(frmAnEntregas.FlexTotUf2.Text)
        frmAnEntregas.FlexTotUf2.Col = 1
        xforaprz = CDbl(frmAnEntregas.FlexTotUf2.Text)
        frmAnEntregas.FlexTotUf2.Col = 2
        If (xforaprz + xnoprazo) > 0 Then
            frmAnEntregas.FlexTotUf2.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
        End If
    Next xy
    
'percentual dos totais Capitais, Interior e Total UF
    
    frmAnEntregas.FlexTotCapitais.Col = 1
    xnoprazo = CDbl(frmAnEntregas.FlexTotCapitais.Text)
    frmAnEntregas.FlexTotCapitais.Col = 2
    xforaprz = CDbl(frmAnEntregas.FlexTotCapitais.Text)
    frmAnEntregas.FlexTotCapitais.Col = 3
    If (xforaprz + xnoprazo) > 0 Then
        frmAnEntregas.FlexTotCapitais.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
    End If
    frmAnEntregas.FlexTotInterior.Col = 1
    xnoprazo = CDbl(frmAnEntregas.FlexTotInterior.Text)
    frmAnEntregas.FlexTotInterior.Col = 2
    xforaprz = CDbl(frmAnEntregas.FlexTotInterior.Text)
    frmAnEntregas.FlexTotInterior.Col = 3
    If (xforaprz + xnoprazo) > 0 Then
        frmAnEntregas.FlexTotInterior.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
    End If
    frmAnEntregas.FlexTotTotuf.Col = 0
    xnoprazo = CDbl(frmAnEntregas.FlexTotTotuf.Text)
    frmAnEntregas.FlexTotTotuf.Col = 1
    xforaprz = CDbl(frmAnEntregas.FlexTotTotuf.Text)
    frmAnEntregas.FlexTotTotuf.Col = 2
    If (xforaprz + xnoprazo) > 0 Then
        frmAnEntregas.FlexTotTotuf.Text = Format(xnoprazo / (xforaprz + xnoprazo), "##0.0%")
    End If
    
'PREENCHE COM "  - " OS CAMPOS QUE ESTIVEREM ZERADOS


    
'Atualiza o TAB de Resumo Geral
    
    
'REGIÃO NORTE

    For xy = 1 To 7
        frmAnEntregas.flexCtcs1.Col = 0
        frmAnEntregas.flexCtcs1.Row = xy
        frmAnEntregas.lblCTCsNO.Caption = CDbl(frmAnEntregas.lblCTCsNO.Caption) + CDbl(frmAnEntregas.flexCtcs1.Text)
        frmAnEntregas.flexCtcs1.Col = 1
        frmAnEntregas.lblNCTCsNO.Caption = CDbl(frmAnEntregas.lblNCTCsNO.Caption) + CDbl(frmAnEntregas.flexCtcs1.Text)
        frmAnEntregas.FlexTotUf1.Col = 0
        frmAnEntregas.FlexTotUf1.Row = xy
        frmAnEntregas.lblOnTimeNO.Caption = CDbl(frmAnEntregas.lblOnTimeNO.Caption) + CDbl(frmAnEntregas.FlexTotUf1.Text)
        frmAnEntregas.FlexTotUf1.Col = 1
        frmAnEntregas.lblDelayNO.Caption = CDbl(frmAnEntregas.lblDelayNO.Caption) + CDbl(frmAnEntregas.FlexTotUf1.Text)
    Next xy
    
    If CDbl(frmAnEntregas.lblCTCsNO.Caption) > 0 Then
        frmAnEntregas.lblPerc1NO.Caption = Format(CDbl(frmAnEntregas.lblNCTCsNO.Caption) / CDbl(frmAnEntregas.lblCTCsNO.Caption), "##0.0%")
    Else
        frmAnEntregas.lblPerc1NO.Caption = Format(0, "##0.0%")
    End If
    If (CDbl(frmAnEntregas.lblOnTimeNO.Caption) + CDbl(frmAnEntregas.lblDelayNO.Caption)) > 0 Then
        frmAnEntregas.lblPerc2NO.Caption = Format(CDbl(frmAnEntregas.lblOnTimeNO.Caption) / (CDbl(frmAnEntregas.lblOnTimeNO.Caption) + CDbl(frmAnEntregas.lblDelayNO.Caption)), "##0.0%")
    Else
        frmAnEntregas.lblPerc2NO.Caption = Format(0, "##0.0%")
    End If
    
'REGIÃO NORDESTE

    For xy = 8 To 16
        frmAnEntregas.flexCtcs1.Col = 0
        frmAnEntregas.flexCtcs1.Row = xy
        frmAnEntregas.lblCTCsND.Caption = CDbl(frmAnEntregas.lblCTCsND.Caption) + CDbl(frmAnEntregas.flexCtcs1.Text)
        frmAnEntregas.flexCtcs1.Col = 1
        frmAnEntregas.lblNCTCsND.Caption = CDbl(frmAnEntregas.lblNCTCsND.Caption) + CDbl(frmAnEntregas.flexCtcs1.Text)
        frmAnEntregas.FlexTotUf1.Col = 0
        frmAnEntregas.FlexTotUf1.Row = xy
        frmAnEntregas.lblOnTimeND.Caption = CDbl(frmAnEntregas.lblOnTimeND.Caption) + CDbl(frmAnEntregas.FlexTotUf1.Text)
        frmAnEntregas.FlexTotUf1.Col = 1
        frmAnEntregas.lblDelayND.Caption = CDbl(frmAnEntregas.lblDelayND.Caption) + CDbl(frmAnEntregas.FlexTotUf1.Text)
    Next xy
    
    If CDbl(frmAnEntregas.lblCTCsND.Caption) > 0 Then
        frmAnEntregas.lblPerc1ND.Caption = Format(CDbl(frmAnEntregas.lblNCTCsND.Caption) / CDbl(frmAnEntregas.lblCTCsND.Caption), "##0.0%")
    Else
        frmAnEntregas.lblPerc1ND.Caption = Format(0, "##0.0%")
    End If
    If (CDbl(frmAnEntregas.lblOnTimeND.Caption) + CDbl(frmAnEntregas.lblDelayND.Caption)) > 0 Then
        frmAnEntregas.lblPerc2ND.Caption = Format(CDbl(frmAnEntregas.lblOnTimeND.Caption) / (CDbl(frmAnEntregas.lblOnTimeND.Caption) + CDbl(frmAnEntregas.lblDelayND.Caption)), "##0.0%")
    Else
        frmAnEntregas.lblPerc2ND.Caption = Format(0, "##0.0%")
    End If
    
'REGIÃO SUDESTE

    For xy = 1 To 4
        frmAnEntregas.FlexCtcs2.Col = 0
        frmAnEntregas.FlexCtcs2.Row = xy
        frmAnEntregas.lblCTCsSD.Caption = CDbl(frmAnEntregas.lblCTCsSD.Caption) + CDbl(frmAnEntregas.FlexCtcs2.Text)
        frmAnEntregas.FlexCtcs2.Col = 1
        frmAnEntregas.lblNCTCsSD.Caption = CDbl(frmAnEntregas.lblNCTCsSD.Caption) + CDbl(frmAnEntregas.FlexCtcs2.Text)
        frmAnEntregas.FlexTotUf2.Col = 0
        frmAnEntregas.FlexTotUf2.Row = xy
        frmAnEntregas.lblOnTimeSD.Caption = CDbl(frmAnEntregas.lblOnTimeSD.Caption) + CDbl(frmAnEntregas.FlexTotUf2.Text)
        frmAnEntregas.FlexTotUf2.Col = 1
        frmAnEntregas.lblDelaySD.Caption = CDbl(frmAnEntregas.lblDelaySD.Caption) + CDbl(frmAnEntregas.FlexTotUf2.Text)
    Next xy
    
    If CDbl(frmAnEntregas.lblCTCsSD.Caption) > 0 Then
        frmAnEntregas.lblPerc1SD.Caption = Format(CDbl(frmAnEntregas.lblNCTCsSD.Caption) / CDbl(frmAnEntregas.lblCTCsSD.Caption), "##0.0%")
    Else
        frmAnEntregas.lblPerc1SD.Caption = Format(0, "##0.0%")
    End If
    If (CDbl(frmAnEntregas.lblOnTimeSD.Caption) + CDbl(frmAnEntregas.lblDelaySD.Caption)) > 0 Then
        frmAnEntregas.lblPerc2SD.Caption = Format(CDbl(frmAnEntregas.lblOnTimeSD.Caption) / (CDbl(frmAnEntregas.lblOnTimeSD.Caption) + CDbl(frmAnEntregas.lblDelaySD.Caption)), "##0.0%")
    Else
        frmAnEntregas.lblPerc2SD.Caption = Format(0, "##0.0%")
    End If
    
'REGIÃO SUL

    For xy = 5 To 7
        frmAnEntregas.FlexCtcs2.Col = 0
        frmAnEntregas.FlexCtcs2.Row = xy
        frmAnEntregas.lblCTCsSU.Caption = CDbl(frmAnEntregas.lblCTCsSU.Caption) + CDbl(frmAnEntregas.FlexCtcs2.Text)
        frmAnEntregas.FlexCtcs2.Col = 1
        frmAnEntregas.lblNCTCsSU.Caption = CDbl(frmAnEntregas.lblNCTCsSU.Caption) + CDbl(frmAnEntregas.FlexCtcs2.Text)
        frmAnEntregas.FlexTotUf2.Col = 0
        frmAnEntregas.FlexTotUf2.Row = xy
        frmAnEntregas.lblOnTimeSU.Caption = CDbl(frmAnEntregas.lblOnTimeSU.Caption) + CDbl(frmAnEntregas.FlexTotUf2.Text)
        frmAnEntregas.FlexTotUf2.Col = 1
        frmAnEntregas.lblDelaySU.Caption = CDbl(frmAnEntregas.lblDelaySU.Caption) + CDbl(frmAnEntregas.FlexTotUf2.Text)
    Next xy
    
    If CDbl(frmAnEntregas.lblCTCsSU.Caption) > 0 Then
        frmAnEntregas.lblPerc1SU.Caption = Format(CDbl(frmAnEntregas.lblNCTCsSU.Caption) / CDbl(frmAnEntregas.lblCTCsSU.Caption), "##0.0%")
    Else
        frmAnEntregas.lblPerc1SU.Caption = Format(0, "##0.0%")
    End If
    If (CDbl(frmAnEntregas.lblOnTimeSU.Caption) + CDbl(frmAnEntregas.lblDelaySU.Caption)) > 0 Then
        frmAnEntregas.lblPerc2SU.Caption = Format(CDbl(frmAnEntregas.lblOnTimeSU.Caption) / (CDbl(frmAnEntregas.lblOnTimeSU.Caption) + CDbl(frmAnEntregas.lblDelaySU.Caption)), "##0.0%")
    Else
        frmAnEntregas.lblPerc2SU.Caption = Format(0, "##0.0%")
    End If

'REGIÃO CENTRO OESTE

    For xy = 8 To 11
        frmAnEntregas.FlexCtcs2.Col = 0
        frmAnEntregas.FlexCtcs2.Row = xy
        frmAnEntregas.lblCTCsCO.Caption = CDbl(frmAnEntregas.lblCTCsCO.Caption) + CDbl(frmAnEntregas.FlexCtcs2.Text)
        frmAnEntregas.FlexCtcs2.Col = 1
        frmAnEntregas.lblNCTCsCO.Caption = CDbl(frmAnEntregas.lblNCTCsCO.Caption) + CDbl(frmAnEntregas.FlexCtcs2.Text)
        frmAnEntregas.FlexTotUf2.Col = 0
        frmAnEntregas.FlexTotUf2.Row = xy
        frmAnEntregas.lblOnTimeCO.Caption = CDbl(frmAnEntregas.lblOnTimeCO.Caption) + CDbl(frmAnEntregas.FlexTotUf2.Text)
        frmAnEntregas.FlexTotUf2.Col = 1
        frmAnEntregas.lblDelayCO.Caption = CDbl(frmAnEntregas.lblDelayCO.Caption) + CDbl(frmAnEntregas.FlexTotUf2.Text)
    Next xy
    
    If CDbl(frmAnEntregas.lblCTCsCO.Caption) > 0 Then
        frmAnEntregas.lblPerc1CO.Caption = Format(CDbl(frmAnEntregas.lblNCTCsCO.Caption) / CDbl(frmAnEntregas.lblCTCsCO.Caption), "##0.0%")
    Else
        frmAnEntregas.lblPerc1CO.Caption = Format(0, "##0.0%")
    End If
    If (CDbl(frmAnEntregas.lblOnTimeCO.Caption) + CDbl(frmAnEntregas.lblDelayCO.Caption)) > 0 Then
        frmAnEntregas.lblPerc2CO.Caption = Format(CDbl(frmAnEntregas.lblOnTimeCO.Caption) / (CDbl(frmAnEntregas.lblOnTimeCO.Caption) + CDbl(frmAnEntregas.lblDelayCO.Caption)), "##0.0%")
    Else
        frmAnEntregas.lblPerc2CO.Caption = Format(0, "##0.0%")
    End If
    
'TOTAL BRASIL

    frmAnEntregas.lblCTCsBR.Caption = CDbl(frmAnEntregas.lblCTCsNO) + CDbl(frmAnEntregas.lblCTCsND) + CDbl(frmAnEntregas.lblCTCsSD) + _
                                       CDbl(frmAnEntregas.lblCTCsSU) + CDbl(frmAnEntregas.lblCTCsCO)
    frmAnEntregas.lblNCTCsBR.Caption = CDbl(frmAnEntregas.lblNCTCsNO) + CDbl(frmAnEntregas.lblNCTCsND) + CDbl(frmAnEntregas.lblNCTCsSD) + _
                                       CDbl(frmAnEntregas.lblNCTCsSU) + CDbl(frmAnEntregas.lblNCTCsCO)
    
    If CDbl(frmAnEntregas.lblCTCsBR) > 0 Then
        frmAnEntregas.lblPerc1BR.Caption = Format(CDbl(frmAnEntregas.lblNCTCsBR) / CDbl(frmAnEntregas.lblCTCsBR), "##0.0%")
    Else
        frmAnEntregas.lblPerc1BR.Caption = Format(0, "##0.0%")
    End If
                                      
    frmAnEntregas.lblOnTimeBR.Caption = CDbl(frmAnEntregas.lblOnTimeNO) + CDbl(frmAnEntregas.lblOnTimeND) + CDbl(frmAnEntregas.lblOnTimeSD) + _
                                       CDbl(frmAnEntregas.lblOnTimeSU) + CDbl(frmAnEntregas.lblOnTimeCO)
    frmAnEntregas.lblDelayBR.Caption = CDbl(frmAnEntregas.lblDelayNO) + CDbl(frmAnEntregas.lblDelayND) + CDbl(frmAnEntregas.lblDelaySD) + _
                                       CDbl(frmAnEntregas.lblDelaySU) + CDbl(frmAnEntregas.lblDelayCO)
    
    If CDbl(frmAnEntregas.lblCTCsBR) > 0 Then
        frmAnEntregas.lblPerc2BR.Caption = Format(CDbl(frmAnEntregas.lblOnTimeBR) / (CDbl(frmAnEntregas.lblDelayBR) + CDbl(frmAnEntregas.lblOnTimeBR)), "##0.0%")
    Else
        frmAnEntregas.lblPerc2BR.Caption = Format(0, "##0.0%")
    End If
    
'ATUALIZA OS DADOS DOS GRÁFICOS

'BRASIL
    If CDbl(frmAnEntregas.lblOnTimeBR.Caption) > 0 Then
        frmAnEntregas.GrafBR.Visible = True
    End If
    frmAnEntregas.GrafBR.Column = 1
    frmAnEntregas.GrafBR.Data = CDbl(frmAnEntregas.lblOnTimeBR.Caption)
    frmAnEntregas.GrafBR.Column = 2
    frmAnEntregas.GrafBR.Data = CDbl(frmAnEntregas.lblDelayBR.Caption)
'NORTE
    If CDbl(frmAnEntregas.lblOnTimeNO.Caption) > 0 Then
        frmAnEntregas.GrafNO.Visible = True
    End If
    frmAnEntregas.GrafNO.Column = 1
    frmAnEntregas.GrafNO.Data = CDbl(frmAnEntregas.lblOnTimeNO.Caption)
    frmAnEntregas.GrafNO.Column = 2
    frmAnEntregas.GrafNO.Data = CDbl(frmAnEntregas.lblDelayNO.Caption)
'NORDESTE
    If CDbl(frmAnEntregas.lblOnTimeND.Caption) > 0 Then
        frmAnEntregas.GrafND.Visible = True
    End If
    frmAnEntregas.GrafND.Column = 1
    frmAnEntregas.GrafND.Data = CDbl(frmAnEntregas.lblOnTimeND.Caption)
    frmAnEntregas.GrafND.Column = 2
    frmAnEntregas.GrafND.Data = CDbl(frmAnEntregas.lblDelayND.Caption)
'SUDESTE
    If CDbl(frmAnEntregas.lblOnTimeSD.Caption) > 0 Then
        frmAnEntregas.GrafSD.Visible = True
    End If
    frmAnEntregas.GrafSD.Column = 1
    frmAnEntregas.GrafSD.Data = CDbl(frmAnEntregas.lblOnTimeSD.Caption)
    frmAnEntregas.GrafSD.Column = 2
    frmAnEntregas.GrafSD.Data = CDbl(frmAnEntregas.lblDelaySD.Caption)
'SUL
    If CDbl(frmAnEntregas.lblOnTimeSU.Caption) > 0 Then
        frmAnEntregas.GrafSU.Visible = True
    End If
    frmAnEntregas.GrafSU.Column = 1
    frmAnEntregas.GrafSU.Data = CDbl(frmAnEntregas.lblOnTimeSU.Caption)
    frmAnEntregas.GrafSU.Column = 2
    frmAnEntregas.GrafSU.Data = CDbl(frmAnEntregas.lblDelaySU.Caption)
'CENTRO-OESTE
    If CDbl(frmAnEntregas.lblOnTimeCO.Caption) > 0 Then
        frmAnEntregas.GrafCO.Visible = True
    End If
    frmAnEntregas.GrafCO.Column = 1
    frmAnEntregas.GrafCO.Data = CDbl(frmAnEntregas.lblOnTimeCO.Caption)
    frmAnEntregas.GrafCO.Column = 2
    frmAnEntregas.GrafCO.Data = CDbl(frmAnEntregas.lblDelayCO.Caption)
    Call zeracampos

End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkModal_Click()
    If chkModal.Value = 1 Then
        optRodo.Enabled = False
        optAir.Enabled = False
    Else
        optRodo.Enabled = True
        optAir.Enabled = True
        'If optAir.Enabled = True Then optAir.SetFocus
    End If
End Sub

Private Sub chkSair_Click()
    Unload Me
    'frmAnOcorr.Caption = "SAIR"
End Sub

Private Sub chkModal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        If cmdProcessa.Enabled = True Then cmdProcessa.SetFocus
        'SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub chkTodosCli_Click()
    If chkTodosCli.Value = 0 Then  'escolher por CGC
        lblNomeCli.Caption = ""
        txtCGCCli.Enabled = True
        txtBuscaNome.Enabled = True
        txtCGCCli.BackColor = &HC0FFFF       'amarelo
        txtBuscaNome.BackColor = &HC0FFFF    'amarelo
        chkTodosEstab.Enabled = True
        fraConsCli.Enabled = True
        txtCGCCli.SetFocus
    ElseIf chkTodosCli.Value = 1 Then  'todos os clientes
        lblNomeCli.Caption = "TODOS"
        txtCGCCli.Enabled = False
        txtBuscaNome.Enabled = False
        txtCGCCli.BackColor = &H80000005        'branco
        txtBuscaNome.BackColor = &H80000005     'branco
        chkTodosEstab.Enabled = False
        fraConsCli.Enabled = False
        mskPer1.SetFocus
    End If
End Sub

Private Sub chkTodosEstab_Click()
    If chkTodosEstab.Value = 0 Then
        txtCGCCli.MaxLength = 14
    ElseIf chkTodosEstab.Value = 1 Then
        txtCGCCli.MaxLength = 8
    End If
    txtCGCCli.SetFocus
End Sub

Private Sub cmdBusca_Click()
    If txtBuscaNome.Text = "" Then
    Else
        If de_informa.rsSel_ConsCadCliNome.State = 1 Then de_informa.rsSel_ConsCadCliNome.Close
        If optBuscaInic = True Then
            de_informa.Sel_ConsCadCliNome Trim(txtBuscaNome) & "%"
        Else
            de_informa.Sel_ConsCadCliNome "%" & Trim(txtBuscaNome) & "%"
        End If
        GridConsCli.DataMember = "sel_consCadCliNome"
        GridConsCli.Refresh
    End If
End Sub

Private Sub cmdProcessa_Click()
    Dim xcgc As String, xmodal As String

If CDate(mskPer2) - CDate(mskPer1) > 32 Then
    MsgBox "Período Escolhido Maior que 30 Dias ! Escolha um Período Menor."
    mskPer1.SetFocus
    Exit Sub
End If

If Me.Caption = "Análise Estatística" Then
    fraConsCli.Visible = False
    fraDados.Enabled = False
    fraMensagem.Visible = True
    frmEscCliPer.Height = 4950
    Animation1.Open App.Path & "\filemove.avi"
    DoEvents
    If chkTodosCli.Value = 0 Then  'escolha de um CGC
        xcgc = Trim(txtCGCCli.Text) & "%"
    ElseIf chkTodosCli.Value = 1 Then  'todos os clientes
        xcgc = "%"
    End If
    If chkModal.Value = 1 Then
        frmAnEstat.lblModal.Caption = "AÉREO/RODO"
        xmodal = "%"
    Else
        If optAir = True Then
            frmAnEstat.lblModal.Caption = "AÉREO"
            xmodal = "A%"
        Else
            frmAnEstat.lblModal.Caption = "RODOVIÁRIO"
            xmodal = "R%"
        End If
    End If
    frmAnEstat.lblCliente.Caption = lblNomeCli
    frmAnEstat.lblDataPer1.Caption = mskPer1.Text
    frmAnEstat.lblDataPer2.Caption = mskPer2.Text
    
    If de_informa.rsSel_AnOper.State = 1 Then de_informa.rsSel_AnOper.Close
    de_informa.Sel_AnOper xcgc, xmodal, mskPer1.Text, mskPer2.Text, txtFilial & "%"
    If de_informa.rsSel_AnOper.RecordCount > 0 Then
            
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "PROCESSO", xusuario, "ANALISE ESTATÍSTICA: " & lblNomeCli
        
        Call proc_dados
        'Call proc_dadosmutuo
    Else
        MsgBox "Não há dados para a seleção optada."
    End If
    
    frmEscCliPer.Height = 7125
    fraConsCli.Visible = True
    fraDados.Enabled = True
    fraMensagem.Visible = False
    Animation1.Close
    DoEvents
    
    Me.Hide
    
    frmAnEstat.Show

ElseIf Me.Caption = "Análise de Entregas" Then
    Dim xy As Integer
    Dim xCTC_AC, xCTC_AM, xCTC_AP, xCTC_PA, xCTC_RO, xCTC_RR, xCTC_TO, xCTC_AL, xCTC_BA, xCTC_SE, xCTC_PE As Long
    Dim xCTC_PB, xCTC_RN, xCTC_CE, xCTC_PI, xCTC_MA, xCTC_ES, xCTC_MG, xCTC_RJ, xCTC_SP As Long
    Dim xCTC_PR, xCTC_RS, xCTC_SC, xCTC_DF, xCTC_GO, xCTC_MS, xCTC_MT, xTotalCTC As Long
    Dim xCtcN_AC, xCtcN_AM, xCtcN_AP, xCtcN_PA, xCtcN_RO, xCtcN_RR, xCtcN_TO, xCtcN_AL, xCtcN_BA, xCtcN_SE, xCtcN_PE As Long
    Dim xCtcN_PB, xCtcN_RN, xCtcN_CE, xCtcN_PI, xCtcN_MA, xCtcN_ES, xCtcN_MG, xCtcN_RJ, xCtcN_SP, xCtcN_PR As Long
    Dim xCtcN_RS, xCtcN_SC, xCtcN_DF, xCtcN_GO, xCtcN_MS, xCtcN_MT, xTotalCtcN As Long
    Call zeracampos2
    fraConsCli.Visible = False
    fraDados.Enabled = False
    fraMensagem.Visible = True
    frmEscCliPer.Height = 4950
    Animation1.Open App.Path & "\filemove.avi"
    DoEvents
    If chkTodosCli.Value = 0 Then  'escolha de um CGC
        xcgc = Trim(txtCGCCli.Text) & "%"
    ElseIf chkTodosCli.Value = 1 Then  'todos os clientes
        xcgc = "%"
    End If

    'checar modal rodo / air
    
'    If chkModal.Value = 1 Then
'        xmodal = ""
'        frmAnEntregas.lblModal.Caption = "AÉREO/RODO"
'    Else
        If optAir = True Then
            xmodal = "AEREO"
            frmAnEntregas.lblModal.Caption = "AÉREO"
        Else
            xmodal = "RODOVIARIO"
            frmAnEntregas.lblModal.Caption = "RODOVIÁRIO"
        End If
'    End If
    frmAnEntregas.lblCliente.Caption = lblNomeCli
    frmAnEntregas.lblDataPer1.Caption = mskPer1.Text
    frmAnEntregas.lblDataPer2.Caption = mskPer2.Text
    frmAnEntregas.lblCgcRemet.Caption = xcgc
    
'    On erro GoTo TrataErro1
    
   If de_informa.rsSel_CTCsAnEntr.State = 1 Then de_informa.rsSel_CTCsAnEntr.Close
   If de_informa.rsSel_CtcNaoAnEntr.State = 1 Then de_informa.rsSel_CtcNaoAnEntr.Close
   de_informa.Sel_CTCsAnEntr xcgc, CDate(mskPer1.Text), CDate(mskPer2.Text), xmodal
   de_informa.Sel_CtcNaoAnEntr xcgc, CDate(mskPer1.Text), CDate(mskPer2.Text), xmodal, "1"
   If de_informa.rsSel_CTCsAnEntr.RecordCount > 0 Then
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "PROCESSO", xusuario, "ANALISE DE ENTREGAS: " & lblNomeCli
   
      de_informa.rsSel_CTCsAnEntr.MoveFirst
        frmAnEntregas.flexCtcs1.Col = 0
        frmAnEntregas.FlexCtcs2.Col = 0
      Do Until de_informa.rsSel_CTCsAnEntr.EOF
    
'DADOS DAS QUANTIDADES DE CTCs POR ESTADOS
    
'AC
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "AC" Then
            frmAnEntregas.flexCtcs1.Row = 1
            xCTC_AC = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_AC
        End If
'AM
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "AM" Then
            frmAnEntregas.flexCtcs1.Row = 2
            xCTC_AM = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_AM
        End If
'AP
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "AP" Then
            frmAnEntregas.flexCtcs1.Row = 3
            xCTC_AP = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_AP
        End If
'PA
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "PA" Then
            frmAnEntregas.flexCtcs1.Row = 4
            xCTC_PA = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_PA
        End If
'RO
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "RO" Then
            frmAnEntregas.flexCtcs1.Row = 5
            xCTC_RO = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_RO
        End If
'RR
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "RR" Then
            frmAnEntregas.flexCtcs1.Row = 6
            xCTC_RR = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_RR
        End If
'TO
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "TO" Then
            frmAnEntregas.flexCtcs1.Row = 7
            xCTC_TO = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_TO
        End If
'AL
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "AL" Then
            frmAnEntregas.flexCtcs1.Row = 8
            xCTC_AL = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_AL
        End If
'BA
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "BA" Then
            frmAnEntregas.flexCtcs1.Row = 9
            xCTC_BA = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_BA
        End If
'SE
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "SE" Then
            frmAnEntregas.flexCtcs1.Row = 10
            xCTC_SE = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_SE
        End If
'PE
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "PE" Then
            frmAnEntregas.flexCtcs1.Row = 11
            xCTC_PE = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_PE
        End If
'PB
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "PB" Then
            frmAnEntregas.flexCtcs1.Row = 12
            xCTC_PB = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_PB
        End If
'RN
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "RN" Then
            frmAnEntregas.flexCtcs1.Row = 13
            xCTC_RN = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_RN
        End If
'CE
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "CE" Then
            frmAnEntregas.flexCtcs1.Row = 14
            xCTC_CE = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_CE
        End If
'PI
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "PI" Then
            frmAnEntregas.flexCtcs1.Row = 15
            xCTC_PI = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_PI
        End If
'MA
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "MA" Then
            frmAnEntregas.flexCtcs1.Row = 16
            xCTC_MA = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCTC_MA
        End If
'ES
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "ES" Then
            frmAnEntregas.FlexCtcs2.Row = 1
            xCTC_ES = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_ES
        End If
'MG
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "MG" Then
            frmAnEntregas.FlexCtcs2.Row = 2
            xCTC_MG = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_MG
        End If
'RJ
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "RJ" Then
            frmAnEntregas.FlexCtcs2.Row = 3
            xCTC_RJ = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_RJ
        End If
'SP
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "SP" Then
            frmAnEntregas.FlexCtcs2.Row = 4
            xCTC_SP = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_SP
        End If
'PR
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "PR" Then
            frmAnEntregas.FlexCtcs2.Row = 5
            xCTC_PR = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_PR
        End If
'RS
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "RS" Then
            frmAnEntregas.FlexCtcs2.Row = 6
            xCTC_RS = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_RS
        End If
'SC
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "SC" Then
            frmAnEntregas.FlexCtcs2.Row = 7
            xCTC_SC = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_SC
        End If
'DF
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "DF" Then
            frmAnEntregas.FlexCtcs2.Row = 8
            xCTC_DF = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_DF
        End If
'GO
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "GO" Then
            frmAnEntregas.FlexCtcs2.Row = 9
            xCTC_GO = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_GO
        End If
'MS
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "MS" Then
            frmAnEntregas.FlexCtcs2.Row = 10
            xCTC_MS = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_MS
        End If
'MT
        If de_informa.rsSel_CTCsAnEntr.Fields("uf_dest") = "MT" Then
            frmAnEntregas.FlexCtcs2.Row = 11
            xCTC_MT = de_informa.rsSel_CTCsAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCTC_MT
        End If
       de_informa.rsSel_CTCsAnEntr.MoveNext
     Loop
  End If
  
'TOTAL DE CTCs

  xTotalCTC = xCTC_AC + xCTC_AM + xCTC_AP + xCTC_PA + xCTC_RO + xCTC_RR + xCTC_TO + xCTC_AL _
              + xCTC_BA + xCTC_SE + xCTC_PE + xCTC_PB + xCTC_RN + xCTC_CE + xCTC_PI + xCTC_MA _
              + xCTC_ES + xCTC_MG + xCTC_RJ + xCTC_SP + xCTC_PR + xCTC_RS + xCTC_SC + xCTC_DF _
              + xCTC_GO + xCTC_MS + xCTC_MT
              
  frmAnEntregas.flexTotCTCs.Col = 0
  frmAnEntregas.flexTotCTCs.Row = 0
  frmAnEntregas.flexTotCTCs.Text = xTotalCTC

   If de_informa.rsSel_CtcNaoAnEntr.RecordCount > 0 Then
      de_informa.rsSel_CtcNaoAnEntr.MoveFirst
        frmAnEntregas.flexCtcs1.Col = 1
        frmAnEntregas.FlexCtcs2.Col = 1
      Do Until de_informa.rsSel_CtcNaoAnEntr.EOF
    
    'DADOS DAS QUANTIDADES DE CTCs NAO ENTREGUES POR ESTADOS
    
'AC
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "AC" Then
            frmAnEntregas.flexCtcs1.Row = 1
            xCtcN_AC = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_AC
        End If
'AM
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "AM" Then
            frmAnEntregas.flexCtcs1.Row = 2
            xCtcN_AM = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_AM
        End If
'AP
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "AP" Then
            frmAnEntregas.flexCtcs1.Row = 3
            xCtcN_AP = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_AP
        End If
'PA
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "PA" Then
            frmAnEntregas.flexCtcs1.Row = 4
            xCtcN_PA = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_PA
        End If
        
'RO
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "RO" Then
            frmAnEntregas.flexCtcs1.Row = 5
            xCtcN_RO = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_RO
        End If
'RR
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "RR" Then
            frmAnEntregas.flexCtcs1.Row = 6
            xCtcN_RR = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_RR
        End If
'TO
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "TO" Then
            frmAnEntregas.flexCtcs1.Row = 7
            xCtcN_TO = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_TO
        End If
'AL
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "AL" Then
            frmAnEntregas.flexCtcs1.Row = 8
            xCtcN_AL = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_AL
        End If
'BA
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "BA" Then
            frmAnEntregas.flexCtcs1.Row = 9
            xCtcN_BA = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_BA
        End If
'SE
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "SE" Then
            frmAnEntregas.flexCtcs1.Row = 10
            xCtcN_SE = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_SE
        End If
'PE
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "PE" Then
            frmAnEntregas.flexCtcs1.Row = 11
            xCtcN_PE = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_PE
        End If
'PB
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "PB" Then
            frmAnEntregas.flexCtcs1.Row = 12
            xCtcN_PB = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_PB
        End If
'RN
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "RN" Then
            frmAnEntregas.flexCtcs1.Row = 13
            xCtcN_RN = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_RN
        End If
'CE
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "CE" Then
            frmAnEntregas.flexCtcs1.Row = 14
            xCtcN_CE = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_CE
        End If
'PI
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "PI" Then
            frmAnEntregas.flexCtcs1.Row = 15
            xCtcN_PI = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_PI
        End If
'MA
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "MA" Then
            frmAnEntregas.flexCtcs1.Row = 16
            xCtcN_MA = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.flexCtcs1.Text = xCtcN_MA
        End If
'ES
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "ES" Then
            frmAnEntregas.FlexCtcs2.Row = 1
            xCtcN_ES = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_ES
        End If
'MG
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "MG" Then
            frmAnEntregas.FlexCtcs2.Row = 2
            xCtcN_MG = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_MG
        End If
'RJ
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "RJ" Then
            frmAnEntregas.FlexCtcs2.Row = 3
            xCtcN_RJ = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_RJ
        End If
'SP
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "SP" Then
            frmAnEntregas.FlexCtcs2.Row = 4
            xCtcN_SP = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_SP
        End If
'PR
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "PR" Then
            frmAnEntregas.FlexCtcs2.Row = 5
            xCtcN_PR = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_PR
        End If
'RS
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "RS" Then
            frmAnEntregas.FlexCtcs2.Row = 6
            xCtcN_RS = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_RS
        End If
'SC
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "SC" Then
            frmAnEntregas.FlexCtcs2.Row = 7
            xCtcN_SC = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_SC
        End If
'DF
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "DF" Then
            frmAnEntregas.FlexCtcs2.Row = 8
            xCtcN_DF = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_DF
        End If
'GO
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "GO" Then
            frmAnEntregas.FlexCtcs2.Row = 9
            xCtcN_GO = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_GO
        End If
'MS
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "MS" Then
            frmAnEntregas.FlexCtcs2.Row = 10
            xCtcN_MS = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_MS
        End If
'MT
        If de_informa.rsSel_CtcNaoAnEntr.Fields("uf_dest") = "MT" Then
            frmAnEntregas.FlexCtcs2.Row = 11
            xCtcN_MT = de_informa.rsSel_CtcNaoAnEntr.Fields("tot")
            frmAnEntregas.FlexCtcs2.Text = xCtcN_MT
        End If
        de_informa.rsSel_CtcNaoAnEntr.MoveNext
     Loop
  End If
    
'TOTAL DE CTCs NAO ENTREGUES

  xTotalCtcN = xCtcN_AC + xCtcN_AM + xCtcN_AP + xCtcN_PA + xCtcN_RO + xCtcN_RR + xCtcN_TO + xCtcN_AL _
              + xCtcN_BA + xCtcN_SE + xCtcN_PE + xCtcN_PB + xCtcN_RN + xCtcN_CE + xCtcN_PI + xCtcN_MA _
              + xCtcN_ES + xCtcN_MG + xCtcN_RJ + xCtcN_SP + xCtcN_PR + xCtcN_RS + xCtcN_SC + xCtcN_DF _
              + xCtcN_GO + xCtcN_MS + xCtcN_MT
              
  frmAnEntregas.flexTotCTCs.Col = 1
  frmAnEntregas.flexTotCTCs.Text = xTotalCtcN
        
'% CTCs Nao Entregues / Total de CTC
    
    frmAnEntregas.flexCtcs1.Col = 2
    frmAnEntregas.flexCtcs1.Row = 1
    If xCTC_AC > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_AC / xCTC_AC, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 2
    If xCTC_AM > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_AM / xCTC_AM, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 3
    If xCTC_AP > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_AP / xCTC_AP, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 4
    If xCTC_PA > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_PA / xCTC_PA, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 5
    If xCTC_RO > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_RO / xCTC_RO, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 6
    If xCTC_RR > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_RR / xCTC_RR, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 7
    If xCTC_TO > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_TO / xCTC_TO, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 8
    If xCTC_AL > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_AL / xCTC_AL, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 9
    If xCTC_BA > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_BA / xCTC_BA, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 10
    If xCTC_SE > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_SE / xCTC_SE, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 11
    If xCTC_PE > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_PE / xCTC_PE, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 12
    If xCTC_PB > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_PB / xCTC_PB, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 13
    If xCTC_RN > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_RN / xCTC_RN, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 14
    If xCTC_CE > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_CE / xCTC_CE, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 15
    If xCTC_PI > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_PI / xCTC_PI, "##0.0%")
    End If
    frmAnEntregas.flexCtcs1.Row = 16
    If xCTC_MA > 0 Then
        frmAnEntregas.flexCtcs1.Text = Format(xCtcN_MA / xCTC_MA, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Col = 2
    frmAnEntregas.FlexCtcs2.Row = 1
    If xCTC_ES > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_ES / xCTC_ES, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 2
    If xCTC_MG > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_MG / xCTC_MG, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 3
    If xCTC_RJ > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_RJ / xCTC_RJ, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 4
    If xCTC_SP > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_SP / xCTC_SP, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 5
    If xCTC_PR > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_PR / xCTC_PR, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 6
    If xCTC_RS > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_RS / xCTC_RS, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 7
    If xCTC_SC > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_SC / xCTC_SC, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 8
    If xCTC_DF > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_DF / xCTC_DF, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 9
    If xCTC_GO > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_GO / xCTC_GO, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 10
    If xCTC_MS > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_MS / xCTC_MS, "##0.0%")
    End If
    frmAnEntregas.FlexCtcs2.Row = 11
    If xCTC_MT > 0 Then
        frmAnEntregas.FlexCtcs2.Text = Format(xCtcN_MT / xCTC_MT, "##0.0%")
    End If
    frmAnEntregas.flexTotCTCs.Col = 2
    If xTotalCTC > 0 Then
        frmAnEntregas.flexTotCTCs.Text = Format(xTotalCtcN / xTotalCTC, "##0.0%")
    End If
    
'chama sub de atualização dos prazos no flexgrid
    
    Call gridprazos
    
    If chkAnalise.Value = 1 Then
        frmAnEntregas.chkAnalise.Value = 1
        If de_informa.rsSel_CtcsAtrasosComAbono.State = 1 Then de_informa.rsSel_CtcsAtrasosComAbono.Close
        de_informa.Sel_CtcsAtrasosComAbono xcgc, CDate(mskPer1), CDate(mskPer2), xmodal, "%"
        frmAnEntregas.gridAtrasos.DataMember = "sel_ctcsatrasoscomabono"
        frmAnEntregas.gridAtrasos.Refresh
        frmAnEntregas.fraAtraso.Caption = "CTCs em Atraso: " & de_informa.rsSel_CtcsAtrasosComAbono.RecordCount
        If de_informa.rsSel_CtcsAtrasosComAbono.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
            frmAnEntregas.GridConsOcorr.Refresh
            
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
            frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
            frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
            
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
            Else
                frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
            End If
            
            'frmAnEntregas.cmdAbona.Enabled = True
        End If
        
    Else
        frmAnEntregas.chkAnalise.Value = 0
        If de_informa.rsSel_CtcsAtrasosSemAbono.State = 1 Then de_informa.rsSel_CtcsAtrasosSemAbono.Close
        de_informa.Sel_CtcsAtrasosSemAbono xcgc, CDate(mskPer1), CDate(mskPer2), xmodal, "%"
        frmAnEntregas.gridAtrasos.DataMember = "sel_ctcsatrasossemabono"
        frmAnEntregas.gridAtrasos.Refresh
        frmAnEntregas.fraAtraso.Caption = "CTCs em Atraso: " & de_informa.rsSel_CtcsAtrasosSemAbono.RecordCount
        If de_informa.rsSel_CtcsAtrasosSemAbono.RecordCount > 0 Then
            If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
            de_informa.Sel_ConsOcorr2 frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.GridConsOcorr.DataMember = "Sel_ConsOcorr2"
            frmAnEntregas.GridConsOcorr.Refresh
            
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr frmAnEntregas.gridAtrasos.Columns(0), "01"
            
            frmAnEntregas.lblDtBxPre = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
            frmAnEntregas.lblRecebPre = de_informa.rsSel_ConsOcorr.Fields("recebpre")
            frmAnEntregas.lblUsuBxPre = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
            frmAnEntregas.lblUsuDtBaixaPre = de_informa.rsSel_ConsOcorr.Fields("usu_datapre")
            
            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) Then
                frmAnEntregas.lblDtBx = ""
                frmAnEntregas.lblReceb = ""
                frmAnEntregas.lblUsuBx = ""
                frmAnEntregas.lblUsuDtBaixa = ""
            Else
                frmAnEntregas.lblDtBx = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                frmAnEntregas.lblReceb = de_informa.rsSel_ConsOcorr.Fields("receb")
                frmAnEntregas.lblUsuBx = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                frmAnEntregas.lblUsuDtBaixa = de_informa.rsSel_ConsOcorr.Fields("usu_databx")
            End If
           ' frmAnEntregas.cmdAbona.Enabled = True
        End If
        
    End If
    
    frmEscCliPer.Height = 7125
    fraConsCli.Visible = True
    fraDados.Enabled = True
    fraMensagem.Visible = False
    Animation1.Close
    DoEvents
    
    Me.Hide
    
    frmAnEntregas.Show

ElseIf Me.Caption = "Análise de Ocorrências" Then
    Call calc_anocorr
End If



End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    GridConsCli.DataMember = ""
    GridConsCli.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEscCliPer = Nothing
End Sub

Private Sub GridConsCli_Click()
    If chkTodosEstab.Value = 0 Then
        txtCGCCli.Text = GridConsCli.Columns(0)
    Else
        txtCGCCli.Text = Mid(GridConsCli.Columns(0), 1, 8)
    End If
    lblNomeCli.Caption = GridConsCli.Columns(1)
    txtCGCCli.SetFocus
    mskPer1.SetFocus
End Sub

Private Sub mskPer1_GotFocus()
    mskPer1.SelStart = 0
    mskPer1.SelLength = 10
End Sub

Private Sub mskPer1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer1_LostFocus()
    If mskPer1.Text <> "__/__/____" Then
        mskPer1.Text = century(mskPer1.Text)
        If IsDate(mskPer1.Text) = False Or Mid(mskPer1.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
        If CDate(mskPer1.Text) > datahora("data") Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
        If IsDate(mskPer2.Text) = True And lblNomeCli.Caption <> "" Then
            cmdProcessa.Enabled = True
        End If
    Else
        cmdProcessa.Enabled = False
    End If
End Sub

Private Sub mskPer2_GotFocus()
    mskPer2.SelStart = 0
    mskPer2.SelLength = 10
End Sub

Private Sub mskPer2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub mskPer2_LostFocus()
    If mskPer2.Text <> "__/__/____" Then
        mskPer2.Text = century(mskPer2.Text)
        If IsDate(mskPer2.Text) = False Or Mid(mskPer2.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
        If CDate(mskPer2.Text) > datahora("data") Then
            MsgBox "Data Maior que Hoje", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
        If IsDate(mskPer1.Text) = True And lblNomeCli.Caption <> "" Then
            cmdProcessa.Enabled = True
        End If
    Else
        cmdProcessa.Enabled = False
    End If
End Sub

Private Sub optAir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        If cmdProcessa.Enabled = True Then cmdProcessa.SetFocus
        'SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optRodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        If cmdProcessa.Enabled = True Then cmdProcessa.SetFocus
        'SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtBuscaNome_GotFocus()
    txtBuscaNome.SelStart = 0
    txtBuscaNome.SelLength = 25
End Sub

Private Sub txtBuscaNome_LostFocus()
    txtBuscaNome.Text = UCase(txtBuscaNome)
End Sub

Private Sub txtCGCCli_Change()
    If Len(txtCGCCli) = txtCGCCli.MaxLength Then mskPer1.SetFocus
End Sub

Private Sub txtCGCCli_GotFocus()
    txtCGCCli.SelStart = 0
    txtCGCCli.SelLength = 14
End Sub

Private Sub txtCGCCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCGCCli_LostFocus()
    If txtCGCCli.Text <> "" Then
        If de_informa.rsSel_ConsCadCli.State = 1 Then de_informa.rsSel_ConsCadCli.Close
        de_informa.Sel_ConsCadCli Trim(txtCGCCli) & "%"
        If de_informa.rsSel_ConsCadCli.RecordCount > 0 Then
            lblNomeCli.Caption = de_informa.rsSel_ConsCadCli.Fields("nome")
            If IsDate(mskPer1.Text) = True And IsDate(mskPer2.Text) = True Then
                cmdProcessa.Enabled = True
            End If
        Else
            txtCGCCli.SetFocus
            cmdProcessa.Enabled = False
        End If
    Else
        lblNomeCli.Caption = ""
        cmdProcessa.Enabled = False
    End If
End Sub
Private Sub zeracampos()
Dim xy As Integer, xyz As Integer
    For xyz = 1 To 3
        frmAnEntregas.flexCtcs1.Col = xyz - 1
        frmAnEntregas.flexTotCTCs.Col = xyz - 1
        frmAnEntregas.FlexCapitais1.Col = xyz
        frmAnEntregas.FlexTotCapitais.Col = xyz
        frmAnEntregas.FlexInterior1.Col = xyz
        frmAnEntregas.FlexTotInterior.Col = xyz
        frmAnEntregas.FlexTotUf1.Col = xyz - 1
        frmAnEntregas.FlexTotTotuf.Col = xyz - 1
        frmAnEntregas.FlexCtcs2.Col = xyz - 1
        frmAnEntregas.FlexCapitais2.Col = xyz
        frmAnEntregas.FlexInterior2.Col = xyz
        frmAnEntregas.FlexTotUf2.Col = xyz - 1

        For xy = 1 To 16
            frmAnEntregas.flexCtcs1.Row = xy
            frmAnEntregas.FlexCapitais1.Row = xy
            frmAnEntregas.FlexInterior1.Row = xy
            frmAnEntregas.FlexTotUf1.Row = xy
            If Val(frmAnEntregas.flexCtcs1.Text) = 0 Then frmAnEntregas.flexCtcs1.Text = "        -"
            If Val(frmAnEntregas.FlexCapitais1.Text) = 0 Then frmAnEntregas.FlexCapitais1.Text = "        -"
            If Val(frmAnEntregas.FlexInterior1.Text) = 0 Then frmAnEntregas.FlexInterior1.Text = "        -"
            If Val(frmAnEntregas.FlexTotUf1.Text) = 0 Then frmAnEntregas.FlexTotUf1.Text = "        -"
        Next xy
        For xy = 1 To 11
            frmAnEntregas.FlexCtcs2.Row = xy
            frmAnEntregas.FlexCapitais2.Row = xy
            frmAnEntregas.FlexInterior2.Row = xy
            frmAnEntregas.FlexTotUf2.Row = xy
            If Val(frmAnEntregas.FlexCtcs2.Text) = 0 Then frmAnEntregas.FlexCtcs2.Text = "        -"
            If Val(frmAnEntregas.FlexCapitais2.Text) = 0 Then frmAnEntregas.FlexCapitais2.Text = "        -"
            If Val(frmAnEntregas.FlexInterior2.Text) = 0 Then frmAnEntregas.FlexInterior2.Text = "        -"
            If Val(frmAnEntregas.FlexTotUf2.Text) = 0 Then frmAnEntregas.FlexTotUf2.Text = "        -"
        Next xy
        If Val(frmAnEntregas.flexTotCTCs.Text) = 0 Then frmAnEntregas.flexTotCTCs.Text = "        -"
        If Val(frmAnEntregas.FlexTotCapitais.Text) = 0 Then frmAnEntregas.FlexTotCapitais.Text = "        -"
        If Val(frmAnEntregas.FlexTotInterior.Text) = 0 Then frmAnEntregas.FlexTotInterior.Text = "        -"
        If Val(frmAnEntregas.FlexTotTotuf.Text) = 0 Then frmAnEntregas.FlexTotTotuf.Text = "        -"
    Next xyz
        frmAnEntregas.FlexTotCapitais.Col = 0
        frmAnEntregas.FlexTotInterior.Col = 0
        frmAnEntregas.FlexTotCapitais.Text = "     ---"
        frmAnEntregas.FlexTotInterior.Text = "     ---"
End Sub


Private Sub zeracampos2()
Dim xy As Integer, xyz As Integer
    For xyz = 1 To 3
        frmAnEntregas.flexCtcs1.Col = xyz - 1
        frmAnEntregas.flexTotCTCs.Col = xyz - 1
        frmAnEntregas.FlexCapitais1.Col = xyz
        frmAnEntregas.FlexTotCapitais.Col = xyz
        frmAnEntregas.FlexInterior1.Col = xyz
        frmAnEntregas.FlexTotInterior.Col = xyz
        frmAnEntregas.FlexTotUf1.Col = xyz - 1
        frmAnEntregas.FlexTotTotuf.Col = xyz - 1
        frmAnEntregas.FlexCtcs2.Col = xyz - 1
        frmAnEntregas.FlexCapitais2.Col = xyz
        frmAnEntregas.FlexInterior2.Col = xyz
        frmAnEntregas.FlexTotUf2.Col = xyz - 1

        For xy = 1 To 16
            frmAnEntregas.flexCtcs1.Row = xy
            frmAnEntregas.FlexCapitais1.Row = xy
            frmAnEntregas.FlexInterior1.Row = xy
            frmAnEntregas.FlexTotUf1.Row = xy
            frmAnEntregas.flexCtcs1.Text = "0"
            frmAnEntregas.FlexCapitais1.Text = "0"
            frmAnEntregas.FlexInterior1.Text = "0"
            frmAnEntregas.FlexTotUf1.Text = "0"
        Next xy
        For xy = 1 To 11
            frmAnEntregas.FlexCtcs2.Row = xy
            frmAnEntregas.FlexCapitais2.Row = xy
            frmAnEntregas.FlexInterior2.Row = xy
            frmAnEntregas.FlexTotUf2.Row = xy
            frmAnEntregas.FlexCtcs2.Text = 0
            frmAnEntregas.FlexCapitais2.Text = "0"
            frmAnEntregas.FlexInterior2.Text = "0"
            frmAnEntregas.FlexTotUf2.Text = "0"
        Next xy
        frmAnEntregas.flexTotCTCs.Text = "0"
        frmAnEntregas.FlexTotCapitais.Text = "0"
        frmAnEntregas.FlexTotInterior.Text = "0"
        frmAnEntregas.FlexTotTotuf.Text = "0"
    Next xyz
        frmAnEntregas.FlexTotCapitais.Col = "0"
        frmAnEntregas.FlexTotInterior.Col = "0"
        frmAnEntregas.FlexTotCapitais.Text = "0"
        frmAnEntregas.FlexTotInterior.Text = "0"
        frmAnEntregas.lblCTCsBR = "0"
        frmAnEntregas.lblCTCsCO = "0"
        frmAnEntregas.lblCTCsND = "0"
        frmAnEntregas.lblCTCsNO = "0"
        frmAnEntregas.lblCTCsSD = "0"
        frmAnEntregas.lblCTCsSU = "0"
        frmAnEntregas.lblNCTCsBR = "0"
        frmAnEntregas.lblNCTCsCO = "0"
        frmAnEntregas.lblNCTCsND = "0"
        frmAnEntregas.lblNCTCsNO = "0"
        frmAnEntregas.lblNCTCsSD = "0"
        frmAnEntregas.lblNCTCsSU = "0"
        frmAnEntregas.lblOnTimeBR = "0"
        frmAnEntregas.lblOnTimeCO = "0"
        frmAnEntregas.lblOnTimeND = "0"
        frmAnEntregas.lblOnTimeNO = "0"
        frmAnEntregas.lblOnTimeSD = "0"
        frmAnEntregas.lblOnTimeSU = "0"
        frmAnEntregas.lblDelayBR = "0"
        frmAnEntregas.lblDelayCO = "0"
        frmAnEntregas.lblDelayND = "0"
        frmAnEntregas.lblDelayNO = "0"
        frmAnEntregas.lblDelaySD = "0"
        frmAnEntregas.lblDelaySU = "0"
        
End Sub
Private Sub calc_anocorr()
    Dim xcgc As String, xmodal As String
    Dim xdata As Date
    Dim xtotinf1 As Long, xtotinf2 As Long
    Dim xtotinf41 As Long, xtotinf42 As Long, xtotinf6 As Long, xseries As Long, xoutros As Long
    Dim xtotinf41b As Long, xtotinf42b As Long, xtotinf6b As Long
    fraConsCli.Visible = False
    fraDados.Enabled = False
    fraMensagem.Visible = True
    frmEscCliPer.Height = 4950
    Animation1.Open App.Path & "\filemove.avi"
    DoEvents
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "ANALISE DE OCORRÊNCIAS: " & lblNomeCli
    
    If chkTodosCli.Value = 0 Then  'escolha de um CGC
        xcgc = Trim(txtCGCCli.Text) & "%"
    ElseIf chkTodosCli.Value = 1 Then  'todos os clientes
        xcgc = "%"
    End If
    
    If chkModal.Value = 1 Then
        xmodal = "%"
        frmAnOcorr.lblModal = "RODO/AÉREO"
    Else
        If optAir = True Then
            xmodal = "A%"
            frmAnOcorr.lblModal = "AÉREO"
        Else
            xmodal = "R%"
            frmAnOcorr.lblModal = "RODOVIÁRIO"
        End If
    End If
        
        frmAnOcorr.lblCliente = frmEscCliPer.lblNomeCli
        frmAnOcorr.lblDataPer1 = frmEscCliPer.mskPer1
        frmAnOcorr.lblDataPer2 = frmEscCliPer.mskPer2
        frmAnOcorr.lblCgcCli = xcgc
        
'quantidade de CTCs do período
        
        If de_informa.rsSel_QtdeCTC1.State = 1 Then de_informa.rsSel_QtdeCTC1.Close
        de_informa.Sel_QtdeCTC1 CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, xmodal
        xtotinf1 = de_informa.rsSel_QtdeCTC1.Fields("qtdtot")
        frmAnOcorr.lblInf1 = de_informa.rsSel_QtdeCTC1.Fields("qtdtot")
        
'quantidade de NFs do período

        If de_informa.rsSel_QtdeNF1.State = 1 Then de_informa.rsSel_QtdeNF1.Close
        de_informa.Sel_QtdeNF1 CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, xmodal
        xtotinf2 = de_informa.rsSel_QtdeNF1.Fields("qtdtot")
        frmAnOcorr.lblInf2 = de_informa.rsSel_QtdeNF1.Fields("qtdtot")
        
'quantidade de CTCs ENTREGUES: tem_ocorr = 1
        
'POR CTCS
        If de_informa.rsSel_QtdCtcOcorr.State = 1 Then de_informa.rsSel_QtdCtcOcorr.Close
        de_informa.Sel_QtdCtcOcorr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "1", xmodal
        If xtotinf1 > 0 Then frmAnOcorr.lblPercInf3.Caption = Format(de_informa.rsSel_QtdCtcOcorr.Fields("qtdtot") / xtotinf1, "##0.0%")
        frmAnOcorr.lblInf3 = de_informa.rsSel_QtdCtcOcorr.Fields("qtdtot")
'POR NFS
        If de_informa.rsSel_QtdNfOcorr.State = 1 Then de_informa.rsSel_QtdNfOcorr.Close
        de_informa.Sel_QtdNFOcorr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "1", xmodal
        If xtotinf2 > 0 Then frmAnOcorr.lblPercInf3b.Caption = Format(de_informa.rsSel_QtdNfOcorr.Fields("qtdtot") / xtotinf2, "##0.0%")
        frmAnOcorr.lblInf3b = de_informa.rsSel_QtdNfOcorr.Fields("qtdtot")


'quantidade de CTCs NAO ENTREGUES com OCORRENCIAS FECHADAS: tem_ocorr = 0

'POR CTCS
        If de_informa.rsSel_QtdCtcOcorr.State = 1 Then de_informa.rsSel_QtdCtcOcorr.Close
        de_informa.Sel_QtdCtcOcorr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "0", xmodal
        xtotinf41 = de_informa.rsSel_QtdCtcOcorr.Fields("qtdtot")
        frmAnOcorr.lblInf41 = xtotinf41
'POR NFS
        If de_informa.rsSel_QtdNfOcorr.State = 1 Then de_informa.rsSel_QtdNfOcorr.Close
        de_informa.Sel_QtdNFOcorr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "0", xmodal
        xtotinf41b = de_informa.rsSel_QtdNfOcorr.Fields("qtdtot")
        frmAnOcorr.lblInf41b = xtotinf41b
        
'quantidade de CTCs NAO ENTREGUES COM OCORRÊNCIAS PENDENTES: tem_ocorr = 2
        
'POR CTCS
        If de_informa.rsSel_QtdCtcOcorr.State = 1 Then de_informa.rsSel_QtdCtcOcorr.Close
        de_informa.Sel_QtdCtcOcorr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "2", xmodal
        xtotinf42 = de_informa.rsSel_QtdCtcOcorr.Fields("qtdtot")
        frmAnOcorr.lblInf42 = xtotinf42
'POR NFS
        If de_informa.rsSel_QtdNfOcorr.State = 1 Then de_informa.rsSel_QtdNfOcorr.Close
        de_informa.Sel_QtdNFOcorr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "2", xmodal
        xtotinf42b = de_informa.rsSel_QtdNfOcorr.Fields("qtdtot")
        frmAnOcorr.lblInf42b = xtotinf42b

'ATUALIZA O LABEL DO TOTAL DE CTCs COM OCORRÊNCIAS DO ÍTEM DE INFORMAÇÃO 4
        
'POR CTCs
        If xtotinf1 > 0 Then frmAnOcorr.lblPercInf4.Caption = Format((CDbl(xtotinf41) + CDbl(xtotinf42)) / xtotinf1, "##0.0%")
        frmAnOcorr.lblInf4tot = CDbl(xtotinf41) + CDbl(xtotinf42)
'POR NFS
        If xtotinf2 > 0 Then frmAnOcorr.lblPercInf4b.Caption = Format((CDbl(xtotinf41b) + CDbl(xtotinf42b)) / xtotinf2, "##0.0%")
        frmAnOcorr.lblInf4totb = CDbl(xtotinf41b) + CDbl(xtotinf42b)
        
'quantidade de CTCs EM TRÂNSITO: tem_ocorr = N e Prev_Entrega <= Hoje

'POR CTCS
        If de_informa.rsSel_CtcEmTransito.State = 1 Then de_informa.rsSel_CtcEmTransito.Close
        de_informa.Sel_CtcEmTransito CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "N", datahora("data"), xmodal
        
            frmAnOcorr.lblInf5 = de_informa.rsSel_CtcEmTransito.Fields("qtdtot")
            If xtotinf1 > 0 Then frmAnOcorr.lblPercInf5.Caption = Format(frmAnOcorr.lblInf5 / xtotinf1, "##0.0%")
'POR NFS
        If de_informa.rsSel_NfEmTransito.State = 1 Then de_informa.rsSel_NfEmTransito.Close
        de_informa.Sel_nfEmTransito CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "N", datahora("data"), xmodal
        
            frmAnOcorr.lblInf5b = de_informa.rsSel_NfEmTransito.Fields("qtdtot")
            If xtotinf2 > 0 Then frmAnOcorr.lblPercInf5b.Caption = Format(frmAnOcorr.lblInf5b / xtotinf2, "##0.0%")
        
'quantidade de CTCs SEM POSIÇÃO: tem_ocorr = N e Prev_Entrega >=  Hoje
        
'POR CTCs
        If de_informa.rsSel_CtcPendenteEntr.State = 1 Then de_informa.rsSel_CtcPendenteEntr.Close
        de_informa.Sel_CtcPendenteEntr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "N", datahora("data"), xmodal
        xtotinf6 = de_informa.rsSel_CtcPendenteEntr.Fields("qtdtot")
        frmAnOcorr.lblInf6 = xtotinf6
        If xtotinf1 > 0 Then frmAnOcorr.lblPercInf6.Caption = Format(xtotinf6 / xtotinf1, "##0.0%")
'POR NFS
        If de_informa.rsSel_NfPendenteEntr.State = 1 Then de_informa.rsSel_NfPendenteEntr.Close
        de_informa.Sel_nfPendenteEntr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "N", datahora("data"), xmodal
        xtotinf6b = de_informa.rsSel_NfPendenteEntr.Fields("qtdtot")
        frmAnOcorr.lblInf6b = xtotinf6b
        If xtotinf2 > 0 Then frmAnOcorr.lblPercInf6b.Caption = Format(xtotinf6b / xtotinf2, "##0.0%")

'ATUALIZA O LABEL DO TOTAL DE CTCs COM PENDÊNCIAS DO ÍTEM DE INFORMAÇÃO 6 (SOMA ÍTEM 4.2 + 5)
        
'POR CTCs
        If xtotinf1 > 0 Then frmAnOcorr.lblPercInf7.Caption = Format((CDbl(xtotinf42) + CDbl(xtotinf6)) / xtotinf1, "##0.0%")
        frmAnOcorr.lblInf7 = CDbl(xtotinf42) + CDbl(xtotinf6)
'POR NFS
        If xtotinf2 > 0 Then frmAnOcorr.lblPercInf7b.Caption = Format((CDbl(xtotinf42b) + CDbl(xtotinf6b)) / xtotinf2, "##0.0%")
        frmAnOcorr.lblInf7b = CDbl(xtotinf42b) + CDbl(xtotinf6b)
        
'ATUALIZA O GRID DE OCORRÊNCIAS MAIS FREQUENTES

        If de_informa.rsSel_ABCOcorr.State = 1 Then de_informa.rsSel_ABCOcorr.Close
        de_informa.Sel_ABCOcorr CDate(mskPer1.Text), CDate(mskPer2.Text), xcgc, "02" 'OCORRÊNCIAS >= 02 (01=ENTREGA E 00=OCORR FECHADA)
        Set frmAnOcorr.GridOcorrABC.DataSource = de_informa
        frmAnOcorr.GridOcorrABC.DataMember = "sel_abcocorr"
        frmAnOcorr.GridOcorrABC.Refresh
        
'CONFIGURA OS DADOS PARA O GRÁFICO

        If de_informa.rsSel_ABCOcorr.RecordCount > 15 Then
            frmAnOcorr.GrafOcorrABC.RowCount = 15
        Else
            frmAnOcorr.GrafOcorrABC.RowCount = de_informa.rsSel_ABCOcorr.RecordCount
        End If
        If de_informa.rsSel_ABCOcorr.RecordCount > 0 Then
            de_informa.rsSel_ABCOcorr.MoveFirst
            For xseries = 1 To frmAnOcorr.GrafOcorrABC.RowCount
                If xseries = 15 Then
                    xoutros = 0
                    Do Until de_informa.rsSel_ABCOcorr.EOF
                        xoutros = xoutros + de_informa.rsSel_ABCOcorr.Fields("qtdtot")
                        de_informa.rsSel_ABCOcorr.MoveNext
                    Loop
                    frmAnOcorr.GrafOcorrABC.Row = xseries
                    frmAnOcorr.GrafOcorrABC.Data = xoutros
                    frmAnOcorr.GrafOcorrABC.RowLabel = "xx"
                Else
                    frmAnOcorr.GrafOcorrABC.Row = xseries
                    frmAnOcorr.GrafOcorrABC.Data = de_informa.rsSel_ABCOcorr.Fields("qtdtot")
                    frmAnOcorr.GrafOcorrABC.RowLabel = de_informa.rsSel_ABCOcorr.Fields("cod_ocorr")
                    de_informa.rsSel_ABCOcorr.MoveNext
                End If
            Next
            frmAnOcorr.GrafOcorrABC.Visible = True
            de_informa.rsSel_ABCOcorr.MoveFirst
        Else
            frmAnOcorr.GrafOcorrABC.Visible = False
        End If
        

'EXIBE O FORM
 
 Me.Hide
 
    frmEscCliPer.Height = 7125
    fraConsCli.Visible = True
    fraDados.Enabled = True
    fraMensagem.Visible = False
    Animation1.Close
    DoEvents
 
 frmAnOcorr.Show

End Sub
Private Sub proc_dadosmutuo()
    Dim xcgc As String, xcontreg As Long, I As Long
    Dim xTotValmerc As Currency, xTotCtc As Long, xTotPeso As Currency, xTotVol As Long
    Dim xTotFreteCobr As Currency, xTotFretePago As Currency, xTotReceita As Currency
    
    'trata aba POR UF
    
    If optAir.Value = True Then Exit Sub
    
    If chkTodosCli.Value = 0 Then  'escolha de um CGC
        xcgc = Trim(txtCGCCli.Text) & "%"
    ElseIf chkTodosCli.Value = 1 Then  'todos os clientes
        xcgc = "%"
    End If
    
    If de_informa.rsSel_AnEstatSubPorUF.State = 1 Then de_informa.rsSel_AnEstatSubPorUF.Close
    de_informa.sel_anestatsubporuf CDate(mskPer1), CDate(mskPer2), xcgc
    
    If de_informa.rsSel_AnEstatSubPorUF.RecordCount > 0 Then
    
        Set frmAnEstat.FlexMutuoPorUF.DataSource = de_informa
        frmAnEstat.FlexMutuoPorUF.DataMember = "Sel_AnEstatSubPorUF"
        frmAnEstat.FlexMutuoPorUF.Refresh
        
        frmAnEstat.FlexMutuoPorUF.ColWidth(0) = 300
        frmAnEstat.FlexMutuoPorUF.ColWidth(1) = 1350
        frmAnEstat.FlexMutuoPorUF.ColWidth(2) = 1350
        frmAnEstat.FlexMutuoPorUF.ColWidth(3) = 350
        frmAnEstat.FlexMutuoPorUF.ColWidth(4) = 1000
        frmAnEstat.FlexMutuoPorUF.ColWidth(5) = 1400
        frmAnEstat.FlexMutuoPorUF.ColWidth(6) = 1300
        frmAnEstat.FlexMutuoPorUF.ColWidth(7) = 1300
        frmAnEstat.FlexMutuoPorUF.ColWidth(8) = 1000
        frmAnEstat.FlexMutuoPorUF.ColWidth(9) = 800
        frmAnEstat.FlexMutuoPorUF.ColWidth(10) = 800
        
        frmAnEstat.FlexMutuoPorUF.Row = 0
        frmAnEstat.FlexMutuoPorUF.ColAlignmentFixed = 4
        frmAnEstat.FlexMutuoPorUF.Col = 1

        frmAnEstat.FlexMutuoPorUF.Text = "T.SubContratada"
        frmAnEstat.FlexMutuoPorUF.Col = 2
        frmAnEstat.FlexMutuoPorUF.Text = "Região"
        frmAnEstat.FlexMutuoPorUF.Col = 3
        frmAnEstat.FlexMutuoPorUF.Text = "UF"
        frmAnEstat.FlexMutuoPorUF.Col = 4
        frmAnEstat.FlexMutuoPorUF.Text = "Expedições"
        frmAnEstat.FlexMutuoPorUF.Col = 5
        frmAnEstat.FlexMutuoPorUF.Text = "Val. Mercadoria"
        frmAnEstat.FlexMutuoPorUF.Col = 6
        frmAnEstat.FlexMutuoPorUF.Text = "Frete Cobrado"
        frmAnEstat.FlexMutuoPorUF.Col = 7
        frmAnEstat.FlexMutuoPorUF.Text = "Frete Pg. Sub."
        frmAnEstat.FlexMutuoPorUF.Col = 8
        frmAnEstat.FlexMutuoPorUF.Text = "Receita"
        frmAnEstat.FlexMutuoPorUF.Col = 9
        frmAnEstat.FlexMutuoPorUF.Text = "Peso"
        frmAnEstat.FlexMutuoPorUF.ColAlignment(9) = 7
        frmAnEstat.FlexMutuoPorUF.Col = 10
        frmAnEstat.FlexMutuoPorUF.Text = "Volumes"
        frmAnEstat.FlexMutuoPorUF.ColAlignment(10) = 7

        'formatando as células
        
        For I = 1 To frmAnEstat.FlexMutuoPorUF.Rows - 1
            frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 4) = Format(Val(frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 4)), "###,##0")
            frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 5) = Format(Val(frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 5)), "#,###,###,##0.00")
            frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 6) = Format(Val(frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 6)), "##,###,##0.00")
            frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 7) = Format(Val(frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 7)), "##,###,##0.00")
            frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 8) = Format(Val(frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 8)), "##,###,##0.00")
            frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 9) = Format(Val(frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 9)), "###,##0.0")
            frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 10) = Format(Val(frmAnEstat.FlexMutuoPorUF.TextMatrix(I, 10)), "###,##0")
        Next

        'trata aba por resumo
        
        If de_informa.rsSel_AnEstatSubGrid.State = 1 Then de_informa.rsSel_AnEstatSubGrid.Close
        de_informa.sel_AnEstatSubGrid CDate(mskPer1), CDate(mskPer2), xcgc
        
        Set frmAnEstat.flexMutuoResumo.DataSource = de_informa
        frmAnEstat.flexMutuoResumo.DataMember = "sel_AnEstatSubGrid"
        frmAnEstat.flexMutuoResumo.Refresh
        
        frmAnEstat.flexMutuoResumo.ColWidth(0) = 300
        frmAnEstat.flexMutuoResumo.ColWidth(1) = 1300
        frmAnEstat.flexMutuoResumo.ColWidth(2) = 1130
        frmAnEstat.flexMutuoResumo.ColWidth(3) = 1130
        frmAnEstat.flexMutuoResumo.ColWidth(4) = 930
        frmAnEstat.flexMutuoResumo.ColWidth(5) = 1
        frmAnEstat.flexMutuoResumo.ColWidth(6) = 1
        frmAnEstat.flexMutuoResumo.ColWidth(7) = 1
        frmAnEstat.flexMutuoResumo.ColWidth(8) = 1
        
        frmAnEstat.flexMutuoResumo.Row = 0
        frmAnEstat.flexMutuoResumo.Col = 1
        frmAnEstat.flexMutuoResumo.Text = "T.SubContratada"
        frmAnEstat.flexMutuoResumo.Col = 2
        frmAnEstat.flexMutuoResumo.Text = "Frete Cobrado"
        frmAnEstat.flexMutuoResumo.Col = 3
        frmAnEstat.flexMutuoResumo.Text = "Frete Pg. Sub."
        frmAnEstat.flexMutuoResumo.Col = 4
        frmAnEstat.flexMutuoResumo.Text = "Receita"
        
        'formatando as células
         frmAnEstat.flexMutuoResumo.ColAlignmentFixed = 4
        For I = 1 To frmAnEstat.flexMutuoResumo.Rows - 1
            frmAnEstat.flexMutuoResumo.TextMatrix(I, 2) = Format(Val(frmAnEstat.flexMutuoResumo.TextMatrix(I, 2)), "##,###,##0.00")
            frmAnEstat.flexMutuoResumo.TextMatrix(I, 3) = Format(Val(frmAnEstat.flexMutuoResumo.TextMatrix(I, 3)), "##,###,##0.00")
            frmAnEstat.flexMutuoResumo.TextMatrix(I, 4) = Format(Val(frmAnEstat.flexMutuoResumo.TextMatrix(I, 4)), "##,###,##0.00")
        Next
        
        de_informa.rsSel_AnEstatSubGrid.MoveFirst
        
        xTotValmerc = 0
        xTotCtc = 0
        xTotPeso = 0
        xTotVol = 0
        xTotFreteCobr = 0
        xTotFretePago = 0
        xTotReceita = 0
        
        Do Until de_informa.rsSel_AnEstatSubGrid.EOF
            xTotValmerc = xTotValmerc + de_informa.rsSel_AnEstatSubGrid.Fields("tvalmerc")
            xTotCtc = xTotCtc + de_informa.rsSel_AnEstatSubGrid.Fields("qtd")
            xTotPeso = xTotPeso + de_informa.rsSel_AnEstatSubGrid.Fields("tpeso")
            xTotVol = xTotVol + de_informa.rsSel_AnEstatSubGrid.Fields("tvolumes")
            xTotFreteCobr = xTotFreteCobr + de_informa.rsSel_AnEstatSubGrid.Fields("tfretetotal")
            xTotFretePago = xTotFretePago + de_informa.rsSel_AnEstatSubGrid.Fields("tfretepago")
            xTotReceita = xTotReceita + de_informa.rsSel_AnEstatSubGrid.Fields("receita")
            de_informa.rsSel_AnEstatSubGrid.MoveNext
        Loop
        
        frmAnEstat.lblTotValMerc = Format(xTotValmerc, "#,###,###,##0.00")
        frmAnEstat.lblTotCtc = Format(xTotCtc, "###,##0")
        frmAnEstat.lblTotPeso = Format(xTotPeso, "###,##0.0")
        frmAnEstat.lblTotVol = Format(xTotVol, "###,##0")
        frmAnEstat.lblTotFreteCobr = Format(xTotFreteCobr, "##,###,##0.00")
        frmAnEstat.lblTotFretePago = Format(xTotFretePago, "##,###,##0.00")
        frmAnEstat.lblTotReceita = Format(xTotReceita, "##,###,##0.00")
        
        frmAnEstat.lblTotPercCobrPago = Format(xTotFretePago / xTotFreteCobr, "##0.00%")
        frmAnEstat.lblTotFreteCobrValor = Format(xTotFreteCobr / xTotValmerc, "##0.000%")
        frmAnEstat.lblTotFretePagoValor = Format(xTotFretePago / xTotValmerc, "##0.000%")
        frmAnEstat.lblTotReceitaValor = Format(xTotReceita / xTotValmerc, "##0.000%")
        frmAnEstat.lblTotReceitaporCTC = Format(xTotReceita / xTotCtc, "###,##0.00")
        
        'dados do gráfico
        
        frmAnEstat.grafReceitaTotal.Column = 1
        frmAnEstat.grafReceitaTotal.Data = xTotReceita
        frmAnEstat.grafReceitaTotal.ColumnLabel = "Receita"
        frmAnEstat.grafReceitaTotal.Column = 2
        frmAnEstat.grafReceitaTotal.Data = xTotFretePago
        frmAnEstat.grafReceitaTotal.ColumnLabel = "Fr.Pago"
        
        frmAnEstat.grafReceita.ColumnCount = de_informa.rsSel_AnEstatSubGrid.RecordCount
        
        frmAnEstat.grafValmerc.ColumnCount = de_informa.rsSel_AnEstatSubGrid.RecordCount
        
        de_informa.rsSel_AnEstatSubGrid.MoveFirst
        
        For xcontreg = 1 To de_informa.rsSel_AnEstatSubGrid.RecordCount
            frmAnEstat.grafReceita.Column = xcontreg
            frmAnEstat.grafValmerc.Column = xcontreg
            frmAnEstat.grafReceita.Data = de_informa.rsSel_AnEstatSubGrid.Fields("receita")
            frmAnEstat.grafReceita.ColumnLabel = LCase(Mid$(de_informa.rsSel_AnEstatSubGrid.Fields("transp_sub"), 1, 8))
            frmAnEstat.grafValmerc.Data = de_informa.rsSel_AnEstatSubGrid.Fields("tvalmerc")
            frmAnEstat.grafValmerc.ColumnLabel = LCase(Mid$(de_informa.rsSel_AnEstatSubGrid.Fields("transp_sub"), 1, 8))
            de_informa.rsSel_AnEstatSubGrid.MoveNext
        Next
    End If
End Sub

Private Sub TxtFilial_GotFocus()
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub

Private Sub txtfilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
