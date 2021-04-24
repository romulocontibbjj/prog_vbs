VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDevolução 
   Caption         =   "Controle de Devolução"
   ClientHeight    =   7815
   ClientLeft      =   810
   ClientTop       =   615
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10455
   Begin VB.Frame Frame5 
      Caption         =   "Devoluções Pendentes"
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
      TabIndex        =   27
      Top             =   3960
      Width           =   10095
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2143
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sair"
      Height          =   495
      Left            =   8880
      TabIndex        =   26
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Procurar..."
      Height          =   495
      Left            =   6000
      TabIndex        =   25
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Baixar Devolução"
      Height          =   495
      Left            =   7440
      TabIndex        =   24
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nova Devolução"
      Height          =   495
      Left            =   4560
      TabIndex        =   23
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Listagem Geral de Devoluções"
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
      TabIndex        =   21
      Top             =   5640
      Width           =   10095
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1215
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2143
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados da Devolução"
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
      TabIndex        =   7
      Top             =   960
      Width           =   6135
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4800
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2520
         TabIndex        =   17
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Previsão:"
         Height          =   195
         Left            =   6480
         TabIndex        =   13
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente que Está Devolvendo:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente  da NF:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nota Fiscal:"
         Height          =   195
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Coleta:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Finalização"
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
      Left            =   6480
      TabIndex        =   4
      Top             =   960
      Width           =   3735
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1200
         TabIndex        =   29
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Responsável:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Data de Envio Doctos Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data da Entrega:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
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
      Height          =   705
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
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
         Left            =   2025
         TabIndex        =   3
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   10095
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frmDevolução"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Set frmDevolução = Nothing
End Sub
