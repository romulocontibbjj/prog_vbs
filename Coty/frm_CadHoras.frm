VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_CadHoras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Valores"
   ClientHeight    =   3060
   ClientLeft      =   4440
   ClientTop       =   4050
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6150
   Begin VB.Frame Frame3 
      Caption         =   "Histórico de Alteração "
      Height          =   1455
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   3855
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1931
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
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdgravarvh 
      Caption         =   "Gravar"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txt_valor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "0,00"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Horas / Km"
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
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton opt_KM 
            Caption         =   "Valor KM"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton opt_Hora 
            Caption         =   "Valor Hora"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Label Label1 
         Caption         =   "R$"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_cadHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub opt_Hora_Click()

If opt_Hora.Value = True Then
    
    Label1.Caption = "VALOR HORA:"
    
End If

End Sub

Private Sub opt_KM_Click()

If opt_KM.Value = True Then
    
        Label1.Caption = "VALOR KM:"

End If


End Sub
