VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_cadExJucos 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   990
   ClientTop       =   1725
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   11415
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "CADASTRO DE EX - JUCOS"
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6495
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   11456
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
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11040
         Y1              =   1440
         Y2              =   1440
      End
   End
End
Attribute VB_Name = "frm_cadExJucos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

