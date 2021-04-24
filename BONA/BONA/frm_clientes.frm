VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_clientes 
   Caption         =   "CLIENTES"
   ClientHeight    =   3045
   ClientLeft      =   4095
   ClientTop       =   3420
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   6480
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&Sair"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   2520
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid grd_cliente 
         Bindings        =   "frm_clientes.frx":0000
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3836
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
         DataMember      =   "sel_busca_cli"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "CGC"
            Caption         =   "CGC"
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
            DataField       =   "CLIENTE"
            Caption         =   "CLIENTE"
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
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub grd_cliente_DblClick()
Dim cgc As String

frm_diversos.lab_nome.Caption = deb_bona.rssel_busca_cli.Fields("CLIENTE")

cgc = Trim(deb_bona.rssel_busca_cli.Fields("CGC"))

frm_diversos.lab_cgc.Caption = cgc
frm_diversos.cmd_protocolos.Visible = True


Unload Me




End Sub
