VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_festa 
   Caption         =   "CADASTRO DA FESTA DE 07 DE SETEMBRO DE 2004"
   ClientHeight    =   5580
   ClientLeft      =   4185
   ClientTop       =   3375
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7800
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox CMB_SEX 
         Height          =   315
         ItemData        =   "frm_festa.frx":0000
         Left            =   4440
         List            =   "frm_festa.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox cmb_pago 
         Height          =   315
         ItemData        =   "frm_festa.frx":0014
         Left            =   6720
         List            =   "frm_festa.frx":0021
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmd_ok 
         Caption         =   "OK"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TXT_CONVITE 
         Height          =   285
         Left            =   6720
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TXT_NOME 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.ComboBox cmb_convite 
         Height          =   315
         ItemData        =   "frm_festa.frx":002E
         Left            =   1440
         List            =   "frm_festa.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   4800
         Width           =   7335
         Begin VB.Label Label1 
            Caption         =   "QTD. CONV:"
            Height          =   255
            Left            =   5400
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lab_qtd_convite 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   6360
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid GRD_FESTA 
         Bindings        =   "frm_festa.frx":0063
         Height          =   3375
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5953
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
         DataMember      =   "sel_festa_convite"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "CONVITE_NUMERO"
            Caption         =   "CONVITE_NUMERO"
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
            DataField       =   "NOME"
            Caption         =   "NOME"
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
            DataField       =   "CONVIDADO"
            Caption         =   "CONVIDADO"
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
            DataField       =   "PAGO"
            Caption         =   "PAGO"
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
               ColumnWidth     =   1590,236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "SEXO:"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "ORGANIZADOR:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "CONVITE:"
         Height          =   255
         Left            =   5880
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "NOME:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         FillColor       =   &H00000080&
         Height          =   3615
         Left            =   120
         Top             =   1080
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frm_festa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ok_Click()

'If TXT_NOME.Text = Empty Then
  '  MsgBox "Digite o Nome", vbInformation, "NOME"
  '  TXT_NOME.SetFocus
'ElseIf TXT_CONVITE.Text = Empty Then
   '     MsgBox "Digite o Nº do Convite", vbInformation, "CONVITE"
    '    TXT_CONVITE.SetFocus
'Else

   ' deb_pend.in_festa TXT_NOME, cmb_convite.Text, TXT_CONVITE, cmb_pago,
    
'CADASTRO DE CLIENTES....

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
