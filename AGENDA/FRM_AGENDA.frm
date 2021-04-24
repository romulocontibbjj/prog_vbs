VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRM_AGENDA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENDÊNCIAS DIÁRIAS"
   ClientHeight    =   9270
   ClientLeft      =   4140
   ClientTop       =   1710
   ClientWidth     =   7215
   Icon            =   "FRM_AGENDA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   7215
   Begin VB.Frame frm_agenda 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmd_fone 
         Caption         =   "TELS"
         Height          =   255
         Left            =   5520
         TabIndex        =   26
         Top             =   720
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "DESCRIÇÃO"
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   3960
         Width           =   6135
         Begin VB.Label lab_cod 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2760
            TabIndex        =   25
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "COD:"
            Height          =   255
            Left            =   2280
            TabIndex        =   24
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lab_descr 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2760
            TabIndex        =   23
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label8 
            Caption         =   "DESCR.:"
            Height          =   255
            Left            =   2040
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lab_hora 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   720
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "HORA:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lab_data 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   720
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "DATA:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton cmd_calendario 
         Caption         =   "Dia"
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin MSComCtl2.MonthView calendario 
         Height          =   2370
         Left            =   2280
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   19136513
         CurrentDate     =   38226
      End
      Begin VB.TextBox TXT_DATA 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton CMD_LIMPAR 
         Caption         =   "LIMPAR TUDO"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   8040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid grd_fechados 
         Bindings        =   "FRM_AGENDA.frx":1272
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   5760
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3836
         _Version        =   393216
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   18
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
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "SEL_FECHADOS"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "DATA"
            Caption         =   "DATA"
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
            DataField       =   "HORA"
            Caption         =   "HORA"
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
            DataField       =   "DESCRICAO"
            Caption         =   "DESCRICAO"
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
            DataField       =   "FECHAMENTO"
            Caption         =   "FECHAMENTO"
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
            DataField       =   "DATA_FECH"
            Caption         =   "DATA_FECH"
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
               ColumnWidth     =   1739,906
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
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CMD_FECHAR 
         Caption         =   "FECHAR"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   5040
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid grd_abertos 
         Bindings        =   "FRM_AGENDA.frx":1289
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3836
         _Version        =   393216
         ForeColor       =   128
         HeadLines       =   1
         RowHeight       =   18
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
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "sel_tudo_aberto"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "DATA"
            Caption         =   "DATA"
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
            DataField       =   "HORA"
            Caption         =   "HORA"
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
            DataField       =   "DESCRICAO"
            Caption         =   "DESCRICAO"
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
            DataField       =   "CÓDIGO"
            Caption         =   "CÓDIGO"
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   915,024
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CMD_INSERIR 
         Caption         =   "&OK"
         Height          =   255
         Left            =   6480
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TXT_DESCR 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lab_fechados 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   6360
         TabIndex        =   14
         Top             =   8040
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "DATA:"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "AGENDA DIARIA"
         BeginProperty Font 
            Name            =   "BatmanForeverAlternate"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   1440
         TabIndex        =   11
         Top             =   8520
         Width           =   4575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "FECHADOS"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   5520
         Width           =   6975
      End
      Begin VB.Label LAB_QTD_ABERTOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   6360
         TabIndex        =   6
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ABERTOS"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   6975
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   7080
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7080
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "DESCRIÇÃO:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "FRM_AGENDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_calendario_Click()

If calendario.Visible = True Then
    calendario.Visible = False
Else
    calendario.Visible = True
End If



End Sub

Private Sub CMD_FECHAR_Click()
Dim xcod As Integer

xcod = deb_pend.rssel_tudo_aberto.Fields("CÓDIGO")

deb_pend.UP_FECHAR Date, Time, xcod

If deb_pend.rssel_tudo_aberto.State = 1 Then deb_pend.rssel_tudo_aberto.Close

deb_pend.rssel_tudo_aberto.Open
grd_abertos.DataMember = "SEL_TUDO_ABERTO"
grd_abertos.Refresh
LAB_QTD_ABERTOS.Caption = deb_pend.rssel_tudo_aberto.RecordCount

If deb_pend.rsSEL_FECHADOS.State = 1 Then deb_pend.rsSEL_FECHADOS.Close
    deb_pend.SEL_FECHADOS Date
    
    'deb_pend.rsSEL_FECHADOS.Open
    
    grd_fechados.DataMember = "sel_fechados"
    grd_fechados.Refresh
    
    lab_fechados.Caption = deb_pend.rsSEL_FECHADOS.RecordCount


End Sub

Private Sub cmd_fone_Click()
frm_agenda_fone.Show

DoEvents

End Sub

Private Sub cmd_inserir_Click()

If Trim(TXT_DESCR.Text) = Empty Then
    MsgBox "Digite a Descrição da Pendência", vbInformation, "DESCRIÇÃO"
    TXT_DESCR.SetFocus
    Exit Sub
    
Else

deb_pend.in_pend TXT_DATA, Time, TXT_DESCR.Text

TXT_DESCR.Text = Empty
TXT_DESCR.SetFocus

If deb_pend.rssel_tudo_aberto.State = 1 Then deb_pend.rssel_tudo_aberto.Close

deb_pend.rssel_tudo_aberto.Open
grd_abertos.DataMember = "SEL_TUDO_ABERTO"
grd_abertos.Refresh
LAB_QTD_ABERTOS.Caption = deb_pend.rssel_tudo_aberto.RecordCount

End If


End Sub

Private Sub CMD_LIMPAR_Click()
Dim XQTD As Integer

deb_pend.DEL_TUDO

XQTD = deb_pend.rsSEL_FECHADOS.RecordCount

If deb_pend.rsSEL_FECHADOS.State = 1 Then deb_pend.rsSEL_FECHADOS.Close
    deb_pend.SEL_FECHADOS Date
    
    'deb_pend.rsSEL_FECHADOS.Open
    
    grd_fechados.DataMember = "sel_fechados"
    grd_fechados.Refresh
    lab_fechados.Caption = deb_pend.rsSEL_FECHADOS.RecordCount

MsgBox XQTD & " DELETADOS", vbInformation, "LIMPANDO TABELA"

End Sub

Private Sub Command1_Click()
grd_abertos.DataMember = "sel_tudo_aberto"
grd_abertos.Refresh
End Sub

Private Sub Command2_Click()
grd_abertos.DataMember = ""
grd_abertos.Refresh
End Sub


Private Sub Form_Load()
TXT_DATA = Date
calendario.Value = Date
grd_abertos.DataMember = "sel_tudo_aberto"
LAB_QTD_ABERTOS.Caption = deb_pend.rssel_tudo_aberto.RecordCount

If deb_pend.rsSEL_FECHADOS.State = 1 Then deb_pend.rsSEL_FECHADOS.Close
    deb_pend.SEL_FECHADOS Date
    
    'deb_pend.rsSEL_FECHADOS.Open
    
    grd_fechados.DataMember = "sel_fechados"
    grd_fechados.Refresh
    
    lab_fechados.Caption = deb_pend.rsSEL_FECHADOS.RecordCount

    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub grd_abertos_Click()
lab_data.Caption = deb_pend.rssel_tudo_aberto.Fields("DATA")
lab_hora.Caption = deb_pend.rssel_tudo_aberto.Fields("HORA")
lab_descr.Caption = deb_pend.rssel_tudo_aberto.Fields("DESCRICAO")
lab_cod.Caption = deb_pend.rssel_tudo_aberto.Fields("CÓDIGO")

End Sub
