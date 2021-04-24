VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_formata_edi 
   Caption         =   "LAYOUT DE EDI´S"
   ClientHeight    =   6960
   ClientLeft      =   2400
   ClientTop       =   1770
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   8820
   Begin VB.Frame Frame1 
      Height          =   6915
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      Begin VB.Frame Frame2 
         Height          =   3390
         Left            =   75
         TabIndex        =   11
         Top             =   150
         Width           =   8565
         Begin VB.Frame Frame3 
            Height          =   840
            Left            =   75
            TabIndex        =   19
            Top             =   1200
            Width           =   6015
            Begin VB.TextBox TXT_TEXTO 
               Height          =   285
               Left            =   3300
               TabIndex        =   6
               Top             =   450
               Width           =   2640
            End
            Begin VB.ComboBox cmb_campos 
               Height          =   315
               ItemData        =   "frm_formata_edi.frx":0000
               Left            =   75
               List            =   "frm_formata_edi.frx":01A8
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   450
               Width           =   2640
            End
            Begin VB.Label Label9 
               Caption         =   "TEXTO:"
               Height          =   165
               Left            =   3300
               TabIndex        =   22
               Top             =   150
               Width           =   615
            End
            Begin VB.Label Label8 
               Caption         =   "ou"
               Height          =   165
               Left            =   2925
               TabIndex        =   21
               Top             =   300
               Width           =   240
            End
            Begin VB.Label Label4 
               Caption         =   "NOME DO CAMPO:"
               Height          =   240
               Left            =   75
               TabIndex        =   20
               Top             =   150
               Width           =   1440
            End
         End
         Begin VB.TextBox txt_arquivoNome 
            Height          =   285
            Left            =   1200
            TabIndex        =   1
            Top             =   225
            Width           =   840
         End
         Begin VB.CommandButton cmd_sair 
            Caption         =   "&Sair"
            Height          =   390
            Left            =   6600
            TabIndex        =   9
            Top             =   2100
            Width           =   1815
         End
         Begin VB.CommandButton cmg_Gravar 
            Caption         =   "&Gravar"
            Height          =   390
            Left            =   6600
            TabIndex        =   8
            Top             =   1650
            Width           =   1815
         End
         Begin VB.TextBox txt_reg 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3375
            TabIndex        =   2
            Top             =   225
            Width           =   840
         End
         Begin VB.TextBox TXT_descr 
            Height          =   765
            Left            =   75
            TabIndex        =   7
            Top             =   2475
            Width           =   6015
         End
         Begin VB.TextBox txt_pesoate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3375
            TabIndex        =   4
            Text            =   "0"
            Top             =   825
            Width           =   840
         End
         Begin VB.TextBox txt_pesode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Text            =   "0"
            Top             =   825
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "ARQUIVO:"
            Height          =   240
            Left            =   75
            TabIndex        =   18
            Top             =   225
            Width           =   1065
         End
         Begin VB.Image Image1 
            Height          =   915
            Left            =   6525
            Top             =   225
            Width           =   1890
         End
         Begin VB.Label Label7 
            Caption         =   "REGISTRO:"
            Height          =   240
            Left            =   2175
            TabIndex        =   17
            Top             =   225
            Width           =   1215
         End
         Begin VB.Label lab_qtdChar 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   315
            Left            =   5325
            TabIndex        =   16
            Top             =   825
            Width           =   465
         End
         Begin VB.Label Label6 
            Caption         =   "QTD. CHAR:"
            Height          =   240
            Left            =   4350
            TabIndex        =   15
            Top             =   825
            Width           =   990
         End
         Begin VB.Label Label5 
            Caption         =   "DESCRIÇÃO:"
            Height          =   240
            Left            =   75
            TabIndex        =   14
            Top             =   2175
            Width           =   990
         End
         Begin VB.Label Label3 
            Caption         =   "POSICAO ATE:"
            Height          =   240
            Left            =   2175
            TabIndex        =   13
            Top             =   825
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "POSICAO DE:"
            Height          =   240
            Left            =   75
            TabIndex        =   12
            Top             =   825
            Width           =   1065
         End
      End
      Begin MSDataGridLib.DataGrid grd_Edi 
         Bindings        =   "frm_formata_edi.frx":0805
         Height          =   3165
         Left            =   75
         TabIndex        =   10
         Top             =   3675
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   5583
         _Version        =   393216
         BackColor       =   12648447
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
         DataMember      =   "Sel_Edi"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "REGISTRO"
            Caption         =   "REGISTRO"
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
            DataField       =   "POSICAODE"
            Caption         =   "POSICAODE"
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
            DataField       =   "POSICAOATE"
            Caption         =   "POSICAOATE"
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
            DataField       =   "QTD_CHAR"
            Caption         =   "QTD_CHAR"
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
            DataField       =   "CAMPO"
            Caption         =   "CAMPO"
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
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_formata_edi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub cmg_Gravar_Click()
Dim xcampo As String

If cmb_campos.ListIndex = -1 Then
    xcampo = TXT_TEXTO.Text
Else
    xcampo = cmb_campos.Text
End If

    

deb_edi.In_Edi txt_arquivoNome.Text, Int(txt_pesode.Text), Int(txt_pesoate.Text), Int(lab_qtdChar.Caption), xcampo, _
                TXT_descr.Text, txt_reg.Text

txt_pesode.Text = Int(txt_pesoate.Text) + 1
txt_pesoate.Text = txt_pesode.Text
TXT_descr.Text = Empty
TXT_TEXTO.Text = Empty
cmb_campos.ListIndex = -1


txt_pesoate.SelStart = 0
txt_pesoate.SelLength = Len(txt_pesoate.Text)
txt_pesoate.SetFocus

If deb_edi.rsSel_Edi.State = 1 Then deb_edi.rsSel_Edi.Close
    deb_edi.Sel_Edi txt_arquivoNome.Text
    
    grd_Edi.DataMember = "Sel_Edi"
    grd_Edi.Refresh



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift _
      As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txt_pesoate_GotFocus()
txt_pesoate.SelStart = 0
txt_pesoate.SelLength = Len(txt_pesoate.Text)
txt_pesoate.SetFocus
End Sub

Private Sub txt_pesoate_LostFocus()

    lab_qtdChar.Caption = (Val(txt_pesoate.Text) - Val(txt_pesode.Text)) + 1

End Sub

Private Sub txt_pesode_GotFocus()
txt_pesode.SelStart = 0
txt_pesode.SelLength = Len(txt_pesode.Text)
txt_pesode.SetFocus

End Sub
