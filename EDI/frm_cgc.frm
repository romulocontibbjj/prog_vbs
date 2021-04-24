VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_cgc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busca de Cgc´s"
   ClientHeight    =   3165
   ClientLeft      =   3885
   ClientTop       =   3810
   ClientWidth     =   6570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6570
   Begin VB.Frame Frame1 
      Height          =   3090
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6465
      Begin VB.CommandButton cmd_Buscar 
         Caption         =   "&Buscar"
         Height          =   240
         Left            =   5025
         TabIndex        =   2
         Top             =   225
         Width           =   840
      End
      Begin VB.TextBox txt_Cliente 
         Height          =   285
         Left            =   900
         TabIndex        =   1
         Top             =   225
         Width           =   3990
      End
      Begin MSDataGridLib.DataGrid grd_clientes 
         Bindings        =   "frm_cgc.frx":0000
         Height          =   2340
         Left            =   75
         TabIndex        =   3
         Top             =   675
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   4128
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
         DataMember      =   "Sel_Clientes"
         ColumnCount     =   4
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
         BeginProperty Column03 
            DataField       =   "uf"
            Caption         =   "uf"
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
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   540,284
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "CLIENTE:"
         Height          =   240
         Left            =   75
         TabIndex        =   4
         Top             =   225
         Width           =   765
      End
   End
End
Attribute VB_Name = "frm_cgc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Buscar_Click()

If deb_edi.rsSel_Clientes.State = 1 Then deb_edi.rsSel_Clientes.Close
    deb_edi.Sel_Clientes "%" & txt_Cliente.Text & "%"
    
    If deb_edi.rsSel_Clientes.RecordCount > 0 Then
    
        grd_clientes.DataMember = "Sel_Clientes"
        grd_clientes.Refresh
    Else
        MsgBox "Cliente Não Localizado", vbInformation, "CLIENTE"
    End If
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift _
      As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub grd_clientes_DblClick()
Dim xcgc As String

xcgc = Mid$(grd_clientes.Columns(0), 1, 8)

frm_CadEdis.txt_cgc.Text = xcgc
frm_CadEdis.txt_Cliente.Text = grd_clientes.Columns(1)

        grd_clientes.DataMember = ""
        grd_clientes.Refresh

Unload Me

End Sub
