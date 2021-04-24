VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLogUsuarios 
   Caption         =   "LOG de USUÁRIOS"
   ClientHeight    =   6270
   ClientLeft      =   840
   ClientTop       =   1440
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   9240
   Begin VB.CommandButton cmdImpr 
      Caption         =   "Imprimir ..."
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período..."
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   240
      Width           =   3015
      Begin MSMask.MaskEdBox mskPer2 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
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
         Left            =   120
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
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Text            =   "%"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtAcao 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "%"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "Atualizar"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid gridLog 
      Bindings        =   "frmLogUsuarios.frx":0000
      Height          =   5055
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8916
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
      DataMember      =   "Sel_LogUsuarios"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "acao"
         Caption         =   "Ação"
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
         DataField       =   "data"
         Caption         =   "Data e Hora"
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
         DataField       =   "usuario"
         Caption         =   "Usuário"
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
         DataField       =   "descricao"
         Caption         =   "Descrição da Ação"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2069,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4020,095
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Usuário:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ação:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   420
   End
End
Attribute VB_Name = "frmLogUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualizar_Click()
    cmdAtualizar.Enabled = False
    cmdAtualizar.Caption = "AGUARDE ..."
    If de_informa.rsSel_LogUsuarios.State = 1 Then de_informa.rsSel_LogUsuarios.Close
    de_informa.Sel_LogUsuarios mskPer1.Text, mskPer2.Text & " 23:59:59", Trim$(txtUsuario), Trim$(txtAcao)
    If de_informa.rsSel_LogUsuarios.RecordCount < 1 Then
        MsgBox "Dados não Encontrados !"
    End If
    gridLog.DataMember = "sel_logusuarios"
    gridLog.Refresh
    cmdAtualizar.Enabled = True
    cmdAtualizar.Caption = "Atualizar"
End Sub

Private Sub cmdImpr_Click()
    Dim xlinha As Integer
    
    cmdImpr.Enabled = False
    cmdAtualizar.Enabled = False
    
    de_informa.rsSel_LogUsuarios.MoveFirst
    xlinha = 0
    
    Printer.FontName = "Courier New"
    
    Do Until de_informa.rsSel_LogUsuarios.EOF
        
        
        If xlinha = 0 Then
            Printer.Print
            Printer.Print
            Printer.FontBold = True
            Printer.FontUnderline = True
            Printer.Print Spc(2); "INTEC TRANSPORTES"
            Printer.FontUnderline = False
            Printer.Print
            Printer.Print Spc(2); "RELATÓRIO DE LOG DE USUÁRIOS"
            Printer.FontStrikethru = True
            Printer.Print Spc(2); String(230, " ")
            Printer.FontStrikethru = False
            Printer.Print Spc(2); "AÇÃO";
            Printer.Print Spc(8); "DATA/HORA";
            Printer.Print Spc(11); "USUÁRIO";
            Printer.Print Spc(4); "DESCRIÇÃO"
            Printer.FontStrikethru = True
            Printer.Print Spc(2); String(230, " ")
            Printer.FontStrikethru = False
            Printer.FontBold = False
            Printer.FontUnderline = False
        End If
        
        Printer.Print Spc(2); de_informa.rsSel_LogUsuarios.Fields("acao");
        Printer.Print Spc(12 - Len(de_informa.rsSel_LogUsuarios.Fields("acao"))); de_informa.rsSel_LogUsuarios.Fields("data");
        Printer.Print Spc(19 - Len(de_informa.rsSel_LogUsuarios.Fields("data"))); de_informa.rsSel_LogUsuarios.Fields("usuario");
        Printer.Print Spc(11 - Len(de_informa.rsSel_LogUsuarios.Fields("usuario"))); de_informa.rsSel_LogUsuarios.Fields("descricao")
        
        xlinha = xlinha + 1
        If xlinha > 70 Then
            xlinha = 0
            Printer.NewPage
        End If
        
        de_informa.rsSel_LogUsuarios.MoveNext
    
    Loop
        Printer.NewPage
        Printer.EndDoc
        
        MsgBox "RELATÓRIO ENVIADO PARA IMPRESSÃO"
        
        cmdImpr.Enabled = True
        cmdAtualizar.Enabled = True
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogUsuarios = Nothing
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
    End If
End Sub

Private Sub txtAcao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
