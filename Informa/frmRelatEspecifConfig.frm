VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRelatEspecifConfig 
   Caption         =   "Configuração de Relatórios"
   ClientHeight    =   6990
   ClientLeft      =   1740
   ClientTop       =   1665
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   10815
   Begin VB.Frame Frame4 
      Height          =   2535
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Width           =   10575
      Begin VB.CommandButton cmdExcluirUsu 
         Caption         =   "----  Excluir Usuários  --->"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cmdIncluirUsu 
         Caption         =   "<---  Incluir Usuários  ----"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   28
         Top             =   960
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid gridUsuariosAutoriz 
         Bindings        =   "frmRelatEspecifConfig.frx":0000
         Height          =   1935
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3413
         _Version        =   393216
         Enabled         =   0   'False
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
         DataMember      =   "Sel_RelatUsu"
         ColumnCount     =   3
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
         BeginProperty Column02 
            DataField       =   "nome"
            Caption         =   "Nome Completo"
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
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2445,166
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid gridUsuariosDemais 
         Bindings        =   "frmRelatEspecifConfig.frx":0019
         Height          =   1935
         Left            =   6360
         TabIndex        =   30
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3413
         _Version        =   393216
         Enabled         =   0   'False
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
         DataMember      =   "Sel_RelatNaoUsu"
         ColumnCount     =   2
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "nome"
            Caption         =   "Nome Completo"
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
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2415,118
            EndProperty
         EndProperty
      End
      Begin VB.Label lblusu2 
         AutoSize        =   -1  'True
         Caption         =   "**   Demais Usuários do Sistema   **"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6840
         TabIndex        =   31
         Top             =   240
         Width           =   3075
      End
      Begin VB.Label lblusu1 
         AutoSize        =   -1  'True
         Caption         =   "**   Usuários do Relatório   **"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   2505
      End
   End
   Begin VB.Frame fraDados 
      Caption         =   "Dados do Relatório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   10575
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar Dados ..."
         Height          =   330
         Left            =   5640
         TabIndex        =   35
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox txtRelQuery 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   975
         Left            =   1560
         MaxLength       =   2000
         TabIndex        =   2
         Top             =   1080
         Width           =   8895
      End
      Begin VB.TextBox txtRelDescr 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   1
         Top             =   720
         Width           =   8895
      End
      Begin VB.TextBox txtRelNome 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblNumRelat 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1560
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblAcao 
         AutoSize        =   -1  'True
         Caption         =   "INCLUSÃO"
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
         Left            =   9000
         TabIndex        =   33
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Query SQL:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Relatório:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Parâmetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   2360
      Width           =   10575
      Begin VB.CommandButton cmdGravarRelat 
         Caption         =   "Gravar Relatório"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   32
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CheckBox chkParaPeriodo 
         Caption         =   "Parâmetro de Período"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtPara1 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara2 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8520
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara3 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara4 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara6 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara7 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara5 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8520
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtPara8 
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   8520
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblPara1 
         AutoSize        =   -1  'True
         Caption         =   "Parâmetro 1 ........:"
         Height          =   195
         Left            =   3600
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPara2 
         AutoSize        =   -1  'True
         Caption         =   "Parâmetro 2 ........:"
         Height          =   195
         Left            =   7080
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPara3 
         AutoSize        =   -1  'True
         Caption         =   "Parâmetro 3 ........:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPara4 
         AutoSize        =   -1  'True
         Caption         =   "Parâmetro 4 ........:"
         Height          =   195
         Left            =   3600
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPara6 
         AutoSize        =   -1  'True
         Caption         =   "Parâmetro 6 ........:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPara7 
         AutoSize        =   -1  'True
         Caption         =   "Parâmetro 7 ........:"
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPara5 
         AutoSize        =   -1  'True
         Caption         =   "Parâmetro 5 ........:"
         Height          =   195
         Left            =   7080
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblPara8 
         AutoSize        =   -1  'True
         Caption         =   "Parâmetro 8 ........:"
         Height          =   195
         Left            =   7080
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmRelatEspecifConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGravarRelat_Click()
    Dim xParaPeriodo As String
    
    If Len(Trim$(txtRelNome)) < 3 Then
        MsgBox "Nome de Relatório Inválido !!", vbCritical, "Ops"
        txtRelNome.SetFocus
        Exit Sub
    ElseIf Len(Trim$(txtRelDescr)) < 3 Then
        MsgBox "Descrição de Relatório Inválido !!", vbCritical, "Ops"
        txtRelDescr.SetFocus
        Exit Sub
    ElseIf Len(Trim$(txtRelQuery)) < 3 Then
        MsgBox "Query do Relatório Inválido !!", vbCritical, "Ops"
        txtRelQuery.SetFocus
        Exit Sub
    End If
    
    If chkParaPeriodo.Value = 1 Then
        xParaPeriodo = "S"
    Else
        xParaPeriodo = "N"
    End If

    If txtRelNome.Enabled = True Then
        'Inclusão do Relatório
        de_informa.Ins_Relatorios Trim$(txtRelNome), Trim$(txtRelDescr), Trim$(txtRelQuery), xParaPeriodo, _
                                  Trim$(txtPara1), Trim$(txtPara2), Trim$(txtPara3), Trim$(txtPara4), _
                                  Trim$(txtPara5), Trim$(txtPara6), Trim$(txtPara7), Trim$(txtPara8)
                                  
        If de_informa.rsSel_RelatUltimo.State = 1 Then de_informa.rsSel_RelatUltimo.Close
        de_informa.Sel_RelatUltimo
        
        If de_informa.rsSel_RelatoriosNumero.State = 1 Then de_informa.rsSel_RelatoriosNumero.Close
        de_informa.Sel_RelatoriosNumero de_informa.rsSel_RelatUltimo.Fields("num")
        
        lblNumRelat = de_informa.rsSel_RelatUltimo.Fields("num")
        
        fraDados.Enabled = False
        fraParametros.Enabled = False
                                  
        MsgBox "Relatório Incluído !! Cadastre Agora os Usuários Para Este Relatório.", vbInformation
        
        lblusu1.Enabled = True
        lblusu2.Enabled = True
        gridUsuariosAutoriz.Enabled = True
        gridUsuariosDemais.Enabled = True
        cmdExcluirUsu.Enabled = True
        cmdIncluirUsu.Enabled = True
        
    Else
        'Alteração do Relatório
        de_informa.Alt_Relatorios Trim$(txtRelDescr), Trim$(txtRelQuery), xParaPeriodo, _
                                  Trim$(txtPara1), Trim$(txtPara2), Trim$(txtPara3), Trim$(txtPara4), _
                                  Trim$(txtPara5), Trim$(txtPara6), Trim$(txtPara7), Trim$(txtPara8), lblNumRelat
                                  
        MsgBox "Dados Gravados !!", vbInformation
        
    End If
    

End Sub

Private Sub txtPara1_GotFocus()
    txtPara1.SelStart = 0
    txtPara1.SelLength = 20
End Sub

Private Sub txtPara1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara1_LostFocus()
    txtPara1 = UCase(txtPara1)
End Sub

Private Sub txtPara2_GotFocus()
    txtPara2.SelStart = 0
    txtPara2.SelLength = 20
End Sub

Private Sub txtPara2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara2_LostFocus()
    txtPara2 = UCase(txtPara2)
End Sub

Private Sub txtPara3_GotFocus()
    txtPara3.SelStart = 0
    txtPara3.SelLength = 20
End Sub

Private Sub txtPara3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara3_LostFocus()
    txtPara3 = UCase(txtPara3)
End Sub

Private Sub txtPara4_GotFocus()
    txtPara4.SelStart = 0
    txtPara4.SelLength = 20
End Sub

Private Sub txtPara4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara4_LostFocus()
    txtPara4 = UCase(txtPara4)
End Sub

Private Sub txtPara5_GotFocus()
    txtPara5.SelStart = 0
    txtPara5.SelLength = 20
End Sub

Private Sub txtPara5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara5_LostFocus()
    txtPara5 = UCase(txtPara5)
End Sub

Private Sub txtPara6_GotFocus()
    txtPara6.SelStart = 0
    txtPara6.SelLength = 20
End Sub

Private Sub txtPara6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara6_LostFocus()
    txtPara6 = UCase(txtPara6)
End Sub

Private Sub txtPara7_GotFocus()
    txtPara7.SelStart = 0
    txtPara7.SelLength = 20
End Sub

Private Sub txtPara7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara7_LostFocus()
    txtPara7 = UCase(txtPara7)
End Sub

Private Sub txtPara8_GotFocus()
    txtPara8.SelStart = 0
    txtPara8.SelLength = 20
End Sub

Private Sub txtPara8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtPara8_LostFocus()
    txtPara8 = UCase(txtPara8)
End Sub

Private Sub txtRelDescr_GotFocus()
    txtRelDescr.SelStart = 0
    txtRelDescr.SelLength = 100
End Sub

Private Sub txtRelDescr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRelDescr_LostFocus()
    txtRelDescr.Text = UCase(txtRelDescr)
End Sub

Private Sub txtRelNome_GotFocus()
    txtRelNome.SelStart = 0
    txtRelNome.SelLength = 20
End Sub

Private Sub txtRelNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRelNome_LostFocus()
    txtRelNome.Text = UCase(txtRelNome)
End Sub

Private Sub txtRelQuery_GotFocus()
    txtRelQuery.SelStart = 0
    txtRelQuery.SelLength = 2000
End Sub

Private Sub txtRelQuery_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

