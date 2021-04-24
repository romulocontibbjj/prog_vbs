VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCadFeriados 
   Caption         =   "Cadastro de Feriados"
   ClientHeight    =   6825
   ClientLeft      =   1290
   ClientTop       =   1020
   ClientWidth     =   9270
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   9270
   Begin VB.Frame Frame1 
      Caption         =   "Emails Informando Feriados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   28
      Top             =   2880
      Width           =   9015
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5280
         TabIndex        =   36
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5280
         TabIndex        =   35
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   33
         Top             =   720
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Envio para os Endereços dos Cadastros de Clientes"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmCadFeriados.frx":0000
         Left            =   5760
         List            =   "frmCadFeriados.frx":0016
         TabIndex        =   29
         Text            =   "5"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Email 4:"
         Height          =   195
         Left            =   4680
         TabIndex        =   40
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Email 3:"
         Height          =   195
         Left            =   4680
         TabIndex        =   39
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Email 2:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Email 1:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Enviar com"
         Height          =   195
         Left            =   4800
         TabIndex        =   32
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dias de Antecedência"
         Height          =   195
         Left            =   6480
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraCalendar 
      Caption         =   "Calendário"
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
      Height          =   2655
      Left            =   5040
      TabIndex        =   26
      Top             =   120
      Width           =   2775
      Begin MSComCtl2.MonthView CalenFeriado 
         Height          =   2310
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         ShowToday       =   0   'False
         StartOfWeek     =   75366401
         CurrentDate     =   37300
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7920
      TabIndex        =   17
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Canc/Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "Alterar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "Novo"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraData 
      Caption         =   "Data / Descrição / Local"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   4815
      Begin VB.TextBox txtUf 
         Height          =   285
         Left            =   4320
         TabIndex        =   25
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtCidade 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   24
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtDia 
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtMes 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtAno 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descr."
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   3960
         TabIndex        =   15
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Dia:"
         Height          =   195
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mês:"
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ano:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame fraAbrangencia 
      Caption         =   "Abrangência"
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
      Height          =   975
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton optMunicipal 
         Caption         =   "Municipal"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optEstadual 
         Caption         =   "Estadual"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optNacional 
         Caption         =   "Nacional"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame fraTipo 
      Caption         =   "Tipo"
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
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton optVariavel 
         Caption         =   "Variável - Dia/Mês diferentes anualmente"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton optFixo 
         Caption         =   "Fixo - Mesmo Dia/Mês todos os Anos"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   3135
      End
   End
   Begin MSDataGridLib.DataGrid GridFeriado 
      Bindings        =   "frmCadFeriados.frx":0031
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   8388608
      ForeColor       =   8454143
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
      DataMember      =   "Sel_CadFeriado"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ano"
         Caption         =   "ano"
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
         DataField       =   "mes"
         Caption         =   "mes"
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
         DataField       =   "dia"
         Caption         =   "dia"
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
         DataField       =   "descricao"
         Caption         =   "descricao"
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
      BeginProperty Column05 
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
      BeginProperty Column06 
         DataField       =   "tipo"
         Caption         =   "tipo"
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
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   390,047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3509,858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   374,74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2759,811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   390,047
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CalenFeriado_DateClick(ByVal DateClicked As Date)
    txtAno = CalenFeriado.Year
    txtMes = CalenFeriado.Month
    txtDia = CalenFeriado.Day
    TxtDescricao.SetFocus
End Sub
Private Sub cmdAlterar_Click()
    If Mid$(xdireitos, 7, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        fraAbrangencia.Enabled = True
        fraData.Enabled = True
        fraTipo.Enabled = True
        fraCalendar.Enabled = True
        GridFeriado.Enabled = False
        cmdNovo.Enabled = False
        cmdAlterar.Enabled = False
        cmdGravar.Enabled = True
        'txtCidade.BackColor = &HC0FFFF
        txtUf.BackColor = &HC0FFFF
        txtAno.BackColor = &HC0FFFF
        txtMes.BackColor = &HC0FFFF
        txtDia.BackColor = &HC0FFFF
        TxtDescricao.BackColor = &HC0FFFF
        frmCadFeriados.Caption = "Cadastro de Feriados - Alterar"
    End If
End Sub
Private Sub cmdGravar_Click()
    Dim xano, xTipo As String
    If frmCadFeriados.Caption = "Cadastro de Feriados - Novo" Then
        xano = Val(txtAno)
        If optFixo.Value = True Then
            xTipo = "F"
        Else
            xTipo = "V"
        End If
        de_informa.ins_CadFeriado xano, Val(txtMes), Val(txtDia), TxtDescricao, txtCidade, txtUf, xTipo
        If de_informa.rsSel_CadFeriado.State = 1 Then de_informa.rsSel_CadFeriado.Close
        de_informa.Sel_CadFeriado
        GridFeriado.DataMember = "sel_cadferiado"
        GridFeriado.Refresh
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "CAD. DE FERIADOS: " & TxtDescricao
        
    Else
        'alterar
        xano = Val(txtAno)
        If optFixo.Value = True Then
            xTipo = "F"
        Else
            xTipo = "V"
        End If
        de_informa.alt_cadferiado de_informa.rsSel_CadFeriado.Fields("codigo"), xano, Val(txtMes), Val(txtDia), TxtDescricao, txtCidade, txtUf, xTipo
        If de_informa.rsSel_CadFeriado.State = 1 Then de_informa.rsSel_CadFeriado.Close
        de_informa.Sel_CadFeriado
        GridFeriado.DataMember = "sel_cadferiado"
        GridFeriado.Refresh
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "ALTERAÇÃO", xusuario, "CAD. DE FERIADOS: " & TxtDescricao
        
    End If
        'frmCadFeriados.Caption = "Cadastro de Feriados"
        cmdSair_Click
End Sub
Private Sub cmdNovo_Click()
    If Mid$(xdireitos, 7, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        fraAbrangencia.Enabled = True
        fraData.Enabled = True
        fraTipo.Enabled = True
        fraCalendar.Enabled = True
        GridFeriado.Enabled = False
        cmdNovo.Enabled = False
        cmdAlterar.Enabled = False
        cmdGravar.Enabled = True
        txtAno.Text = ""
        txtMes.Text = ""
        txtDia.Text = ""
        TxtDescricao.Text = ""
        txtCidade.Text = ""
        txtUf.Text = ""
        txtAno.BackColor = &HC0FFFF
        txtMes.BackColor = &HC0FFFF
        txtDia.BackColor = &HC0FFFF
        TxtDescricao.BackColor = &HC0FFFF
        'txtCidade.BackColor = &HC0FFFF
        txtUf.BackColor = &HC0FFFF
        frmCadFeriados.Caption = "Cadastro de Feriados - Novo"
        txtAno.SetFocus
    End If
End Sub
Private Sub cmdSair_Click()
    If cmdGravar.Enabled = True Then
        txtDia = ""
        txtDia.BackColor = &H8000000E
        txtMes = ""
        txtMes.BackColor = &H8000000E
        txtAno = ""
        txtAno.BackColor = &H8000000E
        TxtDescricao = ""
        TxtDescricao.BackColor = &H8000000E
        txtCidade = ""
        txtCidade.BackColor = &H8000000E
        txtUf = ""
        txtUf.BackColor = &H8000000E
        fraAbrangencia.Enabled = False
        fraData.Enabled = False
        fraTipo.Enabled = False
        fraCalendar.Enabled = False
        GridFeriado.Enabled = True
        cmdNovo.Enabled = True
        cmdGravar.Enabled = False
        frmCadFeriados.Caption = "Cadastro de Feriados"
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Enabled = False
    mdiInforma.mnuArquivos.Enabled = False
    mdiInforma.mnuCad.Enabled = False
    mdiInforma.mnuProcesso.Enabled = False
    mdiInforma.mnuSair.Enabled = False
    mdiInforma.mnuInformacao.Enabled = False
    mdiInforma.mnuRelatorios.Enabled = False
    GridFeriado.DataMember = "sel_cadferiado"
    GridFeriado.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
    mdiInforma.mnuArquivos.Enabled = True
    mdiInforma.mnuCad.Enabled = True
    mdiInforma.mnuProcesso.Enabled = True
    mdiInforma.mnuSair.Enabled = True
    mdiInforma.mnuInformacao.Enabled = True
    mdiInforma.mnuRelatorios.Enabled = True
    Set frmCadFeriados = Nothing
End Sub

Private Sub GridFeriado_Click()
    Dim xdata As Date
    If GridFeriado.Columns(6) = "F" Then
        optFixo = True
    Else
        optVariavel = True
    End If
    If GridFeriado.Columns(4) = "BR" Then       'feriado nacional
        optNacional = True
    ElseIf GridFeriado.Columns(5) = "" Then     'feriado estadual
        optEstadual = True
    Else                                        'feriado municipal
        optMunicipal = True
    End If
    If GridFeriado.Columns(0) > 0 Then
        txtAno = GridFeriado.Columns(0)
    Else
        txtAno = ""
    End If
    txtMes = GridFeriado.Columns(1)
    txtDia = GridFeriado.Columns(2)
    txtCidade = GridFeriado.Columns(5)
    txtUf = GridFeriado.Columns(4)
    TxtDescricao = GridFeriado.Columns(3)
    If Val(txtAno) > 1000 Then
        xdata = CDate(txtDia & "/" & txtMes & "/" & txtAno)
    Else
        xdata = CDate(txtDia & "/" & txtMes & "/" & Year(datahora("data")))
    End If
    CalenFeriado.Value = xdata
    cmdAlterar.Enabled = True
    
    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "CONSULTA", xusuario, "CAD. DE FERIADOS: " & TxtDescricao
    
End Sub

Private Sub optEstadual_Click()
    txtCidade = ""
    txtUf = ""
    txtCidade.Enabled = False
    txtCidade.BackColor = &H8000000E
End Sub

Private Sub optFixo_Click()
    txtAno = ""
End Sub

Private Sub optMunicipal_Click()
    txtUf = ""
    txtCidade.Enabled = True
    txtCidade.BackColor = &HC0FFFF
End Sub

Private Sub optNacional_Click()
    txtCidade = ""
    txtUf = "BR"
    txtCidade.Enabled = False
    txtCidade.BackColor = &H8000000E
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAno_LostFocus()
    If optFixo = True Then
        txtAno = ""
    End If
End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCidade_LostFocus()
    txtCidade = UCase(txtCidade)
End Sub

Private Sub TxtDescricao_GotFocus()
    If optFixo = True Then
        txtAno = ""
    End If
End Sub

Private Sub TxtDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub TxtDescricao_LostFocus()
    TxtDescricao = UCase(TxtDescricao)
End Sub

Private Sub txtDia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtUf_GotFocus()
    If optNacional = True Then
        txtUf = "BR"
    End If
End Sub

Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtUF_LostFocus()
    txtUf = UCase(txtUf)
    If optNacional = True Then
        txtUf = "BR"
    End If
End Sub
