VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVLManutBase 
   Caption         =   "Manutenção do Base de Dados"
   ClientHeight    =   6735
   ClientLeft      =   2850
   ClientTop       =   1980
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9945
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cadastro de Clientes - Dados Faltantes"
      TabPicture(0)   =   "frmVLManutBase.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdVerificar"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "cmdSair"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Definição de Pacotes por Títulos"
      TabPicture(1)   =   "frmVLManutBase.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame5 
         Caption         =   "Consulta Pacotes"
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
         TabIndex        =   31
         Top             =   4200
         Width           =   9495
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   975
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1720
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1815
            Left            =   5280
            TabIndex        =   36
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   3201
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
         Begin VB.CommandButton Command1 
            Caption         =   "Buscar"
            Height          =   375
            Left            =   3840
            TabIndex        =   35
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1800
            TabIndex        =   34
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Por Título"
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Por Pacote"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   3735
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   9495
         Begin VB.CommandButton cmdBuscarTit 
            Caption         =   "Buscar Títulos Pendentes"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   4215
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "Incluir"
            Height          =   375
            Left            =   4320
            TabIndex        =   27
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton cmdExclue 
            Caption         =   "Excluir"
            Height          =   375
            Left            =   4320
            TabIndex        =   26
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox txtPacote 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   5640
            TabIndex        =   24
            Top             =   360
            Width           =   3615
         End
         Begin VB.CommandButton cmdGravarPacote 
            Caption         =   "Gravar Pacote"
            Height          =   855
            Left            =   4320
            TabIndex        =   23
            Top             =   1800
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid flexPacotes 
            Height          =   2775
            Left            =   5280
            TabIndex        =   25
            Top             =   840
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4895
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid gridTitulos 
            Bindings        =   "frmVLManutBase.frx":0038
            Height          =   2775
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4895
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
            DataMember      =   "Sel_VLTitulosPend"
            ColumnCount     =   1
            BeginProperty Column00 
               DataField       =   "material"
               Caption         =   "material"
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
                  ColumnWidth     =   3630,047
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Pacote:"
            Height          =   195
            Left            =   4920
            TabIndex        =   30
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   -67080
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cliente"
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
         Left            =   -74880
         TabIndex        =   10
         Top             =   960
         Width           =   9495
         Begin VB.Frame Frame4 
            Caption         =   "Dados a Serem Atualizados"
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
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton cmdGravar 
               Caption         =   "Gravar Dados ..."
               Height          =   615
               Left            =   7440
               TabIndex        =   5
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtNomeFantasia 
               BackColor       =   &H00C0FFFF&
               Height          =   285
               Left            =   120
               MaxLength       =   30
               TabIndex        =   1
               Top             =   600
               Width           =   2895
            End
            Begin VB.Frame Frame3 
               Height          =   855
               Left            =   3240
               TabIndex        =   17
               Top             =   240
               Width           =   3975
               Begin VB.CommandButton cmdAtualizar 
                  Caption         =   "!"
                  Height          =   255
                  Left            =   90
                  TabIndex        =   2
                  Top             =   480
                  Width           =   330
               End
               Begin VB.ComboBox cmbGrupo 
                  Height          =   315
                  Left            =   480
                  TabIndex        =   3
                  Top             =   440
                  Width           =   1815
               End
               Begin VB.ComboBox cmbTipo 
                  Height          =   315
                  ItemData        =   "frmVLManutBase.frx":0051
                  Left            =   2400
                  List            =   "frmVLManutBase.frx":0053
                  TabIndex        =   4
                  Top             =   440
                  Width           =   1455
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Grupo:"
                  Height          =   195
                  Left            =   480
                  TabIndex        =   19
                  Top             =   165
                  Width           =   480
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo:"
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   18
                  Top             =   165
                  Width           =   360
               End
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Nome Fantasia:"
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   1110
            End
         End
         Begin VB.Label lblCnpj 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3360
            TabIndex        =   21
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
            Height          =   195
            Left            =   2760
            TabIndex        =   20
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblCliente 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   5640
            TabIndex        =   14
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label lblCodSap 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   5040
            TabIndex        =   12
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código SAP:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clientes com Dados Faltantes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74880
         TabIndex        =   9
         Top             =   3240
         Width           =   9495
         Begin MSDataGridLib.DataGrid gridFaltantes 
            Bindings        =   "frmVLManutBase.frx":0055
            Height          =   2775
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4895
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
            DataMember      =   "Sel_VLManutBasecli"
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "codclinf"
               Caption         =   "codclinf"
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
               DataField       =   "cgcclinf"
               Caption         =   "cgcclinf"
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
               DataField       =   "clientenf"
               Caption         =   "clientenf"
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
                  ColumnWidth     =   1260,284
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1574,929
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   5834,835
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "Verificar ..."
         Height          =   375
         Left            =   -74760
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmVLManutBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscarTit_Click()
    If de_informa.rsSel_VLTitulosPend.State = 1 Then de_informa.rsSel_VLTitulosPend.Close
    de_informa.Sel_VLTitulosPend
    
    gridTitulos.DataMember = "Sel_VLTitulosPend"
    gridTitulos.Refresh
    
End Sub

Private Sub Label4_Click()

End Sub

Private Sub cmbGrupo_Click()
    
    Me.MousePointer = 11
    DoEvents
    DoEvents
    
    'preenche o combo
    If de_informa.rsSel_VLBuscaTipoCli.State = 1 Then de_informa.rsSel_VLBuscaTipoCli.Close
    de_informa.Sel_VLBuscaTipoCli cmbGrupo.List(cmbGrupo.ListIndex)
    
    cmbTipo.Clear
    
    Do Until de_informa.rsSel_VLBuscaTipoCli.EOF
        cmbTipo.AddItem de_informa.rsSel_VLBuscaTipoCli.Fields("tipo")
        de_informa.rsSel_VLBuscaTipoCli.MoveNext
    Loop
    
    Me.MousePointer = 0
    DoEvents
    DoEvents
    

End Sub

Private Sub cmbGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub cmdAtualizar_Click()
    'preenche combo de grupo
    
    Me.MousePointer = 11
    DoEvents
    DoEvents
    
    If de_informa.rsSel_VLBuscaGrupoCli.State = 1 Then de_informa.rsSel_VLBuscaGrupoCli.Close
    de_informa.Sel_VLBuscaGrupoCli
    
    cmbGrupo.Clear
    
    Do Until de_informa.rsSel_VLBuscaGrupoCli.EOF
        cmbGrupo.AddItem de_informa.rsSel_VLBuscaGrupoCli.Fields("grupo")
        de_informa.rsSel_VLBuscaGrupoCli.MoveNext
    Loop
    
    Me.MousePointer = 0
    DoEvents
    DoEvents
    

End Sub

Private Sub cmdGravar_Click()
            
    If Len(Trim$(txtNomeFantasia)) < 2 Then
        MsgBox "Nome Fantasia Inválido !"
        txtNomeFantasia.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(cmbGrupo.List(cmbGrupo.ListIndex))) < 1 Then
        MsgBox "Você Deve Escolhe o Grupo e o Tipo do Cliente !"
        cmbGrupo.SetFocus
        Exit Sub
    End If
    
    If Len(Trim$(cmbTipo.List(cmbTipo.ListIndex))) < 1 Then
        MsgBox "Você Deve Escolhe o Grupo e o Tipo do Cliente !"
        cmbGrupo.SetFocus
        Exit Sub
    End If
    
    de_informa.Alt_VLManutBasecli Trim$(txtNomeFantasia), cmbGrupo.List(cmbGrupo.ListIndex), _
                                 cmbTipo.List(cmbTipo.ListIndex), Trim$(lblCodSap)
                                 
    de_informa.Alt_VLManutVLClientes Trim$(txtNomeFantasia), cmbGrupo.List(cmbGrupo.ListIndex), _
                                 cmbTipo.List(cmbTipo.ListIndex), Trim$(lblCodSap)
                                 
    MsgBox "Dados Gravados !"
    
    txtNomeFantasia = ""
    lblCodSap = ""
    lblCliente = ""
    lblCnpj = ""
    cmbGrupo.ListIndex = -1
    cmbTipo.ListIndex = -1
    
    cmdVerificar_Click
                                 
End Sub

Private Sub cmdIncluir_Click()
    flexPacotes.Rows = flexPacotes.Rows + 1
    flexPacotes.TextMatrix(flexPacotes.Rows - 2, 1) = gridTitulos.Columns(0)
        
    
    
    
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVerificar_Click()

    Me.MousePointer = 11
    DoEvents

    If de_informa.rsSel_VLManutBasecli.State = 1 Then de_informa.rsSel_VLManutBasecli.Close
    de_informa.Sel_VLManutBasecli
    
    gridFaltantes.DataMember = "sel_vlmanutbasecli"
    gridFaltantes.Refresh
    
    If de_informa.rsSel_VLManutBasecli.RecordCount < 1 Then
        MsgBox "Não Há Clientes com Dados Faltantes !"
        gridFaltantes.Enabled = False
        cmdSair.SetFocus
    Else
        gridFaltantes.Enabled = True
        gridFaltantes_Click
        txtNomeFantasia.SetFocus
    End If
    
    Me.MousePointer = 0
    DoEvents
    
End Sub

Private Sub Form_Load()
    gridFaltantes.DataMember = ""
    gridFaltantes.Refresh
    gridTitulos.DataMember = ""
    gridTitulos.Refresh
    flexPacotes.Cols = 2
    flexPacotes.TextMatrix(0, 1) = "Título / Material"
    flexPacotes.ColWidth(0) = 200
    flexPacotes.ColWidth(1) = 2000
End Sub

Private Sub gridFaltantes_Click()
    lblCodSap = gridFaltantes.Columns(0)
    lblCnpj = gridFaltantes.Columns(1)
    lblCliente = gridFaltantes.Columns(2)
End Sub

Private Sub txtNomeFantasia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtNomeFantasia_LostFocus()
    txtNomeFantasia = UCase(txtNomeFantasia)
End Sub
