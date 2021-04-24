VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_fox 
   Caption         =   "Form1"
   ClientHeight    =   9630
   ClientLeft      =   1785
   ClientTop       =   1410
   ClientWidth     =   12960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   12960
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.Frame Frame4 
         Caption         =   "IMP - CLIENTES FOX (TB_VLCLIENTES)"
         Height          =   4935
         Left            =   120
         TabIndex        =   17
         Top             =   4560
         Width           =   7575
         Begin VB.CommandButton cmd_inserir_fox 
            Caption         =   "VL_CLIENTES"
            Height          =   255
            Left            =   5880
            TabIndex        =   23
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton cmd_tb_fox_imp 
            Caption         =   "TB_FOX_IMP"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1935
         End
         Begin MSDataGridLib.DataGrid grd_fox_imp 
            Bindings        =   "frm_fox.frx":0000
            Height          =   3735
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   6588
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "INDICE"
               Caption         =   "INDICE"
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
               DataField       =   "SERIE"
               Caption         =   "SERIE"
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
            BeginProperty Column03 
               DataField       =   "SAP"
               Caption         =   "SAP"
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
            BeginProperty Column05 
               DataField       =   "FANTASIA"
               Caption         =   "FANTASIA"
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
               DataField       =   "GRUPO"
               Caption         =   "GRUPO"
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
            BeginProperty Column07 
               DataField       =   "TIPO"
               Caption         =   "TIPO"
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
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.Label lab_vl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6600
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "QTD_VL:"
            Height          =   255
            Left            =   5880
            TabIndex        =   21
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lab_tb_fox_imp 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2160
            TabIndex        =   20
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.CommandButton cmd_sair 
         Caption         =   "SAIR"
         Height          =   375
         Left            =   11400
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "VIDEOLAR"
         Height          =   1095
         Left            =   7800
         TabIndex        =   11
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton cmd_atualiza_prazos 
            Caption         =   "PRAZOS"
            Height          =   255
            Left            =   840
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lab_vid 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lab_filialctc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   840
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Filial CTC:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "FOX - FILMES"
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7575
         Begin VB.CommandButton cmd_busca_cgc_branco 
            Caption         =   "Busca ""CGC"" = """""
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmd_altera_cgc 
            Caption         =   "ALTERA - CGC"
            Height          =   375
            Left            =   2520
            TabIndex        =   25
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSComctlLib.ProgressBar PRG_BASECLI 
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   3840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.CommandButton cmd_busca_clibrancos 
            Caption         =   "Buscar ""tb_basecli"""
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton cmd_alteracao 
            Caption         =   "Alterar ""GRUPO"" e ""TIPO"""
            Height          =   375
            Left            =   2520
            TabIndex        =   2
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSDataGridLib.DataGrid grd_clientes 
            Bindings        =   "frm_fox.frx":0017
            Height          =   2775
            Left            =   240
            TabIndex        =   4
            Top             =   960
            Width           =   4575
            _ExtentX        =   8070
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
            ColumnCount     =   1
            BeginProperty Column00 
               DataField       =   "CLIENTENF"
               Caption         =   "CLIENTENF"
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
                  ColumnWidth     =   4004,788
               EndProperty
            EndProperty
         End
         Begin VB.Label Label6 
            Caption         =   "QTD:"
            Height          =   255
            Left            =   3480
            TabIndex        =   10
            Top             =   3840
            Width           =   495
         End
         Begin VB.Label lab_qtd_basecli 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4080
            TabIndex        =   9
            Top             =   3840
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "TIPO:"
            Height          =   255
            Left            =   4920
            TabIndex        =   8
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "GRUPO:"
            Height          =   255
            Left            =   4920
            TabIndex        =   7
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4920
            TabIndex        =   6
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4920
            TabIndex        =   5
            Top             =   2280
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frm_fox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_altera_cgc_Click()
Dim xcliente    As String
Dim xgrupo      As String
Dim xtipo       As String
Dim xqtd        As Integer
Dim xprg        As Integer
Dim xcgc        As String
Dim xfantasia   As String

PRG_BASECLI.Min = 0
PRG_BASECLI.Max = deb_bona.rssel_cgc_branco.RecordCount
PRG_BASECLI.Value = PRG_BASECLI.Min
PRG_BASECLI.Visible = True
xcliente = deb_bona.rssel_cgc_branco.Fields("CLIENTENF")

deb_bona.rssel_cgc_branco.MoveFirst

Do Until deb_bona.rssel_cgc_branco.EOF
    xcliente = deb_bona.rssel_cgc_branco.Fields("CLIENTENF")
    
    If deb_bona.rssel_busca_grupo_tipo.State = 1 Then deb_bona.rssel_busca_grupo_tipo.Close
        deb_bona.sel_busca_grupo_tipo xcliente
        
        If deb_bona.rssel_busca_grupo_tipo.RecordCount > 0 Then
            xcgc = deb_bona.rssel_busca_grupo_tipo.Fields("CGC")
            xfantasia = deb_bona.rssel_busca_grupo_tipo.Fields("FANTASIA")
           
            xqtd = xqtd + 1
        
            deb_bona.up_cgc_base xcgc, xfantasia, xcliente
            
        End If
        
        
        deb_bona.rssel_cgc_branco.MoveNext
        
        xprg = xprg + 1
        
        PRG_BASECLI.Value = xprg
Loop
        
        MsgBox xqtd & " ATUALIZADOS"
         deb_bona.rssel_cgc_branco.Close
End Sub

Private Sub cmd_alteracao_Click()
Dim xcliente    As String
Dim xgrupo      As String
Dim xtipo       As String
Dim xqtd        As Integer
Dim xprg        As Integer
Dim xcgc        As String

PRG_BASECLI.Min = 0
PRG_BASECLI.Max = deb_bona.rssel_busca_cli_brancos.RecordCount
PRG_BASECLI.Value = PRG_BASECLI.Min
PRG_BASECLI.Visible = True
xcliente = deb_bona.rssel_busca_cli_brancos.Fields("CLIENTENF")

deb_bona.rssel_busca_cli_brancos.MoveFirst

Do Until deb_bona.rssel_busca_cli_brancos.EOF
    xcliente = deb_bona.rssel_busca_cli_brancos.Fields("CLIENTENF")
    
    If deb_bona.rssel_busca_grupo_tipo.State = 1 Then deb_bona.rssel_busca_grupo_tipo.Close
        deb_bona.sel_busca_grupo_tipo xcliente
        
        If deb_bona.rssel_busca_grupo_tipo.RecordCount > 0 Then
            xcgc = deb_bona.rssel_busca_grupo_tipo.Fields("CGC")
            xgrupo = deb_bona.rssel_busca_grupo_tipo.Fields("GRUPO")
            Label1.Caption = xgrupo
            xtipo = deb_bona.rssel_busca_grupo_tipo.Fields("TIPO")
            Label2.Caption = xtipo
            deb_bona.up_tb_basecli xgrupo, xtipo, xcliente
            
            xqtd = xqtd + 1
        
        
        End If
        
        
        deb_bona.rssel_busca_cli_brancos.MoveNext
        
        xprg = xprg + 1
        
        PRG_BASECLI.Value = xprg
Loop
        
        MsgBox xqtd & " ATUALIZADOS"
         deb_bona.rssel_busca_cli_brancos.Close
         cmd_alteracao.Visible = False

End Sub

Private Sub cmd_atualiza_prazos_Click()
Dim xfilialctc As String
Dim xdata As Date
Dim xqtd As Integer

deb_bona.rssel_busca_prazos.Open
lab_vid.Caption = deb_bona.rssel_busca_prazos.RecordCount

deb_bona.rssel_busca_prazos.MoveFirst

Do Until deb_bona.rssel_busca_prazos.EOF
xfilialctc = deb_bona.rssel_busca_prazos.Fields("filialctc")

    If deb_bona.rssel_busca_prazos_salete.State = 1 Then deb_bona.rssel_busca_prazos_salete.Close
        deb_bona.sel_busca_prazos_salete xfilialctc
        
        xdata = deb_bona.rssel_busca_prazos_salete.Fields("DATA")
        
    deb_bona.up_prazo_ctc xdata, xfilialctc
    xqtd = xqtd + 1
    deb_bona.rssel_busca_prazos.MoveNext
    
        
    Loop
    
    MsgBox xqtd & " ALTERADOS", vbInformation, "VIDEOLAR"
    
        


End Sub

Private Sub cmd_busca_cgc_branco_Click()

deb_bona.rssel_cgc_branco.Open
lab_qtd_basecli.Caption = deb_bona.rssel_cgc_branco.RecordCount
If deb_bona.rssel_cgc_branco.RecordCount < 1 Then
    MsgBox "NÃO"

Else
grd_clientes.DataMember = "sel_cgc_branco"
grd_clientes.Refresh

End If

cmd_alteracao.Visible = False
cmd_altera_cgc.Visible = True


End Sub

Private Sub cmd_busca_clibrancos_Click()

deb_bona.rssel_busca_cli_brancos.Open
lab_qtd_basecli.Caption = deb_bona.rssel_busca_cli_brancos.RecordCount
If deb_bona.rssel_busca_cli_brancos.RecordCount < 1 Then
    MsgBox "NÃO"

Else
grd_clientes.DataMember = "sel_busca_cli_brancos"
grd_clientes.Refresh

End If

cmd_alteracao.Visible = True
cmd_altera_cgc.Visible = False




End Sub

Private Sub cmd_inserir_fox_Click()
Dim xserie As String
Dim xcgc As String
Dim xsap As String
Dim xnome As String
Dim xfanatasia As String
Dim xgrupo As String
Dim xtipo As String
Dim xcont As Integer
Dim xcgc1 As String



deb_bona.rssel_fox_imp.MoveFirst

Do Until deb_bona.rssel_fox_imp.EOF

 deb_bona.rssel_count_vlclientes.Open
        lab_vl.Caption = deb_bona.rssel_count_vlclientes.Fields("TUDO")
        deb_bona.rssel_count_vlclientes.Close

If deb_bona.rssel_compara_cgc_fox.State = 1 Then deb_bona.rssel_compara_cgc_fox.Close
    deb_bona.sel_compara_cgc_fox deb_bona.rssel_fox_imp.Fields("CGC")
    
        If deb_bona.rssel_compara_cgc_fox.RecordCount < 1 Then
                       
        xserie = deb_bona.rssel_fox_imp.Fields("serie")
        
        xcgc1 = deb_bona.rssel_fox_imp.Fields("CGC")
       
        xsap = deb_bona.rssel_fox_imp.Fields("SAP")
        xnome = deb_bona.rssel_fox_imp.Fields("CLIENTE")
        xfantasia = deb_bona.rssel_fox_imp.Fields("FANTASIA")
        xgrupo = deb_bona.rssel_fox_imp.Fields("GRUPO")
        xtipo = deb_bona.rssel_fox_imp.Fields("TIPO")
                
        deb_bona.in_vl_clientes xserie, xcgc, xsap, xnome, xfantasia, xgrupo, xtipo
        
    xcont = xcont + 1

    If xcont = 150 Then
    MsgBox xcont
    End If
        
        
    End If
    
deb_bona.rssel_fox_imp.MoveNext




Loop



End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub cmd_tb_fox_imp_Click()

deb_bona.rssel_count_vlclientes.Open
lab_vl.Caption = deb_bona.rssel_count_vlclientes.Fields("TUDO")
deb_bona.rssel_count_vlclientes.Close

deb_bona.rssel_fox_imp.Open

If deb_bona.rssel_fox_imp.RecordCount < 1 Then
    MsgBox "Tabela Sem Registros", vbInformation, "FOX"
Else
grd_fox_imp.DataMember = "sel_fox_imp"
grd_fox_imp.Refresh
lab_tb_fox_imp.Caption = deb_bona.rssel_fox_imp.RecordCount
End If




End Sub
