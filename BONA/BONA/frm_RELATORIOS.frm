VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_RELATORIOS 
   Caption         =   "Relatórios - Divesos"
   ClientHeight    =   9105
   ClientLeft      =   540
   ClientTop       =   1725
   ClientWidth     =   11190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   11190
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&SAIR"
      Height          =   375
      Left            =   9540
      TabIndex        =   10
      Top             =   8700
      Width           =   1635
   End
   Begin TabDlg.SSTab tab_Doni 
      Height          =   8595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   15161
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Relatórios Donizette"
      TabPicture(0)   =   "frm_RELATORIOS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txt_nome"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmd_procura"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "grd_rel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt_local"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frm_RELATORIOS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frm_RELATORIOS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   795
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Width           =   3195
         Begin MSMask.MaskEdBox mask_data_inicio 
            Height          =   315
            Left            =   360
            TabIndex        =   12
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mask_data_final 
            Height          =   315
            Left            =   1800
            TabIndex        =   13
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "á"
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label3 
            Caption         =   "De:"
            Height          =   255
            Left            =   60
            TabIndex        =   14
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.TextBox txt_local 
         Height          =   285
         Left            =   6000
         TabIndex        =   9
         Text            =   "C:\"
         Top             =   8160
         Width           =   2715
      End
      Begin MSDataGridLib.DataGrid grd_rel 
         Height          =   5175
         Left            =   60
         TabIndex        =   3
         Top             =   2820
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   9128
         _Version        =   393216
         BackColor       =   0
         ForeColor       =   65280
         HeadLines       =   1
         RowHeight       =   16
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
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "sel_pesquisa"
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "DOC"
            Caption         =   "DOC"
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
         BeginProperty Column02 
            DataField       =   "FILIAL_CTC"
            Caption         =   "FILIAL_CTC"
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
            DataField       =   "NFS"
            Caption         =   "NFS"
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
            DataField       =   "MANIFESTO"
            Caption         =   "MANIFESTO"
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
            DataField       =   "PLACA"
            Caption         =   "PLACA"
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
         BeginProperty Column07 
            DataField       =   "DESTINATARIO"
            Caption         =   "DESTINATARIO"
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
         BeginProperty Column08 
            DataField       =   "CIDADE_DEST"
            Caption         =   "CIDADE_DEST"
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
         BeginProperty Column09 
            DataField       =   "UF_DEST"
            Caption         =   "UF_DEST"
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
         BeginProperty Column10 
            DataField       =   "MODAL"
            Caption         =   "MODAL"
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
         BeginProperty Column11 
            DataField       =   "TRANSPORTADORA"
            Caption         =   "TRANSPORTADORA"
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
         BeginProperty Column12 
            DataField       =   "VOLUMES"
            Caption         =   "VOLUMES"
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
         BeginProperty Column13 
            DataField       =   "PESO"
            Caption         =   "PESO"
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
         BeginProperty Column14 
            DataField       =   "VAL_MERC"
            Caption         =   "VAL_MERC"
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
         BeginProperty Column15 
            DataField       =   "FRETE"
            Caption         =   "FRETE"
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
               ColumnWidth     =   494,929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1649,764
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   854,929
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1395,213
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   120
         TabIndex        =   4
         Top             =   7980
         Width           =   10755
         Begin VB.CommandButton CMD_GERAR 
            Caption         =   "&Gerar Relatório .TXT"
            Height          =   255
            Left            =   8640
            TabIndex        =   5
            Top             =   180
            Width           =   1815
         End
         Begin VB.Label LAB_TOTAL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   600
            TabIndex        =   7
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "QTD.:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.CommandButton cmd_procura 
         Caption         =   "Procurar"
         Height          =   255
         Left            =   900
         TabIndex        =   2
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txt_nome 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   900
         TabIndex        =   1
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "CLIENTE:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   795
      End
   End
End
Attribute VB_Name = "frm_RELATORIOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_GERAR_Click()
 Dim xprocura As String


xprocura = UCase(txt_nome.Text)
      
    If deb_consulta.rssel_pesquisa.State = 0 Then deb_consulta.sel_pesquisa mask_data_inicio, mask_data_final, "%" & xprocura & "%"
        
    
    
    If deb_consulta.rssel_pesquisa.RecordCount > 0 Then
        Open txt_local.Text & UCase(txt_nome.Text) & ".TXT" For Output As #1
        XDOC = "DOC"
        XDATA = "DATA"
        XFILIAL = "FILIAL"
        XNFS = "NFS"
        XMANIFESTO = "MANIFESTO"
        XPLACA = "PLACA"
        XCLIENTE = "CLIENTE"
        XDESTINATARIO = "DESTINATARIO"
        XCIDADE = "CIDADE"
        XUF = "UF"
        XMODAL = "MODAL"
        XTRANSPORTADORA = "TRANSPORTADORA"
        XVOLUMES = "VOLUMES"
        XPESO = "PESO"
        XVALMERC = "VALOR_MERCADORIA"
        XFRETE = "FRETE"
        xlinha = XDOC & "#" & XDATA & "#" & XFILIAL & "#" & XNFS & "#" & XMANIFESTO & "#" & XPLACA & "#" & _
        XCLIENTE & "#" & XDESTINATARIO & "#" & XCIDADE & "#" & XUF & "#" & XMODAL & "#" & _
        XTRANSPORTADORA & "#" & XVOLUMES & "#" & XPESO & "#" & XVALMERC & "#" & XFRETE
        
        Print #1, xlinha
        deb_consulta.rssel_pesquisa.MoveNext
            Do Until deb_consulta.rssel_pesquisa.EOF
            XDOC = deb_consulta.rssel_pesquisa.Fields("DOC")
            XDATA = deb_consulta.rssel_pesquisa.Fields("DATA")
        XFILIAL = deb_consulta.rssel_pesquisa.Fields("FILIAL_CTC")
        XNFS = deb_consulta.rssel_pesquisa.Fields("NFS")
        XMANIFESTO = deb_consulta.rssel_pesquisa.Fields("MANIFESTO")
        XPLACA = deb_consulta.rssel_pesquisa.Fields("PLACA")
        XCLIENTE = deb_consulta.rssel_pesquisa.Fields("CLIENTE")
        XDESTINATARIO = deb_consulta.rssel_pesquisa.Fields("DESTINATARIO")
        XCIDADE = deb_consulta.rssel_pesquisa.Fields("CIDADE_DEST")
        XUF = deb_consulta.rssel_pesquisa.Fields("UF_DEST")
        XMODAL = deb_consulta.rssel_pesquisa.Fields("MODAL")
        XTRANSPORTADORA = deb_consulta.rssel_pesquisa.Fields("TRANSPORTADORA")
        XVOLUMES = deb_consulta.rssel_pesquisa.Fields("VOLUMES")
        XPESO = deb_consulta.rssel_pesquisa.Fields("PESO")
        XVALMERC = deb_consulta.rssel_pesquisa.Fields("VAL_MERC")
        XFRETE = deb_consulta.rssel_pesquisa.Fields("FRETE")
        xlinha = XDOC & "#" & XDATA & "#" & XFILIAL & "#" & XNFS & "#" & XMANIFESTO & "#" & XPLACA & "#" & _
        XCLIENTE & "#" & XDESTINATARIO & "#" & XCIDADE & "#" & XUF & "#" & XMODAL & "#" & _
        XTRANSPORTADORA & "#" & XVOLUMES & "#" & XPESO & "#" & XVALMERC & "#" & XFRETE
           Print #1, xlinha
            deb_consulta.rssel_pesquisa.MoveNext
            Loop
        Close #1
        MsgBox "Arquivo Gerado com Sucesso" + Chr$(13) + "Arquivo: C:\" & xprocura & ".TXT", vbInformation, "Arquivo Gerado"
                
        End If
End Sub

Private Sub cmd_procura_Click()
Dim xprocura As String

xprocura = txt_nome.Text

If deb_consulta.rssel_pesquisa.State = 1 Then deb_consulta.rssel_pesquisa.Close
    deb_consulta.sel_pesquisa mask_data_inicio, mask_data_final, "%" & xprocura & "%"
    
If deb_consulta.rssel_pesquisa.RecordCount < 1 Then
    MsgBox "Não há Dados para Pesquisa", vbInformation, "Arquivos Não Localizados"

Else
grd_rel.DataMember = "sel_pesquisa"
grd_rel.Refresh
End If

LAB_TOTAL.Caption = deb_consulta.rssel_pesquisa.RecordCount
End Sub

Private Sub cmd_sair_Click()
Unload Me
End Sub
