VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManifesto 
   BackColor       =   &H8000000A&
   Caption         =   "MANIFESTO"
   ClientHeight    =   8670
   ClientLeft      =   720
   ClientTop       =   915
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   11790
   WindowState     =   2  'Maximized
   Begin VB.Frame MANIFESTO 
      Caption         =   "Manifesto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   11535
      Begin VB.CommandButton cmdGerar 
         Caption         =   "Gerar Relatório *.TXT"
         Height          =   375
         Left            =   9720
         TabIndex        =   51
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados Manifesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   9495
         Begin VB.Label labFrete 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6600
            TabIndex        =   50
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label labValmerc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   6600
            TabIndex        =   49
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label labconferente 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6600
            TabIndex        =   48
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label labPeso 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4320
            TabIndex        =   47
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label labEmissor 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4320
            TabIndex        =   46
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label labhorario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2520
            TabIndex        =   40
            Top             =   240
            Width           =   975
         End
         Begin VB.Label labvolumes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2520
            TabIndex        =   39
            Top             =   600
            Width           =   975
         End
         Begin VB.Label labCTCS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   720
            TabIndex        =   38
            Top             =   600
            Width           =   975
         End
         Begin VB.Label labDataMani 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   720
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "CTC´S:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Volumes:"
            Height          =   255
            Left            =   1800
            TabIndex        =   31
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Peso:"
            Height          =   255
            Left            =   3600
            TabIndex        =   30
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Val. Frete:"
            Height          =   255
            Left            =   5760
            TabIndex        =   29
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Val. Merc:"
            Height          =   255
            Left            =   5760
            TabIndex        =   28
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "Horário:"
            Height          =   255
            Left            =   1800
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "Emissor:"
            Height          =   255
            Left            =   3600
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "Conferente:"
            Height          =   255
            Left            =   5760
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid gridMani 
         Height          =   3375
         Left            =   120
         TabIndex        =   21
         Top             =   4200
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5953
         _Version        =   393216
         BackColor       =   0
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
         DataMember      =   "sel_gride"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "FILIAL_MANIFESTO"
            Caption         =   "FILIAL_MANIFESTO"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "CIDADE"
            Caption         =   "CIDADE"
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
            DataField       =   "UF"
            Caption         =   "UF"
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
         BeginProperty Column07 
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
               ColumnWidth     =   1560,189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140,095
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
            BeginProperty Column05 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdPesquisar 
         BackColor       =   &H80000007&
         Caption         =   "Pesquisar"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
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
         Left            =   10080
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Veículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   11295
         Begin VB.CheckBox chkCamFria 
            BackColor       =   &H8000000B&
            Caption         =   "Cam. Fria"
            Height          =   255
            Left            =   9720
            TabIndex        =   18
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CheckBox chkPlatHidr 
            Caption         =   "Plat. Hidr."
            Height          =   255
            Left            =   9720
            TabIndex        =   17
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CheckBox ChkSuspAr 
            Caption         =   "Susp. AR"
            Height          =   255
            Left            =   9720
            TabIndex        =   16
            Top             =   960
            Width           =   975
         End
         Begin VB.Label labAjudantes 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   45
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label labAno 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9720
            TabIndex        =   44
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label labTipo 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   9720
            TabIndex        =   43
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label labMotorista 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   42
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label labCGC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   41
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label labPlaca 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   36
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label labMarca 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   35
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label labRastreamento 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6000
            TabIndex        =   34
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label labCodigo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label labProprietario 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label20 
            Caption         =   "Ajudantes:"
            Height          =   255
            Left            =   4800
            TabIndex        =   20
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Rastreamento:"
            Height          =   255
            Left            =   4800
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Codigo:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Proprietario CGC:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Proprietário Veic.:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Motorista:"
            Height          =   255
            Left            =   4800
            TabIndex        =   11
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Ano:"
            Height          =   255
            Left            =   8880
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Veíc.:"
            Height          =   255
            Left            =   8880
            TabIndex        =   9
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Marca/Modelo:"
            Height          =   255
            Left            =   4800
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Placa:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Caption         =   "Buscar"
         Height          =   255
         Left            =   2760
         MaskColor       =   &H00FF8080&
         TabIndex        =   2
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtManifesto 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Manifesto Nº :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmManifesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPesquisar_Click()
PESQUISA.Show
DoEvents

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim xfilialmanifesto As String


xfilialmanifesto = txtFilial & String(6 - Len(Trim$(txtManifesto)), "0") & Trim$(txtManifesto)
If debManifesto.rssel_manifesto.State = 1 Then debManifesto.rssel_manifesto.Close
    debManifesto.sel_manifesto xfilialmanifesto
    
xfilialmanifesto = txtFilial & String(6 - Len(Trim$(txtManifesto)), "0") & Trim$(txtManifesto)
If debManifesto.rssel_calculos.State = 1 Then debManifesto.rssel_calculos.Close
    debManifesto.sel_calculos xfilialmanifesto
  
xfilialmanifesto = txtFilial & String(6 - Len(Trim$(txtManifesto)), "0") & Trim$(txtManifesto)
If debManifesto.rssel_gride.State = 1 Then debManifesto.rssel_gride.Close
    debManifesto.sel_gride xfilialmanifesto
  
  
If debManifesto.rssel_manifesto.RecordCount < 1 Then
    MsgBox "Filial/Manifesto não Localizados", vbInformation, "Manifesto Não Localizado"

Else
    labCodigo = debManifesto.rssel_manifesto.Fields("codigo")
    labPlaca = debManifesto.rssel_manifesto.Fields("placaveic")
    labDataMani = debManifesto.rssel_manifesto.Fields("dtemissao")
    labProprietario = debManifesto.rssel_manifesto.Fields("proprietario")
    labRastreamento = debManifesto.rssel_manifesto.Fields("rastreamento")
    labMarca = debManifesto.rssel_manifesto.Fields("marca")
    labMotorista = debManifesto.rssel_manifesto.Fields("motorista")
    labTipo = debManifesto.rssel_manifesto.Fields("tipo")
    labAno = debManifesto.rssel_manifesto.Fields("ano")
    labCGC = debManifesto.rssel_manifesto.Fields("proprietariocgc")
    labhorario = debManifesto.rssel_manifesto.Fields("hsemissao")
    labconferente = debManifesto.rssel_manifesto.Fields("conferente")
    labAjudantes = debManifesto.rssel_manifesto.Fields("ajudantes")
    ChkSuspAr.Enabled = False
    chkPlatHidr.Enabled = False
    chkCamFria.Enabled = False
    If debManifesto.rssel_manifesto.Fields("suspensaoar") = "S" Then
        ChkSuspAr.Value = 1
    Else
        ChkSuspAr.Value = 0
    End If
    If debManifesto.rssel_manifesto.Fields("plataformahidr") = "S" Then
        chkPlatHidr.Value = 1
    Else
        chkPlatHidr.Value = 0
    End If
    If debManifesto.rssel_manifesto.Fields("camarafria") = "S" Then
        chkCamFria.Value = 1
    Else
        chkCamFria.Value = 0
    End If
    If IsNull(debManifesto.rssel_manifesto.Fields("usuariocad")) Then
        labEmissor = ""
    Else
        labEmissor = debManifesto.rssel_manifesto.Fields("usuariocad")
    End If
    labCTCS = debManifesto.rssel_calculos.Fields("QTDCTC")
    labvolumes = debManifesto.rssel_calculos.Fields("QTDVOLUMES")
    labPeso = debManifesto.rssel_calculos.Fields("QTDPESO")
    labFrete = debManifesto.rssel_calculos.Fields("FRETE")
    labValmerc = debManifesto.rssel_calculos.Fields("VALMERC")
    gridMani.DataMember = "sel_gride"
    gridMani.Refresh
    
End If



End Sub

Private Sub cmdGerar_Click()
Dim xfilialmanifesto As String

xfilialmanifesto = txtFilial & String(6 - Len(Trim$(txtManifesto)), "0") & Trim$(txtManifesto)
If debManifesto.rssel_gride.State = 0 Then debManifesto.sel_gride txtfilialmanifesto

If debManifesto.rssel_gride.RecordCount > 0 Then
    Open "C:\MANIFESTO.txt" For Output As #1
    xfilial = "FILIAL_MANIFESTO"
    xfilialctc = "FILIAL_CTC"
    xcliente = "CLIENTE"
    xdestinatario = "DESTINATARIO"
    xCidade = "CIDADE"
    xuf = "UF"
    xvolumes = "VOLUMES"
    xpeso = "PESO"
    xvalmerc = "VAL_MERC"
    xfrete = "FRETE"
    xlinha = xfilial & "#" & xfilialctc & "#" & xcliente & "#" & xdestinatario & "#" & xCidade & "#" & xuf & "#" & xvolumes & "#" & xpeso & "#" & xvalmerc & "#" & xfrete
    Print #1, xlinha
    debManifesto.rssel_gride.MoveFirst
        Do Until debManifesto.rssel_gride.EOF
        xfilial = debManifesto.rssel_gride.Fields("FILIAL_MANIFESTO")
        xfilialctc = debManifesto.rssel_gride.Fields("FILIAL_CTC")
        xcliente = debManifesto.rssel_gride.Fields("CLIENTE")
        xdestinatario = debManifesto.rssel_gride.Fields("DESTINATARIO")
        xCidade = debManifesto.rssel_gride.Fields("CIDADE")
        xuf = debManifesto.rssel_gride.Fields("UF")
        xvolumes = debManifesto.rssel_gride.Fields("VOLUMES")
        xpeso = debManifesto.rssel_gride.Fields("PESO")
        xvalmerc = debManifesto.rssel_gride.Fields("VAL_MERC")
        xfrete = debManifesto.rssel_gride.Fields("FRETE")
        xlinha = xfilial & "#" & xfilialctc & "#" & xcliente & "#" & xdestinatario & "#" & xCidade & "#" & xuf & "#" & xvolumes & "#" & xpeso & "#" & xvalmerc & "#" & xfrete
        Print #1, xlinha
        debManifesto.rssel_gride.MoveNext
        Loop
    Close #1
    MsgBox "ARQUIVO MANIFESTO.TXT GRAVADO EM C:\.O Arquivo Gerado é do Tipo Texto ( TXT com Delimitador # ) e você pode abrí-lo em diversos aplicativos. Para Abrí-lo no MS-Excel, em ABRIR escolha ARQUIVOS DO TIPO = Arquivos de Texto e selecione o arquivo no local indicado acima. Na Caixa ASSISTENTE DE IMPORTAÇÃO escolha DELIMITADO e o caracter delimitador escolha OUTROS e digite # . Clique em Concluir e o arquivo será importado para o MS-Excel.", vbInformation, "Geração Finalizada"
End If

    


End Sub

Private Sub Form_Load()
    mdiInforma.Toolbar1.Visible = False
    mdiInforma.StatusBar1.Visible = False
    gridMani.DataMember = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Visible = True
    mdiInforma.StatusBar1.Visible = True
    Set frmManifesto = Nothing
End Sub
