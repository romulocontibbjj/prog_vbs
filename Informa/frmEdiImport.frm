VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEdiImport 
   Caption         =   "Importa EDI de Ocorrências (Padrão PROCEDA)"
   ClientHeight    =   7350
   ClientLeft      =   765
   ClientTop       =   1620
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_importacorreio 
      Caption         =   "CORREIOS"
      Height          =   375
      Left            =   6480
      TabIndex        =   70
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdImportCorreio 
      Caption         =   "Importar Objetos Correio ( intec.xls )"
      Height          =   495
      Left            =   6480
      TabIndex        =   69
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Status da Importação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Width           =   3375
      Begin VB.Label lblBxGrav 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   63
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblOcorrGrav 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   61
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblReg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   51
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblTotReg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   59
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Total Reg. Gravados.....:"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblTotGrav 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   65
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Baixas Gravadas............:"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Ocorrências Gravadas...:"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Total de Registros..........:"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Registros Criticados.......:"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblCrit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   57
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Baixas Lidas...................:"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblBxLida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   55
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Ocorrências Lidas..........:"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblOcorrLida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   53
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label35 
         Caption         =   "Registros Lidos..............:"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Processo de Recepção do Arquivo EDI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   11775
      Begin VB.Frame Frame13 
         Caption         =   "Status da Ocorrência"
         Height          =   735
         Left            =   6720
         TabIndex        =   46
         Top             =   2520
         Width           =   2415
         Begin VB.Label lblEDIStatus 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Nome do Recebedor"
         Height          =   735
         Left            =   9240
         TabIndex        =   38
         Top             =   2520
         Width           =   2415
         Begin VB.Label lblEDIReceb 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Ocorrência"
         Height          =   735
         Left            =   6720
         TabIndex        =   37
         Top             =   1560
         Width           =   4935
         Begin VB.Label lblEDIDescOcorr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   600
            TabIndex        =   44
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label lblEDICodOcorr 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Data/Hora Ocorr."
         Height          =   735
         Left            =   9840
         TabIndex        =   36
         Top             =   600
         Width           =   1815
         Begin VB.Label lblEDIHs 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1200
            TabIndex        =   42
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblEDIData 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "NF"
         Height          =   735
         Left            =   8520
         TabIndex        =   35
         Top             =   600
         Width           =   1215
         Begin VB.Label lblEDINf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "CGC Remetente"
         Height          =   735
         Left            =   6720
         TabIndex        =   34
         Top             =   600
         Width           =   1695
         Begin VB.Label lblEdiCGC 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Destino"
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   6375
         Begin VB.Label lblDestUf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5880
            TabIndex        =   48
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblDestCid 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3720
            TabIndex        =   24
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblDestCli 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Origem"
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   6375
         Begin VB.Label lblOriUf 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5880
            TabIndex        =   21
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblOriCid 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3720
            TabIndex        =   20
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblOriCli 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Inf. de Entrega"
         Height          =   1455
         Left            =   4440
         TabIndex        =   17
         Top             =   2040
         Width           =   2055
         Begin VB.Label lblBxEntrega 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   840
            TabIndex        =   33
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblRecEntrega 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   840
            TabIndex        =   32
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblHsEntrega 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   840
            TabIndex        =   49
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblDtEntrega 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   840
            TabIndex        =   31
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Quem Bx:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   690
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Receb:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   525
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Hs:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   390
         End
      End
      Begin MSDataGridLib.DataGrid gridOcorr 
         Bindings        =   "frmEdiImport.frx":0000
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
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
         DataMember      =   "Sel_ConsOcorr2"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "data"
            Caption         =   "Data"
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
            DataField       =   "hora"
            Caption         =   "hora"
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
            DataField       =   "cod_ocorr"
            Caption         =   "Cd"
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
            DataField       =   "descr_ocorr"
            Caption         =   "Ocorrência"
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
            DataField       =   "usu_ocorr"
            Caption         =   "usu_ocorr"
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
            DataField       =   "usu_dataocorr"
            Caption         =   "usu_dataocorr"
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
            DataField       =   "obs_ocorr"
            Caption         =   "obs_ocorr"
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
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   285,165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2550,047
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   68
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Emissão:"
         Height          =   255
         Left            =   2400
         TabIndex        =   67
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblConfirma 
         AutoSize        =   -1  'True
         Caption         =   "confirma"
         Height          =   195
         Left            =   8880
         TabIndex        =   50
         Top             =   3360
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblFilialCTC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Filial - CTC:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dados do Arquivo EDI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   9960
         TabIndex        =   15
         Top             =   240
         Width           =   1590
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   6600
         X2              =   6600
         Y1              =   240
         Y2              =   3480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dados no Sistema INTEC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Identificação do Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   8175
      Begin VB.Label lblIdent 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6480
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblRemEDI 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Identificação do Arquivo:"
         Height          =   195
         Left            =   4560
         TabIndex        =   11
         Top             =   360
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rem.EDI:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdProcessa 
      Caption         =   "Processar..."
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Arquivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.FileListBox fileImport 
         Height          =   1065
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtArq 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
      Begin VB.DirListBox DirImporta 
         Height          =   1890
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo Escolhido"
         Height          =   195
         Left            =   4080
         TabIndex        =   4
         Top             =   1320
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmEdiImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_importacorreio_Click()
Dim xnf As String
Dim xdata As String
Dim xhora As String
Dim xcod_status As String
Dim xStatus As String
Dim xlocal As String
Dim xregistro As String
Dim xlinha As String
Dim xLin As Integer
Dim x40 As Integer
Dim xdif As Integer
Dim xserie As Integer
Dim aux As String
Dim X As Integer
Dim xocorr As String
Dim xnomearq As String

Me.MousePointer = 11

xnomearq = fileImport.Path & "\" & fileImport.FileName


Open xnomearq For Input As #1
'CONTA OS REGISTROS
    Do Until EOF(1)
        xLin = xLin + 1
        
        Line Input #1, xlinha
            
            If Mid(xlinha, 39, 2) = "40" Then
                x40 = x40 + 1
                lblBxLida.Caption = x40
            Else
                xdif = xdif + 1
                lblCrit.Caption = xdif
             End If
        frmEdiImport.Refresh
    Loop
    
'COMEÇA A LEITURA
    lblTotReg.Caption = xLin
    Close #1
    xLin = 0
    x40 = 0
    
Open xnomearq For Input As #1
    
     Do Until EOF(1)
        
        xLin = xLin + 1
        
        Line Input #1, xlinha
            If Mid(xlinha, 39, 2) = "40" Then
                
                                
                xnf = Trim$(Mid$(xlinha, 1, 20))
                
                If IsNumeric(Mid(xnf, 1, 1)) = True Then
                aux = 0
                X = 0
                                
                               
                                
                'verifica caracter por caracter
                For X = 1 To Len(xnf)
                    aux = Mid(xnf, X, 1)
                   If IsNumeric(aux) = False Then
                       Exit For
                    End If
                Next
                
                aux = Mid(xnf, 1, (X - 1))
                
                'SEPARA NF DO DIGITO
                xserie = Mid(xnf, Len(xnf), 1)
                xnf = aux
                
                'VÊ SE XNF É NUMÉRICO
                
                If xnf <> "" Then
                
                xdata = Trim(Mid$(xlinha, 21, 10))
                xhora = Trim(Mid$(xlinha, 31, 8))
                xcod_status = "01"
                xStatus = UCase(Trim$(Mid$(xlinha, 41, 20)))
                xlocal = Mid$(xlinha, 61, 130)
                xregistro = Mid$(xlinha, 191, 13)
                lblEDINf.Caption = xnf
                lblEDIData.Caption = xdata
                lblEDIHs.Caption = xhora
                lblEDICodOcorr.Caption = xcod_status
                lblEDIDescOcorr.Caption = xStatus
                lblRemEDI.Caption = "CORREIOS"
                lblIdent.Caption = "INTEC.TXT"
                lblEdiCGC.Caption = "04229761000413"
                              
                frmEdiImport.Refresh
                
                'SE NÃO FOR NUMÉRICO ELE VAI PARA A PRÓXIMA NF
                'Else
                
                 '   xdif = xdif + 1
                  '  lblCrit = xdif
                   ' Me.Refresh
                   ' Loop
                
               End If
                
                
                    If de_informa.rsSel_BuscaPorCGC_NF_Serie.State = 1 Then de_informa.rsSel_BuscaPorCGC_NF_Serie.Close
                        de_informa.Sel_BuscaPorCGC_NF_Serie "04229761000413", xnf, xserie
                        
                        If de_informa.rsSel_BuscaPorCGC_NF_Serie.RecordCount > 0 Then
                            If de_informa.rsSel_BuscaPorCGC_NF_Serie.RecordCount > 1 Then
                                de_informa.rsSel_BuscaPorCGC_NF_Serie.MoveLast
                            End If
                            
                            lblFilialCTC.Caption = de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("filialctc")
                            lblEmissao.Caption = de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("data")
                                
                            If de_informa.rsSel_CTC_NFeCGC.State = 1 Then de_informa.rsSel_CTC_NFeCGC.Close
                                de_informa.Sel_CTC_NFeCGC "04229761" & "%", xnf
                               
                                    If de_informa.rsSel_CTC_NFeCGC.RecordCount > 0 Then
                                        lblDestCid.Caption = de_informa.rsSel_CTC_NFeCGC.Fields("cidade_dest")
                                        lblDestUf.Caption = de_informa.rsSel_CTC_NFeCGC.Fields("uf_dest")
                                        lblDestCli.Caption = de_informa.rsSel_CTC_NFeCGC.Fields("dest_nome")
                                        lblOriCli.Caption = de_informa.rsSel_CTC_NFeCGC.Fields("remet_nome")
                                        lblOriCid.Caption = de_informa.rsSel_CTC_NFeCGC.Fields("remet_cidade")
                                        lblOriUf.Caption = de_informa.rsSel_CTC_NFeCGC.Fields("remet_uf")
                                        lblEDIReceb.Caption = xregistro
                                        lblEDIStatus.Caption = "ENTREGUE"
                                        xocorr = de_informa.rsSel_CTC_NFeCGC.Fields("tem_ocorr")
                                        lblEDIReceb.Caption = ""
                                            If xocorr = "2" Or xocorr = "N" Then
                                    
                                                    If lblDtEntrega = "" Then  'Nova Baixa ...
                                                     'inicia a transação
                                                        de_informa.cn_informa.BeginTrans
                                                        de_informa.ins_ocorr1 lblFilialCTC, de_informa.rsSel_CTC_NFeCGC.Fields("data"), lblEdiCGC, lblEDICodOcorr, lblEDIDescOcorr, _
                                                        CDate(lblEDIData), lblEDIHs, CDate(lblEDIData), lblEDIHs, lblEDIReceb & "", "EDI-" & Mid$(lblRemEDI, 1, 6), datahora("datahora"), "S", datahora("data")
                                                        de_informa.alt_temocorr_sn "1", lblFilialCTC
                                                        de_informa.Alt_AtClienteNFBranco lblFilialCTC
                                                        de_informa.cn_informa.CommitTrans
                                                        x40 = x40 + 1
                                                        lblBxGrav.Caption = x40
                                                        
                                                    End If
                                                    If Mid$(de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("obs_emissao"), 1, 15) <> "OBJETO CORREIO:" Then
                                                    de_informa.Alt_ObjCorreioObs "OBJETO CORREIO:" & xregistro & "//" & de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("obs_emissao"), de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("filialctc")
                                                   
                                                    End If
                                                                                        
                                            Else
                                                xdif = xdif + 1
                                                    lblCrit = xdif
                                            
                                            End If
            
                                    End If
                         
                        End If
                                                  
                    End If
                                           
                 End If
                    
                       
                        
                        
                        
    Loop
    Close #1
    
    FileCopy xnomearq, fileImport.Path & "\" & "BACKUP\" & fileImport.FileName
    Kill xnomearq
    fileImport.Refresh
    Me.MousePointer = 0
MsgBox "Arquivo Gerado"

    
txtArq = ""
    lblFilialCTC = ""
    lblEmissao = ""
    lblOriCli = ""
    lblOriCid = ""
    lblOriUf = ""
    lblDestCid = ""
    lblDestCli = ""
    lblDestUf = ""
    'Set gridOcorr.DataSource = Null
    gridOcorr.DataMember = ""
    gridOcorr.Refresh
    lblDtEntrega = ""
    lblHsEntrega = ""
    lblRecEntrega = ""
    lblBxEntrega = ""
    lblRemEDI = ""
    lblIdent = ""
    lblReg = ""
    lblEdiCGC = ""
    lblEDINf = ""
    lblEDIData = ""
    lblEDIHs = ""
    lblEDICodOcorr = ""
    lblEDIDescOcorr = ""
    lblEDIStatus = ""
    lblEDIReceb = ""
    lblReg = ""
    lblTotReg = ""
    lblOcorrGrav = ""
    lblOcorrLida = ""
    lblBxLida = ""
    lblBxGrav = ""
    lblTotGrav = ""
    lblCrit = ""
    DirImporta.Enabled = True
    fileImport.Enabled = True
    txtArq.Enabled = True
    cmdProcessa.Enabled = True
    cmdSair.Enabled = True
  


End Sub

Private Sub cmdImportCorreio_Click()
    Dim Excel As Excel.Application
    Dim ExcelA1 As Excel.Worksheet
    Dim LinhaAtual As Integer, LinhaMAX As Integer
    
    LinhaMAX = 0
    
    cmdImportCorreio.Caption = "AGUARDE ..."
    
    Set Excel = CreateObject("EXCEL.APPLICATION")
    Excel.Visible = False
    Excel.Interactive = False
    Excel.Workbooks.Open FileName:="C:\INFORMA\CORREIOS\INTEC.xls"
    Set ExcelA1 = Excel.Worksheets(1)
    
    LinhaAtual = 0
    
    Do While True
        LinhaAtual = LinhaAtual + 1
        LinhaMAX = LinhaMAX + 1
        If Len(Trim(ExcelA1.Cells(LinhaAtual, 1))) = 0 And LinhaAtual >= 2 Then
            Exit Do
        End If
    Loop
    
    DoEvents
    
    For LinhaAtual = 1 To LinhaMAX - 1
    
        xobjeto = Trim(ExcelA1.Cells(LinhaAtual, 1))
        xnfserie = Trim(ExcelA1.Cells(LinhaAtual, 2))
        
        If InStr(1, xnfserie, "-", vbTextCompare) = 0 Then
        Else
        
            xnf = Mid$(xnfserie, 1, InStr(1, xnfserie, "-", vbTextCompare) - 1)
            xserie = Mid$(xnfserie, InStr(1, xnfserie, "-", vbTextCompare) + 1, 1)
            
            If de_informa.rsSel_BuscaPorCGC_NF_Serie.State = 1 Then de_informa.rsSel_BuscaPorCGC_NF_Serie.Close
            de_informa.Sel_BuscaPorCGC_NF_Serie "04229761000413", CDbl(xnf), xserie
            
            If de_informa.rsSel_BuscaPorCGC_NF_Serie.RecordCount > 0 Then
                
                Do Until de_informa.rsSel_BuscaPorCGC_NF_Serie.EOF
                    
                    If Mid$(de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("obs_emissao"), 1, 15) = "OBJETO CORREIO:" Then
                        
                        de_informa.Alt_ObjCorreioObs "OBJETO CORREIO: " & xobjeto & " // " & _
                        Mid$(de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("obs_emissao"), 31, 285 + 30), _
                        de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("filialctc")
                    
                    Else
                        
                        de_informa.Alt_ObjCorreioObs "OBJETO CORREIO: " & xobjeto & " // " & _
                        Mid$(de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("obs_emissao"), 1, 285), _
                        de_informa.rsSel_BuscaPorCGC_NF_Serie.Fields("filialctc")
                    
                    End If
                    
                    de_informa.rsSel_BuscaPorCGC_NF_Serie.MoveNext
                    
                Loop
            
            End If
            
        End If
        
    Next
    
    Excel.Quit
    Set ExcelA1 = Nothing
    Set Excel = Nothing
    
    cmdImportCorreio.Caption = "Importar Objetos Correio ( intec.xls )"
    
    MsgBox "Arquivo Importado !"

End Sub

Private Sub cmdProcessa_Click()

'DECLARAR AS VARIAVEIS DESTA SUB

    Dim xtranspcgc As String, xabonodias As Integer

     If txtArq.Text = "" Then
        MsgBox "Erro ! Nome de arquivo Inválido !"
        txtArq.SetFocus
        Exit Sub
     ElseIf Dir(DirImporta.Path & "\" & txtArq.Text) = "" Then
        MsgBox "Erro ! Arquivo Não Encontrado !"
        txtArq.SetFocus
        Exit Sub
     End If

'trava controles durante processamento

DirImporta.Enabled = False
fileImport.Enabled = False
txtArq.Enabled = False
cmdProcessa.Enabled = False
cmdSair.Enabled = False

'antes de contar a quantidade de registros verificar se o arquivo é realmente de OCORRÊNCIAS
'Abre o arquivo inicialmente para verificar a quantidade de registros
     
     xtamarq = 0
     xLin = 0
     Close #1
     Open DirImporta.Path & "\" & txtArq.Text For Input As #1
     Do Until EOF(1)
        xLin = xLin + 1
        Line Input #1, xlinha
        If xLin = 1 Then
            If Mid(xlinha, 1, 3) <> "000" Then
                MsgBox "Arquivo Inválido ! Processamento Cancelado."
                    DirImporta.Enabled = True
                    fileImport.Enabled = True
                    txtArq.Enabled = True
                    cmdProcessa.Enabled = True
                    cmdSair.Enabled = True
                    Close #1
                Exit Sub
            End If
        End If
        If xLin = 2 Then
            If Mid(xlinha, 1, 3) <> "340" And Mid(xlinha, 1, 3) <> "040" And Mid(xlinha, 4, 5) <> "OCORR" Then
                MsgBox "Arquivo Inválido ! Processamento Cancelado."
                    DirImporta.Enabled = True
                    fileImport.Enabled = True
                    txtArq.Enabled = True
                    cmdProcessa.Enabled = True
                    cmdSair.Enabled = True
                    Close #1
                Exit Sub
            End If
        End If
        If Mid(xlinha, 1, 3) = "342" Or Mid(xlinha, 1, 3) = "042" Then
            xtamarq = xtamarq + 1
        End If
     Loop
     Close #1
     
     If xtamarq = 0 Then
        MsgBox "Arquivo Inválido !"
        DirImporta.Enabled = True
        fileImport.Enabled = True
        txtArq.Enabled = True
        cmdProcessa.Enabled = True
        cmdSair.Enabled = True
        Close #1
        Exit Sub
     End If
     
'rotina para verificar os dados do arquivo e verificar se é versão 3 ou 3.1 ou mesmo se é PROCEDA

     lblTotReg.Caption = xtamarq

'Abre o arquivo para leitura (txt)
     
     Open DirImporta.Path & "\" & txtArq.Text For Input As #1

    'zera lbls contadoras de status
     lblReg = 0
     lblOcorrLida = 0
     lblOcorrGrav = 0
     lblBxLida = 0
     lblBxGrav = 0
     lblTotGrav = 0
     lblCrit = 0
     xtranspcgc = ""
     
     'Inicia o loop para leitura linha a linha do TXT
     Do Until EOF(1)
        Line Input #1, xlinha
        xid = Mid(xlinha, 1, 3)
        lblConfirma = ""

        If xid = "000" Then   'IDENTIFICA REGISTRO DE CABEÇALHO
            lblRemEDI = Mid(xlinha, 4, 35)
            lblIdent = Mid(xlinha, 84, 12)
        ElseIf xid = "341" Or Mid(xlinha, 1, 3) = "041" Then 'IDENTIFICA REGISTRO DE DADOS TRANSPORTADOR
            xtranspcgc = Mid(xlinha, 4, 8)
        ElseIf xid = "342" Or Mid(xlinha, 1, 3) = "042" Then 'IDENTIFICA REGISTRO DE DADOS OCORRENCIA/ENTREGA
            'trata e informa dados do arquivo do transportador
            lblReg = Val(lblReg) + 1
            xcgc = Mid(xlinha, 4, 14)
            xnf = Val(Mid(xlinha, 21, 8))
            xserie = Trim$(Mid$(xlinha, 18, 3))
            If IsNumeric(xserie) Then
                
            End If
            xcodocorr = Mid(xlinha, 29, 2)
            If de_informa.rsSel_ConsCadOcor.State = 1 Then de_informa.rsSel_ConsCadOcor.Close
            de_informa.Sel_ConsCadOcor xcodocorr  'busca ocorrência
            If de_informa.rsSel_ConsCadOcor.RecordCount > 0 Then
                lblEDIDescOcorr = de_informa.rsSel_ConsCadOcor.Fields("descricao")
                If de_informa.rsSel_ConsCadOcor.Fields("cod_ocorr") = "00" Then
                    lblEDIDescOcorr = "??"
                    If MsgBox("Arquivo com Código de Ocorrência Inválido e este Lançamento Não Poderá Ser Realizado   (Cód: 00 - Uso Interno). Deseja Continuar Este Processo ?", vbQuestion + vbYesNo, "Erro") = vbNo Then
                        DirImporta.Enabled = True
                        fileImport.Enabled = True
                        txtArq.Enabled = True
                        cmdProcessa.Enabled = True
                        cmdSair.Enabled = True
                        Exit Sub
                    End If
                End If
            Else
                lblEDIDescOcorr = "??"     'se não encontra ocorrência
            End If
            'traz dados do arq EDI para campos (lado dados do arquivo EDI)
            If (Mid(xlinha, 31, 2) & "/" & Mid(xlinha, 33, 2) & "/" & Mid(xlinha, 35, 4)) = "00/00/0000" Then
                MsgBox "Erro no Arquivo Data Inválida: " & (Mid(xlinha, 31, 2) & "/" & Mid(xlinha, 33, 2) & "/" & Mid(xlinha, 35, 4))
            Else
                xdata = CDate(Mid(xlinha, 31, 2) & "/" & Mid(xlinha, 33, 2) & "/" & Mid(xlinha, 35, 4))
                xhora = Mid(xlinha, 39, 2) & ":" & Mid(xlinha, 41, 2)
                xStatus = Mid(xlinha, 43, 2)
                xfilialctc = Mid(xlinha, 45, 8)
                xreceb = RTrim(Mid(xlinha, 53, 25))
                lblEdiCGC = xcgc
                lblEDINf = xnf
                xserie = Mid(xlinha, 18, 3)
                lblEDIData = xdata
                lblEDIHs = xhora
                lblEDICodOcorr = xcodocorr
                If xStatus = "01" Then
                    lblEDIStatus = "Devolução/Recusa Total"
                ElseIf xStatus = "02" Then
                    lblEDIStatus = "Devolução/Recusa Parcial"
                ElseIf xStatus = "03" Then
                    lblEDIStatus = "Aceite/Entrega por Acordo"
                End If
                If xcodocorr = "01" Then
                    lblEDIReceb = xreceb
                Else
                    lblEDIReceb = "Não se Aplica"
                End If
                
                'busca ctc por CGC + NF do cliente remetente
    
                If de_informa.rsSel_CTC_NFeCGC.State = 1 Then de_informa.rsSel_CTC_NFeCGC.Close
                If Val(lblEDINf) = 0 Then lblEDINf = -999999
                de_informa.Sel_CTC_NFeCGC xcgc, Val(lblEDINf)
                If de_informa.rsSel_CTC_NFeCGC.RecordCount < 1 Then  'se nao encontra NF ...
                    lblFilialCTC = ""
                    lblEmissao = ""
                    lblOriCli = ""
                    lblOriCid = ""
                    lblOriUf = ""
                    lblDestCid = ""
                    lblDestCli = ""
                    lblDestUf = ""
                    'Atualiza o Recorset que traz as ocorrências para zero
                    If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                    de_informa.Sel_ConsOcorr2 lblFilialCTC, "01"
                    Set gridOcorr.DataSource = de_informa
                    gridOcorr.DataMember = "Sel_ConsOcorr2"
                    gridOcorr.Refresh
                    'Atualiza o Recorset que traz as baixas para zero
                    If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                    de_informa.Sel_ConsOcorr lblFilialCTC, "01"
                    lblDtEntrega = ""
                    lblHsEntrega = ""
                    lblRecEntrega = ""
                    lblBxEntrega = ""
                Else   'NF encontrada
                    'se foi encontrada mais de uma NF para este cliente traz a última NF (pela data do CTC) _
                     desde que seja anterior a ocorrência do EDI
                    If de_informa.rsSel_CTC_NFeCGC.RecordCount > 1 Then
                        de_informa.rsSel_CTC_NFeCGC.MoveLast
                        Do While True
                            If de_informa.rsSel_CTC_NFeCGC.Fields("data") > CDate(lblEDIData) Then
                                de_informa.rsSel_CTC_NFeCGC.MovePrevious
                                If de_informa.rsSel_CTC_NFeCGC.BOF Then
                                    de_informa.rsSel_CTC_NFeCGC.MoveFirst
                                    Exit Do
                                End If
                            Else
                                Exit Do
                            End If
                        Loop
                    End If
                    'Dados do CTC - Origem e Destino
                    lblFilialCTC = de_informa.rsSel_CTC_NFeCGC.Fields("filialctc")
                    lblEmissao = de_informa.rsSel_CTC_NFeCGC.Fields("data")
                    lblDestCli = de_informa.rsSel_CTC_NFeCGC.Fields("dest_nome")
                    lblDestCid = de_informa.rsSel_CTC_NFeCGC.Fields("cidade_dest")
                    lblDestUf = de_informa.rsSel_CTC_NFeCGC.Fields("uf_dest")
                    'Dados do Cliente Remetente
                    If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
                    de_informa.Sel_CadCliCGC xcgc
                    If de_informa.rsSel_CadCliCGC.RecordCount > 0 Then
                        lblOriCli = de_informa.rsSel_CadCliCGC.Fields("nome")
                        lblOriCid = de_informa.rsSel_CadCliCGC.Fields("cidade")
                        lblOriUf = de_informa.rsSel_CadCliCGC.Fields("uf")
                    Else
                        lblOriCli = ""
                        lblOriCid = ""
                        lblOriUf = ""
                    End If
                    'Dados de Ocorrência
                    If de_informa.rsSel_ConsOcorr2.State = 1 Then de_informa.rsSel_ConsOcorr2.Close
                    de_informa.Sel_ConsOcorr2 lblFilialCTC, "01"
                    Set gridOcorr.DataSource = de_informa
                    gridOcorr.DataMember = "Sel_ConsOcorr2"
                    gridOcorr.Refresh
                    'consulta que traz os campos = 01 que é dado de entrega (ENTREGA REALIZADA)
                    If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
                    de_informa.Sel_ConsOcorr lblFilialCTC, "01"
                    If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
                        'atualiza os campos referente a dados de entrega
                        If de_informa.rsSel_ConsOcorr.Fields("baixadofinal") = "S" Then
                            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixa")) = False Then lblDtEntrega = de_informa.rsSel_ConsOcorr.Fields("dtbaixa")
                            If IsNull(de_informa.rsSel_ConsOcorr.Fields("hsbaixa")) = False Then lblHsEntrega = de_informa.rsSel_ConsOcorr.Fields("hsbaixa")
                            If IsNull(de_informa.rsSel_ConsOcorr.Fields("receb")) = False Then lblRecEntrega = de_informa.rsSel_ConsOcorr.Fields("receb")
                            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_bx")) = False Then lblBxEntrega = de_informa.rsSel_ConsOcorr.Fields("usu_bx")
                        Else
                            If IsNull(de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")) = False Then lblDtEntrega = de_informa.rsSel_ConsOcorr.Fields("dtbaixapre")
                            If IsNull(de_informa.rsSel_ConsOcorr.Fields("hsbaixapre")) = False Then lblHsEntrega = de_informa.rsSel_ConsOcorr.Fields("hsbaixapre")
                            If IsNull(de_informa.rsSel_ConsOcorr.Fields("recebpre")) = False Then lblRecEntrega = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                            If IsNull(de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")) = False Then lblBxEntrega = de_informa.rsSel_ConsOcorr.Fields("usu_bxpre")
                        End If
                    Else
                        lblDtEntrega = ""
                        lblHsEntrega = ""
                        lblRecEntrega = ""
                        lblBxEntrega = ""
                    End If
                End If
                'atualiza contadores de controle do processo (Status)
                If lblEDICodOcorr = "01" Then
                    lblBxLida = Val(lblBxLida) + 1
                Else
                    If lblEDIDescOcorr = "??" Then
                        lblCrit = Val(lblCrit) + 1
                    Else
                        lblOcorrLida = Val(lblOcorrLida) + 1
                    End If
                End If
                'consistências:
                'Se já houver baixa e no registro traz uma outra baixa, verifica se esta é a
                'mesma data e hora que a que já está baixada
                If de_informa.rsSel_ConsOcorr.RecordCount > 0 And lblEDICodOcorr = "01" Then
                    'If de_informa.rsSel_ConsOcorr.Fields("data") = CDate(lblEDIData) Then
                        'se for a mesma data/hora, despreza a gravação deste registro
                        lblConfirma = "NÃO"
                    'Else
                        'se for diferente aciona form para confirmação do lançamento pelo usuário
                    '    frmEDIConf.Show 1
                    'End If
                ElseIf lblEDIDescOcorr = "??" Then  'se a ocorrência é desconhecida ...
                    lblConfirma = "NÃO"
                    'frmEDIConf.cmdConfirma.Enabled = False  'trava botão para confirmar lancamento
                    'frmEDIConf.Show 1   'exibe form de confirmação
                    'frmEDIConf.cmdConfirma.Enabled = True  'destrava botão ...
                ElseIf de_informa.rsSel_ConsOcorr2.RecordCount > 0 And lblEDICodOcorr <> "01" Then
                    'verifica se esta ocorrência já não foi cadastrada anteriormente (consiste Ocorrencia, Data e Hora)
                    de_informa.rsSel_ConsOcorr2.MoveFirst
                    Do Until de_informa.rsSel_ConsOcorr2.EOF
                        If de_informa.rsSel_ConsOcorr2.Fields("cod_ocorr") = lblEDICodOcorr And _
                        de_informa.rsSel_ConsOcorr2.Fields("data") = CDate(lblEDIData) And _
                        de_informa.rsSel_ConsOcorr2.Fields("hora") = lblEDIHs Then
                            lblConfirma = "NÃO"   'se encontrou com a mesma data/hora, despreza...
                            Exit Do
                        Else
                            lblConfirma = "SIM"
                        End If
                        de_informa.rsSel_ConsOcorr2.MoveNext
                    Loop
                Else  'não há nem baixa nem ocorrência anterior
                    If lblFilialCTC <> "" Then  'se nao é critica (NF não encontrado) ... confirma
                        lblConfirma = "SIM"
                    Else  'NF não encontrada
                        'frmEDIConf.cmdConfirma.Enabled = False
                        'frmEDIConf.Show 1
                        'frmEDIConf.cmdConfirma.Enabled = True
                        lblCrit = Val(lblCrit) + 1
                    End If
                End If
                'processo para confirmar casos em que uma baixa a ser lançada tem data anterior a
                'alguma ocorrência ou, uma ocorrência tem data posterior a uma baixa
                'se for baixa
                If lblEDICodOcorr = "01" And de_informa.rsSel_ConsOcorr2.RecordCount > 0 And lblConfirma = "SIM" Then 'é uma baixa e existe ocorrências
                    de_informa.rsSel_ConsOcorr2.MoveFirst
                    Do Until de_informa.rsSel_ConsOcorr2.EOF
                        If CDate(de_informa.rsSel_ConsOcorr2.Fields("data")) > CDate(CDate(lblEDIData)) Then 'se a data da ocorrência já cadastrada é maior que a da baixa a ser lançada ...
                            If MsgBox("ATENÇÃO ! Existe uma OCORRÊNCIA com data maior que a data desta BAIXA. Você confirma o lançamento deste dado no Sistema ?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
                                lblConfirma = "SIM"
                            Else
                                lblConfirma = "NÃO"
                            End If
                        End If
                        de_informa.rsSel_ConsOcorr2.MoveNext
                    Loop
                'se for ocorrência
                ElseIf lblEDICodOcorr <> "01" And de_informa.rsSel_ConsOcorr.RecordCount > 0 And lblConfirma = "SIM" Then 'é ocorrência e existe baixa
                    If CDate(de_informa.rsSel_ConsOcorr.Fields("data")) < CDate(lblEDIData) Then 'data da ocorrência a ser lançada é maior que a baixa ...
                        If MsgBox("ATENÇÃO ! Este OCORRÊNCIA tem data superior a data que este CTC foi BAIXADO. Você confirma o lançamento deste dado no Sistema ?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
                            lblConfirma = "SIM"
                        Else
                            lblConfirma = "NÃO"
                        End If
                    End If
                End If
                
    
                If de_informa.rsSel_CTC_NFeCGC.RecordCount > 0 Then
                    'caso em que um CTC já está baixado SEM ENTREGA (CodOcorr 00) e tenta-se baixar como ENTREGUE
                    If de_informa.rsSel_CTC_NFeCGC.Fields("tem_ocorr") = "0" And lblEDICodOcorr = "01" And lblConfirma = "SIM" Then
                        MsgBox "Não é possível lançar POD/ENTREGA (Cod. 01) em um CTC Baixado Sem Entrega (Cod. 00)."
                        lblConfirma = "NÃO"
                    End If
                    'data da ocorrência anterior a emissão do CTC
                    If CDate(lblEmissao) > CDate(lblEDIData) And lblConfirma = "SIM" Then
                        MsgBox "Não é possível lançar esta ocorrência pois a mesma é anterior a emissão do CTC !"
                        lblConfirma = "NÃO"
                    End If
                    'transp_sub = "?" = Intec ou Representante
                    If de_informa.rsSel_CTC_NFeCGC.Fields("transp_sub") = "?" And lblConfirma = "SIM" Then
                        If MsgBox("Este CTC não tem a Indicação de que foi redespachado via Tráfego Mútuo. Você confirma a inclusão desta ocorrência no Sistema ?", vbYesNo) = vbNo Then
                            'MsgBox "Esta Ocorrência não será lançada pois não há indicação de Tráfego Mútuo no CTC da Intec"
                            lblConfirma = "NÃO"
                        End If
                    End If
                    'confirmando a transportadora
                    If lblConfirma = "SIM" And de_informa.rsSel_CTC_NFeCGC.Fields("transp_sub") <> "?" Then
                        If de_informa.rsSel_BuscaTrafMutuo.State = 1 Then de_informa.rsSel_BuscaTrafMutuo.Close
                        de_informa.Sel_BuscaTrafMutuo de_informa.rsSel_CTC_NFeCGC.Fields("transp_sub") & "%"
                        If de_informa.rsSel_BuscaTrafMutuo.RecordCount >= 1 Then
                            If Trim$(xtranspcgc) <> Trim$(de_informa.rsSel_BuscaTrafMutuo.Fields("cgc")) Then
                                If de_informa.rsSel_CTC_NFeCGC.Fields("remet_cgc") <> "04229761000413" Then
                                    MsgBox "Esta Ocorrência Não Será Lançada Pois há Inconsistência do Transportador SubContratado (CGC do CTC da Intec / CGC do Arquivo EDI) ! Avisar a Área de SAC."
                                    lblConfirma = "NÃO"
                                End If
                            End If
                        Else
                            If de_informa.rsSel_BuscaSubContra.State = 1 Then de_informa.rsSel_BuscaSubContra.Close
                            de_informa.Sel_BuscaSubContra de_informa.rsSel_CTC_NFeCGC.Fields("transp_sub") & "%"
                            If de_informa.rsSel_BuscaSubContra.RecordCount >= 1 Then
                                If Trim$(xtranspcgc) <> Mid$(Trim$(de_informa.rsSel_BuscaSubContra.Fields("cgc")), 1, 8) Then
                                    If de_informa.rsSel_CTC_NFeCGC.Fields("remet_cgc") <> "04229761000413" Then
                                        MsgBox "Esta Ocorrência Não Será Lançada Pois há Inconsistência do Transportador SubContratado (CGC do CTC da Intec / CGC do Arquivo EDI) ! Avisar a Área de SAC."
                                        lblConfirma = "NÃO"
                                    End If
                                End If
                            Else
                                MsgBox "Tráfego Mutuo do CTC não Encontrado na Tabela de Tráfego Mutuo ! Chame Suporte."
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                
                'efetua os lançamentos
                
                If lblConfirma = "SIM" And Mid$(lblRemEDI, 1, 10) = "TRANVALLE" Then
                    If lblDestUf = "RS" Or lblDestUf = "PR" Or lblDestUf = "SC" Then
                    Else
                        If MsgBox("TRANSVALLE: Nao é para SUL. Confirma ?", vbYesNo) = vbNo Then lblConfirma = "NÃO"
                    End If
                End If
                
                If lblConfirma = "SIM" Then
                    
                    If lblEDICodOcorr <> "01" Then  'Tratar confirmação do lancamento: OCORRÊNCIA
                        'inicia a transação
                        de_informa.cn_informa.BeginTrans
                        de_informa.ins_ocorr4 lblFilialCTC, de_informa.rsSel_CTC_NFeCGC.Fields("data"), lblEdiCGC, lblEDICodOcorr, lblEDIDescOcorr, _
                        CDate(lblEDIData), lblEDIHs, "EDI-" & Mid$(lblRemEDI, 1, 6), datahora("datahora")
                        lblOcorrGrav = Val(lblOcorrGrav) + 1
                        de_informa.Alt_AtClienteNFBranco lblFilialCTC
                        If de_informa.rsSel_CTC_NFeCGC.Fields("tem_ocorr") = "0" Or _
                            de_informa.rsSel_CTC_NFeCGC.Fields("tem_ocorr") = "1" Or _
                            de_informa.rsSel_CTC_NFeCGC.Fields("tem_ocorr") = "C" Then
                            
                            'abono automático de atraso na entrega
                            If de_informa.rsSel_CTC_NFeCGC.Fields("tem_ocorr") = "1" Then
                                If lblEDICodOcorr = "26" Or lblEDICodOcorr = "85" Then  'abono automático para atraso
                                    If de_informa.rsSel_CTCEntrega.State = 1 Then de_informa.rsSel_CTCEntrega.Close
                                    de_informa.Sel_CTCEntrega lblFilialCTC
                                    If de_informa.rsSel_CTCEntrega.RecordCount > 0 Then
                                        If de_informa.rsSel_CTCEntrega.Fields("diasuteis") - _
                                        de_informa.rsSel_CTCEntrega.Fields("abonodias") > _
                                        de_informa.rsSel_CTCEntrega.Fields("prazoentr") Then
                                            'está em atraso, lançar abono automático
                                            xabonodias = de_informa.rsSel_CTCEntrega.Fields("diasuteis") - de_informa.rsSel_CTCEntrega.Fields("prazoentr")
                                            de_informa.Alt_AbonoAtraso xabonodias, "AUTOMATIC", datahora("DATAHORA"), "Abono Automático Devido Ocorrência", lblFilialCTC
                                        End If
                                    End If
                                End If
                            End If
                            
                        Else
                            de_informa.alt_temocorr_sn "2", lblFilialCTC   'atualiza arquivo de CTC com tem_ocorr = 2
                            If lblEDICodOcorr = "39" Or lblEDICodOcorr = "84" Then  'pre-baixa automática por ser CTC/NF Retido para COnferência
                                de_informa.ins_ocorr1 lblFilialCTC, de_informa.rsSel_CTC_NFeCGC.Fields("data"), lblEdiCGC, "01", "ENTREGA REALIZADA", _
                                CDate(lblEDIData), lblEDIHs, CDate(lblEDIData), lblEDIHs, ".", "AUTO-PREBX", datahora("datahora"), "S", datahora("data")
                                de_informa.alt_temocorr_sn "1", lblFilialCTC
                            End If
                        End If
                        de_informa.cn_informa.CommitTrans
                    Else   'Tratar confirmação do lancamento: ENTREGA
                        If Not IsDate(lblEDIData) Then
                            MsgBox "Campo Data Inválido ! Lançamento não Realizado !"
                        Else
                            'verifica se o CTC já está baixado ou se é nova baixa
                            If lblDtEntrega = "" Then  'Nova Baixa ...
                                'inicia a transação
                                de_informa.cn_informa.BeginTrans
                                de_informa.ins_ocorr1 lblFilialCTC, de_informa.rsSel_CTC_NFeCGC.Fields("data"), lblEdiCGC, lblEDICodOcorr, lblEDIDescOcorr, _
                                CDate(lblEDIData), lblEDIHs, CDate(lblEDIData), lblEDIHs, lblEDIReceb & "", "EDI-" & Mid$(lblRemEDI, 1, 6), datahora("datahora"), "S", datahora("data")
                                de_informa.alt_temocorr_sn "1", lblFilialCTC
                                de_informa.Alt_AtClienteNFBranco lblFilialCTC
                                lblBxGrav = Val(lblBxGrav) + 1
                                de_informa.cn_informa.CommitTrans
                            Else   'já está baixado e baixa por cima (overwrite)
                                'inicia a transação
                                de_informa.cn_informa.BeginTrans
                                de_informa.alt_ocorr1 lblFilialCTC, CDate(lblEDIData), lblEDIHs, CDate(lblEDIData), lblEDIHs, lblEDIReceb & "", "EDI-" & Mid$(lblRemEDI, 1, 6), datahora("datahora"), "S", datahora("data")
                                de_informa.Alt_AtClienteNFBranco lblFilialCTC
                                de_informa.alt_temocorr_sn "1", lblFilialCTC
                                lblBxGrav = Val(lblBxGrav) + 1
                                de_informa.cn_informa.CommitTrans
                            End If
                        End If
                    End If
                    'atualiza contador de totais gravados
                    lblTotGrav = Val(lblBxGrav) + Val(lblOcorrGrav)
                ElseIf lblConfirma = "NÃO" Then
                    'Tratar não confirmação do lancamento
                ElseIf lblConfirma = "CANCELAR" Then
                    Close #1
                    DirImporta.Enabled = True
                    fileImport.Enabled = True
                    txtArq.Enabled = True
                    cmdProcessa.Enabled = True
                    cmdSair.Enabled = True
                    Exit Sub
                End If
            End If
        End If
        lblConfirma = ""
        DoEvents
    Loop
    Close #1 'fecha arquivo

    'LOG DE USUÁRIO
    de_informa.ins_LogUsuario "PROCESSO", xusuario, "LEITURA DE EDI DE OCORRENCIAS/POD: " & lblRemEDI

    'If Dir(DirImporta.Path & "\backup") = "" Then  'se o diretório de backup existe ...
    'Else
        FileCopy DirImporta.Path & "\" & txtArq.Text, DirImporta.Path & "\backup\" & txtArq.Text
        Kill DirImporta.Path & "\" & txtArq.Text
        fileImport.Refresh
    'End If
    MsgBox "Arquivo Processado !"
    txtArq = ""
    lblFilialCTC = ""
    lblEmissao = ""
    lblOriCli = ""
    lblOriCid = ""
    lblOriUf = ""
    lblDestCid = ""
    lblDestCli = ""
    lblDestUf = ""
    'Set gridOcorr.DataSource = Null
    gridOcorr.DataMember = ""
    gridOcorr.Refresh
    lblDtEntrega = ""
    lblHsEntrega = ""
    lblRecEntrega = ""
    lblBxEntrega = ""
    lblRemEDI = ""
    lblIdent = ""
    lblReg = ""
    lblEdiCGC = ""
    lblEDINf = ""
    lblEDIData = ""
    lblEDIHs = ""
    lblEDICodOcorr = ""
    lblEDIDescOcorr = ""
    lblEDIStatus = ""
    lblEDIReceb = ""
    lblReg = ""
    lblTotReg = ""
    lblOcorrGrav = ""
    lblOcorrLida = ""
    lblBxLida = ""
    lblBxGrav = ""
    lblTotGrav = ""
    lblCrit = ""
    DirImporta.Enabled = True
    fileImport.Enabled = True
    txtArq.Enabled = True
    cmdProcessa.Enabled = True
    cmdSair.Enabled = True
End Sub

Private Sub cmdSair_Click()
    frmAtualPrazos.Show 1
    Set frmEdiImport = Nothing
    Set frmEDIConf = Nothing
    Unload frmEDIConf
    Unload Me
End Sub

Private Sub DirImporta_Change()
    fileImport.Path = (DirImporta.Path)  'Quando mudado o diretório, atualiza o path do Dir
    fileImport.Refresh
End Sub

Private Sub fileImport_Click()
txtArq.Text = fileImport.FileName
End Sub

Private Sub Form_Load()
DirImporta.Path = "C:\INFORMA\EDI_IMP"
End Sub

Private Sub lblRegOcorr_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEdiImport = Nothing
End Sub
