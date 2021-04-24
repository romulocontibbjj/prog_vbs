VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{E6C4280E-288E-41E1-B348-A0E583B65166}#1.1#0"; "AnimatedGif.ocx"
Begin VB.Form frm_Rels 
   Caption         =   "Form2"
   ClientHeight    =   8610
   ClientLeft      =   2595
   ClientTop       =   1965
   ClientWidth     =   10410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8610
   ScaleWidth      =   10410
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   9975
         Begin BONAGURA.ProgBar prg_bar 
            Height          =   375
            Left            =   120
            Top             =   2400
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   661
            BackColor       =   12648447
            BarColor        =   128
            Value           =   0
         End
         Begin BONAGURA.isButton cmd_Stop 
            Height          =   420
            Left            =   4200
            TabIndex        =   17
            Top             =   1200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
            Icon            =   "Frm_rels.frx":0000
            Style           =   5
            Caption         =   "Stop"
            IconAlign       =   1
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
         End
         Begin BONAGURA.isButton isButton2 
            Height          =   375
            Left            =   4200
            TabIndex        =   16
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Icon            =   "Frm_rels.frx":001C
            Style           =   0
            Caption         =   "isButton2"
            IconAlign       =   1
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
         End
         Begin BONAGURA.isButton isButton1 
            Height          =   375
            Left            =   4200
            TabIndex        =   15
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            Icon            =   "Frm_rels.frx":0038
            Style           =   0
            Caption         =   "isButton1"
            IconAlign       =   1
            iNonThemeStyle  =   0
            Tooltiptitle    =   ""
            ToolTipIcon     =   0
            ToolTipType     =   0
            ttForeColor     =   0
         End
         Begin VB.CommandButton cmd_Stop1 
            Caption         =   "-"
            Height          =   375
            Left            =   3000
            TabIndex        =   14
            Top             =   1200
            Width           =   975
         End
         Begin AnimatedGif.AniGif AniGif1 
            Height          =   735
            Left            =   8400
            TabIndex        =   13
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
         End
         Begin VB.CommandButton cmd_Relatorio 
            Caption         =   "Relatorio"
            Height          =   375
            Left            =   3000
            TabIndex        =   5
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmd_Gerar 
            Caption         =   "Gerar"
            Height          =   375
            Left            =   3000
            TabIndex        =   4
            Top             =   120
            Width           =   975
         End
         Begin VB.Frame Frame3 
            Caption         =   "Periodo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   2655
            Begin MSMask.MaskEdBox mask_data1 
               Height          =   300
               Left            =   120
               TabIndex        =   2
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "99/99/9999"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Mask_dt2 
               Height          =   300
               Left            =   1440
               TabIndex        =   3
               Top             =   360
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "99/99/9999"
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "FIM"
               Height          =   255
               Left            =   1440
               TabIndex        =   11
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "INÍCIO"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   720
               Width           =   1095
            End
         End
         Begin VB.TextBox txt_cgc 
            Height          =   285
            Left            =   600
            TabIndex        =   1
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "CGC:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
      End
      Begin MSDataGridLib.DataGrid grid_Rels 
         Bindings        =   "Frm_rels.frx":0054
         Height          =   4335
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7646
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
         DataMember      =   "Sel_DadoPrincipais"
         ColumnCount     =   12
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "FILIALCTC"
            Caption         =   "FILIALCTC"
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
            DataField       =   "NUMNFNUM"
            Caption         =   "NUMNFNUM"
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
            DataField       =   "VALMERC"
            Caption         =   "VALMERC"
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
            DataField       =   "REMET_CGC"
            Caption         =   "REMET_CGC"
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
            DataField       =   "REMET_NOME"
            Caption         =   "REMET_NOME"
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
            DataField       =   "REMET_CIDADE"
            Caption         =   "REMET_CIDADE"
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
            DataField       =   "REMET_UF"
            Caption         =   "REMET_UF"
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
            DataField       =   "DEST_NOME"
            Caption         =   "DEST_NOME"
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739,906
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
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column11 
            EndProperty
         EndProperty
      End
      Begin VB.Label lab_Qtd_reg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   7920
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   5055
         Left            =   0
         Top             =   3360
         Width           =   9975
      End
   End
End
Attribute VB_Name = "frm_Rels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_gerar_Click()
Dim x As Integer

prg_bar.Min = 0
prg_bar.Value = 0

On Error GoTo Trata_erro
Trata_erro:
    If Err.Number = -2147217871 Then
        MsgBox "TEMPO EXCEDIDO"
        Exit Sub
    End If
    

With deb_Relatorios
DoEvents
If .rsSel_DadoPrincipais.State = 1 Then .rsSel_DadoPrincipais.Close
    .Sel_DadoPrincipais CDate(mask_data1), CDate(Mask_dt2), Trim$(txt_cgc.Text)
    
    If .rsSel_DadoPrincipais.RecordCount = 0 Then
        
        MsgBox "Sem Dados neste Período"
    
    Else
        lab_Qtd_reg.Caption = .rsSel_DadoPrincipais.RecordCount
        grid_Rels.DataMember = "Sel_DadoPrincipais"
        grid_Rels.Refresh
        
        prg_bar.Max = .rsSel_DadoPrincipais.RecordCount
            
        cmd_Relatorio.SetFocus
    
    End If
End With


End Sub

Private Sub cmd_Relatorio_Click()
Dim z As Integer

Dim xData As String
Dim xfilialctc As String
Dim xNUMNFNUM As String
Dim xvalmerc As Currency
Dim xremet_cgc As String
Dim xremet_nome As String
Dim xREMET_CIDADE As String
Dim xREMET_UF As String
Dim xDEST_NOME As String
Dim xCIDADE_dest As String
Dim xUF_dest As String
Dim xModal As String

Dim xMotorista As String
Dim xCPF As String
Dim xplaca As String
Dim x As Integer
Dim y As Integer


Dim Excel As Excel.Application
Dim ExcelWBk As Excel.Workbook
Dim ExcelA1 As Excel.Worksheet

        
        Set Excel = CreateObject("Excel.Application")
        Set Excel = GetObject(, "Excel.Application")
        Excel.Visible = True
        Excel.Interactive = True
        Set ExcelWBk = Excel.Workbooks.Add
        Set ExcelA1 = Excel.Worksheets(1)

x = 3

With Excel
    

    .Cells(x, 1) = "DATA"
    .Cells(x, 2) = "FILIALCTC"
    .Cells(x, 3) = "NUMNFNUM"
    .Cells(x, 4) = "VALMERC"
    .Cells(x, 5) = "REMET_CGC"
    .Cells(x, 6) = "REMET_NOME"
    .Cells(x, 7) = "REMET_CIDADE"
    .Cells(x, 8) = "REMET_UF"
    .Cells(x, 9) = "DEST_NOME"
    .Cells(x, 10) = "CIDADE_DEST"
    .Cells(x, 11) = "UF_DEST"
    .Cells(x, 12) = "MODAL"
    .Cells(x, 13) = "MOTORISTA"
    .Cells(x, 14) = "CPF"
    .Cells(x, 15) = "PLACA"

    


End With

x = 4

With deb_Relatorios
    
    .rsSel_DadoPrincipais.MoveFirst
    
    Do Until .rsSel_DadoPrincipais.EOF
    
    z = z + 1
    
    prg_bar.Value = z
    
        xfilialctc = grid_Rels.Columns(1)
        DoEvents
        If .rsSel_Dadosos1.State = 1 Then .rsSel_Dadosos1.Close
            .Sel_Dadosos1 xfilialctc
        
            If .rsSel_Dadosos1.RecordCount > 0 Then
                .rsSel_Dadosos1.MoveFirst
                xMotorista = .rsSel_Dadosos1.Fields("MOTORISTA")
                xCPF = .rsSel_Dadosos1.Fields("CPF")
                xplaca = .rsSel_Dadosos1("PLACA")
                
                
                xData = CDate(grid_Rels.Columns(0))
                xfilialctc = grid_Rels.Columns(1)
                xNUMNFNUM = grid_Rels.Columns(2)
                xvalmerc = grid_Rels.Columns(3)
                xremet_cgc = grid_Rels.Columns(4)
                xremet_nome = grid_Rels.Columns(5)
                xREMET_CIDADE = grid_Rels.Columns(6)
                xREMET_UF = grid_Rels.Columns(7)
                xDEST_NOME = grid_Rels.Columns(8)
                xCIDADE_dest = grid_Rels.Columns(9)
                xUF_dest = grid_Rels.Columns(10)
                xModal = grid_Rels.Columns(11)

                
                
                Excel.Cells(x, 1) = CDate(xData)
                Excel.Cells(x, 2) = "'" & xfilialctc
                Excel.Cells(x, 3) = xNUMNFNUM
                Excel.Cells(x, 4) = xvalmerc
                Excel.Cells(x, 5) = "'" & xremet_cgc
                Excel.Cells(x, 6) = xremet_nome
                Excel.Cells(x, 7) = xREMET_CIDADE
                Excel.Cells(x, 8) = xREMET_UF
                Excel.Cells(x, 9) = xDEST_NOME
                Excel.Cells(x, 10) = xCIDADE_dest
                Excel.Cells(x, 11) = xUF_dest
                Excel.Cells(x, 12) = xModal
                Excel.Cells(x, 13) = xMotorista
                Excel.Cells(x, 14) = "'" & xCPF
                Excel.Cells(x, 15) = UCase(xplaca)
                
                x = x + 1
DoEvents
                
            End If
        
        .rsSel_DadoPrincipais.MoveNext
    Loop
    
    ExcelA1.Range("A:DZ").EntireColumn.AutoFit
    
    Excel.Cells(1, 1) = "PERÍODO:(" & CDate(mask_data1) & " - " & CDate(Mask_dt2) & ")"
    Excel.Range(Excel.Cells(1, 1), Excel.Cells(3, 15)).Font.Bold = True
    Excel.Range(Excel.Cells(3, 1), Excel.Cells(3, 15)).Borders.ColorIndex = 1
    
MsgBox x - 3

End With

Set ExcelA1 = Nothing
Set ExcelWBk = Nothing



End Sub

Private Sub cmd_Stop_Click()
AniGif1.FinishCycleFast
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = &H80000012
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = &HFF&
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
        KeyCode = 0
    Unload Me
    
ElseIf KeyCode = 13 Then
    SendKeys "{TAB}"
    KeyCode = 0
End If



End Sub

Private Sub Form_Load()
AniGif1.BackColor = &H8000000F
AniGif1.LoadFile "C:\Documents and Settings\rconti\Meus documentos\Minhas imagens\animais.gif", False
End Sub
