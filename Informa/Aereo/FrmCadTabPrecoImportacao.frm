VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCadTabPrecoImportacao 
   Caption         =   "Importação de Planilhas"
   ClientHeight    =   5715
   ClientLeft      =   825
   ClientTop       =   1815
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9510
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdNovo 
      Caption         =   "Novo Processamento"
      Height          =   435
      Left            =   6420
      TabIndex        =   16
      Top             =   1520
      Width           =   2895
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   6420
      TabIndex        =   15
      Top             =   2100
      Width           =   2895
   End
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar Dados"
      Enabled         =   0   'False
      Height          =   435
      Left            =   6420
      TabIndex        =   14
      Top             =   940
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colunas"
      Height          =   1995
      Left            =   120
      TabIndex        =   10
      Top             =   4620
      Width           =   11715
      Begin VB.CommandButton CmdColuna 
         Caption         =   "Cruza Coluna"
         Enabled         =   0   'False
         Height          =   1155
         Left            =   5340
         TabIndex        =   11
         Top             =   480
         Width           =   1035
      End
      Begin MSFlexGridLib.MSFlexGrid FlexGridColunaDisp 
         Height          =   1635
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2884
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid FlexGridColunaCorr 
         Height          =   1635
         Left            =   6420
         TabIndex        =   13
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2884
         _Version        =   393216
         Enabled         =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Localidades"
      Height          =   1995
      Left            =   120
      TabIndex        =   6
      Top             =   2580
      Width           =   11715
      Begin VB.CommandButton CmdLocalidade 
         Caption         =   "Cruzar Localidade"
         Enabled         =   0   'False
         Height          =   1155
         Left            =   5340
         TabIndex        =   8
         Top             =   480
         Width           =   1035
      End
      Begin MSFlexGridLib.MSFlexGrid FlexGridLocalidadeDisp 
         Height          =   1635
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2884
         _Version        =   393216
         Enabled         =   0   'False
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid FlexGridLocalidadeCorr 
         Height          =   1635
         Left            =   6420
         TabIndex        =   9
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2884
         _Version        =   393216
         Enabled         =   0   'False
         SelectionMode   =   1
      End
   End
   Begin VB.CommandButton CmdProcessar 
      Caption         =   "Processar Planilha"
      Height          =   435
      Left            =   6420
      TabIndex        =   5
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox TxtArquivo 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   6195
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2220
      Width           =   3075
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3075
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   3075
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Arquivo"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "FrmCadTabPrecoImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xColunas As Long
Public xLinhas As Long

Public Excel As Excel.Application
Public ExcelWBK As Excel.Workbook
Public ExcelWS As Excel.Worksheet
Public ExcelWS2 As Excel.Worksheet
Public ExcelWS3 As Excel.Worksheet
Public ExcelWS4 As Excel.Worksheet
Public ExcelA1 As Excel.Worksheet
Public ExcelA2 As Excel.Worksheet
Public ExcelA3 As Excel.Worksheet


Private Sub CmdCancelar_Click()
Excel.Quit
Unload Me
End Sub

Private Sub CmdColuna_Click()
xtransf = True
    For Y = 1 To FlexGridColunaCorr.Rows - 1
        If FlexGridColunaCorr.TextMatrix(Y, 1) = FlexGridColunaDisp.TextMatrix(FlexGridColunaDisp.Row, 0) Then
        xtransf = False
        End If
    Next
    
    If xtransf = False Then
    MsgBox "Não é possível inserir este valor para esta coluna, visto que já existe uma coluna definida como " & FlexGridColunaDisp.TextMatrix(FlexGridColunaDisp.Row, 0), vbCritical, ""
    Else
    FlexGridColunaCorr.TextMatrix(FlexGridColunaCorr.Row, 1) = FlexGridColunaDisp.TextMatrix(FlexGridColunaDisp.Row, 0)
    End If
End Sub

Private Sub CmdImportar_Click()

    For X = 1 To FlexGridLocalidadeCorr.Rows - 1
        If Len(Trim(FlexGridLocalidadeCorr.TextMatrix(X, 1))) = 0 Then
        MsgBox "Não é possível importar esta planilha visto que existe Localidades em branco.", vbCritical, ""
        Exit Sub
        End If
    Next
    
    For X = 1 To FlexGridColunaCorr.Rows - 1
        If Len(Trim(FlexGridColunaCorr.TextMatrix(X, 1))) = 0 Then
        MsgBox "Não é possível importar esta planilha visto que existe Colunas em branco.", vbCritical, ""
        Exit Sub
        End If
    Next
Me.MousePointer = 11
CmdImportar.Enabled = False
DoEvents
frmCadTabPrecoINCLUSAO.FlexGridImportacao.Clear
frmCadTabPrecoINCLUSAO.FlexGridImportacao.Rows = FlexGridLocalidadeCorr.Rows
frmCadTabPrecoINCLUSAO.FlexGridImportacao.Cols = FlexGridColunaCorr.Rows - 1

    For Y = 1 To FlexGridColunaCorr.Rows - 1
    frmCadTabPrecoINCLUSAO.FlexGridImportacao.TextMatrix(0, Y - 1) = FlexGridColunaCorr.TextMatrix(Y, 1)
    Next
    
    For Y = 1 To FlexGridLocalidadeCorr.Rows - 1
    frmCadTabPrecoINCLUSAO.FlexGridImportacao.TextMatrix(Y, 0) = FlexGridLocalidadeCorr.TextMatrix(Y, 1)
    Next

'Excel.Visible = True
'Excel.Interactive = True
    
    For Y = 2 To xLinhas
        For X = 2 To xColunas
        frmCadTabPrecoINCLUSAO.FlexGridImportacao.TextMatrix(0, X - 1) = frmCadTabPrecoINCLUSAO.FlexGridImportacao.TextMatrix(0, X - 1)
        frmCadTabPrecoINCLUSAO.FlexGridImportacao.TextMatrix(Y - 1, X - 1) = Format(ExcelA1.Cells(Y, X), "##,##0.00")
        Next
    Next

Excel.Quit
Unload Me
End Sub

Private Sub CmdLocalidade_Click()
xtransf = True
    For Y = 1 To FlexGridLocalidadeCorr.Rows - 1
        If FlexGridLocalidadeCorr.TextMatrix(Y, 1) = FlexGridLocalidadeDisp.TextMatrix(FlexGridLocalidadeDisp.Row, 0) Then
        xtransf = False
        End If
    Next
    
    If xtransf = False Then
    MsgBox "Não é possível inserir este valor para esta localidade, visto uma que já existe uma localidade definida como " & FlexGridLocalidadeDisp.TextMatrix(FlexGridLocalidadeDisp.Row, 0), vbCritical, ""
    Else
    FlexGridLocalidadeCorr.TextMatrix(FlexGridLocalidadeCorr.Row, 1) = FlexGridLocalidadeDisp.TextMatrix(FlexGridLocalidadeDisp.Row, 0)
    End If
End Sub

Private Sub CmdNovo_Click()
CmdLocalidade.Enabled = False
CmdColuna.Enabled = False
CmdCancelar.Enabled = True
CmdImportar.Enabled = False
CmdProcessar.Enabled = True
TxtArquivo.Enabled = True
File1.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
FlexGridLocalidadeCorr.Rows = 0
FlexGridLocalidadeDisp.Rows = 0
FlexGridColunaCorr.Rows = 0
FlexGridColunaDisp.Rows = 0
FlexGridLocalidadeCorr.Enabled = False
FlexGridLocalidadeDisp.Enabled = False
FlexGridColunaCorr.Enabled = False
FlexGridColunaDisp.Enabled = False
DoEvents
End Sub

Private Sub CmdProcessar_Click()

CmdLocalidade.Enabled = False
CmdColuna.Enabled = False
CmdCancelar.Enabled = False
CmdImportar.Enabled = False
CmdProcessar.Enabled = False
TxtArquivo.Enabled = False
File1.Enabled = False
Dir1.Enabled = False
Drive1.Enabled = False

If Mid(File1.Path, Len(File1.Path)) = "\" Then xarquivo1 = File1.Path & TxtArquivo.Text & ".xls"
If Mid(File1.Path, Len(File1.Path)) <> "\" Then xarquivo1 = File1.Path & "\" & TxtArquivo.Text & ".xls"

Set Excel = CreateObject("EXCEL.APPLICATION")
Excel.Visible = False
Excel.Interactive = False
Excel.DisplayAlerts = False

Excel.Workbooks.OpenText FileName:=xarquivo1, Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
Other:=True, OtherChar:="#", FieldInfo:=Array(1, 1)
Set ExcelA1 = Excel.Worksheets(1)

xColunas = 0
    Do While True
    xColunas = xColunas + 1
        If Len(Trim(ExcelA1.Cells(1, xColunas))) = 0 Then
        xColunas = xColunas - 1
        Exit Do
        End If
    Loop

xLinhas = 0
    Do While True
    xLinhas = xLinhas + 1
        If Len(Trim(ExcelA1.Cells(xLinhas, 1))) = 0 Then
        xLinhas = xLinhas - 1
        Exit Do
        End If
    Loop

If de_informa.rsSel_CadLocalAirGROUP.State = 1 Then de_informa.rsSel_CadLocalAirGROUP.Close
de_informa.Sel_CadLocalAirgroup

FlexGridLocalidadeDisp.Clear
FlexGridLocalidadeDisp.Rows = de_informa.rsSel_CadLocalAirGROUP.RecordCount + 1
FlexGridLocalidadeDisp.Cols = 1
FlexGridLocalidadeDisp.FixedRows = 1
FlexGridLocalidadeDisp.FixedCols = 0
FlexGridLocalidadeDisp.ColWidth(0) = 4000
Y = 0
FlexGridLocalidadeDisp.TextMatrix(Y, 0) = "Localidade Correta"

    Do Until de_informa.rsSel_CadLocalAirGROUP.EOF
    Y = Y + 1
    FlexGridLocalidadeDisp.TextMatrix(Y, 0) = PriMaiuscula(de_informa.rsSel_CadLocalAirGROUP.Fields("localidade")) ' & " - " & de_informa.rsSel_CadLocalAirgroup.Fields("SIGLA")
    de_informa.rsSel_CadLocalAirGROUP.MoveNext
    Loop
    

If de_informa.rsSel_CadIATA.State = 1 Then de_informa.rsSel_CadIATA.Close
de_informa.Sel_Cadiata "%"
        
FlexGridColunaDisp.Clear
FlexGridColunaDisp.Rows = de_informa.rsSel_CadIATA.RecordCount + 9
FlexGridColunaDisp.Cols = 1
FlexGridColunaDisp.FixedRows = 1
FlexGridColunaDisp.FixedCols = 0
FlexGridColunaDisp.ColWidth(0) = 4000
FlexGridColunaDisp.TextMatrix(0, 0) = "Coluna Correta"
FlexGridColunaDisp.TextMatrix(1, 0) = "Localidades"
FlexGridColunaDisp.TextMatrix(2, 0) = "Taxa Mínima"
FlexGridColunaDisp.TextMatrix(3, 0) = "Até 25,5"
FlexGridColunaDisp.TextMatrix(4, 0) = "Até 50,5"
FlexGridColunaDisp.TextMatrix(5, 0) = "Até 300,5"
FlexGridColunaDisp.TextMatrix(6, 0) = "Até 500,5"
FlexGridColunaDisp.TextMatrix(7, 0) = "Até 1000,5"
FlexGridColunaDisp.TextMatrix(8, 0) = "Acima de 1000,5"

Y = 8

    Do Until de_informa.rsSel_CadIATA.EOF
    Y = Y + 1
        If de_informa.rsSel_CadIATA.Fields("codigo") <> "000" Then
        FlexGridColunaDisp.TextMatrix(Y, 0) = "Cód. " & de_informa.rsSel_CadIATA.Fields("codigo")
        End If
    de_informa.rsSel_CadIATA.MoveNext
    Loop
    
    
FlexGridColunaCorr.Clear
FlexGridColunaCorr.Rows = xColunas + 1
FlexGridColunaCorr.Cols = 2
FlexGridColunaCorr.FixedRows = 1
FlexGridColunaCorr.FixedCols = 1
FlexGridColunaCorr.ColWidth(0) = 2350
FlexGridColunaCorr.ColWidth(1) = 2350
Y = 0
FlexGridColunaCorr.TextMatrix(Y, 0) = "Col. Planilha"
FlexGridColunaCorr.TextMatrix(Y, 1) = "Col. Correta"

    For X = 1 To xColunas
    Y = Y + 1
 FlexGridColunaCorr.TextMatrix(Y, 0) = UCase(Trim(ExcelA1.Cells(1, X)))
    Next


FlexGridLocalidadeCorr.Clear
FlexGridLocalidadeCorr.Rows = xLinhas
FlexGridLocalidadeCorr.Cols = 2
FlexGridLocalidadeCorr.FixedRows = 1
FlexGridLocalidadeCorr.FixedCols = 1
FlexGridLocalidadeCorr.ColWidth(0) = 2350
FlexGridLocalidadeCorr.ColWidth(1) = 2350
Y = 0
FlexGridLocalidadeCorr.TextMatrix(Y, 0) = "Local. Planilha"
FlexGridLocalidadeCorr.TextMatrix(Y, 1) = "Local. Correta"

    For X = 2 To xLinhas
    Y = Y + 1
    FlexGridLocalidadeCorr.TextMatrix(Y, 0) = UCase(Trim(ExcelA1.Cells(X, 1)))
    Next
    

    For Y = 1 To FlexGridLocalidadeCorr.Rows - 1
        For X = 1 To FlexGridLocalidadeDisp.Rows - 1
            If UCase(Trim(FlexGridLocalidadeCorr.TextMatrix(Y, 0))) = UCase(Trim(FlexGridLocalidadeDisp.TextMatrix(X, 0))) Then
            FlexGridLocalidadeCorr.TextMatrix(Y, 1) = FlexGridLocalidadeDisp.TextMatrix(X, 0)
            End If
        Next
    Next
'Excel.Quit


CmdLocalidade.Enabled = True
CmdColuna.Enabled = True
CmdCancelar.Enabled = True
CmdImportar.Enabled = True
FlexGridLocalidadeCorr.Enabled = True
FlexGridLocalidadeDisp.Enabled = True
FlexGridColunaCorr.Enabled = True
FlexGridColunaDisp.Enabled = True

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Pattern = "*.xls"
File1.Refresh
DoEvents
End Sub

Private Sub Drive1_Change()
Dir1.Path = Mid(Drive1.Drive, 1, 2) & "\"
File1.Path = Dir1.Path
File1.Pattern = "*.xls"
File1.Refresh
Dir1.Refresh
DoEvents
End Sub

Private Sub File1_Click()
TxtArquivo.Text = Mid(File1.List(File1.ListIndex), 1, Len(File1.List(File1.ListIndex)) - 4)
DoEvents
End Sub

Private Sub FlexGridColunaCorr_Click()
Call ColoreCelulaFlex(FlexGridColunaCorr)
End Sub

Private Sub FlexGridColunaDisp_Click()
Call ColoreCelulaFlex(FlexGridColunaDisp)
End Sub

Private Sub FlexGridLocalidadeCorr_Click()
Call ColoreCelulaFlex(FlexGridLocalidadeCorr)
End Sub

Private Sub FlexGridLocalidadeDisp_Click()
Call ColoreCelulaFlex(FlexGridLocalidadeDisp)
End Sub

Private Sub Form_Load()
Drive1.Drive = "c"
Dir1.Path = Drive1.Drive & "\"
File1.Path = Dir1.Path
File1.Pattern = "*.xls"
End Sub

Public Sub ColoreCelulaFlex(xFlexGridRequired As MSFlexGrid)
XCOLUNAPRIMARIA = xFlexGridRequired.Row
    For X = 1 To xFlexGridRequired.Rows - 1
        xFlexGridRequired.Row = X
        xFlexGridRequired.CellBackColor = xBranco
        xFlexGridRequired.CellForeColor = xPreto
    Next
xFlexGridRequired.Row = XCOLUNAPRIMARIA
xFlexGridRequired.CellBackColor = xAzul
xFlexGridRequired.CellForeColor = xBranco
DoEvents
End Sub
