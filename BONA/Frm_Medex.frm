VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Frm_Medex 
   Caption         =   "Relatório Medex"
   ClientHeight    =   2505
   ClientLeft      =   5610
   ClientTop       =   4305
   ClientWidth     =   3195
   LinkTopic       =   "Form2"
   ScaleHeight     =   2505
   ScaleWidth      =   3195
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.CheckBox chk_excel 
         Caption         =   "Ver Excel"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mask_data2 
         Height          =   300
         Left            =   1800
         TabIndex        =   5
         Top             =   1320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mask_data1 
         Height          =   300
         Left            =   720
         TabIndex        =   4
         Top             =   1320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_Cgc 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Text            =   "06019570000100"
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmd_gerar 
         Caption         =   "Gera Arquivo"
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Data Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Data Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "CGC:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "Frm_Medex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_gerar_Click()
Dim Excel As Excel.Application
Dim ExcelWBk As Excel.Workbook
Dim ExcelA1 As Excel.Worksheet
Dim xData As String

Dim xmax As Integer
Dim xlinha As Integer
Dim x As Integer
Dim y As Integer

Dim xfilialctc As String
Dim xnfs As String
Dim xremet_cgc As String
Dim xremet_nome As String
Dim xREMET_CIDADE As String
Dim xREMET_UF As String
Dim xDEST_NOME As String
Dim xCIDADE_dest As String
Dim xUF_dest As String
Dim xplacaveic As String
Dim xarquivo As String
Dim xenvia As String
Dim xData As Single
Dim xMotorista As String
Dim xvalmerc As Currency
Dim xModal As String
Dim xCPF As String


'Data
'FILIALCTC
'NOTAFISCAL
'VALOR
'REMET_CGC
'REMET_NOME
'REMET_CIDADE
'REMET_UF
'DEST_NOME
'CIDADE_DEST
'UF_DEST
'PLACAVEIC
'MOTORISTA
'CPF
'MOTORISTA
'MODAL






If deb_fnac.rsSel_MEDEX.State = 1 Then deb_fnac.rsSel_MEDEX.Close
    deb_fnac.Sel_MEDEX txt_cgc, CDate(mask_data1), CDate(mask_data2)
    
    If deb_fnac.rsSel_MEDEX.RecordCount > 0 Then
        
        Set Excel = CreateObject("Excel.Application")
        Set Excel = GetObject(, "Excel.Application")
        
        If chk_excel.Value = 0 Then
            Excel.Visible = False
        Else
            Excel.Visible = True
        End If
        
        Excel.Interactive = True
        
        Set ExcelWBk = Excel.Workbooks.Add
        Set ExcelA1 = Excel.Worksheets(1)
    
        ExcelA1.Name = "MEDEX"
        
        
        With deb_fnac.rsSel_MEDEX
        
        ExcelA1.Cells.Font.Name = "Verdana"
        
        Excel.Cells(1, 1) = "Relatório Entregas"
        Excel.Cells(2, 1) = "Cliente: " & .Fields("remet_nome")
        Excel.Cells(3, 1) = "Período: (" & CDate(mask_data1) & " a " & CDate(mask_data2) & ")"
        
        Excel.Range(Excel.Cells(1, 1), Excel.Cells(5, 15)).Font.Bold = True
        
        Excel.Cells(5, 1) = "DATA"
        Excel.Cells(5, 2) = "FILIALCTC"
        Excel.Cells(5, 3) = "NOTA FISCAL"
        Excel.Cells(5, 4) = "VALOR"
        Excel.Cells(5, 5) = "REMET_CGC"
        Excel.Cells(5, 6) = "REMET_NOME"
        Excel.Cells(5, 7) = "REMET_CIDADE"
        Excel.Cells(5, 8) = "REMET_UF"
        Excel.Cells(5, 9) = "DEST_NOME"
        Excel.Cells(5, 10) = "CIDADE_DEST"
        Excel.Cells(5, 11) = "UF_DEST"
        Excel.Cells(5, 12) = "PLACAVEIC"
        Excel.Cells(5, 13) = "MOTORISTA"
        Excel.Cells(5, 14) = "CPF"
        Excel.Cells(5, 15) = "MODAL"
        
'Data
'FILIALCTC
'NOTAFISCAL
'VALOR
'REMET_CGC
'REMET_NOME
'REMET_CIDADE
'REMET_UF
'DEST_NOME
'CIDADE_DEST
'UF_DEST
'PLACAVEIC
'MOTORISTA
'CPF
'MOTORISTA
'MODAL
        
        Excel.Range(Excel.Cells(5, 1), Excel.Cells(5, 15)).Interior.ColorIndex = 15
        
        
        x = 6
        .MoveFirst
        
        Do Until .EOF
        
            xData = .Fields("data")
            xfilialctc = .Fields("filialctc")
            xnfs = .Fields("nfs")
            xremet_cgc = .Fields("remet_cgc")
            xremet_nome = .Fields("remet_nome")
            xREMET_CIDADE = .Fields("REMET_CIDADE")
            xREMET_UF = .Fields("REMET_UF")
            xDEST_NOME = .Fields("DEST_NOME")
            xCIDADE_dest = .Fields("CIDADE_dest")
            xUF_dest = .Fields("UF_dest")
            xplacaveic = .Fields("placaveic")
            
            y = 0
            Excel.Cells(x, y + 1) = xData
            Excel.Cells(x, y + 2) = "'" & xfilialctc
            Excel.Cells(x, y + 3) = xnfs
            Excel.Cells(x, y + 4) = "'" & xremet_cgc
            Excel.Cells(x, y + 5) = xremet_nome
            Excel.Cells(x, y + 6) = xREMET_CIDADE
            Excel.Cells(x, y + 7) = xREMET_UF
            Excel.Cells(x, y + 8) = xDEST_NOME
            Excel.Cells(x, y + 9) = xCIDADE_dest
            Excel.Cells(x, y + 10) = xUF_dest
            Excel.Cells(x, y + 11) = xplacaveic

                        
            Excel.Range(Excel.Cells(x, 1), Excel.Cells(x, 15)).Borders.ColorIndex = 1
            
            x = x + 1
            
            If x Mod 2 = 0 Then
            
                Excel.Range(Excel.Cells(x, 1), Excel.Cells(x, 15)).Interior.ColorIndex = 19
                
            End If
            
            .MoveNext
            
        Loop
        
        Excel.Range(ExcelA1.Cells(5, 1), ExcelA1.Cells(x, 15)).EntireColumn.AutoFit
        
        xarquivo = "C:\MEDEX" & String(2 - Len(Day(Date)), "0") & String(2 - Len(Month(Date)), "0") & _
                    String(2 - Len(Hour(Time)), "0") & String(2 - Len(Minute(Time)), "0")
                    
        ExcelWBk.SaveAs xarquivo, , , , , , xlExclusive
        
        Excel.Quit
        Set ExcelA1 = Nothing
        Set ExcelWBk = Nothing
        
        'xenvia = EMAIL("romulo@intec.com.br; rzupelo@intec.com.br; fred@bomibrasil.com.br", xarquivo, "ENTREGAS MEDEX", _
                        "Segue em Anexo Arquivo de Entregas MEDEX." & Chr$(13) & Chr$(13) & "[]´S" & Chr$(13) & "ROMULO CONTI")
        
        End With
    
    Else
        
        MsgBox " Sem dados Para Geração de Arquivo"
    
    End If
    





End Sub

Private Sub MaskEdBox1_Change()

End Sub
