VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_imp 
   Caption         =   "CONFERÊNCIA DE PRÉ - FATURAS"
   ClientHeight    =   4530
   ClientLeft      =   2835
   ClientTop       =   4785
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9585
   Begin VB.Frame Frame1 
      Height          =   4515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9540
      Begin VB.Frame Frame2 
         Height          =   4215
         Left            =   75
         TabIndex        =   1
         Top             =   225
         Width           =   9390
         Begin VB.Frame Frame5 
            Height          =   540
            Left            =   7200
            TabIndex        =   17
            Top             =   1950
            Width           =   2040
            Begin VB.TextBox txt_aliq 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1125
               TabIndex        =   19
               Text            =   "0"
               Top             =   225
               Width           =   840
            End
            Begin VB.Label Label4 
               Caption         =   "Aliquota:"
               Height          =   240
               Left            =   75
               TabIndex        =   18
               Top             =   225
               Width           =   615
            End
         End
         Begin MSComctlLib.ProgressBar prg_bar 
            Height          =   240
            Left            =   75
            TabIndex        =   16
            Top             =   3825
            Width           =   8190
            _ExtentX        =   14446
            _ExtentY        =   423
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   7200
            TabIndex        =   14
            Top             =   1275
            Width           =   2040
            Begin VB.CheckBox chk_Subs 
               Caption         =   "Subs. Tributária"
               Height          =   240
               Left            =   150
               TabIndex        =   15
               Top             =   225
               Width           =   1740
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1140
            Left            =   7200
            TabIndex        =   11
            Top             =   150
            Width           =   2040
            Begin VB.OptionButton opt_devolucoes 
               Caption         =   "Devoluções"
               Height          =   240
               Left            =   150
               TabIndex        =   13
               Top             =   675
               Width           =   1365
            End
            Begin VB.OptionButton opt_Entregas 
               Caption         =   "Entregas"
               Height          =   315
               Left            =   150
               TabIndex        =   12
               Top             =   225
               Value           =   -1  'True
               Width           =   1740
            End
         End
         Begin VB.ComboBox cmb_tabelas 
            Height          =   315
            Left            =   5175
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1725
            Width           =   1665
         End
         Begin VB.CommandButton cmd_sair 
            Caption         =   "&Sair"
            Height          =   315
            Left            =   5100
            TabIndex        =   7
            Top             =   375
            Width           =   1740
         End
         Begin VB.CommandButton cmd_Processar 
            Caption         =   "Processar"
            Height          =   315
            Left            =   5100
            TabIndex        =   4
            Top             =   750
            Width           =   1740
         End
         Begin VB.FileListBox File1 
            Height          =   2625
            Left            =   2775
            Pattern         =   "*.xls"
            TabIndex        =   3
            Top             =   225
            Width           =   2190
         End
         Begin VB.DirListBox Dir1 
            Height          =   2565
            Left            =   75
            TabIndex        =   2
            Top             =   225
            Width           =   2565
         End
         Begin VB.Label Label3 
            Caption         =   "TABELA:"
            Height          =   240
            Left            =   5175
            TabIndex        =   9
            Top             =   1500
            Width           =   840
         End
         Begin VB.Label Label2 
            Caption         =   "By Conti"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8400
            TabIndex        =   8
            Top             =   3900
            Width           =   840
         End
         Begin VB.Label lab_reg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   240
            Left            =   5925
            TabIndex        =   6
            Top             =   1125
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Total Reg.:"
            Height          =   315
            Left            =   5100
            TabIndex        =   5
            Top             =   1125
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "frm_imp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xtipo As Integer

Private Sub cmb_tabelas_Change()

cmd_Processar.Enabled = True



End Sub

Private Sub cmd_Processar_Click()
Dim Excel As Excel.Application
Dim ExcelPlan As Excel.Worksheet
Dim xmax As Integer
Dim xatual As Integer
Dim x As Integer

If opt_Entregas.Value = True Then

    xtipo = 1
    
Else

    xtipo = 0

End If


Set Excel = CreateObject("EXCEL.Application")

Excel.Visible = True
Excel.Interactive = True
Excel.Workbooks.Open FileName:=File1.Path & "\" & File1.FileName

Set ExcelPlan = Excel.Worksheets(1)

xatual = 2

Do While True
    
    xatual = xatual + 1
    xmax = xmax + 1
    
    If Len(Trim$(Excel.Cells(xatual, 1))) = 0 And Len(Trim$(Excel.Cells(xatual, 2))) = 0 And xatual >= 2 Then
    
        Exit Do
        
    End If
    
    lab_reg.Caption = xmax
    Me.Refresh
    




Loop

Excel.Cells(1, 29) = "TODOS VALORES"
Excel.Range(Excel.Cells(1, 29), Excel.Cells(1, 29)).HorizontalAlignment = xlCenter
Excel.Range(Excel.Cells(1, 29), Excel.Cells(1, 29)).Font.Bold = True




Excel.Cells(2, 26) = "CTRC"
Excel.Cells(2, 27) = "FT PESO"
Excel.Cells(2, 28) = "PESO VLR"
Excel.Cells(2, 29) = "COLETA"
Excel.Cells(2, 30) = "TOTAL"
Excel.Cells(2, 31) = "ALQT"
Excel.Cells(2, 32) = "ICMS"

Excel.Range(Excel.Cells(2, 26), Excel.Cells(2, 32)).Interior.ColorIndex = 33
Excel.Range(Excel.Cells(2, 26), Excel.Cells(2, 32)).Borders.ColorIndex = 1

Excel.Cells(1, 35) = "VALORES S/IMPOSTOS"
Excel.Range(Excel.Cells(1, 35), Excel.Cells(1, 29)).HorizontalAlignment = xlCenter
Excel.Range(Excel.Cells(1, 35), Excel.Cells(1, 29)).Font.Bold = True

Excel.Cells(2, 34) = "FT PESO"
Excel.Cells(2, 35) = "FT VALOR"
Excel.Cells(2, 36) = "TOTAL"
Excel.Cells(2, 37) = "ENTREGA"

Excel.Range(Excel.Cells(2, 34), Excel.Cells(2, 36)).Interior.ColorIndex = 42
Excel.Range(Excel.Cells(2, 34), Excel.Cells(2, 36)).Borders.ColorIndex = 1


Excel.Cells(1, 40) = "VALORES COM IMPOSTOS"
Excel.Cells(1, 43) = "ACERTO"
Excel.Range(Excel.Cells(1, 40), Excel.Cells(1, 43)).HorizontalAlignment = xlCenter
Excel.Range(Excel.Cells(1, 40), Excel.Cells(1, 43)).Font.Bold = True

Excel.Cells(2, 38) = "FT PESO"
Excel.Cells(2, 39) = "FT VLR"
Excel.Cells(2, 40) = "TOTAL"
Excel.Cells(2, 41) = "ALQT"
Excel.Cells(2, 42) = "ICMS"

Excel.Range(Excel.Cells(2, 37), Excel.Cells(2, 42)).Interior.ColorIndex = 36
Excel.Range(Excel.Cells(2, 37), Excel.Cells(2, 42)).Borders.ColorIndex = 1

'calculos VALORES INCORRETOS
Dim XResult As Currency
Dim Xcol12 As Currency
Dim Xcol13 As Currency
Dim Xcol14 As Currency
Dim xaliq As Integer
Dim xicms As Currency
Dim xCIM As String
Dim xcidade As String
Dim xpeso As Integer
Dim xuf As String
Dim xvlrFreteNf As Currency
Dim xadval As Double
Dim xvalNf As Currency
Dim xftvalor As Currency
Dim xcoletavalor As Currency
Dim xentregavalor As Currency


prg_bar.Min = 0
prg_bar.Max = xmax
prg_bar.Value = 0

For xatual = 3 To xmax + 1

If Excel.Cells(xatual, 2) = "5351" Then
    MsgBox "ESTE"
End If




Me.Refresh

prg_bar.Value = xatual - 1


If xtipo = 0 Then

    xcidade = xsepara(Excel.Cells(xatual, 4))
    xuf = Mid(Excel.Cells(xatual, 4), Len(Trim$(Excel.Cells(xatual, 4))) - 1, 2)
Else

    xcidade = Excel.Cells(xatual, 6) & "-" & Excel.Cells(xatual, 7)
    xuf = Mid(xcidade, Len(Trim$(xcidade)) - 1, 2)
    xcidade = xsepara(xcidade)
End If

xpeso = Excel.Cells(xatual, 11)

'pesquisa primeiro a Cidade se exite diferença

If Deb_Bio.rsSel_CIDADE.State = 1 Then Deb_Bio.rsSel_CIDADE.Close
    Deb_Bio.Sel_CIDADE cmb_tabelas.Text, xuf, xcidade
    
    If Deb_Bio.rsSel_CIDADE.RecordCount = 0 Then

        If Deb_Bio.rsSel_CI.State = 1 Then Deb_Bio.rsSel_CI.Close
            Deb_Bio.Sel_CI xcidade, xuf
         

            'ARRUMAR SE FOR RECORDCOUNT 0
            xCIM = Deb_Bio.rsSel_CI.Fields("cim")
            xuf = Deb_Bio.rsSel_CI.Fields("uf")

            If xCIM = "C" Then
            xCIM = "CAP"

            ElseIf xCIM = "I" Then
    
            xCIM = "INT"
    
            End If
        
    Else
        
         xCIM = "CID"
            
    End If
        


Dim xpesode As Double
Dim xpesoate As Double

If Deb_Bio.rsSel_tr02Pesos.State = 1 Then Deb_Bio.rsSel_tr02Pesos.Close
    Deb_Bio.Sel_tr02Pesos cmb_tabelas.Text, xuf, xCIM, xcidade
    
            If Deb_Bio.rsSel_tr02Pesos.RecordCount = 0 Then
            
                If Deb_Bio.rsSel_tr02Pesos.State = 1 Then Deb_Bio.rsSel_tr02Pesos.Close
                   Deb_Bio.Sel_tr02Pesos cmb_tabelas.Text, xuf, xCIM, "%"
                
                    If Deb_Bio.rsSel_tr02Pesos.RecordCount = 0 Then
                    
                        If Deb_Bio.rsSel_tr02Pesos.State = 1 Then Deb_Bio.rsSel_tr02Pesos.Close
                            Deb_Bio.Sel_tr02Pesos cmb_tabelas.Text, xuf, xCIM, " "
                    End If
                                       
                                   
            End If
            
            
            
            Deb_Bio.rsSel_tr02Pesos.MoveFirst
    
            Do Until Deb_Bio.rsSel_tr02Pesos.EOF
        
                xpesode = Deb_Bio.rsSel_tr02Pesos.Fields("pesode")
                xpesoate = Deb_Bio.rsSel_tr02Pesos.Fields("pesoate")
        
        
                
            
                If xpesode <= xpeso And xpeso <= xpesoate Then
                    
                        If Deb_Bio.rsSel_tr02Pesos.Fields("porkilo") > 0 Then
            
                            xvlrFreteNf = xpeso * Deb_Bio.rsSel_tr02Pesos.Fields("porkilo")
                            xadval = Deb_Bio.rsSel_tr02Pesos.Fields("ADVAL") / 100
                            xcoletavalor = Deb_Bio.rsSel_tr02Pesos.Fields("coleta_valor")
                            xentregavalor = Deb_Bio.rsSel_tr02Pesos.Fields("entrega_valor")
                            
                            Excel.Cells(xatual, 35) = Excel.Cells(xatual, 35) + xcoletavalor + xentregavalor
                            
                            'Excel.Range(Excel.Cells(xatual, 30), Excel.Cells(xatual, 30)).Comment = XResult & " + Valor Coleta: " & xcoletavalor & " + Valor Entrega: " & xentregavalor
                            
            
                            Exit Do
                        
                        Else
                      
                            xvlrFreteNf = Deb_Bio.rsSel_tr02Pesos.Fields("fretepeso")
            
                            xadval = Deb_Bio.rsSel_tr02Pesos.Fields("ADVAL") / 100
                            
                            xcoletavalor = Deb_Bio.rsSel_tr02Pesos.Fields("coleta_valor")
                            xentregavalor = Deb_Bio.rsSel_tr02Pesos.Fields("entrega_valor")
                            
                            
                        Exit Do
                        
                       End If
                       
                
                End If
            
                Deb_Bio.rsSel_tr02Pesos.MoveNext
    
    
                Loop
                
    

xvalNf = Excel.Cells(xatual, 9)
    
    
Xcol12 = Excel.Cells(xatual, 15)
Xcol13 = Excel.Cells(xatual, 16)
Xcol14 = Excel.Cells(xatual, 19)

XResult = Xcol13 + Xcol12 + Xcol14

If txt_aliq.Text = 0 Then
    xicms = Excel.Cells(xatual, 25)
    'Format(XResult * (xaliq / 100), "##.##")
    xaliq = ((xicms * 100) / Excel.Cells(xatual, 14))
Else
    xaliq = Int(Val(txt_aliq.Text))
    xicms = (Excel.Cells(xatual, 14) * (txt_aliq.Text / 100))
End If


'VALORES TODOS
Excel.Cells(xatual, 26) = Excel.Cells(xatual, 2)

Excel.Cells(xatual, 27) = Xcol12
Excel.Cells(xatual, 28) = Xcol13
Excel.Cells(xatual, 29) = Excel.Cells(xatual, 19)
Excel.Cells(xatual, 30) = XResult
Excel.Cells(xatual, 31) = xaliq
Excel.Cells(xatual, 32) = xicms

Excel.Range(Excel.Cells(xatual, 26), Excel.Cells(xatual, 32)).Interior.ColorIndex = 34
Excel.Range(Excel.Cells(xatual, 26), Excel.Cells(xatual, 32)).Borders.ColorIndex = 1



'VALORES S/ IMPOSTOS
Excel.Cells(xatual, 34) = xvlrFreteNf

xftvalor = xvalNf * xadval

Excel.Cells(xatual, 35) = xftvalor
Excel.Cells(xatual, 36) = (Excel.Cells(xatual, 34) + Excel.Cells(xatual, 35))
Excel.Cells(xatual, 36) = Excel.Cells(xatual, 36) + xcoletavalor + xentregavalor
Excel.Range(Excel.Cells(xatual, 34), Excel.Cells(xatual, 36)).Interior.ColorIndex = 35
Excel.Range(Excel.Cells(xatual, 34), Excel.Cells(xatual, 36)).Borders.ColorIndex = 1


'VALORES COM IMPOSTOS
Dim xftpeso As Currency
Dim xtot As Currency
Dim xacerto As Currency


If (xsepara(Excel.Cells(xatual, 4)) = "SAO PAULO" And Excel.Cells(xatual, 6) = "SAO PAULO") Or (xsepara(Excel.Cells(xatual, 6)) = "SAO PAULO" And Excel.Cells(xatual, 4) = "SAO PAULO") Then
    Excel.Cells(xatual, 32) = 0
    Excel.Cells(xatual, 38) = xvlrFreteNf
    Excel.Cells(xatual, 39) = xftvalor
    xicms = 0
Else

    If chk_Subs.Value = 0 Then
        xvlrFreteNf = (xvlrFreteNf / 0.88)
        xftvalor = (xftvalor / 0.88)
    End If
    
    Excel.Cells(xatual, 38) = xvlrFreteNf
    Excel.Cells(xatual, 39) = xftvalor
    xtot = (Excel.Cells(xatual, 38) + Excel.Cells(xatual, 39)) + xcoletavalor + xentregavalor
    xicms = Format(xtot * (xaliq / 100), "##.####")
    Excel.Cells(xatual, 38) = xvlrFreteNf

    Excel.Cells(xatual, 39) = xftvalor
    
End If

Excel.Cells(xatual, 37) = xentregavalor

If xentregavalor > 0 Then

    Excel.Range(Excel.Cells(xatual, 37), Excel.Cells(xatual, 37)).Font.ColorIndex = 3
    Excel.Range(Excel.Cells(xatual, 37), Excel.Cells(xatual, 37)).Font.Bold = True
End If

    


xtot = (Excel.Cells(xatual, 38) + Excel.Cells(xatual, 39)) + xcoletavalor + xentregavalor
Excel.Cells(xatual, 40) = xtot
Excel.Cells(xatual, 41) = xaliq
Excel.Cells(xatual, 42) = xicms

If chk_Subs.Value = 0 Then

    xacerto = (Excel.Cells(xatual, 30) - Excel.Cells(xatual, 40))

Else

    xacerto = (Excel.Cells(xatual, 30) - Excel.Cells(xatual, 32)) - Excel.Cells(xatual, 40)

End If


Excel.Cells(xatual, 43) = xacerto

If xacerto < 0 Then

    Excel.Range(Excel.Cells(xatual, 43), Excel.Cells(xatual, 43)).Font.ColorIndex = 10
    
End If

Excel.Range(Excel.Cells(xatual, 38), Excel.Cells(xatual, 43)).Interior.ColorIndex = 19
Excel.Range(Excel.Cells(xatual, 38), Excel.Cells(xatual, 43)).Borders.ColorIndex = 1
    
    


Excel.Range(Excel.Cells(xatual, 26), Excel.Cells(xatual, 43)).EntireColumn.AutoFit




Next


Excel.Interactive = True





MsgBox "OK"




'Excel.Quit

'Set ExcelPlan = Nothing
'Set Excel = Nothing



End Sub


Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()


If Xbusca("C:\Prefatura", True) = False Then
    MkDir ("c:\Prefatura")
End If



Dir1.Path = "c:\PREFATURA"

If Deb_Bio.rsSel_Tr02.State = 1 Then Deb_Bio.rsSel_Tr02.Close
    Deb_Bio.Sel_Tr02
    
    Deb_Bio.rsSel_Tr02.MoveFirst
    
    Do Until Deb_Bio.rsSel_Tr02.EOF
    
        cmb_tabelas.AddItem Deb_Bio.rsSel_Tr02.Fields("CODIGO")
        
        Deb_Bio.rsSel_Tr02.MoveNext
    
    Loop
    


End Sub

Private Sub opt_devolucoes_Click()
If opt_devolucoes.Value = True Then

    xtipo = 0
    
End If

End Sub

Private Sub opt_Entregas_Click()

If opt_Entregas.Value = True Then

    xtipo = 1
    
End If



End Sub
