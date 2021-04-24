VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_fnac 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   1710
   ClientTop       =   1725
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   12030
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&SAIR"
         Height          =   375
         Left            =   10200
         TabIndex        =   13
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   11535
         Begin VB.CommandButton cmd_refresh 
            Caption         =   "Refresh Emails"
            Height          =   375
            Left            =   9960
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox List2 
            Height          =   2010
            Left            =   9960
            TabIndex        =   21
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmd_list1 
            Caption         =   "<<"
            Height          =   375
            Left            =   9480
            TabIndex        =   20
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton cmd_list2 
            Caption         =   ">>"
            Height          =   375
            Left            =   9480
            TabIndex        =   19
            Top             =   1080
            Width           =   375
         End
         Begin VB.ListBox List1 
            Height          =   2010
            Left            =   7920
            TabIndex        =   18
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chk_envia 
            Caption         =   "Envia Email"
            Height          =   255
            Left            =   8160
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame frm_gera_arquivo 
            Caption         =   "GERANDO ARQUIVO DA FNAC"
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
            Left            =   960
            TabIndex        =   14
            Top             =   720
            Visible         =   0   'False
            Width           =   5295
            Begin MSComctlLib.ProgressBar prg_bar1 
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   480
               Width           =   4815
               _ExtentX        =   8493
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   1
            End
            Begin VB.Label lab_progress 
               Alignment       =   2  'Center
               Caption         =   "ARQUIVO"
               Height          =   375
               Left            =   240
               TabIndex        =   16
               Top             =   840
               Width           =   4815
            End
         End
         Begin MSMask.MaskEdBox mask_data2 
            Height          =   300
            Left            =   6840
            TabIndex        =   12
            Top             =   2400
            Width           =   1000
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mask_data1 
            Height          =   300
            Left            =   5520
            TabIndex        =   10
            Top             =   2400
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmd_busca 
            Caption         =   "BUSCA REGISTROS E GERA *.TXT"
            Height          =   615
            Left            =   5520
            TabIndex        =   9
            Top             =   1560
            Width           =   2175
         End
         Begin VB.CommandButton cmd_import 
            Caption         =   "IMPORTA .XLS"
            Height          =   375
            Left            =   5520
            TabIndex        =   4
            Top             =   360
            Width           =   2175
         End
         Begin VB.FileListBox File1 
            Height          =   2625
            Left            =   3120
            Pattern         =   "*.xls"
            TabIndex        =   3
            Top             =   240
            Width           =   2175
         End
         Begin VB.DirListBox Dir1 
            Height          =   2565
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "à"
            Height          =   255
            Left            =   6600
            TabIndex        =   11
            Top             =   2400
            Width           =   135
         End
         Begin VB.Label lab_gravando 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   6600
            TabIndex        =   8
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "GRAVANDO:"
            Height          =   255
            Left            =   5520
            TabIndex        =   7
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "QTD DE REG:"
            Height          =   255
            Left            =   5520
            TabIndex        =   6
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lab_qtd 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   6600
            TabIndex        =   5
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.Shape Shape1 
         Height          =   3135
         Left            =   120
         Top             =   240
         Width           =   11775
      End
   End
End
Attribute VB_Name = "frm_fnac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_busca_Click()
Dim Excel As Excel.Application
Dim ExcelWBk As Excel.Workbook
Dim ExcelA1 As Excel.Worksheet
Dim xdata As String
Dim xfrete As Currency
Dim xmax As Integer
Dim xlinha As Integer
Dim X As Integer
Dim Y As Integer
Dim xnomearq As String
Dim xfilialctc As String
Dim xserie As String
Dim xstatus As String
Dim xcnpj As String
Dim xnf As String
Dim xdest As String

xdatamenos = CDate(mask_data1) - 5


If deb_fnac.rsSel_Tbfnac.State = 1 Then deb_fnac.rsSel_Tbfnac.Close
    deb_fnac.Sel_Tbfnac CDate(mask_data1), CDate(mask_data2), CDate(xdatamenos), CDate(mask_data2)
    
    If deb_fnac.rsSel_Tbfnac.RecordCount > 1 Then
        
        xmax = deb_fnac.rsSel_Tbfnac.RecordCount
        
        Set Excel = CreateObject("Excel.Application")
        Set Excel = GetObject(, "Excel.Application")
        Excel.Visible = False
        Set ExcelWBk = Excel.Workbooks.Add
        Set ExcelA1 = Excel.Worksheets(1)
    
        deb_fnac.rsSel_Tbfnac.MoveFirst
    
    
        ExcelA1.Cells.Font.Name = "Verdana"
        
        ExcelA1.Cells(1, 1) = "FILIALCTC"
        ExcelA1.Cells(1, 2) = "SERIE"
        ExcelA1.Cells(1, 3) = "NF"
        ExcelA1.Cells(1, 4) = "DATA"
        ExcelA1.Cells(1, 5) = "FRETE"
        ExcelA1.Cells(1, 6) = "STATUS"
        ExcelA1.Cells(1, 7) = "NR_CNPJ"
                
        X = 0
        
        frm_gera_arquivo.Visible = True
        Me.Refresh
        prg_bar1.Min = 0
        prg_bar1.Max = xmax
        prg_bar1.Value = X
        
        X = 1
        
        deb_fnac.rsSel_Tbfnac.MoveFirst
        
        'MONTA O ARQUIVO
        Do Until deb_fnac.rsSel_Tbfnac.EOF
            
            xfilialctc = deb_fnac.rsSel_Tbfnac.Fields("filialctc")
            xserie = deb_fnac.rsSel_Tbfnac.Fields("serie")
            xdata = Year(deb_fnac.rsSel_Tbfnac.Fields("data")) & "/" & Month(deb_fnac.rsSel_Tbfnac.Fields("data")) _
                    & "/" & Day(deb_fnac.rsSel_Tbfnac.Fields("data"))
            xfrete = deb_fnac.rsSel_Tbfnac.Fields("fretetotal")
            xstatus = deb_fnac.rsSel_Tbfnac.Fields("status")
            xcnpj = deb_fnac.rsSel_Tbfnac.Fields("nr_cnpj")
            xnf = deb_fnac.rsSel_Tbfnac.Fields("numnf")
            
            ExcelA1.Cells(X + 1, 1) = xfilialctc
            ExcelA1.Cells(X + 1, 2) = xserie
            ExcelA1.Cells(X + 1, 3) = xnf
            ExcelA1.Cells(X + 1, 4) = xdata
            ExcelA1.Cells(X + 1, 5) = xfrete
            ExcelA1.Cells(X + 1, 6) = xstatus
            ExcelA1.Cells(X + 1, 7) = xcnpj
            
            deb_fnac.rsSel_Tbfnac.MoveNext
                                   
            X = X + 1
            
                        
            'COLOCA BORDAS
            
            lab_progress.Caption = (X - 1) & " / " & xmax
            prg_bar1.Value = X - 1
            Me.Refresh
        
        Loop
        
        xmax = X

                
                'COLOCA A PRIMEIRA LINHA EM NEGRITO
        For X = 1 To 7
            
            ExcelA1.Range(ExcelA1.Cells(1, 1), ExcelA1.Cells(1, X)).Font.Bold = True
            ExcelA1.Range(ExcelA1.Cells(1, 1), ExcelA1.Cells(1, X)).Borders.ColorIndex = 1
            lab_progress.Caption = "FORMATANDO PLANILHA"
            Me.Refresh
            
        Next
         'FORMATA PLANILHA
        'For x = 1 To xmax
        
         '    If x Mod 2 = 0 Then
          '        For y = 1 To 7
                  
           '         ExcelA1.Range(ExcelA1.Cells(x, 1), ExcelA1.Cells(x, y)).Font.ColorIndex = 5
                    
                    
            '    Next
                
           ' End If
           ' lab_progress.Caption = "FORMATANDO PLANILHA - " & x
           ' Me.Refresh
        'Next
               
               
                
                
        frm_gera_arquivo.Visible = False
        
        ExcelA1.Range("A:DZ").EntireColumn.AutoFit
        ExcelA1.Name = "FNAC"
        xmes = UCase(MonthName(Month(Date)))
        xnomearq = "C:\FNAC DE " & Mid(mask_data1, 1, 2) & " A " & Mid(mask_data2, 1, 2) & " DE " & xmes
        
        
        
        ExcelWBk.SaveAs xnomearq, , , , , , xlExclusive
    
        Excel.Quit
        Set ExcelWS1 = Nothing
        Set ExcelA1 = Nothing
        Set Excel = Nothing
        
        
        'ENVIA ARQUIVO POR EMAIL
        If chk_envia.Value = 1 Then
              
            Dim xMail As Outlook.Application
            Dim xanexo As Outlook.Items
            Dim xMensagem As Outlook.MailItem
            Set xMail = CreateObject("Outlook.Application")
            Set xMensagem = xMail.CreateItem(olMailItem)
           xdest = ""
            For X = 0 To List2.ListCount - 1
                
                xdest = xdest & List2.List(X) & "; "
            
            Next
                        
            xMensagem.To = xdest
            xMensagem.Subject = "ARQUIVO FNAC DE " & mask_data1 & " À " & mask_data2
            
            xMensagem.Body = "SEGUE ARQUIVO EM ANEXO." & Chr$(13) & Chr$(13) & "[]´s" & Chr$(13) _
                             & Chr$(13) & "ROMULO CONTI"
            
            xMensagem.Importance = olImportanceNormal
            xMensagem.Attachments.Add xnomearq & ".xls"
            
            DoEvents
            
            'ENVIA MENSAGEM
            xMensagem.Send
            
            'FECHA EMAIL
            Set xMensagem = Nothing
            Set xMail = Nothing
                   
            MsgBox "Mensagem Enviada com Sucesso", vbInformation, "EMAIL"
            
                   
        Else
            
            MsgBox "ARQUIVO GERADO COM SUCESSO", vbOKOnly, "FNAC"
        
        End If
        
            
     
    End If
    
        


End Sub

Private Sub cmd_import_Click()
Dim Excel As Excel.Application
Dim ExcelA1 As Excel.Worksheet
Dim LinhaAtual As Integer, LinhaMAX As Integer
Dim xarq As String
Dim nr_cnpj As String
Dim nr_nf As String
Dim ds_serie As String
Dim data_emissao As String
Dim cd_cfop As String
Dim ds_cfop As String
Dim Status As String
Dim numnfnum As Integer


If File1.ListIndex = -1 Then

    MsgBox "selecione o Arquivo"
    Exit Sub

End If

    
xarq = File1.Path & "\" & File1.FileName
    
    LinhaMAX = 0
    
  Set Excel = CreateObject("EXCEL.APPLICATION")
    Excel.Visible = False
    Excel.Interactive = False
    Excel.Workbooks.Open FileName:=xarq
    Set ExcelA1 = Excel.Worksheets(1)
    

If Trim$(ExcelA1.Cells(1, 1)) <> "nr_cnpj" Then
    Excel.Quit
    Set ExcelA1 = Nothing
    Set Excel = Nothing
    
    MsgBox "Arquivo Não pode ser Importado", vbCritical, "FNAC"
        
    Exit Sub
    
End If
    
    
    
    LinhaAtual = 0
    
    Do While True
        LinhaAtual = LinhaAtual + 1
        LinhaMAX = LinhaMAX + 1
        
        If Len(Trim(ExcelA1.Cells(LinhaAtual, 1))) = 0 And LinhaAtual >= 2 Then
            Exit Do
        End If
        
        lab_qtd.Caption = LinhaAtual - 1
        Me.Refresh
    
    Loop
    
   

For LinhaAtual = 2 To LinhaMAX - 1
    
    nr_cnpj = Trim$(ExcelA1.Cells(LinhaAtual, 1))
    nr_nf = Trim(ExcelA1.Cells(LinhaAtual, 2))
    ds_serie = Trim$(ExcelA1.Cells(LinhaAtual, 3))
    data_emissao = Trim$(ExcelA1.Cells(LinhaAtual, 4))
    cd_cfop = Trim$(ExcelA1.Cells(LinhaAtual, 5))
    ds_cfop = Trim$(ExcelA1.Cells(LinhaAtual, 6))
    Status = UCase(Trim$(ExcelA1.Cells(LinhaAtual, 7)))

    deb_fnac.in_Tbfnac nr_cnpj, nr_nf, ds_serie, CDate(data_emissao), cd_cfop, ds_cfop, Status
        
    lab_gravando.Caption = LinhaAtual
    Me.Refresh
    
Next


Excel.Quit
Set ExcelA1 = Nothing
Set Excel = Nothing

FileCopy Dir1.Path & "\" & File1.FileName, "c:\INFORMA\EDI_IMP\FOX\BACKUP\" & File1.FileName
Kill Dir1.Path & "\" & File1.FileName




deb_fnac.up_tb_fnac

MsgBox "ARQUIVO IMPORTADO", vbInformation, "FNAC"
    


End Sub

Private Sub cmd_list2_Click()

List2.AddItem List1.List(List1.ListIndex)
List1.RemoveItem List1.ListIndex


End Sub

Private Sub cmd_refresh_Click()
Dim xcont As Integer
Dim cont As Integer
deb_hora.rssel_emails.Open
deb_hora.rssel_emails.MoveFirst

xcont = deb_hora.rssel_emails.RecordCount

For X = 1 To xcont
    List1.AddItem LCase(Trim$(deb_hora.rssel_emails.Fields("email")))
    deb_hora.rssel_emails.MoveNext
    Me.Refresh
    
Next


deb_hora.rssel_emails.Close
End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
Dir1.Path = "c:\informa"
File1.Path = Dir1.Path


End Sub
