VERSION 5.00
Begin VB.Form frmExportSitla 
   Caption         =   "Exporta Arquivo de POD para o SITLA"
   ClientHeight    =   2325
   ClientLeft      =   2145
   ClientTop       =   1515
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   7215
   Begin VB.CommandButton cmdGeraArq 
      Caption         =   "Gerar Arquivo"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5415
      Begin VB.Label lblCont 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3600
         TabIndex        =   6
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CTCs gravados para atualização: "
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   2400
      End
      Begin VB.Label Label2 
         Caption         =   "Nome do Arquivo: DTENTARQ.TXT"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Gera arquivo para o SITLA para atualização de POD's (Data de Entrega)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmExportSitla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGeraArq_Click()
    Dim xdata As String, xctc As String, xentrega As String, xhora As String, xreceb As String, xcont As Long
    
    If de_informa.rsSel_ExportSitla.State = 1 Then de_informa.rsSel_ExportSitla.Close
    de_informa.Sel_ExportSitla
    
    If de_informa.rsSel_ExportSitla.RecordCount > 0 Then
        de_informa.rsSel_ExportSitla.MoveFirst
        
        'trata xdata como variavel contendo a data no formato DDMMAAAA
        
        xdata = ""
        If Len(Trim$(Str(Day(datahora("data"))))) = 1 Then
            xdata = xdata & "0" & Trim$(Str(Day(datahora("data"))))
        Else
            xdata = xdata & Trim$(Str(Day(datahora("data"))))
        End If
        If Len(Trim$(Str(Month(datahora("data"))))) = 1 Then
            xdata = xdata & "0" & Trim$(Str(Month(datahora("data"))))
        Else
            xdata = xdata & Trim$(Str(Month(datahora("data"))))
        End If
        xdata = xdata & Trim$(Str(Year(datahora("data"))))
        xlinha = ""
        
        'ABRE O ARQUIVO E MONTA O CABEÇARIO 'REG 01
        
        Open "C:\INFORMA\SITLA\DTENTARQ.TXT" For Output As #1
        xlinha = "1;52134798000320;INTEC TRANSPORTES;" & xdata
        Print #1, xlinha
        
        Do Until de_informa.rsSel_ExportSitla.EOF

            
            'trata número do ctc
            
            xctc = Mid(de_informa.rsSel_ExportSitla.Fields("filialctc"), 1, 2) & "-" & _
                   Mid(de_informa.rsSel_ExportSitla.Fields("filialctc"), 5, 6)
            
            'trata data de entrega formado DDMMAAAA
            
            xentrega = ""
            If Len(Trim$(Str(Day(de_informa.rsSel_ExportSitla.Fields("data"))))) = 1 Then
                xentrega = xentrega & "0" & Trim$(Str(Day(de_informa.rsSel_ExportSitla.Fields("data"))))
            Else
                xentrega = xentrega & Trim$(Str(Day(de_informa.rsSel_ExportSitla.Fields("data"))))
            End If
            If Len(Trim$(Str(Month(de_informa.rsSel_ExportSitla.Fields("data"))))) = 1 Then
                xentrega = xentrega & "0" & Trim$(Str(Month(de_informa.rsSel_ExportSitla.Fields("data"))))
            Else
                xentrega = xentrega & Trim$(Str(Month(de_informa.rsSel_ExportSitla.Fields("data"))))
            End If
            xentrega = xentrega & Trim$(Str(Year(de_informa.rsSel_ExportSitla.Fields("data"))))
            
            'trata hora no formato HHMMSS
            
            xhora = Mid$(de_informa.rsSel_ExportSitla.Fields("hora"), 1, 2) & _
                    Mid$(de_informa.rsSel_ExportSitla.Fields("hora"), 4, 2) & "00"
                    
            'trata Nome de Recebedor
            
            If de_informa.rsSel_ExportSitla.Fields("baixadofinal") = "S" Then
                If Not IsNull(de_informa.rsSel_ExportSitla.Fields("receb")) Then
                    xreceb = Trim$(de_informa.rsSel_ExportSitla.Fields("receb"))
                Else
                    xreceb = ""
                End If
            Else
                If Not IsNull(de_informa.rsSel_ExportSitla.Fields("recebpre")) Then
                    xreceb = Trim$(de_informa.rsSel_ExportSitla.Fields("recebpre"))
                Else
                    xreceb = ""
                End If
            End If
                        
            'grava dados no arquivo TXT
            xlinha = "3;" & xctc & ";" & xentrega & ";" & xhora & ";" & xreceb & ";"
            Print #1, xlinha
            'atualiza ATUAL_SITLA = "N" ou seja, não atualiza sitla, pois já foi atualizado
            de_informa.Alt_ExportSitlaNao de_informa.rsSel_ExportSitla.Fields("filialctc")
            de_informa.rsSel_ExportSitla.MoveNext
            xcont = xcont + 1
            lblCont = xcont
            DoEvents
        Loop
        Close #1
        MsgBox "Processo Finalizado !"
    Else
        MsgBox "Não há Dados a serem atualizados !"
    End If
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmExportSitla = Nothing
End Sub
