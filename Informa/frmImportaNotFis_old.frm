VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmImportaNotFis 
   Caption         =   "Importa NOTFIS - Sistema Informa"
   ClientHeight    =   6270
   ClientLeft      =   3195
   ClientTop       =   1620
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7935
   Begin VB.Frame Frame4 
      Caption         =   "LOG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   7695
      Begin MSDataGridLib.DataGrid gridLog 
         Bindings        =   "frmImportaNotFis.frx":0000
         Height          =   3615
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   6376
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
         DataMember      =   "Sel_LogNotFis"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "data"
            Caption         =   "data"
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
            DataField       =   "arquivo"
            Caption         =   "arquivo"
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
            DataField       =   "tipo"
            Caption         =   "tipo"
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
            DataField       =   "descricao"
            Caption         =   "descricao"
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
               ColumnWidth     =   1695,118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1769,953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1305,071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   6900,095
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   4200
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
      Begin MSComctlLib.ProgressBar Progress1 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Arquivo em Processamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   3615
      Begin VB.Label lblArquivo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   7440
      Top             =   0
   End
   Begin VB.Label lblCont 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   6960
      TabIndex        =   8
      Top             =   0
      Width           =   90
   End
End
Attribute VB_Name = "frmImportaNotFis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    File1.Path = "c:\edi"
End Sub
Private Sub Timer1_Timer()
    Dim xid_notfis As String, xremet_cgc As String, xremet_nome As String, xdest_cgc As String, xdest_nome As String
    Dim xdest_ie As String, xdest_end As String, xdest_bairro As String, xdest_cidade As String, xdest_uf As String
    Dim xdest_cep As String, xtipocarga As String, xtipofrete As String, xNumNF As String, xNumNfNum As Single
    Dim xserie As String, xemissaonf As Date, xnatureza As String, xespecie As String, xvolumes As Single
    Dim xvalmerc As Currency, xpeso As Currency, xpesocub As Currency, xdataimp As Date, xgravar As String
    Dim xdatainterface As Date, xIdRomaneio As String, xQtdeItem As Integer
    Dim xCodigoItem As String, xDescricaoItem As String, xPosicaoItem As Integer
    
    On Error GoTo TrataErro
    
    Timer1.Interval = 0
    Timer2.Interval = 0
    
    File1.Path = "c:\edi"
    File1.Refresh
    
    If File1.ListCount > 0 Then
        lblArquivo.Caption = File1.List(0)
    Else
        lblArquivo.Caption = ""
        Timer1.Interval = 15000
        Exit Sub
    End If
    
    'antes de contar a quantidade de registros verificar se o arquivo é NOTFIS
    xLin = 0
    xIdRomaneio = "N"
    
    Open File1.Path & "\" & lblArquivo.Caption For Input As #1
    
    de_informa.Ins_LogNotFis "INICIO", "IMPORTACAO DO ARQUIVO", lblArquivo.Caption
    
    If de_informa.rsSel_LogNotFis.State = 1 Then de_informa.rsSel_LogNotFis.Close
    de_informa.Sel_LogNotFis
    gridLog.DataMember = "sel_lognotfis"
    gridLog.Refresh
    DoEvents

    
    Do Until EOF(1)
        xLin = xLin + 1
        Line Input #1, xlinha
        If xLin = 1 Then
            If Mid(xlinha, 1, 3) <> "000" Then
                de_informa.Ins_LogNotFis "CRITICA", "ARQUIVO INVALIDO", lblArquivo.Caption
                
                If de_informa.rsSel_LogNotFis.State = 1 Then de_informa.rsSel_LogNotFis.Close
                de_informa.Sel_LogNotFis
                gridLog.DataMember = "sel_lognotfis"
                gridLog.Refresh
                DoEvents
                
                Close #1
                On Error Resume Next
                FileCopy File1.Path & "\" & lblArquivo.Caption, File1.Path & "\critica\" & lblArquivo.Caption
                Kill File1.Path & "\" & lblArquivo.Caption
                File1.Path = "C:\EDI"
                File1.Refresh
                lblArquivo.Caption = ""
                Timer1.Interval = 15000
                Timer2.Interval = 60000
                Exit Sub
            Else
                If Mid(xlinha, 84, 3) <> "NOT" Then
                    de_informa.Ins_LogNotFis "CRITICA", "ARQUIVO INVALIDO", lblArquivo.Caption
                    
                    If de_informa.rsSel_LogNotFis.State = 1 Then de_informa.rsSel_LogNotFis.Close
                    de_informa.Sel_LogNotFis
                    gridLog.DataMember = "sel_lognotfis"
                    gridLog.Refresh
                    DoEvents
                    
                    Close #1
                    On Error Resume Next
                    FileCopy File1.Path & "\" & lblArquivo.Caption, File1.Path & "\critica\" & lblArquivo.Caption
                    Kill File1.Path & "\" & lblArquivo.Caption
                    File1.Path = "C:\EDI"
                    File1.Refresh
                    lblArquivo.Caption = ""
                    Timer1.Interval = 15000
                    Timer2.Interval = 60000
                    Exit Sub
                End If
            End If
        End If
        
        If Mid(xlinha, 1, 3) = "313" Then
            xtamarq = xtamarq + 1
            If xIdRomaneio = "N" Then xIdRomaneio = Trim$(Mid(xlinha, 4, 15))
        End If
        DoEvents
    Loop
    Close #1
    
    If xtamarq = 0 Then
        de_informa.Ins_LogNotFis "CRITICA", "ARQUIVO INVALIDO", lblArquivo.Caption
        
        If de_informa.rsSel_LogNotFis.State = 1 Then de_informa.rsSel_LogNotFis.Close
        de_informa.Sel_LogNotFis
        gridLog.DataMember = "sel_lognotfis"
        gridLog.Refresh
        DoEvents
        
        Close #1
        On Error Resume Next
        FileCopy File1.Path & "\" & lblArquivo.Caption, File1.Path & "\critica\" & lblArquivo.Caption
        Kill File1.Path & "\" & lblArquivo.Caption
        File1.Path = "C:\EDI"
        File1.Refresh
        lblArquivo.Caption = ""
        Timer1.Interval = 15000
        Timer2.Interval = 60000
        Exit Sub
    End If
    
    'rotina para verificar os dados do arquivo e verificar se é versão 3 ou 3.1 ou mesmo se é PROCEDA
    
    Progress1.Max = xtamarq
    DoEvents
    
    'Abre o arquivo para leitura (txt)
    
    Open File1.Path & "\" & lblArquivo.Caption For Input As #1
    xLin = 0
    
    xdataimp = datahora("DATAHORA")
    
    Do Until EOF(1)
        
        Line Input #1, xlinha
        xgravar = "N"
        DoEvents
        
        If Mid$(xlinha, 1, 3) = "000" Then   'CABEÇALHO DO ARQUIVO
            xdatainterface = "20" & Mid(xlinha, 78, 2) & "/" & Mid(xlinha, 76, 2) & "/" & Mid(xlinha, 74, 2)
        End If
        
        If Mid$(xlinha, 1, 3) = "310" Then   'IDENTIFICAÇÃO DO ARQUIVO
            xid_notfis = Mid$(xlinha, 4, 14)
            xid_notfis = xid_notfis + xIdRomaneio
        End If
        
        If Mid$(xlinha, 1, 3) = "311" Then   'REMETENTE/EMBARCADOR
            xremet_cgc = Mid$(xlinha, 4, 14)
            
            If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
            de_informa.Sel_CadCliCGC xremet_cgc
            
            If de_informa.rsSel_CadCliCGC.RecordCount > 0 Then
                xremet_nome = Trim$(de_informa.rsSel_CadCliCGC.Fields("nome"))
            Else
                de_informa.Ins_LogNotFis "CRITICA", "CLIENTE REMETENTE NAO CADASTRADO. CNPJ: " & xremet_cgc & " - ARQUIVO NÃO PROCESSADO.", lblArquivo.Caption
                
                If de_informa.rsSel_LogNotFis.State = 1 Then de_informa.rsSel_LogNotFis.Close
                de_informa.Sel_LogNotFis
                gridLog.DataMember = "sel_lognotfis"
                gridLog.Refresh
                DoEvents
                
                Close #1
                On Error Resume Next
                FileCopy File1.Path & "\" & lblArquivo.Caption, File1.Path & "\critica\" & lblArquivo.Caption
                Kill File1.Path & "\" & lblArquivo.Caption
                File1.Path = "C:\EDI"
                File1.Refresh
                lblArquivo.Caption = ""
                Timer1.Interval = 15000
                Timer2.Interval = 60000
                Exit Sub
            End If
            
        End If
    
        If Mid$(xlinha, 1, 3) = "312" Then   'DESTINATÁRIO
            xdest_cgc = zeros2(Mid$(xlinha, 44, 14), 14)
            xdest_nome = UCase(Trim$(Mid$(xlinha, 4, 40)))
            xdest_ie = Trim$(Mid$(xlinha, 58, 15))
            xdest_end = UCase(Trim$(Mid$(xlinha, 73, 40)))
            xdest_bairro = UCase(Trim$(Mid$(xlinha, 113, 20)))
            xdest_cidade = UCase(Trim$(Mid$(xlinha, 133, 35)))
            xdest_uf = UCase(Mid$(xlinha, 186, 2))
            xdest_cep = SoNumeros(Trim$(Mid$(xlinha, 168, 9)))
            
            'cadastrar destinatário caso não esteja cadastrado
            'If de_informa.rsSel_CadCliCGC.State = 1 Then de_informa.rsSel_CadCliCGC.Close
            'de_informa.Sel_CadCliCGC xdest_cgc
            
            'If de_informa.rsSel_CadCliCGC.RecordCount > 0 Then
            '    If Len(Trim$(xdest_end)) > 4 And _
            '       Len(Trim$(xdest_cidade)) > 2 And _
            '       Len(Trim$(xdest_uf)) = 2 Then  'tem o número do endereço
'           '         de_informa.Alt_CadCliEDI xdest_nome, xdest_end, xdest_cidade, xdest_uf, xdest_cep, xdest_ie, xdest_cgc
            '    End If
            'Else
            '    de_informa.Ins_CadCli xdest_cgc, xdest_nome, "", "", xdest_end, "", xdest_cep, xdest_cidade, _
            '                          xdest_uf, xdest_ie, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
            '                          "", "", "", "", "", "", "", "", "", "TAB000", "AUTO-EDI", "DES", ""
            'End If
            
        End If
        
        If Mid$(xlinha, 1, 3) = "313" Then   'NOTAS FISCAIS
        
            xgravar = "S"
            xtipocarga = Mid$(xlinha, 28, 1)
            If xtipocarga = "1" Then
                xtipocarga = "FRIA"
            ElseIf xtipocarga = "2" Then
                xtipocarga = "SECA"
            ElseIf xtipocarga = "3" Then
                xtipocarga = "MISTA"
            Else
                xtipocarga = ""
            End If
            xtipofrete = Mid$(xlinha, 29, 1)
            If xtipofrete = "C" Then
                xtipofrete = "CIF"
            ElseIf xtipofrete = "F" Then
                xtipofrete = "FOB"
            Else
                xtipofrete = ""
            End If
            xNumNF = Trim$(CVar(CDbl(SoNumeros(Mid$(xlinha, 33, 8)))))
            xNumNfNum = CDbl(xNumNF)
            xserie = Trim$(Mid$(xlinha, 30, 3))
            If Not IsNumeric(xserie) Then
                de_informa.Ins_LogNotFis "CRITICA DADO", "NF " & xNumNF & " COM SERIE INVALIDA: " & xserie, lblArquivo.Caption
            End If
            If IsDate((Mid$(xlinha, 45, 4) & "/" & Mid$(xlinha, 43, 2) & "/" & Mid$(xlinha, 41, 2))) Then
                xemissaonf = CDate(Mid$(xlinha, 45, 4) & "/" & Mid$(xlinha, 43, 2) & "/" & Mid$(xlinha, 41, 2))
            Else
                xemissaonf = CDate("1900/01/01")
            End If
            xnatureza = UCase(Trim$(Mid$(xlinha, 49, 15)))
            xespecie = UCase(Trim$(Mid$(xlinha, 64, 15)))
            xvalmerc = CDbl(Mid$(xlinha, 86, 15)) / 100
            If xvalmerc = 0 Then
                de_informa.Ins_LogNotFis "CRITICA DADO", "VALOR DE MERCADORIA ZERADO. NF: " & xNumNF, lblArquivo.Caption
            End If
            xvolumes = CDbl(Mid$(xlinha, 79, 7)) / 100
            xpeso = CDbl(Mid$(xlinha, 101, 7)) / 100
            If xpeso = 0 Then
                de_informa.Ins_LogNotFis "CRITICA DADO", "PESO DA MERCADORIA ZERADO. NF: " & xNumNF, lblArquivo.Caption
            End If
            If IsNumeric(Mid$(xlinha, 108, 5)) Then
                xpesocub = CDbl(Mid$(xlinha, 108, 5)) / 100
            Else
                xpesocub = 0
            End If
            xLin = xLin + 1
            Progress1.Value = xLin
            lblTotReg = xLin
            DoEvents
        End If
        
        If xgravar = "S" Then
            If de_informa.rsSel_NFNotFis.State = 1 Then de_informa.rsSel_NFNotFis.Close
            de_informa.Sel_NFNotFis xremet_cgc, xNumNfNum, xserie
            If de_informa.rsSel_NFNotFis.RecordCount > 0 Then
                de_informa.Alt_AcertoNotFis xid_notfis, xdest_cgc, xdest_nome, xdest_ie, _
                                            xdest_end, xdest_bairro, xdest_cidade, xdest_uf, xdest_cep, xtipocarga, _
                                            xtipofrete, xemissaonf, xnatureza, xespecie, _
                                            xvolumes, xvalmerc, xpeso, xpesocub, xdataimp, _
                                            xdatainterface, "", "0", xserie, CDbl(xNumNF)
            Else
                de_informa.Ins_NotFis xid_notfis, xremet_cgc, xremet_nome, xdest_cgc, xdest_nome, xdest_ie, _
                                      xdest_end, xdest_bairro, xdest_cidade, xdest_uf, xdest_cep, xtipocarga, _
                                      xtipofrete, xNumNF, xNumNfNum, xserie, xemissaonf, xnatureza, xespecie, _
                                      xvolumes, xvalmerc, xpeso, xpesocub, xdataimp, xdatainterface, "", "0"
            End If
        End If
    Loop
    
    Close #1
    de_informa.Ins_LogNotFis "FINALIZADO", "IMPORTACAO DO ARQUIVO", lblArquivo.Caption
    
    If de_informa.rsSel_LogNotFis.State = 1 Then de_informa.rsSel_LogNotFis.Close
    de_informa.Sel_LogNotFis
    gridLog.DataMember = "sel_lognotfis"
    gridLog.Refresh
    DoEvents
    
'    On Error Resume Next
    FileCopy File1.Path & "\" & lblArquivo.Caption, File1.Path & "\backup\" & lblArquivo.Caption
    Kill File1.Path & "\" & lblArquivo.Caption
    File1.Path = "C:\EDI"
    File1.Refresh
    lblArquivo.Caption = ""
    Timer1.Interval = 15000
    Timer2.Interval = 60000
    Exit Sub
    
TrataErro:
    de_informa.Ins_LogNotFis "ERRO", "NF: " & xNumNF & " - " & Mid$(Err.Description, 1, 200), lblArquivo.Caption
    
    If de_informa.rsSel_LogNotFis.State = 1 Then de_informa.rsSel_LogNotFis.Close
    de_informa.Sel_LogNotFis
    gridLog.DataMember = "sel_lognotfis"
    gridLog.Refresh
    DoEvents
    
    On Error Resume Next
    Close #1
    FileCopy File1.Path & "\" & lblArquivo.Caption, File1.Path & "\backup\" & lblArquivo.Caption
    Kill File1.Path & "\" & lblArquivo.Caption
    File1.Path = "C:\EDI"
    File1.Refresh
    lblArquivo.Caption = ""
    Timer1.Interval = 15000
    Timer2.Interval = 60000
    Exit Sub
End Sub
Private Sub Timer2_Timer()
    Timer2.Interval = 0
    lblCont = Int(lblCont) + 1
    On Error Resume Next
    If Int(lblCont) >= 15 Then
        If de_informa.rsSel_Teste.State = 1 Then de_informa.rsSel_Teste.Close
        de_informa.Sel_Teste
        lblCont = "0"
    End If
    Timer2.Interval = 60000
End Sub
