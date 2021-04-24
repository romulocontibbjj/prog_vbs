VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListaFaturas 
   Caption         =   "Lista Faturas Gravadas"
   ClientHeight    =   7905
   ClientLeft      =   1185
   ClientTop       =   1590
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FraPeriodo 
      Caption         =   "Opções de Seleção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   11655
      Begin VB.Frame Frame2 
         Caption         =   "Filial"
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
         Left            =   6180
         TabIndex        =   34
         Top             =   1080
         Width           =   735
         Begin VB.TextBox txtfilial 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   180
            TabIndex        =   11
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Relatório / Arquivo"
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
         TabIndex        =   29
         Top             =   1920
         Width           =   7335
         Begin VB.OptionButton optRelAnalitico 
            Caption         =   "Analítico"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4920
            TabIndex        =   33
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optRelSintetico 
            Caption         =   "Sintético"
            Enabled         =   0   'False
            Height          =   195
            Left            =   6120
            TabIndex        =   32
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdImpressao 
            Caption         =   "Imprimir Relatório ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   31
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton cmdGerarArq 
            Caption         =   "Gerar Arquivo ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Relatório - Ordernar/Totalizar Por"
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
         Left            =   6960
         TabIndex        =   28
         Top             =   1080
         Width           =   4575
         Begin VB.OptionButton optTotEmissao 
            Caption         =   "Emissão"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTotCliente 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   1200
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optTotVenc 
            Caption         =   "Vencimento"
            Height          =   255
            Left            =   2160
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optTotPagto 
            Caption         =   "Pagamento"
            Height          =   255
            Left            =   3360
            TabIndex        =   15
            Top             =   360
            Width           =   1120
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Faturas"
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
         Left            =   7440
         TabIndex        =   27
         Top             =   300
         Width           =   4110
         Begin VB.OptionButton optFatTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optFatAVencer 
            Caption         =   "À Vencer"
            Enabled         =   0   'False
            Height          =   195
            Left            =   960
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optFatVencidas 
            Caption         =   "Vencidas"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1950
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optFatQuitadas 
            Caption         =   "Quitadas"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3000
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cliente"
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
         TabIndex        =   24
         Top             =   1080
         Width           =   6015
         Begin VB.CommandButton cmdBuscaCli 
            Caption         =   "?"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2160
            TabIndex        =   10
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtCnpj 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            MaxLength       =   14
            TabIndex        =   9
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblCliente 
            BackColor       =   &H8000000A&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2640
            TabIndex        =   26
            Top             =   360
            Width           =   3240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data do Período"
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
         Left            =   3840
         TabIndex        =   23
         Top             =   300
         Width           =   3495
         Begin VB.OptionButton optDataPagto 
            Caption         =   "Pagamento"
            Height          =   255
            Left            =   2280
            TabIndex        =   4
            Top             =   360
            Width           =   1120
         End
         Begin VB.OptionButton optDataVenc 
            Caption         =   "Vencimento"
            Height          =   255
            Left            =   1080
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optDataEmissao 
            Caption         =   "Emissão"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   495
         Left            =   9720
         TabIndex        =   17
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar"
         Height          =   495
         Left            =   7800
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "No Período de"
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
         TabIndex        =   20
         Top             =   300
         Width           =   3705
         Begin VB.CheckBox chkCTC 
            Caption         =   "CTC"
            Height          =   255
            Left            =   3000
            TabIndex        =   35
            Top             =   360
            Width           =   615
         End
         Begin MSMask.MaskEdBox mskPer2 
            Height          =   285
            Left            =   1680
            TabIndex        =   1
            Top             =   360
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskPer1 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12648447
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "à"
            Height          =   195
            Left            =   1440
            TabIndex        =   21
            Top             =   360
            Width           =   90
         End
      End
   End
   Begin VB.Frame fraDados 
      Height          =   4935
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   11655
      Begin MSDataGridLib.DataGrid gridRelFatura 
         Bindings        =   "frmListaFaturas.frx":0000
         Height          =   4575
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8070
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
         DataMember      =   "Sel_RelFaturas1"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Fatura"
            Caption         =   "Filial-Fatura"
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
            DataField       =   "emissao"
            Caption         =   "Emissão"
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
            DataField       =   "cliente_nome"
            Caption         =   "Cliente Nome"
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
            DataField       =   "vencimento"
            Caption         =   "Vencimento"
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
            DataField       =   "valorfatura"
            Caption         =   "Valor Fatura"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "obsfatura"
            Caption         =   "Observação"
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
            DataField       =   "status"
            Caption         =   "St."
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
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2910,047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2954,835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   299,906
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmListaFaturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGerarArq_Click()
    Dim xrs As Recordset, xlinha As String, xFiles As String
    
    If optDataEmissao = True Then  'periodo: data de emissao
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3
        End If
        
    ElseIf optDataVenc = True Then
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0A
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1A
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2A
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3A
        End If
        
    ElseIf optDataPagto = True Then
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0B
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1B
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2B
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3B
        End If
        
    End If
    
    If xrs.RecordCount > 0 Then
    
        xFiles = "C:\INFORMA\FAT" & "_" & zeros(Day(Date), 2) & zeros(Month(Date), 2) & "_" & Mid$(Trim$(CVar(Time())), 1, 2) & Mid$(Trim$(CVar(Time())), 4, 2) & ".txt"
        
        Open xFiles For Output As #1
        
        xlinha = "Num.Fatura#Emissao#Cliente#Vencimento#Pagamento#Valor#OBS#Status#"
        
        Print #1, xlinha
    
        Do Until xrs.EOF
        
            xlinha = xrs.Fields("fatura") & "#" & _
                     xrs.Fields("emissao") & "#" & _
                     xrs.Fields("cliente_nome") & "#" & _
                     xrs.Fields("vencimento") & "#" & _
                     IIf(IsNull(xrs.Fields("pagamento")), "", xrs.Fields("pagamento")) & "#" & _
                     xrs.Fields("valorfatura") & "#" & _
                     xrs.Fields("obsfatura") & "#" & _
                     xrs.Fields("status") & "#"
            Print #1, xlinha
            
            xrs.MoveNext
            
        Loop
        
        xrs.MoveFirst
        
        Close #1
        
        MsgBox "Arquivo Gerado ! " & xFiles & Chr(10) + Chr(13) + Chr(10) + Chr(13) + _
               "O Arquivo Gerado é do Tipo TEXTO Delimitado e pode ser Aberto no Excel utilizando o caracter # como delimitador.", vbInformation, "Arquivo Gerado"
        
    End If
        
    
    

End Sub

Private Sub cmdImpressao_Click()
    Dim xcont As Integer, xrs As Recordset, xsubtot As Currency, xtotger As Currency, xcontsub As Integer
    Dim xpag As Long
    
    If optDataEmissao = True Then  'periodo: data de emissao
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3
        End If
        
    ElseIf optDataVenc = True Then
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0A
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1A
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2A
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3A
        End If
        
    ElseIf optDataPagto = True Then
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0B
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1B
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2B
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3B
        End If
        
    End If
        
    fraDados.Enabled = False
    FraPeriodo.Enabled = False
    cmdImpressao.Caption = "Aguarde Impressão"
    
    'busca impressora para este documento
    If Dir("C:\informa.cfg") <> "" Then
        
        Open "C:\informa.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "REL" Then
                ximpr_cfg = Trim$(Mid$(xlinha, 5))
                Exit Do
            End If
        Loop
        
        Close #1
        
    Else
    
        MsgBox "Não está Configurado a Impressora para Este Documento: REL "
        fraDados.Enabled = True
        FraPeriodo.Enabled = True
        cmdImpressao.Caption = "Imprimir Relatório ..."
        Exit Sub
        
    End If
    
    'seta impressora
    
    For Each ximpr_inst In Printers
        If ximpr_inst.DeviceName = ximpr_cfg Then
            Set Printer = ximpr_inst
            DoEvents
            Exit For
        End If
    Next
    
    xcont = 1
    xpag = 1
    xcontsub = 0
    xrs.MoveFirst
    
    Printer.FontName = "Courier New"
    Printer.FontSize = 8
    
    xsubtot = 0
    xtotger = 0
        
    Do Until xrs.EOF
    
        If xcont = 1 Then
                
            Printer.FontBold = True
            Printer.Print "RELATÓRIO DE FATURAMENTO"
            If Len(Trim$(txtCnpj)) >= 8 Then
                Printer.Print "CLIENTE: " & Mid$(lblCliente, 25)
            Else
                Printer.Print "CLIENTE: TODOS"
            End If
            Printer.Print "PERÍODO: " & mskPer1.Text & " à " & mskPer2.Text;
            If optDataEmissao = True Then
                Printer.Print " (EMISSÃO)   ";
            ElseIf optDataVenc = True Then
                Printer.Print " (VENCIMENTO)";
            ElseIf optDataPagto = True Then
                Printer.Print " (PAGAMENTO) ";
            End If
            If optFatAVencer.Enabled = True Then
                Printer.Print "    FATURAS: À VENCER";
            ElseIf optFatQuitadas = True Then
                Printer.Print "    FATURAS: QUITADAS";
            ElseIf optFatVencidas = True Then
                Printer.Print "    FATURAS: VENCIDAS";
            ElseIf optFatTodas = True Then
                Printer.Print "    FATURAS: TODAS   ";
            End If
            If optTotEmissao = True Then
                Printer.Print "ORDEM: POR EMISSAO                          ";
            ElseIf optTotCliente = True Then
                Printer.Print "ORDEM: POR CLIENTE                          ";
            ElseIf optTotVenc = True Then
                Printer.Print "ORDEM: POR VENCIMENTO                       ";
            ElseIf optTotPagto = True Then
                Printer.Print "ORDEM: POR PAGAMENTO                        ";
            End If
            Printer.Print "Pág: " & zeros(xpag, 3)
            Printer.Print "----------------------------------------------------------------------------------------------------------------------"
            Printer.Print "Num.Fatura  Emissão   Cliente                   Vencimento  Pagamento         Valor R$  OBS                        St."
            Printer.Print "----------------------------------------------------------------------------------------------------------------------"
            Printer.FontBold = False
            xcont = 8
        End If
        
        Printer.Print xrs.Fields("fatura"); Spc(2);
        Printer.Print Trim$(xrs.Fields("emissao")); Spc(1);
        Printer.Print Trim$(Mid$(xrs.Fields("cliente_nome"), 1, 25)) & _
                      String(25 - Len(Trim$(Mid$(xrs.Fields("cliente_nome"), 1, 25))), " "); Spc(1);
        Printer.Print Trim$(xrs.Fields("vencimento")); Spc(2);
        If IsNull(xrs.Fields("pagamento")) Then
            Printer.Print "          "; Spc(3);
        Else
            If xrs.Fields("pagamento") = CDate("1900/01/01") Then
                Printer.Print "          "; Spc(3);
            Else
                Printer.Print Trim$(xrs.Fields("pagamento")); Spc(3);
            End If
        End If
        Printer.Print String(13 - Len(Format(xrs.Fields("valorfatura"), "##,###,##0.00")), " "); Format(xrs.Fields("valorfatura"), "##,###,##0.00"); Spc(2);
        Printer.Print Trim$(Mid$(xrs.Fields("obsfatura"), 1, 25)) & _
                      String(25 - Len(Trim$(Mid$(xrs.Fields("obsfatura"), 1, 25))), " "); Spc(3);
        Printer.Print xrs.Fields("status")
        xcont = xcont + 1
        
        xsubtot = xsubtot + xrs.Fields("valorfatura")
        xcontsub = xcontsub + 1
        xtotger = xtotger + xrs.Fields("valorfatura")
        
        If optTotEmissao.Value = True Then
            xemi_ant = Trim$(xrs.Fields("emissao"))
        ElseIf optTotCliente.Value = True Then
            xcli_ant = Mid$(xrs.Fields("cliente_nome"), 1, 25)
        ElseIf optTotVenc.Value = True Then
            xvenc_ant = Trim$(xrs.Fields("vencimento"))
        ElseIf optTotPagto.Value = True Then
            xpag_ant = Trim$(xrs.Fields("pagamento"))
        End If
        
        xrs.MoveNext
        
        'sub total
        If xrs.EOF Then
            If xcontsub > 1 Then
                Printer.Print Space(46); "Sub Total .............."; Spc(3);
                Printer.Print String(13 - Len(Format(xsubtot, "##,###,##0.00")), " "); Format(xsubtot, "##,###,##0.00")
            End If
        Else
            If optTotEmissao.Value = True Then
                If Trim$(xrs.Fields("emissao")) <> xemi_ant And xcontsub > 1 Then
                    Printer.Print Space(46); "Sub Total .............."; Spc(3);
                    Printer.Print String(13 - Len(Format(xsubtot, "##,###,##0.00")), " "); Format(xsubtot, "##,###,##0.00")
                    Printer.Print
                    xcontsub = 0
                    xsubtot = 0
                    xcont = xcont + 1
                Else
                    If Trim$(xrs.Fields("emissao")) <> xemi_ant Then
                        Printer.Print
                        xcontsub = 0
                        xsubtot = 0
                        xcont = xcont + 1
                    End If
                End If
            ElseIf optTotCliente.Value = True Then
                If Mid$(xrs.Fields("cliente_nome"), 1, 25) <> xcli_ant And xcontsub > 1 Then
                    Printer.Print Space(46); "Sub Total .............."; Spc(3);
                    Printer.Print String(13 - Len(Format(xsubtot, "##,###,##0.00")), " "); Format(xsubtot, "##,###,##0.00")
                    Printer.Print
                    xcontsub = 0
                    xsubtot = 0
                    xcont = xcont + 1
                Else
                    If Mid$(xrs.Fields("cliente_nome"), 1, 25) <> xcli_ant Then
                        Printer.Print
                        xcontsub = 0
                        xsubtot = 0
                        xcont = xcont + 1
                    End If
                End If
            ElseIf optTotVenc.Value = True Then
                If Trim$(xrs.Fields("vencimento")) <> xvenc_ant And xcontsub > 1 Then
                    Printer.Print Space(46); "Sub Total .............."; Spc(3);
                    Printer.Print String(13 - Len(Format(xsubtot, "##,###,##0.00")), " "); Format(xsubtot, "##,###,##0.00")
                    Printer.Print
                    xcontsub = 0
                    xsubtot = 0
                    xcont = xcont + 1
                Else
                    If Trim$(xrs.Fields("vencimento")) <> xvenc_ant Then
                        Printer.Print
                        xcontsub = 0
                        xsubtot = 0
                        xcont = xcont + 1
                    End If
                End If
            ElseIf optTotPagto.Value = True Then
                If Trim$(xrs.Fields("pagamento")) <> xpag_ant And xcontsub > 1 Then
                    Printer.Print Space(46); "Sub Total .............."; Spc(3);
                    Printer.Print String(13 - Len(Format(xsubtot, "##,###,##0.00")), " "); Format(xsubtot, "##,###,##0.00")
                    Printer.Print
                    xcontsub = 0
                    xsubtot = 0
                    xcont = xcont + 1
                Else
                    If Trim$(xrs.Fields("pagamento")) <> xpag_ant Then
                        Printer.Print
                        xcontsub = 0
                        xsubtot = 0
                        xcont = xcont + 1
                    End If
                End If
            End If
        End If
        
        If xcont >= 80 Then
            Printer.FontBold = True
            Printer.Print "----------------------------------------------------------------------------------------------------------------------"
            xcont = 1
            xpag = xpag + 1
            Printer.FontBold = False
            Printer.NewPage
        End If
        
    Loop
    
    Printer.FontBold = True
    Printer.Print "----------------------------------------------------------------------------------------------------------------------"
    Printer.Print Space(46); "TOTAL GERAL ............"; Spc(3);
    Printer.Print String(13 - Len(Format(xtotger, "##,###,##0.00")), " "); Format(xtotger, "##,###,##0.00")
    Printer.Print "----------------------------------------------------------------------------------------------------------------------"
    Printer.FontBold = False
    
    xrs.MoveFirst

    Printer.EndDoc

    MsgBox "Dados Enviados para Impressão !", vbInformation

    fraDados.Enabled = True
    FraPeriodo.Enabled = True
    cmdImpressao.Caption = "Imprimir Relatório ..."
    mskPer1.SetFocus
    
End Sub
Private Sub cmdProcessar_Click()
    Dim xrs As Recordset
    
    If optDataEmissao = True Then  'periodo: data de emissao
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3
        End If
        
        If xrs.State = 1 Then xrs.Close
        
        If optTotEmissao.Value = True Then
            de_informa.Sel_RelFaturas0 CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas0"
            gridRelFatura.Refresh
        ElseIf optTotCliente.Value = True Then
            de_informa.Sel_RelFaturas1 CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas1"
            gridRelFatura.Refresh
        ElseIf optTotVenc.Value = True Then
            de_informa.Sel_RelFaturas2 CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas2"
            gridRelFatura.Refresh
        ElseIf optTotPagto.Value = True Then
            de_informa.Sel_RelFaturas3 CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas3"
            gridRelFatura.Refresh
        End If
        
    ElseIf optDataVenc = True Then
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0A
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1A
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2A
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3A
        End If
        
        If xrs.State = 1 Then xrs.Close
        
        If optTotEmissao.Value = True Then
            de_informa.Sel_RelFaturas0a CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas0a"
            gridRelFatura.Refresh
        ElseIf optTotCliente.Value = True Then
            de_informa.Sel_RelFaturas1a CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas1a"
            gridRelFatura.Refresh
        ElseIf optTotVenc.Value = True Then
            de_informa.Sel_RelFaturas2a CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas2a"
            gridRelFatura.Refresh
        ElseIf optTotPagto.Value = True Then
            de_informa.Sel_RelFaturas3a CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas3a"
            gridRelFatura.Refresh
        End If
    
    ElseIf optDataPagto = True Then
    
        If optTotEmissao = True Then
            Set xrs = de_informa.rsSel_RelFaturas0B
        ElseIf optTotCliente = True Then
            Set xrs = de_informa.rsSel_RelFaturas1B
        ElseIf optTotVenc = True Then
            Set xrs = de_informa.rsSel_RelFaturas2B
        ElseIf optTotPagto = True Then
            Set xrs = de_informa.rsSel_RelFaturas3B
        End If
        
        If xrs.State = 1 Then xrs.Close
        
        If optTotEmissao.Value = True Then
            de_informa.Sel_RelFaturas0b CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas0b"
            gridRelFatura.Refresh
        ElseIf optTotCliente.Value = True Then
            de_informa.Sel_RelFaturas1b CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas1b"
            gridRelFatura.Refresh
        ElseIf optTotVenc.Value = True Then
            de_informa.Sel_RelFaturas2b CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas2b"
            gridRelFatura.Refresh
        ElseIf optTotPagto.Value = True Then
            de_informa.Sel_RelFaturas3b CDate(mskPer1), CDate(mskPer2), txtCnpj & "%", "%", Trim$(txtFilial) & "%"
            gridRelFatura.DataMember = "Sel_RelFaturas3b"
            gridRelFatura.Refresh
        End If
            
    End If
        
    If xrs.RecordCount < 1 Then
        MsgBox "Não Há Dados Para as Seleções Escolhidas !", vbInformation, "Ops"
        mskPer1.SetFocus
        Exit Sub
    End If
    
    If xrs.RecordCount > 0 Then
        cmdImpressao.Enabled = True
        cmdGerarArq.Enabled = True
        optRelAnalitico.Enabled = True
        cmdImpressao.SetFocus
    Else
        cmdImpressao.Enabled = False
        cmdGerarArq.Enabled = False
        optRelAnalitico.Enabled = False
        cmdImpressao.SetFocus
    End If
    
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    mdiFatura.ToolFaturamento.Visible = False
    If de_informa.rsSel_RelFaturas1.State = 1 Then de_informa.rsSel_RelFaturas1.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1"
    gridRelFatura.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiFatura.ToolFaturamento.Visible = True
End Sub

Private Sub mskPer1_GotFocus()
    mskPer1.SelStart = 0
    mskPer1.SelLength = 10
End Sub
Private Sub mskPer1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub mskPer1_LostFocus()
    If mskPer1.Text <> "__/__/____" Then
        mskPer1.Text = century(mskPer1.Text)
        If IsDate(mskPer1.Text) = False Or Mid(mskPer1.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPer1.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub mskPer2_GotFocus()
    mskPer2.SelStart = 0
    mskPer2.SelLength = 10
End Sub
Private Sub mskPer2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub mskPer2_LostFocus()
    If mskPer2.Text <> "__/__/____" Then
        mskPer2.Text = century(mskPer2.Text)
        If IsDate(mskPer2.Text) = False Or Mid(mskPer2.Text, 4, 2) > 12 Then
            MsgBox "Data Inválida !", vbCritical, "Erro"
            mskPer2.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub optDataEmissao_Click()

    If de_informa.rsSel_RelFaturas1.State = 1 Then de_informa.rsSel_RelFaturas1.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1"
    If de_informa.rsSel_RelFaturas2.State = 1 Then de_informa.rsSel_RelFaturas2.Close
    gridRelFatura.DataMember = "Sel_RelFaturas2"
    If de_informa.rsSel_RelFaturas3.State = 1 Then de_informa.rsSel_RelFaturas3.Close
    gridRelFatura.DataMember = "Sel_RelFaturas3"
    If de_informa.rsSel_RelFaturas1A.State = 1 Then de_informa.rsSel_RelFaturas1A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas2A.State = 1 Then de_informa.rsSel_RelFaturas2A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas3A.State = 1 Then de_informa.rsSel_RelFaturas3A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas1B.State = 1 Then de_informa.rsSel_RelFaturas1B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas2B.State = 1 Then de_informa.rsSel_RelFaturas2B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas3B.State = 1 Then de_informa.rsSel_RelFaturas3B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    gridRelFatura.Refresh
    

End Sub

Private Sub optDataEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optDataPagto_Click()

    If de_informa.rsSel_RelFaturas1.State = 1 Then de_informa.rsSel_RelFaturas1.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1"
    If de_informa.rsSel_RelFaturas2.State = 1 Then de_informa.rsSel_RelFaturas2.Close
    gridRelFatura.DataMember = "Sel_RelFaturas2"
    If de_informa.rsSel_RelFaturas3.State = 1 Then de_informa.rsSel_RelFaturas3.Close
    gridRelFatura.DataMember = "Sel_RelFaturas3"
    If de_informa.rsSel_RelFaturas1A.State = 1 Then de_informa.rsSel_RelFaturas1A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas2A.State = 1 Then de_informa.rsSel_RelFaturas2A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas3A.State = 1 Then de_informa.rsSel_RelFaturas3A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas1B.State = 1 Then de_informa.rsSel_RelFaturas1B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas2B.State = 1 Then de_informa.rsSel_RelFaturas2B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas3B.State = 1 Then de_informa.rsSel_RelFaturas3B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    gridRelFatura.Refresh
    

End Sub

Private Sub optDataPagto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optDataVenc_Click()

    If de_informa.rsSel_RelFaturas1.State = 1 Then de_informa.rsSel_RelFaturas1.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1"
    If de_informa.rsSel_RelFaturas2.State = 1 Then de_informa.rsSel_RelFaturas2.Close
    gridRelFatura.DataMember = "Sel_RelFaturas2"
    If de_informa.rsSel_RelFaturas3.State = 1 Then de_informa.rsSel_RelFaturas3.Close
    gridRelFatura.DataMember = "Sel_RelFaturas3"
    If de_informa.rsSel_RelFaturas1A.State = 1 Then de_informa.rsSel_RelFaturas1A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas2A.State = 1 Then de_informa.rsSel_RelFaturas2A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas3A.State = 1 Then de_informa.rsSel_RelFaturas3A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas1B.State = 1 Then de_informa.rsSel_RelFaturas1B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas2B.State = 1 Then de_informa.rsSel_RelFaturas2B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas3B.State = 1 Then de_informa.rsSel_RelFaturas3B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    gridRelFatura.Refresh
    

End Sub

Private Sub optDataVenc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optFatAVencer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optFatQuitadas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optFatTodas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optFatVencidas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub optTotCliente_Click()

    If de_informa.rsSel_RelFaturas1.State = 1 Then de_informa.rsSel_RelFaturas1.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1"
    If de_informa.rsSel_RelFaturas2.State = 1 Then de_informa.rsSel_RelFaturas2.Close
    gridRelFatura.DataMember = "Sel_RelFaturas2"
    If de_informa.rsSel_RelFaturas3.State = 1 Then de_informa.rsSel_RelFaturas3.Close
    gridRelFatura.DataMember = "Sel_RelFaturas3"
    If de_informa.rsSel_RelFaturas1A.State = 1 Then de_informa.rsSel_RelFaturas1A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas2A.State = 1 Then de_informa.rsSel_RelFaturas2A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas3A.State = 1 Then de_informa.rsSel_RelFaturas3A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas1B.State = 1 Then de_informa.rsSel_RelFaturas1B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas2B.State = 1 Then de_informa.rsSel_RelFaturas2B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas3B.State = 1 Then de_informa.rsSel_RelFaturas3B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    gridRelFatura.Refresh
    
End Sub

Private Sub optTotCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optTotEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optTotPagto_Click()

    If de_informa.rsSel_RelFaturas1.State = 1 Then de_informa.rsSel_RelFaturas1.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1"
    If de_informa.rsSel_RelFaturas2.State = 1 Then de_informa.rsSel_RelFaturas2.Close
    gridRelFatura.DataMember = "Sel_RelFaturas2"
    If de_informa.rsSel_RelFaturas3.State = 1 Then de_informa.rsSel_RelFaturas3.Close
    gridRelFatura.DataMember = "Sel_RelFaturas3"
    If de_informa.rsSel_RelFaturas1A.State = 1 Then de_informa.rsSel_RelFaturas1A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas2A.State = 1 Then de_informa.rsSel_RelFaturas2A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas3A.State = 1 Then de_informa.rsSel_RelFaturas3A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas1B.State = 1 Then de_informa.rsSel_RelFaturas1B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas2B.State = 1 Then de_informa.rsSel_RelFaturas2B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas3B.State = 1 Then de_informa.rsSel_RelFaturas3B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    gridRelFatura.Refresh
    
End Sub

Private Sub optTotPagto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub optTotVenc_Click()

    If de_informa.rsSel_RelFaturas1.State = 1 Then de_informa.rsSel_RelFaturas1.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1"
    If de_informa.rsSel_RelFaturas2.State = 1 Then de_informa.rsSel_RelFaturas2.Close
    gridRelFatura.DataMember = "Sel_RelFaturas2"
    If de_informa.rsSel_RelFaturas3.State = 1 Then de_informa.rsSel_RelFaturas3.Close
    gridRelFatura.DataMember = "Sel_RelFaturas3"
    If de_informa.rsSel_RelFaturas1A.State = 1 Then de_informa.rsSel_RelFaturas1A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas2A.State = 1 Then de_informa.rsSel_RelFaturas2A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas3A.State = 1 Then de_informa.rsSel_RelFaturas3A.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1a"
    If de_informa.rsSel_RelFaturas1B.State = 1 Then de_informa.rsSel_RelFaturas1B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas2B.State = 1 Then de_informa.rsSel_RelFaturas2B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    If de_informa.rsSel_RelFaturas3B.State = 1 Then de_informa.rsSel_RelFaturas3B.Close
    gridRelFatura.DataMember = "Sel_RelFaturas1b"
    gridRelFatura.Refresh
    
End Sub

Private Sub optTotVenc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If

End Sub

Private Sub txtCnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtFilial_Change()
    If Len(Trim$(txtFilial)) > 0 Then
        If Not IsNumeric(txtFilial) Or Mid$(txtFilial, Len(txtFilial), 1) = "," Or Mid$(txtFilial, Len(txtFilial), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
End Sub

Private Sub txtFilial_GotFocus()
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub
Private Sub txtFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
