VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmConfBonagura 
   Caption         =   "Gera Arquivo de Conferência Bonagura"
   ClientHeight    =   2775
   ClientLeft      =   1230
   ClientTop       =   1185
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   5190
   Begin VB.Frame FraPeriodo 
      Caption         =   "Conferência Bonagura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
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
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   3465
         Begin MSMask.MaskEdBox mskPer2 
            Height          =   285
            Left            =   1920
            TabIndex        =   1
            Top             =   320
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   240
            TabIndex        =   0
            Top             =   320
            Width           =   1290
            _ExtentX        =   2275
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
            Left            =   1680
            TabIndex        =   6
            Top             =   320
            Width           =   90
         End
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "Processar"
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   495
         Left            =   2640
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmConfBonagura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcessar_Click()
    Dim xfilial As String, xfatura As String, xemissao As String, xcliente As String
    Dim xvalor As String, xvencimento As String, xperiodo As String

    If de_informa.rsSel_ConfBonagura.State = 1 Then de_informa.rsSel_ConfBonagura.Close
    de_informa.Sel_ConfBonagura CDate(mskPer1), CDate(mskPer2)
    
    If de_informa.rsSel_ConfBonagura.RecordCount < 1 Then
        MsgBox "Não Há Faturas Para Este Período !", vbCritical
        mskPer1.SetFocus
        Exit Sub
    Else
    
        xFiles = "C:\INFORMA\BONA" & zeros2(Str(Day(CDate(mskPer1))), 2) & _
                   zeros2(Str(Month(CDate(mskPer1))), 2) & _
                   Trim$(Str(Year(CDate(mskPer1)))) & "_" & _
                   zeros2(Str(Day(CDate(mskPer2))), 2) & _
                   zeros2(Str(Month(CDate(mskPer2))), 2) & _
                   Trim$(Str(Year(CDate(mskPer2)))) & ".txt"
        
        Open xFiles For Output As #1
    
        xperiodo = zeros2(Str(Day(CDate(mskPer1))), 2) & _
                   zeros2(Str(Month(CDate(mskPer1))), 2) & _
                   Trim$(Str(Year(CDate(mskPer1)))) & _
                   zeros2(Str(Day(CDate(mskPer2))), 2) & _
                   zeros2(Str(Month(CDate(mskPer2))), 2) & _
                   Trim$(Str(Year(CDate(mskPer2))))
        
        Print #1, xperiodo
        
        Do Until de_informa.rsSel_ConfBonagura.EOF
        
            xfilial = Mid$(de_informa.rsSel_ConfBonagura.Fields("filialfatura"), 1, 2)
            xfatura = Mid$(de_informa.rsSel_ConfBonagura.Fields("filialfatura"), 3, 6)
            xemissao = zeros2(Str(Day(de_informa.rsSel_ConfBonagura.Fields("emissao"))), 2) & _
                       zeros2(Str(Month(de_informa.rsSel_ConfBonagura.Fields("emissao"))), 2) & _
                       Trim$(Str(Year(de_informa.rsSel_ConfBonagura.Fields("emissao"))))
            xcliente = de_informa.rsSel_ConfBonagura.Fields("cliente_cgc")
            xvalor = zeros2(SoNumeros(Format(de_informa.rsSel_ConfBonagura.Fields("valorfatura"), "##,###,##0.00")), 15)
            xvencimento = zeros2(Str(Day(de_informa.rsSel_ConfBonagura.Fields("vencimento"))), 2) & _
                       zeros2(Str(Month(de_informa.rsSel_ConfBonagura.Fields("vencimento"))), 2) & _
                       Trim$(Str(Year(de_informa.rsSel_ConfBonagura.Fields("vencimento"))))
            
            Print #1, xfilial & xfatura & xemissao & xcliente & xvalor & xvencimento
            
            de_informa.rsSel_ConfBonagura.MoveNext
            
        Loop
        
        Close #1
        
        MsgBox "Arquivo Gerado: " & xFiles, vbInformation
        Exit Sub
                       
    End If
    
    
End Sub

Private Sub cmdSair_Click()
Unload Me
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

