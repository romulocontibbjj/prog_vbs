VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_diversosII 
   Caption         =   "DIVERSOS 2 - A REVANCHE"
   ClientHeight    =   9345
   ClientLeft      =   1185
   ClientTop       =   1395
   ClientWidth     =   13845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   13845
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      Begin VB.Frame Frame3 
         Caption         =   "JOHNSON - AVERBAÇÃO"
         Height          =   2415
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   5775
         Begin VB.CommandButton cmd_johnson 
            Caption         =   "JOHNSON"
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&SAIR"
         Height          =   495
         Left            =   11640
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "FATURAMENTO - JOSE"
         Height          =   2415
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5775
         Begin VB.OptionButton opt_urgencia 
            Caption         =   "URGÊNCIA"
            Height          =   255
            Left            =   3720
            TabIndex        =   14
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton opt_normal 
            Caption         =   "NORMAL"
            Height          =   255
            Left            =   3720
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin MSComctlLib.ProgressBar PRG_JOSE 
            Height          =   2055
            Left            =   5280
            TabIndex        =   11
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   3625
            _Version        =   393216
            Appearance      =   1
            Min             =   1e-4
            Orientation     =   1
         End
         Begin VB.CommandButton cmd_rel_jose 
            Caption         =   "Gerar Arquivo"
            Height          =   255
            Left            =   1080
            TabIndex        =   10
            Top             =   1680
            Width           =   1815
         End
         Begin MSMask.MaskEdBox mas_data2 
            Height          =   300
            Left            =   2280
            TabIndex        =   4
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mas_data1 
            Height          =   300
            Left            =   480
            TabIndex        =   3
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txt_doc 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2880
            TabIndex        =   2
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txt_cgc 
            Height          =   285
            Left            =   600
            TabIndex        =   1
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "DATA FINAL:"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "DATA INICIAL:"
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "DOC:"
            Height          =   255
            Left            =   2400
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "CGC:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frm_diversosII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xurg As String

Private Sub cmd_rel_jose_Click()
Dim xdata       As String
Dim xfilial     As String
Dim xctc        As String
Dim xremet      As String
Dim xdest       As String
Dim xcidade     As String
Dim xuf         As String
Dim xfretetotal As String
Dim xmodal      As String
Dim xpeso       As String
Dim xpesotax    As String
Dim xvolumes    As String
Dim xtxurgencias As String
Dim xvalmerc    As String
Dim xobs        As String
Dim xtipodoc    As String
Dim xlinha      As String
Dim xcont       As Integer

If xurg = "N" Then

If deb_bona.rssel_rel_jose.State = 1 Then deb_bona.rssel_rel_jose.Close
    deb_bona.sel_rel_jose txt_cgc.Text & "%", mas_data1, mas_data2, "%" & txt_doc.Text & "%"
    
    If deb_bona.rssel_rel_jose.RecordCount < 1 Then
        MsgBox "Não Há Dados neste período", vbInformation, "FATURAMENTO"
    Else
        PRG_JOSE.Min = xcont
        PRG_JOSE.Max = deb_bona.rssel_rel_jose.RecordCount
        PRG_JOSE.Value = xcont
        
        Open "C:\ABBOTT_NORMAL.TXT" For Output As #1
        xdata = "DATA"
        xfilial = "FILIAL"
        xctc = "CTC"
        xremet = "REMET_NOME"
        xdest = "DEST_NOME"
        xcidade = "CIDADE_DEST"
        xuf = "UF_DEST"
        xfretetotal = "FRETE_TOTAL"
        xmodal = "MODAL"
        xpeso = "PESO"
        xpesotax = "PESOTAX"
        xtxurgencia = "TXURGENCIA"
        xvolumes = "VOLUMES"
        xvalmerc = "VALMERC"
        xobs = "OBS_EMISSAO"
        xtipodoc = "TIPODOC"
        xlinha = xdata & "#" & xfilial & "#" & xctc & "#" & xremet & "#" & _
                xdest & "#" & xcidade & "#" & xuf & "#" & xfretetotal & "#" & _
                xmodal & "#" & xpeso & "#" & xpesotax & "#" & xvolumes & "#" & _
                xvalmerc & "#" & xobs & "#" & xtipodoc
        Print #1, xlinha
        With deb_bona.rssel_rel_jose
        .MoveFirst
        Do Until .EOF
        xdata = .Fields("DATA")
        xfilial = .Fields("FILIAL")
        xctc = .Fields("CTC")
        xremet = .Fields("REMET_NOME")
        xdest = .Fields("DEST_NOME")
        xcidade = .Fields("CIDADE_DEST")
        xuf = .Fields("UF_DEST")
        xfretetotal = .Fields("FRETETOTAL")
        xmodal = .Fields("MODAL")
        xpeso = .Fields("PESO")
        xpesotax = .Fields("PESOTAX")
        xtxurgencia = .Fields("TXURGENCIA")
        xvolumes = .Fields("VOLUMES")
        xvalmerc = .Fields("VALMERC")
        xobs = .Fields("OBS_EMISSAO")
        xtipodoc = .Fields("TIPODOC")
        xlinha = xdata & "#" & xfilial & "#" & xctc & "#" & xremet & "#" & _
                xdest & "#" & xcidade & "#" & xuf & "#" & xfretetotal & "#" & _
                xmodal & "#" & xpeso & "#" & xpesotax & "#" & xvolumes & "#" & _
                xvalmerc & "#" & xobs & "#" & xtipodoc
        Print #1, xlinha
        .MoveNext
        
        xcont = xcont + 1
        
        PRG_JOSE.Value = xcont
        
        
        Loop
        End With
        
        Close #1
        
        MsgBox "ARQUIVO C:\ABBOTT_NOMAL.TXT GERADO COM SUCESSO" & Chr$(13) & deb_bona.rssel_rel_jose.RecordCount & " REGISTROS"
        
        
        
    End If
    

Else

If deb_bona.rssel_rel_jose2.State = 1 Then deb_bona.rssel_rel_jose2.Close
    deb_bona.sel_rel_jose2 txt_cgc.Text & "%", mas_data1, mas_data2, "%" & txt_doc.Text & "%"
    
    If deb_bona.rssel_rel_jose2.RecordCount < 1 Then
        MsgBox "Não Há Dados neste período", vbInformation, "FATURAMENTO"
    Else
        PRG_JOSE.Min = xcont
        PRG_JOSE.Max = deb_bona.rssel_rel_jose2.RecordCount
        PRG_JOSE.Value = xcont
        
        Open "C:\ABBOTT_URG.TXT" For Output As #1
        xdata = "DATA"
        xfilial = "FILIAL"
        xctc = "CTC"
        xremet = "REMET_NOME"
        xdest = "DEST_NOME"
        xcidade = "CIDADE_DEST"
        xuf = "UF_DEST"
        xfretetotal = "FRETE_TOTAL"
        xmodal = "MODAL"
        xpeso = "PESO"
        xpesotax = "PESOTAX"
        xtxurgencia = "TXURGENCIA"
        xvolumes = "VOLUMES"
        xvalmerc = "VALMERC"
        xobs = "OBS_EMISSAO"
        xtipodoc = "TIPODOC"
        xlinha = xdata & "#" & xfilial & "#" & xctc & "#" & xremet & "#" & _
                xdest & "#" & xcidade & "#" & xuf & "#" & xfretetotal & "#" & _
                xmodal & "#" & xpeso & "#" & xpesotax & "#" & xvolumes & "#" & _
                xvalmerc & "#" & xobs & "#" & xtipodoc
        Print #1, xlinha
        With deb_bona.rssel_rel_jose2
        .MoveFirst
        Do Until .EOF
        xdata = .Fields("DATA")
        xfilial = .Fields("FILIAL")
        xctc = .Fields("CTC")
        xremet = .Fields("REMET_NOME")
        xdest = .Fields("DEST_NOME")
        xcidade = .Fields("CIDADE_DEST")
        xuf = .Fields("UF_DEST")
        xfretetotal = .Fields("FRETETOTAL")
        xmodal = .Fields("MODAL")
        xpeso = .Fields("PESO")
        xpesotax = .Fields("PESOTAX")
        xtxurgencia = .Fields("TXURGENCIA")
        xvolumes = .Fields("VOLUMES")
        xvalmerc = .Fields("VALMERC")
        xobs = .Fields("OBS_EMISSAO")
        xtipodoc = .Fields("TIPODOC")
        xlinha = xdata & "#" & xfilial & "#" & xctc & "#" & xremet & "#" & _
                xdest & "#" & xcidade & "#" & xuf & "#" & xfretetotal & "#" & _
                xmodal & "#" & xpeso & "#" & xpesotax & "#" & xvolumes & "#" & _
                xvalmerc & "#" & xobs & "#" & xtipodoc
        Print #1, xlinha
        .MoveNext
        
        xcont = xcont + 1
        
        PRG_JOSE.Value = xcont
        
        
        Loop
        End With
        
        Close #1
        
        MsgBox "ARQUIVO C:\ABBOTT_URG.TXT GERADO COM SUCESSO" & Chr$(13) & deb_bona.rssel_rel_jose2.RecordCount & " REGISTROS"


End If

End If

End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub


Private Sub opt_normal_Click()
If opt_normal.Value = True Then
    xurg = "N"
End If

    
End Sub

Private Sub opt_urgencia_Click()

If opt_urgencia.Value = True Then
    xurg = "S"
End If

    

End Sub
