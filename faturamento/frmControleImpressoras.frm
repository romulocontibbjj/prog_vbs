VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmControleImpressoras 
   Caption         =   "Configuração de Impressão"
   ClientHeight    =   3885
   ClientLeft      =   825
   ClientTop       =   1140
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   9255
   Begin VB.TextBox txtNFS 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdNFS 
      Caption         =   "<="
      Height          =   285
      Left            =   4620
      TabIndex        =   18
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtFaturas 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   17
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdManif 
      Caption         =   "<="
      Height          =   285
      Left            =   4620
      TabIndex        =   15
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdRelat 
      Caption         =   "<="
      Height          =   285
      Left            =   4620
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdCtcs 
      Caption         =   "<="
      Height          =   285
      Left            =   4620
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdFaturas 
      Caption         =   "<="
      Height          =   285
      Left            =   4620
      TabIndex        =   12
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdCtr 
      Caption         =   "<="
      Height          =   285
      Left            =   4620
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressoras Instaladas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Width           =   3855
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexImpressoras 
         Height          =   2295
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4048
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "GRAVAR"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtCtcs 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtRelat 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtManifesto 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtCTR 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   $"frmControleImpressoras.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   22
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Nota de Serviço (Matricial):"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   1920
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fatura (Matricial):"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CTC (Matricial):"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Relatório (Laser/InkJet):"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Manifesto (Matricial):"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CTR (Laser):"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frmControleImpressoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCtcs_Click()
    txtCtcs.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdCtr_Click()
    txtCTR.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdFaturas_Click()
    txtFaturas.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdGravar_Click()
    Open "C:\informa.cfg" For Output As #1
    Print #1, "CTR=" & txtCTR
    Print #1, "CTC=" & txtCtcs
    Print #1, "MNF=" & txtManifesto
    Print #1, "REL=" & txtRelat
    Print #1, "FAT=" & txtFaturas
    Print #1, "NFS=" & txtNFS
    Close #1
    MsgBox "Impressoras Gravadas !"
End Sub

Private Sub cmdManif_Click()
    txtManifesto.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdNFS_Click()
txtNFS.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdRelat_Click()
    txtRelat.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub Command1_Click()
    
    If de_informa.rsSel_Acerto1.State = 1 Then de_informa.rsSel_Acerto1.Close
    de_informa.Sel_Acerto1 Text1.Text
    
    Do Until de_informa.rsSel_Acerto1.EOF
    
        de_informa.Alt_Acerto1 de_informa.rsSel_Acerto1.Fields("fretetotal"), Text1.Text, de_informa.rsSel_Acerto1.Fields("filialctc")
        
        de_informa.rsSel_Acerto1.MoveNext
    
    Loop
    
    MsgBox "OK"

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Impr As Printer, xlinha As String

    flexImpressoras.Row = 0
    flexImpressoras.ColWidth(0) = 300
    flexImpressoras.ColWidth(1) = 3100
    flexImpressoras.Text = "Impressoras Instaladas"

    'busca impressoras instaladas
    
    For Each Impr In Printers
        flexImpressoras.Row = flexImpressoras.Rows - 1
        flexImpressoras.Text = Impr.DeviceName
        flexImpressoras.Rows = flexImpressoras.Rows + 1
    Next
     
    'busca a configuração atual
    
    If Dir("C:\informa.cfg") <> "" Then
        
        Open "C:\informa.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "CTR" Then
                txtCTR = Mid$(xlinha, 5)
            ElseIf Mid$(xlinha, 1, 3) = "CTC" Then
                txtCtcs = Mid$(xlinha, 5)
            ElseIf Mid$(xlinha, 1, 3) = "MNF" Then
                txtManifesto = Mid$(xlinha, 5)
            ElseIf Mid$(xlinha, 1, 3) = "REL" Then
                txtRelat = Mid$(xlinha, 5)
            ElseIf Mid$(xlinha, 1, 3) = "FAT" Then
                txtFaturas = Mid$(xlinha, 5)
            ElseIf Mid$(xlinha, 1, 3) = "NFS" Then
                txtNFS = Mid$(xlinha, 5)
            End If
        Loop
        
        Close #1
        
    End If
    
    flexImpressoras.Rows = flexImpressoras.Rows - 1
    
End Sub

