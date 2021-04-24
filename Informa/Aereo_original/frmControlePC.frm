VERSION 5.00
Begin VB.Form frmControlePC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração de PC"
   ClientHeight    =   975
   ClientLeft      =   3090
   ClientTop       =   3765
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmControlePC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   2235
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   2235
   End
   Begin VB.TextBox txtAWBs 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nome deste PC"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "frmControlePC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCtr_Click()
    txtCTR.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdCtcs_Click()
txtAWBs.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub CmdETL_Click()
TxtETL.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub CmdETV_Click()
TxtETV.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdGravar_Click()
    Open App.Path & "\PC.cfg" For Output As #1
    Print #1, "PC=" & txtAWBs
    Close #1
    MsgBox "Nomedo PC Gravado!", vbInformation, ""
Unload Me
End Sub

Private Sub cmdManif_Click()
txtManifesto.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdRelat_Click()
    txtRelat.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    'busca a configuração atual
    
    If Dir(App.Path & "\PC.cfg") <> "" Then
        
        Open App.Path & "\PC.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "PC" Then
                txtAWBs = Mid$(xlinha, 5)
            End If
        Loop
        
        Close #1
        
    End If
End Sub

