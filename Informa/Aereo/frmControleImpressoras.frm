VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmControleImpressoras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração de Impressoras"
   ClientHeight    =   4710
   ClientLeft      =   2400
   ClientTop       =   3195
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmControleImpressoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtETV 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   1140
      Width           =   2415
   End
   Begin VB.CommandButton CmdETV 
      Caption         =   "<="
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   1140
      Width           =   555
   End
   Begin VB.TextBox TxtETL 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   780
      Width           =   2415
   End
   Begin VB.CommandButton CmdETL 
      Caption         =   "<="
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   780
      Width           =   555
   End
   Begin VB.CommandButton cmdManif 
      Caption         =   "<="
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Top             =   420
      Width           =   555
   End
   Begin VB.CommandButton cmdCtcs 
      Caption         =   "<="
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   60
      Width           =   555
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
      Left            =   60
      TabIndex        =   6
      Top             =   1500
      Width           =   4575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexImpressoras 
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4048
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   4260
      Width           =   2235
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4260
      Width           =   2235
   End
   Begin VB.TextBox txtAWBs 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   60
      Width           =   2415
   End
   Begin VB.TextBox txtManifesto 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   420
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Etiquetas de Vols.:"
      Height          =   195
      Left            =   195
      TabIndex        =   15
      Top             =   1185
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Etiquetas de Lote:"
      Height          =   195
      Left            =   225
      TabIndex        =   12
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "AWBs (Matricial):"
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Manifesto (Matricial):"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   465
      Width           =   1455
   End
End
Attribute VB_Name = "frmControleImpressoras"
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
    'Open "c:\printer.cfg" For Output As #1
    Open "c:\printer.cfg" For Output As #1
    Print #1, "AWB=" & txtAWBs
    Print #1, "MNF=" & txtManifesto
    Print #1, "ETL=" & TxtETL
    Print #1, "ETV=" & TxtETV
    Close #1
    MsgBox "Impressoras Gravadas !", vbInformation, ""
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
    
    'If Dir("c:\printer.cfg") <> "" Then
    
    If Dir("c:\printer.cfg") <> "" Then
        
        'Open "c:\printer.cfg" For Input As #1
                      
        Open "c:\printer.cfg" For Input As #1
     
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "AWB" Then
                txtAWBs = Mid$(xlinha, 5)
            ElseIf Mid$(xlinha, 1, 3) = "MNF" Then
                txtManifesto = Mid$(xlinha, 5)
            ElseIf Mid$(xlinha, 1, 3) = "ETV" Then
            TxtETV = Mid$(xlinha, 5)
            ElseIf Mid$(xlinha, 1, 3) = "ETL" Then
            TxtETL = Mid$(xlinha, 5)
            End If
        Loop
               
        
        Close #1
    End If
    
    flexImpressoras.Rows = flexImpressoras.Rows - 1
    
End Sub

