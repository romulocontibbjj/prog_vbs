VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmColetaImpressoras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração de Impressoras"
   ClientHeight    =   3615
   ClientLeft      =   3495
   ClientTop       =   3150
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmColetaImpressoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCtcs 
      Caption         =   "<="
      Height          =   285
      Left            =   4080
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   420
      Width           =   4575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexImpressoras 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
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
      TabIndex        =   3
      Top             =   3180
      Width           =   2235
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3180
      Width           =   2235
   End
   Begin VB.TextBox txtCOL 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Impressora Default:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmColetaImpressoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCtcs_Click()
txtCOL.Text = flexImpressoras.TextMatrix(flexImpressoras.RowSel, flexImpressoras.ColSel)
End Sub

Private Sub cmdGravar_Click()
    Open App.Path & "\coletaimp.cfg" For Output As #1
    Print #1, "COL=" & txtCOL.Text
    Close #1
    MsgBox "Impressoras Gravadas !", vbInformation, ""
Unload Me
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
    
    If Dir(App.Path & "\coletaimp.cfg") <> "" Then
        
        Open App.Path & "\coletaimp.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "COL" Then
                txtCOL = Mid$(xlinha, 5)
            End If
        Loop
        
        Close #1
        
    End If
    
    flexImpressoras.Rows = flexImpressoras.Rows - 1
    
End Sub

