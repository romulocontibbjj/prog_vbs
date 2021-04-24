VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoFiltraClienteDest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Especifique seu Cliente"
   ClientHeight    =   2955
   ClientLeft      =   225
   ClientTop       =   2385
   ClientWidth     =   11595
   Icon            =   "frmEmissaoFiltraClienteDest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "< ESC > : Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   2640
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "< ENTER > : Seleciona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9540
      TabIndex        =   1
      Top             =   2640
      Width           =   1995
   End
End
Attribute VB_Name = "frmEmissaoFiltraClienteDest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Flex_KeyPress(KeyAscii As Integer)
Dim x As Integer

x = Flex.Row

    If KeyAscii = 27 Then
    Unload Me
    ElseIf KeyAscii = 13 Then
    frmEmissao.TxtCGCDestinatario.Text = Flex.TextMatrix(x, 0)
    frmEmissao.TxtNomeDestinatario.Text = Flex.TextMatrix(x, 1)
    frmEmissao.TxtCidadeDestinatario.Text = Flex.TextMatrix(x, 2)
    frmEmissao.TxtUFDestinatario.Text = Flex.TextMatrix(x, 3)
    Unload Me
    End If
        
End Sub

Private Sub Form_Load()
Dim x As Integer
Flex.Clear
Flex.Cols = 5
Flex.Rows = de_informa.rsSelCliente.RecordCount + 1
Flex.FixedRows = 1
Flex.FixedCols = 0
Flex.TextMatrix(0, 0) = "CNPJ"
Flex.TextMatrix(0, 1) = "Nome"
Flex.TextMatrix(0, 2) = "Cidade"
Flex.TextMatrix(0, 3) = "UF"
Flex.TextMatrix(0, 4) = "Endereço"
Flex.ColWidth(0) = 1500
Flex.ColWidth(1) = 4000
Flex.ColWidth(2) = 3000
Flex.ColWidth(3) = 500
Flex.ColWidth(4) = 6000


de_informa.rsSelCliente.MoveFirst
x = 1
    Do Until de_informa.rsSelCliente.EOF
    Flex.TextMatrix(x, 0) = PriMaiuscula(de_informa.rsSelCliente.Fields("cgc"))
    Flex.TextMatrix(x, 1) = PriMaiuscula(de_informa.rsSelCliente.Fields("nome"))
    Flex.TextMatrix(x, 2) = PriMaiuscula(de_informa.rsSelCliente.Fields("cidade"))
    Flex.TextMatrix(x, 3) = de_informa.rsSelCliente.Fields("uf")
    Flex.TextMatrix(x, 4) = PriMaiuscula(de_informa.rsSelCliente.Fields("endereco"))
    x = x + 1
    de_informa.rsSelCliente.MoveNext
    Loop
    
End Sub

