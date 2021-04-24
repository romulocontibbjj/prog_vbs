VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmManifestoFiltraMotorista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Especifique o Motorista"
   ClientHeight    =   5715
   ClientLeft      =   3075
   ClientTop       =   1335
   ClientWidth     =   6735
   Icon            =   "frmManifestoFiltraMotorista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   5235
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9234
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
      Top             =   5460
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
      Left            =   4680
      TabIndex        =   1
      Top             =   5460
      Width           =   1995
   End
End
Attribute VB_Name = "frmManifestoFiltraMotorista"
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
    frmManifesto.TxtMotorista.Text = Flex.TextMatrix(x, 0)
    Unload Me
    End If
        
End Sub

Private Sub Form_Load()
Dim x As Integer
Flex.Clear
Flex.Cols = 2
Flex.Rows = de_informa.rsMotorista.RecordCount + 1
Flex.FixedRows = 1
Flex.FixedCols = 0
Flex.TextMatrix(0, 0) = "Nome"
Flex.TextMatrix(0, 1) = "Função"
Flex.ColWidth(0) = 4000
Flex.ColWidth(1) = 2000

de_informa.rsMotorista.MoveFirst
x = 1
    Do Until de_informa.rsMotorista.EOF
    Flex.TextMatrix(x, 0) = PriMaiuscula(de_informa.rsMotorista.Fields("nome"))
    Flex.TextMatrix(x, 1) = PriMaiuscula(de_informa.rsMotorista.Fields("funcao"))
    x = x + 1
    de_informa.rsMotorista.MoveNext
    Loop
End Sub

