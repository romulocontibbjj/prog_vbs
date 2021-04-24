VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoBuscaCiaAerea 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3075
   ClientLeft      =   2895
   ClientTop       =   3015
   ClientWidth     =   6615
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoBuscaCiaAerea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid FlexGridCiaAerea 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4683
      _Version        =   393216
      SelectionMode   =   1
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
      Left            =   4440
      TabIndex        =   2
      Top             =   2820
      Width           =   1995
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
      Left            =   120
      TabIndex        =   1
      Top             =   2820
      Width           =   1635
   End
End
Attribute VB_Name = "frmEmissaoBuscaCiaAerea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlexGridCiaAerea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    Unload Me
    ElseIf KeyAscii = 13 Then
    frmEmissao.TxtSiglaCiaAerea.Text = FlexGridCiaAerea.TextMatrix(FlexGridCiaAerea.Row, 0)
    frmEmissao.TxtNomeCiaAerea.Caption = FlexGridCiaAerea.TextMatrix(FlexGridCiaAerea.Row, 1)
    frmEmissao.TxtCGCCiaAerea.Caption = FlexGridCiaAerea.TextMatrix(FlexGridCiaAerea.Row, 2)
    frmEmissao.TxtInscrEstCiaAerea.Caption = FlexGridCiaAerea.TextMatrix(FlexGridCiaAerea.Row, 3)
    Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim xCont As Integer
FlexGridCiaAerea.Clear

If de_informa.rsSelCiaAerea.State = 1 Then de_informa.rsSelCiaAerea.Close
de_informa.SelCiaAerea "%"

FlexGridCiaAerea.Rows = de_informa.rsSelCiaAerea.RecordCount + 1
FlexGridCiaAerea.Cols = 4

FlexGridCiaAerea.FixedCols = 0
FlexGridCiaAerea.FixedRows = 1

    With FlexGridCiaAerea
    .TextMatrix(0, 0) = "Sigla"
    .TextMatrix(0, 1) = "Nome Cia."
    .TextMatrix(0, 2) = "CNPJ"
    .TextMatrix(0, 3) = "Inscr. Est."
    .ColWidth(0) = 500
    .ColWidth(1) = 3100
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    End With
    
    xCont = 1
    Do Until de_informa.rsSelCiaAerea.EOF
        With FlexGridCiaAerea
        .TextMatrix(xCont, 0) = de_informa.rsSelCiaAerea.Fields("codcia")
        .TextMatrix(xCont, 1) = PriMaiuscula(de_informa.rsSelCiaAerea.Fields("fantasia"))
        .TextMatrix(xCont, 2) = de_informa.rsSelCiaAerea.Fields("cgc")
        .TextMatrix(xCont, 3) = de_informa.rsSelCiaAerea.Fields("inscrest")
        End With
    xCont = xCont + 1
    de_informa.rsSelCiaAerea.MoveNext
    Loop


End Sub
