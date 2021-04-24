VERSION 5.00
Begin VB.Form frmCadTabPrecoLocalidades 
   Caption         =   "Adicionar ou Remover Localidades"
   ClientHeight    =   5235
   ClientLeft      =   2535
   ClientTop       =   1590
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "frmCadTabPrecoLocalidades.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   6375
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   4800
      Width           =   3075
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Frame FraLocalidades 
      Caption         =   "Localidades de Atendimento"
      Height          =   4635
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6255
      Begin VB.CommandButton CmdRemoveLocalidade 
         Caption         =   "REMOVER"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   2940
         TabIndex        =   1
         Top             =   2820
         Width           =   315
      End
      Begin VB.CommandButton CmdAdicionaLocalidade 
         Caption         =   "ADI C IONAR"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   2940
         TabIndex        =   6
         Top             =   900
         Width           =   315
      End
      Begin VB.CommandButton CmdLocalidades 
         Caption         =   "Cadastro de Localidades"
         Height          =   375
         Left            =   3180
         TabIndex        =   5
         Top             =   300
         Width           =   2895
      End
      Begin VB.CommandButton CmdTodasLocalidades 
         Caption         =   "Todas as Localidades"
         Height          =   375
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   2895
      End
      Begin VB.ListBox ListLocalidadesSel 
         Height          =   3570
         Left            =   3300
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   900
         Width           =   2775
      End
      Begin VB.ListBox ListLocalidadesDisponives 
         Height          =   3570
         Left            =   180
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   900
         Width           =   2715
      End
   End
End
Attribute VB_Name = "frmCadTabPrecoLocalidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdicionaLocalidade_Click()
CmdAdicionaLocalidade.Enabled = False
Call TransfereItemDeListBox(ListLocalidadesDisponives, ListLocalidadesSel)
Call OrdenaListBox(ListLocalidadesSel)
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdContinuar_Click()
Dim Y, A, B, C As Integer
    
    If ListLocalidadesSel.ListCount = 0 Then
    MsgBox "Não existe nenhuma Localidade selecionada. Por favor, tente novamente...", vbExclamation, ""
    Exit Sub
    End If
    A = frmCadTabPrecoALTERACAO_UNIT.FlexGridAjustar.Rows
    For Y = 0 To ListLocalidadesSel.ListCount - 1
    frmCadTabPrecoALTERACAO_UNIT.FlexGridAjustar.
    
    
End Sub

Private Sub CmdLocalidades_Click()
    frmCadLocalidade.Show 1

If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
de_informa.Sel_CadLocalAir "%"

    Do Until de_informa.rsSel_CadLocalAir.EOF
    ListLocalidadesDisponives.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade")) & " - " & de_informa.rsSel_CadLocalAir.Fields("SIGLA")
    de_informa.rsSel_CadLocalAir.MoveNext
    Loop
    
    'INICIO DO TRECHO QUE AVERIGUA LIST BOX
    For Y = 0 To ListLocalidadesDisponives.ListCount - 1
    ListLocalidadesDisponives.Selected(Y) = False
    Next
    For X = 0 To ListLocalidadesSel.ListCount - 1
        For Y = 0 To ListLocalidadesDisponives.ListCount - 1
            If ListLocalidadesDisponives.List(Y) = ListLocalidadesSel.List(X) Then
            ListLocalidadesDisponives.Selected(Y) = True
            End If
        Next
    Next
    Y = 0
    Do While True
        If Y > ListLocalidadesDisponives.ListCount - 1 Then
        Exit Sub
        End If
        If ListLocalidadesDisponives.Selected(Y) = True Then
        ListLocalidadesDisponives.RemoveItem (Y)
            If Y > 0 Then
            Y = Y - 1
            Else
            Y = 0
            End If
        Else
        Y = Y + 1
        End If
    Loop
    'FIM DO TRECHO QUE AVERIGUA LIST BOX
    
End Sub

Private Sub CmdRemoveLocalidade_Click()
CmdRemoveLocalidade.Enabled = False
Call TransfereItemDeListBox(ListLocalidadesSel, ListLocalidadesDisponives)
Call OrdenaListBox(ListLocalidadesDisponives)
End Sub

Private Sub CmdTodasLocalidades_Click()

CmdTodasLocalidades.Enabled = False
Dim xCont As Integer

    For xCont = 0 To ListLocalidadesDisponives.ListCount - 1
    ListLocalidadesDisponives.Selected(xCont) = True
    Next
    
Call TransfereItemDeListBox(ListLocalidadesDisponives, ListLocalidadesSel)
Call OrdenaListBox(ListLocalidadesSel)

CmdAdicionaLocalidade.Enabled = False
CmdRemoveLocalidade.Enabled = False

CmdTodasLocalidades.Enabled = True
End Sub

Private Sub Form_Load()
Dim X, Y As Integer
    
If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
de_informa.Sel_CadLocalAir "%"

    Do Until de_informa.rsSel_CadLocalAir.EOF
    ListLocalidadesDisponives.AddItem PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade")) & " - " & de_informa.rsSel_CadLocalAir.Fields("SIGLA")
    de_informa.rsSel_CadLocalAir.MoveNext
    Loop

    
    
    For X = 1 To frmCadTabPrecoALTERACAO_UNIT.FlexGridAjustar.Rows - 1
    ListLocalidadesSel.AddItem Trim(frmCadTabPrecoALTERACAO_UNIT.FlexGridAjustar.TextMatrix(X, 0)) & " - " & Trim(frmCadTabPrecoALTERACAO_UNIT.FlexGridAjustar.TextMatrix(X, 1))
    Next
    
    'INICIO DO TRECHO QUE AVERIGUA LIST BOX
    For Y = 0 To ListLocalidadesDisponives.ListCount - 1
    ListLocalidadesDisponives.Selected(Y) = False
    Next
    For X = 0 To ListLocalidadesSel.ListCount - 1
        For Y = 0 To ListLocalidadesDisponives.ListCount - 1
            If ListLocalidadesDisponives.List(Y) = ListLocalidadesSel.List(X) Then
            ListLocalidadesDisponives.Selected(Y) = True
            End If
        Next
    Next
    Y = 0
    Do While True
        If Y > ListLocalidadesDisponives.ListCount - 1 Then
        Exit Sub
        End If
        If ListLocalidadesDisponives.Selected(Y) = True Then
        ListLocalidadesDisponives.RemoveItem (Y)
            If Y > 0 Then
            Y = Y - 1
            Else
            Y = 0
            End If
        Else
        Y = Y + 1
        End If
    Loop
    'FIM DO TRECHO QUE AVERIGUA LIST BOX
    
CmdAdicionaLocalidade.Enabled = False
End Sub

Private Sub ListLocalidadesDisponives_Click()
CmdAdicionaLocalidade.Enabled = True
End Sub

Private Sub ListLocalidadesSel_Click()
CmdRemoveLocalidade.Enabled = True
End Sub
