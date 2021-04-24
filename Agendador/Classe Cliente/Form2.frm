VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adiciona Tarefa"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancela"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   12
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.OptionButton opTipo 
      Caption         =   "Mensal"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.OptionButton opTipo 
      Caption         =   "Semanal"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton opTipo 
      Caption         =   "Diario"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton opTipo 
      Caption         =   "Manual"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "Nome"
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   120
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Começa em"
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label4 
      Caption         =   "Minuto"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Hora"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Dia"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivo"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me
    
End Sub

Private Sub Command2_Click()

    Dim T As Integer
    
    With Form1
        .Ag.Arquivo = Text1
        
        For i = 0 To opTipo.UBound
            If opTipo(i).Value = True Then
                .Ag.Frequencia = i + 1
            End If
        Next
        
        If Combo1.Enabled = False Then
            T = 0
        Else
            T = Combo1.ListIndex + 1
        End If
        
        .Ag.Dia = T
        .Ag.Hora = Text3
        .Ag.Minuto = Text4
        .Ag.StartDate = Text2
        .Ag.Save Text5
    End With

    Form1.LoadList
    Unload Me
    
End Sub

Private Sub opTipo_Click(Index As Integer)

    Combo1.Clear
    
    If Index < 2 Then
        Combo1.Enabled = False
    Else
        Combo1.Enabled = True
        
        If Index = 2 Then
            Combo1.AddItem "Domingo"
            Combo1.AddItem "Segunda"
            Combo1.AddItem "Terça"
            Combo1.AddItem "Quarta"
            Combo1.AddItem "Quinta"
            Combo1.AddItem "Sexta"
            Combo1.AddItem "Sábado"
        Else
            For i = 1 To 31
                Combo1.AddItem i
            Next
        End If
    End If
    
End Sub
