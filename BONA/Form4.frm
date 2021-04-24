VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   9180
   ClientLeft      =   3225
   ClientTop       =   1290
   ClientWidth     =   11475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   9180
   ScaleWidth      =   11475
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   7200
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexImpressoras 
      Height          =   2295
      Left            =   8160
      TabIndex        =   12
      Top             =   4800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4048
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   5880
      TabIndex        =   11
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   615
      Left            =   5880
      TabIndex        =   10
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   5880
      TabIndex        =   9
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   4920
      Width           =   1935
   End
   Begin VB.ListBox List3 
      Height          =   3960
      Left            =   2640
      TabIndex        =   7
      Top             =   4680
      Width           =   2535
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   8655
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2760
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim xlin As Integer
Dim xtamanhoarq As Integer
Dim xlinhas As String

Open Label1.Caption For Input As #1

Do Until EOF(1)
xlin = xlin + 1
Line Input #1, xlinhas

    List1.AddItem xlinhas
    
    Call OrdenaLista(Form4.List1, List1.ListCount, xlinhas)


Loop


Close #1


End Sub

Private Sub Command2_Click()
Dim xteste As String

xteste = InputBox("Teste", "Testando", "")

List1.AddItem xteste

Call OrdenaLista(Form4.List1, List1.ListCount, xteste)


End Sub

Private Sub Command3_Click()
Dim x As Integer
Dim y As Integer

Randomize

For x = 1 To 10
    y = Int(Rnd * 10)
    List2.AddItem y

Next



End Sub

Private Sub Command4_Click()
Dim x As Integer
Dim z As String

For x = 0 To List2.ListCount - 1
    z = List2.List(x)
    List3.AddItem z
    
    Call OrdenaLista(Form4.List3, List3.ListCount, z)
    
Next


End Sub

Private Sub Command5_Click()
Call OrdenaListagem(Form4.List2)
End Sub

Private Sub Command6_Click()
Dim ximp As Printer
flexImpressoras.ColWidth(0) = 200
flexImpressoras.ColWidth(1) = 2500


    For Each Impr In Printers
        
        
        flexImpressoras.Row = flexImpressoras.Rows - 1
        flexImpressoras.Text = Impr.DeviceName
        flexImpressoras.Rows = flexImpressoras.Rows + 1
    Next








End Sub

Private Sub Command7_Click()



Open flexImpressoras.TextMatrix(flexImpressoras.Row, flexImpressoras.Col) For Output As #1

    Printer.Print "TESTE"
    Printer.EndDoc

Close #1



End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()

If Dir1.Path <> "C:\" Then
    Label1.Caption = File1.Path & "\" & File1.FileName
Else
    Label1.Caption = File1.Path & File1.FileName
End If

End Sub

Private Sub Form_Load()
Dir1.Path = "C:\"
File1.Path = Dir1.Path

End Sub

Private Sub Form_KeyDown(KeyAsc As Integer, Shift As Integer)

If KeyAsc = 27 Then
    Unload Me
End If


End Sub

