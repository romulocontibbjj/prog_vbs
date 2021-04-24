VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_horarios 
   Caption         =   "HORÁRIO DE SAÍDA - EU"
   ClientHeight    =   4920
   ClientLeft      =   5385
   ClientTop       =   3090
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   4380
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   840
      Top             =   4440
   End
   Begin VB.ListBox List3 
      Height          =   1815
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   3000
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   1680
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin MSMask.MaskEdBox mask_hora 
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   550
         _ExtentX        =   979
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmd_hora 
         Caption         =   "MARCAR PTO"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmd_fui 
         Caption         =   "FUI"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   600
         Top             =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Horário de Saída:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LAB_HORA 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "HORA:"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lab_data 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "DATA:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox MediaPlayer1 
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   3915
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frm_horarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_fui_Click()
Unload Me

End Sub

Private Sub cmd_hora_Click()


deb_hora.in_hora lab_data, LAB_HORA, mask_hora.Text

If mask_hora < Mid(Time, 1, 5) Then

    MsgBox "Hora Extra Cadastrada", vbInformation, "HORA EXTRA"

Else
    
    MsgBox "Hora Cadastrada", vbInformation, "HORA NORMAL"

End If


End Sub

Private Sub Form_Load()
Dim xdata As String
Dim xtime As String
Dim xnumdia As Integer
Dim xnomedia As String

mask_hora.SelStart = 0
mask_hora.SelLength = Len(mask_hora)



lab_data.Caption = Date
LAB_HORA.Caption = Time


With deb_hora.rssel_td_hora
.Open

If .RecordCount > 0 Then

.MoveFirst

Do Until .EOF

    xdata = .Fields("DATA")
    xtime = .Fields("HORA")
    
        xnumdia = Weekday(xdata)
        xnomedia = WeekdayName(xnumdia)
        List3.AddItem xnomedia
        List1.AddItem xdata
        List2.AddItem xtime
        
      
    .MoveNext
    
Loop
    
End If

    
End With






End Sub

Private Sub List3_Click()
List2.ListIndex = List3.ListIndex
List1.ListIndex = List3.ListIndex
End Sub

Private Sub Timer1_Timer()

LAB_HORA.Caption = Time


End Sub

