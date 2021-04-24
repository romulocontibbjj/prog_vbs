VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Frm_Diversos 
   Caption         =   "Diversos"
   ClientHeight    =   4035
   ClientLeft      =   2805
   ClientTop       =   3045
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5085
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   970
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Fontes"
      TabPicture(0)   =   "Frm_Diversos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fra_estilo"
      Tab(0).Control(1)=   "Fra_Cor"
      Tab(0).Control(2)=   "fra_Fontes"
      Tab(0).Control(3)=   "Lab_fontes"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Objetos"
      TabPicture(1)   =   "Frm_Diversos.frx":0452
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lst_obj"
      Tab(1).Control(1)=   "cmd_limpar"
      Tab(1).Control(2)=   "cmd_remove"
      Tab(1).Control(3)=   "cmd_ok"
      Tab(1).Control(4)=   "txt_item"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Videos"
      TabPicture(2)   =   "Frm_Diversos.frx":08A4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fil_videos"
      Tab(2).Control(1)=   "mp_videos"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Fotos"
      TabPicture(3)   =   "Frm_Diversos.frx":0CF6
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "IMG_LIST"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lst_fotos"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "tmr_fotos"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "chk_automatico"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.FileListBox fil_videos 
         Height          =   2820
         Left            =   -74880
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chk_automatico 
         Caption         =   "Automático"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Timer tmr_fotos 
         Enabled         =   0   'False
         Interval        =   700
         Left            =   120
         Top             =   3120
      End
      Begin VB.ListBox lst_fotos 
         BackColor       =   &H00C0FFFF&
         Height          =   2010
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1695
      End
      Begin VB.ListBox lst_obj 
         Height          =   2400
         Left            =   -72120
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmd_limpar 
         Caption         =   "Limpar"
         Height          =   795
         Left            =   -73200
         Picture         =   "Frm_Diversos.frx":0D12
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmd_remove 
         Caption         =   "Remover"
         Height          =   915
         Left            =   -73200
         Picture         =   "Frm_Diversos.frx":1154
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmd_ok 
         Caption         =   "OK"
         Height          =   735
         Left            =   -73200
         Picture         =   "Frm_Diversos.frx":1596
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txt_item 
         Height          =   285
         Left            =   -74760
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Fra_estilo 
         Caption         =   "Estilo"
         Height          =   1815
         Left            =   -71880
         TabIndex        =   4
         Top             =   1755
         Width           =   1335
         Begin VB.CheckBox Chk_Tachado 
            Caption         =   "Tachado"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox Chk_Italico 
            Caption         =   "Itálico"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox Chk_Sublinhado 
            Caption         =   "Sublinhado"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chk_Negrito 
            Caption         =   "Negrito"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Fra_Cor 
         Caption         =   "Cor"
         Height          =   1815
         Left            =   -73320
         TabIndex        =   3
         Top             =   1755
         Width           =   1335
         Begin VB.OptionButton Opt_Azul 
            Caption         =   "Azul"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Opt_Vinho 
            Caption         =   "Vinho"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1095
         End
         Begin VB.OptionButton Opt_Vermelho 
            Caption         =   "Vermelho"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Opt_Verde 
            Caption         =   "Verde"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fra_Fontes 
         Caption         =   "Fontes"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   2
         Top             =   1755
         Width           =   1335
         Begin VB.OptionButton Opt_Script 
            Caption         =   "Script"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   1095
         End
         Begin VB.OptionButton Opt_Currier 
            Caption         =   "Currier"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Opt_Garamond 
            Caption         =   "Garamond"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton opt_Comic 
            Caption         =   "Comic"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
      End
      Begin MediaPlayerCtl.MediaPlayer mp_videos 
         Height          =   2895
         Left            =   -73200
         TabIndex        =   24
         Top             =   720
         Width           =   2775
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
      Begin VB.Image IMG_LIST 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Lab_fontes 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Visual Basic Essentials"
         Height          =   255
         Left            =   -74760
         TabIndex        =   1
         Top             =   1395
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Frm_Diversos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer


Private Sub chk_automatico_Click()
tmr_fotos.Enabled = Not tmr_fotos.Enabled


End Sub

Private Sub Chk_Italico_Click()
If Chk_Italico.Value = 1 Then
Lab_fontes.FontItalic = True
Else
Lab_fontes.FontItalic = False
End If


End Sub

Private Sub chk_Negrito_Click()
If chk_Negrito.Value = 1 Then
Lab_fontes.FontBold = True
Else
Lab_fontes.FontBold = False
End If

End Sub

Private Sub Chk_Sublinhado_Click()
If Chk_Sublinhado.Value = 1 Then
Lab_fontes.FontUnderline = True
Else
Lab_fontes.FontUnderline = False
End If

End Sub

Private Sub Chk_Tachado_Click()
If Chk_Tachado.Value = 1 Then
Lab_fontes.FontStrikethru = True
Else
Lab_fontes.FontStrikethru = False
End If

End Sub

Private Sub cmd_limpar_Click()
lst_obj.Clear
End Sub

Private Sub cmd_ok_Click()
Dim i As Integer
Dim x As Integer
Dim y As Integer

If Trim(txt_item.Text) = Empty Then
MsgBox "Não Se pode cadastrar Dados em branco", vbInformation, "Não é Possível"
txt_item.SetFocus
Else
x = lst_obj.ListCount
Do While x <> -1
If Trim(txt_item.Text) = lst_obj.List(x) Then
MsgBox "NAO"
txt_item.SelStart = 0
txt_item.SelLength = Len(Trim(txt_item.Text))
txt_item.SetFocus
Exit Sub
Else
x = x - 1
End If

Loop

lst_obj.AddItem txt_item.Text
txt_item.Text = Empty
txt_item.SetFocus


End If



End Sub

Private Sub cmd_remove_Click()
If lst_obj.ListCount = 0 Then
MsgBox "LISTA VAZIA", vbCritical, "ERRO"
ElseIf lst_obj.ListIndex = -1 Then
MsgBox "Marque o Item", vbCritical, "ERRO"
Else
lst_obj.RemoveItem lst_obj.ListIndex
End If


End Sub

Private Sub fil_videos_Click()
mp_videos.FileName = "C:\VIDEOS\" & fil_videos.FileName

End Sub

Private Sub Form_Load()
Dim vfoto As String
'ABRINDO ARQ_FOTOS.TXT
Open "C:\FOTOS\ARQ_FOTOS.TXT" For Input As #1
Do While EOF(1) = False
    Line Input #1, vfoto 'lendo a linha
        lst_fotos.AddItem Trim(vfoto) 'Adicionando a lista
Loop

Close #1 'Fechando arquivo

'_______________________________________
fil_videos.Path = "C:\VIDEOS"





End Sub

Private Sub img_List_Click()
Form1.Show
End Sub

Private Sub lst_fotos_Click()
IMG_LIST.Picture = LoadPicture("C:\FOTOS\" & lst_fotos.Text)
End Sub

Private Sub MediaPlayer1_DVDNotify(ByVal EventCode As Long, ByVal EventParam1 As Long, ByVal EventParam2 As Long)

End Sub

Private Sub Opt_Azul_Click()
Lab_fontes.ForeColor = vbBlue

End Sub

Private Sub opt_Comic_Click(Index As Integer)
    Lab_fontes.FontName = "Comic Sans MS"
    
End Sub

Private Sub Opt_Currier_Click(Index As Integer)
Lab_fontes.FontName = "Courier"

End Sub

Private Sub Opt_Garamond_Click(Index As Integer)
 Lab_fontes.FontName = "Garamond"
 
End Sub

Private Sub Opt_Script_Click(Index As Integer)
Lab_fontes.FontName = "Script"

End Sub

Private Sub Opt_Verde_Click()
Lab_fontes.ForeColor = vbGreen

End Sub

Private Sub Opt_Vermelho_Click()
Lab_fontes.ForeColor = vbRed

End Sub

Private Sub Opt_Vinho_Click()
Lab_fontes.ForeColor = &H80&

End Sub

Private Sub tmr_fotos_Timer()
Static vcont As Integer
IMG_LIST.Picture = LoadPicture("C:\FOTOS\" & lst_fotos.List(vcont))

lst_fotos.Selected(vcont) = True

vcont = vcont + 1

If vcont = lst_fotos.ListCount Then vcont = 0
    





End Sub
