VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_Pagamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SOLICITÇÃO DE PAGAMENTOS"
   ClientHeight    =   8235
   ClientLeft      =   2310
   ClientTop       =   1260
   ClientWidth     =   10260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10260
   Begin VB.Frame Frame1 
      Caption         =   "PAGAMENTOS"
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&Sair"
         Height          =   255
         Left            =   8280
         TabIndex        =   32
         Top             =   7680
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mask_valor 
         Height          =   300
         Left            =   8520
         TabIndex        =   7
         Top             =   2640
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "9.999,99"
         PromptChar      =   "_"
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   8160
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txt_departamento 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   3
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmd_imprimi 
         Caption         =   "&Imprimir"
         Height          =   255
         Left            =   8280
         TabIndex        =   13
         Top             =   7440
         Width           =   1455
      End
      Begin VB.TextBox txt_ramal2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8400
         TabIndex        =   12
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox txt_autorização 
         Height          =   285
         Left            =   8400
         MaxLength       =   10
         TabIndex        =   11
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox txt_ramal 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   10
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox txt_coneferente 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   9
         Top             =   6360
         Width           =   1815
      End
      Begin VB.TextBox txt_obs 
         Height          =   2295
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3840
         Width           =   8415
      End
      Begin MSMask.MaskEdBox mask_vencimento 
         Height          =   300
         Left            =   8520
         TabIndex        =   6
         Top             =   2280
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mask_recebimento 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   1920
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_fatura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8520
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1920
         Width           =   1000
      End
      Begin VB.TextBox txt_fornecedor 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   2
         Top             =   2280
         Width           =   4095
      End
      Begin VB.ComboBox cmd_doc 
         Height          =   315
         ItemData        =   "frm_Pagamentos.frx":0000
         Left            =   1200
         List            =   "frm_Pagamentos.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lab_cont_obs 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "(             )"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Valor R$:"
         Height          =   255
         Left            =   7560
         TabIndex        =   29
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Departamento:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Ramal.........:"
         Height          =   255
         Left            =   7320
         TabIndex        =   26
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Autorização:"
         Height          =   255
         Left            =   7320
         TabIndex        =   25
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Ramal........:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   6720
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Conferente:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   6360
         Width           =   855
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   10200
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   10080
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label8 
         Caption         =   "Descrição.....:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Vencimento:"
         Height          =   255
         Left            =   7560
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lab_cont 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "(MAX. 30 -          )"
         Height          =   255
         Left            =   5400
         TabIndex        =   19
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Fatura Nº:"
         Height          =   255
         Left            =   7560
         TabIndex        =   18
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo Doc......:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Fornecedor...:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Recebimento:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.Line Line1 
         DrawMode        =   14  'Copy Pen
         X1              =   120
         X2              =   10080
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "RELATÓRIO PARA CONEFERÊNCIA DE PAGAMENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   720
         Width           =   4695
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   240
         Picture         =   "frm_Pagamentos.frx":002B
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frm_Pagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_imprimi_Click()
Dim Excel As Excel.Application
Dim ExcelA1 As Excel.Worksheet
Dim ExcelWBk As Excel.Workbook
Dim xarq As String
Dim xarqnum As String
Dim xarqnovo As String

xarq = File1.Path & "\" & File1.List(0)
xarqnum = Mid(xarq, Len(xarq) - 4, 1)

If xarqnum = 9 Then
    
    xarqnovo = Mid(xarq, 1, Len(xarq) - 5) & "0"
    
ElseIf xarqnum >= 0 Then
    
     xarqnovo = UCase(Mid(xarq, 1, Len(xarq) - 5) & (xarqnum + 1))
     
End If



Set Excel = CreateObject("Excel.Application")
Set Excel = GetObject(, "Excel.Application")
Excel.Visible = True
Excel.Interactive = True
Excel.Workbooks.Open FileName:=xarq

Set ExcelA1 = Excel.Worksheets(1)

ExcelA1.Cells(5, 2) = CDate(mask_recebimento)
ExcelA1.Cells(6, 2) = txt_fornecedor.Text
ExcelA1.Cells(7, 2) = cmd_doc.Text
ExcelA1.Cells(8, 2) = txt_departamento
ExcelA1.Cells(15, 2) = txt_coneferente.Text
ExcelA1.Cells(5, 5) = txt_fatura.Text
ExcelA1.Cells(6, 5) = CDate(mask_vencimento)
ExcelA1.Cells(10, 1) = Mid(txt_obs.Text, 1, 60)
ExcelA1.Cells(11, 1) = Mid(txt_obs.Text, 61, 60)
ExcelA1.Cells(12, 1) = Mid(txt_obs.Text, 121, 60)
ExcelA1.Cells(13, 1) = Mid(txt_obs.Text, 181, 60)
ExcelA1.Cells(14, 1) = Mid(txt_obs.Text, 241, 60)
ExcelA1.Cells(15, 5) = txt_autorização.Text
ExcelA1.Cells(16, 5) = txt_ramal2.Text
ExcelA1.Cells(15, 2) = txt_coneferente.Text
ExcelA1.Cells(16, 2) = txt_ramal.Text
ExcelA1.Cells(7, 5) = mask_valor


ExcelA1.PrintOut

'ExcelWBk.SaveAs xarqnovo, , , , , , xlExclusive



Excel.Quit
Set ExcelA1 = Nothing
Set Excel = Nothing

'Kill xarq




End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub Form_Load()

mask_recebimento.SelStart = 0
mask_recebimento.SelLength = Len(mask_recebimento)

File1.Path = "C:\PAGAMENTOS"


End Sub

Private Sub mask_vencimento_Change()

mask_vencimento.SelStart = 0
mask_vencimento.SelLength = Len(mask_vencimento)


End Sub

Private Sub txt_fornecedor_Change()
lab_cont.Caption = Len(txt_fornecedor)

End Sub

Private Sub txt_obs_Change()
lab_cont_obs.Caption = Len(txt_obs.Text)



End Sub
