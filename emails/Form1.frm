VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Compor E-Mail"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   1890
      TabIndex        =   12
      Top             =   5520
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin TabDlg.SSTab sst1 
      Height          =   3015
      Left            =   0
      TabIndex        =   11
      Top             =   2790
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5318
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Normal"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtBody"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Visualização"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "wb1"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtBody 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   7335
      End
      Begin SHDocVwCtl.WebBrowser wb1 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   14
         Top             =   120
         Width           =   7335
         ExtentX         =   12938
         ExtentY         =   4895
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   1482
      BandCount       =   1
      _CBWidth        =   7560
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "tblCommands"
      MinHeight1      =   780
      Width1          =   2655
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tblCommands 
         Height          =   780
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   1376
         ButtonWidth     =   1879
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "imlDisabled"
         DisabledImageList=   "imlDisabled"
         HotImageList    =   "imlEnabled"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Enviar"
               Object.ToolTipText     =   "Envia o Email"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Prioridade"
               Object.ToolTipText     =   "Seleciona um nível de prioridade"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Alta"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Normal"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Baixa"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Acusar leitura"
               Object.ToolTipText     =   "Leia a ajuda para maiores informações."
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Sim"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Não"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2025
      Left            =   0
      TabIndex        =   5
      Top             =   750
      Width           =   7575
      Begin VB.TextBox txtAssunto 
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   1605
         Width           =   6390
      End
      Begin VB.TextBox txtAnexo 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1290
         Width           =   6375
      End
      Begin MSComctlLib.ImageList iml1 
         Left            =   105
         Top             =   1965
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   255
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":047A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":07CE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   4800
         TabIndex        =   1
         Text            =   "smtp.servidor.com.br"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Text            =   "seu@email.com.br"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtTo1 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   615
         Width           =   6375
      End
      Begin VB.TextBox txtTo2 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   945
         Width           =   6375
      End
      Begin MSComctlLib.Toolbar cmdTo 
         Height          =   1320
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   2328
         ButtonWidth     =   1588
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "iml1"
         DisabledImageList=   "iml1"
         HotImageList    =   "iml1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "De:"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Para:"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "CPara:"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anexo:"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Assunto:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "Servidor SMTP"
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Servidor:"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         ToolTipText     =   "Servidor SMTP"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "De:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList imlDisabled 
      Left            =   0
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B22
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1776
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlEnabled 
      Left            =   720
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":301E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":48C6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Prioridade As MAIL_PRIORITY
Private Recebimento As Boolean
Private WithEvents SendMail As clsSendMail
Attribute SendMail.VB_VarHelpID = -1

Private Sub cmdTo_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.index
    Case 1
        '
        'MsgBox "Método não implementado", vbInformation + vbOKOnly, "De"
        frmServers.Show 1
    Case 2
        'MsgBox "Método não implementado", vbInformation + vbOKOnly, "Para"
        frmEmailList.Show 1
    Case 3
        'MsgBox "Método não implementado", vbInformation + vbOKOnly, "Cópia Para"
        frmEmailList.Show 1
    Case 4
        txtAnexo.Text = GetFile
End Select
End Sub

Private Sub Form_Load()
Prioridade = NORMAL_PRIORITY
'pBar.Value = 100
wb1.Offline = True
wb1.Navigate "about:blank"
Set SendMail = New clsSendMail

If GetSetting("VBSender", "Geral", "Exibir", True) Then
    Load frmLogin
    frmLogin.Show vbModal
End If

Debug.Print Me.Caption & " Initialize=True"
End Sub

Private Sub Form_Resize()
Frame1.Left = 0
Frame1.Width = Me.ScaleWidth
'lef tool 120
'left label 240
Dim mW As Long, mH As Long
If Me.WindowState = vbMinimized Then Exit Sub
mW = (((Me.ScaleWidth - Label1.Width) / 2) - (240 * 3))
Label1.Left = 240
txtEmail.Move Label1.Left + Label1.Width, Label1.Top, mW
Label2.Left = txtEmail.Left + txtEmail.Width
Label2.Top = Label1.Top
txtServer.Move Label2.Left + Label2.Width, Label2.Top, mW
mW = (txtServer.Left + txtServer.Width) - txtEmail.Left
cmdTo.Left = 120
txtTo1.Move txtEmail.Left, txtTo1.Top, mW
txtTo2.Move txtEmail.Left, txtTo2.Top, mW
txtAnexo.Move txtEmail.Left, txtAnexo.Top, mW
Label3.Left = 240
txtAssunto.Move txtEmail.Left, txtAssunto.Top, mW
mW = Me.ScaleWidth
mH = ((Me.ScaleHeight) - (Frame1.Top + Frame1.Height + 10))
sst1.Move 0, Frame1.Top + Frame1.Height + 10, mW, mH
pBar.Move 1890, Me.ScaleHeight - 255, Me.ScaleWidth - 1890, 255
txtBody.Move 120, 120, sst1.Width - 240, sst1.Height - (240 * 2.5)
wb1.Offline = True
wb1.Move 120, 120, sst1.Width - 240, sst1.Height - (240 * 2.5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim I As clsFolders
    Set I = New clsFolders
    If I.Fexist(I.InAppPath("temp.htm")) Then Kill I.InAppPath("temp.htm")
    Set I = Nothing
    Set SendMail = Nothing
    'Load frmDetails
    'Load frmEmailList
    'Load frmSck
    'Load Form1
    'Load frmServers
    'Unload frmDetails
    'Unload frmEmailList
    'Unload frmSck
    'Unload Form1
    'Unload frmServers
'End
Debug.Print Me.Caption & " Terminate = True"
End Sub

Private Sub SendMail_Progress(PercentComplete As Long)
pBar.Min = 0
pBar.Max = 100
pBar.Value = PercentComplete
End Sub

Private Sub SendMail_SendFailed(Explanation As String)
MsgBox "Seu email não foi enviado", vbInformation + vbOKOnly, "NÃO Enviado"
frmDetails.List1.AddItem "Erro ao enviar o e-mail: " & Explanation
End Sub

Private Sub SendMail_SendSuccesful()
MsgBox "Seu email foi enviado com successo", vbInformation + vbOKOnly, "Enviado"
frmDetails.List1.AddItem "E-mail enviado com sucesso."
End Sub

Private Sub SendMail_Status(Status As String)
frmDetails.List1.AddItem Status
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
Dim fFile As Integer
wb1.Offline = True
Select Case sst1.Tab
    Case 0
        txtBody.Visible = True
        wb1.Visible = False
    Case 1
        'wb1.Document.innerHTML = txtBody.Text
        fFile = FreeFile
        Open App.Path & "\temp.htm" For Output As #fFile
            Print #fFile, txtBody.Text
        Close #fFile
        wb1.Navigate2 App.Path & "\temp.htm"
        txtBody.Visible = False
        wb1.Visible = True
End Select
End Sub

Private Sub tblCommands_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Caption
    Case "Prioridade"
        MsgBox strPrioridade(Prioridade), vbInformation + vbOKOnly, "Prioridade"
    Case "Acusar leitura"
        MsgBox strRecebimento(Recebimento), vbInformation + vbOKOnly, "Acusar Leitura"
    Case Else
        Send
End Select
End Sub

Private Sub tblCommands_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Text
    Case "Alta"
        Prioridade = HIGH_PRIORITY
    Case "Normal"
        Prioridade = NORMAL_PRIORITY
    Case "Baixa"
        Prioridade = LOW_PRIORITY
    Case "Sim"
        Recebimento = True
    Case "Não"
        Recebimento = False
End Select
End Sub

Private Function strPrioridade(ByVal prior As MAIL_PRIORITY) As String
Dim s As String
Select Case prior
    Case 1
        s = "Alta"
    Case 3
        s = "Normal"
    Case 5
        s = "Baixa"
    Case Else
        s = "Normal"
End Select
strPrioridade = s
End Function

Private Function strRecebimento(ByVal rec As Boolean) As String
    strRecebimento = IIf(rec, "Sim", "Não")
End Function
Private Function GetFile() As String
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    'Set the parent window
    OFName.hwndOwner = Me.hWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Select a filter
    OFName.lpstrFilter = "Todos os Arquivos (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the title
    OFName.lpstrTitle = "Anexar Arquivo"
    'No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        GetFile = Trim$(OFName.lpstrFile)
    Else
        GetFile = ""
    End If
End Function

Private Sub Send()
    Screen.MousePointer = vbHourglass
    Load frmDetails
    frmDetails.Show
    With SendMail
        .SMTPHostValidation = VALIDATE_NONE
        .EmailAddressValidation = VALIDATE_SYNTAX
        .Delimiter = ";"
        .SMTPHost = txtServer.Text
        
        'Aqui entra o código para identificação.
        If GetSetting("VBSender", "Geral", "Exigir Autenticação", False) Then
            .UseAuthentication = True
            If GetSetting("VBSender", "Geral", "Salvo", True) Then
                .Username = GetSetting("VBSender", "Geral", "User")
                .Password = GetSetting("VBSender", "Geral", "Pass")
            Else
                Screen.MousePointer = vbDefault
                Load frmLogIn2
                frmLogIn2.Show vbModal
                Screen.MousePointer = vbHourglass
                .Username = frmLogIn2.txtName.Text
                .Password = frmLogIn2.txtPassword.Text
                Unload frmLogIn2
            End If
        End If
        .from = txtEmail.Text
        .FromDisplayName = txtEmail.Text
        .Recipient = txtTo1.Text
        .RecipientDisplayName = txtTo1.Text
        .CcRecipient = txtTo2.Text
        .CcDisplayName = txtTo2.Text
        .ReplyToAddress = txtEmail.Text
        .Subject = txtAssunto.Text
        .Message = txtBody.Text
        .Attachment = Trim(txtAnexo.Text)
        .AsHTML = True
        .ContentBase = ""
        .EncodeType = MIME_ENCODE
        .Priority = Prioridade
        .Receipt = Recebimento
        .UseAuthentication = False
        .MaxRecipients = 100
        .Send
    End With
    Screen.MousePointer = vbDefault
End Sub
