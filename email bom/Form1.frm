VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   1500
   ClientTop       =   1875
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "sair"
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "envia"
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Text            =   "romulo@intec.com.br"
      Top             =   720
      Width           =   1695
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3720
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()


MAPISession1.Action = 1
MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.Compose

'DESTINATÁRIO
MAPIMessages1.RecipAddress = Text1.Text
MAPIMessages1.AddressResolveUI = True
MAPIMessages1.ResolveName

'TITULO
MAPIMessages1.MsgSubject = UCase(Text2.Text)
'MENSAGEM
MAPIMessages1.MsgNoteText = UCase(Text3.Text) & Chr$(13) & "ROMULO CONTI"
MAPIMessages1.Send False

'anexa no final da mensagem
MAPIMessages1.AttachmentPosition = Len(MAPIMessages1.MsgNoteText)
'define o tipo de dados do anexo
MAPIMessages1.AttachmentType = mapData
'da um nome ao anexo
MAPIMessages1.AttachmentName = "GLAXO.xls"
'define o caminho e nome do arquivo a anexar
MAPIMessages1.AttachmentPathName = "C:\Glaxo.xls"



'http://www.freecode.com.br/drArtigos/art_detalhe.asp?s=432947450UUXQ65123QIWFFM
'SITE



MAPISession1.SignOff






'Private Sub Command1_Click()
'MAPISession1.SignOn

'MAPIMessages1.SessionID = MAPISession1.SessionID

'MAPIMessages1.Compose
'MAPIMessages1.RecipAddress = Text1.Text
'MAPIMessages1.MsgSubject = Text2.Text
'MAPIMessages1.MsgNoteText = Text3.Text

'anexa no final da mensagem
'MAPIMessages1.AttachmentPosition = Len(MAPIMessages1.MsgNoteText)
'define o tipo de dados do anexo
'MAPIMessages1.AttachmentType = mapData
'da um nome ao anexo
'MAPIMessages1.AttachmentName = "Anexos"
'define o caminho e nome do arquivo a anexar
'MAPIMessages1.AttachmentPathName = Text4.Text

'envia o arquivo
'MAPIMessages1.Send True

'MAPISession1.SignOff

'End Sub

'Como funciona:

'- Apenas definimos as propriedades AttachmentPosition, AttachmentType , AttachmentName e AttachmentPathName do controle MAPIMessages

'Deletando e repassando mensagens

'Para encerrar nosso assunto veremos agora como podemos deletar e repassar mensagens recebidas usando o MAPI.

'Para excluir uma mensagem basta usar o método Delete , assim: MAPI.Messages1.Delete

'Para repassar uma mensagem usamos o método Forward : MAPI.Messages1.Forward

'O controle MAPIMessages contém ainda outras propriedades para que possamos gerenciar mensagens de e-mail; embora elas não tenha sido aqui abordadas , abaixo temos uma relação com a descrição resumida de cada uma delas:

'Propriedade Descrição

'MsgConversationID = Determina o valor do identificador da conversação para a mensagem atual.
'MsgCount = Retorna o numero total de mensagens presentes no conjunto de mensagens recebidas para a sessão atual.
'MsgDateReceived = Retorna a data na qual a mensagem foi recebida.
'MsgID = Retona uma string que identifica a indice da mensagem atual.
'MsgIndex = Determina o número do indice da mensagem atual.
'MsgOrigAddress = Retorna o endereço de email do remetente da mensagem atual.
'MsgOrigDisplayName = Retorna o nome original para a mensagem atual.
'MsgRead = Retorna uma expressão Boleana indicando se a mensagem ja foi lida.
'MsgSent = Determina se a mensagem atual já foi enviada ao servidor de email para distribuição.
'MsgType = Define o tipo da mensagem atual.







End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If


End Sub
