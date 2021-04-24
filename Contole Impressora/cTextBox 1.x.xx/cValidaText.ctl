VERSION 5.00
Begin VB.UserControl cText 
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   PropertyPages   =   "cValidaText.ctx":0000
   ScaleHeight     =   345
   ScaleWidth      =   1545
   ToolboxBitmap   =   "cValidaText.ctx":000F
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1230
      Picture         =   "cValidaText.ctx":0321
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   1
      Top             =   15
      Width           =   270
   End
   Begin VB.TextBox Txt 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "cText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim CorFundo_GotFocus  As OLE_COLOR
Dim CorFundo_LostFocus As OLE_COLOR
Dim CorFundo_Erro      As OLE_COLOR
Dim CorBotao           As OLE_COLOR
Dim Mensagem_Erro      As String
Dim AssumirTextoErro   As String
Dim TipoTexto          As Byte
Dim AutoSelecao        As Boolean
Dim TrocaTAB           As Boolean
Dim AlertarErro        As Boolean
Dim MensagensPadrao    As Boolean 'mensagem de erro padrao
Dim LetrasGrandes      As Boolean
Dim sWidth             As Integer
Dim sAceitaAcentos     As Boolean
Dim sAceitaNegativo    As Boolean
Dim AceitaAspaSimples  As Boolean
Dim ExibirBotao        As Boolean
Dim ControleTrancado   As Boolean 'controlar propriedade Locked quando o texttype=6
Dim seForVazio         As String

Const VK_TAB = &H9

Enum Alinhamento
     ctLeft = 0
     ctRight = 1
     ctCenter = 2
End Enum

Enum Aparencia
     ctFlat = 0
     ct3D = 1
End Enum

Enum TipoText
        Numero = 1
        Data = 2
        Moeda = 3
        Letra = 4
        Livre = 5
        Personalizado = 6
End Enum

Enum EstiloBorda
     ctNone = 0
     ctFixed = 1
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Event Change()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Function ValidaTexto(sAscii As Integer) As Long
Select Case sAscii
Case 39 'aspa simples
        If AceitaAspaSimples = False Then
            ValidaTexto = 0
        Else
            ValidaTexto = 39
        End If
Case 32
        ValidaTexto = sAscii 'espaco
Case 65 To 90   ' A a Z maiuculo
        ValidaTexto = sAscii
Case 97 To 122 'A a Z minusculo
        ValidaTexto = sAscii
        
        If LetrasGrandes = True Then
            ValidaTexto = sAscii - 32
        End If
Case 192 To 197, 199 To 207, 210 To 214, 217 To 220, 224 To 229, 231 To 254
        If sAceitaAcentos = True Then
            ValidaTexto = sAscii
        End If
Case Else
        ValidaTexto = 0
End Select
End Function

Private Function ValidaData(sAscii As Integer, sData As String) As Long
Dim i           As Byte
Dim Ocorrencias As Byte 'quantas vezes achou "/"
Dim LocalAntigo As Byte

Select Case sAscii
Case 47 ' "/"
        For i = 1 To Len(sData)
               If InStr(i, sData, "/") > LocalAntigo Then
                    Ocorrencias = Ocorrencias + 1
                    LocalAntigo = InStr(i, sData, "/")
               End If
        Next i

        If Ocorrencias > 1 Then
            ValidaData = 0
        Else
            ValidaData = 47 'barra
        End If
Case 8, 48 To 57  'backspace e numeros
        ValidaData = sAscii
        If Txt.Selstart = 1 Or Txt.Selstart = 4 Then SendKeys ("{/}")
Case Else
        ValidaData = 0
End Select

End Function

Private Function ValidaMonetario(sAscii As Integer, sValor As String, SoNumero As Boolean) As Long
Dim Ocorrencias As Byte 'quantas vezes achou virgula
Dim LocalAntigo As Byte
Dim i           As Byte

Select Case sAscii
Case 8, 48 To 57 'backspace, virgula ou numeros
        ValidaMonetario = sAscii
Case 46 ' ponto
        If SoNumero = True Then ValidaMonetario = 0: Exit Function 'se for so numero para aceita ponto
            
        For i = 1 To Len(sValor)
               Select Case InStr(1, sValor, ",") 'pesquisa virgula
               Case 0 'nao acha virgula
                      ValidaMonetario = 46
                      Exit For
               Case Is < Txt.Selstart 'se tentar inserir o ponto apos a virgula
                      ValidaMonetario = 0
                      Exit For
               Case Is > Txt.Selstart 'se tentar inserir o ponto antes a virgula insere normal
                    ValidaMonetario = 46
                    Exit For
               End Select
        Next i
            

Case 44
        If SoNumero = True Then
            ValidaMonetario = 0
            Exit Function
        End If

        For i = 1 To Len(sValor)
               If InStr(i, sValor, ",") > LocalAntigo Then
                    Ocorrencias = Ocorrencias + 1
                    LocalAntigo = InStr(i, sValor, ",")
               End If
        Next i
        
        If Ocorrencias > 0 Then 'se achar somente uma virgula
            ValidaMonetario = 0 'valor ascii da virgula
        Else
            ValidaMonetario = 44 'anula o pressionamento
        End If
Case 45 'sinal de menos
        If sAceitaNegativo = False Then
            ValidaMonetario = 0
        Else
            If Txt.Selstart = 0 Then
                If InStr(1, Txt.Text, "-") = 0 Then ValidaMonetario = 45
            End If
        End If
Case Else
        ValidaMonetario = 0
End Select
    
End Function

Private Sub Picture1_Click()
On Error GoTo Erro
Dim sLocal  As RECT

GetWindowRect UserControl.hwnd, sLocal

Load CalendarioSDI
DoEvents

MoveWindow CalendarioSDI.hwnd, sLocal.Left, sLocal.Bottom, 184, 167, 1

CalendarioSDI.Visible = True
CalendarioSDI.ZOrder 0
Set CalendarioSDI.ControleSolicitante = Extender

Erro:
     If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, Err.Number
End Sub

Private Sub Picture1_GotFocus()
If CalendarioSDI.Visible = False Then
    keybd_event VK_TAB, 0, 0, 0
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.Line (10, 0)-(Picture1.ScaleWidth, 0)
Picture1.Line (0, 0)-(0, Picture1.ScaleHeight)

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.Cls
End Sub

Private Sub Txt_Change()
RaiseEvent Change
End Sub

Private Sub Txt_Click()
On Error GoTo Erro

If TipoTexto = 2 Then 'data
    Txt.Selstart = InStr(1, Txt.Text, "_") - 1
End If

Erro:
Select Case Err.Number
Case 0, 380 'gera um erro por imcompatibilidade entre a myMascara e os eventos que adicionei ao text
        Err.Clear
Case Else
        MsgBox "Erro Inesperado" & vbCrLf & vbCrLf & "Informações Técnicas:" & Err.Description, vbExclamation, "N° Erro: " & Err.Number
End Select
End Sub

Private Sub Txt_DblClick()
Txt_Click
End Sub

Private Sub Txt_GotFocus()
On Error GoTo Erro
Txt.BackColor = CorFundo_GotFocus

If TipoTexto = 2 Or TipoTexto = 6 Then
    If InStr(1, MyMascara, myPromptChar) > 0 Then 'se encontrar algum caractere do promptChar
        Txt.Selstart = InStr(1, MyMascara, myPromptChar) - 1
    End If
End If

If AutoSelecao = True Then
    Txt.Selstart = 0
    Txt.SelLength = Len(Txt.Text)
End If
Txt_Click

Unload CalendarioSDI
Erro:
If Err.Number <> 0 Then
    MsgBox Err.Number & "-" & Err.Description, vbInformation, "cTextBox"
End If
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Erro
Select Case TipoTexto
Case 2, 6 'data e personalizado
      sKeyDown KeyCode, Shift, Txt.Text, Txt.Selstart, Txt
End Select
    
Select Case KeyCode
Case vbKeyDown
        keybd_event VK_TAB, 0, 0, 0
Case vbKeyUp
        SendKeys ("+{TAB}")
End Select
    
RaiseEvent KeyDown(KeyCode, Shift)

Erro:
    Select Case Err.Number
    Case 5
            Err.Clear
    Case Else
            If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, Err.Number
    End Select
End Sub

Private Sub Txt_KeyPress(KeyAscii As Integer)

Select Case TipoTexto
Case 2, 6 'data e mascara
        If Txt.SelLength > 0 Then Txt.Selstart = 0
        
        If KeyAscii = vbKeyBack Then
            ValidaMascara KeyAscii, Txt.Text, Txt.Selstart, Txt
            Exit Sub
        End If
End Select

If TrocaTAB = True Then
    If KeyAscii = 13 Then
        keybd_event VK_TAB, 0, 0, 0
        KeyAscii = 0
        Exit Sub
    End If
Else
    If KeyAscii = 13 Then
       RaiseEvent KeyPress(KeyAscii)
       KeyAscii = 0
       Exit Sub
    End If
End If

Select Case TipoTexto
Case 1 'numero
        If KeyAscii = vbKeyBack Then Exit Sub
        If Txt.MaxLength > 0 Then If Len(Txt.Text) = Txt.MaxLength Then KeyAscii = 0
        KeyAscii = ValidaMonetario(KeyAscii, Txt.Text, True)
Case 2 ' data
        If ControleTrancado = False Then
            ValidaMascara KeyAscii, Txt.Text, Txt.Selstart, Txt
        End If
Case 3 'moeda
        If KeyAscii = vbKeyBack Then Exit Sub
        If Txt.MaxLength > 0 Then If Len(Txt.Text) = Txt.MaxLength Then KeyAscii = 0
        KeyAscii = ValidaMonetario(KeyAscii, Txt.Text, False)
Case 4 'letra
        If KeyAscii = vbKeyBack Then Exit Sub
        KeyAscii = ValidaTexto(KeyAscii)
Case 5 'livre (nao faz nada)
        If LetrasGrandes = True Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 39 Then If AceitaAspaSimples = False Then KeyAscii = 0
Case 6 'personalizado
        RaiseEvent KeyPress(KeyAscii)
        
        If KeyAscii = 0 Then Exit Sub
        
        If ControleTrancado = False Then
            ValidaMascara KeyAscii, Txt.Text, Txt.Selstart, Txt
        Else
            KeyAscii = 0
        End If
        Exit Sub
End Select

RaiseEvent KeyPress(KeyAscii)


End Sub

Private Sub Txt_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Txt_LostFocus()
Dim MsgPadrao As String
Dim i         As Byte

Txt.BackColor = CorFundo_LostFocus

Select Case TipoTexto
Case 1 'valor numerico
      If Txt.Text = "" Then
          Txt.Text = seForVazio
          Exit Sub
      End If

      If Not IsNumeric(Txt.Text) Then
         MsgPadrao = "Número (" & Txt.Text & ") informado inválido"
         Txt.Text = ErrorTextValue
         Txt.BackColor = CorFundo_Erro
      End If
      
Case 2 ' data
      If Txt.Text = "__/__/____" Then Exit Sub

      If Not IsDate(Txt.Text) Then
         MsgPadrao = "Data (" & Txt.Text & ") informada inválido"
         Txt.Text = "__/__/____"
         Txt.BackColor = CorFundo_Erro
         Txt.Selstart = 0
      End If
Case 3 'moeda
      If Txt.Text = "" Then
          Txt.Text = seForVazio
          Exit Sub
      End If
      
      If Not IsNumeric(Txt.Text) Then
         MsgPadrao = "Valor ( " & Txt.Text & " ) informado inválido"
         Txt.Text = ErrorTextValue
         Txt.BackColor = CorFundo_Erro
      Else
         Txt.Text = Format(Txt.Text, "Standard")
      End If
Case Else 'se for texto ou livre

End Select

If AlertarErro = True Then
    If MensagensPadrao = True Then
        If MsgPadrao <> "" Then MsgBox MsgPadrao, vbInformation, ""
    Else
        If Mensagem_Erro <> "" Then
            MsgBox Mensagem_Erro, vbInformation, ""
        End If
    End If
End If

If TipoTexto = 3 Then Txt.Text = Replace(Txt.Text, ".", "") 'se for moeda tira os pontos

End Sub

Private Sub UserControl_GotFocus()
If Txt.Enabled = True Then Txt.SetFocus

End Sub

Private Sub UserControl_Initialize()
UserControl.Refresh

End Sub

Private Sub UserControl_InitProperties()
Me.TextType = 5
Me.BackColorGotFocus = &HFFFFFF
Me.BackColorLostFocus = &HFFFFFF
Me.ErrorBackColor = &HFFFFFF
Me.ColorButton = &H8000000F

Me.Calendar_FormBackcolor = &HC0C0C0
Me.Calendar_ComboBackColor = &H80000005
Me.Calendar_BackColor = &H80000005
Me.Calendar_ColorWeekDay = &H800000
Me.Calendar_DayActive = &H80000012
Me.Calendar_DayInactive = &H808080
Me.Calendar_Selected = &HFF&

myPromptChar = "_"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'carrega
 CorFundo_GotFocus = PropBag.ReadProperty("BackColorGotFocus", &HFFFFFF)
CorFundo_LostFocus = PropBag.ReadProperty("BackColorLostFocus", &HFFFFFF)
         TipoTexto = PropBag.ReadProperty("TextType", 5)
       AutoSelecao = PropBag.ReadProperty("AutoSelect", False)
          TrocaTAB = PropBag.ReadProperty("SendTAB", False)
       AlertarErro = PropBag.ReadProperty("ErrorAlert", False)
     Mensagem_Erro = PropBag.ReadProperty("ErrorTextMessage", "")
     CorFundo_Erro = PropBag.ReadProperty("ErrorBackColor", &HFFFFFF)
  AssumirTextoErro = PropBag.ReadProperty("ErrorTextValue", "")
   MensagensPadrao = PropBag.ReadProperty("ErrorMessagesDefault", False)
     LetrasGrandes = PropBag.ReadProperty("textUcase", False)
    sAceitaAcentos = PropBag.ReadProperty("AceitaAcentos", False)
   sAceitaNegativo = PropBag.ReadProperty("AceitaNegativos", False)
 AceitaAspaSimples = PropBag.ReadProperty("AceitaAspas", False)
        seForVazio = PropBag.ReadProperty("IF_IsEmpty", "")
          CorBotao = PropBag.ReadProperty("ColorButton", &H8000000F)
      myPromptChar = PropBag.ReadProperty("PromptChar", "_")
         MyMascara = PropBag.ReadProperty("Mask", "")
       ExibirBotao = PropBag.ReadProperty("ShowButton", True)
  ControleTrancado = PropBag.ReadProperty("Locked", False)

 Calendario_FormCorFundo = PropBag.ReadProperty("Calendar_FormBackcolor", &HC0C0C0)
     Calendario_CorFundo = PropBag.ReadProperty("Calendar_BackColor", &H80000005)
Calendario_ComboCorFundo = PropBag.ReadProperty("Calendar_ComboBackColor", &H80000005)
   Calendario_DiasSemana = PropBag.ReadProperty("Calendar_ColorWeekDay", &H800000)
    Calendario_DiaAtivos = PropBag.ReadProperty("Calendar_DayActive", &H80000012)
 Calendario_DiasInativos = PropBag.ReadProperty("Calendar_DayInactive", &H808080)
  Calendario_Selecionado = PropBag.ReadProperty("Calendar_Selected", vbRed)

Txt.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
Txt.Text = PropBag.ReadProperty("Text", "")
Txt.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
Txt.MaxLength = PropBag.ReadProperty("MaxLength", 0)

Txt.ForeColor = PropBag.ReadProperty("ForeColor", 0)
Txt.Enabled = PropBag.ReadProperty("Enabled", True)
Txt.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
Txt.Appearance = PropBag.ReadProperty("Appeareance", 1)
Txt.Alignment = PropBag.ReadProperty("Alignment", 0)
Txt.FontBold = PropBag.ReadProperty("FontBold", False)
Txt.FontSize = PropBag.ReadProperty("FontSize", 8)
Txt.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")

Select Case TipoTexto
Case 2 'data
    Txt.Text = "__/__/____"
    Txt.Locked = True
Case 6 'personalizado
    Txt.Locked = True
Case Else 'outros
    Txt.Text = Me.Text
End Select

UserControl_Resize
UserControl.Refresh

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'salvando
Call PropBag.WriteProperty("Text", Txt.Text, "")
Call PropBag.WriteProperty("BackColorGotFocus", CorFundo_GotFocus, 0)
Call PropBag.WriteProperty("BackColorLostFocus", CorFundo_LostFocus, 0)
Call PropBag.WriteProperty("PasswordChar", Txt.PasswordChar, "")
Call PropBag.WriteProperty("MaxLength", Txt.MaxLength, 0)
Call PropBag.WriteProperty("Locked", ControleTrancado, False)
Call PropBag.WriteProperty("ForeColor", Txt.ForeColor, 0)
Call PropBag.WriteProperty("Enabled", Txt.Enabled, True)
Call PropBag.WriteProperty("BorderStyle", Txt.BorderStyle)
Call PropBag.WriteProperty("Appeareance", Txt.Appearance)
Call PropBag.WriteProperty("Alignment", Txt.Alignment)
Call PropBag.WriteProperty("FontBold", Txt.FontBold)
Call PropBag.WriteProperty("FontSize", Txt.FontSize)
Call PropBag.WriteProperty("FontName", Txt.FontName)
Call PropBag.WriteProperty("BackColor", Txt.BackColor)
Call PropBag.WriteProperty("TextType", TipoTexto, 5)
Call PropBag.WriteProperty("AutoSelect", AutoSelecao, False)
Call PropBag.WriteProperty("SendTAB", TrocaTAB, False)
Call PropBag.WriteProperty("ErrorAlert", AlertarErro, False)
Call PropBag.WriteProperty("ErrorTextMessage", Mensagem_Erro, "")
Call PropBag.WriteProperty("ErrorBackColor", CorFundo_Erro, &HFFFFFF)
Call PropBag.WriteProperty("ErrorTextValue", AssumirTextoErro, "")
Call PropBag.WriteProperty("ErrorMessagesDefault", MensagensPadrao, False)
Call PropBag.WriteProperty("textUcase", LetrasGrandes, False)
Call PropBag.WriteProperty("AceitaAcentos", sAceitaAcentos, False)
Call PropBag.WriteProperty("AceitaNegativos", sAceitaNegativo, False)
Call PropBag.WriteProperty("IF_IsEmpty", seForVazio, "")
Call PropBag.WriteProperty("AceitaAspas", AceitaAspaSimples, False)
Call PropBag.WriteProperty("ColorButton", CorBotao, &H8000000F)
Call PropBag.WriteProperty("PromptChar", myPromptChar, "_")
Call PropBag.WriteProperty("Mask", MyMascara, "")
Call PropBag.WriteProperty("ShowButton", ExibirBotao, True)

'calendario
Call PropBag.WriteProperty("Calendar_FormBackcolor", Calendario_FormCorFundo, &HC0C0C0)
Call PropBag.WriteProperty("Calendar_ComboBackColor", Calendario_ComboCorFundo, &H80000005)
Call PropBag.WriteProperty("Calendar_BackColor", Calendario_CorFundo, &H80000005)
Call PropBag.WriteProperty("Calendar_ColorWeekDay", Calendario_DiasSemana, &H800000)
Call PropBag.WriteProperty("Calendar_DayActive", Calendario_DiaAtivos, &H80000012)
Call PropBag.WriteProperty("Calendar_DayInactive", Calendario_DiasInativos, &H808080)
Call PropBag.WriteProperty("Calendar_Selected", Calendario_Selecionado, vbRed)


UserControl.Refresh
End Sub

Private Sub UserControl_Resize()

If TipoTexto = 2 Then 'data
    If ExibirBotao = True Then
        Picture1.Visible = ExibirBotao
        Txt.Width = (UserControl.Width - Picture1.Width) - 20
    Else
        Txt.Width = UserControl.Width - 20
    End If
    
    Txt.Height = UserControl.Height

    Picture1.Top = 0
    Picture1.Height = UserControl.Height
    Picture1.Left = Txt.Width + 20
    
Else
    Picture1.Visible = False
    Txt.Width = UserControl.Width
    Txt.Height = UserControl.Height
End If

UserControl.Refresh

End Sub

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
Text = Txt.Text
End Property

Public Property Let Text(ByVal Valor As String)
Select Case TipoTexto
Case 2 'data
      Txt.Text() = Format(Valor, "dd/mm/yyyy")
Case 3 'moeda
      Txt.Text() = Format(Valor, "Standard")
Case Else
      Txt.Text() = Valor
End Select


End Property

Public Property Get FontName() As String
FontName = Txt.FontName
End Property

Public Property Let FontName(ByVal Valor As String)
Txt.FontName() = Valor
End Property

Public Property Get FontSize() As Byte
FontSize = Txt.FontSize
End Property

Public Property Let FontSize(ByVal Valor As Byte)
Txt.FontSize() = Valor
End Property

Public Property Get FontBold() As Boolean
FontBold = Txt.FontBold
End Property

Public Property Let FontBold(ByVal Valor As Boolean)
Txt.FontBold() = Valor
End Property

Public Property Get Alignment() As Alinhamento
Alignment = Txt.Alignment
End Property

Public Property Let Alignment(ByVal Valor As Alinhamento)
Txt.Alignment() = Valor
End Property

Public Property Get Appeareance() As Aparencia
Appeareance = Txt.Appearance
End Property

Public Property Let Appeareance(ByVal Valor As Aparencia)
Txt.Appearance() = Valor
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = Txt.BackColor
End Property

Public Property Let BackColor(ByVal Valor As OLE_COLOR)
Txt.BackColor() = Valor
End Property

Public Property Get BorderStyle() As EstiloBorda
BorderStyle = Txt.BorderStyle
End Property

Public Property Let BorderStyle(ByVal Valor As EstiloBorda)
Txt.BorderStyle() = Valor
End Property

Public Property Get Enabled() As Boolean
Enabled = Txt.Enabled
UserControl.Enabled = Txt.Enabled
End Property

Public Property Let Enabled(ByVal Valor As Boolean)
Txt.Enabled() = Valor
UserControl.Enabled = Valor
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = Txt.ForeColor
End Property

Public Property Let ForeColor(ByVal Valor As OLE_COLOR)
Txt.ForeColor() = Valor
End Property

Public Property Get Locked() As Boolean
Select Case TipoTexto
Case 2, 6
        Txt.Locked = True
        Locked = ControleTrancado
Case Else
        Locked = Txt.Locked
End Select

End Property

Public Property Let Locked(ByVal Valor As Boolean)
ControleTrancado = Valor
End Property

Public Property Get MaxLength() As Integer
MaxLength = Txt.MaxLength
End Property

Public Property Let MaxLength(ByVal Valor As Integer)
If TipoTexto = 2 Or TipoTexto = 6 Then MsgBox "Acesso Negado !", vbExclamation, "": Exit Property
Txt.MaxLength = Valor
End Property

Public Property Get PasswordChar() As String
PasswordChar = Txt.PasswordChar
End Property

Public Property Let PasswordChar(ByVal Valor As String)
Txt.PasswordChar() = Valor
End Property

Public Property Get TextType() As TipoText
TextType = TipoTexto
UserControl_Resize
End Property

Public Property Let TextType(ByVal Valor As TipoText)
TipoTexto = Valor

Select Case Valor
Case 1, 3 'numero moeda
    Me.Alignment = ctRight
    Txt.Text = ""
Case 2 'data
    Me.Alignment = ctCenter
    UserControl.Height = 310
    Me.Mask = "__/__/____"
    Me.Text = "__/__/____"
    Txt.Text = "__/__/____"
    ExibirBotao = False
    
Case 4, 5 'texto
    Me.Alignment = ctLeft
    Txt.Text = ""
Case 6 'personalizado
    Txt.Text = MyMascara
End Select

UserControl.Refresh
End Property

Public Property Get BackColorGotFocus() As OLE_COLOR
BackColorGotFocus = CorFundo_GotFocus
End Property

Public Property Let BackColorGotFocus(ByVal Valor As OLE_COLOR)
CorFundo_GotFocus = Valor
End Property

Public Property Get BackColorLostFocus() As OLE_COLOR
BackColorLostFocus = CorFundo_LostFocus
End Property

Public Property Let BackColorLostFocus(ByVal Valor As OLE_COLOR)
CorFundo_LostFocus = Valor
End Property

Private Sub UserControl_Terminate()
Unload CalendarioSDI

End Sub

Public Property Get AutoSelect() As Boolean
Attribute AutoSelect.VB_Description = "Habilita a AutoSelecao do texto quando a Caixa ganhar o Foco"
AutoSelect = AutoSelecao
End Property

Public Property Let AutoSelect(ByVal Valor As Boolean)
AutoSelecao = Valor
End Property

Public Property Get SendTAB() As Boolean
Attribute SendTAB.VB_Description = "Pressiona um TAB quando e pressionado o ENTER"
SendTAB = TrocaTAB
End Property

Public Property Let SendTAB(ByVal Valor As Boolean)
TrocaTAB = Valor
End Property

Public Property Get ErrorAlert() As Boolean
Attribute ErrorAlert.VB_Description = "Exibir alerta que o valor digitado esta incorreto"
ErrorAlert = AlertarErro
End Property

Public Property Let ErrorAlert(ByVal Valor As Boolean)
AlertarErro = Valor
End Property

Public Property Get ErrorBackColor() As OLE_COLOR
Attribute ErrorBackColor.VB_Description = "Muda a Cor de fundo quando o valor informado nao e aceito"
ErrorBackColor = CorFundo_Erro
End Property

Public Property Let ErrorBackColor(ByVal Valor As OLE_COLOR)
CorFundo_Erro = Valor
End Property

Public Property Get ErrorTextMessage() As String
Attribute ErrorTextMessage.VB_Description = "Mostra o texto informado nesta propriedade quando houver erro"
ErrorTextMessage = Mensagem_Erro
End Property

Public Property Let ErrorTextMessage(ByVal Valor As String)
Mensagem_Erro = Valor
End Property

Public Property Get ErrorTextValue() As String
Attribute ErrorTextValue.VB_Description = "A caixa de texto assume o valor informado nesta propriedade quando houver erro"
ErrorTextValue = AssumirTextoErro
End Property

Public Property Let ErrorTextValue(ByVal Valor As String)
AssumirTextoErro = Valor
End Property

Public Property Get ErrorMessagesDefault() As Boolean
Attribute ErrorMessagesDefault.VB_Description = "Habilita para que o componente use as mensagens de erro padrao do Componente"
ErrorMessagesDefault = MensagensPadrao
End Property

Public Property Let ErrorMessagesDefault(ByVal Valor As Boolean)
MensagensPadrao = Valor
End Property

Public Property Get textUcase() As Boolean
Attribute textUcase.VB_Description = "Se for TRUE a caixa sodigita caracteres em maiusculo"
textUcase = LetrasGrandes
End Property

Public Property Let textUcase(ByVal Valor As Boolean)
LetrasGrandes = Valor
End Property

Public Property Get AceitaAcentos() As Boolean
AceitaAcentos = sAceitaAcentos
End Property

Public Property Let AceitaAcentos(ByVal Valor As Boolean)
sAceitaAcentos = Valor
End Property

Public Property Get Calendar_FormBackcolor() As OLE_COLOR
Calendar_FormBackcolor = Calendario_FormCorFundo
End Property

Public Property Let Calendar_FormBackcolor(ByVal Valor As OLE_COLOR)
Calendario_FormCorFundo = Valor
End Property

Public Property Get Calendar_ComboBackColor() As OLE_COLOR
Calendar_ComboBackColor = Calendario_ComboCorFundo
End Property

Public Property Let Calendar_ComboBackColor(ByVal Valor As OLE_COLOR)
Calendario_ComboCorFundo = Valor
End Property

Public Property Get Calendar_BackColor() As OLE_COLOR
Calendar_BackColor = Calendario_CorFundo
End Property

Public Property Let Calendar_BackColor(ByVal Valor As OLE_COLOR)
Calendario_CorFundo = Valor
End Property

Public Property Get Calendar_DayActive() As OLE_COLOR
Calendar_DayActive = Calendario_DiaAtivos
End Property

Public Property Let Calendar_DayActive(ByVal Valor As OLE_COLOR)
Calendario_DiaAtivos = Valor
End Property

Public Property Get Calendar_DayInactive() As OLE_COLOR
Calendar_DayInactive = Calendario_DiasInativos
End Property

Public Property Let Calendar_DayInactive(ByVal Valor As OLE_COLOR)
Calendario_DiasInativos = Valor

End Property

Public Property Get Calendar_ColorWeekDay() As OLE_COLOR
Calendar_ColorWeekDay = Calendario_DiasSemana
End Property

Public Property Let Calendar_ColorWeekDay(ByVal Valor As OLE_COLOR)
Calendario_DiasSemana = Valor

End Property

Public Property Get Calendar_Selected() As OLE_COLOR
Attribute Calendar_Selected.VB_Description = "Muda Cor do Circulo que seleciona as datas"
Calendar_Selected = Calendario_Selecionado
End Property

Public Property Let Calendar_Selected(ByVal Valor As OLE_COLOR)
Calendario_Selecionado = Valor
End Property

Public Property Get AceitaNegativos() As Boolean
Attribute AceitaNegativos.VB_Description = "Habilita o uso o sinal de nagativo quando a propriedade TextType for Numero ou Moeda"
AceitaNegativos = sAceitaNegativo
End Property

Public Property Let AceitaNegativos(ByVal Valor As Boolean)
sAceitaNegativo = Valor
End Property

Public Property Get IF_IsEmpty() As String
Attribute IF_IsEmpty.VB_Description = "Assume o texto informado nesta propriedade quando a caixa de texto perde o foco vazia"
IF_IsEmpty = seForVazio
End Property

Public Property Let IF_IsEmpty(ByVal Valor As String)
seForVazio = Valor
End Property

Public Property Get AceitaAspas() As Boolean
Attribute AceitaAspas.VB_Description = "Para Evitar o Uso de Aspas Simples no cTextBox"
AceitaAspas = AceitaAspaSimples
End Property

Public Property Let AceitaAspas(ByVal Valor As Boolean)
AceitaAspaSimples = Valor
End Property

Public Property Get ColorButton() As OLE_COLOR
ColorButton = CorBotao '&H8000000F&
'Picture1.BackColor = CorBotao (Depois fazer o triangulo via codigo)
End Property

Public Property Let ColorButton(ByVal Valor As OLE_COLOR)
'MsgBox "Propriedade Inoperante no momento", vbExclamation, ""
CorBotao = Valor
End Property

Public Property Get PromptChar() As String
PromptChar = myPromptChar
End Property

Public Property Let PromptChar(ByVal Valor As String)
If Len(Valor) > 1 Then MsgBox "Parâmetro Incorreto" & vbCrLf & vbCrLf & "Informe apenas 01 digito", vbExclamation, "": Exit Property

myPromptChar = Valor
End Property

Public Property Get Mask() As String
Mask = MyMascara
End Property

Public Property Let Mask(ByVal Valor As String)
Select Case TipoTexto
Case 1, 3, 4, 5 'data e personalizado
    MyMascara = ""
End Select

If InStr(1, Valor, myPromptChar) = 0 Then
    MsgBox "Nenhum caractere do PromptChar foi informado", vbExclamation, "Cancelado !"
    Exit Property
End If

MyMascara = Valor
Me.Text = Valor
Txt.MaxLength = Len(Valor)
Txt.Text = Valor
End Property

Public Property Get ShowButton() As Boolean
ShowButton = ExibirBotao
End Property

Public Property Let ShowButton(ByVal Valor As Boolean)
If TipoTexto <> 2 Then MsgBox "Mude a propriedade TextType para (2-Data)", vbExclamation, "Cancelado !": Exit Property
Picture1.Visible = Valor
ExibirBotao = Valor
End Property

Public Property Get Value() As String
Dim Str As String
Dim i   As Integer

If MyMascara = "" Then
    Value = Txt.Text
Else
    For i = 1 To Len(Txt.Text)
            If Mid(MyMascara, i, 1) = PromptChar Then
                Str = Str & Mid(Txt.Text, i, 1)
            End If
    Next i
    
    Value = Str
End If

End Property

Public Property Get Selstart() As Integer
Attribute Selstart.VB_MemberFlags = "400"
Selstart = Txt.Selstart
End Property

Public Property Let Selstart(ByVal Valor As Integer)
Txt.Selstart = Valor
End Property
