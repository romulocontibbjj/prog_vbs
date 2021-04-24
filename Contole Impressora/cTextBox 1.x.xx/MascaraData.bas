Attribute VB_Name = "Mascaras"
Option Explicit

Public MyMascara      As String
Public myPromptChar As String

Public Calendario_CorFundo      As OLE_COLOR
Public Calendario_FormCorFundo  As OLE_COLOR
Public Calendario_ComboCorFundo As OLE_COLOR
Public Calendario_DiaAtivos     As OLE_COLOR
Public Calendario_DiasInativos  As OLE_COLOR
Public Calendario_DiasSemana    As OLE_COLOR
Public Calendario_Selecionado   As OLE_COLOR 'item selecionado(shape)

Public Function sKeyDown(KeyCode As Integer, Shift As Integer, TextoAtual As String, sSelStart As Integer, sTextBox As TextBox)
Dim TextoFrente As String
Dim TextoAtras  As String

Select Case KeyCode
Case 8 'BackSpace
        If sSelStart = 0 Then Exit Function

        If Mid(MyMascara, sSelStart, 1) = myPromptChar Then
             TextoAtras = Mid(TextoAtual, sSelStart + 1)
            TextoFrente = Mid(TextoAtual, 1, sSelStart - 1)

            sTextBox.Text = TextoFrente & myPromptChar & TextoAtras
            sTextBox.SelStart = sSelStart - 1
        
            If Mid(MyMascara, sSelStart - 1, 1) <> myPromptChar Then sTextBox.SelStart = sSelStart - 2
            
        End If
Case 46 'Delete
        If sSelStart = 0 Then sTextBox.SelStart = sSelStart + 1
        
        If Mid(MyMascara, sSelStart + 1, 1) = myPromptChar Then
             TextoFrente = Mid(TextoAtual, sSelStart + 2)
              TextoAtras = Mid(TextoAtual, 1, sSelStart)

            sTextBox.Text = TextoAtras & myPromptChar & TextoFrente
            sTextBox.SelStart = sSelStart + 1
        
            If Mid(MyMascara, sSelStart + 2, 1) <> myPromptChar Then sTextBox.SelStart = sSelStart + 2
        End If
End Select
End Function
Public Function ValidaMascara(sKeyAscii As Integer, TextoAtual As String, sSelStart As Integer, sTextBox As TextBox)
Dim TextoAtras  As String
Dim LetraFrente As String 'mostra a letra da frente

Select Case sKeyAscii
Case 48 To 57, 65 To 90, 97 To 122
    If Mid(MyMascara, sSelStart + 1, 1) = myPromptChar Then 'Encontrou o myPromptChar
        TextoAtras = Mid(TextoAtual, 1, sSelStart)
        LetraFrente = Mid(TextoAtual, sSelStart + 1, 1)
    
        sTextBox.Text = TextoAtras & Replace(TextoAtual, LetraFrente, Chr(sKeyAscii), sSelStart + 1, 1)
    
    
        sTextBox.SelStart = sSelStart + 1
    
        If Mid(MyMascara, sSelStart + 2, 1) <> myPromptChar Then sTextBox.SelStart = sSelStart + 2
    Else
        sTextBox.SelStart = sSelStart + 1
    End If
End Select
End Function
