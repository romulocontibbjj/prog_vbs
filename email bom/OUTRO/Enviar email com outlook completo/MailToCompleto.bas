Attribute VB_Name = "MailToCompleto"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum Prioridade
    Alta = 1
    Normal = 3
    Baixa = 5
End Enum
Public Function GeraMensagem(NomeDoArquivoEML As String, Remetente As String, Destinatario As String, Prioridade As Prioridade, _
                            Optional CopiaPara As String, Optional CopiaOcultaPara As String, _
                            Optional Assunto, Optional Mensagem, Optional Anexo As String)
    Screen.MousePointer = 11
    Open NomeDoArquivoEML For Output As #2 ''inicia a geração EML
    Print #2, "From: <"; Remetente; ">" ''REMETENTE
    Print #2, "To: <"; Destinatario; ">" ''DESTINO
    Print #2, "Cc: <"; CopiaPara; ">" ''COPIA PARA
    Print #2, "Bcc: <"; CopiaOcultaPara; ">" ''COPIA OCULTA PARA
    Print #2, "Subject: "; Assunto ''ASSUNTO
    Print #2, "MIME-Version: 1.0"
    Print #2, "X-Priority:"; Prioridade ''PRIORIDADE
    Select Case Prioridade
           Case 1
                Print #2, "X-MSMail-Priority: Hight" ''PRIORIDADE ALTA
           Case 3
                Print #2, "X-MSMail-Priority: Normal" ''PRIORIDADE NORMAL
           Case 5
                Print #2, "X-MSMail-Priority: Low" ''PRIORIDADE BAIXA
    End Select
    Print #2, "X-Unsent: 1"
    Print #2, "Content-Type: multipart/mixed;"
    Print #2, " boundary="; Chr(34); "----=_NextPart_000_0006_01C0294E.2CB018E0"; Chr(34)
    Print #2, "X-MimeOLE: Programado por http://pagina.de/luciochaves"
    Print #2, ""
    Print #2, "This is a multi-part message in MIME format."
    Print #2, ""
    Print #2, "------=_NextPart_000_0006_01C0294E.2CB018E0"
    Print #2, "Content-Type: text/plain;"
    Print #2, " charset="; Chr(34); "iso-8859-1"; Chr(34)
    Print #2, "Content-Transfer-Encoding: 7bit"
    Print #2, ""
    '''IMPRIME O TEXTO DA MENSAGEM
    Print #2, Mensagem
    '''ENCERRA A IMPRESSAO DA MENSAGEM
    Print #2, ""
    If Anexo <> "" Then
        Print #2, "------=_NextPart_000_0006_01C0294E.2CB018E0"
        Print #2, "Content-Type: text/plain;"
        Print #2, " name="; Chr(34); Dir(Anexo); Chr(34) ''NOME DO ARQUIVO ANEXO
        Print #2, "Content-Transfer-Encoding: 7bit"
        Print #2, "Content-Disposition: attachment;"
        Print #2, " filename="; Chr(34); Dir(Anexo); Chr(34) ''NOME DO ARQUIVO ANEXO
        Print #2, ""
        
        '''INICIA A IMPRESSAO DO CONTEUDO DO ARQUIVO ANEXO
            Open Anexo For Binary Shared As #1
            Do
                buf = Input(2048000, 1)
                Print #2, buf
            Loop Until EOF(1)
            Close #1
        ''''ENCERRA A IMPRESSAO DO ARQUIVO ANEXO
    End If
    Print #2, ""
    Print #2, "------=_NextPart_000_0006_01C0294E.2CB018E0--"
    Close #2
    Screen.MousePointer = 0
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


