VERSION 5.00
Begin VB.Form frmReimprProt 
   Caption         =   "Reimpressão de Protocolo para Arquivo (CTCs)"
   ClientHeight    =   1515
   ClientLeft      =   3255
   ClientTop       =   2175
   ClientWidth     =   5970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   5970
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtNumProt 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número do Protocolo:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1545
   End
End
Attribute VB_Name = "frmReimprProt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
Dim xcopias As Single, xcontctc As Integer, xcoluna As Single, xlinha As Single
If IsNumeric(txtNumProt) Then
    If de_informa.rsSel_ReRelProt.State = 1 Then de_informa.rsSel_ReRelProt.Close
    de_informa.Sel_ReRelProt Val(txtNumProt), xusuario
    If de_informa.rsSel_ReRelProt.RecordCount < 1 Then
        MsgBox "Não há Protocolo Com este Número e com Este Usuário"
        Exit Sub
    Else
        For xcopias = 1 To 1
            de_informa.rsSel_ReRelProt.MoveFirst
            xcontctc = 0
            xcoluna = 1
            xlinha = 0
            Do Until de_informa.rsSel_ReRelProt.EOF
                xcontctc = xcontctc + 1   'contador de quantidade
                If xlinha = 0 And xcoluna = 1 Then   'identifica inicio da página/cabeçário
                    Printer.Print
                    Printer.Print
                    Printer.FontSize = 12
                    Printer.FontBold = True
                    Printer.FontUnderline = True
                    Printer.Print Spc(5); "INTEC TRANSPORTES"
                    Printer.FontUnderline = False
                    Printer.Print
                    Printer.Print Spc(5); "RELATÓRIO DE CTCs FÍSICOS BAIXADOS"
                    Printer.Print Spc(5); "DOCUMENTOS DESPACHADOS PARA SETOR DE ARQUIVO EM " & de_informa.rsSel_ReRelProt.Fields("rel_arq_data")
                    Printer.Print Spc(5); "USUÁRIO / DIGITADOR: " & xusuario
                    Printer.Print Spc(5); "PROTOCOLO NÚMERO: " & String(6 - Len(Trim$(txtNumProt)), "0") & Trim$(txtNumProt)
                    Printer.FontStrikethru = True
                    Printer.Print Spc(5); String(132, " ")
                    Printer.FontSize = 10
                    Printer.FontStrikethru = False
                    Printer.FontBold = False
                    Printer.FontUnderline = False
                End If
                'impressão por 6 colunas
                If xcoluna = 1 Then
                    Printer.Print Spc(6); Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 1, 2) & "-" & _
                    Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 3, 8);
                    xcoluna = xcoluna + 1
                    xlinha = xlinha + 1
                ElseIf xcoluna = 2 Then
                    Printer.Print Spc(3); Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 1, 2) & "-" & _
                    Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 3, 8);
                    xcoluna = xcoluna + 1
                ElseIf xcoluna = 3 Then
                    Printer.Print Spc(3); Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 1, 2) & "-" & _
                    Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 3, 8);
                    xcoluna = xcoluna + 1
                ElseIf xcoluna = 4 Then
                    Printer.Print Spc(3); Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 1, 2) & "-" & _
                    Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 3, 8);
                    xcoluna = xcoluna + 1
                ElseIf xcoluna = 5 Then
                    Printer.Print Spc(3); Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 1, 2) & "-" & _
                    Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 3, 8);
                    xcoluna = xcoluna + 1
                ElseIf xcoluna = 6 Then
                    Printer.Print Spc(3); Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 1, 2) & "-" & _
                    Mid(de_informa.rsSel_ReRelProt.Fields("filialctc"), 3, 8);
                    'na coluna 6 volta para a coluna 1
                    xcoluna = 1
                    Printer.Print
                    'e se a linha for = 26 e não for último CTC ...
                    If xlinha = 26 And de_informa.rsSel_ReRelProt.RecordCount <> xcontctc Then
                        xlinha = 0
                        Printer.FontSize = 12
                        Printer.FontBold = True
                        Printer.FontStrikethru = True
                        Printer.Print Spc(5); String(132, " ")
                        Printer.FontSize = 10
                        Printer.FontBold = False
                        Printer.FontStrikethru = False
                        Printer.Print
                        Printer.Print
                        Printer.Print Spc(7); "Visto Conferência: ______________"
                        Printer.NewPage
                    Else
                        Printer.Print
                    End If
                End If
                de_informa.rsSel_ReRelProt.MoveNext
            Loop
            'final do relatório, fecha com rodapé
            Printer.Print
            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.FontStrikethru = True
            Printer.Print Spc(5); String(132, " ")
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.FontStrikethru = False
            Printer.Print
            Printer.Print Spc(7); "Quantidade de CTCs: "; xcontctc
            Printer.Print Spc(7); "Data: "; datahora("data"); Space(35); "Assinatura Conferência: ______________________"
            Printer.NewPage
        Next
        Printer.EndDoc   'finaliza spool da impressão
        DoEvents
    
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "IMPRESSÃO", xusuario, "RE-IMPRESSÃO DO PROTOCOLO: " & txtNumProt
        
        MsgBox "RELATÓRIO ENVIADO PARA IMPRESSÃO !"
        txtNumProt.SetFocus
    End If
Else
    MsgBox "Número de Protocolo Inválido !"
    txtNumProt.SetFocus
End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReimprProt = Nothing
End Sub

Private Sub txtNumProt_Change()
    If Len(txtNumProt) > 0 Then
        cmdImprimir.Enabled = True
    Else
        cmdImprimir.Enabled = False
    End If
End Sub
