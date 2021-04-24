VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoIncluiNF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Incluir NFs"
   ClientHeight    =   2655
   ClientLeft      =   2835
   ClientTop       =   3525
   ClientWidth     =   8235
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoIncluiNF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6360
      TabIndex        =   5
      Top             =   2160
      Width           =   1755
   End
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "Continuar"
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   1680
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Importar NF de CTC"
      Height          =   1455
      Left            =   6360
      TabIndex        =   6
      Top             =   60
      Width           =   1755
      Begin VB.CommandButton CmdImportarNF 
         Caption         =   "Importar NFs"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1020
         Width           =   1515
      End
      Begin VB.TextBox TxtCTC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   600
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox TxtFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   0
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CTC"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   300
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexNF 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmEmissaoIncluiNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdCancelar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub CmdContinuar_Click()
Dim X, Y As Integer
Dim xSoma As Double

frmEmissao.FlexGridNFs.Clear
frmEmissao.FlexGridNFs.Rows = FlexNF.Rows
frmEmissao.FlexGridNFs.Cols = FlexNF.Cols
frmEmissao.FlexGridNFs.FixedCols = FlexNF.FixedCols
frmEmissao.FlexGridNFs.FixedRows = FlexNF.FixedRows
    For X = 0 To FlexNF.Cols - 1
    frmEmissao.FlexGridNFs.ColWidth(X) = FlexNF.ColWidth(X)
    Next

frmEmissao.FlexGridNFs.LeftCol = frmEmissao.FlexGridNFs.Cols - 1

    For Y = 0 To FlexNF.Rows - 1
        For X = 0 To FlexNF.Cols - 1
        frmEmissao.FlexGridNFs.TextMatrix(Y, X) = FlexNF.TextMatrix(Y, X)
        Next
    Next
    
    xSoma = 0
    
    For Y = 1 To frmEmissao.FlexGridNFs.Rows - 1
        If Len(frmEmissao.FlexGridNFs.TextMatrix(Y, 2)) > 0 Then
        xSoma = xSoma + CDbl(frmEmissao.FlexGridNFs.TextMatrix(Y, frmEmissao.FlexGridNFs.Cols - 1))
        End If
    Next

frmEmissao.TxtTotalVM.Text = Format((SemPonto(Str(xSoma * 100)) / 100), "###,##0.00")
Unload Me
End Sub

Private Sub CmdContinuar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub CmdImportarNF_Click()
Dim X, Y As Integer
Dim Continua As Boolean

Continua = True

If de_informa.rsSelNFdeCTC.State = 1 Then de_informa.rsSelNFdeCTC.Close
de_informa.SelNFdeCTC String(2 - Len(Trim(TxtFilial.Text)), "0") & Trim(TxtFilial.Text) & String(8 - Len(Trim(TxtCTC.Text)), "0") & Trim(TxtCTC.Text)

    If de_informa.rsSelNFdeCTC.RecordCount > 0 Then
        For Y = 1 To FlexNF.Rows - 1
            If Val(FlexNF.TextMatrix(Y, 1)) = Val(Mid(de_informa.rsSelNFdeCTC.Fields("filialctc"), 1, 2)) And Val(FlexNF.TextMatrix(Y, 2)) = Val(Mid(de_informa.rsSelNFdeCTC.Fields("filialctc"), 3)) Then
            MsgBox "Você já incluiu as Notas deste CTC!", vbExclamation, ""
            Continua = False
            Exit For
            End If
        Next
            If Continua = True Then
                Do Until de_informa.rsSelNFdeCTC.EOF
                    If de_informa.rsSelNFdeCTC.Fields("valornf") > 0 Then
                        If FlexNF.Rows <> 2 Then
                        FlexNF.AddItem ("")
                        Else
                            If Len(FlexNF.TextMatrix(1, 0)) > 0 Then
                            FlexNF.AddItem ("")
                            End If
                        End If
                    FlexNF.TextMatrix(FlexNF.Rows - 1, 0) = "N1"
                    FlexNF.TextMatrix(FlexNF.Rows - 1, 1) = Mid(de_informa.rsSelNFdeCTC.Fields("filialctc"), 1, 2)
                    FlexNF.TextMatrix(FlexNF.Rows - 1, 2) = Val(Mid(de_informa.rsSelNFdeCTC.Fields("filialctc"), 3))
                    FlexNF.TextMatrix(FlexNF.Rows - 1, 3) = de_informa.rsSelNFdeCTC.Fields("numnf")
                    If de_informa.rsSelNFdeCTC.Fields("serie") > 0 Then FlexNF.TextMatrix(FlexNF.Rows - 1, 4) = de_informa.rsSelNFdeCTC.Fields("serie")
                    FlexNF.TextMatrix(FlexNF.Rows - 1, 5) = Format(de_informa.rsSelNFdeCTC.Fields("valornf"), "###,##0.00")
                    End If
                    
                    If FlexNF.Rows >= 7 Then
                    FlexNF.TopRow = FlexNF.Rows - 6
                    DoEvents
                    End If
                de_informa.rsSelNFdeCTC.MoveNext
                Loop
            End If
    
    TxtFilial.SetFocus
    Else
    MsgBox "Não foram encontradas Notas Fiscais para este CTC.", vbCritical, ""
    TxtFilial.SetFocus
    End If
End Sub

Private Sub CmdImportarNF_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub FlexNF_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X, Y As Integer
Dim xCTC, xFilial As String

        

        If KeyCode = 46 Then
            If Len(Trim(FlexNF.TextMatrix(FlexNF.Row, 0))) > 0 Then
                xCTC = FlexNF.TextMatrix(FlexNF.Row, 2)
                xFilial = FlexNF.TextMatrix(FlexNF.Row, 1)
                For X = 1 To FlexNF.Rows - 1
                    If FlexNF.TextMatrix(X, 2) = xCTC And FlexNF.TextMatrix(X, 1) = xFilial Then
                    FlexNF.AddItem ("")
                    FlexNF.RemoveItem (X)
                    X = 0
                    End If
                Next
            Else
            FlexNF.RemoveItem (FlexNF.Row)
            If FlexNF.Rows = 1 Then FlexNF.AddItem ("")
            End If
                
            Y = 1
            Do While True
                If Y > FlexNF.Rows - 1 Then Exit Do
                If Len(FlexNF.TextMatrix(Y, 0)) = 0 Then
                    If FlexNF.Rows > 2 Then
                    FlexNF.RemoveItem (Y)
                    Y = Y - 1
                    Else
                    Exit Do
                    End If
                Else
                Y = Y + 1
                End If
            Loop
        End If

End Sub

Private Sub FlexNF_KeyPress(KeyAscii As Integer)
Dim X, Y As Integer
If KeyAscii = 27 Then Unload Me
X = FlexNF.Col
Y = FlexNF.Row

If FlexNF.TextMatrix(Y, 0) = "N1" And KeyAscii <> 13 Then KeyAscii = 0

    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
            If X = 0 Then
                If Len(FlexNF.TextMatrix(FlexNF.Row, X)) = 0 Then
                MsgBox "Você precisa informar se esta linha refere-se à uma NF ou Declaração.", vbExclamation, ""
                Exit Sub
                Else
                FlexNF.Col = X + 1
                End If
            ElseIf X = 1 Then
                If Len(FlexNF.TextMatrix(FlexNF.Row, X)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
                    If FlexNF.Row = 1 Then
                    MsgBox "Você deve informar a filial antes de prosseguir.", vbExclamation, ""
                    Exit Sub
                    Else
                    FlexNF.TextMatrix(FlexNF.Row, X) = FlexNF.TextMatrix(FlexNF.Row - 1, X)
                    FlexNF.Col = X + 1
                    End If
                Else
                FlexNF.Col = X + 1
                End If
            ElseIf X = 2 Then
                If Len(FlexNF.TextMatrix(FlexNF.Row, X)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
                    If FlexNF.Row = 1 Then
                    MsgBox "Você deve informar o CTC antes de prosseguir.", vbExclamation, ""
                    Exit Sub
                    Else
                    FlexNF.TextMatrix(FlexNF.Row, X) = FlexNF.TextMatrix(FlexNF.Row - 1, X)
                    FlexNF.Col = X + 1
                    End If
                Else
                FlexNF.Col = X + 1
                End If
            ElseIf X = 3 Then
                If Len(FlexNF.TextMatrix(FlexNF.Row, X)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
                MsgBox "Você deve informar a NF antes de prosseguir.", vbExclamation, ""
                Exit Sub
                Else
                FlexNF.Col = X + 1
                End If
            ElseIf X = 4 Then
                FlexNF.Col = X + 1
            ElseIf X = 5 Then
                If Len(FlexNF.TextMatrix(FlexNF.Row, X)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
                MsgBox "Você deve informar o Valor da NF antes de prosseguir.", vbExclamation, ""
                Exit Sub
                Else
                FlexNF.AddItem ("")
                FlexNF.Col = 0
                FlexNF.Row = FlexNF.Row + 1
                If FlexNF.Rows >= 7 Then
                    FlexNF.TopRow = FlexNF.Rows - 6
                    DoEvents
                    End If
                End If
            End If
        ElseIf KeyAscii = 8 Then
            If Len(FlexNF.Text) = 0 Then
            KeyAscii = 0
            Else
                If X = 5 Then
                FlexNF.Text = Mid(FlexNF.Text, 1, Len(FlexNF.Text) - 1)
                FlexNF.Text = Format((SemPonto(FlexNF.Text) / 100), "###,##0.00")
                Else
                FlexNF.Text = Mid(FlexNF.Text, 1, Len(FlexNF.Text) - 1)
                End If
            End If
        ElseIf KeyAscii = 78 Or KeyAscii = 110 Then
            If X = 0 Then
            FlexNF.Text = "N"
            End If
        ElseIf KeyAscii = 68 Or KeyAscii = 100 Then
            If X = 0 Then
            FlexNF.Text = "D"
            End If
        End If
    Else
        If X = 1 Then
            If Len(FlexNF.TextMatrix(FlexNF.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informado se isto é uma NF ou declaração.", vbExclamation, ""
            Exit Sub
            Else
            FlexNF.TextMatrix(FlexNF.Row, X) = FlexNF.TextMatrix(FlexNF.Row, X) & Chr(KeyAscii)
            End If
        ElseIf X = 2 Then
            If Len(FlexNF.TextMatrix(FlexNF.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informado se isto é uma NF ou declaração.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 1)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a Filial.", vbExclamation, ""
            Exit Sub
            Else
            FlexNF.TextMatrix(FlexNF.Row, X) = FlexNF.TextMatrix(FlexNF.Row, X) & Chr(KeyAscii)
            End If
        ElseIf X = 3 Then
            If Len(FlexNF.TextMatrix(FlexNF.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informado se isto é uma NF ou declaração.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 1)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a Filial.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 2)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informado o CTC.", vbExclamation, ""
            Exit Sub
            Else
            FlexNF.TextMatrix(FlexNF.Row, X) = FlexNF.TextMatrix(FlexNF.Row, X) & Chr(KeyAscii)
            End If
        ElseIf X = 4 Then
            If Len(FlexNF.TextMatrix(FlexNF.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informado se isto é uma NF ou declaração.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 1)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a Filial.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 2)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informado o CTC.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 3)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a NF.", vbExclamation, ""
            Exit Sub
            Else
            FlexNF.TextMatrix(FlexNF.Row, X) = FlexNF.TextMatrix(FlexNF.Row, X) & Chr(KeyAscii)
            End If
        ElseIf X = 5 Then
            If Len(FlexNF.TextMatrix(FlexNF.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informado se isto é uma NF ou declaração.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 1)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a Filial.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 2)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informado o CTC.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexNF.TextMatrix(FlexNF.Row, 3)) = 0 And FlexNF.TextMatrix(FlexNF.Row, 0) = "N" Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a NF.", vbExclamation, ""
            Exit Sub
            Else
            FlexNF.TextMatrix(FlexNF.Row, X) = FlexNF.TextMatrix(FlexNF.Row, X) & Chr(KeyAscii)
            FlexNF.Text = Format((SemPonto(FlexNF.Text) / 100), "###,##0.00")
            End If
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim X, Y As Integer
    
    If frmEmissao.FlexGridNFs.Rows > 0 Then
    FlexNF.Rows = frmEmissao.FlexGridNFs.Rows
    Else
    FlexNF.Rows = 2
    End If
    
FlexNF.Cols = 6
FlexNF.FixedRows = 1
FlexNF.FixedCols = 0
FlexNF.TextMatrix(0, 0) = "N/D"
FlexNF.TextMatrix(0, 1) = "Filial"
FlexNF.TextMatrix(0, 2) = "CTC"
FlexNF.TextMatrix(0, 3) = "Nº NF"
FlexNF.TextMatrix(0, 4) = "Série"
FlexNF.TextMatrix(0, 5) = "Valor"
FlexNF.ColWidth(0) = 500
FlexNF.ColWidth(1) = 900
FlexNF.ColWidth(2) = 1200
FlexNF.ColWidth(3) = 1200
FlexNF.ColWidth(4) = 700
FlexNF.ColWidth(5) = 1200

    For Y = 1 To frmEmissao.FlexGridNFs.Rows - 1
        For X = 0 To frmEmissao.FlexGridNFs.Cols - 1
        FlexNF.TextMatrix(Y, X) = frmEmissao.FlexGridNFs.TextMatrix(Y, X)
        Next
    Next
    

End Sub

Private Sub TxtCTC_Change()
    If Len(TxtCTC.Text) >= 8 Then SendKeys "{TAB}"
End Sub

Private Sub TxtCTC_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 200
End Sub

Private Sub TxtCTC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub TxtFilial_Change()
    If Len(TxtFilial.Text) >= 2 Then SendKeys "{TAB}"
End Sub

Private Sub TxtFilial_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 200
End Sub

Private Sub TxtFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If KeyAscii = 27 Then Unload Me
End Sub
