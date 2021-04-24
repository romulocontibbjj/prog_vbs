VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoIncluiVolume 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe as Medidas do Volume"
   ClientHeight    =   3015
   ClientLeft      =   2700
   ClientTop       =   3885
   ClientWidth     =   6570
   Icon            =   "frmEmissaoIncluiVolume.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "Continuar"
      Height          =   315
      Left            =   4380
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid FlexVol 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4260
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
   End
End
Attribute VB_Name = "frmEmissaoIncluiVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdContinuar_Click()
Dim X, Y, Z, A, B, xVolumes, xVolumesTotal As Integer
Dim xSoma, xCub, xPesoReal, xCubTotal, xPesoTotal As Double
Dim xQuit As Boolean
xPesoTotal = 0
xCubTotal = 0

xQuit = True
A = FlexVol.Rows - 1
    If Val(FlexVol.TextMatrix(A, 0)) = 0 Or Val(FlexVol.TextMatrix(A, 1)) = 0 Or Val(FlexVol.TextMatrix(A, 2)) = 0 Or Val(FlexVol.TextMatrix(A, 3)) = 0 Or Val(SemPonto(FlexVol.TextMatrix(A, 4))) = 0 Then
        If FlexVol.Rows > 2 Then
        FlexVol.RemoveItem (A)
        xQuit = False
        End If
    Else
    xQuit = False
    End If

If xQuit = False Then
    For A = 1 To FlexVol.Rows - 1
    xVolumes = Val(FlexVol.TextMatrix(A, 0))
    X = Val(FlexVol.TextMatrix(A, 1))
    Y = Val(FlexVol.TextMatrix(A, 2))
    Z = Val(FlexVol.TextMatrix(A, 3))
    xPesoReal = Val(FlexVol.TextMatrix(A, 4))
    xCub = ((X * Y * Z) / 6000) * xVolumes
    xPesoTotal = xPesoTotal + xPesoReal ' * xVolumes
    xCubTotal = xCubTotal + xCub
    Next

frmEmissao.FlexGridVolumes.Clear
frmEmissao.FlexGridVolumes.Rows = FlexVol.Rows
frmEmissao.FlexGridVolumes.Cols = FlexVol.Cols
frmEmissao.FlexGridVolumes.FixedCols = FlexVol.FixedCols
frmEmissao.FlexGridVolumes.FixedRows = FlexVol.FixedRows
frmEmissao.FlexGridVolumes.TextMatrix(0, 0) = "Volumes (Cm)"
frmEmissao.FlexGridVolumes.TextMatrix(0, 1) = "Comprimento (Cm)"
frmEmissao.FlexGridVolumes.TextMatrix(0, 2) = "Largura (Cm)"
frmEmissao.FlexGridVolumes.TextMatrix(0, 3) = "Altura (Cm)"
frmEmissao.FlexGridVolumes.TextMatrix(0, 4) = "Peso Real"

    For Y = 0 To FlexVol.Rows - 1
        For X = 0 To FlexVol.Cols - 1
        frmEmissao.FlexGridVolumes.TextMatrix(Y, X) = FlexVol.TextMatrix(Y, X)
        frmEmissao.FlexGridVolumes.ColWidth(X) = FlexVol.ColWidth(X)
        Next
    Next

'***********Alteração - Lincoln - Pesos em intervalos de 0.5 **********************Start
'***********Calcula o decimal -> valor menos a parte INTEIRA **************************

If frmEmissao.TxtSiglaCiaAerea.Text = "RG" Then
    'Calcular c/ PESO CUBADO
    dec_peso = (CDbl(xCubTotal)) - (Int(xCubTotal))
   
    If (dec_peso >= "0,01") And (dec_peso <= "0,49") Then
        xCubTotal = Int(xCubTotal) + 0.5
    
    ElseIf (dec_peso >= "0,51") And (dec_peso <= "0,99") Then
        xCubTotal = Int(xCubTotal) + 1
    End If
End If

'***********Alteração - Lincoln - Pesos em intervalos de 0.5 **********************End


frmEmissao.TxtPesoCubado = Format(xCubTotal, "#####0.0")
frmEmissao.TxtPesoReal = Format(xPesoTotal, "###,##0.0")
End If

Unload Me
End Sub

Private Sub FlexVol_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X, Y As Integer
Y = FlexVol.Row
If KeyCode = 46 Then
    If FlexVol.Rows > 2 Then
    FlexVol.RemoveItem (Y)
    Else
    FlexVol.AddItem ("")
    FlexVol.RemoveItem (Y)
    End If
End If
End Sub

Private Sub FlexVol_KeyPress(KeyAscii As Integer)
Dim X, Y As Integer

X = FlexVol.Col
Y = FlexVol.Row

    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
            If X = 0 Then
                If Len(FlexVol.TextMatrix(FlexVol.Row, X)) = 0 Then
                MsgBox "Você precisa informar a qunatidade de Volumes antes de continuar.", vbExclamation, ""
                Exit Sub
                ElseIf Val(FlexVol.TextMatrix(FlexVol.Row, X)) = 0 Then
                MsgBox "Não são aceitos valores nulos para este campo.", vbExclamation, ""
                Exit Sub
                Else
                FlexVol.Col = X + 1
                End If
            ElseIf X = 1 Then
                If Len(FlexVol.TextMatrix(FlexVol.Row, X)) = 0 Then
                MsgBox "Você deve informar o COMPRIMENTO antes de prosseguir.", vbExclamation, ""
                Exit Sub
                ElseIf Val(FlexVol.TextMatrix(FlexVol.Row, X)) = 0 Then
                MsgBox "Não são aceitos valores nulos para este campo.", vbExclamation, ""
                Exit Sub
                Else
                FlexVol.Col = X + 1
                End If
            ElseIf X = 2 Then
                If Len(FlexVol.TextMatrix(FlexVol.Row, X)) = 0 Then
                MsgBox "Você deve informar a LARGURA antes de prosseguir.", vbExclamation, ""
                Exit Sub
                ElseIf Val(FlexVol.TextMatrix(FlexVol.Row, X)) = 0 Then
                MsgBox "Não são aceitos valores nulos para este campo.", vbExclamation, ""
                Exit Sub
                Else
                FlexVol.Col = X + 1
                End If
            ElseIf X = 3 Then
                If Len(FlexVol.TextMatrix(FlexVol.Row, X)) = 0 Then
                MsgBox "Você deve informar a ALTURA antes de prosseguir.", vbExclamation, ""
                Exit Sub
                ElseIf Val(FlexVol.TextMatrix(FlexVol.Row, X)) = 0 Then
                MsgBox "Não são aceitos valores nulos para este campo.", vbExclamation, ""
                Exit Sub
                Else
                FlexVol.Col = X + 1
                End If
            ElseIf X = 4 Then
                FlexVol.AddItem ("")
                FlexVol.Col = 0
                FlexVol.Row = FlexVol.Row + 1
            End If
        ElseIf KeyAscii = 8 Then
            If Len(FlexVol.Text) = 0 Then
            KeyAscii = 0
            Else
            FlexVol.Text = Mid(FlexVol.Text, 1, Len(FlexVol.Text) - 1)
            End If
        End If
    Else
        If X = 0 Then
        FlexVol.TextMatrix(FlexVol.Row, X) = FlexVol.TextMatrix(FlexVol.Row, X) & Chr(KeyAscii)
        ElseIf X = 1 Then
            If Len(FlexVol.TextMatrix(FlexVol.Row, 0)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a qte. de Volumes.", vbExclamation, ""
            Exit Sub
            Else
            FlexVol.TextMatrix(FlexVol.Row, X) = FlexVol.TextMatrix(FlexVol.Row, X) & Chr(KeyAscii)
            End If
        ElseIf X = 2 Then
            If Len(FlexVol.TextMatrix(FlexVol.Row, 0)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a qte. de Volumes.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexVol.TextMatrix(FlexVol.Row, 1)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 1)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informado o COMPRIMENTO.", vbExclamation, ""
            Exit Sub
            Else
            FlexVol.TextMatrix(FlexVol.Row, X) = FlexVol.TextMatrix(FlexVol.Row, X) & Chr(KeyAscii)
            End If
        ElseIf X = 3 Then
            If Len(FlexVol.TextMatrix(FlexVol.Row, 0)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a qte. de Volumes.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexVol.TextMatrix(FlexVol.Row, 1)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 1)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informado o COMPRIMENTO.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexVol.TextMatrix(FlexVol.Row, 2)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 2)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a LARGURA.", vbExclamation, ""
            Exit Sub
            Else
            FlexVol.TextMatrix(FlexVol.Row, X) = FlexVol.TextMatrix(FlexVol.Row, X) & Chr(KeyAscii)
            End If
        ElseIf X = 4 Then
            If Len(FlexVol.TextMatrix(FlexVol.Row, 0)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 0)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a qte. de Volumes.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexVol.TextMatrix(FlexVol.Row, 1)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 1)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informado o COMPRIMENTO.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexVol.TextMatrix(FlexVol.Row, 2)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 2)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a LARGURA.", vbExclamation, ""
            Exit Sub
            ElseIf Len(FlexVol.TextMatrix(FlexVol.Row, 3)) = 0 Or Val(FlexVol.TextMatrix(FlexVol.Row, 3)) = 0 Then
            MsgBox "Não é possível editar esta célula, já que não foi informada a ALTURA.", vbExclamation, ""
            Exit Sub
            Else
            FlexVol.TextMatrix(FlexVol.Row, X) = FlexVol.TextMatrix(FlexVol.Row, X) & Chr(KeyAscii)
            End If
        End If
    End If
    
End Sub

Private Sub Form_Load()


    If frmEmissao.FlexGridVolumes.Rows > 0 Then
    FlexVol.Rows = frmEmissao.FlexGridVolumes.Rows
    Else
    FlexVol.Rows = 2
    End If

FlexVol.Cols = 5
FlexVol.FixedCols = 0
FlexVol.FixedRows = 1
FlexVol.TextMatrix(0, 0) = "Vol. (Cm)"
FlexVol.TextMatrix(0, 1) = "Comp. (Cm)"
FlexVol.TextMatrix(0, 2) = "Larg. (Cm)"
FlexVol.TextMatrix(0, 3) = "Alt. (Cm)"
FlexVol.TextMatrix(0, 4) = "Peso (Kg)"
FlexVol.ColWidth(0) = 900
FlexVol.ColWidth(1) = 900
FlexVol.ColWidth(2) = 900
FlexVol.ColWidth(3) = 900
FlexVol.ColWidth(4) = 900


    For Y = 1 To frmEmissao.FlexGridVolumes.Rows - 1
        For X = 0 To frmEmissao.FlexGridVolumes.Cols - 1
        FlexVol.TextMatrix(Y, X) = frmEmissao.FlexGridVolumes.TextMatrix(Y, X)
        Next
    Next
End Sub
