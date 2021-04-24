Attribute VB_Name = "Module2"
Public Function DiasUteis(dtInicio As Variant, dtFinal As Variant) As Integer

Dim intSemanas As Integer
Dim varDataCont As Variant
Dim intFimDias As Integer

dtInicio = DateValue(dtInicio)
dtFinal = DateValue(dtFinal)
intSemanas = DateDiff("w", dtInicio, dtFinal)
varDataCont = DateAdd("ww", intSemanas, dtInicio)
intFimDias = 0

Do While varDataCont < dtFinal
If Format(varDataCont, "ddd") <> "Sun" And _
Format(varDataCont, "ddd") <> "Sat" Then
intFimDias = intFimDias + 1
End If
varDataCont = DateAdd("d", 1, varDataCont)
Loop

DiasUteis = intSemanas * 5 + intFimDias

End Function

 Public Function ImpressoraInstalada() As Boolean
    On Error Resume Next
    
    Dim strVerifica As String
    strVerifica = Printer.DeviceName
    
    If Err.Number Then
        ImpressoraInstalada = False
    Else
        ImpressoraInstalada = True
    End If
    
End Function



