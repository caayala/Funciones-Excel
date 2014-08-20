Option Explicit

Function errormuestralinf(muestra As Long, Optional signif As Double = 95, _
    Optional p As Double = 0.5) As Single

' Formula de error muestral si la poblacion es infinita
' Supone muestreo aleatorio simple

If signif > 1 Then
    signif = signif / 100
End If

        errormuestralinf = _
            (Application.WorksheetFunction.NormSInv((1 + (signif)) / 2) _
            * Sqr(p * (1 - p) / muestra) * 100)

End Function

Function errormuestralfin(muestra As Long, pobTot As Long, Optional signif As Double = 95, _
    Optional p As Double = 0.5) As Double

' Formula de error muestral si la poblacion es finita
' Supone muestreo aleatorio simple

If signif > 1 Then
    signif = signif / 100
End If

    errormuestralfin = _
        (Application.WorksheetFunction.NormSInv((1 + (signif)) / 2) _
        * Sqr(p * (1 - p) / muestra) * 100) * Sqr((pobTot - muestra) / (pobTot - 1))

End Function


Function tamuestra(error As Double, pobTot As Long, Optional signif As Double = 95, _
    Optional p As Double = 0.5) As Double

' Formula para calcular el tamano de una muestra dado un error muestral.
' Supone muestreo aleatorio simple

Dim pobinf As Double

If signif > 1 Then
    signif = signif / 100
End If

pobinf = WorksheetFunction.NormSInv(1 / 2 + (signif) / 2) ^ 2 * p * (1 - p) / error ^ 2

tamuestra = WorksheetFunction.RoundUp(pobinf / (1 + (pobinf - 1) / pobTot), 0)

End Function




