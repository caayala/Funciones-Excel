Option Explicit

Function interconfianza(prom As Single, DesStd As Single, casos As Single, _
    Optional sup As Boolean = True, Optional signif As Double = 95) As Double
    'C‡lculo de intervalos de confianza a partir de promedios.

If sup = True Then
    interconfianza = prom + WorksheetFunction.TInv(1 - signif / 100, casos - 1) * DesStd / (casos) ^ (1 / 2)
Else
    interconfianza = prom - WorksheetFunction.TInv(1 - signif / 100, casos - 1) * DesStd / (casos) ^ (1 / 2)
End If

End Function


Function errormuestralinf(muestra As Long, Optional signif As Double = 95, _
    Optional p As Double = 0.5) As Single
' Formula de error muestral si la poblacion es infinita

If signif > 1 Then
    signif = signif / 100
End If

        errormuestralinf = _
            (Application.WorksheetFunction.NormSInv((1 + (signif)) / 2) _
            * Sqr(p * (1 - p) / muestra) * 100)

End Function

Function errormuestralfin(muestra As Long, pobTot As Long, Optional signif As Double = 95, _
    Optional p As Double = 0.5) As Double
' Formula de error muestral si la poblaci—n es finita
If signif > 1 Then
    signif = signif / 100
End If

    errormuestralfin = _
        (Application.WorksheetFunction.NormSInv((1 + (signif)) / 2) _
        * Sqr(p * (1 - p) / muestra) * 100) * Sqr((pobTot - muestra) / (pobTot - 1))

End Function


Function tamuestra(error As Double, pobTot As Long, Optional signif As Double = 95, _
    Optional p As Double = 0.5) As Double
' F—rmula de tama–o de muestra.

Dim pobinf As Double

If signif > 1 Then
    signif = signif / 100
End If

pobinf = WorksheetFunction.NormSInv(1 / 2 + (signif) / 2) ^ 2 * p * (1 - p) / error ^ 2

tamuestra = WorksheetFunction.RoundUp(pobinf / (1 + (pobinf - 1) / pobTot), 0)

End Function




