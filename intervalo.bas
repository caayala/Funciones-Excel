Option Explicit

Function interconfianza(prom As Single, DesStd As Single, casos As Single, _
    Optional sup As Boolean = True, Optional signif As Double = 95) As Double
    'Cálculo de intervalos de confianza a partir de promedios.

If sup = True Then
    interconfianza = prom + WorksheetFunction.TInv(1 - signif / 100, casos - 1) * DesStd / (casos) ^ (1 / 2)
Else
    interconfianza = prom - WorksheetFunction.TInv(1 - signif / 100, casos - 1) * DesStd / (casos) ^ (1 / 2)
End If

End Function