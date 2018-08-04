Option Explicit

Function errormuestralinf(muestra As Long, _ 
                          Optional p As Single = 0.5, _ 
                          Optional signif As Single = 95) As Single

' Formula de error muestral si la poblacion es infinita
' Supone muestreo aleatorio simple

If signif > 1 Then
    signif = signif / 100
End If

        errormuestralinf = _
            (Application.WorksheetFunction.NormSInv((1 + (signif)) / 2) _
            * Sqr(p * (1 - p) / muestra) * 100)

End Function

Function errormuestralfin(muestra As Long, _ 
                          pobTot As Long, _ 
                          Optional p As Single = 0.5, _ 
                          Optional signif As Single = 95) As Single

' Formula de error muestral si la poblacion es finita
' Supone muestreo aleatorio simple

If signif > 1 Then
    signif = signif / 100
End If

    errormuestralfin = _
        (Application.WorksheetFunction.NormSInv((1 + (signif)) / 2) _
        * Sqr(p * (1 - p) / muestra) * 100) * Sqr((pobTot - muestra) / (pobTot - 1))

End Function


Function tamuestra(error As Double, _ 
                   pobTot As Long, _ 
                   Optional p As Single = 0.5, _ 
                   Optional signif As Single = 95) As Single

' Formula para calcular el tamano de una muestra dado un error muestral.
' Supone muestreo aleatorio simple

Dim pobinf As Double

If signif > 1 Then
    signif = signif / 100
End If

pobinf = WorksheetFunction.NormSInv(1 / 2 + (signif) / 2) ^ 2 * p * (1 - p) / error ^ 2

tamuestra = WorksheetFunction.RoundUp(pobinf / (1 + (pobinf - 1) / pobTot), 0)

End Function


Function errormuestral_dist(muestra As Range, _ 
                            pob As Range, _ 
                            Optional p As Single = 0.5, _ 
                            Optional signif As Single = 95) As Single

' Funcion para calculo de error muestral considerando la distribucion de estratos

Dim pob_total As Double
Dim muestra_total As Double
Dim a As Double, b As Double

Dim pob_address  As String
Dim muestra_address  As String

pob_address = pob.Address
muestra_address = muestra.Address

pob_total = WorksheetFunction.Sum(pob)
muestra_total = WorksheetFunction.Sum(muestra)

If signif > 1 Then
    signif = signif / 100
End If

signif = WorksheetFunction.NormSInv(1 - (1 - signif) / 2)

a = Evaluate("SUM((" & pob_address & ") ^ 2*" & p & "*(1-" & p & ") / (" _
    & muestra_address & " / " & muestra_total & "))")

b = Evaluate("SUM((" & pob_address & ")*" & p & "*(1-" & p & "))")

errormuestral_dist = signif * Sqr(a - b * muestra_total) / (Sqr(muestra_total) * pob_total)

End Function


