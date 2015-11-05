Option Explicit

Function cuadrato(pregunta As String, alternativa As Variant, _
    segmento As String, rgMat As string, Optional posicion As Integer = 0, _
    Optional error As Boolean = False) As Variant

' by Cristian Ayala

Dim fil As Integer, col As Integer, aux As Integer
Dim rg As Range
Dim Msg As String

' Set up error handling
' On Error Resume Next
On Error GoTo BadEntry

Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.ScreenUpdating = False

Set rg = Range(rgMat)
With rg
    ' Busca el numero de fila en la primera columna
    fil = WorksheetFunction.Match(pregunta, .Columns(1).value, 0)
    ' Aux es el numero de fila donde termina el rango de la pregunta en la que busco
    aux = .Cells(fil, 1).MergeArea.Count
    
    ' Busca el numero de fila de en la segunda columna
    fil = WorksheetFunction.Match(alternativa, Range(.Cells(fil, 2), .Cells(fil + aux - 1, 2)).value, 0) + fil - 1
    
    ' Busca el numero de columna en la segunda fila
    col = WorksheetFunction.Match(segmento, .Rows(2).value, 0)
    
    ' Ajusta el numero de columna a la posicion en que esta el dato de interes
    cuadrato = .Cells(fil, col + posicion).Value

    If Right(.Cells(fil, col + posicion).NumberFormat, 1) = "%" Then
        cuadrato = cuadrato * 100
    End If

End With

Set rg = Nothing

GoTo Done

BadEntry:
    Set rg = Nothing
    
    If error = True Then
        cuadrato = 0
    Else
        cuadrato = vbNullString ' Igual que "" pero no ocupa memoria.
    End If

Done:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

'    Msg = "An error occurred." & vbNewLine
'    MsgBox Msg
End Function

Function cuadratoRango(pregunta As String, alternativa As Range, _
    segmento As String, rgMat As String, Optional posicion As Integer = 0, _
    Optional error As Boolean = False) As Variant

' suma los valores asignados al rango de alternativas ingresadas.
' necesita de funcion cuadrato para funcionar

Dim c As Range

For Each c In alternativa.Cells
    If IsEmpty(c) = False Then
        cuadratoRango = cuadrato(pregunta, c, segmento, rgMat, posicion, error) + cuadratoRango
    End If
Next

End Function