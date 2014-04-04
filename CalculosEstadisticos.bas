Attribute VB_Name = "CalculosEstadisticos"

Sub calculaMediaCaracteristicas(inicio As Integer, max As Integer)
 Dim j As Integer
 Dim divisor As Integer

 divisor = 1
 For j = inicio + 1 To max
    If clsTipos.varNombres(0, j) = clsTipos.varNombres(0, inicio) Then
        Call sumaCaracteristicas(inicio, j)
        divisor = divisor + 1
        'marcamos los elementos tratados
        clsTipos.varNombres(1, j) = "T"
    End If
 Next j
  
 Call divideCaracteristicas(inicio, divisor)
   
End Sub

Private Sub sumaCaracteristicas(inicio As Integer, j As Integer)
    Dim i As Integer
    
    For i = 0 To UBound(clsTipos.varCaracteristicas)
        clsTipos.varCaracteristicas(i, inicio) = clsTipos.varCaracteristicas(i, inicio) + clsTipos.varCaracteristicas(i, j)
    Next i

End Sub

Private Sub divideCaracteristicas(inicio As Integer, divisor As Integer)
    Dim i As Integer
    
    For i = 0 To UBound(clsTipos.varCaracteristicas)
        clsTipos.varCaracteristicas(i, inicio) = clsTipos.varCaracteristicas(i, inicio) / divisor
    Next i
    
End Sub
