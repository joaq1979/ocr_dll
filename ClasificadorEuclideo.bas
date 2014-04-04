Attribute VB_Name = "ClasificadorEuclideo"

Function clasifica(distanciaMinima As Double) As Integer
'Obtiene los vectores sobre los cuales se realiza la distancia euclidea
 Dim elem As Variant
 Dim distancia As Double
 Dim elemProximo As Integer
 
 distanciaMinima = 999999999

  For Each elem In clsTipos.varCaracteristicasMedias
  
        distancia = clasificadorEuclideo(CInt(elem))
  
        If distancia < distanciaMinima Then
            distanciaMinima = distancia
            elemProximo = CInt(elem)
        End If

  Next elem

  clasifica = elemProximo
End Function


Private Function clasificadorEuclideo(elem As Integer) As Double
'Aplica la funcion de la distancia Euclidea
 Dim i As Integer
 Dim distancia As Double
 distancia = 0
 
  For i = 0 To UBound(clsTipos.varCaracteristicasImg)
    distancia = distancia + (clsTipos.varCaracteristicasImg(i) - clsTipos.varCaracteristicas(i, elem)) ^ 2
  Next i
  clasificadorEuclideo = Sqr(distancia)
End Function
