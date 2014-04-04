Attribute VB_Name = "clsMomentos"
Private Function momentosimple(p As Integer, q As Integer) As Double

    Dim x As Integer, y As Integer
    Dim res As Double
    res = 0
    
    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            res = res + ((x + 1) ^ p) * ((y + 1) ^ q) * clsTipos.varImagen.MatrizGris(x, y)
        Next y
    Next x
    momentosimple = res

End Function

Private Function momentocentral(p As Integer, q As Integer) As Double
    
    Dim x As Integer, y As Integer
    Dim res As Double
    Dim xmedio As Double, ymedio As Double
    
    xmedio = momentosimple(1, 0) / momentosimple(0, 0)
    ymedio = momentosimple(0, 1) / momentosimple(0, 0)
       
    
    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            res = res + (((x + 1) - xmedio) ^ p) * (((y + 1) - ymedio) ^ q) * clsTipos.varImagen.MatrizGris(x, y)
        Next y
    Next x
    
    momentocentral = res
End Function

Private Function mcn(p As Integer, q As Integer) As Double
    Dim alfa As Double
    
    alfa = ((p + q) / 2) + 1
    mcn = momentocentral(p, q) / (momentocentral(0, 0) ^ alfa)
    
End Function

Sub momentosHu(i As Integer)
    
    If i >= 0 Then
   
       'centroide del caracter
       clsTipos.varCaracteristicas(0, i) = (momentosimple(1, 0) / momentosimple(0, 0)) / 150
       clsTipos.varCaracteristicas(1, i) = (momentosimple(0, 1) / momentosimple(0, 0)) / 150
       'los momentos de Hu
       clsTipos.varCaracteristicas(15, i) = (mcn(2, 0) + mcn(0, 2)) / 1000
       clsTipos.varCaracteristicas(16, i) = ((mcn(2, 0) - mcn(0, 2)) ^ 2 + 4 * mcn(1, 1) ^ 2) / 1000
       clsTipos.varCaracteristicas(17, i) = ((mcn(3, 0) - 3 * mcn(1, 2)) ^ 2 + (3 * mcn(2, 1) - mcn(0, 3)) ^ 2) / 1000
 '      clsTipos.varCaracteristicas(18, i) = ((mcn(3, 0) + mcn(1, 2)) ^ 2 + (mcn(2, 1) + mcn(0, 3)) ^ 2) / 1000
 '      clsTipos.varCaracteristicas(19, i) = ((mcn(3, 0) - 3 * mcn(1, 2)) * (mcn(3, 0) + mcn(1, 2)) * _
                                        ((mcn(3, 0) + mcn(1, 2)) ^ 2 - 3 * (mcn(2, 1) + mcn(0, 3)) ^ 2) + _
                                        (3 * mcn(2, 1) - mcn(0, 3)) * (mcn(2, 1) + mcn(0, 3)) * ((3 * mcn(3, 0) + mcn(1, 2)) ^ 2 - _
                                        (mcn(2, 1) + mcn(0, 3)) ^ 2)) / 1000
 '      clsTipos.varCaracteristicas(20, i) = ((mcn(2, 0) - mcn(0, 2)) * ((mcn(3, 0) + mcn(1, 2)) ^ 2 - (mcn(2, 1) + mcn(0, 3)) ^ 2) + _
                                        4 * mcn(1, 1) * (mcn(3, 0) + mcn(1, 2)) * (mcn(2, 1) + mcn(0, 3))) / 1000
    
    Else
 
       'centroide del caracter
       clsTipos.varCaracteristicasImg(0) = (momentosimple(1, 0) / momentosimple(0, 0)) / 150
       clsTipos.varCaracteristicasImg(1) = (momentosimple(0, 1) / momentosimple(0, 0)) / 150
       'Momentos de Hu
       clsTipos.varCaracteristicasImg(15) = (mcn(2, 0) + mcn(0, 2)) / 1000
       clsTipos.varCaracteristicasImg(16) = ((mcn(2, 0) - mcn(0, 2)) ^ 2 + 4 * mcn(1, 1) ^ 2) / 1000
       clsTipos.varCaracteristicasImg(17) = ((mcn(3, 0) - 3 * mcn(1, 2)) ^ 2 + (3 * mcn(2, 1) - mcn(0, 3)) ^ 2) / 1000
 '      clsTipos.varCaracteristicasImg(18) = ((mcn(3, 0) + mcn(1, 2)) ^ 2 + (mcn(2, 1) + mcn(0, 3)) ^ 2) / 1000
 '      clsTipos.varCaracteristicasImg(19) = ((mcn(3, 0) - 3 * mcn(1, 2)) * (mcn(3, 0) + mcn(1, 2)) * _
                                        ((mcn(3, 0) + mcn(1, 2)) ^ 2 - 3 * (mcn(2, 1) + mcn(0, 3)) ^ 2) + _
                                        (3 * mcn(2, 1) - mcn(0, 3)) * (mcn(2, 1) + mcn(0, 3)) * ((3 * mcn(3, 0) + mcn(1, 2)) ^ 2 - _
                                        (mcn(2, 1) + mcn(0, 3)) ^ 2)) / 1000
                                        
 '      clsTipos.varCaracteristicasImg(20) = ((mcn(2, 0) - mcn(0, 2)) * ((mcn(3, 0) + mcn(1, 2)) ^ 2 - (mcn(2, 1) + mcn(0, 3)) ^ 2) + _
                                        4 * mcn(1, 1) * (mcn(3, 0) + mcn(1, 2)) * (mcn(2, 1) + mcn(0, 3))) / 1000
 
    End If
    
End Sub










