Attribute VB_Name = "Esqueleto"

'hago mas fino el trazo de las imagenes
Sub esqueletoImagen()
    Dim salir1 As Boolean, salir2 As Boolean
    Dim x As Integer, y As Integer
    Dim count As Integer
    Dim tran As Integer
    
    salir1 = True
    salir2 = True
    While (salir1 And salir2)
        salir1 = False
        salir2 = False
        
        'primera iteración
        For x = 1 To clsTipos.varImagen.ancho - 2
            For y = 1 To clsTipos.varImagen.Alto - 2
                'si el pixel pertenece al caracter
                If clsTipos.varImagen.MatrizGris(x, y) = 255 Then
                    count = vecinos(x, y)
                    'si el pixel pertenece al contorno
                    If count <> 8 Then
                        If (count >= 2 And count <= 6) Then
                            tran = transiciones(x, y)
                            If tran = 1 Then
                                If clsTipos.varImagen.MatrizGris(x, y - 1) = 0 Or clsTipos.varImagen.MatrizGris(x + 1, y) = 0 Or clsTipos.varImagen.MatrizGris(x, y + 1) = 0 Then
                                    If clsTipos.varImagen.MatrizGris(x + 1, y) = 0 Or clsTipos.varImagen.MatrizGris(x, y + 1) = 0 Or clsTipos.varImagen.MatrizGris(x - 1, y) = 0 Then
                                        clsTipos.varImagen.MatrizGris(x, y) = 0
                                        salir1 = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next y
        Next x
        
        'segunda iteración
        For x = 1 To clsTipos.varImagen.ancho - 2
            For y = 1 To clsTipos.varImagen.Alto - 2
                'si el pixel pertenece al caracter
                If clsTipos.varImagen.MatrizGris(x, y) = 255 Then
                    count = vecinos(x, y)
                    'si el pixel pertenece al contorno
                    If count <> 8 Then
                        If (count >= 2 And count <= 6) Then
                            tran = transiciones(x, y)
                            If tran = 1 Then
                                If clsTipos.varImagen.MatrizGris(x, y - 1) = 0 Or clsTipos.varImagen.MatrizGris(x + 1, y) = 0 Or clsTipos.varImagen.MatrizGris(x - 1, y) = 0 Then
                                    If clsTipos.varImagen.MatrizGris(x, y - 1) = 0 Or clsTipos.varImagen.MatrizGris(x, y + 1) = 0 Or clsTipos.varImagen.MatrizGris(x - 1, y) = 0 Then
                                        clsTipos.varImagen.MatrizGris(x, y) = 0
                                        salir2 = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next y
        Next x
    Wend
    
    
    'Elimina mucha de las ramificaciones secundarias
    Call eliminaArtefactos
    'Elimina pixeles que no cortan la conectividad del esqueleto
    Call eliminaEsquinas
    'Elimina mucha de las ramificaciones secundarias
    Call eliminaArtefactos
    'elimina pequeños cortes en el esqueleto
    Call uneFlancos
    'Elimina mucha de las ramificaciones secundarias
    Call eliminaArtefactos
  
    
End Sub

Sub puntosTerminales(indicador As Integer)
 Dim x As Integer, y As Integer
 
 If indicador >= 0 Then
        clsTipos.varCaracteristicas(14, indicador) = 0
        clsTipos.varCaracteristicas(21, indicador) = 0
 Else
        clsTipos.varCaracteristicasImg(14) = 0
        clsTipos.varCaracteristicasImg(21) = 0
 End If
    
 'Calculo una propiedad numérica en base al número de puntos terminales
 For x = 1 To clsTipos.varImagen.ancho - 2
        For y = 1 To clsTipos.varImagen.Alto - 2
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And vecinos(x, y) = 1 Then
                'MsgBox "pntos terminales " + Str(x) + " " + Str(y)
                If indicador >= 0 Then
                   clsTipos.varCaracteristicas(14, indicador) = clsTipos.varCaracteristicas(14, indicador) + 0.1
                Else
                   clsTipos.varCaracteristicasImg(14) = clsTipos.varCaracteristicasImg(14) + 0.1
                End If
                'Ahora llamo a una función que calcula una característica en función
                'de la posicion relativa del punto terminal
                Call posicionRelativaPuntoTerminal(x, y, indicador)
            End If
        Next y
 Next x

End Sub

Private Sub posicionRelativaPuntoTerminal(x As Integer, y As Integer, indicador As Integer)
'La imagen la dividiré en cuatro cuadrantes, asignando un valor numerico en
'base a la posicion relativa del punto terminal
 Dim valor As Single
 
 valor = 0
    
'Si esta en el primer cuadrante
If (x <= clsTipos.varImagen.ancho / 2) And (y <= clsTipos.varImagen.Alto / 2) Then
    valor = 0.05
Else 'segundo cuadrante
    If (x > clsTipos.varImagen.ancho / 2) And (y <= clsTipos.varImagen.Alto / 2) Then
        valor = 0.1
    Else 'tercer cuadrante
        If (x <= clsTipos.varImagen.ancho / 2) And (y > clsTipos.varImagen.Alto / 2) Then
            valor = 0.15
        Else
            If (x > clsTipos.varImagen.ancho / 2) And (y > clsTipos.varImagen.Alto / 2) Then
                valor = 0.2
            End If
        End If
    End If
End If

 If indicador >= 0 Then
     clsTipos.varCaracteristicas(21, indicador) = clsTipos.varCaracteristicas(21, indicador) + valor
 Else
     clsTipos.varCaracteristicasImg(21) = clsTipos.varCaracteristicasImg(21) + valor
 End If
 
End Sub


'calcula el número de vecinos pertenecientes a la imagen del pixel
Private Function vecinos(x As Integer, y As Integer) As Integer
    Dim resul As Integer
    
    resul = 0
    If clsTipos.varImagen.MatrizGris(x - 1, y - 1) = 255 Then
        resul = resul + 1
    End If
    
    If clsTipos.varImagen.MatrizGris(x, y - 1) = 255 Then
        resul = resul + 1
    End If
    
    If clsTipos.varImagen.MatrizGris(x + 1, y - 1) = 255 Then
        resul = resul + 1
    End If
    
    If clsTipos.varImagen.MatrizGris(x + 1, y) = 255 Then
        resul = resul + 1
    End If
    
    If clsTipos.varImagen.MatrizGris(x + 1, y + 1) = 255 Then
        resul = resul + 1
    End If
    
    If clsTipos.varImagen.MatrizGris(x, y + 1) = 255 Then
        resul = resul + 1
    End If
    
    If clsTipos.varImagen.MatrizGris(x - 1, y + 1) = 255 Then
        resul = resul + 1
    End If
    
    If clsTipos.varImagen.MatrizGris(x - 1, y) = 255 Then
         resul = resul + 1
    End If
        
    vecinos = resul
End Function

'indica los cambios entre fondo e imagen
Private Function transiciones(x As Integer, y As Integer) As Integer
    Dim t As Integer
    
    If clsTipos.varImagen.MatrizGris(x, y - 1) = 0 And clsTipos.varImagen.MatrizGris(x + 1, y - 1) = 255 Then
        t = t + 1
    End If
    If clsTipos.varImagen.MatrizGris(x + 1, y - 1) = 0 And clsTipos.varImagen.MatrizGris(x + 1, y) = 255 Then
        t = t + 1
    End If
    If clsTipos.varImagen.MatrizGris(x + 1, y) = 0 And clsTipos.varImagen.MatrizGris(x + 1, y + 1) = 255 Then
        t = t + 1
    End If
    If clsTipos.varImagen.MatrizGris(x + 1, y + 1) = 0 And clsTipos.varImagen.MatrizGris(x, y + 1) = 255 Then
        t = t + 1
    End If
    If clsTipos.varImagen.MatrizGris(x, y + 1) = 0 And clsTipos.varImagen.MatrizGris(x - 1, y + 1) = 255 Then
        t = t + 1
    End If
    If clsTipos.varImagen.MatrizGris(x - 1, y + 1) = 0 And clsTipos.varImagen.MatrizGris(x - 1, y) = 255 Then
        t = t + 1
    End If
    If clsTipos.varImagen.MatrizGris(x - 1, y) = 0 And clsTipos.varImagen.MatrizGris(x - 1, y - 1) = 255 Then
        t = t + 1
    End If
    If clsTipos.varImagen.MatrizGris(x - 1, y - 1) = 0 And clsTipos.varImagen.MatrizGris(x, y - 1) = 255 Then
        t = t + 1
    End If
    
    transiciones = t
    
End Function

'procedimiento que elimina las ramificaciones secundarias que no pertenecen al esqueleto
Private Sub eliminaArtefactos()
   Dim x As Integer, y As Integer, ven As Integer

   For x = 1 To clsTipos.varImagen.ancho - 2
    For y = 1 To clsTipos.varImagen.Alto - 2
        If clsTipos.varImagen.MatrizGris(x, y) = 255 Then
            ven = vecinos(x, y)
            'es un pixel terminal
            If ven = 1 Then
                Call tratarRamificacion(x, y)
            End If
        End If
    Next y
  Next x

End Sub

'Establezco si la ramificación es de la propia imagen o es un artefac
Private Sub tratarRamificacion(a As Integer, b As Integer)
    Dim tamanyo As Integer
    Dim x As Integer, y As Integer
    Dim xant As Integer, yant As Integer
    Dim puntos() As Integer
    Dim ven As Integer
    Dim borrados As Integer
   
    x = a
    y = b
    xant = x
    yant = y
    
    'determino que si la ramificacion es inferior al 10% de la imagen
    'entonces se trata de una ramificación falsa del esqueleto
    For tamanyo = 0 To Round(clsTipos.varImagen.Alto * 0.1)
    
        Call proximo(xant, yant, x, y)
        
        ReDim Preserve puntos(tamanyo * 2 + 1)
        puntos(tamanyo * 2 + 0) = xant
        puntos(tamanyo * 2 + 1) = yant
        ven = vecinos(x, y)
        If ven > 2 Then
            borrados = 0
            While borrados < UBound(puntos)
                clsTipos.varImagen.MatrizGris(puntos(borrados), puntos(borrados + 1)) = 0
                borrados = borrados + 2
            Wend
            tamanyo = 100
        End If
        
    Next tamanyo

End Sub
'determina el proximo punto del esqueleto
Private Sub proximo(xant As Integer, yant As Integer, x As Integer, y As Integer)

    If clsTipos.varImagen.MatrizGris(x, y - 1) = 255 And (y - 1) <> yant Then
        xant = x
        yant = y
        y = y - 1
    Else
        If clsTipos.varImagen.MatrizGris(x + 1, y - 1) = 255 And ((x + 1) <> xant Or (y - 1) <> yant) Then
            xant = x
            yant = y
            x = x + 1
            y = y - 1
        Else
            If clsTipos.varImagen.MatrizGris(x + 1, y) = 255 And (x + 1) <> xant Then
                xant = x
                yant = y
                x = x + 1
            Else
                If clsTipos.varImagen.MatrizGris(x + 1, y + 1) = 255 And ((x + 1) <> xant Or (y + 1) <> yant) Then
                    xant = x
                    yant = y
                    x = x + 1
                    y = y + 1
                Else
                    If clsTipos.varImagen.MatrizGris(x, y + 1) = 255 And (y + 1) <> yant Then
                        xant = x
                        yant = y
                        y = y + 1
                    Else
                        If clsTipos.varImagen.MatrizGris(x - 1, y + 1) = 255 And ((x - 1) <> xant Or (y + 1) <> yant) Then
                            xant = x
                            yant = y
                            x = x - 1
                            y = y + 1
                        Else
                            If clsTipos.varImagen.MatrizGris(x - 1, y) = 255 And (x - 1) <> xant Then
                                xant = x
                                yant = y
                                x = x - 1
                            Else
                                If clsTipos.varImagen.MatrizGris(x - 1, y - 1) = 255 And ((x - 1) <> xant Or (y - 1) <> yant) Then
                                    xant = x
                                    yant = y
                                    x = x - 1
                                    y = y - 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
                  
End Sub
'elimina los pixeles que estan en las esquinas, de forma que solo habrá
'pixeles intermedios con 2 vecinos, y terminales con 1 vecino
Private Sub eliminaEsquinas()
    Dim x As Integer, y As Integer
    
    For x = 1 To clsTipos.varImagen.ancho - 2
        For y = 1 To clsTipos.varImagen.Alto - 2
            If clsTipos.varImagen.MatrizGris(x, y) = 255 Then
                If clsTipos.varImagen.MatrizGris(x + 1, y) = 255 And clsTipos.varImagen.MatrizGris(x, y + 1) = 255 Then
                    clsTipos.varImagen.MatrizGris(x, y) = 0
                Else
                    If clsTipos.varImagen.MatrizGris(x - 1, y) = 255 And clsTipos.varImagen.MatrizGris(x, y + 1) = 255 Then
                        clsTipos.varImagen.MatrizGris(x, y) = 0
                    Else
                        If clsTipos.varImagen.MatrizGris(x, y - 1) = 255 And clsTipos.varImagen.MatrizGris(x + 1, y) = 255 Then
                            clsTipos.varImagen.MatrizGris(x, y) = 0
                        Else
                            If clsTipos.varImagen.MatrizGris(x - 1, y) = 255 And clsTipos.varImagen.MatrizGris(x, y - 1) = 255 Then
                                clsTipos.varImagen.MatrizGris(x, y) = 0
                            End If
                        End If
                    End If
                End If
            End If
        Next y
    Next x

End Sub

'Elimina los cortes producidos en el esqueleto para lo cual hace cuatro recorridos
'siguiendo las diagonales en los dos sentidos
Private Sub uneFlancos()
    Dim x As Integer, y As Integer
    Dim distancia As Integer
    
    If clsTipos.varImagen.Alto < clsTipos.varImagen.ancho Then
        distancia = Round(clsTipos.varImagen.Alto * 0.1)
    Else
        distancia = Round(clsTipos.varImagen.ancho * 0.1)
    End If
    
    For x = 1 To clsTipos.varImagen.ancho - 1 - distancia
        For y = 1 To clsTipos.varImagen.Alto - 1 - distancia
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And vecinos(x, y) = 1 And clsTipos.varImagen.MatrizGris(x + 1, y) = 0 And clsTipos.varImagen.MatrizGris(x + 1, y + 1) = 0 And clsTipos.varImagen.MatrizGris(x, y + 1) = 0 Then
                Call uneFlancosEntorno(x, y, distancia)
            End If
        Next y
    Next x
    
    For x = clsTipos.varImagen.ancho - 2 To distancia + 1 Step -1
        For y = 1 To clsTipos.varImagen.Alto - 1 - distancia
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And vecinos(x, y) = 1 And clsTipos.varImagen.MatrizGris(x - 1, y) = 0 And clsTipos.varImagen.MatrizGris(x, y + 1) = 0 And clsTipos.varImagen.MatrizGris(x - 1, y + 1) = 0 Then
                Call uneFlancosEntorno2(x, y, distancia)
            End If
        Next y
    Next x
    
    For x = 1 To clsTipos.varImagen.ancho - 1 - distancia
        For y = clsTipos.varImagen.Alto - 2 To distancia + 1 Step -1
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And vecinos(x, y) = 1 And clsTipos.varImagen.MatrizGris(x, y - 1) = 0 And clsTipos.varImagen.MatrizGris(x + 1, y) = 0 And clsTipos.varImagen.MatrizGris(x + 1, y - 1) = 0 Then
                Call uneFlancosEntorno3(x, y, distancia)
            End If
        Next y
    Next x
    
    For x = clsTipos.varImagen.ancho - 2 To distancia + 1 Step -1
        For y = clsTipos.varImagen.Alto - 2 To distancia + 1 Step -1
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And vecinos(x, y) = 1 And clsTipos.varImagen.MatrizGris(x - 1, y) = 0 And clsTipos.varImagen.MatrizGris(x - 1, y - 1) = 0 And clsTipos.varImagen.MatrizGris(x, y - 1) = 0 Then
                Call uneFlancosEntorno4(x, y, distancia)
            End If
        Next y
    Next x
    
End Sub

'Une un pixel terminal al pixel del caracter más cercano
Private Sub uneFlancosEntorno(x As Integer, y As Integer, distancia As Integer)

Dim i As Integer, j As Integer, k As Integer
Dim incX As Double, incY As Double
Dim seguir As Boolean

seguir = True
While i < distancia And seguir
    j = 0
    While j < distancia And seguir
        
        If i = 0 And j = 0 Then
            
        Else
            If clsTipos.varImagen.MatrizGris(x + i, y + j) = 0 Then
            
            Else
               If i >= j Then
                   incX = j / i
                   k = 0
                   While k <= i
                       clsTipos.varImagen.MatrizGris(x + k, y + Round(incX * k + 0.5)) = 255
                       k = k + 1
                       seguir = False
                   Wend
               Else
                   incY = i / j
                   k = 0
                   While k <= j
                       clsTipos.varImagen.MatrizGris(x + Round(incY * k + 0.5), y + k) = 255
                       k = k + 1
                       seguir = False
                   Wend
               End If
            End If
        End If
        j = j + 1
    Wend
    i = i + 1
Wend

End Sub


Private Sub uneFlancosEntorno2(x As Integer, y As Integer, distancia As Integer)

Dim i As Integer, j As Integer, k As Integer
Dim incX As Double, incY As Double
Dim seguir As Boolean

seguir = True
While i < distancia And seguir
    j = 0
    While j < distancia And seguir
        
        If i = 0 And j = 0 Then
            
        Else
            If clsTipos.varImagen.MatrizGris(x - i, y + j) = 0 Then
            
            Else
                If i > j Then
                    incX = j / i
                    k = 0
                    While k <= i
                        clsTipos.varImagen.MatrizGris(x - k, y + Round(incX * k + 0.5)) = 255
                        k = k + 1
                        seguir = False
                    Wend
                Else
                    incY = i / j
                    k = 0
                    While k <= j
                        clsTipos.varImagen.MatrizGris(x - Round(incY * k + 0.5), y + k) = 255
                        k = k + 1
                        seguir = False
                    Wend
                End If
            End If
        End If
        j = j + 1
    Wend
    i = i + 1
Wend

End Sub

Private Sub uneFlancosEntorno3(x As Integer, y As Integer, distancia As Integer)

Dim i As Integer, j As Integer, k As Integer
Dim incX As Double, incY As Double
Dim seguir As Boolean

seguir = True
While i < distancia And seguir
    j = 0
    While j < distancia And seguir
        
        If i = 0 And j = 0 Then
            
        Else
            If clsTipos.varImagen.MatrizGris(x + i, y - j) = 0 Then
            
            Else
                If i > j Then
                    incX = j / i
                    k = 0
                    While k <= i
                        clsTipos.varImagen.MatrizGris(x + k, y - Round(incX * k + 0.5)) = 255
                        k = k + 1
                        seguir = False
                    Wend
                Else
                    incY = i / j
                    k = 0
                    While k <= j
                        clsTipos.varImagen.MatrizGris(x + Round(incY * k + 0.5), y - k) = 255
                        k = k + 1
                        seguir = False
                    Wend
                End If
            End If
        End If
        j = j + 1
    Wend
    i = i + 1
Wend

End Sub


Private Sub uneFlancosEntorno4(x As Integer, y As Integer, distancia As Integer)

Dim i As Integer, j As Integer, k As Integer
Dim incX As Double, incY As Double
Dim seguir As Boolean

seguir = True
While i < distancia And seguir
    j = 0
    While j < distancia And seguir
        
        If i = 0 And j = 0 Then
            
        Else
            If clsTipos.varImagen.MatrizGris(x - i, y - j) = 0 Then
            
            Else
                If i > j Then
                    incX = j / i
                    k = 0
                    While k <= i
                        clsTipos.varImagen.MatrizGris(x - k, y - Round(incX * k + 0.5)) = 255
                        k = k + 1
                        seguir = False
                    Wend
                Else
                    incY = i / j
                    k = 0
                    While k <= j
                        clsTipos.varImagen.MatrizGris(x - Round(incY * k + 0.5), y - k) = 255
                        k = k + 1
                        seguir = False
                    Wend
                End If
            End If
        End If
        j = j + 1
    Wend
    i = i + 1
Wend

End Sub



