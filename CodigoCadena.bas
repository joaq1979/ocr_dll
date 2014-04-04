Attribute VB_Name = "clsCodigoCadena"

Sub obtenerCadena(indicador As Integer)

    Dim Xini As Integer, Yini As Integer
    Dim x As Integer, y As Integer
    Dim elem As Integer
    Dim cierto As Boolean
    Dim Contador As Integer
    Dim xant As Integer, yant As Integer
    
    cierto = True
    'Obtengo el primer punto del contorno en Xini Yini
    For y = clsTipos.varImagen.Alto - 1 To 0 Step -1
        For x = clsTipos.varImagen.ancho - 1 To 0 Step -1
            If clsTipos.varImagen.Contorno(x, y) = 255 And cierto Then
                Xini = x
                Yini = y
                cierto = False
            End If
        Next x
    Next y
        
    Contador = 0
    x = Xini
    y = Yini
    
    elem = siguientePunto(0, 0, x, y)
    ReDim clsTipos.varImagen.CodigoCadena(Contador)
    clsTipos.varImagen.CodigoCadena(Contador) = elem
    
     xant = Xini
     yant = Yini
      
    'recorro todos los pixeles del contorno hasta llegar a los iniciales
    'y almaceno el codigo cadena de cada pixel
    Do
        Contador = Contador + 1
        elem = siguientePunto(xant, yant, x, y)
        ReDim Preserve clsTipos.varImagen.CodigoCadena(Contador)
        clsTipos.varImagen.CodigoCadena(Contador) = elem
    
    Loop Until ((x = Xini) And (y = Yini))
    
    'calculo el promedio de tipos de elementos del contorno cadena
    Call promediosCadena(Contador, indicador)
    
    
End Sub


Private Function siguientePunto(xant As Integer, yant As Integer, x As Integer, y As Integer)

    If clsTipos.varImagen.Contorno(x, y - 1) = 255 And (y - 1) <> yant Then
        siguientePunto = 3
        xant = x
        yant = y
        y = y - 1
    Else
        If clsTipos.varImagen.Contorno(x + 1, y - 1) = 255 And ((x + 1) <> xant Or (y - 1) <> yant) Then
            siguientePunto = 2
            xant = x
            yant = y
            x = x + 1
            y = y - 1
        Else
            If clsTipos.varImagen.Contorno(x + 1, y) = 255 And (x + 1) <> xant Then
                siguientePunto = 1
                xant = x
                yant = y
                x = x + 1
            Else
                If clsTipos.varImagen.Contorno(x + 1, y + 1) = 255 And ((x + 1) <> xant Or (y + 1) <> yant) Then
                    siguientePunto = 8
                    xant = x
                    yant = y
                    x = x + 1
                    y = y + 1
                Else
                    If clsTipos.varImagen.Contorno(x, y + 1) = 255 And (y + 1) <> yant Then
                        siguientePunto = 7
                        xant = x
                        yant = y
                        y = y + 1
                    Else
                        If clsTipos.varImagen.Contorno(x - 1, y + 1) = 255 And ((x - 1) <> xant Or (y + 1) <> yant) Then
                            siguientePunto = 6
                            xant = x
                            yant = y
                            x = x - 1
                            y = y + 1
                        Else
                            If clsTipos.varImagen.Contorno(x - 1, y) = 255 And (x - 1) <> xant Then
                                siguientePunto = 5
                                xant = x
                                yant = y
                                x = x - 1
                            Else
                                If clsTipos.varImagen.Contorno(x - 1, y - 1) = 255 And ((x - 1) <> xant Or (y - 1) <> yant) Then
                                    siguientePunto = 4
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
                  
End Function

'calculo los promedios del codigo cadena y los almaceno en las tablas de caracteristicas
'de las imagenes
Private Sub promediosCadena(totalpixel As Integer, indicador As Integer)
    Dim promediouno As Double, promediodos As Double
    Dim promediotres As Double, promediocuatro As Double
    Dim promediocinco As Double, promedioseis As Double
    Dim promediosiete As Double, promedioocho As Double
    Dim i As Integer
    
    promediouno = 0
    promediodos = 0
    promediotres = 0
    promediocuatro = 0
    promediocinco = 0
    promedioseis = 0
    promediosiete = 0
    promedioocho = 0
    
    For i = 0 To UBound(clsTipos.varImagen.CodigoCadena)
        Select Case clsTipos.varImagen.CodigoCadena(i)
        Case 1
            promediouno = promediouno + 1
        Case 2
            promediodos = promediodos + 1
        Case 3
            promediotres = promediotres + 1
        Case 4
            promediocuatro = promediocuatro + 1
        Case 5
            promediocinco = promediocinco + 1
        Case 6
            promedioseis = promedioseis + 1
        Case 7
            promediosiete = promediosiete + 1
        Case 8
            promedioocho = promedioocho + 1
        End Select
    Next i
        
    'se trata de una imagen de la B.D
    If indicador >= 0 Then
       clsTipos.varCaracteristicas(1, indicador) = promediouno / totalpixel
       clsTipos.varCaracteristicas(2, indicador) = promediodos / totalpixel
       clsTipos.varCaracteristicas(3, indicador) = promediotres / totalpixel
       clsTipos.varCaracteristicas(4, indicador) = promediocuatro / totalpixel
       clsTipos.varCaracteristicas(5, indicador) = promediocinco / totalpixel
       clsTipos.varCaracteristicas(6, indicador) = promedioseis / totalpixel
       clsTipos.varCaracteristicas(7, indicador) = promediosiete / totalpixel
       clsTipos.varCaracteristicas(8, indicador) = promedioocho / totalpixel
    Else 'se trata de la imagen a comparar
       clsTipos.varCaracteristicasImg(1) = promediouno / totalpixel
       clsTipos.varCaracteristicasImg(2) = promediodos / totalpixel
       clsTipos.varCaracteristicasImg(3) = promediotres / totalpixel
       clsTipos.varCaracteristicasImg(4) = promediocuatro / totalpixel
       clsTipos.varCaracteristicasImg(5) = promediocinco / totalpixel
       clsTipos.varCaracteristicasImg(6) = promedioseis / totalpixel
       clsTipos.varCaracteristicasImg(7) = promediosiete / totalpixel
       clsTipos.varCaracteristicasImg(8) = promedioocho / totalpixel
    End If
End Sub
