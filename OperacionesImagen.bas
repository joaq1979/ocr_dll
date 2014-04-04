Attribute VB_Name = "OperacionesImagen"
Option Explicit

'Obtiene el tamaño de la imagen y lo almacena en los campos de TipoImagen
'además redimensiona la matriz de valores para grises al tamaño de la imagen
Sub obtenerDimensionImagen(imagen As PictureBox)
    
    clsTipos.varImagen.ancho = imagen.ScaleX(imagen.Picture.Width, vbHimetric, vbPixels)
    clsTipos.varImagen.Alto = imagen.ScaleY(imagen.Picture.Height, vbHimetric, vbPixels)
   
    ReDim clsTipos.varImagen.MatrizGris(clsTipos.varImagen.ancho - 1, clsTipos.varImagen.Alto - 1)
    ReDim clsTipos.varImagen.Contorno(clsTipos.varImagen.ancho - 1, clsTipos.varImagen.Alto - 1)
End Sub

'se le pasa un color y lo descompone en el gris de luminancia adecuado, despues
'hace una binarización en funcion del nivel de gris
Private Function colorToGris(ByVal color As Long) As Double
    Dim red As Long, green As Long, blue As Long, temp As Long
    
    blue = (color \ &H10000) And &HFF
    green = (color \ &H100) And &HFF
    red = color And &HFF
    
    temp = 0.299 * red + 0.587 * green + 0.114 * blue
    
    colorToGris = temp
End Function

Private Function binarizar(pixel As Double, umbral As Integer) As Integer
    'binarizo a fondo negro y caracter blanco
    If pixel > umbral Then
        binarizar = 0
    Else
        binarizar = 255
    End If
End Function

Private Function obtenerUmbral() As Integer
    Dim x As Integer, y As Integer
    Dim histo(255) As Integer
    Dim max1 As Integer, max2 As Integer
    Dim ind1 As Integer, ind2 As Integer
    
    'calcula el histograma
    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            histo(clsTipos.varImagen.MatrizGris(x, y)) = histo(clsTipos.varImagen.MatrizGris(x, y)) + 1
        Next y
    Next x

    max1 = 0
    ind1 = 0
    
    For x = 0 To 255
        If histo(x) > max1 Then
            max1 = histo(x)
            ind1 = x
        End If
    Next x

    For x = 0 To 255
        If histo(x) > max2 And Abs(x - ind1) > 50 Then
            max2 = histo(x)
            ind2 = x
        End If
    Next x
      
    obtenerUmbral = (ind1 + ind2) \ 2

End Function


'Obtengo una matriz con la imagen binarizada y normalizada
Sub obtenerMatrizImagen(imagen As PictureBox)
    Dim x As Single, y As Single
    Dim factorEscala As Integer
    Dim umbral As Integer
    
    'se pasa a escala de grises
    For x = 0 To clsTipos.varImagen.ancho - 1
      For y = 0 To clsTipos.varImagen.Alto - 1
        clsTipos.varImagen.MatrizGris(x, y) = colorToGris(imagen.Point(x, y))
      Next y
    Next x
    
    'se obtiene el umbral de binarización automáticamente
    umbral = obtenerUmbral()
    
    'se binariza
    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            clsTipos.varImagen.MatrizGris(x, y) = binarizar(clsTipos.varImagen.MatrizGris(x, y), umbral)
        Next y
    Next x
    
    'se elimina el ruido de sal y pimienta
    Call eliminaRuido
    
    'se normaliza el tamaño de las imagenes
    factorEscala = 100 \ clsTipos.varImagen.Alto
    
    If factorEscala <> 1 And factorEscala > 1 Then
        Call crearMatrizNormalizada(factorEscala)
    End If
    
    Call limpiaBorde
End Sub
'elimina el ruido de sal y pimienta
Private Sub eliminaRuido()
 
 Dim y As Integer, x As Integer
 
    For x = 1 To clsTipos.varImagen.ancho - 2
        For y = 1 To clsTipos.varImagen.Alto - 2
        
            If (clsTipos.varImagen.MatrizGris(x, y) = 255) Then
                If (clsTipos.varImagen.MatrizGris(x - 1, y) = 0 And clsTipos.varImagen.MatrizGris(x - 1, y - 1) = 0 And _
                clsTipos.varImagen.MatrizGris(x, y - 1) = 0 And clsTipos.varImagen.MatrizGris(x + 1, y - 1) = 0 And _
                clsTipos.varImagen.MatrizGris(x + 1, y) = 0 And clsTipos.varImagen.MatrizGris(x + 1, y + 1) = 0 And _
                clsTipos.varImagen.MatrizGris(x, y + 1) = 0 And clsTipos.varImagen.MatrizGris(x - 1, y + 1) = 0) Then
         
                    clsTipos.varImagen.MatrizGris(x, y) = 0
                End If
            End If
        Next y
    Next x
End Sub

'Normaliza el tamaño de la imagen a una escala aproximada de 100x100
Private Sub crearMatrizNormalizada(factorEscala As Integer)
    Dim aux() As Double
    Dim x As Integer, y As Integer
    Dim newAncho As Integer, newAlto As Integer
    Dim i As Integer, j As Integer
       
    'creo la nueva matriz de pixeles con el tamaño normalizado
    newAncho = clsTipos.varImagen.ancho * factorEscala
    newAlto = clsTipos.varImagen.Alto * factorEscala
    ReDim aux(newAncho, newAlto)
    
    
    'marco los pixeles originales en la nueva posición en la matriz normalizada
    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            aux(((x + 1) * factorEscala) - 1, ((y + 1) * factorEscala) - 1) = clsTipos.varImagen.MatrizGris(x, y)
            If aux(((x + 1) * factorEscala) - 1, ((y + 1) * factorEscala) - 1) = 255 Then
                For i = ((x + 1) * factorEscala) - 1 To ((x + 1) * factorEscala) - 1 - factorEscala Step -1
                    For j = ((y + 1) * factorEscala) - 1 To ((y + 1) * factorEscala) - 1 - factorEscala Step -1
                        aux(i, j) = 255
                    Next j
                Next i
            End If
        Next y
    Next x

    'pongo los datos de la variable auxiliar en la variable de la imagen
    ReDim clsTipos.varImagen.MatrizGris(newAncho, newAlto)
    ReDim clsTipos.varImagen.Contorno(newAncho, newAlto)
    clsTipos.varImagen.Alto = newAlto
    clsTipos.varImagen.ancho = newAncho
    
    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            clsTipos.varImagen.MatrizGris(x, y) = aux(x, y)
        Next y
    Next x

End Sub

'limpiamos los bordes externos de la imagen para que correspondan al fondo
Private Sub limpiaBorde()
Dim x As Integer, y As Integer

    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To 3
            clsTipos.varImagen.MatrizGris(x, y) = 0
        Next y
        For y = clsTipos.varImagen.Alto - 4 To clsTipos.varImagen.Alto - 1
            clsTipos.varImagen.MatrizGris(x, y) = 0
        Next y
    Next x
    
    For y = 0 To clsTipos.varImagen.Alto - 1
        For x = 0 To 3
            clsTipos.varImagen.MatrizGris(x, y) = 0
        Next x
        For x = clsTipos.varImagen.ancho - 4 To clsTipos.varImagen.ancho - 1
            clsTipos.varImagen.MatrizGris(x, y) = 0
        Next x
    Next y
End Sub


'hago mas grueso el trazo de las imagenes
Sub engordarImagen()
    Dim x As Single, y As Single
    
    ' recorridos horizontales
    For y = 0 To clsTipos.varImagen.Alto - 1
        x = 0
        While x < clsTipos.varImagen.ancho
            
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And x > 2 Then
                clsTipos.varImagen.MatrizGris(x - 1, y) = 255
                clsTipos.varImagen.MatrizGris(x - 2, y) = 255
              
                While clsTipos.varImagen.MatrizGris(x, y) = 255 And x < clsTipos.varImagen.ancho - 1
                    x = x + 1
                Wend
            Else
                x = x + 1
            End If
        Wend
    Next y
    
    ' recorridos verticales
    For x = 0 To clsTipos.varImagen.ancho - 1
        y = 0
        While y < clsTipos.varImagen.Alto
            
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And y > 2 Then
                clsTipos.varImagen.MatrizGris(x, y - 1) = 255
                clsTipos.varImagen.MatrizGris(x, y - 2) = 255
             
                While clsTipos.varImagen.MatrizGris(x, y) = 255 And y < clsTipos.varImagen.Alto - 1
                    y = y + 1
                Wend
            Else
                y = y + 1
            End If
        Wend
    Next x
    
End Sub

Sub contornoImagen()
    Dim x As Integer, y As Integer
    Dim blanco As Boolean

    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            clsTipos.varImagen.Contorno(x, y) = 0
        Next y
    Next x
    
    'recorrido horizontal
    For y = 0 To clsTipos.varImagen.Alto - 1
        x = 0
        While x < clsTipos.varImagen.ancho
            If clsTipos.varImagen.MatrizGris(x, y) = 255 Then
                clsTipos.varImagen.Contorno(x, y) = 255
                x = x + 1
                blanco = False
                While ((x < clsTipos.varImagen.ancho) And (blanco = False))
                    If clsTipos.varImagen.MatrizGris(x, y) = 255 Then
                        x = x + 1
                    Else
                        blanco = True
                    End If
                Wend
                
                clsTipos.varImagen.Contorno(x - 1, y) = 255
            End If
            x = x + 1
        Wend
    Next y
                
    For x = 0 To clsTipos.varImagen.ancho - 1
        y = 0
        While y < clsTipos.varImagen.Alto
            If clsTipos.varImagen.MatrizGris(x, y) = 255 Then
                clsTipos.varImagen.Contorno(x, y) = 255
                y = y + 1
                blanco = False
                While ((y < clsTipos.varImagen.Alto) And (blanco = False))
                    If clsTipos.varImagen.MatrizGris(x, y) = 255 Then
                        y = y + 1
                    Else
                         blanco = True
                    End If
                Wend
                clsTipos.varImagen.Contorno(x, y - 1) = 255
            End If
            y = y + 1
        Wend
    Next x
                       
    Call eliminaEsquinas
End Sub

Private Sub eliminaEsquinas()
    Dim x As Integer, y As Integer
    
    'elimino las esquinas del contorno, asi cada pixel tendra como máximo 2 vecinos
    For x = 1 To clsTipos.varImagen.ancho - 2
        For y = 1 To clsTipos.varImagen.Alto - 2
            If clsTipos.varImagen.Contorno(x, y) = 255 Then
                If clsTipos.varImagen.Contorno(x + 1, y) = 255 And clsTipos.varImagen.Contorno(x, y + 1) = 255 Then
                    clsTipos.varImagen.Contorno(x, y) = 0
                Else
                    If clsTipos.varImagen.Contorno(x - 1, y) = 255 And clsTipos.varImagen.Contorno(x, y + 1) = 255 Then
                        clsTipos.varImagen.Contorno(x, y) = 0
                    Else
                        If clsTipos.varImagen.Contorno(x, y - 1) = 255 And clsTipos.varImagen.Contorno(x + 1, y) = 255 Then
                            clsTipos.varImagen.Contorno(x, y) = 0
                        Else
                            If clsTipos.varImagen.Contorno(x - 1, y) = 255 And clsTipos.varImagen.Contorno(x, y - 1) = 255 Then
                                clsTipos.varImagen.Contorno(x, y) = 0
                            End If
                        End If
                    End If
                End If
            End If
        Next y
    Next x

End Sub
'calcula la relacion del alto por el ancho
Sub altoAncho(indicador As Integer)
    Dim xmin As Integer, xmax As Integer
    Dim ymin As Integer, ymax As Integer
        
    xmin = clsTipos.varImagen.ancho
    ymin = clsTipos.varImagen.Alto
    xmax = 0
    ymax = 0
    
    Call puntosMin(xmin, ymin)
    Call puntosMax(xmax, ymax)
   
    If indicador >= 0 Then
        clsTipos.varCaracteristicas(10, indicador) = (xmax - xmin) / ((ymax - ymin) * 6)
    Else
        clsTipos.varCaracteristicasImg(10) = (xmax - xmin) / ((ymax - ymin) * 6)
     End If
End Sub

'calcula la relacción del número de transiciones horizontales entre la de verticales
'en la mitad del caracter

Sub transicionesMedias(indicador As Integer)
  Dim xmin As Integer, xmax As Integer
    Dim ymin As Integer, ymax As Integer
    Dim xmedio As Integer, ymedio As Integer
    Dim y As Integer, x As Integer
    Dim transiciones As Integer
        
    xmin = clsTipos.varImagen.ancho
    ymin = clsTipos.varImagen.Alto
    xmax = 0
    ymax = 0
    
    Call puntosMin(xmin, ymin)
    Call puntosMax(xmax, ymax)
    
    xmedio = ((xmax - xmin) / 2) + xmin
    ymedio = ((ymax - ymin) / 2) + ymin
    
    transiciones = 0
    'recorrido vertical
    For y = 0 To clsTipos.varImagen.Alto - 1
        If clsTipos.varImagen.MatrizGris(xmedio, y) = 255 Then
            transiciones = transiciones + 1
            
            While clsTipos.varImagen.MatrizGris(xmedio, y) = 255
                y = y + 1
            Wend
        End If
    Next y
    
    'MsgBox "transicones verticales" + Str(transiciones)
    If indicador >= 0 Then
        clsTipos.varCaracteristicas(12, indicador) = transiciones * 0.1
    Else
        clsTipos.varCaracteristicasImg(12) = transiciones * 0.1
    End If

    
    transiciones = 0
    'recorrido horizontal
    For x = 0 To clsTipos.varImagen.ancho - 1
        If clsTipos.varImagen.MatrizGris(x, ymedio) = 255 Then
            transiciones = transiciones + 1
            
            While clsTipos.varImagen.MatrizGris(x, ymedio) = 255
                x = x + 1
            Wend
        End If
    Next x
    
    'MsgBox "transicones horizontales" + Str(transiciones)
    If indicador >= 0 Then
        clsTipos.varCaracteristicas(13, indicador) = transiciones * 0.1
    Else
        clsTipos.varCaracteristicasImg(13) = transiciones * 0.1
    End If

    
End Sub


'establezco los puntos minimos del caracter en la imagen
Private Sub puntosMin(xmin As Integer, ymin As Integer)
 Dim x As Integer, y As Integer
 'hallo los puntos minimos
    For y = 0 To clsTipos.varImagen.Alto - 1
        For x = 0 To clsTipos.varImagen.ancho - 1
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And x < xmin Then
                xmin = x
            End If
        Next x
    Next y
    
    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And y < ymin Then
                ymin = y
            End If
        Next y
    Next x
    
End Sub

Private Sub puntosMax(xmax As Integer, ymax As Integer)
  Dim x As Integer, y As Integer
   'hallo los puntos maximos
    For y = 0 To clsTipos.varImagen.Alto - 1
        For x = 0 To clsTipos.varImagen.ancho - 1
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And x > xmax Then
                xmax = x
            End If
        Next x
    Next y
    
    For x = 0 To clsTipos.varImagen.ancho - 1
        For y = 0 To clsTipos.varImagen.Alto - 1
            If clsTipos.varImagen.MatrizGris(x, y) = 255 And y > ymax Then
                ymax = y
            End If
        Next y
    Next x
    
End Sub


