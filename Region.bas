Attribute VB_Name = "Region"
'Se establecen los procedimientos necesarios para determinar el número de zonas
'separadas pertenecientes al fondo de la imagen, con ello conseguimos saber el número
'de agujeros del caracter, pues será igual al número de zonas del fondo -1

Sub regionesFondo(indicador As Integer)
Dim x As Integer, y As Integer
Dim nueva_etiqueta As Integer
Dim etiqueta() As Integer
Dim cambio As Boolean
Dim valor As Integer
Dim regiones As Integer
Dim arrayRegiones() As Integer


ReDim etiqueta(clsTipos.varImagen.ancho, clsTipos.varImagen.Alto)
nueva_etiqueta = 0
cambio = True
regiones = 0

'rellenamos las etiquetas del borde
For x = 0 To clsTipos.varImagen.ancho - 1
    etiqueta(x, 0) = 1
    etiqueta(x, clsTipos.varImagen.Alto - 1) = 1
Next x
For y = 0 To clsTipos.varImagen.Alto - 1
    etiqueta(0, y) = 1
    etiqueta(clsTipos.varImagen.ancho - 1, y) = 1
Next y


'inicialización de cada pixel del fondo con una etiqueta
For x = 1 To clsTipos.varImagen.ancho - 2
    For y = 1 To clsTipos.varImagen.Alto - 2
        If clsTipos.varImagen.MatrizGris(x, y) = 0 Then
            etiqueta(x, y) = etiquetar(nueva_etiqueta)
            regiones = regiones + 1
        Else
            etiqueta(x, y) = 255
        End If
    Next y
Next x

'repetir hasta que ningun pixel se etiquete
While cambio = True
    cambio = False
    'recorrido de arriba a abajo y de izquierda a derecha
    For x = 1 To clsTipos.varImagen.ancho - 2
        For y = 1 To clsTipos.varImagen.Alto - 2
            If etiqueta(x, y) <> 255 Then
                 valor = minVecinosUno(etiqueta(x - 1, y), etiqueta(x, y - 1))
                 If valor <> -1 And valor <> etiqueta(x, y) Then
                    etiqueta(x, y) = valor
                    regiones = regiones - 1
                    cambio = True
                 End If
            End If
         Next y
    Next x
                         
    'recorrido de abajo a arriba y de izquierda a derecha
    For x = clsTipos.varImagen.ancho - 2 To 1 Step -1
        For y = clsTipos.varImagen.Alto - 2 To 1 Step -1
            If etiqueta(x, y) <> 255 Then
                 valor = minVecinosUno(etiqueta(x + 1, y), etiqueta(x, y + 1))
                 If valor <> -1 And valor <> etiqueta(x, y) Then
                    etiqueta(x, y) = valor
                    regiones = regiones - 1
                    cambio = True
                 End If
            End If
         Next y
    Next x
                       
    ReDim arrayRegiones(0)
    arrayRegiones(0) = etiqueta(0, 0)
    regiones = 0
    For x = 1 To clsTipos.varImagen.ancho - 1
        For y = 1 To clsTipos.varImagen.Alto - 1
           If noEsta(arrayRegiones, etiqueta(x, y)) Then
            regiones = regiones + 1
            ReDim Preserve arrayRegiones(regiones)
            arrayRegiones(regiones) = etiqueta(x, y)
           End If
        Next y
    Next x
    
Wend
    
If indicador >= 0 Then
    clsTipos.varCaracteristicas(11, indicador) = (UBound(arrayRegiones) - 1) * 0.15
Else
    clsTipos.varCaracteristicasImg(11) = (UBound(arrayRegiones) - 1) * 0.15
End If
                    
End Sub

'obtiene el siguiente valor de las etiquetas
Private Function etiquetar(nueva_etiqueta) As Integer

     If nueva_etiqueta = 254 Then
          nueva_etiqueta = nueva_etiqueta + 2
     Else
          nueva_etiqueta = nueva_etiqueta + 1
     End If
        
     etiquetar = nueva_etiqueta
End Function
'devuelve la etiqueta menor del vecino, si no pertenece al caracter
Private Function minVecinosUno(izq As Integer, arr As Integer) As Integer

    If izq <> 255 And arr <> 255 Then
        If izq < arr Then
            minVecinosUno = izq
        Else
            minVecinosUno = arr
        End If
    Else
        If izq = 255 And arr <> 255 Then
            minVecinosUno = arr
        Else
            If izq <> 255 And arr = 255 Then
                minVecinosUno = izq
            Else
                minVecinosUno = -1
            End If
        End If
    End If

End Function

Private Function noEsta(vector() As Integer, elemento As Integer) As Boolean
Dim i As Integer
noEsta = True

    For i = 0 To UBound(vector)
        If vector(i) = elemento Then
            noEsta = False
        End If
    Next i
End Function

