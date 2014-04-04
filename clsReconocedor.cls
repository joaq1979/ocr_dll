VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReconocedor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub cargarImgBD(suNombre As String)
   Dim i As Integer
   
   'enlaza con la base de datos
   conectaBD (suNombre)
   Call enlazarContenido
   
   'redimensiona las tablas para almacenar los datos
   ReDim clsTipos.varNombres(1, frmDemo.Adodc1.Recordset.RecordCount - 1)
   ReDim clsTipos.varCaracteristicas(22, frmDemo.Adodc1.Recordset.RecordCount - 1)
   
   i = 0
   Do
      'obtiene los datos de la imagen
      Call obtenerDimensionImagen(frmDemo.image1)
      Call obtenerMatrizImagen(frmDemo.image1)
     
     'cargo los nombres de las imagenes, que también las identifica
     clsTipos.varNombres(0, i) = frmDemo.lblNombre.Caption
     'marco los nombres de las imagenes con "N" indicando elemento No tratado
     clsTipos.varNombres(1, i) = "N"
       
       
     'debo meter el esqueleto
     Call esqueletoImagen
     
     Call puntosTerminales(i)
     'relaccion alto-ancho
     Call altoAncho(i)
     Call transicionesMedias(i)
         
       
     'engorda la imagen al menos en 2 pixeles mas
     Call engordarImagen
     'obtengo el contorno de la imagen
     Call contornoImagen
        
     'obtener el codigo cadena
     Call obtenerCadena(i)
    
     Call engordarImagen
     Call engordarImagen
     Call regionesFondo(i)
     Call momentosHu(i)
     
     i = i + 1
     frmDemo.Adodc1.Recordset.MoveNext
   Loop Until frmDemo.Adodc1.Recordset.EOF
   
   
End Sub


Public Function CompararImag(nombre As String) As String
  Dim i As Integer
  Dim clases As Integer
  Dim elemProximo As Double
  Dim distanciaMinima As Double
  
  clases = 0
  frmDemo.image2.Picture = LoadPicture(nombre)
  Call obtenerDimensionImagen(frmDemo.image2)
  Call obtenerMatrizImagen(frmDemo.image2)
                                   
  'debo meter el esqueleto
  Call esqueletoImagen
  
  Call puntosTerminales(-1)
  'obtengo la relacción alto-ancho
  Call altoAncho(-1)
  Call transicionesMedias(-1)
  Call engordarImagen
  Call contornoImagen
        
  'obtengo el codigo cadena de las imagenes
  Call obtenerCadena(-1)
  
  Call engordarImagen
  Call engordarImagen
  Call regionesFondo(-1)
  Call momentosHu(-1)
  
   
  'Calculo las caracteristicas medias de cada clase
  For i = 0 To frmDemo.Adodc1.Recordset.RecordCount - 1
    
    If clsTipos.varNombres(1, i) = "N" Then
        'Lo marco como tratado
        clsTipos.varNombres(1, i) = "T"
        'LLamo a la funcion que obtiene la media aritmetica de los momentos
        Call calculaMediaCaracteristicas(i, frmDemo.Adodc1.Recordset.RecordCount - 1)
        'En varMomentosMedios almaceno los indices de la tabla varMomentos en
        'donde he almacenado las medias de las clases de las imagenes
        ReDim Preserve clsTipos.varCaracteristicasMedias(clases)
        clsTipos.varCaracteristicasMedias(clases) = i
        
        'Print Tipos.varMomentosMedios(clases)
        clases = clases + 1
     End If
  Next i

  'realiza la clasificacion
  elemProximo = clasifica(distanciaMinima)
  CompararImag = clsTipos.varNombres(0, elemProximo)
  
End Function

