Attribute VB_Name = "clsTipos"
'Datos sobre las imagenes cargadas
Type tipoImagen
    Alto As Integer
    ancho As Integer
    MatrizGris() As Double
    Contorno() As Double
    CodigoCadena() As Integer
End Type

Public varImagen As tipoImagen

'Para almacenar el tipo de cada una de las imagenes
Public varNombres() As String

'Para almacenar los momentos de cada una de las imagenes de la BD
Public varCaracteristicas() As Double

'Para obtener los momentos medios por cada clase de imagen de la BD
Public varCaracteristicasMedias() As Integer

'para la imagen a clasificar
Public varCaracteristicasImg(22) As Double
