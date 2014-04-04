Attribute VB_Name = "clsBaseDeDatos"
 
 Public Sub conectaBD(nombre As String)
    frmDemo.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + nombre
    frmDemo.Adodc1.CursorType = adOpenDynamic
    frmDemo.Adodc1.RecordSource = "Fotos"
    frmDemo.Adodc1.Refresh
 End Sub
 
Public Sub enlazarContenido()

   Set frmDemo.lblNombre.DataSource = frmDemo.Adodc1
   frmDemo.lblNombre.DataField = "Nombre"
    
   Set frmDemo.image1.DataSource = frmDemo.Adodc1
   frmDemo.image1.DataField = "foto"
    
End Sub
