VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const adOpenForwardOnly = 3
Const adUseClient = 3

Dim StrQuery             As String
Public Function Recuperar(recproductos As Variant) As Long
     
    StrQuery = "Select * from Productos"
  
    On Error GoTo Recuperar_Err
       GetData StrQuery, adOpenForwardOnly, adUseClient, recproductos
       Recuperar = 0&
       
       
Exit Function

Recuperar_Err:
    Recuperar = vbObjectError + 5
    Err.Raise vbObjectError + 5

End Function
Private Function GetData(ByVal StrQuery As String, ByVal Cursor As Integer, ByVal CurLoc As Integer, adoRecordset As Variant) As Long

On Error GoTo GetData_Err
    
   Set adoRecordset = New Recordset
        With adoRecordset
            .CursorType = Cursor
            .CursorLocation = CurLoc
            .Open StrQuery, DSN
        End With
   GetData = 0&
   Exit Function

GetData_Err:
   Set adoRecordset = Nothing
   GetData = vbObjectError + 2
   Err.Raise vbObjectError + 2

End Function
Public Function Alta_Movimiento(ByVal nombre As String, _
   ByVal ApellidoPat As String, ByVal ApellidoMat As String, _
   ByVal CalleN As String, ByVal Colonia As String, _
   ByVal Ciudad As String, ByVal DelegMunicip As String, _
   ByVal CP As String, ByVal Telefono As String, _
   ByVal Observaciones As String, ByVal RFC As String _
   ) As Long

    Dim RecID As Variant
    ''Dim ID As Long
    
    On Error GoTo Alta_Movimiento_Err
                      
       StrQuery = "Insert Into Clientes(Nombre,ApellidoPat,ApellidoMat,CalleN,Colonia,Ciudad,DelegMunicip,CP,Telefono,Observaciones,RFC) Values"
       StrQuery = StrQuery & "(" & "'" & Trim$(nombre) & "'"
       StrQuery = StrQuery & ",'" & Trim$(ApellidoPat) & "'"
       StrQuery = StrQuery & ",'" & Trim$(ApellidoMat) & "'"
       StrQuery = StrQuery & ",'" & Trim$(CalleN) & "'"
       StrQuery = StrQuery & ",'" & Trim$(Colonia) & "'"
       If Trim$(Ciudad) <> "" Then
          StrQuery = StrQuery & ",'" & Trim$(Ciudad) & "'"
       Else
          Ciudad = " "
          StrQuery = StrQuery & ",'" & Trim$(Ciudad) & "'"
       End If
       If Trim$(DelegMunicip) <> "" Then
          StrQuery = StrQuery & ",'" & Trim$(DelegMunicip) & "'"
       Else
          DelegMunicip = " "
          StrQuery = StrQuery & ",'" & Trim$(DelegMunicip) & "'"
       End If
       If Trim$(CP) <> "" Then
          StrQuery = StrQuery & ",'" & Trim$(CP) & "'"
       Else
          CP = " "
          StrQuery = StrQuery & ",'" & Trim$(CP) & "'"
       End If
       If Trim$(Telefono) <> "" Then
          StrQuery = StrQuery & ",'" & Trim$(Telefono) & "'"
       Else
          Telefono = " "
          StrQuery = StrQuery & ",'" & Trim$(Telefono) & "'"
       End If
       
       If Trim$(Observaciones) <> "" Then
          StrQuery = StrQuery & ",'" & Trim$(Observaciones) & "'"
       Else
          Observaciones = " "
          StrQuery = StrQuery & ",'" & Trim$(Observaciones) & "'"
       End If
       StrQuery = StrQuery & ",'" & Trim$(RFC) & "')"
       
       Execute StrQuery
       
       Alta_Movimiento = 0&
       
Exit Function

Alta_Movimiento_Err:
    Alta_Movimiento = vbObjectError + 1
    Err.Raise vbObjectError + 1
End Function
Public Function Execute(ByVal StrQuery As String) As Long

Dim Cnn As ADODB.Connection
 
On Error GoTo Execute_Err
    
   Set Cnn = New ADODB.Connection
        Cnn.Open DSN
        Cnn.Execute StrQuery
        
        Cnn.Close
    Set Cnn = Nothing
    Execute = 0&
    Exit Function

Execute_Err:
    
    Set Cnn = Nothing
    Execute = vbObjectError + 3
    Err.Raise vbObjectError + 3

End Function

Public Function RecuperarClientes(RecClientes As Variant) As Long
     
    StrQuery = "Select * from Clientes"
  
    On Error GoTo RecuperarClientes_Err
       GetData StrQuery, adOpenForwardOnly, adUseClient, RecClientes
       RecuperarClientes = 0&
       
       
Exit Function

RecuperarClientes_Err:
    RecuperarClientes = vbObjectError + 5
    Err.Raise vbObjectError + 5

End Function
Public Function Modificacion(ByVal idcliente As Integer, ByVal nombre As String, _
   ByVal ApellidoPat As String, ByVal ApellidoMat As String, _
   ByVal CalleN As String, ByVal Colonia As String, _
   ByVal Ciudad As String, ByVal DelegMunicip As String, _
   ByVal CP As String, ByVal Telefono As String, _
   ByVal Observaciones As String, ByVal RFC As String _
   ) As Long
    
    StrQuery = "Update Clientes Set "
    StrQuery = StrQuery & "Nombre='" & Trim$(nombre) & "',"
    StrQuery = StrQuery & "ApellidoPat='" & Trim$(ApellidoPat) & "',"
    StrQuery = StrQuery & "ApellidoMat='" & Trim$(ApellidoMat) & "',"
    StrQuery = StrQuery & "CalleN='" & Trim$(CalleN) & "',"
    StrQuery = StrQuery & "Colonia='" & Trim$(Colonia) & "',"
    StrQuery = StrQuery & "Ciudad='" & Trim$(Ciudad) & "',"
    StrQuery = StrQuery & "DelegMunicip='" & Trim$(DelegMunicip) & "',"
    StrQuery = StrQuery & "CP='" & Trim$(CP) & "',"
    StrQuery = StrQuery & "Telefono='" & Trim$(Telefono) & "',"
    StrQuery = StrQuery & "RFC='" & Trim$(RFC) & "',"
    StrQuery = StrQuery & "Observaciones='" & Trim$(Observaciones) & "'"
    StrQuery = StrQuery & "Where IDCliente=" & idcliente
    
    On Error GoTo Modificacion_Err
       Execute StrQuery
       Modificacion = 0&
       
              
Exit Function
  
Modificacion_Err:
    
    
    Modificacion = vbObjectError + 4
    Err.Raise vbObjectError + 4

End Function
Public Function RecuperarUsuarios(RecUsuarios As Variant, RecUsuariounico As Variant, Optional Usuario As String, _
                Optional Parcial As Boolean) As Long
    On Error GoTo RecuperarUsuarios_err
    
    If Parcial = True And Usuario <> "" Then
       StrQuery = "Select * from Usuarios " _
                & "where (SNombre = '" & Usuario & "')"
             ''& "where (idproducto = " & Idproducto & ")"
       GetData StrQuery, adOpenForwardOnly, adUseClient, RecUsuariounico
    Else
        StrQuery = "Select * from usuarios"
        GetData StrQuery, adOpenForwardOnly, adUseClient, RecUsuarios
    End If
    RecuperarUsuarios = 0&

Exit Function
RecuperarUsuarios_err:
  RecuperarUsuarios = vbObjectError + 5
  Err.Raise vbObjectError + 5
End Function
Public Function BorraUsuario(ByVal idusuario As Integer) As Long
    StrQuery = "Delete * from Usuarios where pnidusuario = " & idusuario & " "
    On Error GoTo BorraUsuario_Err
    Execute StrQuery
    BorraUsuario = 0&
Exit Function
BorraUsuario_Err:

    BorraUsuario = vbObjectError + 2
    Err.Raise vbObjectError + 2
End Function


Public Function RecuperarUnidades(ByVal Idproducto As Integer, RecUnidades As Variant) As Long
     
    StrQuery = "Select * from Unidades " _
             & "where (idproducto = " & Idproducto & ")"
    
    On Error GoTo RecuperarUnidades_Err
       GetData StrQuery, adOpenForwardOnly, adUseClient, RecUnidades
       RecuperarUnidades = 0&
       
       
Exit Function

RecuperarUnidades_Err:
    RecuperarUnidades = vbObjectError + 5
    Err.Raise vbObjectError + 5

End Function
Public Function Alta_Usuario(ByVal nombre As String, ByVal Contraseña As String, _
        Optional opciones As Integer, Optional idusuario As Integer) As Long
Dim recusua As Variant
Dim id As Integer
On Error GoTo Alta_Usuario_err
    If idusuario <> 0 Then
       StrQuery = "Update Usuarios set "
       StrQuery = StrQuery & "SNombre = '" & nombre & "', "
       StrQuery = StrQuery & "SLogin = '" & Contraseña & "', "
       StrQuery = StrQuery & "nopciones = " & opciones & " "
       StrQuery = StrQuery & "where pnidusuario = " & idusuario & ""
    Else
        StrQuery = "Select * from usuarios"
            GetData StrQuery, adOpenForwardOnly, adUseClient, recusua
        If recusua.RecordCount > 0 Then
           recusua.MoveLast
           id = recusua(0) + 1
        Else
           id = 1
        End If
        ''StrQuery = "Insert into Usuarios(PNIdUsuario, SNombre,SLogin,SOpciones)Values & ('" & id & "' & '" & Nombre & "' & '" & Contraseña & "' & '" & opciones & "')"
            StrQuery = "Insert into Usuarios(PNIdUsuario, SNombre,SLogin,NOpciones)Values"
            StrQuery = StrQuery & "(" & "" & id & ""
            StrQuery = StrQuery & ",'" & nombre & "'"
            StrQuery = StrQuery & ",'" & Contraseña & "'"
            StrQuery = StrQuery & "," & opciones & ")"
     End If
            Execute StrQuery
            Alta_Usuario = 0&
Exit Function
Alta_Usuario_err:
    Alta_Usuario = vbObjectError + 5
    Err.Raise vbObjectError + 5

End Function

Public Function Alta_Productos(ByVal NombrepRODUCTO As String, _
   ByVal uNIDADMEDIDA As String, ByVal Observaciones As Double, ByVal precio As Double) As Long

Dim RecID As Variant
    ''Dim ID As Long
    
    On Error GoTo Alta_Productos_Err
                      
       StrQuery = "Insert Into Productos(NombreProducto,UnidadMedida,Observaciones,Precio) Values"
       StrQuery = StrQuery & "(" & "'" & Trim$(NombrepRODUCTO) & "'"
       StrQuery = StrQuery & ",'" & Trim$(uNIDADMEDIDA) & "'"
       If Observaciones > 0 Then
          StrQuery = StrQuery & "," & Trim$(Observaciones) & ""
       Else
          Observaciones = "1"
          StrQuery = StrQuery & "," & Trim$(Observaciones) & ""
       End If
       StrQuery = StrQuery & "," & CDbl(precio) & ")"
       Execute StrQuery
       
       Alta_Productos = 0&
       
Exit Function

Alta_Productos_Err:
     Alta_Productos = vbObjectError + 1
    Err.Raise vbObjectError + 1
End Function

Public Function ModificacionProd(ByVal Idproducto As Integer, ByVal NombrepRODUCTO As String, _
   ByVal uNIDADMEDIDA As String, ByVal Observaciones As Double, ByVal precio As Double _
   ) As Long
    
    StrQuery = "Update pRODUCTOS Set "
    StrQuery = StrQuery & "NombrePRODUCTO='" & Trim$(NombrepRODUCTO) & "',"
    StrQuery = StrQuery & "uNIDADMEDIDA='" & Trim$(uNIDADMEDIDA) & "',"
    StrQuery = StrQuery & "Observaciones= " & Trim$(Observaciones) & ","
    StrQuery = StrQuery & "PRECIO= " & CDbl(precio) & " "
    StrQuery = StrQuery & "Where IDpRODUCTO=" & Idproducto
    
    On Error GoTo ModificacionProd_Err
       Execute StrQuery
       ModificacionProd = 0&
       
              
Exit Function
  
ModificacionProd_Err:
    
    
    ModificacionProd = vbObjectError + 4
    Err.Raise vbObjectError + 4

End Function
Public Function Baja(ByVal Idproducto As Integer) As Long
    Dim i As Integer

    On Error GoTo Baja_Err

    StrQuery = "Delete from productos where IDProducto = " & CInt(Idproducto) & ""

    Execute StrQuery
'    StrQuery = "Delete from CodResultExamen where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    StrQuery = "Delete from CodResultListaTrab where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    StrQuery = "Delete from CodigosResultado where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    objContext.SetComplete
    Baja = 0&

Exit Function

Baja_Err:
    
    Baja = vbObjectError + 2
    Err.Raise vbObjectError + 2
End Function

Public Function BajaCliente(ByVal idcliente As Integer) As Long
    Dim i As Integer

    On Error GoTo BajaCliente_Err

    StrQuery = "Delete from clientes where IDCliente = " & CInt(idcliente) & ""

    Execute StrQuery
'    StrQuery = "Delete from CodResultExamen where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    StrQuery = "Delete from CodResultListaTrab where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    StrQuery = "Delete from CodigosResultado where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    objContext.SetComplete
    BajaCliente = 0&

Exit Function

BajaCliente_Err:
    
    BajaCliente = vbObjectError + 2
    Err.Raise vbObjectError + 2
End Function
Public Function ALta_Venta(ByVal idcliente As String, _
   ByVal FechaVenta As String, ByVal arreglo As Variant, ByVal numelem As Integer) As Long

    Dim RecID As Variant
    Dim i As Integer
    On Error GoTo Alta_Venta_Err
                      
    For i = 0 To numelem - 1
        StrQuery = "Insert Into Ventas (idproducto,montosiniva,montoconiva,unidadprod,idcliente,Costounitario,FechaVenta,Conversion,Metros) Values"
        StrQuery = StrQuery & "(" & "'" & Trim$(arreglo(i, 0)) & "'"
       StrQuery = StrQuery & ",'" & Trim$(arreglo(i, 4)) & "'"
       StrQuery = StrQuery & ",'" & Trim$(arreglo(i, 5)) & "'"
       StrQuery = StrQuery & ",'" & Trim$(arreglo(i, 3)) & "'"
       StrQuery = StrQuery & ",'" & Trim$(idcliente) & "'"
       StrQuery = StrQuery & ",'" & Trim$(arreglo(i, 2)) & "'"
       StrQuery = StrQuery & ",'" & FechaVenta & "'"
       StrQuery = StrQuery & ",'" & Trim$(arreglo(i, 6)) & "'"
       StrQuery = StrQuery & ",'" & Trim$(arreglo(i, 7)) & "')"
        Execute StrQuery
    Next i
        ' Update Inventario set InventarioActual = (InventarioActual - unidadprod)
                     
                     
                     
     '  StrQuery = "Insert Into Ventas (idproducto,montosiniva,montoconiva,unidadprod,idcliente,fechaventa,Costounitario) Values"
      ' StrQuery = StrQuery & "(" & "'" & Trim$(idproducto) & "'"
      ' StrQuery = StrQuery & ",'" & Trim$(montosiniva) & "'"
      ' StrQuery = StrQuery & ",'" & Trim$(montoconiva) & "'"
      ' StrQuery = StrQuery & ",'" & Trim$(unidadprod) & "'"
      ' StrQuery = StrQuery & ",'" & Trim$(idcliente) & "'"
      ' StrQuery = StrQuery & ",'" & Trim$(Fechaventa) & "'"
      ' StrQuery = StrQuery & ",'" & Trim$(Costounitario) & ""
       
      '' Execute StrQuery, DSN
       
       ALta_Venta = 0&
       
Exit Function

Alta_Venta_Err:
    ALta_Venta = vbObjectError + 1
    Err.Raise vbObjectError + 1
End Function
Public Function RecuperarVentas(RecVentas As Variant) As Long
     
    StrQuery = "Select * from Ventas"
            
    
    On Error GoTo RecuperarVentas_Err
       GetData StrQuery, adOpenForwardOnly, adUseClient, RecVentas
       RecuperarVentas = 0&
       
       
Exit Function

RecuperarVentas_Err:
    RecuperarVentas = vbObjectError + 5
    Err.Raise vbObjectError + 5

End Function
Public Function RecuperarVentasTodo(RecVentas As Variant, _
                             ByVal Fechain As String, ByVal fechafin As String) As Long
     
    StrQuery = "Select * from Ventas "
    StrQuery = StrQuery & "where (Fechaventa >= " & "#" & Format(Fechain, "mm/dd/yyyy") & "#" & ") "
    StrQuery = StrQuery & "and (Fechaventa <= " & "#" & Format(fechafin, "mm/dd/yyyy") & "#" & ")"
        
    
    On Error GoTo RecuperarVentastodo_Err
       GetData StrQuery, adOpenForwardOnly, adUseClient, RecVentas
       RecuperarVentasTodo = 0&
       
       
Exit Function

RecuperarVentastodo_Err:
    RecuperarVentasTodo = vbObjectError + 5
    Err.Raise vbObjectError + 5

End Function
Public Function RecuperarVentasFecha(ByVal Idproducto As String, _
                ByVal Fechain As String, ByVal fechafin As String, _
                RecVtasFecha As Variant) As Long
     
    StrQuery = "Select * from Ventas "
    StrQuery = StrQuery & "where ((Fechaventa >= " & "#" & Format(Fechain, "mm/dd/yyyy") & "#" & ") "
    StrQuery = StrQuery & "and (Fechaventa <= " & "#" & Format(fechafin, "mm/dd/yyyy") & "#" & "))"
    StrQuery = StrQuery & "and idproducto = '" & Idproducto & "'"
            
    On Error GoTo RecuperarVentasFecha_Err
       GetData StrQuery, adOpenForwardOnly, adUseClient, RecVtasFecha
       RecuperarVentasFecha = 0&
       
       
Exit Function

RecuperarVentasFecha_Err:
    RecuperarVentasFecha = vbObjectError + 5
    Err.Raise vbObjectError + 5

End Function



Public Function Alta_Invent(ByVal Idproducto As Integer, _
   ByVal InventarioActual As Integer, ByVal Fecha As String, _
   ByVal UnidadesADescontar As Integer) As Long

Dim RecID As Variant
    ''Dim ID As Long
    
    On Error GoTo Alta_Invent_Err
                      
       StrQuery = "Insert Into Inventario(idProducto,InventarioActual,Fecha,UnidadesADescontar) Values"
       StrQuery = StrQuery & "(" & "'" & CInt(Idproducto) & "'"
       StrQuery = StrQuery & ",'" & CInt(InventarioActual) & "'"
       StrQuery = StrQuery & ",'" & Trim$(Fecha) & "'"
       StrQuery = StrQuery & "," & CInt(UnidadesADescontar) & ")"
       Execute StrQuery
       
       Alta_Invent = 0&
       
Exit Function

Alta_Invent_Err:
    Alta_Invent = vbObjectError + 1
    Err.Raise vbObjectError + 1
End Function
Public Function RecuperarInvent(recinvent As Variant) As Long
     
    StrQuery = "Select * from Inventario"
            
    
    On Error GoTo RecuperarInvent_Err
       GetData StrQuery, adOpenForwardOnly, adUseClient, recinvent
       RecuperarInvent = 0&
       
       
Exit Function

RecuperarInvent_Err:
    RecuperarInvent = vbObjectError + 5
    Err.Raise vbObjectError + 5

End Function
Public Function BajaTodoVenta() As Long
    Dim i As Integer

    On Error GoTo BajaTodoVenta_Err

    StrQuery = "Delete * from Ventas"

    Execute StrQuery
'    StrQuery = "Delete from CodResultExamen where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    StrQuery = "Delete from CodResultListaTrab where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    StrQuery = "Delete from CodigosResultado where IDCodigoResultado = '" & Trim$(strCampo) & "'"
'    objDatos.Execute SQLStmt, DSN
'    objContext.SetComplete
    BajaTodoVenta = 0&

Exit Function

BajaTodoVenta_Err:
    
    BajaTodoVenta = vbObjectError + 2
    Err.Raise vbObjectError + 2
End Function
