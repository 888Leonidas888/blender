Attribute VB_Name = "managerProducts"
' Este m�dulo se encarga de generar las colecciones con las instancias
' de clase apartir del acceso a los archivos Excel:
'   - ProductGeneralStock
'   - ProductShipmentLine
' Tambi�n hacer el cruce datos de ambas colleciones.


Function productsInStock2(rs As ADODB.Recordset) As Collection
    ' Descripci�n:
    '   Esta funci�n toma un ADODB.Recordset como entrada y recorre sus filas,
    '   creando una colecci�n de objetos ProductGeneralStock para cada registro en el Recordset.
    '   Llena los atributos de cada objeto ProductGeneralStock con datos de los campos
    '   del Recordset y devuelve una colecci�n de estos objetos.
    '
    ' Par�metros:
    '   rs (ADODB.Recordset) - El recordset que contiene la informaci�n de los productos a procesar.
    '
    ' Retorno:
    '   Collection - Una colecci�n de objetos ProductGeneralStock que representan los productos en stock.

    Dim products As New Collection
    Dim p As ProductGeneralStock
    
    With rs
        If Not (.EOF) And Not (.BOF) Then
            .MoveFirst
            Do While Not (.EOF)
                Set p = New ProductGeneralStock
                p.sku = .fields("sku").value
                p.centro = .fields("centro").value
                p.description = .fields("descripci�n").value
                p.LPN = .fields("LPN").value
                p.amount = .fields("cantidad").value
                p.ubication = .fields("ubicaci�n").value
                products.Add p
                .MoveNext
            Loop
        End If
    End With
    
    Set productsInStock2 = products

End Function
Function productsInStock(rs As ADODB.Recordset) As Collection
    ' Descripci�n:
    '   Esta funci�n toma un ADODB.Recordset como entrada y recorre sus filas,
    '   creando una colecci�n de objetos ProductGeneralStock para cada registro en el Recordset.
    '   Llena los atributos de cada objeto ProductGeneralStock con los datos de los campos
    '   del Recordset ("sku", "centro" y "descripci�n") y devuelve una colecci�n de estos objetos.
    '
    ' Par�metros:
    '   rs (ADODB.Recordset) - El recordset que contiene la informaci�n de los productos a procesar.
    '
    ' Retorno:
    '   Collection - Una colecci�n de objetos ProductGeneralStock que representan los productos en stock.
    
    Dim products As New Collection
    Dim p As ProductGeneralStock
    
    With rs
        If Not (.EOF) And Not (.BOF) Then
            .MoveFirst
            Do While Not (.EOF)
                Set p = New ProductGeneralStock
                p.sku = .fields("sku").value
                p.centro = .fields("centro").value
                p.description = .fields("descripci�n").value
                products.Add p
                .MoveNext
            Loop
        End If
    End With
    
    Set productsInStock = products

End Function
Function productsInShipment(rs As ADODB.Recordset) As Collection
    ' Descripci�n:
    '   Esta funci�n toma un ADODB.Recordset como entrada y recorre sus filas,
    '   creando una colecci�n de objetos ProductShipmentLine. Llena los atributos
    '   de cada objeto ProductShipmentLine con los datos de los campos del Recordset
    '   ("sku", "canal", "descripci�n" y "total") y devuelve una colecci�n de estos objetos.
    '
    ' Par�metros:
    '   rs (ADODB.Recordset) - El recordset que contiene la informaci�n de los productos para el env�o.
    '
    ' Retorno:
    '   Collection - Una colecci�n de objetos ProductShipmentLine que representan los productos para el env�o.
    
    Dim products As New Collection
    Dim p As ProductShipmentLine
    
    With rs
        If Not (.EOF) And Not (.BOF) Then
            .MoveFirst
            Do While Not (.EOF)
                Set p = New ProductShipmentLine
                p.sku = .fields("sku").value
                p.channel = .fields("canal").value
                p.description = .fields("descripci�n").value
                p.total = .fields("total").value
                products.Add p
                .MoveNext
            Loop
        End If
    End With
    
    Set productsInShipment = products
    
End Function
Function IntersectCollections(p1 As Collection, p2 As Collection) As Collection
    ' Descripci�n:
    '   Esta funci�n toma dos colecciones de objetos como entrada y devuelve una nueva colecci�n
    '   con los elementos comunes entre ambas, bas�ndose en el atributo `sku` de cada objeto.
    '   Los objetos de la segunda colecci�n que tengan un `sku` que tambi�n est� presente en la primera
    '   colecci�n ser�n agregados a la colecci�n resultante.
    '
    ' Par�metros:
    '   p1 (Collection) - La primera colecci�n que contiene objetos con el atributo `sku`.
    '   p2 (Collection) - La segunda colecci�n que contiene objetos con el atributo `sku`.
    '
    ' Retorno:
    '   Collection - Una colecci�n de los elementos de la segunda colecci�n (`p2`) cuyo `sku`
    '   est� presente en la primera colecci�n (`p1`).

    Dim conjUnique As New Collection
    Dim dict As New Dictionary

    For Each conj1 In p1
        If Not dict.Exists(conj1.sku) Then
            dict.Add conj1.sku, True
        End If
    Next conj1
    
    For Each conj2 In p2
        If dict.Exists(conj2.sku) Then
            conjUnique.Add conj2
        End If
    Next conj2

    Set IntersectCollections = conjUnique
    
End Function
Function productsUniqueInStock(sql As String, ddbb As String) As Collection
    ' Descripci�n:
    '   Esta funci�n ejecuta una consulta SQL personalizada en una base de datos especificada, obtiene
    '   un conjunto de productos del inventario y devuelve una colecci�n con los productos resultantes
    '   que son �nicos. Utiliza la clase `Storage` para gestionar la conexi�n y consulta a la base de datos.
    '
    ' Par�metros:
    '   sql (String) - La consulta SQL personalizada que se desea ejecutar.
    '   ddbb (String) - El nombre o la ruta de la base de datos a la cual conectarse.
    '
    ' Retorno:
    '   Collection - Una colecci�n que contiene los objetos de tipo `ProductGeneralStock` devueltos por
    '   la consulta SQL.
    
    Dim st As New Storage
    Dim rs As ADODB.Recordset
    Dim products As Collection
    Dim p As New ProductGeneralStock

    With st
        .connect configConnection(ddbb)
        Set rs = .customQuery(sql)
        Set products = productsInStock(rs)
        .disconnect
    End With
    
    Set productsUniqueInStock = products
    
End Function
Function productsAllInShipment(sql As String, ddbb As String) As Collection
    ' Descripci�n:
    '   Esta funci�n ejecuta una consulta SQL personalizada en una base de datos especificada, obtiene
    '   un conjunto de productos para los env�os y devuelve una colecci�n con todos los productos encontrados.
    '   Utiliza la clase `Storage` para gestionar la conexi�n y consulta a la base de datos.
    '
    ' Par�metros:
    '   sql (String) - La consulta SQL personalizada que se desea ejecutar.
    '   ddbb (String) - El nombre o la ruta de la base de datos a la cual conectarse.
    '
    ' Retorno:
    '   Collection - Una colecci�n que contiene los objetos de tipo `ProductShipmentLine` devueltos por
    '   la consulta SQL.

    Dim st As New Storage
    Dim rs As ADODB.Recordset
    Dim products As Collection
    Dim p As New ProductShipmentLine

    With st
        .connect configConnection(ddbb)
        Set rs = .customQuery(sql)
        Set products = productsInShipment(rs)
        .disconnect
    End With
    
    Set productsAllInShipment = products
    
End Function
Function getListPiking(shipmentInStock As Collection, ddbbStock As String) As Collection
    ' Descripci�n:
    '   Esta funci�n genera una lista de productos para realizar el picking (recolecci�n) de un env�o
    '   en funci�n del stock disponible en la base de datos. Eval�a diferentes niveles de frescura
    '   del producto y selecciona productos hasta alcanzar el total requerido para el env�o.
    '
    ' Par�metros:
    '   shipmentInStock (Collection) - Una colecci�n de productos del env�o, que incluye SKU, canal, descripci�n y cantidad.
    '   ddbbStock (String) - El nombre o la ruta de la base de datos donde se encuentra el stock de productos.
    '
    ' Retorno:
    '   Collection - Una colecci�n de productos que cumplen con los requisitos de picking, basada en los productos en stock.
    '
    ' Notas:
    '   - La funci�n eval�a los niveles de frescura del stock (baja, media, cuentas) y busca
    '     productos en la base de datos del stock que coincidan con el SKU y la frescura.
    '   - Se conecta a la base de datos de stock utilizando la clase `Storage` y ejecuta consultas SQL personalizadas.
    '   - Los productos seleccionados se almacenan en una colecci�n de productos de tipo `ProductGeneralStock` para el picking.
    '
    Dim total  As Integer
    Dim totalFount As Integer
    Dim completed As Boolean
    Dim freshness As String
    Dim IDFreshness As Integer
    Dim sku As String
    Dim tmpProduct As Collection
    Dim listProductByPicking As New Collection
    Dim p As ProductGeneralStock
    Dim LPNs As New Dictionary
    Dim freshnessToWork As New Collection
    Dim product As New ProductShipmentLine
    Dim st As Storage
    Dim rs As ADODB.Recordset
    
    With freshnessToWork
        .Add "baja"
        .Add "media"
        .Add "cuentas"
    End With

    For Each product In shipmentInStock

        total = product.total
        totalFount = 0
        freshness = getFeshness(product.channel)
        IDFreshness = getIDFreshness(freshness)
        sku = product.sku
        completed = False

        For i = IDFreshness To freshnessToWork.count
            Set st = New Storage

            With st
                .connect configConnection(ddbbStock)
                Set rs = .customQuery(skuByPicking(sku, freshnessToWork.Item(i)))
                Set tmpProduct = productsInStock2(rs)
                .disconnect
            End With
            
            If tmpProduct.count > 0 Then
                For Each p In tmpProduct
                    If totalFount <= total Then
                        totalFount = totalFount + p.amount
                        If Not LPNs.Exists(p.LPN) Then
                            LPNs.Add p.LPN, True
                            p.channel = product.channel
                            p.total = product.total
                            listProductByPicking.Add p
                        End If
                    Else
                        completed = True
                        Exit For
                    End If
                Next p
            End If
            
            If completed Then Exit For
        Next i
    Next product
    
    Set getListPiking = listProductByPicking
       
'    i = 1
'    For Each p In listProductByPicking
'        Cells(i, 1).value = p.sku
'        Cells(i, 2).value = p.description
'        Cells(i, 3).value = "'" & p.LPN
'        Cells(i, 4).value = p.ubication
'        Cells(i, 5).value = p.amount
'        i = i + 1
'    Next p
End Function
