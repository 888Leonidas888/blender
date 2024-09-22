Attribute VB_Name = "ProcedureStore"
' Este m�dulo proporciona las consultas SQL personalizadas, todas las consultas son solo de lectura.
' Estas consultas estas ajustadas para conectarse a arhivos Excel.

Function skusUniqueOfStock() As String
    ' Descripci�n:
    '   Esta funci�n genera una consulta SQL que selecciona todos los SKUs �nicos, junto con su descripci�n y centro,
    '   desde una hoja de c�lculo en un rango espec�fico [datos$A2:T600], y los ordena por SKU.
    '
    ' Par�metros:
    '   Ninguno.
    '
    ' Retorno:
    '   String - Una cadena de texto que contiene la consulta SQL para obtener los SKUs �nicos.
    '
    ' Notas:
    '   - La consulta est� dise�ada para extraer informaci�n de un rango en una hoja de c�lculo de Excel denominada '[datos$A2:T600]'.
    '   - Por qu� desde [datos$A2:T600]?; los datos empiezan en A2 y se estiman hasta T600 pudiendo ser menos.
    '   - Se seleccionan los campos `sku`, `descripci�n` y `centro`, y los resultados se ordenan por `sku`.
     
    Dim sql As String
    
    sql = "SELECT DISTINCT sku, descripci�n,centro " & _
          "FROM [datos$A2:T600] " & _
          "ORDER BY sku"
          
    skusUniqueOfStock = sql
    
End Function
Function skuByChannelByFreshnessByTotal() As String
    ' Descripci�n:
    '   Esta funci�n genera una consulta SQL que selecciona SKUs, descripciones y canales,
    '   junto con la suma de la cantidad esperada agrupada por SKU, descripci�n y canal.
    '   La consulta filtra los resultados para incluir solo ciertos canales.
    '
    ' Par�metros:
    '   Ninguno.
    '
    ' Retorno:
    '   String - Una cadena de texto que contiene la consulta SQL para obtener SKUs por canal y frescura total.
    '
    ' Notas:
    '   - La consulta est� dise�ada para extraer informaci�n de una hoja de c�lculo denominada '[datos$]'.
    '   - Se filtran los resultados para incluir solo los canales: 'cuentas claves', 'detallista' y 'cruceristas'.
    '   - Los resultados se agrupan por SKU, descripci�n y canal, y se ordenan por SKU.
    
    Dim sql As String
    
    sql = "SELECT sku, descripci�n, canal, SUM([cantidad esperada]) AS total " & _
          "FROM " & _
          "(SELECT * " & _
          "FROM [datos$] " & _
          "WHERE canal = 'cuentas claves' " & _
          "OR canal = 'detallista' " & _
          "OR canal = 'cruceristas') " & _
          "AS sale_by_channel " & _
          "GROUP BY sku, descripci�n, canal " & _
          "ORDER BY sku;"
        
    skuByChannelByFreshnessByTotal = sql
    
End Function
Function skuByPicking(sku As String, freshness As String) As String
    ' Descripci�n:
    '   Esta funci�n genera una consulta SQL que selecciona todos los registros de una hoja de c�lculo
    '   donde el SKU y la frescura coinciden con los valores proporcionados. Adem�s, filtra los resultados
    '   para incluir solo aquellas ubicaciones que comienzan con ciertos prefijos.
    '
    ' Par�metros:
    '   sku (String): El SKU que se desea filtrar en la consulta.
    '   freshness (String): La frescura que se desea filtrar en la consulta.
    '
    ' Retorno:
    '   String - Una cadena de texto que contiene la consulta SQL para obtener datos filtrados por SKU y frescura.
    '
    ' Notas:
    '   - La consulta est� dise�ada para extraer informaci�n de una hoja de c�lculo denominada '[datos$A2:T600]'.
    '   - Se filtran los resultados para incluir solo aquellos registros cuya ubicaci�n comience con 'SE', 'DI' o 'PS'.
    '   - Los resultados se ordenan por la columna 'cantidad'.
    
    Dim sql As String
    
    sql = "SELECT * " & _
          "FROM [datos$A2:T600] " & _
          "WHERE sku = '" & sku & "' " & _
          "AND frescura = '" & freshness & "' " & _
          "AND (ubicaci�n LIKE 'SE%' OR ubicaci�n LIKE 'DI%' OR ubicaci�n LIKE 'PS%') " & _
          "ORDER BY cantidad;"

    skuByPicking = sql
    
End Function

