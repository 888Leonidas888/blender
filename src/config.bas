Attribute VB_Name = "config"
'Este m�dulo proporciona la configuraci�n para generar la cadena de conexi�n
'para la instancia de Storage


Function configConnection(pathDDBB As String) As ADODB.Connection
    Dim config As New ADODB.Connection
    
    With config
        .provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Data Source") = pathDDBB
'        .Properties("Data Source") = ThisWorkbook.Path & "\" & pathDDBB
        .Properties("Extended Properties") = "Excel 12.0 Xml;HDR=YES"
    End With
    
    Set configConnection = config
End Function
