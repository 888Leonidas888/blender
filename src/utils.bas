Attribute VB_Name = "utils"
' Este módulo de utildades de apoyo para las operaciones de este proyecto

Private Const CRUCERISTA As String = "cruceristas"
Private Const DETALLISTA As String = "detallista"
Private Const CUENTAS_CLAVES As String = "cuentas claves"

Private Const BAJA As String = "baja"
Private Const MEDIA As String = "media"
Private Const CUENTAS As String = "cuentas"

Function getFeshness(channel As String) As String
    ' Descripción:
    '   Se encarga de recuperar la frescura recibiendo el canal de despacho.
    '
    ' Parámetros:
    '   channel(string): el canal al cual pertenece el producto.
    '
    ' Retorno:
    '   string(baja | media | cuentas | sin asignar): la frescura que corresponde al canal recibido.
    
    Select Case LCase(channel)
        Case Is = CRUCERISTA
            getFeshness = BAJA
        Case Is = DETALLISTA
            getFeshness = MEDIA
        Case Is = CUENTAS_CLAVES
            getFeshness = CUENTAS
        Case Else
            getFeshness = "sin asignar"
    End Select
End Function
Function getIDFreshness(freshness As String) As Integer
    ' Descripción:
    '   Se encarga de recuperar el ID de la frescura
    '
    ' Parámetros:
    '   freshness(string): la frescura del producto.
    '
    ' Retorno:
    '   integer(1 | 2 | 3): ID correspondiente a la frescura.
    Select Case LCase(freshness)
        Case Is = BAJA
            getIDFreshness = 1
        Case Is = MEDIA
            getIDFreshness = 2
        Case Is = CUENTAS
            getIDFreshness = 3
    End Select
End Function
Function toString(value As Variant) As String
    ' Descripción:
    '   envuelve cualquier valor entre apóstrofe, función de apoyo
    '   para incluirla de una consulta SQL junto con Storage.
    '
    ' Parámetros:
    '   value(variant): el valor que se quiere envolver con apóstrofes.
    '
    ' Retorno:
    '   string: valor envuelto entre apóstrofes.
    
    toString = "'" & value & "'"
End Function
Public Function converterRGBForVBA(hexColor As String) As Variant
    
    Dim R As String
    Dim G As String
    Dim B As String
    
    If Len(hexColor) <> 7 Then: Err.Raise Number:=2000, description:="Len invalid"
    If Left(hexColor, 1) <> "#" Then: Err.Raise Number:=2001, description:="Firts character invalid"
    
    hexColor = Right(hexColor, Len(hexColor) - 1)
    
    R = Left(hexColor, 2)
    G = Mid(hexColor, 3, 2)
    B = Right(hexColor, 2)
    
    converterRGBForVBA = "&H" & B & G & R
    
End Function
Function selectFileXlsx() As String
    
    Dim fDialog As FileDialog
    Dim selectedFile As String
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fDialog
        .Title = "Seleccionar archivo .xlsx"
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xlsx"
        .AllowMultiSelect = False
        .InitialFileName = Environ$("userprofile") '"C:\"
        
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            selectedFile = Empty
        End If
    End With
    
    selectFileXlsx = selectedFile
    
End Function
Sub addNewSheet()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim count As Integer
    Dim sheetExists As Boolean
    
    count = 1
    sheetExists = True
    
    Do While sheetExists
        sheetName = "licuad" & count
        sheetExists = False
        
        For Each ws In ThisWorkbook.Sheets
            If ws.name = sheetName Then
                sheetExists = True
                Exit For
            End If
        Next ws
        
        If sheetExists Then
            count = count + 1
        End If
    Loop
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.name = sheetName
    
End Sub


