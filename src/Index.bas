Attribute VB_Name = "Index"
'Punto de entrada a nuestra proyecto

Sub main(ddbbStock As String, ddbbShipment As String)
    
    'stock general in Excel
    Dim sqlStock As String
'    Dim ddbbStock As String
    Dim pStockGeneral As Collection
    
    'shipment line in Excel
    Dim sqlShipment As String
'    Dim ddbbShipment As String
    Dim pShipmentLine As Collection
    
    ' shipment line in stock
    Dim shipmentInStock As Collection
    
    Dim product As New ProductShipmentLine
    
    'freshness to work
    Dim freshnessToWork As New Collection
    
    'Connection Excel stock
    Dim st As New Storage
    Dim rs As ADODB.Recordset
    Dim params As Dictionary
    Dim listProductByPicking As Collection
    
    sqlStock = ProcedureStore.skusUniqueOfStock()
'    ddbbStock = "Huach-Prod-Lindley.GeneralStockList-2024-09-13-23.04.xlsx"
    Set pStockGeneral = productsUniqueInStock(sqlStock, ddbbStock)
    
    sqlShipment = ProcedureStore.skuByChannelByFreshnessByTotal
'    ddbbShipment = "Huach-Prod-Lindley.ShipmentLine-2024-09-13-23.05.xlsx"
    Set pShipmentLine = productsAllInShipment(sqlShipment, ddbbShipment)
    
    Set shipmentInStock = IntersectCollections(pStockGeneral, pShipmentLine)

    Set listProductByPicking = getListPiking(shipmentInStock, ddbbStock)
    
    i = 1
    Cells(i, 1).value = "SKU"
    Cells(i, 2).value = "DESCRIPCION"
    Cells(i, 3).value = "LPN"
    Cells(i, 4).value = "UBICACIÓN"
    Cells(i, 5).value = "CANTIDAD"
    Cells(i, 6).value = "CANAL"
    Cells(i, 7).value = "TOTAL_POR_CANAL"

    i = 2
    For Each p In listProductByPicking
        Cells(i, 1).value = p.sku
        Cells(i, 2).value = p.description
        Cells(i, 3).value = "'" & p.LPN
        Cells(i, 4).value = p.ubication
        Cells(i, 5).value = p.amount
        Cells(i, 6).value = p.channel
        Cells(i, 7).value = p.total
        i = i + 1
    Next p
    
End Sub

