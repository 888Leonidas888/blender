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
    
    Dim shipmentInStock As Collection
    Dim listProductByPicking As Collection
    
    sqlStock = ProcedureStore.skusUniqueOfStock()
'    ddbbStock = "Huach-Prod-Lindley.GeneralStockList-2024-09-13-23.04.xlsx"
    Set pStockGeneral = productsUniqueInStock(sqlStock, ddbbStock)
    
    sqlShipment = ProcedureStore.skuByChannelByFreshnessByTotal
'    ddbbShipment = "Huach-Prod-Lindley.ShipmentLine-2024-09-13-23.05.xlsx"
    Set pShipmentLine = productsAllInShipment(sqlShipment, ddbbShipment)
    
    Set shipmentInStock = IntersectCollections(pStockGeneral, pShipmentLine)

    Set listProductByPicking = getListPiking(shipmentInStock, ddbbStock)
    
    Call printListProducts(listProductByPicking)

End Sub

