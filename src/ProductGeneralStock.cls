VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductGeneralStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private newProduct As ProductBase
Private addShitmpentLine As ProductShipmentLine

Private Type fields
    centro As String
    LPN As String
    amount As Integer
    ubication As String
    vto As Date
End Type

Private self As fields

Private Sub Class_Initialize()
    Set newProduct = New ProductBase
    Set addShitmpentLine = New ProductShipmentLine
End Sub

Public Property Get sku() As String
    sku = newProduct.sku
End Property

Public Property Let sku(ByVal value As String)
    newProduct.sku = value
End Property

Public Property Get description() As String
    description = newProduct.description
End Property

Public Property Let description(ByVal value As String)
    newProduct.description = value
End Property

Public Property Get centro() As String
    centro = self.centro
End Property

Public Property Let centro(ByVal value As String)
    self.centro = value
End Property

Public Property Get LPN() As String
    LPN = self.LPN
End Property

Public Property Let LPN(ByVal value As String)
    self.LPN = value
End Property
Public Property Get amount() As String
    amount = self.amount
End Property

Public Property Let amount(ByVal value As String)
    self.amount = value
End Property
Public Property Get ubication() As String
    ubication = self.ubication
End Property

Public Property Let ubication(ByVal value As String)
    self.ubication = value
End Property

Public Property Let vto(ByVal value As Date)
    self.vto = value
End Property
Public Property Get vto() As Date
    vto = self.vto
End Property

Public Property Get channel() As String
    channel = addShitmpentLine.channel
End Property

Public Property Let channel(ByVal value As String)
    addShitmpentLine.channel = value
End Property

Public Property Get total() As Integer
    total = addShitmpentLine.total
End Property

Public Property Let total(ByVal value As Integer)
    addShitmpentLine.total = value
End Property
