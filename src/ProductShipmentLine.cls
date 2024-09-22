VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductShipmentLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private newProduct As ProductBase

Private Type fields
    channel As String
    total As Integer
End Type

Private self As fields

Private Sub Class_Initialize()
    Set newProduct = New ProductBase
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

Public Property Get channel() As String
    channel = self.channel
End Property

Public Property Let channel(ByVal value As String)
    self.channel = value
End Property

Public Property Get total() As Integer
    total = self.total
End Property

Public Property Let total(ByVal value As Integer)
    self.total = value
End Property
