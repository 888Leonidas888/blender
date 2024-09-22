VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type fields
    sku As String
    description As String
End Type

Private self As fields

Public Property Get sku() As String
    sku = self.sku
End Property

Public Property Let sku(ByVal value As String)
    self.sku = value
End Property

Public Property Get description() As String
    description = self.description
End Property

Public Property Let description(ByVal value As String)
    self.description = value
End Property

