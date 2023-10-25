VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event elementAdded(ByVal element_Id As String, ByVal element_Code As String, ByVal element_Name As String, ByVal element_price As Double)



Private Type TItem
    BoxPrice As Double
    BoxQty As Integer
    Code As String
    Cost As Double
    Description As String
    Id As Integer
    IsNew As Boolean
    Name As String
    Qty As Integer
    SellingPrice As Double
    Supplier As String
    WholeSalePrice As Double
End Type

Private this As TItem

Private Sub class_initialize()
    With this
        .Id = 0
        .Qty = 1
        .IsNew = False
        .Code = "A000"
        .Name = "baseItemName Base desc"
        .Description = "Empty Desc"
        .Supplier = "Los Incas"
        .Cost = 9.99
        .WholeSalePrice = 999.99
        .SellingPrice = 9999.99
        .BoxQty = 24
        .BoxPrice = 99.99
    End With
End Sub
Public Sub Create(ByVal ICode As String, ByVal IName As String, ByVal IDescription As String, _
                    ByVal ISupplier As String, ByVal ICost As String, ByVal IWholeSalePrice As String, _
                    ByVal IPrice _
                    )
    With this
        .Code = ICode
        .Name = IName
        .Description = IDescription
        .Supplier = ISupplier
        .Cost = ICost
        .WholeSalePrice = IWholeSalePrice
        .SellingPrice = IPrice
    End With
    
End Sub



Public Property Get Qty() As Integer
    Qty = this.Qty
End Property

Public Property Let Qty(ByVal RHS As Integer)
    this.Qty = RHS
End Property

Public Property Get Id() As Integer
    Id = this.Id
End Property

Public Property Let Id(ByVal value As Integer)
    this.Id = value
End Property

Public Property Get IsNew() As Boolean
    IsNew = this.IsNew
End Property

Public Property Let IsNew(ByVal value As Boolean)
    this.IsNew = value
End Property

Public Property Get Code() As String
    Code = this.Code
End Property

Public Property Let Code(ByVal value As String)
    this.Code = value
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal value As String)
    this.Name = value
End Property

Public Property Get Description() As String
    Description = this.Description
End Property

Public Property Let Description(ByVal value As String)
    this.Description = value
End Property

Public Property Get Supplier() As String
    Supplier = this.Supplier
End Property

Public Property Let Supplier(ByVal value As String)
    this.Supplier = value
End Property

Public Property Get Cost() As Double
    Cost = this.Cost
End Property

Public Property Let Cost(ByVal value As Double)
    this.Cost = value
End Property

Public Property Get WholeSalePrice() As Double
    WholeSalePrice = this.WholeSalePrice
End Property

Public Property Let WholeSalePrice(ByVal value As Double)
    this.WholeSalePrice = value
End Property

Public Property Get SellingPrice() As Double
    SellingPrice = this.SellingPrice
End Property

Public Property Let SellingPrice(ByVal value As Double)
    this.SellingPrice = value
End Property


Public Property Get BoxQty() As Integer
    BoxQty = this.BoxQty
End Property

Public Property Let BoxQty(ByVal RHS As Integer)
    this.BoxQty = RHS
End Property

Public Property Get BoxPrice() As Double
    BoxPrice = this.BoxPrice
End Property

Public Property Let BoxPrice(ByVal RHS As Double)
    this.BoxPrice = RHS
End Property




