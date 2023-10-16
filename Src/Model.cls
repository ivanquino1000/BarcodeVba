VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Themes
    light
    dark
End Enum

Public Enum datatype
    LongNumber
    ShortNumber
    TitleText
    NormalText
    LabelText
    Date
End Enum

Private Type TModel
    RowSize         As Integer
    ColumnSize      As Integer
    Theme           As Integer
    FallBack        As Dictionary
    GridMap         As Variant
    ProductCoord    As Dictionary
    ProductFormat   As Collection
    ProductResize   As Dictionary
    CompanyCoord    As Dictionary
    CompanyFormat   As Collection
    CompanyResize   As Dictionary
    LabelCoord      As Dictionary
    LabelValues     As Dictionary
    LabelResize     As Dictionary
    LabelFormat     As Collection
End Type

Private this        As TModel

Private Sub class_initialize()
    With this
        .RowSize = 5
        .ColumnSize = 5
        .Theme = 1
        Set .FallBack = New Dictionary
        Set .ProductCoord = New Dictionary
        Set .ProductFormat = New Collection
        Set .ProductResize = New Dictionary
        Set .CompanyCoord = New Dictionary
        Set .CompanyFormat = New Collection
        Set .CompanyResize = New Dictionary
        Set .LabelCoord = New Dictionary
        Set .LabelValues = New Dictionary
        Set .LabelResize = New Dictionary
        Set .LabelFormat = New Collection


        With .ProductCoord
            .Add "BoxPrice", Array(5, 1)
            .Add "BoxQty", Array(4, 1)
            .Add "Code", Array(1, 1)
            .Add "Cost", Array(2, 1)
            .Add "Description", Array(2, 2)
            .Add "Name", Array(1, 2)
            .Add "Qty", Array(5, 3)
            .Add "SellingPrice", Array(4, 3)
            .Add "Supplier", Array(4, 2)
            .Add "WholeSalePrice", Array(5, 2)
        End With

        With .ProductResize
            .Add "BoxPrice", Array("R", 0)
            .Add "BoxQty", Array("R", 0)
            .Add "Code", Array("R", 0)
            .Add "Cost", Array("R", 0)
            .Add "Description", Array("R", 0)
            .Add "Name", Array("R", 0)
            .Add "Qty", Array("R", 0)
            .Add "SellingPrice", Array("R", 0)
            .Add "Supplier", Array("R", 0)
            .Add "WholeSalePrice", Array("R", 0)


            .Add "Date", Array("R", 0)

        End With
        With .ProductFormat
            .Add CreateFormat(datatype.LongNumber, this.Theme), "BoxPrice"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "BoxQty"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "Code"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "Cost"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "Description"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "Name"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "Qty"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "SellingPrice"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "Supplier"
            .Add CreateFormat(datatype.LongNumber, this.Theme), "WholeSalePrice"


        End With





        With .CompanyCoord
            .Add "Address", Array("R", 1)
            .Add "Name", Array("R", 1)
            .Add "Ruc", Array("R", 1)
            .Add "Email", Array("R", 1)
            .Add "Phone", Array("R", 1)
            .Add "facebook", Array("R", 1)

        End With
        With .CompanyResize

            .Add "Address", Array("R", 1)
            .Add "Name", Array("R", 1)
            .Add "Ruc", Array("R", 1)
            .Add "Email", Array("R", 1)
            .Add "Phone", Array("R", 1)
            .Add "facebook", Array("R", 1)

        End With
        With .CompanyFormat

            .Add CreateFormat(1, this.Theme), "Address"
            .Add CreateFormat(1, this.Theme), "Name"
            .Add CreateFormat(1, this.Theme), "Ruc"
            .Add CreateFormat(1, this.Theme), "Email"
            .Add CreateFormat(1, this.Theme), "Phone"
            .Add CreateFormat(1, this.Theme), "facebook"
        End With
'
'        With .LabelCoord
'            .Add "Address", Array("R", 1)
'            .Add "CompanyName", Array("R", 1)
'            .Add "Ruc", Array("R", 1)
'            .Add "Email", Array("R", 1)
'            .Add "Phone", Array("R", 1)
'            .Add "facebook", Array("R", 1)
'
'        End With
'        With .LabelResize
'            .Add "Address", Array("R", 1)
'            .Add "CompanyName", Array("R", 1)
'            .Add "Ruc", Array("R", 1)
'            .Add "Email", Array("R", 1)
'            .Add "Phone", Array("R", 1)
'            .Add "facebook", Array("R", 1)
'
'        End With
'        With .LabelFormat
'            .Add "Address", Array("R", 1)
'            .Add "CompanyName", Array("R", 1)
'            .Add "Ruc", Array("R", 1)
'            .Add "Email", Array("R", 1)
'            .Add "Phone", Array("R", 1)
'            .Add "facebook", Array("R", 1)
'
'        End With
'        With .LabelValues
'            .Add "Address", Array("R", 1)
'            .Add "CompanyName", Array("R", 1)
'            .Add "Ruc", Array("R", 1)
'            .Add "Email", Array("R", 1)
'            .Add "Phone", Array("R", 1)
'            .Add "facebook", Array("R", 1)
'
'        End With
        With FallBack
            .Add "BoxPrice", "Name"
            .Add "BoxQty", "Name"
            .Add "Code", "Name"
            .Add "Cost", "Name"
            .Add "Description", "Name"
            .Add "Name", "Name"
            .Add "Qty", "Name"
            .Add "SellingPrice", "Name"
            .Add "Supplier", "Name"
            .Add "WholeSalePrice", "Name"

        End With

    End With
    this.GridMap = CreateGrid

End Sub


Private Function CreateFormat(ByVal datatype As datatype, ByVal Theme As Integer) As Object
    Dim format      As FormatSettings
    Set format = New FormatSettings
    'Theme Changes
    With format
        Select Case Theme
            Case 1
                .BgColor = RGB(0, 0, 0)
                .FontColor = RGB(255, 255, 255)
            Case Else
                .BgColor = RGB(255, 255, 255)
                .FontColor = RGB(255, 255, 255)
        End Select
        'DataType Changes
        Select Case datatype

            Case 1
                .FontName = "Arial"
                .FontSize = 12
                .Bold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case 2
                .FontName = "Arial"
                .FontSize = 12
                .Bold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case 3
                .FontName = "Arial"
                .FontSize = 12
                .Bold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case 4
                .FontName = "Arial"
                .FontSize = 12
                .Bold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case 5
                .FontName = "Arial"
                .FontSize = 12
                .Bold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case 6
                .FontName = "Arial"
                .FontSize = 12
                .Bold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case Else
                .FontName = "Arial"
                .FontSize = 12
                .Bold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
        End Select
    End With


    Set CreateFormat = format
End Function

Public Function CreateGrid() As Variant
    Dim Grid        As Variant
    Dim key As Variant
    With this
        ReDim Grid(1 To .RowSize, 1 To .ColumnSize)
        With .ProductCoord
            For Each key In .Keys
                Debug.Print key
                Grid(.Item(key)(0), .Item(key)(1)) = key
            Next key
        End With

    End With
    CreateGrid = Grid
End Function


Public Property Get RowSize() As Integer
    RowSize = this.RowSize
End Property

Public Property Let RowSize(ByVal RHS As Integer)
    this.RowSize = RHS
End Property

Public Property Get ColumnSize() As Integer
    ColumnSize = this.ColumnSize
End Property

Public Property Let ColumnSize(ByVal RHS As Integer)
    this.ColumnSize = RHS
End Property

Public Property Get Theme() As Integer
    Theme = this.Theme
End Property

Public Property Let Theme(ByVal RHS As Integer)
    this.Theme = RHS
End Property

Public Property Get ResizeInst() As Dictionary
    Set ResizeInst = this.ResizeInst
End Property

Public Property Set ResizeInst(ByVal RHS As Dictionary)
    Set this.ResizeInst = RHS
End Property

Public Property Get FallBack() As Dictionary
    Set FallBack = this.FallBack
End Property

Public Property Set FallBack(ByVal RHS As Dictionary)
    Set this.FallBack = RHS
End Property

Public Property Get format() As Collection
    Set format = this.format
End Property

Public Property Set format(ByVal RHS As Collection)
    Set this.format = RHS
End Property

Public Property Get GridMap() As Variant
    If IsObject(this.GridMap) Then
        Set GridMap = this.GridMap
    Else
        GridMap = this.GridMap
    End If
End Property

Public Property Let GridMap(ByVal RHS As Variant)
    this.GridMap = RHS
End Property

Public Property Set GridMap(ByVal RHS As Variant)
    Set this.GridMap = RHS
End Property

Public Property Get ProductCoord() As Dictionary
    Set ProductCoord = this.ProductCoord
End Property

Public Property Set ProductCoord(ByVal RHS As Dictionary)
    Set this.ProductCoord = RHS
End Property

Public Property Get CompanyCoord() As Dictionary
    Set CompanyCoord = this.CompanyCoord
End Property

Public Property Set CompanyCoord(ByVal RHS As Dictionary)
    Set this.CompanyCoord = RHS
End Property

Public Property Get LabelCoord() As Dictionary
    Set LabelCoord = this.LabelCoord
End Property

Public Property Set LabelCoord(ByVal RHS As Dictionary)
    Set this.LabelCoord = RHS
End Property

Public Property Get LabelValues() As Dictionary
    Set LabelValues = this.LabelValues
End Property

Public Property Set LabelValues(ByVal RHS As Dictionary)
    Set this.LabelValues = RHS
End Property

Public Property Get ProductFormat() As Collection
    Set ProductFormat = this.ProductFormat
End Property

Public Property Set ProductFormat(ByVal RHS As Collection)
    Set this.ProductFormat = RHS
End Property

Public Property Get CompanyFormat() As Collection
    Set CompanyFormat = this.CompanyFormat
End Property

Public Property Set CompanyFormat(ByVal RHS As Collection)
    Set this.CompanyFormat = RHS
End Property

Public Property Get ProductResize() As Dictionary
    Set ProductResize = this.ProductResize
End Property

Public Property Set ProductResize(ByVal RHS As Dictionary)
    Set this.ProductResize = RHS
End Property

Public Property Get CompanyResize() As Dictionary
    Set CompanyResize = this.CompanyResize
End Property

Public Property Set CompanyResize(ByVal RHS As Dictionary)
    Set this.CompanyResize = RHS
End Property

Public Property Get LabelResize() As Dictionary
    Set LabelResize = this.LabelResize
End Property

Public Property Set LabelResize(ByVal RHS As Dictionary)
    Set this.LabelResize = RHS
End Property

Public Property Get LabelFormat() As Collection
    Set LabelFormat = this.LabelFormat
End Property

Public Property Set LabelFormat(ByVal RHS As Collection)
    Set this.LabelFormat = RHS
End Property

