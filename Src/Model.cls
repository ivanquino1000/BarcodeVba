VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'KEYPROP REQUIREMENTS
'ITEM
'BoxPrice
'BoxQty
'Code
'Cost
'Description
'Name
'Qty
'SellingPrice
'Supplier
'WholeSalePrice
'COMPANY

'Address As String
'Name As String
'Ruc As String
'Email as string
'Phone as string
'facebook as string
'PRINTING PLACE
'Date


Public Enum Theme
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
    Theme           As Theme
    Coord           As Dictionary
    ResizeInst      As Dictionary    ' Direction,Spaces
    FallBack        As Dictionary
    Format          As collection

End Type

Private this        As TModel

Private Sub class_initialize()
    With this
        .Coord = New Dictionary
        .ResizeInst = New Dictionary
        .FallBack = New Dictionary
        .Format = New collection
    End With
    With .Coord
        .add "BoxPrice", array(1, 1)
        .add "BoxQty", array(1, 1)
        .add "Code", array(1, 1)
        .add "Cost", array(1, 1)
        .add "Description", array(1, 1)
        .add "Name", array(1, 1)
        .add "Qty", array(1, 1)
        .add "SellingPrice", array(1, 1)
        .add "Supplier", array(1, 1)
        .add "WholeSalePrice", array(1, 1)

        .add "Address", array(1, 1)
        .add "Name", array(1, 1)
        .add "Ruc", array(1, 1)
        .add "Email", array(1, 1)
        .add "Phone", array(1, 1)
        .add "facebook", array(1, 1)

        .add "Date", array(1, 1)
    End With
    'R:Right, L:Left, U:Up, D:Down
    With ResizeInst

        .add "BoxPrice", array("R", 1)
        .add "BoxQty", array("R", 1)
        .add "Code", array("R", 1)
        .add "Cost", array("R", 1)
        .add "Description", array("R", 1)
        .add "Name", array("R", 1)
        .add "Qty", array("R", 1)
        .add "SellingPrice", array("R", 1)
        .add "Supplier", array("R", 1)
        .add "WholeSalePrice", array("R", 1)

        .add "Address", array("R", 1)
        .add "Name", array("R", 1)
        .add "Ruc", array("R", 1)
        .add "Email", array("R", 1)
        .add "Phone", array("R", 1)
        .add "facebook", array("R", 1)

        .add "Date", array("R", 1)


    End With
    With FallBack
        .add "BoxPrice", "Default"
        .add "BoxQty", "Default"
        .add "Code", "Default"
        .add "Cost", "Default"
        .add "Description", "Default"
        .add "Name", "Default"
        .add "Qty", "Default"
        .add "SellingPrice", "Default"
        .add "Supplier", "Default"
        .add "WholeSalePrice", "Default"

        .add "Address", "Default"
        .add "Name", "Default"
        .add "Ruc", "Default"
        .add "Email", "Default"
        .add "Phone", "Default"
        .add "facebook", "Default"

        .add "Date", "Default"
    End With


    With Format
        .add "BoxPrice", CreateFormat(LongNumber, this.Theme)
        .add "BoxQty", CreateFormat(LongNumber, this.Theme)
        .add "Code", CreateFormat(LongNumber, this.Theme)
        .add "Cost", CreateFormat(LongNumber, this.Theme)
        .add "Description", CreateFormat(LongNumber, this.Theme)
        .add "Name", CreateFormat(LongNumber, this.Theme)
        .add "Qty", CreateFormat(LongNumber, this.Theme)
        .add "SellingPrice", CreateFormat(LongNumber, this.Theme)
        .add "Supplier", CreateFormat(LongNumber, this.Theme)
        .add "WholeSalePrice", CreateFormat(LongNumber, this.Theme)

        .add "Address", CreateFormat(LongNumber, this.Theme)
        .add "Name", CreateFormat(LongNumber, this.Theme)
        .add "Ruc", CreateFormat(LongNumber, this.Theme)
        .add "Email", CreateFormat(LongNumber, this.Theme)
        .add "Phone", CreateFormat(LongNumber, this.Theme)
        .add "facebook", CreateFormat(LongNumber, this.Theme)

        .add "Date", CreateFormat(LongNumber, this.Theme)
    End With
End Sub


Private Function CreateFormat(datatype As Integer, Theme As Theme) As FormatSettings
    Dim Format      As FormatSettings
    Set Format = New FormatSettings
    'Theme Changes
    With Format
        Select Case Theme
            Case light
                .BgColor = RGB(0, 0, 0)
                .FontColor = RGB(255, 255, 255)
            Case Else
                .BgColor = RGB(255, 255, 255)
                .FontColor = RGB(255, 255, 255)
        End Select
        'DataType Changes
        Select Case datatype

            Case LongNumber
                .FontName = "Arial"
                .FontSize = 12
                .FontBold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case ShortNumber
                .FontName = "Arial"
                .FontSize = 12
                .FontBold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case TitleText
                .FontName = "Arial"
                .FontSize = 12
                .FontBold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case NormalText
                .FontName = "Arial"
                .FontSize = 12
                .FontBold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case LabelText
                .FontName = "Arial"
                .FontSize = 12
                .FontBold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case Date
                .FontName = "Arial"
                .FontSize = 12
                .FontBold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
            Case Else
                .FontName = "Arial"
                .FontSize = 12
                .FontBold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shirnk = False
        End Select
    End With


    Set CreateFormat = Format
End Function

Public Property Get RowSize() As Integer
    RowSize = this.RowSize
End Property

Public Property Let RowSize(value As Integer)
    this.RowSize = value
End Property

Public Property Get ColumnSize() As Integer
    ColumnSize = this.ColumnSize
End Property

Public Property Let ColumnSize(value As Integer)
    this.ColumnSize = value
End Property

Public Property Get Theme() As Theme
    Theme = this.Theme
End Property

Public Property Let Theme(value As Theme)
    this.Theme = value
End Property

Public Property Get Coord() As Dictionary
    Set Coord = this.Coord
End Property

Public Property Get ResizeInst() As Dictionary
    Set ResizeInst = this.ResizeInst
End Property

Public Property Get FallBack() As Dictionary
    Set FallBack = this.FallBack
End Property

Public Property Get Format() As collection
    Set Format = this.Format
End Property



