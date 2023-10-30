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
'POO
Public Enum datatype
    longnumber
    ShortNumber
    TitleText
    NormalText
    LabelText
    Date
    Barcode
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
        .RowSize = 6
        .ColumnSize = 3
        .Theme = light
        
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
            .Add "BoxPrice", Array(5, 1)    'Unit
            '.Add "BoxQty", Array(4, 1)
            .Add "Code", Array(1, 1)
            .Add "Cost", Array(2, 1)
            .Add "Description", Array(2, 2)
            .Add "Name", Array(1, 2)
            '.Add "Qty", Array(4, 2)
            .Add "SellingPrice", Array(4, 3)
            .Add "Supplier", Array(6, 3)
            .Add "WholeSalePrice", Array(5, 2)
        End With

        With .ProductResize
            .Add "BoxPrice", Array("R", 0)
            .Add "BoxQty", Array("R", 0)
            .Add "Code", Array("R", 0)
            .Add "Cost", Array("R", 0)
            .Add "Description", Array("R", 1)
            .Add "Name", Array("R", 1)
            .Add "Qty", Array("R", 0)
            .Add "SellingPrice", Array("D", 1)
            .Add "Supplier", Array("R", 0)
            .Add "WholeSalePrice", Array("R", 0)


            '.Add "Date", Array("R", 0)

        End With
        With .ProductFormat
            .Add CreateFormat(datatype.ShortNumber, Themes.light), "BoxPrice"
            .Add CreateFormat(datatype.ShortNumber, Themes.dark), "BoxQty"
            .Add CreateFormat(datatype.TitleText, Themes.dark), "Code"
            .Add CreateFormat(datatype.TitleText, Themes.dark), "Cost"
            .Add CreateFormat(datatype.NormalText, Themes.light), "Description"
            .Add CreateFormat(datatype.NormalText, Themes.light), "Name"
            .Add CreateFormat(datatype.ShortNumber, Themes.dark), "Qty"
            .Add CreateFormat(datatype.longnumber, Themes.dark), "SellingPrice"
            .Add CreateFormat(datatype.LabelText, Themes.dark), "Supplier"
            .Add CreateFormat(datatype.ShortNumber, Themes.light), "WholeSalePrice"
        End With

        With .CompanyCoord
            .Add "Address", Array(2, 2)
            .Add "Name", Array(2, 1)
            '.Add "Ruc", Array(1, 1)
            .Add "Email", Array(4, 1)
            .Add "Phone", Array(5, 1)
            .Add "facebook", Array(6, 1)

        End With
        With .CompanyResize

            .Add "Address", Array("R", 1)
            .Add "Name", Array("R", 0)
            .Add "Ruc", Array("R", 0)
            .Add "Email", Array("R", 0)
            .Add "Phone", Array("R", 0)
            .Add "facebook", Array("R", 0)

        End With
        With .CompanyFormat
            .Add CreateFormat(datatype.NormalText, Themes.light), "Address"
            .Add CreateFormat(datatype.TitleText, Themes.dark), "Name"
            .Add CreateFormat(datatype.NormalText, Themes.light), "Ruc"
            .Add CreateFormat(datatype.NormalText, Themes.light), "Email"
            .Add CreateFormat(datatype.NormalText, Themes.light), "Phone"
            .Add CreateFormat(datatype.NormalText, Themes.light), "facebook"
        End With

        With .LabelCoord
            .Add "BoxLabel", Array(4, 1)
            .Add "BoxTotalPrice", Array(6, 1)
            .Add "DozLabel", Array(4, 2)
            .Add "DozTotalPrice", Array(6, 2)
            .Add "Date", Array(6, 3)
            .Add "Barcode", Array(3, 1)

        End With
        With .LabelResize
            .Add "BoxLabel", Array("R", 0)
            .Add "BoxTotalPrice", Array("R", 0)
            .Add "DozLabel", Array("R", 0)
            .Add "DozTotalPrice", Array("R", 0)
            .Add "Date", Array("R", 0)
            .Add "Barcode", Array("R", 2)
        End With
        With .LabelFormat

            .Add CreateFormat(datatype.LabelText, Themes.dark), "BoxLabel"
            .Add CreateFormat(datatype.ShortNumber, Themes.light), "BoxTotalPrice"
            .Add CreateFormat(datatype.LabelText, Themes.dark), "DozLabel"
            .Add CreateFormat(datatype.ShortNumber, Themes.light), "DozTotalPrice"
            .Add CreateFormat(datatype.Barcode, Themes.dark), "Barcode"

            .Add CreateFormat(datatype.Date, Themes.dark), "Date"
        End With
        With .LabelValues
            .Add "BoxLabel", "Box()"
            .Add "BoxTotalPrice", 88
            .Add "DozLabel", "Docena()"
            .Add "DozTotalPrice", 99
            .Add "Date", Now
            .Add "Barcode", "EmptyBC"
        End With

        With FallBack
            .Add "BoxPrice", "Phone"
            '.Add "BoxQty", "Name"
            .Add "Code", ""
            .Add "Cost", "Name"
            .Add "Description", "Address"
            .Add "Name", ""
            '.Add "Qty", "Name"
            .Add "SellingPrice", ""
            .Add "Supplier", "Date"
            '.Add "WholeSalePrice", "Name"
            .Add "BoxLabel", "Email"
            .Add "BoxTotalPrice", "facebook"
            .Add "DozLabel", ""
            .Add "DozTotalPrice", ""
            '.Add "Date", Array("R", 1)
            '.Add "Barcode", Array("R", 1)
        End With

    End With
    this.GridMap = CreateGrid

End Sub


Private Function CreateFormat(ByVal datatype As datatype, ByVal Theme As Themes) As Object
    Dim format      As FormatSettings
    Set format = New FormatSettings
    If (this.Theme <> light) Then
        Theme = 1 - Theme
    End If
    'Theme Changes
    With format

        Select Case Theme
            Case light
                .BgColor = RGB(255, 255, 255)
                .FontColor = RGB(0, 0, 0)
            Case Else
                .BgColor = RGB(0, 0, 0)
                .FontColor = RGB(255, 255, 255)
        End Select
        'DataType Changes
        Select Case datatype
                'Code-39-Logitogo - 36

            Case longnumber
                .FontName = "Chinese Rocks Rg"
                .FontSize = 24
                .Bold = True
                .HAlign = xlRight
                .VAlign = xlCenter
                .Shrink = True
                .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Case ShortNumber
                .FontName = "Chinese Rocks Rg"
                .FontSize = 16
                .Bold = True
                .HAlign = xlRight
                .VAlign = xlCenter
                .Shrink = True
                .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Case TitleText
                .FontName = "BABAPRO FONT"
                .FontSize = 28
                .Bold = False
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shrink = True
            Case NormalText
                .FontName = "Bahnschrift Light SemiCondensed"
                .FontSize = 12
                .Bold = True
                .HAlign = xlLeft
                .VAlign = xlCenter
                .Shrink = True
            Case LabelText
                .FontName = "Japanese Robot"
                .FontSize = 12
                .Bold = True
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shrink = True
            Case Date
                .FontName = "Bahnschrift SemiLight"
                .FontSize = 12
                .Bold = False
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shrink = True
            Case Else    'Barcode
                .FontName = "Code 39-hoch-Logitogo"
                .FontSize = 60
                .Bold = False
                .HAlign = xlCenter
                .VAlign = xlCenter
                .Shrink = True
        End Select
    End With


    Set CreateFormat = format
End Function

Public Function CreateGrid() As Variant
    Dim Grid        As Variant
    Dim key         As Variant
    With this
        ReDim Grid(1 To .RowSize, 1 To .ColumnSize)
        With .ProductCoord
            For Each key In .Keys
                'Debug.Print key
                Grid(.item(key)(0), .item(key)(1)) = key
            Next key
        End With
        With .LabelCoord
            For Each key In .Keys
                Grid(.item(key)(0), .item(key)(1)) = key
            Next key
        End With
    End With
    CreateGrid = Grid
End Function

Public Function KeyInModel(ByVal item As Object, ByVal Company As Object, key As String, ByRef value As Variant) As Boolean
    On Error Resume Next
    value = CallByName(item, key, VbGet)
    If err.Number = 0 Then
        KeyInModel = True
        Exit Function
    End If

    On Error Resume Next
    value = CallByName(Company, key, VbGet)
    If err.Number = 0 Then
        KeyInModel = True
        Exit Function
    End If

    On Error Resume Next
    If this.LabelValues.Exists(key) Then
        value = this.LabelValues(key)
        KeyInModel = True
        Exit Function
    End If

    KeyInModel = False
End Function

Public Function SetResizeInst(ByVal key As String, ByRef dir As String, ByRef moves As Integer) As Boolean
    If this.ProductResize.Exists(key) Then
        dir = this.ProductResize(key)(0)
        moves = this.ProductResize(key)(1)
        SetResizeInst = True
        Exit Function
    End If
    If this.CompanyResize.Exists(key) Then
        dir = this.CompanyResize(key)(0)
        moves = this.CompanyResize(key)(1)
        SetResizeInst = True
        Exit Function
    End If
    If this.LabelResize.Exists(key) Then
        dir = this.LabelResize(key)(0)
        moves = this.LabelResize(key)(1)
        SetResizeInst = True
        Exit Function
    End If
    SetResizeInst = False
End Function

Public Function SetFormat(ByVal key As String, ByRef format As Object) As Boolean
    On Error Resume Next
    Set format = this.ProductFormat(key)
    If err.Number = 0 Then
        SetFormat = True
        Exit Function
    End If
    
    On Error Resume Next
    Set format = this.CompanyFormat(key)
    If err.Number = 0 Then
        SetFormat = True
        Exit Function
    End If
    
    On Error Resume Next
    Set format = this.LabelFormat(key)
    If err.Number = 0 Then
        SetFormat = True
        Exit Function
    End If
    SetFormat = False
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


