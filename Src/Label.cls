VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Enum Models
    Default
    Reduced
    Full
End Enum

Private Type TLabel
    Row             As Integer
    Column          As Integer
    Model           As Model
    Instances       As Integer
    Ws              As Worksheet
    Product         As Item
    Bussiness As Company
End Type
Private this        As TLabel

Public Sub class_initialize()
    With this
        .Row = 1
        .Column = 1
        Set .Model = New Model
        .Instances = 0
        Set .Ws = ThisWorkbook.Sheets("LabelSheet")
        Set .Product = New Item
        Set .Bussiness = New Company
    End With
End Sub

Public Sub build(ByVal Item As Item, ByVal Bussiness As Company)
    With this
        .Row = 1
        .Column = 1
        Set .Model = Models.Default
        .Instances = 0
        .Ws = thiswowkbook.Sheets("LabelSheet")
        Set .Product = Item
        Set .Bussiness = Bussiness
    End With
End Sub

Public Sub ToRange()
    Dim LabelRange As range
    Dim cell As range
    With this
        Set LabelRange = .Ws.Cells(.Row, .Column).Resize(.Model.RowSize, .Model.ColumnSize)
        LabelRange = .Model.GridMap
        For Each cell In LabelRange
            Dim key As String
            Dim KeyVal As Variant
            Dim Direction As String
            Dim places As Integer
            Dim format As FormatSettings
            'Check For Empty Prop - Call Fallback
            key = cell.value
            If key = "" Then
                GoTo EmptySpace
            End If
            On Error Resume Next
            KeyVal = CallByName(.Product, key, VbGet)
            
            On Error GoTo 0
            
            Direction = .Model.ProductResize(key)(0)
            places = .Model.ProductResize(key)(1)
            Set format = .Model.ProductFormat.Item(key)
            If IsEmpty(KeyVal) Then
                key = this.Model.FallBack(key)
                KeyVal = CallByName(.Bussiness, key, VbGet)
                Direction = .Model.CompanyResize(key)(0)
                places = .Model.CompanyResize(key)(1)
                Set format = .Model.ProductFormat.Item(key)
            End If
            'ReDimention Cell Selection to Merge
            Select Case Direction     'R,L,U,D

                Case "R"
                    Set cell = cell.Resize(1, 1 + places)
                Case "L"
                    Set cell = cell.offset(0, -places).Resize(1, places + 1)
                Case "U"
                    Set cell = cell.offset(-places, 0).Resize(places + 1, 1)
                Case "D"
                    Set cell = cell.Resize(1 + places, 1)
            End Select

            cell.Merge
            
            'Format Property Cell
            ApplyFormat cell, format
            
            'Fill Item Values
            cell.Value2 = KeyVal
EmptySpace:
        Next cell
            
    End With
End Sub


Public Property Get Row() As Integer
    Row = this.Row
End Property

Public Property Let Row(ByVal RHS As Integer)
    this.Row = RHS
End Property

Public Property Get Column() As Integer
    Column = this.Column
End Property

Public Property Let Column(ByVal RHS As Integer)
    this.Column = RHS
End Property

Public Property Get Model() As Model
    Set Model = this.Model
End Property

Public Property Set Model(ByVal RHS As Model)
    Set this.Model = RHS
End Property

Public Property Get Instances() As Integer
    Instances = this.Instances
End Property

Public Property Let Instances(ByVal RHS As Integer)
    this.Instances = RHS
End Property

Public Property Get Ws() As Worksheet
    Set Ws = this.Ws
End Property

Public Property Set Ws(ByVal RHS As Worksheet)
    Set this.Ws = RHS
End Property

Public Property Get Product() As Item
    Set Product = this.Product
End Property

Public Property Set Product(ByVal RHS As Item)
    Set this.Product = RHS
End Property


Public Property Get Bussiness() As Company
    Set Bussiness = this.Bussiness
End Property

Public Property Set Bussiness(ByVal RHS As Company)
    Set this.Bussiness = RHS
End Property

