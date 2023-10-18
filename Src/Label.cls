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
    Product         As item
    Bussiness       As Company
End Type
Private this        As TLabel

Public Sub class_initialize()
    With this
        .Row = 2
        .Column = 2
        Set .Model = New Model
        .Instances = 0
        Set .Ws = ThisWorkbook.Sheets("LabelSheet")
        Set .Product = New item
        Set .Bussiness = New Company
    End With
End Sub

Public Sub build(ByVal item As item, ByVal Bussiness As Company)
    With this
        .Row = 2
        .Column = 2
        Set .Model = Models.Default
        .Instances = 0
        .Ws = thiswowkbook.Sheets("LabelSheet")
        Set .Product = item
        Set .Bussiness = Bussiness
    End With
End Sub

Public Sub ToRange()
    Dim LabelRange  As range
    Dim cell        As range

    With this
        'Define Label Range in Starting Position and Model Dimensions
        'Copy ModelGridMap to LabelRange
        Set LabelRange = .Ws.Cells(.Row, .Column).Resize(.Model.RowSize, .Model.ColumnSize)
        LabelRange = .Model.GridMap
        For Each cell In LabelRange
            Dim key As String
            Dim keyVal As Variant
            Dim origin As String
KeyChecker:
            key = cell.Value2
            'Check CellVal and ItemObj Associated
            If key = "" Then
                GoTo ToNextCell
            End If
            If .Model.KeyInModel(.Product, .Bussiness, key, keyVal) Then
                If keyVal = "" Then
                    If .Model.FallBack.Exists(key) Then
                        key = .Model.FallBack(key)
                        GoTo KeyChecker
                    End If
                    GoTo ToNextCell
                End If
            End If

            GoTo ToResizeCell
ToResizeCell:

            Dim Direction As String
            Dim Places As Integer
            With .Model
                If Not .SetResizeInst(key, Direction, Places) Then
                    GoTo ToFormatCell
                End If

            End With


            Redimension cell, Direction, Places
ToFormatCell:
            Dim format As FormatSettings

            With .Model
                If Not .SetFormat(key, format) Then
                    GoTo ToNextCell
                End If

            End With

            ApplyFormat cell, format
            cell.value = keyVal

ToNextCell:
        Next cell
    End With

    LabelRange.Columns.AutoFit
    'LabelRange.Rows.AutoFit
End Sub

Private Function Redimension(ByRef cell As range, ByVal Direction As String, ByVal Places As Integer)
    Select Case Direction
        Case "R"
            Set cell = cell.Resize(1, 1 + Places)
        Case "L"
            Set cell = cell.offset(0, -Places).Resize(1, Places + 1)
        Case "U"
            Set cell = cell.offset(-Places, 0).Resize(Places + 1, 1)
        Case "D"
            Set cell = cell.Resize(1 + Places, 1)
        Case Else
    End Select
    cell.Merge
End Function


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

Public Property Get Product() As item
    Set Product = this.Product
End Property

Public Property Set Product(ByVal RHS As item)
    Set this.Product = RHS
End Property


Public Property Get Bussiness() As Company
    Set Bussiness = this.Bussiness
End Property

Public Property Set Bussiness(ByVal RHS As Company)
    Set this.Bussiness = RHS
End Property

