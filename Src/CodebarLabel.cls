VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodebarLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Type TCodebarLabel
    ItemObj         As Item
    CompanyObj      As Company
    PropConfig      As collection
    PropFormat      As collection
    Theme           As Integer
    ws              As Worksheet
    PrintedDate     As Date
End Type

Private this        As TCodebarLabel

Private Sub class_initialize()
    With this
        .ws = thiswowkbook.sheets("LabelSheet")
        .Theme = 0
        .PropConfig = New collection
        .PropFormat = New collection
        .PrintedDate = Now

        'Item Pos Properties Configuration
        AddPropConfig "Qty", 1, 1, 0, "Default"
        AddPropConfig "Code"
        AddPropConfig "Name", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropConfig "Description", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropConfig "Supplier", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropConfig "Cost", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlRight, False, xlCenter
        AddPropConfig "WholeSalePrice", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlRight, False, xlCenter
        AddPropConfig "SellingPrice", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlRight, False, xlCenter
        AddPropConfig "BoxQty", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial Narrow", 14, xlCenter, False, xlCenter
        AddPropConfig "BoxPrice", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlRight, False, xlCenter
        AddPropConfig "IsNew", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter    'Company Properties Configuration
        'Company Pos Properties Configuration

        AddPropConfig "Address", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropConfig "Name", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropConfig "Ruc", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropConfig "Phone", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial Narrow", 14, xlCenter, False, xlCenter
        AddPropConfig "Email", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial Narrow", 14, xlCenter, False, xlCenter
        AddPropConfig "facebook", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Webdings", 14, xlCenter, False, xlCenter

        AddPropConfig "Default", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Webdings", 14, xlCenter, False, xlCenter

        'Item Properties Configuration
        AddPropFormat "Qty", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropFormat "Code", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropFormat "Name", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropFormat "Description", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropFormat "Supplier", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropFormat "Cost", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlRight, False, xlCenter
        AddPropFormat "WholeSalePrice", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlRight, False, xlCenter
        AddPropFormat "SellingPrice", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlRight, False, xlCenter
        AddPropFormat "BoxQty", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial Narrow", 14, xlCenter, False, xlCenter
        AddPropFormat "BoxPrice", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlRight, False, xlCenter
        AddPropFormat "IsNew", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter    'Company Properties Configuration
        'Company Properties Configuration

        AddPropFormat "Address", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropFormat "Name", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropFormat "Ruc", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial", 14, xlCenter, False, xlCenter
        AddPropFormat "Phone", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial Narrow", 14, xlCenter, False, xlCenter
        AddPropFormat "Email", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Arial Narrow", 14, xlCenter, False, xlCenter
        AddPropFormat "facebook", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Webdings", 14, xlCenter, False, xlCenter

        AddPropFormat "Default", RGB(1, 1, 1), False, xlDashDot, xlMedium, RGB(1, 1, 1), "Webdings", 14, xlCenter, False, xlCenter


    End With
End Sub

Private Sub AddPropConfig( _
        ByVal PropName As String, _
        ByVal x As Double, _
        ByVal y As Double, _
        ByVal offset As Double, _
        ByVal FallBack As String)

    Dim Config      As New PosSettings
    With Config
        .x = x
        .y = y
        .offset = offset
        .FallBack = FallBack
    End With
    this.PropConfig.add PropName, Config
End Sub

Private Sub AddPropFormat( _
        ByVal PropName As String, _
        ByVal BgColor As Long, _
        ByVal Bold As Boolean, _
        ByVal BorderStyle As XlLineStyle, _
        ByVal BorderWeight As XlBorderWeight, _
        ByVal FontColor As Long, _
        ByVal FontName As String, _
        ByVal FontSize As Long, _
        ByVal HAlign As XlHAlign, _
        ByVal Shirnk As Boolean, _
        ByVal VAlign As XlVAlign)

    Dim Format      As New FormatSettings
    With Format
        .BgColor = BgColor
        .Bold = Bold
        .BorderStyle = BorderStyle
        .BorderWeight = BorderWeight
        .FontColor = FontColor
        .FontName = FontName
        .FontSize = FontSize
        .HAlign = HAlign
        .Shrink = Shirnk
        .VAlign = VAlign
    End With
    this.PropSettings.add PropName, Format
End Sub

Private Sub Build(ByVal CItem As Item, ByVal CTheme As Integer)
    With this
        .ItemObj = CItem
        .Theme = CTheme
    End With
End Sub

Public Function CreateRange() As range
    Dim CbRange     As range
    With this
        For Each prop In .ItemObj
            Dim key As String
            key = prop.Name
            If .PropConfig.Exists(prop) Then
                If IsEmpty(CallByName(.ItemObj, prop, VbGet)) Then
                    key = .PropConfig.Item(prop).FallBack
                End If
                CbRange(.PosSettings(key).x, .PosSettings(key).y).value = CallByName(.ItemObj, prop.Name, VbGet)
                ApplyFormat CbRange(.PosSettings(key).x, .PosSettings(key).y), .PropSettings.Item(prop.Name)
            End If

        Next prop
    End With
End Function

Private Sub ToUI()

End Sub

Public Property Get ItemObj() As Item
    Set ItemObj = this.ItemObj
End Property

Public Property Set ItemObj(ByVal RHS As Item)
    Set this.ItemObj = RHS
End Property

Public Property Get Theme() As Item
    Set Theme = this.Theme
End Property

Public Property Set Theme(ByVal RHS As Item)
    Set this.Theme = RHS
End Property

Public Property Get PropSettings() As Dictionary
    Set PropSettings = this.PropSettings
End Property

Public Property Set PropSettings(ByVal RHS As Dictionary)
    Set this.PropSettings = RHS
End Property

Public Property Get DefaultProps() As Item
    Set DefaultProps = this.DefaultProps
End Property

Public Property Set DefaultProps(ByVal RHS As Item)
    Set this.DefaultProps = RHS
End Property

