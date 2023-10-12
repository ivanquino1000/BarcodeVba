VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodebarLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Enum Theme
    Black
    White
End Enum

Public Enum PropSett
    row
    column
    BgColor
    FontSize
    FontColor
    FontName
    Bold
    Shrink
    VAlign
    HAlign
    BorderStyle
    BorderWeight
    RightOffset
End Enum


Private Type TCodebarLabel
    ItemObj         As Item
    Theme           As Integer
    PropSettings    As Dictionary
    RequiredProps   As Variant
End Type

Private this        As TCodebarLabel

Private Sub Class_Initialize()
    With this

        .Theme = 0
        .PropSettings = New Dictionary
        With .PropSettings
            'r,c,Font,

            .Add "Cant", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "Codigo", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "Nombre", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "Adicional", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "Proveedor", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "Costo?", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "xMayor?", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "Precio", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "Cant*Caja", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
            .Add "Precio*Caja", Array(1, 1, RGB(1, 1, 1), 14, RGB(1, 1, 1), _
                    "Arial", False, False, x1center, x1center, _
                    xlDashDot, xmedium, 0)
        End With


        .DefaultProps = Array()

    End With
End Sub

Private Sub Build(ByVal CItem As Item, ByVal CTheme As Integer)
    With this
        .ItemObj = CItem
        .Theme = CTheme
    End With
End Sub

Public Function CreateRange() As range
    Dim cbRange As range
    With this
    For Each prop In .ItemObj
        If .PropSettings.Exists(prop) Then
            If IsEmpty(CallByName(.ItemObj, prop, VbGet)) Then
            
            End If
        End If
        
    Next prop
    End With
End Function


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

