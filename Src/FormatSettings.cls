VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
VERSION 1.0 CLASS
Option Explicit

Private Type TFormatSettings
    BgColor As Long
    Bold As Boolean
    BorderStyle As XlLineStyle
    BorderWeight As XlBorderWeight
    FontColor As Long
    FontName As String
    FontSize As Long
    HAlign As XlHAlign
    Shirnk As Boolean
    VAlign As XlVAlign
End Type

Private this As TFormatSettings

Public Event DataInitialized(ByVal BgColor As Long, ByVal FontColor As Long, ByVal FontSize As Long, ByVal FontName As String)

Private Sub Class_Initialize()
    With this
        .BgColor = RGB(112, 88, 124)
        .Bold = False
        .BorderStyle = xlDashDot
        .BorderWeight = xlMedium
        .FontColor = vbWhite
        .FontName = "Courier New"
        .FontSize = 16
        .HAlign = xlCenter
        .Shirnk = False
        .VAlign = xlCenter
    End With
End Sub

Public Property Get BgColor() As Long
    BgColor = this.BgColor
End Property

Public Property Let BgColor(ByVal value As Long)
    this.BgColor = value
End Property

Public Property Get Bold() As Boolean
    Bold = this.Bold
End Property

Public Property Let Bold(ByVal value As Boolean)
    this.Bold = value
End Property

Public Property Get BorderStyle() As XlLineStyle
    BorderStyle = this.BorderStyle
End Property

Public Property Let BorderStyle(ByVal value As XlLineStyle)
    this.BorderStyle = value
End Property

Public Property Get BorderWeight() As XlBorderWeight
    BorderWeight = this.BorderWeight
End Property

Public Property Let BorderWeight(ByVal value As XlBorderWeight)
    this.BorderWeight = value
End Property

Public Property Get FontColor() As Long
    FontColor = this.FontColor
End Property

Public Property Let FontColor(ByVal value As Long)
    this.FontColor = value
End Property

Public Property Get FontName() As String
    FontName = this.FontName
End Property

Public Property Let FontName(ByVal value As String)
    this.FontName = value
End Property

Public Property Get FontSize() As Long
    FontSize = this.FontSize
End Property

Public Property Let FontSize(ByVal value As Long)
    this.FontSize = value
End Property

Public Property Get HAlign() As XlHAlign
    HAlign = this.HAlign
End Property

Public Property Let HAlign(ByVal value As XlHAlign)
    this.HAlign = value
End Property

Public Property Get Shirnk() As Boolean
    Shirnk = this.Shirnk
End Property

Public Property Let Shirnk(ByVal value As Boolean)
    this.Shirnk = value
End Property

Public Property Get VAlign() As XlVAlign
    VAlign = this.VAlign
End Property

Public Property Let VAlign(ByVal value As XlVAlign)
    this.VAlign = value
End Property

Private Sub Build(ByVal BgColor As Long, ByVal FontSize As Long, ByVal FontColor As Long, _
                ByVal FontName As String, ByVal Bold As Boolean, ByVal Shirnk As Boolean, _
                ByVal VAlign As XlVAlign, ByVal HAlign As XlHAlign, ByVal BorderStyle As XlLineStyle, _
                ByVal BorderWeight As XlBorderWeight)
    With this
        .BgColor = BgColor
        .Bold = Bold
        .BorderStyle = BorderStyle
        .BorderWeight = BorderWeight
        .FontColor = FontColor
        .FontName = FontName
        .FontSize = FontSize
        .HAlign = HAlign
        .Shirnk = Shirnk
        .VAlign = VAlign
    End With
End Sub