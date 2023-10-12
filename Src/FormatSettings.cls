VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event DataInitialized(ByVal BgColor As String, ByVal textColor As String, ByVal textSize As Long, ByVal font As String)





Private Type TFormatSettings
    BgColor         As Long
    FontSize        As Long
    FontColor       As String
    FontName        As String
    Bold            As Boolean
    VAlign          As Boolean
    Shirnk          As Integer
    HAlign          As Integer
    BorderStyle     As Integer
    BorderWeight    As Integer
End Type
Private this        As TFormatSettings


Private Sub Class_Initialize()
    'Headers Default Format
    With this
        .BgColor = RGB(112, 88, 124)
        .FontSize = 16
        .FontColor = vbWhite
        .FontName = "Courier New"
        .Bold = False
        .Shirnk = False
        .VAlign = xlCenter
        .HAlign = xlCenter
        .BorderStyle = xlDashDot
        .BorderWeight = xmedium
    End With
End Sub

Private Sub Build(ByVal FBgColor As Long, ByVal FFontSize As Long, ByVal FFontColor As String, _
                ByVal FFontName As String, ByVal FBold As Boolean, ByVal FShirnk As Boolean, _
                ByVal FVAlign As Integer, ByVal FHAlign As Integer, ByVal FBorderStyle As Integer, _
                ByVal FBorderWeight As Integer)
    With this
        .BgColor = FBgColor
        .FontSize = FFontSize
        .FontColor = FFontColor
        .FontName = FFontName
        .Bold = FBold
        .Shirnk = FShrink
        .VAlign = FVAlign
        .HAlign = FHAlign
        .BorderStyle = FBorderStyle
        .BorderWeight = FBorderWeight
    End With
End Sub

'@Properties
Public Property Get BgColor() As String
    BgColor = this.BgColor
End Property

Public Property Let BgColor(ByVal value As String)
    this.BgColor = value
End Property

Public Property Get FontSize() As Long
    FontSize = this.FontSize
End Property

Public Property Let FontSize(ByVal value As Long)
    this.FontSize = value
End Property

Public Property Get FontColor() As String
    FontColor = this.FontColor
End Property

Public Property Let FontColor(ByVal value As String)
    this.FontColor = value
End Property

Public Property Get FontName() As String
    FontName = this.FontName
End Property

Public Property Let FontName(ByVal value As String)
    this.FontName = value
End Property


Public Property Get Bold() As Boolean
    Bold = this.Bold
End Property

Public Property Let Bold(ByVal RHS As Boolean)
    this.Bold = RHS
End Property

Public Property Get VAlign() As Boolean
    VAlign = this.VAlign
End Property

Public Property Let VAlign(ByVal RHS As Boolean)
    this.VAlign = RHS
End Property

Public Property Get Shirnk() As Boolean
    Shirnk = this.Shirnk
End Property

Public Property Let Shirnk(ByVal RHS As Boolean)
    this.Shirnk = RHS
End Property

Public Property Get HAlign() As Boolean
    HAlign = this.HAlign
End Property

Public Property Let HAlign(ByVal RHS As Boolean)
    this.HAlign = RHS
End Property


Public Property Get BorderStyle() As String
    BorderStyle = this.BorderStyle
End Property

Public Property Let BorderStyle(ByVal RHS As String)
    this.BorderStyle = RHS
End Property

Public Property Get BorderWeight() As String
    BorderWeight = this.BorderWeight
End Property

Public Property Let BorderWeight(ByVal RHS As String)
    this.BorderWeight = RHS
End Property
