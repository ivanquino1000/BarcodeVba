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
    BgColor As Long
    FontSize As Long
    FontColor As String
    FontName As String
End Type

Private this As TFormatSettings

Private Sub Class_Initialize()
'Headers Default Format
    With this
        .BgColor = RGB(112, 88, 124)
        .FontSize = 16
        .FontColor = vbWhite
        .FontName = "Courier New"
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

