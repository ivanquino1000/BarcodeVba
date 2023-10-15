VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PosSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TPosition
    x As Double
    y As Double
    offset As Double
    FallBack As String
End Type

Private this As TPosition

Private Sub class_initialize()
    this.x = 1
    this.y = 1
    this.offset = 0
    this.FallBack = "Default"
End Sub

Public Property Get x() As Double
    x = this.x
End Property

Public Property Let x(ByVal value As Double)
    this.x = value
End Property

Public Property Get y() As Double
    y = this.y
End Property

Public Property Let y(ByVal value As Double)
    this.y = value
End Property

Public Property Get offset() As Double
    offset = this.offset
End Property

Public Property Let offset(ByVal value As Double)
    this.offset = value
End Property

Public Property Get FallBack() As String
    FallBack = this.FallBack
End Property

Public Property Let FallBack(ByVal value As String)
    this.FallBack = value
End Property



