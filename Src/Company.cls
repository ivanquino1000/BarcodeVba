VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private Type TCompany
    Address As String
    Name As String
    Ruc As String
    Email As String
    Phone As String
    facebook As String

End Type

Private this As TCompany


Private Sub class_initialize()
    With this
        .Address = "Octavio Mu�oz Najar 202"
        .Name = "SORAZA"
        .Ruc = "20608875426"
        .Email = "sorazaimport@gmail.com"
        .Phone = "958190331"
        .facebook = "tinyurl.com/soraza"
    End With
End Sub

Public Sub init(ByVal IName As String, ByVal IAddress As String, ByVal IRuc As String)
    With this
        .Address = IAddress
        .Name = IName
        .Ruc = IRuc
    End With
End Sub



Public Property Get Address() As String
    Address = this.Address
End Property

Public Property Let Address(ByVal RHS As String)
    this.Address = RHS
End Property

Public Property Get Ruc() As Long
    Ruc = this.Ruc
End Property

Public Property Let Ruc(ByVal RHS As Long)
    this.Ruc = RHS
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal RHS As String)
    this.Name = RHS
End Property
