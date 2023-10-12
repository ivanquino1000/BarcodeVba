VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrintingForm 
   Caption         =   "Configuracion Inicial"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "PrintingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrintingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OkButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    Dim Bussiness_1 As New Company
    Dim Bussiness_2 As New Company
    Dim Bussiness_3 As New Company
    Dim Bussiness_4 As New Company
    Dim Bussiness_5 As New Company

    Bussiness_1.init "SORAZA", "Octavio Munoz Najar 202", 20608875426#
    Bussiness_2.init "G&R", "San Juan de Dios 601A", 1075699791#
    Bussiness_3.init "LARICO", "Octavio Munoz Najar 248", 20609183650#
    Bussiness_4.init "BENKI", "Teniente Ferre ", 10402036899#
    Bussiness_5.init "ALMACEN", "Mercedez Benz Mz F Lt 7", 20608875426#

    With CompanyBox

        .AddItem Bussiness_1.Name
        .AddItem Bussiness_2.Name
        .AddItem Bussiness_3.Name
        .AddItem Bussiness_4.Name
        .AddItem Bussiness_5.Name
    End With

'    Dim printerName As Variant
'    For Each printerName In Application.Printers
'        ComboBox1.AddItem printerName.DeviceName
'    Next printerName
    PrintingForm.Show
End Sub
