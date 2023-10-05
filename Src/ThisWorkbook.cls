VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const IntroductoryMessage As String = "Welcome to Codebar Printer You will first need to know " & _
        vbNewLine & "       PASOS A SEGUIR ANTES DE INICIAR:        " & _
        vbNewLine & "- Descargar de internet la base de datos" & _
        vbNewLine & "- Teclas Especiales " & vbNewLine & _
        vbNewLine & "      @ENTER NUMERICO: borra barra de busqueda" & _
        vbNewLine & "      @TEXTO *?*: Datos Opcionales de Relleno" & vbNewLine & _
        vbNewLine & "- Al finalizar enviar y subir los productos creados en esta sesion" & _
        vbNewLine & "-IvanQuino1000"

Private WithEvents listTable As DynamicTable
Attribute listTable.VB_VarHelpID = -1
Private WithEvents searchTable As StaticTable
Attribute searchTable.VB_VarHelpID = -1
Private WithEvents resultsTable As StaticTable
Attribute resultsTable.VB_VarHelpID = -1
Private WithEvents databaseTable As DynamicTable
Attribute databaseTable.VB_VarHelpID = -1


Private Sub Workbook_Open()
    Application.EnableAutoComplete = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim sheet       As Worksheet
    For Each sheet In Me.Sheets
        With sheet.Cells
            .ClearContents
            .ClearFormats
            .Interior.Color = RGB(17, 21, 28)
        End With
    Next sheet
    'Set & Create Tables

    Set listTable = New DynamicTable
    Set databaseTable = New DynamicTable
    Set searchTable = New StaticTable
    Set resultsTable = New StaticTable

    listTable.Init 5, 5, ThisWorkbook.Sheets("MainSheet"), _
            Array( _
            "Cant", "Codigo", "Nombre", _
            "Adicional", "Proveedor", "Costo?", _
            "xMayor?", "Precio")
    databaseTable.Init 1, 1, ThisWorkbook.Sheets("DataSheet"), _
            Array( _
            "Codigo", "Nombre", "Precio" _
            )
    searchTable.Init 2, 15, ThisWorkbook.Sheets("MainSheet"), _
            Array( _
            "Codigo", "Nombre", "Precio" _
            )
    resultsTable.Init 5, 15, ThisWorkbook.Sheets("MainSheet"), _
            Array( _
            "Codigo", "Nombre", "Precio" _
            )


    databaseTable.GetData

    ThisWorkbook.Sheets("MainSheet").Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    If Not Intersect(Target, searchTable.ElementsRange) Is Nothing Then
        resultsTable.ClearPrevRange
        Dim Elements As Integer
        databaseTable.Range.AdvancedFilter Action:=xlFilterCopy, _
                criteriarange:=searchTable.Range, CopyToRange:=resultsTable.HeadersComp.Range, _
                Unique:=False
             With resultsTable
            Elements = .ws.Cells(.ws.Rows.Count, .column).End(xlUp).row - .row
        End With
        resultsTable.ElementsCounter = Elements
        
    End If
End Sub

