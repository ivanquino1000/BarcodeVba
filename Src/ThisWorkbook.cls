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

'ON MAIN SHEET REMAP OVERRIDE DEFAULT KEYS AND BUTTONS
Const IntroductoryMessage As String = "Welcome to Codebar Printer You will first need to know " & _
        vbNewLine & "       PASOS A SEGUIR ANTES DE INICIAR:        " & _
        vbNewLine & "- Descargar de internet la base de datos" & _
        vbNewLine & "- Teclas Especiales " & vbNewLine & _
        vbNewLine & "      @ENTER NUMERICO: borra barra de busqueda" & _
        vbNewLine & "      @TEXTO *?*: Datos Opcionales de Relleno" & vbNewLine & _
        vbNewLine & "- Al finalizar enviar y subir los productos creados en esta sesion" & _
        vbNewLine & "-IvanQuino1000"
Dim InitTime, EndTime As Double
Public PrinterName  As String

Private WithEvents listTable As DynamicTable
Attribute listTable.VB_VarHelpID = -1
Private WithEvents searchTable As StaticTable
Attribute searchTable.VB_VarHelpID = -1
Private WithEvents resultsTable As StaticTable
Attribute resultsTable.VB_VarHelpID = -1
Private WithEvents databaseTable As StaticTable
Attribute databaseTable.VB_VarHelpID = -1

'TODO : IMPLEMENT USER INPUT MESSAGES - PRINTER NAME -COMPANY
'CUSTOM KEY TO CHANGE SUPPLIER
'TODO: On items  !exist documents  => Move From APPResources
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
    With ThisWorkbook.Worksheets
        .Select
        ActiveWindow.Zoom = 60
    End With

    'Set & Create Tables
    InitTime = Timer

    Set listTable = New DynamicTable
    Set databaseTable = New StaticTable
    Set searchTable = New StaticTable
    Set resultsTable = New StaticTable

    listTable.init 5, 2, ThisWorkbook.Sheets("MainSheet"), _
            Array( _
            "Cant", "Codigo", "Nombre", _
            "Adicional", "Proveedor", "Cant*Caja", "Precio*Caja", _
            "Costo?", "xMayor?", "Precio")
    databaseTable.init 1, 1, ThisWorkbook.Sheets("DataSheet"), _
            Array( _
            "Codigo", "Nombre", "Precio" _
            )
    searchTable.init 2, 13, ThisWorkbook.Sheets("MainSheet"), _
            Array( _
            "Codigo", "Nombre", "Precio" _
            )
    resultsTable.init 5, 13, ThisWorkbook.Sheets("MainSheet"), _
            Array( _
            "Codigo", "Nombre", "Precio" _
            )
    EndTime = Timer
    Debug.Print "UITableBuild- Performance:", EndTime - InitTime & " sec"

    databaseTable.GetData

    ThisWorkbook.Sheets("MainSheet").Select

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub



'TODO: On SearchBar - Create New Item - Send Random Key
'TODO: @ListTable - Avoid Double Click Action on first row if empty

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As range, Cancel As Boolean)

    If Sh.Name = "MainSheet" Then
        '@AddElement / ResultsTable

        If Not Intersect(Target, resultsTable.ElementsRange) Is Nothing Then

            Cancel = True

            Application.EnableEvents = False
            Dim rowId As Integer

            Dim CodeRange As range
            Dim NameRange As range
            Dim PriceRange As range

            Dim Product As New item

            rowId = Target.row - resultsTable.row

            With resultsTable.ElementsRange
                Set CodeRange = .Columns(1)
                Set NameRange = .Columns(2)
                Set PriceRange = .Columns(3)
            End With

            'Fill Product Data
            With Product
                .Code = CodeRange.Cells(rowId).Value2
                .Name = NameRange.Cells(rowId).Value2
                .SellingPrice = PriceRange.Cells(rowId).Value2
            End With

            If Not Intersect(Target, CodeRange) Is Nothing Then
                Product.Code = "AC0001"
            End If

            listTable.ElementsComp.values.Add Product
            listTable.Update

            Application.EnableEvents = True
        End If

        '@ListTable / Delete
        If Not Intersect(Target, listTable.ElementsComp.range) Is Nothing Then
            With Application
                .EnableEvents = False
                .ScreenUpdating = False
                If listTable.ElementsComp.values.Count < 1 Then
                    GoTo EmptyItems
                End If
                Dim ItemId As Integer
                With listTable
                    ItemId = Target.row - .row
                    With .ElementsComp
                        .values.Remove ItemId

                    End With
                    .Update
                End With

                .EnableEvents = True
                .ScreenUpdating = True
            End With
EmptyItems:
        End If

    End If

End Sub

'TODO: ExactMatch - AutoInsertKey - Highlight Option
'TODO: On Elemets Range Changed - Update Element Collection
'TODO: On Elemets RANGE COSTO - WHOLE - BOX => UNITARY In Table
'Labels as placeholders

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As range)


    '@UI - MainSheet
    If Sh.Name = "MainSheet" Then

        '@ListTable - Update Values
        If Not Intersect(Target, listTable.ElementsComp.range) Is Nothing Then
            With Application
                .EnableEvents = False
                .ScreenUpdating = False

                With listTable
                    '@No Elemnts To Update
                    If listTable.ElementsComp.values.Count < 1 Then
                        GoTo EmptyElements
                    End If

                    Dim item As range
                    For Each item In .ElementsComp.range.Rows

                        If IsEmpty(item) Then

                            GoTo NextItem
                        End If

                        Dim cell As range
                        For Each cell In item
'                            Dim prop As String
'                            prop = .ElementsComp.HeaderMap(.HeadersComp.range(, Target.column - .column + 1).value)
'
'                            Dim ElementId As Long: ElementId = Target.row - .row
'                            Dim Element As Object
'                            Set Element = .ElementsComp.values.item(ElementId)
'
'                            Select Case prop
'                                Case "Qty", "BoxPrice", "BoxQty", "Cost", "SellingPrice", "WholeSalePrice"
'                                    If Not IsNumeric(prop) Then
'                                        'Target.Value2 = ""
'                                        'Exit Sub
'                                    End If
'                                    CallByName Element, prop, VbLet, Target.Value2
'                                    .ElementsComp.Update Sh, .HeadersComp.values
'                                Case Else
'                                    If Not IsNumeric(prop) Then
'                                        'Target.Value2 = ""
'                                        'Exit Sub
'                                    End If
'                                    CallByName Element, prop, VbLet, StrConv(Target.Value2, vbProperCase)
'                                    .ElementsComp.Update Sh, .HeadersComp.values
'                            End Select
                        Next cell
NextItem:
                    Next item
                End With
                '########
                '                With listTable
                '
                '                    Dim prop As String
                '                    prop = .ElementsComp.HeaderMap(.HeadersComp.range(, Target.column - .column + 1).value)
                '                    If Not .ElementsComp.HeaderMap.Exists(prop) Then
                '                        'Target.Value2 = ""
                '                        'Debug.Print "missMe"
                '                        'GoTo InvalidValues
                '                    End If
                '
                '                    Dim ElementId As Long: ElementId = Target.row - .row
                '                    Dim Element As Object
                '                    Set Element = .ElementsComp.values.item(ElementId)
                '
                '                    Select Case prop
                '                        Case "Qty", "BoxPrice", "BoxQty", "Cost", "SellingPrice", "WholeSalePrice"
                '                            If Not IsNumeric(prop) Then
                '                                'Target.Value2 = ""
                '                                'Exit Sub
                '                            End If
                '                            CallByName Element, prop, VbLet, Target.Value2
                '                            .ElementsComp.Update Sh, .HeadersComp.values
                '                        Case Else
                '                            If Not IsNumeric(prop) Then
                '                                'Target.Value2 = ""
                '                                'Exit Sub
                '                            End If
                '                            CallByName Element, prop, VbLet, StrConv(Target.Value2, vbProperCase)
                '                            .ElementsComp.Update Sh, .HeadersComp.values
                '                    End Select
                '                End With
                .EnableEvents = True
                .ScreenUpdating = True
            End With
InvalidValues:
        End If

        '@UI [SearchBar]
        If Not Intersect(Target, searchTable.ElementsRange) Is Nothing Then
            With Application
                .EnableEvents = False
                .ScreenUpdating = False
                Dim item As Object

                If Target.Rows.Count > 1 Then
                    For Each item In Target.Rows
                        GoTo UpdateProperties
                    Next item
                End If

                Set item = Target
UpdateProperties:
                Dim cell As range

                For Each cell In item
                    If IsEmpty(item) Then
                        GoTo FillResults
                    End If

                    Dim Header, UserInput As String
                    Dim HeaderIndex As Integer
                    With searchTable
                        HeaderIndex = cell.column - .column
                        Header = listTable.ElementsComp.HeaderMap( _
                                .HeadersComp.values(HeaderIndex) _
                                )
                        UserInput = cell.value
                    End With

                    If UserInput = "" Then
                        GoTo FillResults
                    End If

                    Select Case Header
                        Case "Code":
                            cell.value = UserInput
                        Case "Name":
                            cell.value = "*" & UserInput
                        Case "SellingPrice":
                            If IsNumeric(UserInput) & UserInput > 30 Then
                                cell.value = "<=" & UserInput
                            Else
                                cell.value = ">=" & UserInput
                            End If
                        Case Else
                            Debug.Print "User Input: Header Out of Index"
                    End Select
                Next cell

FillResults:

                resultsTable.ClearPrevRange
                Dim Elements As Integer
                databaseTable.range.AdvancedFilter Action:=xlFilterCopy, _
                        criteriarange:=searchTable.range, CopyToRange:=resultsTable.HeadersComp.range, _
                        Unique:=False

                With resultsTable
                    Elements = .ws.Cells(.ws.Rows.Count, .column).End(xlUp).row - .row
                    .ElementsCounter = Elements
                    .ElementsRange.Sort key1:=.ElementsRange.Columns(UBound(.HeadersComp.values) + 1), _
                            order1:=xlDescending, Header:=xlNo
                End With


                .EnableEvents = True
                .ScreenUpdating = True
            End With
        End If

        '@Headers - Recovery
        If (Not Intersect(Target, resultsTable.HeadersComp.range) Is Nothing) Or (Not Intersect(Target, searchTable.HeadersComp.range) Is Nothing) Or (Not Intersect(Target, listTable.HeadersComp.range) Is Nothing) Then
            With Application
                .EnableEvents = False
                resultsTable.HeadersComp.Update Sh
                listTable.HeadersComp.Update Sh
                searchTable.HeadersComp.Update Sh
                .EnableEvents = True
            End With
        End If

    End If

    '#Data Sheet
    If Sh.Name = "DataSheet" Then

        '@Headers - Recovery
        If Not Intersect(Target, databaseTable.HeadersComp.range) Is Nothing Then
            With Application
                .EnableEvents = False
                databaseTable.HeadersComp.Update Sh
                .EnableEvents = True
            End With
        End If

    End If
End Sub



