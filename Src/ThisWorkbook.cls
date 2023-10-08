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

    listTable.Init 5, 2, ThisWorkbook.Sheets("MainSheet"), _
            Array( _
            "Cant", "Codigo", "Nombre", _
            "Adicional", "Proveedor", "Cant*Caja", "Precio*Caja", _
            "Costo?", "xMayor?", "Precio")
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

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

'
'    '
'    'Results Elemnts Double Click
'    '
'
'    If Not Intersect(Target, resultsTable.ElementsRange) Is Nothing Then
'        'Required Ranges
'        Dim ElemRange, cell As Range
'        Dim ElemProp As String
'        Dim ElemPropId As Integer
'        Dim itemObj As New Item
'
'        With resultsTable
'
'            ElemPropId = Target.row - .column
'            ElemProp = listTable.ElementsComp.HeaderMap( _
'                    .HeadersComp.values(ElemPropId) _
'                    )
'            Set ElemRange = .ws.Range( _
'                    Cells(Target.row, .column), _
'                    Cells(Target.row, .column + UBound(.HeadersComp.values)) _
'                    )
'        End With
'
'        With resultsTable
'            'Item Builder: elemRange Values
'            Dim i   As Integer
'            For Each cell In ElemRange
'                CallByName itemObj, _
'                        listTable.ElementsComp.HeaderMap( _
'                        .HeadersComp.values(i)), VbLet, cell.value
'                i = i + 1
'            Next cell
'        End With
'
'        'If ElemProp  = Code - Update Code Id on Previous Saved Ids
'        With listTable
'            If ElemProp = "Code" Then
'                Dim CodeLet As String
'                Dim LetReg As New RegExp
'                Dim NewCode As String
'                With regex
'                    .Global = True
'                    .IgnoreCase = True
'                    .Pattern = "\^[a-z,A-Z]+"
'                    Set CodeLet = .Execute(inputString)
'                End With
'                ' Find all matches in the input string
'                If Not .ElementsComp.LastNewCodeId.Exists(CodeLet) Then
'
'                Else
'
'                End If
'                NewCode = CodeLet & _
'                        .SearchMissingItem(CodeLet, .ElementsComp.LastNewCodeId(CodeLet))
'
'            End If
'            .ElementsComp.values.Add itemObj
'            .Update
'
'        End With
'
'        Cancel = True

    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    ' SearchField-Change
    If Not Intersect(Target, searchTable.ElementsRange) Is Nothing Then

        Dim Header, UserInput As String
        Dim HeaderIndex As Integer
        With searchTable
            HeaderIndex = Target.column - .column
            Header = listTable.ElementsComp.HeaderMap( _
                    .HeadersComp.values(HeaderIndex) _
                    )
            UserInput = Target.value
        End With

        
        If UserInput = "" Then
        GoTo FillResults
        End If
        Application.EnableEvents = False
        Select Case Header
            Case "Code":
                Target.value = UserInput
            Case "Name":
                Target.value = "*" & UserInput
            Case "SellingPrice":
                                If IsNumeric(UserInput) & UserInput > 30 Then
                                    Target.value = "<=" & UserInput
                                Else
                                    Target.value = ">=" & UserInput
                                    End If
            Case Else
                Debug.Print "User Input: Header Out of Index"
        End Select
        Application.EnableEvents = True

        'Fill & Format Results Table
FillResults:
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

