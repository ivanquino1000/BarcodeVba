VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IDataTable

Public Event HeadersPlaced(ByVal Headers As Variant)
Public Event elementAdded(ByVal element As Item)
Public Event TableUpdated()
Public Event ElementDeleted(ByVal element As Item)
Public Event DataGotten(ByVal elem_num As Long)

Public WithEvents HeadersComp As Headers
Attribute HeadersComp.VB_VarHelpID = -1
Public WithEvents ElementsComp As Elements
Attribute ElementsComp.VB_VarHelpID = -1


Private Type SourceTab
    FirstRow        As Long
    CodeCol         As String
    NameCol         As String
    PriceCol        As String
End Type

Private SourceTableRef As SourceTab

Private Type TDynamicTable
    Row             As Integer
    Column          As Integer
    range           As range
    SrcBookPath     As String
    NewItemsBookPath As String
    TestBookPath    As String
    Ws              As Worksheet
    TempBookPath    As String

End Type


Private this        As TDynamicTable

Private Sub class_initialize()
    With this
        .Row = 1
        .Column = 1

        Set .Ws = ThisWorkbook.ActiveSheet
    End With
    Set HeadersComp = New Headers
    Set ElementsComp = New Elements
    With SourceTableRef
        .CodeCol = "B"
        .FirstRow = 4
        .NameCol = "E"
        .PriceCol = "K"
    End With
End Sub

Public Sub init(ByVal Row As Integer, _
            ByVal Column As Integer, _
            ByVal Ws As Worksheet, _
            ByVal Headers As Variant _
            )
    With this
        .Row = Row
        .Column = Column
        Set .Ws = Ws
        .SrcBookPath = FindLatestXLSXFile(Environ("USERPROFILE") & "\Downloads\")
        .NewItemsBookPath = Environ("USERPROFILE") & "\Documents\IMPORT" & "Items.xlsx"
        .TempBookPath = ThisWorkbook.Path & "\Config" & "\TempData.xmls\"
        With HeadersComp
            .values = Headers
        End With
        Update
    End With


End Sub

' Executed once per cycle - Improve efficiency

Public Sub GetData()

    On Error GoTo getDataError

    Dim InitTime, EndTime As Double
    InitTime = Timer

    Dim SourceBook  As Workbook: Set SourceBook = Workbooks.Open(this.SrcBookPath)
    Dim SourceSheet As Worksheet: Set SourceSheet = SourceBook.Worksheets(1)

    Dim ElemArr     As Variant
    Dim SrcElemntsCounter As Long

    'regex Code Settings
    Dim regex       As Object: Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .Pattern = "^[A-Z,a-z]{1}\d{3}\b"
    End With
    'Clear Not Matching Codes

    With SourceTableRef
        SrcElemntsCounter = SourceSheet.Cells(SourceSheet.Rows.Count, .CodeCol).End(xlUp).Row

        Dim CodesRng, SrcTabRng, cell As range

        Set CodesRng = SourceSheet.range( _
                    .CodeCol & .FirstRow & ":" & _
                    .CodeCol & SrcElemntsCounter)
        Set SrcTabRng = SourceSheet.range( _
                    .CodeCol & .FirstRow & ":" & _
                    .PriceCol & SrcElemntsCounter)
        For Each cell In CodesRng
            If Not IsEmpty(cell.value) Then
                If Not regex.test(cell.value) Then
                    cell.value = ""
                End If
            End If

        Next cell
        'SourceBook Formating: Sorting + elemCounter Update
        SrcTabRng.Sort key1:=SrcTabRng.Columns(1), order1:=xlAscending, Header:=xlYes
        SrcElemntsCounter = SourceSheet.Cells(SourceSheet.Rows.Count, .CodeCol).End(xlUp).Row

        'Save SourceBook Elements to ElementsComp Collection - Dictionary
        Dim items   As New Collection
        Dim IdValuesMap As New Dictionary
        Dim i       As Long
        Dim ItemObj As Item
        Dim Code, Name, Price As String
        Dim IdValues As Variant: IdValues = Array("-1")
        Dim CurrentKey, PrevKey As String
        Dim CodeId, TempIdValCount As Long
        PrevKey = ExtractLetter(SourceSheet.range(.CodeCol & .FirstRow & ":" & .CodeCol & .FirstRow).value)

        For i = .FirstRow To SrcElemntsCounter

            Code = SourceSheet.Cells(i, .CodeCol).value
            Name = SourceSheet.Cells(i, .NameCol).value
            Price = SourceSheet.Cells(i, .PriceCol).value
            CurrentKey = ExtractLetter(Code)
            CodeId = ExtractNumber(Code)
            Set ItemObj = New Item

            With ItemObj
                .Code = Code
                .Name = Name
                .SellingPrice = Price
            End With
            'Match CodeKey to previous - Push to IdValues[]
            If CurrentKey = PrevKey Then
                'Push To IdValues
                TempIdValCount = TempIdValCount + 1
                ReDim Preserve IdValues(0 To TempIdValCount)
                IdValues(TempIdValCount - 1) = CodeId
                PrevKey = CurrentKey
            Else 'Unmatching CodeKeys - Store PrevKey,IdValues[] - Reset IdValues - Init CodeId
                'ElementsComp.CodeIds.Add PrevKey, IdValues
                IdValuesMap.Add PrevKey, IdValues
                ReDim IdValues(0 To 0) 'Erase IdValues
                IdValues(0) = CodeId
                PrevKey = CurrentKey
                TempIdValCount = 0
            End If

            items.Add ItemObj

        Next i
        Set ElementsComp.values = items
        Set ElementsComp.CodeIds = IdValuesMap
    End With
    SourceBook.Close True
    EndTime = Timer
    Debug.Print "GetData - Performance:", EndTime - InitTime & " sec"
    'Update ElementsIds
    Update
    Exit Sub
getDataError:
    Debug.Print "getDataError:  "; err.Description
    If Not SourceBook Is Nothing And SourceSheet.Visible Then
        SourceBook.Close False
    End If

End Sub

Public Sub Update()
    UpdateRanges
    With this
        .Ws.Select
        HeadersComp.Update .Ws
        ElementsComp.Update .Ws, HeadersComp.values
        .Ws.Columns.AutoFit
    End With
End Sub

'Dependencies: Row,Column, _
 Headers.values,Elements.Values

Private Sub UpdateRanges()

    With HeadersComp
        Set .range = this.Ws.range( _
                    this.Ws.Cells(this.Row, this.Column), _
                    this.Ws.Cells(this.Row, this.Column + UBound(.values)) _
                    )
    End With

    With ElementsComp
        Dim EmptyElements As Integer
        If .values.Count = 0 Then
            EmptyElements = 1
        End If
        Set .range = this.Ws.range( _
                    this.Ws.Cells(this.Row + 1, this.Column), _
                    this.Ws.Cells(this.Row + EmptyElements + .values.Count, this.Column + UBound(HeadersComp.values)) _
                    )
    End With
    With this
        Set .range = Union( _
                    HeadersComp.range, _
                    ElementsComp.range _
                    )
    End With
End Sub

Private Sub PrintTest()
    Dim printingRange As range
    Dim Label As New CodebarLabel
    Dim Row, Column As Integer
    'Dim LabelRange As Range
    Label.Build()
    Set printingRange = Label.range




    printingRange.PrintOut Copies:=1, Preview:=True, ActivePrinter:=PrinterName
End Sub
Private Sub PrintList()

End Sub


Private Sub ElementsComp_RangeUpdated()
    UpdateRanges
End Sub

Private Sub HeadersComp_HeadersUpdated(ByVal Headers As Variant)
    UpdateRanges
End Sub

Public Property Get Row() As Integer
    Row = this.Row
End Property

Public Property Let Row(ByVal value As Integer)
    this.Row = value
    UpdateRanges
End Property

Public Property Get Column() As Integer
    Column = this.Column
End Property

Public Property Let Column(ByVal value As Integer)
    this.Column = value
    UpdateRanges
End Property

Public Property Get range() As range
    Set range = this.range
End Property

Public Property Set range(ByVal value As range)
    Set this.range = value
End Property

Public Property Get SrcBookPath() As String
    SrcBookPath = this.SrcBookPath
End Property

Public Property Let SrcBookPath(ByVal value As String)
    this.SrcBookPath = value
End Property

Public Property Get NewItemsBookPath() As String
    NewItemsBookPath = this.NewItemsBookPath
End Property

Public Property Let NewItemsBookPath(ByVal value As String)
    this.NewItemsBookPath = value
End Property

Public Property Get TestBookPath() As String
    TestBookPath = this.TestBookPath
End Property

Public Property Let TestBookPath(ByVal value As String)
    this.TestBookPath = value
End Property

Public Property Get Ws() As Worksheet
    Set Ws = this.Ws
End Property

Public Property Set Ws(ByVal value As Worksheet)
    Set this.Ws = value
End Property

Public Property Get TempBookPath() As String
    TempBookPath = this.TempBookPath
End Property

Public Property Let TempBookPath(ByVal RHS As String)
    this.TempBookPath = RHS
End Property

Private Sub IDataTable_Update()
    err.Raise 5                                  'TODO implement interface member
End Sub


