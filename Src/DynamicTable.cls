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
    row             As Integer
    column          As Integer
    Range           As Range
    SrcBookPath     As String
    NewItemsBookPath As String
    TestBookPath    As String
    ws              As Worksheet
    TempBookPath    As String
    
End Type


Private this        As TDynamicTable

Private Sub Class_Initialize()
    With this
        .row = 1
        .column = 1

        Set .ws = ThisWorkbook.ActiveSheet
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

Public Sub Init(ByVal row As Integer, _
        ByVal column As Integer, _
        ByVal ws As Worksheet, _
        ByVal Headers As Variant _
        )
    With this
        .row = row
        .column = column
        Set .ws = ws
        .SrcBookPath = FindLatestXLSXFile(Environ("USERPROFILE") & "\Downloads\")
        .NewItemsBookPath = Environ("USERPROFILE") & "\Documents\IMPORT" & "Items.xlsx"
        .TempBookPath = ThisWorkbook.Path & "\Config" & "\TempData.xmls\"
        With HeadersComp
            .values = Headers
        End With
        Update
    End With


End Sub

Public Sub GetData()

    On Error GoTo getDataError

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
        SrcElemntsCounter = SourceSheet.Cells(SourceSheet.Rows.Count, .CodeCol).End(xlUp).row
        Dim CodesRng, SrcTabRng, cell As Range
        Set CodesRng = SourceSheet.Range( _
                .CodeCol & .FirstRow & ":" & _
                .CodeCol & SrcElemntsCounter)
        Set SrcTabRng = SourceSheet.Range( _
                .CodeCol & .FirstRow & ":" & _
                .PriceCol & SrcElemntsCounter)
        For Each cell In CodesRng
            If Not IsEmpty(cell.value) Then
                If Not regex.Test(cell.value) Then
                    cell.value = ""
                End If
            End If

        Next cell
        'SourceBook Formating: Sorting + elemCounter Update
        SrcTabRng.Sort key1:=SrcTabRng.Columns(1), order1:=xlAscending, Header:=xlYes
        SrcElemntsCounter = SourceSheet.Cells(SourceSheet.Rows.Count, .CodeCol).End(xlUp).row
        'Save SourceBook Elements to ElementsComp Collection
        Dim items   As New Collection

        Dim i       As Integer
        Dim itemObj As Item
        For i = .FirstRow To SrcElemntsCounter
            Set itemObj = New Item
            With itemObj
                .Code = SourceSheet.Cells(i, SourceTableRef.CodeCol).value
                .Name = SourceSheet.Cells(i, SourceTableRef.NameCol).value
                .SellingPrice = SourceSheet.Cells(i, SourceTableRef.PriceCol).value
            End With
            items.Add itemObj
        Next i
        Set ElementsComp.values = items

    End With
    SourceBook.Close True
    'Update DynamicTable
    Update
    Exit Sub
getDataError:
    Debug.Print "getDataError:  "; err.Description

End Sub

Public Sub Update()
    UpdateRanges
    With this
        .ws.Select
        HeadersComp.Update .ws
        ElementsComp.Update .ws, HeadersComp.values
        .ws.Columns.AutoFit
    End With
End Sub

'Dependencies: Row,Column, _
 Headers.values,Elements.Values

Private Sub UpdateRanges()

    With HeadersComp
        Set .Range = this.ws.Range( _
                this.ws.Cells(this.row, this.column), _
                this.ws.Cells(this.row, this.column + UBound(.values)) _
                )
    End With

    With ElementsComp
        Dim EmptyElements As Integer
        If .values.Count = 0 Then
            EmptyElements = 1
        End If
        Set .Range = this.ws.Range( _
                this.ws.Cells(this.row + 1, this.column), _
                this.ws.Cells(this.row + EmptyElements + .values.Count, this.column + UBound(HeadersComp.values)) _
                )
    End With
    With this
        Set .Range = Union( _
                HeadersComp.Range, _
                ElementsComp.Range _
                )
    End With
End Sub
Public Sub SearchMissingItem(ByVal CodeLetter As String, ByVal LastCodeId As Integer)
    If LastCodeId = 0 Then
    Else
    End If
    
End Sub


Private Sub ElementsComp_RangeUpdated()
    UpdateRanges
End Sub

Private Sub HeadersComp_HeadersUpdated(ByVal Headers As Variant)
    UpdateRanges
End Sub

Public Property Get row() As Integer
    row = this.row
End Property

Public Property Let row(ByVal value As Integer)
    this.row = value
    UpdateRanges
End Property

Public Property Get column() As Integer
    column = this.column
End Property

Public Property Let column(ByVal value As Integer)
    this.column = value
    UpdateRanges
End Property

Public Property Get Range() As Range
    Set Range = this.Range
End Property

Public Property Set Range(ByVal value As Range)
    Set this.Range = value
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

Public Property Get ws() As Worksheet
    Set ws = this.ws
End Property

Public Property Set ws(ByVal value As Worksheet)
    Set this.ws = value
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


