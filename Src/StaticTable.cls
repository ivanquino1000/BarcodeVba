VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaticTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDataTable
'* TODO Created: 04-10-2023 13:51 Author: VBATools
'* Implement Update Procedure _
 * Set StaticElemnts Logic - static element Countes - prevElemRange - Elem Range _
 * Set UpdateLogic - ws.Select - HeadersUpdate - Clear prev elem range + format elem range

Public Event HeadersPlaced(ByVal Headers As Variant)
Public Event elementAdded(ByVal element As Item)
Public Event TableUpdated()
Public Event ElementDeleted(ByVal element As Item)
Public Event DataGotten(ByVal elem_num As Long)

Public WithEvents HeadersComp As Headers
Attribute HeadersComp.VB_VarHelpID = -1

Private Type TStaticTable
    row             As Integer
    column          As Integer
    range           As range
    ws              As Worksheet
    Headers         As Headers
    ElementsRange   As range
    PreviousElemRange As range
    ElementsCounter As Integer
    ElementsFormat  As FormatSettings
End Type
Private this        As TStaticTable


Private Sub Class_Initialize()
    With this
        .row = 1
        .column = 1

        Set .ws = ThisWorkbook.ActiveSheet
        Set .ElementsFormat = New FormatSettings
        With .ElementsFormat
            .BgColor = RGB(249, 244, 245)
            .FontColor = vbBlack
            .FontName = "Tw Cen MT"
            .FontSize = 14
        End With
    End With
    Set HeadersComp = New Headers

End Sub

Public Sub init(ByVal row As Integer, _
        ByVal column As Integer, _
        ByVal ws As Worksheet, _
        ByVal Headers As Variant _
        )
    With this
        .row = row
        .column = column
        Set .ws = ws
        With HeadersComp
            .values = Headers
        End With

        Update
    End With


End Sub

Public Sub Update()
    UpdateRanges
    With this
        .ws.Select
        HeadersComp.Update .ws
        ElementsUpdate
        .ws.Columns.AutoFit
    End With
    
End Sub

Public Sub ClearPrevRange()
    If Not this.PreviousElemRange Is Nothing Then
        With this.PreviousElemRange
            .ClearContents
            .ClearFormats
            .Interior.Color = RGB(17, 21, 28)
        End With
    End If
End Sub

Public Sub ElementsUpdate()
    
    With this
        Call ApplyFormat(.ElementsRange, .ElementsFormat)
        
        .ElementsRange.Sort key1:=.ElementsRange.Columns(UBound(HeadersComp.values) + 1), _
      order1:=xlDescending, Header:=xlNo
      
        Set .PreviousElemRange = .ElementsRange
        
        .range.EntireRow.RowHeight = 40
    End With

End Sub


'Dependencies: Row,Column, _
 Headers.values,ElementsCounter

Private Sub UpdateRanges()
    With HeadersComp
        Set .range = this.ws.range( _
                this.ws.Cells(this.row, this.column), _
                this.ws.Cells(this.row, this.column + UBound(.values)) _
                )
    End With

    With this
        Dim EmptyElements As Integer
        If .ElementsCounter = 0 Then
            EmptyElements = 1
        End If
        Set .ElementsRange = this.ws.range( _
                this.ws.Cells(this.row + 1, this.column), _
                this.ws.Cells(this.row + EmptyElements + .ElementsCounter, this.column + UBound(HeadersComp.values)) _
                )
        Set .range = Union( _
                HeadersComp.range, _
                .ElementsRange _
                )
    End With
End Sub



Private Sub HeadersComp_HeadersUpdated(ByVal Headers As Variant)
    UpdateRanges
End Sub

Public Property Let row(ByVal value As Integer)
    this.row = value
    UpdateRanges
End Property


Public Property Let column(ByVal value As Integer)
    this.column = value
    UpdateRanges
End Property

Public Property Let ElementsCounter(ByVal RHS As Integer)
    this.ElementsCounter = RHS
    Update
End Property


Public Property Get Headers() As Headers
    Set Headers = this.Headers
End Property

Public Property Set Headers(ByVal RHS As Headers)
    Set this.Headers = RHS
End Property

Public Property Get ElementsRange() As range
    Set ElementsRange = this.ElementsRange
End Property

Public Property Set ElementsRange(ByVal RHS As range)
    Set this.ElementsRange = RHS
    UpdateRanges
End Property

Public Property Get PreviousElemRange() As range
    Set PreviousElemRange = this.PreviousElemRange
End Property

Public Property Set PreviousElemRange(ByVal RHS As range)
    Set this.PreviousElemRange = RHS
End Property

Public Property Get ElementsCounter() As Integer
    ElementsCounter = this.ElementsCounter
End Property

Public Property Get range() As range
    Set range = this.range
End Property

Public Property Set range(ByVal value As range)
    Set this.range = value
End Property

Public Property Get row() As Integer
    row = this.row
End Property



Public Property Get ws() As Worksheet
    Set ws = this.ws
End Property

Public Property Set ws(ByVal value As Worksheet)
    Set this.ws = value
End Property

Public Property Get column() As Integer
    column = this.column
End Property


Public Property Get ElementsFormat() As FormatSettings
    Set ElementsFormat = this.ElementsFormat
End Property

Public Property Set ElementsFormat(ByVal RHS As FormatSettings)
    Set this.ElementsFormat = RHS
End Property

Private Sub IDataTable_Update()
    err.Raise 5                                  'TODO implement interface member
End Sub


