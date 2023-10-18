Attribute VB_Name = "PublicFunctions"
Public Sub ApplyFormat(ByVal rng As range, ByVal format As FormatSettings)

    With rng
        .Interior.Color = format.BgColor
        .Font.Name = format.FontName
        .Font.Size = format.FontSize
        .Font.Color = format.FontColor
        .HorizontalAlignment = format.HAlign
        .VerticalAlignment = format.VAlign
        .Borders.Weight = format.BorderWeight
        .Borders.LineStyle = format.BorderStyle
        .NumberFormat = format.NumberFormat
    End With
End Sub
Sub LabelTest()
    'TODO LIST
    'Label Header Custom Format Static
    'Implement Lables in Model Obj
    'AutoCalculate Label on Item
    'DATATABLE
    'Implement New ItemCode
    'List Table -Change Update Elements Collection
    'List Table - DeleteDoubleClick + Confirmation Msg
    '-- Search in a Collection On Id
    '-- headers Mapper
    
    
    Dim Product As New item
    Dim Company As New Company
    Dim Lab As New Label
    With ThisWorkbook.Sheets("LabelSheet")
        .Cells.ClearContents
        .Cells.ClearFormats
    End With
    Lab.ToRange
    
    
End Sub
Sub MergeRange()
    Dim cell As range: Set cell = ThisWorkbook.Sheets("LabelSheet").Cells(5, 5)
    Dim Direction As String
    Dim Places As Integer

    Direction = "L"
    Places = 3

    Select Case Direction 'R, L, U, D
        Case "R"
            Set cell = cell.Resize(1, 1 + Places)
        Case "L"
            Set cell = cell.offset(0, -Places).Resize(1, Places + 1)
        Case "U"
            Set cell = cell.offset(-Places, 0).Resize
        Case "D"
            Set cell = cell.Resize(1 + Places, 1)
    End Select

    ' Merge the resulting range
    cell.Merge
    'Debug.Assert cell.Address
    Debug.Print cell.Address
End Sub

Function BubbleSort(arr As Variant) As Variant
    Dim i As Long, j As Long
    Dim temp        As Double
    Dim n           As Long

    n = UBound(arr)

    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j) > arr(j + 1) Then
                ' Swap arr(j) and arr(j + 1)
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i

    BubbleSort = arr
End Function


Sub test()
    Dim arr1, arr2  As Variant
    arr1 = Array(1, 4, 5, 7, 8, 123, 9)
    arr2 = Array(3, 5)
    Dim S, E        As Double
    S = Timer

    arr3 = FindMissCodeId(arr1, arr2)

    E = Timer
    Debug.Print "Performance - FindMissing:", E - S, "sec"



End Sub

Public Function CodeBuilder(ByVal CodeLet As String, ByVal CodeId As Integer) As String
    Dim Code        As String

    Select Case CodeId
        Case CodeId < 10
            Code = CodeLet & "00" & CodeId
        Case CodeId < 100
            Code = CodeLet & "0" & CodeId
        Case Else
            Code = CodeLet & CodeId
    End Select

    CodeBuilder = Code

End Function

Public Function ExtractNumber(ByVal inputString As String) As Variant

    Dim NumReg      As New RegExp
    With NumReg
        .Global = True
        .IgnoreCase = True
        .Pattern = "^(\w+\d+)"
    End With

    If NumReg.test(inputString) Then
        Set ExtractNumber = NumReg.Execute(inputString)(0)
    Else
        ExtractNumber = "Not Matches Found in Input"
    End If

End Function

Public Function ExtractLetter(ByVal inputString As String) As Variant

    Dim LetReg      As New RegExp
    With LetReg
        .Global = True
        .IgnoreCase = True
        .Pattern = "^([a-zA-Z]+)"
    End With

    If LetReg.test(inputString) Then
        ExtractLetter = UCase(LetReg.Execute(inputString)(0))
    Else
        ExtractLetter = "Not Matches Found in Input"
    End If

End Function
'Deletes Duplicates
Public Function JoinArrays(ByVal MainNumArr As Variant, _
        ByVal OptionalNumArr As Variant) As Variant

    If IsEmpty(MainNumArr) Then
        FindMissCodeId = 0
        Exit Function
    End If

    Dim CombinedArray As Variant
    Dim i As Long, j As Long, k As Long
    Dim isDuplicate As Boolean

    ' Determine the size of the combined array
    ReDim CombinedArray(0 To UBound(MainNumArr) + UBound(OptionalNumArr) + 1)

    ' Merge both arrays into combinedArray
    
    
    'Copy InitialArray to CombinedArray
    For i = LBound(MainNumArr) To UBound(MainNumArr)
        CombinedArray(i) = MainNumArr(i)

    Next i
    'Next IdElem after MainArrayCopied / Unique Elem Counter
    k = UBound(MainNumArr) + 1

    'Duplicates Deletion
    For i = LBound(OptionalNumArr) To UBound(OptionalNumArr)
        isDuplicate = False
        For j = LBound(MainNumArr) To UBound(MainNumArr)
            If OptionalNumArr(i) = MainNumArr(j) Then
                isDuplicate = True
                Exit For
            End If
        Next j

        If Not isDuplicate Then
            CombinedArray(k) = OptionalNumArr(i)
            k = k + 1
        End If
    Next i

    ' Redimension the array to the actual size
    ReDim Preserve CombinedArray(1 To k)  '- 1)
    JoinArrays = CombinedArray
End Function
Public Function FindMissing(ByVal MainNumArr As Variant, _
        Optional ByVal OptionalNumArr As Variant = Empty) As Integer

    Dim CombinedArray As Variant
    ' Sort the merged and deduplicated array using Bubble Sort
    CombinedArray = BubbleSort(JoinArrays())
    Dim elem        As Integer
    For elem = 1 To UBound(CombinedArray)
        Debug.Print CombinedArray(elem)

    Next elem
    ' Iterate for First Missing Id
    For i = LBound(CombinedArray) To UBound(CombinedArray)
    Next i

End Function

Function FindLatestXLSXFile(ByVal pathDir As String) As String

    Dim fileSystem  As New FileSystemObject
    Dim folderPath  As String
    Dim latestFile  As String
    Dim latestDate  As Date
    Dim file        As Object

    ' Replace "C:\Users\EQUIPO\Downloads" with the path to your data_path directory
    folderPath = pathDir
    ' Initialize variables to hold the latest file information
    latestFile = ""
    latestDate = DateSerial(1900, 1, 1)

    ' Loop through each file in the directory
    For Each file In fileSystem.GetFolder(folderPath).Files
        ' Check if the file is an XLSX file and compare its last modified date
        If LCase(Right(file.Name, 5)) = ".xlsx" And file.DateLastModified > latestDate Then
            latestFile = file.Path
            latestDate = file.DateLastModified
        End If
    Next file
    'Debug.Print latestFile
    FindLatestXLSXFile = latestFile
End Function



