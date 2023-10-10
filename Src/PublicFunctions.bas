Attribute VB_Name = "PublicFunctions"
Public Sub ApplyFormat(ByVal rng As Range, ByVal Format As FormatSettings)

    With rng
        .Interior.Color = Format.BgColor
        .font.Name = Format.FontName
        .font.Size = Format.FontSize
        .font.Color = Format.FontColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With

End Sub


Function BubbleSort(arr() As Double) As Double()
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

    arr3 = FindMissCodeId(arr1, arr2)

End Sub

Public Function CodeBuilder(ByVal CodeLet As String, ByVal CodeId As Integer) As String
    Dim Code As String
    
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
    
    Dim NumReg     As New RegExp
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
    
    Dim LetReg     As New RegExp
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
Public Function FindMissCodeId(ByVal MainNumArr As Variant, _
        Optional ByVal OptionalNumArr As Variant = Empty) As Integer

    If IsEmpty(MainNumArr) Then
        FindMissCodeId = 0
        Exit Function
    End If

    Dim combinedArray() As Double
    Dim i As Long, j As Long, k As Long
    Dim isDuplicate As Boolean

    ' Determine the size of the combined array
    ReDim combinedArray(0 To UBound(MainNumArr) + UBound(OptionalNumArr) + 1)

    ' Merge both arrays into combinedArray
    For i = LBound(MainNumArr) To UBound(MainNumArr)
        combinedArray(i) = MainNumArr(i)
    Next i

    k = UBound(MainNumArr) + 1

    For i = LBound(OptionalNumArr) To UBound(OptionalNumArr)
        isDuplicate = False
        For j = LBound(MainNumArr) To UBound(MainNumArr)
            If OptionalNumArr(i) = MainNumArr(j) Then
                isDuplicate = True
                Exit For
            End If
        Next j

        If Not isDuplicate Then
            combinedArray(k) = OptionalNumArr(i)
            k = k + 1
        End If
    Next i

    ' Redimension the array to the actual size
    ReDim Preserve combinedArray(1 To k + 1) '- 1)

    ' Sort the merged and deduplicated array using Bubble Sort
    combinedArray = BubbleSort(combinedArray)

    ' Return the final sorted array
    'MergeAndSortArrays = combinedArray

    '    Dim FixedNumArr As Variant
    '    'Join - Sort Arr Process
    '    If Not IsEmpty(OptionalNumArr) Then
    '        'New Arr Index
    '        Dim i, j    As Integer
    '        'New Arr Index
    '        Dim l, m, h As Integer
    '        'Initial Main Array Index
    '        Dim initMax As Integer
    '        initMax = UBound(MainNumArr)
    '        l = LBound(MainNumArr)
    '        m = UBound(MainNumArr) + 1
    '        h = UBound(MainNumArr) + UBound(OptionalNumArr) + 1
    '        i = m
    '        ReDim Preserve MainNumArr(l To h)
    '
    '        Do While i <= h
    '            MainNumArr(i) = OptionalNumArr(j)
    '            i = i + 1
    '            j = j + 1
    '        Loop
    '        MainNumArr = BubbleSort(MainNumArr)
    '    End If

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


