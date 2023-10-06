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


Function FindLatestXLSXFile(ByVal pathDir As String) As String

    Dim fileSystem As New FileSystemObject
    Dim folderPath As String
    Dim latestFile As String
    Dim latestDate As Date
    Dim file As Object
    
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


