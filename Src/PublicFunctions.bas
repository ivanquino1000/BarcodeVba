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

Public Sub CloseAllVBEWindows()
    Dim wk As Workbook
    Dim CodeWindow As Variant
    On Error GoTo err
    'If VBE is closed the user shall allow it to open first
    If Application.VBE.MainWindow.Visible = False Then
        If MsgBox("VBE is still closed!" & vbCrLf & "The operation requires VBE to be open during its process. Would you like to open it?", vbInformation + vbYesNo) = vbYes Then
            Application.VBE.MainWindow.Visible = True
        Else
            Exit Sub
        End If
    End If
    'Then it will close all the windows
    'Except the Default VBE Windows (Immediate Windows, Locals Windows ...)
    For Each wk In Application.Workbooks
        If wk.Name = ThisWorkbook.Name Then
            For Each CodeWindow In wk.VBProject.VBE.Windows
                If CodeWindow.Visible = True And CodeWindow.Type = 0 Then CodeWindow.Visible = False
            Next CodeWindow
        End If
    Next wk
    Exit Sub
err:
    Debug.Print err.Description
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


