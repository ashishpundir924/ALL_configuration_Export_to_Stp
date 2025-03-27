Option Explicit

Sub ExportConfigsToSTEP_UniqueNames()
    Dim swApp As Object
    Dim swModel As Object
    Dim swConfMgr As Object
    Dim configNames As Variant
    Dim configName As Variant
    Dim sourceFolder As String
    Dim destFolder As String
    Dim fileName As String
    Dim filePath As String
    Dim outputPath As String
    Dim outputName As String
    Dim errors As Long, warnings As Long
    Dim fileList As Collection
    Dim f As Variant

    ' Ask for source folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing .SLDPRT Files"
        If .Show <> -1 Then Exit Sub
        sourceFolder = .SelectedItems(1)
    End With

    ' Ask for destination folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Destination Folder for STEP Files"
        If .Show <> -1 Then Exit Sub
        destFolder = .SelectedItems(1)
    End With

    ' Get list of files (safe method)
    Set fileList = GetFileList(sourceFolder, "*.sldprt")
    If fileList.count = 0 Then
        MsgBox "No .sldprt files found.", vbExclamation
        Exit Sub
    End If

    ' Launch SolidWorks
    On Error Resume Next
    Set swApp = CreateObject("SldWorks.Application")
    On Error GoTo 0

    If swApp Is Nothing Then
        MsgBox "Could not start SolidWorks.", vbCritical
        Exit Sub
    End If

    swApp.Visible = True

    ' Loop through files from list (no Dir())
    For Each f In fileList
        fileName = Mid(f, InStrRev(f, "\") + 1)
        filePath = f
        Debug.Print "Processing: " & filePath

        Set swModel = swApp.OpenDoc6(filePath, 1, 64, "", errors, warnings)
        If Not swModel Is Nothing Then
            Set swConfMgr = swModel.ConfigurationManager
            configNames = swModel.GetConfigurationNames()

            If UBound(configNames) = 0 Then
                swModel.ShowConfiguration2 configNames(0)
                outputName = Left(fileName, Len(fileName) - 7)
                outputPath = GetUniqueFilePath(destFolder, outputName & ".stp")
                swModel.Extension.SaveAs outputPath, 0, 1, Nothing, errors, warnings
                Debug.Print "Exported: " & outputPath
            Else
                For Each configName In configNames
                    swModel.ShowConfiguration2 configName
                    outputPath = GetUniqueFilePath(destFolder, configName & ".stp")
                    swModel.Extension.SaveAs outputPath, 0, 1, Nothing, errors, warnings
                    Debug.Print "Exported: " & outputPath
                Next configName
            End If

            swApp.CloseDoc fileName
        Else
            Debug.Print "Failed to open: " & filePath
        End If
    Next f

    MsgBox "Export complete!", vbInformation
End Sub

Function GetFileList(folder As String, pattern As String) As Collection
    Dim col As New Collection
    Dim file As String
    file = Dir(folder & "\" & pattern)
    Do While file <> ""
        col.Add folder & "\" & file
        file = Dir()
    Loop
    Set GetFileList = col
End Function

Function GetUniqueFilePath(folder As String, baseFile As String) As String
    Dim fPath As String
    Dim count As Integer
    Dim nameOnly As String
    Dim ext As String

    nameOnly = Left(baseFile, InStrRev(baseFile, ".") - 1)
    ext = Mid(baseFile, InStrRev(baseFile, "."))
    fPath = folder & "\" & baseFile

    count = 1
    Do While Dir(fPath) <> ""
        fPath = folder & "\" & nameOnly & "_" & Format(count, "00") & ext
        count = count + 1
    Loop

    GetUniqueFilePath = fPath
End Function

