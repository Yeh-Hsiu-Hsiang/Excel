Attribute VB_Name = "Get_Files_Path"
Sub loopAllSubFolderSelectStartDirectory()

'Another Macro must call LoopAllSubFolders Macro to start to procedure
Call LoopAllSubFolders("\\yeawen\files-server\08_文控\文管中心\4 SIP\航電\010")

End Sub

'List all files in sub folders
Sub LoopAllSubFolders(ByVal folderPath As String)

Dim fileName As String
Dim fullFilePath As String
Dim numFolders As Long
Dim folders() As String
Dim i As Long


If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
fileName = Dir(folderPath & "*.*", vbDirectory)

While Len(fileName) <> 0

    If Left(fileName, 1) <> "." Then
 
        fullFilePath = folderPath & fileName
 
        If (GetAttr(fullFilePath) And vbDirectory) = vbDirectory Then
            ReDim Preserve folders(0 To numFolders) As String
            folders(numFolders) = fullFilePath
            numFolders = numFolders + 1
        Else
            'Insert the actions to be performed on each file
            'This example will print the full file path to the immediate window
            
            j = 2
            Do While True
                If ActiveSheet.Cells(j, "A").Value = "" Then
                    ActiveSheet.Cells(j, "A").Select
                    Exit Do
                End If
                j = j + 1
            Loop

                Cells(j, 1) = folderPath
                
                Debug.Print folderPath

            'Debug.Print "_" & fileName
            
        End If
 
    End If
 
    fileName = Dir()

Wend

For i = 0 To numFolders - 1

    LoopAllSubFolders folders(i)
 
Next i


For k = 2 To Range("A65536").End(xlUp).Row

    If Range("A" & k) = Range("A" & k).Offset(-1, 0) And Range("A" & k) <> "" Then
        Rows(k).Select
        Selection.Delete Shift:=xlUp
        k = k - 1
    End If
Next

End Sub

Sub Del_list()

    SendKeys "^g^a{DEL}"
    
    Range("A2", Range("A65535").End(xlUp)).Select
    Selection.ClearContents
End Sub
