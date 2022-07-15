Attribute VB_Name = "GetFile"
Sub GetFile()

Dim oFSO, oFolder, oFile, oFolderFile As Object
Dim i As Integer

' 建立 FileSystemObject 物件
Set oFSO = CreateObject("Scripting.FileSystemObject")

' 建立目錄物件
Set oFolder = oFSO.GetFolder("C:\Users\ywqa011\Documents\00_工作")

i = 2

' 以迴圈列出所有子目錄
For Each oFolderFile In oFolder.SubFolders

    Cells(i, 1) = oFolderFile.Name              ' 檔案名稱
    Cells(i, 2) = oFolderFile.Path              ' 檔案路徑
    Cells(i, 3) = oFolderFile.Size              ' 檔案大小（位元組）
    Cells(i, 4) = oFolderFile.Type              ' 檔案類型
    Cells(i, 5) = oFolderFile.DateCreated       ' 檔案建立時間
    Cells(i, 6) = oFolderFile.DateLastAccessed  ' 檔案上次存取時間
    Cells(i, 7) = oFolderFile.DateLastAccessed  ' 檔案上次修改時間

    i = i + 1

Next oFolderFile



j = 1
Do While True
    If ActiveSheet.Cells(j, 1).Value = "" Then
        ActiveSheet.Cells(j, 1).Select
        Exit Do
    End If
    j = j + 1
Loop


' 以迴圈列出所有檔案
For Each oFile In oFolder.Files

    Cells(j, 1) = oFile.Name              ' 檔案名稱
    Cells(j, 2) = oFile.Path              ' 檔案路徑
    Cells(j, 3) = oFile.Size              ' 檔案大小（位元組）
    Cells(j, 4) = oFile.Type              ' 檔案類型
    Cells(j, 5) = oFile.DateCreated       ' 檔案建立時間
    Cells(j, 6) = oFile.DateLastAccessed  ' 檔案上次存取時間
    Cells(j, 7) = oFile.DateLastAccessed  ' 檔案上次修改時間

    j = j + 1

Next oFile

End Sub
