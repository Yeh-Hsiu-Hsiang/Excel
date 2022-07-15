Attribute VB_Name = "GetFile"
Sub GetFile()

Dim oFSO, oFolder, oFile, oFolderFile As Object
Dim i As Integer

' �إ� FileSystemObject ����
Set oFSO = CreateObject("Scripting.FileSystemObject")

' �إߥؿ�����
Set oFolder = oFSO.GetFolder("C:\Users\ywqa011\Documents\00_�u�@")

i = 2

' �H�j��C�X�Ҧ��l�ؿ�
For Each oFolderFile In oFolder.SubFolders

    Cells(i, 1) = oFolderFile.Name              ' �ɮצW��
    Cells(i, 2) = oFolderFile.Path              ' �ɮ׸��|
    Cells(i, 3) = oFolderFile.Size              ' �ɮפj�p�]�줸�ա^
    Cells(i, 4) = oFolderFile.Type              ' �ɮ�����
    Cells(i, 5) = oFolderFile.DateCreated       ' �ɮ׫إ߮ɶ�
    Cells(i, 6) = oFolderFile.DateLastAccessed  ' �ɮפW���s���ɶ�
    Cells(i, 7) = oFolderFile.DateLastAccessed  ' �ɮפW���ק�ɶ�

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


' �H�j��C�X�Ҧ��ɮ�
For Each oFile In oFolder.Files

    Cells(j, 1) = oFile.Name              ' �ɮצW��
    Cells(j, 2) = oFile.Path              ' �ɮ׸��|
    Cells(j, 3) = oFile.Size              ' �ɮפj�p�]�줸�ա^
    Cells(j, 4) = oFile.Type              ' �ɮ�����
    Cells(j, 5) = oFile.DateCreated       ' �ɮ׫إ߮ɶ�
    Cells(j, 6) = oFile.DateLastAccessed  ' �ɮפW���s���ɶ�
    Cells(j, 7) = oFile.DateLastAccessed  ' �ɮפW���ק�ɶ�

    j = j + 1

Next oFile

End Sub
