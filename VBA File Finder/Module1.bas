Attribute VB_Name = "Module1"


Public temp() As String

Function ListFiles(FileName As String, FolderPath As String)
Dim k As Long, i As Long
ReDim temp(2, 0)
If Right(FolderPath, 1) <> "\" Then
    FolderPath = FolderPath & "\"
End If
Recursive FileName, FolderPath
k = Range(Application.Caller.Address).Rows.Count

If k < UBound(temp, 2) Then
Else
    For i = UBound(temp, 2) To k
          ReDim Preserve temp(UBound(temp, 1), i)
            temp(0, i) = ""
            temp(1, i) = ""
            temp(2, i) = ""
    Next i
End If
ListFiles = Application.Transpose(temp)
ReDim temp(0)
End Function

Function Recursive(FileName As String, FolderPath As String)
Dim Value As String, Folders() As String
Dim Folder As Variant, a As Long
ReDim Folders(0)
If Right(FolderPath, 2) = "\\" Then Exit Function
Value = Dir(FolderPath, &H1F)
Do Until Value = ""
    If Value = "." Or Value = ".." Then
    Else
        If GetAttr(FolderPath & Value) = 16 Then
            Folders(UBound(Folders)) = Value
            ReDim Preserve Folders(UBound(Folders) + 1)
        Else
            If Value = FileName Then
                temp(0, UBound(temp, 2)) = FolderPath
                temp(1, UBound(temp, 2)) = Value
                temp(2, UBound(temp, 2)) = FileLen(FolderPath & Value)
                ReDim Preserve temp(UBound(temp, 1), UBound(temp, 2) + 1)
            End If
        End If
    End If
    Value = Dir
Loop
For Each Folder In Folders
    Recursive FileName, FolderPath & Folder & "\"
Next Folder
End Function

