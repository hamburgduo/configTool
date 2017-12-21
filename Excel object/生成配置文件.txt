


'=====================================================================
'        选择目录并给单元格赋值
'=====================================================================
 Sub filePath_Click()
    Dim FolderPicker As Object
    Dim filePath As String
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With FolderPicker
        If .Show = -1 Then
                filePath = .SelectedItems(1)
                Sheets("生成配置文件").Range("confFilePath").value = filePath
        Else
            Exit Sub
        End If
    End With
End Sub

