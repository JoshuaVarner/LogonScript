Option Explicit
Dim objNetwork, strUsername

' Create the WScript.Network object
Set objNetwork = CreateObject("WScript.Network")

' Get the username of the current logged-in user
strUsername = objNetwork.UserName

If strUsername = strUsername Then

    Dim objShell, strDocumentsPath, strMusicPath, strVideosPath, strFavoritesPath, strRegKey

    Set objShell = CreateObject("WScript.Shell")

    strRegKey = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\"

    strDocumentsPath = objShell.RegRead(strRegKey & "Personal")
    strMusicPath = objShell.RegRead(strRegKey & "My Music")
    strVideosPath = objShell.RegRead(strRegKey & "My Video")
    strFavoritesPath = objShell.RegRead(strRegKey & "Favorites")

    Dim objFSO

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Replace the placeholders with your actual source and destination folder paths
    CopyFolderContents "C:\Source\CP\" & strUsername & "\Documents", strDocumentsPath
    CopyFolderContents "C:\Source\CP\" & strUsername & "\Music", strMusicPath
    CopyFolderContents "C:\Source\CP\" & strUsername & "\Videos", strVideosPath
    CopyFolderContents "C:\Source\CP\" & strUsername & "\Favorites", strFavoritesPath

    Sub CopyFolderContents(srcFolder, destFolder)
    If objFSO.FolderExists(srcFolder) And objFSO.FolderExists(destFolder) Then
        Dim objSrcFolder, objFile, objSubFolder

        Set objSrcFolder = objFSO.GetFolder(srcFolder)

        ' Copy all files
        For Each objFile In objSrcFolder.Files
            objFSO.CopyFile objFile.Path, destFolder & "\" & objFile.Name, True
        Next

        ' Copy all subfolders
        For Each objSubFolder In objSrcFolder.SubFolders
            objFSO.CopyFolder objSubFolder.Path, destFolder & "\" & objSubFolder.Name, True
        Next

        'WScript.Echo "Successfully copied files from " & srcFolder & " to " & destFolder
    ElseIf Not objFSO.FolderExists(srcFolder) Then
        'WScript.Echo "Source folder not found: " & srcFolder
        WScript.Quit
    ElseIf Not objFSO.FolderExists(destFolder) Then
        'WScript.Echo "Destination folder not found: " & destFolder
        WScript.Quit
    End If
End Sub
Else
    WScript.Quit
End If

Set objNetwork = Nothing
Set objShell = Nothing
Set objFSO = Nothing
