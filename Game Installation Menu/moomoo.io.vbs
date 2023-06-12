Dim objShell, objShortcut, profileNumber

' Get the profile number
profileNumber = GetProfileNumber()

Set objShell = CreateObject("WScript.Shell")
Set objShortcut = objShell.CreateShortcut(objShell.SpecialFolders("Desktop") & "\Sake_Injector_Moomoo.io_FRVR.lnk")

' Set the target path
objShortcut.TargetPath = "C:\Program Files\Google\Chrome\Application\chrome_proxy.exe"
' Set the command line arguments
objShortcut.Arguments = "--profile-directory=""Profile " & profileNumber & """ --app-id=ombagdmbnldpfnooolcabocfbcdbnohm"
' Set the working directory
objShortcut.WorkingDirectory = "C:\Program Files\Google\Chrome\Application"
' Set the icon path
objShortcut.IconLocation = "%USERPROFILE%\AppData\Local\Google\Chrome\User Data\Profile " & profileNumber & "\Web Applications\_crx_ombagdmbnldpfnooolcabocfbcdbnohm\MooMoo FRVR.ico"

' Save the shortcut
objShortcut.Save

' Clean up objects
Set objShortcut = Nothing
Set objShell = Nothing

Function GetProfileNumber()
    Dim objFSO, profilePath, profileNumber

    ' Get the path to the Chrome profile directory
    profilePath = GetChromeProfilePath()

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Check if the profile directory exists
    If objFSO.FolderExists(profilePath) Then
        ' Loop through all subfolders in the profile directory
        For Each objSubFolder In objFSO.GetFolder(profilePath).SubFolders
            ' Check if the subfolder name starts with "Profile "
            If Left(objSubFolder.Name, 8) = "Profile " Then
                profileNumber = Mid(objSubFolder.Name, 9)
                Exit For
            End If
        Next
    End If

    ' Return the profile number
    GetProfileNumber = profileNumber
End Function

Function GetChromeProfilePath()
    Dim objShell, userProfile, profilePath

    Set objShell = CreateObject("WScript.Shell")
    userProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

    ' Construct the Chrome profile path
    profilePath = userProfile & "\AppData\Local\Google\Chrome\User Data\"

    ' Return the profile path
    GetChromeProfilePath = profilePath
End Function
