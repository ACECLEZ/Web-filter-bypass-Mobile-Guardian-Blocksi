Option Explicit

Dim continueProgram
continueProgram = True

Dim ie

Dim password
password = "sake"
Dim enteredPassword

Do While continueProgram
If enteredPassword = "" Then
enteredPassword = InputBox("Enter Password:" & vbCrLf & vbCrLf & _
                     "https://github.com/ACECLEZ/blocksi-bypass/", "Sake Injector Login - {Non Root 2.7.2}")


End If
If enteredPassword = password Then
    Dim input
    input = InputBox("Please enter a command:" & vbCrLf & vbCrLf & _
                     "1 - Current Session Bypass [Blocksi]" & vbCrLf & _
		     "2 - Current Session Bypass [Mobile Guardian]" & vbCrLf & _
                     "3 - Auto Bypass On Startup [Blocksi]" & vbCrLf & _
                     "4 - Disable Auto Bypass [Blocksi]" & vbCrLf & _
                     "5 - Activate Proxy" & vbCrLf & _
	             "6 - Disconnect Proxy" & vbCrLf & _
                     "7 - Exit the program", "Sake Injector - Non Root - Non Jailbreak - V2.7.2")

    Select Case input
        Case "1"
            Dim bypassManualPath
            bypassManualPath = ".\Bypass-Manual.vbs" ' Replace with the actual path to Bypass-Manual.vbs
            CreateObject("WScript.Shell").Run bypassManualPath, 0, False

        Case "2"
            Dim bypassManualPathMG
            bypassManualPathMG = ".\MG-Bypass-Manual.vbs" ' Replace with the actual path to Bypass-Manual.vbs
            CreateObject("WScript.Shell").Run bypassManualPathMG, 0, False

        Case "3"
            Dim bypassAutoPath
            bypassAutoPath = ".\Bypass-Auto.vbs" ' Replace with the actual path to Bypass-Auto.vbs
            CreateObject("WScript.Shell").Run bypassAutoPath, 0, False

        Case "4"
            Dim disableAutoPath
            disableAutoPath = ".\Disable-Auto.vbs" ' Replace with the actual path to Disable-Auto.vbs
            CreateObject("WScript.Shell").Run disableAutoPath, 0, False

        Case "5"
            Dim proxyPath
            proxyPath = ".\Proxy.vbs" ' Replace with the actual path to Proxy.vbs
            CreateObject("WScript.Shell").Run proxyPath, 0, False

	Case "6"

            CreateObject("WScript.Shell").Run proxyPath, 0, False

        Case "7"
            continueProgram = False

        Case Else
            MsgBox "Invalid command. Please try again."
    End Select
Else
    MsgBox "Incorrect password. Exiting the program."
    Exit Do
End If

If continueProgram Then
    Dim continueResponse
    continueResponse = MsgBox("Do you want to continue using the program?", vbQuestion + vbYesNo)
    If continueResponse = vbNo Then
        continueProgram = False
    End If
End If
Loop