@echo off
REM Build ao-libre server with BotAPI enabled
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe" /make ^
/d UsarQueSocket=1:ConUpTime=1:AntiExternos=0:Lac=1:DEBUGGING=0:BOT_SERVER=1 ^
/out vb6build.log ^
SERVER.VBP
type vb6build.log
