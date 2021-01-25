strSendKey = "{F13}"
intSleepTime = 290000

strQuery = "Select * FROM Win32_Process WHERE (Caption = 'wscript.exe' OR Caption = 'cscript.exe') AND " & " CommandLine LIKE '%" & WScript.ScriptName & "%'"
Set Locator = CreateObject("WbemScripting.SWbemLocator")
Set Service = Locator.ConnectServer
Set Res = Service.ExecQuery(strQuery)

Dim cnt
cnt = 0
If Res.Count > 1 Then 
strMsg = "‘—Mˆ—‚ğ’â~‚µ‚Ü‚·"
If MsgBox(strMsg , vbYesNo + vbQuestion) = vbYes Then  
For Each proc In Res
cnt = cnt + 1
If cnt <> Res.Count then 
proc.Terminate
End If  
Next 
End If 
WScript.Quit 0
Else 
strSleepTime = CStr(intSleepTime/1000) 
strMsg = strSleepTime & "•b–ˆ‚É" & strSendKey &"‚ğ‘—M‚µ‚Ü‚·B"
If MsgBox(strMsg , vbYesNo + vbQuestion) = vbYes Then  
Call StopScript(strSendKey, intSleepTime) 
End If
End If
WScript.Quit 0
Sub StopScript(key, sleepTime)
Set WsShell = CreateObject("Wscript.Shell") 
Do  
WsShell.SendKeys(key) 
WScript.Sleep sleepTime 
Loop
End Sub