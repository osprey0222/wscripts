
Set WshShell = WScript.CreateObject("WScript.Shell")
'WshShell.Run "notepad", 9 

Dim sInput
sInput = InputBox("Enter minutes:")
'MsgBox "You entered:" & sInput

WshShell.SendKeys "^%{TAB}"

For i = 1 To 60*sInput
    WScript.Sleep 1000
    WshShell.SendKeys "^%{TAB}"
    Rem WshShell.SendKeys Mid(Msg, i, 1)World!
Next