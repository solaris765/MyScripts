set Scripts = Array("block-telemetry.ps1", "fix-privacy-settings.ps1", "optimize-user-interface.ps1", "optimize-windows-update.ps1", "remove-default-apps.ps1")




Set objArgs = Wscript.Arguments

If WScript.Arguments.length = 0 Then
   set objdir = WScript.CreateObject ("WScript.Shell")
   Set objShell = CreateObject("Shell.Application")
   'Pass a bogus argument, say [ uac]
   objShell.ShellExecute "wscript.exe", Chr(34) & _
      WScript.ScriptFullName & Chr(34) & " " & objdir.CurrentDirectory, "", "runas", 1
Else
 set objShell = WScript.CreateObject ("WScript.Shell")
 Directory = "-executionpolicy bypass -file " & objArgs(0) & "\Debloat-Windows-10\scripts\"
 
 'msgbox directory
 Set oShell = CreateObject("Shell.Application")  

 oShell.ShellExecute "powershell", "-executionpolicy bypass -file ExecutionPolicy.ps1 -policy unrestricted", "", "runas", 1  
 
 For Each Script in Scripts 
  oShell.ShellExecute "powershell", Directory & Script
 Next
 
 oShell.ShellExecute "powershell", "-executionpolicy bypass -file ExecutionPolicy.ps1 -policy restricted", "", "runas", 1 
 
 
 
End If
               
