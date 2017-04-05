Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

For Each os in oss
    value = os.Caption
Next

 If (OSBits() = "64") then
 value = value & " x64"
 End If

 msgbox value
 
Function OSBits ()
 Dim WshShell
 Set WshShell = CreateObject("WScript.Shell")
 OSBits = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
End Function