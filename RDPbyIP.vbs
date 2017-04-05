Option Explicit
'~ On Error Resume Nex
RequireAdmin

'User Settings
const FILENAME = "GrowPortbyIP"

'----------------------------------------------------------------------------------------------------------------
'IP and Computer Name Section
Dim PortbyIP, CompIP, DeviceName

'Get IP and Computer Name
dim NIC1, Nic, StrIP, CompName
Set NIC1 =     GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
For Each Nic in NIC1
 if Nic.IPEnabled then
  Dim i, WshNetwork
  StrIP = Nic.IPAddress(i)
  
  Set WshNetwork = WScript.CreateObject("WScript.Network")
  CompName= WshNetwork.Computername
  
  if Split(StrIP, ".")(0) = "192" Then
   PortbyIP = "4" & Lpad(Split(strIP, ".")(3),3,"0")
   CompIP = strIP
   DeviceName = CompName
  End if
 End if
Next

'---------------------------------------------------------------------------------------------------------------
'Write File Section
Const ForAppending = 8
Const CreateIfNotExist = True
Const OpenAsASCII = 0

Dim strFile, objFSO, objFile, strDrive
Set objFSO = CreateObject("Scripting.FileSystemObject")

strDrive = Split(objFSO.GetParentFolderName(wscript.ScriptFullName), "\")(0)
strFile = strDrive & "\Information\" & FILENAME & ".csv"

' Open file for appending. Create if it does not exist.
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strFile, ForAppending, CreateIfNotExist, OpenAsASCII)

objFile.Write PortbyIP & "," & CompIP & "," & DeviceName
objFile.WriteLine
objFile.Close
'-----------------------------------------------------------------------------------------------------------


'FireWall Object Section
Dim objFirewall, objPolicy

Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
Dim portTag 
portTag = "Port"
'constant for UDP 17 
Dim UDP
UDP = 17
'constant for TCP 6
Dim TCP
TCP = 6
'Enable ICMP
'Set objICMPSettings = objPolicy.ICMPSettings
'objICMPSettings.AllowInboundEchoRequest = TRUE

Call addPorts(PortbyIP, null,"BePC_RDP_TCP", TCP)

'-------------------------------------------------------------------------------------------------------------

'Registry Section
Dim objReg
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

RegWrite "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp", "PortNumber", "REG_DWORD", PortbyIP

Dim objShell
set objShell = wscript.CreateObject("wscript.shell")
objShell.Run "shutdown.exe /R /T 5 /C ""Done! Rebooting your computer now."" "
'-------------------------------------------------------------------------------------------------------------

'Functions

Function addPorts(initial, final, portTag, Protocol)
 Dim objPort, colPorts, errReturn
 if IsNull(final)then
  Set objPort = CreateObject("HNetCfg.FwOpenPort")
		
		objPort.Port = initial
		objPort.Name =  portTag & " " & initial
		objPort.Protocol = Protocol
		objPort.Enabled = TRUE
		Set colPorts = objPolicy.GloballyOpenPorts
		errReturn = colPorts.Add(objPort)
 Else
    For Port= initial To final Step 1
		Set objPort = CreateObject("HNetCfg.FwOpenPort")
		objPort.Port = Port
		objPort.Name = portTag & " " & Port
		objPort.Protocol = Protocol
		objPort.Enabled = TRUE
		Set colPorts = objPolicy.GloballyOpenPorts
		errReturn = colPorts.Add(objPort)
	Next
 End If
End Function

Function LPad(s, l, c)
  Dim n : n = 0
  If l > Len(s) Then n = l - Len(s)
  LPad = String(n, c) & s
End Function

Function RegWrite(reg_keyname, reg_valuename,reg_type,ByVal reg_value)
	Dim aRegKey, Return
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegWrite = 0
		Exit Function
	End If

	Return = RegWriteKey(aRegKey)
	If Return = 0 Then
		RegWrite = 0
		Exit Function
	End If

	Select Case reg_type
		Case "REG_SZ"
			Return = objReg.SetStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case "REG_EXPAND_SZ"
			Return = objReg.SetExpandedStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case "REG_BINARY"
			If IsArray(reg_value) = 0 Then reg_value = Array()
			Return = objReg.SetBinaryValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		Case "REG_DWORD"
			If IsNumeric(reg_value) = 0 Then reg_value = 0
			Return = objReg.SetDWORDValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		Case "REG_MULTI_SZ"
			If IsArray(reg_value) = 0 Then
				If Len(reg_value) = 0 Then
					reg_value = Array()
				Else
					reg_value = Array(reg_value)
				End If
			End If
			Return = objReg.SetMultiStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		'Case "REG_QWORD"
			'Return = oReg.SetQWORDValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case Else
			RegWrite = 0
			Exit Function
	End Select

	If (Return <> 0) Or (Err.Number <> 0) Then
		RegWrite = 0
		Exit Function
	End If
	RegWrite = 1
End Function

Function RegWriteKey(RegKeyName)
	Dim Return
	If IsArray(RegKeyName) = 0 Then
		RegKeyName = RegSplitKey(RegKeyName)
	End If

	If (IsArray(RegKeyName) = 0) Or (UBound(RegKeyName) <> 1) Then
		RegWriteKey = 0
		Exit Function
	End If

	Return = objReg.CreateKey(RegKeyName(0),RegKeyName(1))
	If (Return <> 0) Or (Err.Number <> 0) Then
		RegWriteKey = 0
		Exit Function
	End If
	RegWriteKey = 1
End Function

Function RegDelete(reg_keyname, reg_valuename)
	Dim Return,aRegKey
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegDelete = 0
		Exit Function
	End If

	Return = objReg.DeleteValue(aRegKey(0),aRegKey(1),reg_valuename)
	If (Return <> 0) And (Err.Number <> 0) Then
		RegDelete = 0
		Exit Function
	End If
	RegDelete = 1
End Function

Function RegDeleteKey(reg_keyname)
	Dim Return,aRegKey
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegDeleteKey = 0
		Exit Function
	End If

	'On Error Resume Next
	Return = RegDeleteSubKey(aRegKey(0),aRegKey(1))
	'On Error Goto 0
	If Return = 0 Then
		RegDeleteKey = 0
		Exit Function
	End If
	RegDeleteKey = 1
End Function

Function RegDeleteSubKey(strRegHive, strKeyPath)
	Dim Return,arrSubkeys,strSubkey
    objReg.EnumKey strRegHive, strKeyPath, arrSubkeys
    If IsArray(arrSubkeys) <> 0 Then
        For Each strSubkey In arrSubkeys
            RegDeleteSubKey strRegHive, strKeyPath & "\" & strSubkey
        Next
    End If

	Return = objReg.DeleteKey(strRegHive, strKeyPath)
	If (Return <> 0) Or (Err.Number <> 0) Then
		RegDeleteSubKey = 0
		Exit Function
	End If
	RegDeleteSubKey = 1
End Function

Function RegSplitKey(RegKeyName)
	Dim strHive, strInstr, strLeft
	strInstr=InStr(RegKeyName,"\")
	If strInstr = 0 Then Exit Function
	strLeft=left(RegKeyName,strInstr-1)

	Select Case strLeft
		Case "HKCR","HKEY_CLASSES_ROOT"	strHive = &H80000000
		Case "HKCU","HKEY_CURRENT_USER"	strHive = &H80000001
		Case "HKLM","HKEY_LOCAL_MACHINE"	strHive = &H80000002
		Case "HKU","HKEY_USERS" 	strHive = &H80000003
		Case "HKCC","HKEY_CURRENT_CONFIG"	strHive = &H80000005
	  Case Else Exit Function
	End Select

    RegSplitKey = Array(strHive,Mid(RegKeyName,strInstr+1))
End Function

Function RequireAdmin()
	Dim reg_valuename, WShell, Cmd, CmdLine, I

	GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")_
	.EnumValues &H80000003, "S-1-5-19\Environment",  reg_valuename
	If IsArray(reg_valuename) <> 0 Then
		RequireAdmin = 1
		Exit Function
	End If

	Set Cmd = WScript.Arguments
	For I = 0 to Cmd.Count - 1
		If Cmd(I) = "/admin" Then
			Wscript.Echo "To script you must have administrator rights!"
			'RequireAdmin = 0
			'Exit Function
			WScript.Quit
		End If
		CmdLine = CmdLine & Chr(32) & Chr(34) & Cmd(I) & Chr(34)
	Next
	CmdLine = CmdLine & Chr(32) & Chr(34) & "/admin" & Chr(34)

	Set WShell= WScript.CreateObject( "WScript.Shell")
	CreateObject("Shell.Application").ShellExecute WShell.ExpandEnvironmentStrings(_
	"%SystemRoot%\System32\WScript.exe"),Chr(34) & WScript.ScriptFullName & Chr(34) & CmdLine, "", "runas"
	WScript.Quit
End Function