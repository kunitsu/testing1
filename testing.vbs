dim wshShell
dim strQuery, objWMIService, colItems, objItem
dim ComputerName, IpAddress

Set wshShell = CreateObject( "WScript.Shell" )
ComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

strQuery = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE MACAddress > ''"

Set objWMIService = GetObject("winmgmts://./root/CIMV2")
Set colItems = objWMIService.ExecQuery(strQuery, "WQL", 48)

For Each objItem In colItems
    If IsArray(objItem.IpAddress) Then 
	    If UBound(objItem.IpAddress) = 0 Then
		    IpAddress = objItem.IpAddress(0)
		Else
		    IpAddress = Join(objItem.IpAddress, chr(13) & chr(10))
		End If
	End If
Next

WScript.Echo "Computername := " & ComputerName & chr(13) & chr(10) & "IP Adddress := " & IpAddress