strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery _
  ("SELECT * FROM Win32_Service WHERE Name = 'Alerter'")
Wscript.Echo "Reached here.."
