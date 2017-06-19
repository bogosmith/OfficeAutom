Set wmi = GetObject("winmgmts://./root/cimv2")
For Each xl In wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'excel.exe'")
  Wscript.Echo("Next..")
  xl.Terminate
Next