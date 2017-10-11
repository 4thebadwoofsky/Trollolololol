Set oWMP = CreateObject("WMPlayer.OCX.7")
Set wshShell =wscript.CreateObject("WScript.Shell")
Set colCDROMs = oWMP.cdromCollection
Do
	if colCDROMs.Count >= 1 then
		For i = 0 to colCDROMs.Count - 1
			colCDROMs.Item(i).Eject
		Next
		For i = 0 to colCDROMs.Count - 1
			colCDROMs.Item(i).Eject
		Next
	End If
	wscript.sleep 10000
Loop