
		dim i
		i = 0
		set wsc  = CreateObject("WScript.Shell") 
	Do
		If i = 0 then
			'MsgBox "Gata, antilock activat"
		End If
		WScript.Sleep(4*10000)
		wsc.SendKeys("{CAPSLOCK}")
		wsc.SendKeys("{CAPSLOCK}")
		i = i + 1
	Loop