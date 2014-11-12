EMConnect ""

start_time = timer

For i = 0 to 100
	EMSendKey "<PF4>"
	EMWaitReady -1, 0

	EMSendKey "<enter>"
	EMWaitReady -1, 0

	EMSendKey "<PF3>"
	EMWaitReady -1, 0
Next

end_time = timer
minus_one_time = end_time - start_time

start_time = timer

For i = 0 to 100
	EMSendKey "<PF4>"
	EMWaitReady 0, 0

	EMSendKey "<enter>"
	EMWaitReady 0, 0

	EMSendKey "<PF3>"
	EMWaitReady 0, 0
Next

end_time = timer
zero_time = end_time - start_time

start_time = timer

For i = 0 to 100
	EMSendKey "<PF4>"
	EMWaitReady 1, 1

	EMSendKey "<enter>"
	EMWaitReady 1, 1

	EMSendKey "<PF3>"
	EMWaitReady 1, 1
Next

end_time = timer
one_time = end_time - start_time

MsgBox "-1: " & minus_one_time & chr(10) & _
	"0: " & zero_time & chr(10) & _
	"1: " & one_time