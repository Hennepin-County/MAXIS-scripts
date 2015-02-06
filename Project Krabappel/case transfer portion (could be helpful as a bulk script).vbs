'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect "" 

workers_to_XFER_cases_to = "x102b82"
case_number_array = array("212852", "212853", "212854", "212855", "212856", "212857", "212858")

'========================================================================TRANSFER CASES========================================================================
'Creates an array of the workers selected in the dialog
workers_to_XFER_cases_to = split(replace(workers_to_XFER_cases_to, " ", ""), ",")

'Creates a new two-dimensional array for assigning a worker to each case_number
Dim transfer_array()
ReDim transfer_array(ubound(case_number_array), 1)

'Assigns a case_number to each row in the first column of the array
For x = 0 to ubound(case_number_array)
	transfer_array(x, 0) = case_number_array(x)
Next

'Reassigning x as a 0 for the following do...loop
x = 0

'Assigning y as 0, to be used by the following do...loop for deciding which worker gets which case
y = 0

'Now, it'll assign a worker to each case number in the transfer_array. Does this on a loop so that a worker can get multiple cases if that is indicated.
Do
	transfer_array(x, 1) = workers_to_XFER_cases_to(y)	'Assigns column 2 of the array to a worker in the workers_to_XFER_cases_to array
	x = x + 1											'Adds +1 to X
	y = y + 1											'Adds +1 to Y
	If y > ubound(workers_to_XFER_cases_to) then y = 0	'Resets to allow the first worker in the array to get anonther one
Loop until x > ubound(case_number_array)

'--------Now, the array is two columns (case_number, worker_assigned)!

'Script must figure out who the current worker is, and what agency they are with. This is vital because transferring within an agency uses different screens than inter-agency.
	'To do this, the script will start by analysing the current worker in REPT/ACTV.
call navigate_to_screen("REPT", "ACTV")			'Navigates to ACTV
EMReadScreen current_user, 7, 21, 13			'Reads current user, which will be reused later on to determine if the agency changes or not
If ucase(left(current_user, 2)) = "PW" then		'Needs special handling for DHS staff (we don't use x1 numbers, we use PW numbers, and the numbers become unique in the 3rd character of the string)
	XFER_chars_to_compare = 2
Else
	XFER_chars_to_compare = 4
End if

'Resetting "x" to be a zero placeholder for the following for...next
x = 0

'Now we actually transfer the cases. This for...next does the work (details in comments below)
For x = 0 to ubound(case_number_array)		'case_number_array is the same as the first col of the transfer_array
	'Assigns the number from the array to the case_number variable
	case_number = transfer_array(x, 0)
	
	'Determines interagency transfers by comparing the current active user (gathered above) to the user in the transfer array.
	If ucase(left(transfer_array(x, 1), XFER_chars_to_compare)) = ucase(left(current_user, XFER_chars_to_compare)) then
		county_to_county_XFER = False
	Else
		county_to_county_XFER = True
	End if

	'Now to transfer the cases.
	If county_to_county_XFER = False then
		call navigate_to_screen("SPEC", "XFER")
		EMWriteScreen "x", 7, 16
		transmit
		PF9
		EMWriteScreen transfer_array(x, 1), 18, 61
		transmit
		transmit
	Else
		call navigate_to_screen("SPEC", "XFER")
		EMWriteScreen "x", 9, 16
		transmit
		PF9
		call create_MAXIS_friendly_date(date, 0, 4, 28)
		call create_MAXIS_friendly_date(date, 0, 4, 61)
		EMWriteScreen "N", 5, 28
		call create_MAXIS_friendly_date(date, 0, 5, 61)
		EMWriteScreen transfer_array(x, 1), 18, 61
		transmit
		transmit
	End if
Next