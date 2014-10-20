'Currently, the custom function derails at line 3 for seemingly no reason.

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'----------DIALOGS----------
BeginDialog tikl_dialog, 0, 0, 191, 90, "TIKL"
  EditBox 110, 5, 70, 15, case_number
  EditBox 110, 25, 15, 15, tikl_month
  EditBox 130, 25, 15, 15, tikl_day
  EditBox 150, 25, 15, 15, tikl_year
  EditBox 55, 45, 125, 15, tikl_text
  ButtonGroup ButtonPressed
    OkButton 45, 70, 50, 15
    CancelButton 100, 70, 50, 15
  Text 10, 10, 50, 10, "Case Number"
  Text 10, 30, 90, 10, "TIKL Date (MM DD YYYY)"
  Text 10, 50, 35, 10, "TIKL Text"
EndDialog

EMConnect ""
maxis_check_function

call find_variable ("Case Nbr: ", case_number, 8)
	case_number = replace(case_number, "_", "")
	IF case_number = "" THEN
		call find_variable ("Case Number: ", case_number, 8)
		case_number = replace(case_number, "_", "")
	END IF
	
DO
	DO
	DIALOG tikl_dialog
		IF ButtonPressed = 0 THEN stopscript
		tikl_date = cdate(tikl_month & "/" & tikl_day & "/" & tikl_year)
		IF isdate(tikl_date) = FALSE THEN MsgBox "Please enter a valid date (MM DD YYYY)."
		IF datediff("D", date, tikl_date) < 0 THEN MSGBOX "You must set a TIKL date NOT in the past."
		IF len(tikl_text) > 253 THEN MSGBox "Your TIKL message is too long. A TIKL can be 253 characters and this TIKL is " & len(tikl_text) & " characters."
	LOOP WHILE len(tikl_text) > 253
LOOP UNTIL isdate(tikl_date) = TRUE AND case_number <> "" AND (datediff("D", date, tikl_date) >= 0)

IF len(tikl_month) = 1 THEN tikl_month = "0" & tikl_month
IF len(tikl_day) = 1 THEN tikl_day = "0" & tikl_day
IF len(tikl_year) = 4 THEN tikl_year = right(tikl_year, 2)

call navigate_to_screen("DAIL", "WRIT")
EMWriteScreen tikl_month, 5, 18
EMWriteScreen tikl_day, 5, 21
EMWriteScreen tikl_year, 5, 24
write_TIKL_function
MSGBox "PF3 to approve TIKL and exit."

