'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - PROJECT NOOB SCRIPT.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone



Dialog1 = "" 'blanking out dialog name
'Add dialog here: Add the dialog just before calling the dialog below unless you need it in the dialog due to using COMBO Boxes or other looping reasons. Blank out the dialog name with Dialog1 = "" before adding dialog.
'Shows dialog -----------------------------------------------------------------------------------------------------

' Add dialog to collect case number, footer month, and footer year. Include field validation.

BeginDialog Dialog1, 0, 0, 191, 105, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 80, 85, 50, 15
    CancelButton 140, 85, 50, 15
  Text 10, 10, 80, 15, "Enter the case number:"
  Text 10, 35, 95, 15, "Enter the footer month (MM):"
  Text 10, 55, 95, 15, "Enter the footer year (YY):"
  EditBox 95, 5, 40, 15, MAXIS_case_number
  EditBox 110, 30, 20, 15, footer_month
  EditBox 110, 50, 20, 15, footer_year
EndDialog


DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		' Field validation to make sure that fields cannot be blank, non-numeric, not too long, etc. Unable to figure out how to ensure month and year fall into range without erroring out, likely due to using > or < with non-numeric situations.

		If footer_month < 1 OR footer_month > 12 THEN err_msg = err_msg & "* The footer month must be a 2-digit number between 01 and 12"
		' If footer_month_number + footer_month > = true AND footer_month > 12 THEN err_msg = err_msg & "cannot be more than 12"
		If IsNumeric(MAXIS_case_number) = false OR Len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* The case number must be numeric and 8 digits or less." 
		If IsNumeric(footer_month) = false OR Len(footer_month) <> 2 THEN err_msg = err_msg & vbNewLine & "* The footer month must be a 2 digit number. Be sure to include a 0 before single digit years." 
		If IsNumeric(footer_year) = false OR Len(footer_year) <> 2 THEN err_msg = err_msg & vbNewLine & "* The footer year must be a 2 digit number. Be sure to include a 0 before single digit years." 
		If err_msg <> "" THEN MsgBox "FORM ERROR(S)!" & vbNewLine & err_msg

	Loop UNTIL err_msg = ""

    'Add to all dialogs where you need to work within BLUEZONE
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'End dialog section-----------------------------------------------------------------------------------------------

' Navigate to STAT/JOBS Panel and pull information from JOBS panel
PF3
PF3
PF3
EMWriteScreen "STAT", 16, 43
EMWriteScreen MAXIS_case_number, 18, 43
footer_month_year = footer_month & footer_year
EMWriteScreen footer_month_year, 20, 43
transmit
EMWriteScreen "JOBS", 20, 71
transmit

' Read data from JOBS panel
EMReadScreen client_name, 20, 4, 36
trim(client_name)
EMReadScreen employer_name, 32, 7, 38
employer_name = REPLACE(employer_name, "_", " ")
employer_name = Trim(employer_name)
EMReadScreen inc_type, 1, 5, 34
EMReadScreen inc_start, 8, 9, 35
EMReadScreen updated_date, 8, 21, 55
EMReadScreen retrospective_wage_total, 8, 17, 38
retrospective_wage_total = trim(retrospective_wage_total)
EMReadScreen prospective_wage_total, 8, 17, 67
prospective_wage_total = trim(prospective_wage_total)
EMReadScreen user_login, 7, 21, 71

Dialog1 = "" 'blanking out dialog name
'Add dialog here: Add the dialog just before calling the dialog below unless you need it in the dialog due to using COMBO Boxes or other looping reasons. Blank out the dialog name with Dialog1 = "" before adding dialog.
'Shows dialog -----------------------------------------------------------------------------------------------------

' Add dialog to pull data from the JOBS screen and allow for addition of a case note

BeginDialog Dialog1, 0, 0, 251, 275, "JOBS Panel Info"
  ButtonGroup ButtonPressed
    OkButton 140, 250, 50, 15
    CancelButton 195, 250, 50, 15
  GroupBox 5, 5, 235, 75, "Staff Info"
  Text 10, 40, 50, 15, "Case Number: "
  Text 10, 20, 45, 15, "User Login:"
  Text 10, 60, 45, 15, "Last Updated: "
  Text 60, 20, 65, 15, user_login
  Text 60, 40, 65, 15, MAXIS_case_number
  Text 60, 60, 60, 15, updated_date
  GroupBox 5, 85, 235, 85, "Client Info"
  Text 10, 100, 50, 15, "Income Type: "
  Text 10, 125, 90, 15, "Retrospective Wage Total:"
  Text 10, 150, 85, 15, "Prospective Wage Total:"
  Text 60, 100, 50, 15, inc_type
  Text 100, 125, 85, 15, retrospective_wage_total
  Text 100, 150, 80, 15, prospective_wage_total
  Text 5, 175, 235, 15, "Based on the information above, fill out the case note below:"
  EditBox 5, 190, 235, 40, case_note
  Text 100, 100, 35, 10, "Employer: "
  Text 140, 100, 90, 10, employer_name
EndDialog


DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		' Add validation that case note is not blank

		If case_note = "" THEN err_msg = err_msg & vbNewLine & "* The case note cannot be blank. Add a case note to the field." 
		If err_msg <> "" THEN MsgBox "FORM ERROR(S)!" & vbNewLine & err_msg

	Loop UNTIL err_msg = ""

    'Add to all dialogs where you need to work within BLUEZONE
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'End dialog 


' Navigate to case note to add JOBS panel details to case note
PF4
PF9

' Create array of the variable titles to combine with variables
dim variable_title_array

variable_title_array = array("Client Name: ", "Case Number: ", "Last Updated: ", "Employer Name: ", "Retrospective Wage Total: ", "Prospective Wage Total: ", "Case Note: ", "User Login: ")


' Create array of the variables to enter into case note
dim variable_array
variable_array = array(client_name, MAXIS_case_number, updated_date, employer_name, retrospective_wage_total, prospective_wage_total, case_note, user_login)


' dim row
' row = 4
' EMWriteScreen (variable_title_array(0) & variable_array(0)), row, 3
' row = row + 1
' EMWriteScreen (variable_title_array(1) & variable_array(1)), row, 3
' row = row + 1
' EMWriteScreen (variable_title_array(2) & variable_array(2)), row, 3
' row = row + 1
' EMWriteScreen (variable_title_array(3) & variable_array(3)), row, 3
' row = row + 1
' EMWriteScreen (variable_title_array(4) & variable_array(4)), row, 3
' row = row + 1
' EMWriteScreen (variable_title_array(5) & variable_array(5)), row, 3
' row = row + 1
' EMWriteScreen (variable_title_array(6) & variable_array(6)), row, 3
' row = row + 1
' EMWriteScreen (variable_title_array(7) & variable_array(7)), row, 3
' row = row + 1

' MsgBox UBound(variable_array)

dim row
row = 4
dim array_length
array_length = UBound(variable_array) + 1
dim variable_title_array_index
variable_title_array_index = 0
dim variable_array_index
variable_array_index = 0
dim counter
counter = 1

Do until counter > array_length
	EMWriteScreen (variable_title_array(variable_title_array_index) & variable_array(variable_array_index)), row, 3
	row = row + 1
	variable_array_index = variable_array_index + 1
	variable_title_array_index = variable_title_array_index + 1
	counter = counter + 1
Loop






'code snippet example---------------------------------------------------------------------------------------------
'This is here to show you how we might use the advanced automation library to do something in MAXIS.
'Feel free to build from this or just take the parts that are helpful.

'We have now made sure we are at SELF in MAXIS

'now we are going to STAT/SUMM for a specific case
' EMWriteScreen "STAT", 16, 43				'writing the MAXIS function to enter in the correct place in MAXIS
' EMWriteScreen MAXIS_case_number, 18, 43		'entering  case number in the 'case number' line
' 'TODO - should I be concerned if there is already information on this line?
' EMWriteScreen "SUMM", 21, 70				'writing the MAXIS command to enter in the correct place in MAXIS

' transmit									'function to move in MAXIS

'TODO - how do I make sure that I actually got to STAT/SUMM

















'leave the case note open and in edit mode unless you have a business reason not to (BULK scripts, multiple case notes, etc.)

'End the script. Put any success messages in between the quotes
script_end_procedure("")
