'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<adding for testing purposes
Worker_county_code = "x127"	 

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - SNAP E&T LETTER.vbs"
start_time = timer

'Option Explicit

DIM beta_agency
DIM FuncLib_URL, req, fso

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

FUNCTION create_MAXIS_friendly_date_three_spaces_between(date_variable, variable_length, screen_row, screen_col) 
	var_month = datepart("m", dateadd("d", variable_length, date_variable))		'determines the date based on the variable length: month 
	If len(var_month) = 1 then var_month = "0" & var_month						'adds a '0' in front of a single digit month
	EMWriteScreen var_month, screen_row, screen_col								'writes in var_month at coordinates set in FUNCTION line
	var_day = datepart("d", dateadd("d", variable_length, date_variable)) 		'determines the date based on the variable length: day
	If len(var_day) = 1 then var_day = "0" & var_day 							'adds a '0' in front of a single digit day
	EMWriteScreen var_day, screen_row, screen_col + 5 							'writes in var_day at coordinates set in FUNCTION line, and starts 5 columns into date field in MAXIS
	var_year = datepart("yyyy", dateadd("d", variable_length, date_variable)) 	'determines the date based on the variable length: year
	EMWriteScreen right(var_year, 2), screen_row, screen_col + 10 				'writes in var_year at coordinates set in FUNCTION line , and starts 5 columns into date field in MAXIS
END FUNCTION	

FUNCTION create_MAXIS_friendly_phone_number(phone_number_variable, screen_row, screen_col)
	WITH (new RegExp)                                                            'Uses RegExp to bring in special string functions to remove the unneeded strings
                .Global = True                                                   'I don't know what this means but David made it work so we're going with it
                .Pattern = "\D"                                                	 'Again, no clue. Just do it.
                phone_number_variable = .Replace(phone_number_variable, "")    	 'This replaces the non-digits of the phone number with nothing. That leaves us with a bunch of numbers
	END WITH
	EMWriteScreen left(phone_number_variable, 3), screen_row, screen_col 
	EMWriteScreen mid(phone_number_variable, 4, 3), screen_row, screen_col + 6
	EMWriteScreen right(phone_number_variable, 4), screen_row, screen_col + 12
END FUNCTION
	

'Array listed above Dialog as below the dialog, the droplist appeared blank
'Creates an array of county FSET offices, which can be dynamically called in scripts which need it (SNAP ET LETTER for instance)

county_FSET_offices = array("Select one", "Century Plaza", "Sabathani Community Center")
'IF worker_county_code = "x127" THEN county_FSET_offices = array("Select one", "Century Plaza", "Sabathani Community Center")

call convert_array_to_droplist_items (county_FSET_offices, FSET_list)

If worker_county_code = "x127" THEN 
	SNAPET_contact = "the EZ Info Line"
	SNAPET_phone = "612-596-1300"
END IF

'DIALOGS----------------------------------------------------------------------------------------------------
' FSET_list is a variable not a standard drop down list.  When you copy into dialog editor, it will not work
BeginDialog SNAPET_dialog, 0, 0, 321, 195, "SNAP E&T Appointment Letter"
  EditBox 70, 5, 55, 15, case_number
  EditBox 215, 5, 20, 15, member_number
  EditBox 70, 25, 55, 15, appointment_date
  EditBox 215, 25, 20, 15, appointment_time_prefix_editbox
  EditBox 235, 25, 20, 15, appointment_time_post_editbox
  DropListBox 260, 25, 55, 15, "Select one.."+chr(9)+"AM"+chr(9)+"PM", AM_PM
  DropListBox 175, 45, 140, 15, "county_office_list", interview_location
  EditBox 65, 65, 190, 15, SNAPET_name
  EditBox 65, 85, 190, 15, SNAPET_address_01
  EditBox 65, 105, 95, 15, SNAPET_city
  EditBox 165, 105, 40, 15, SNAPET_ST
  EditBox 210, 105, 45, 15, SNAPET_zip
  EditBox 65, 125, 65, 15, SNAPET_contact
  EditBox 185, 125, 70, 15, SNAPET_phone
  EditBox 140, 175, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 210, 175, 50, 15
    CancelButton 265, 175, 50, 15
  Text 5, 30, 60, 10, "Appointment Date:"
  Text 145, 30, 60, 15, "Appointment Time:"
  Text 5, 50, 170, 10, "Location (select from dropdown, or fill in manually)"
  Text 5, 70, 55, 10, "Provider Name:"
  Text 5, 90, 55, 10, "Address line 1:"
  Text 10, 130, 55, 10, "Contact Name:"
  Text 135, 130, 50, 10, "Contact Phone:"
  Text 80, 180, 60, 10, "Worker Signature:"
  Text 5, 145, 315, 25, "Please note: the dropdown above automatically fills in from your agency office/intake locations.  It may not match your SNAP E&T orientation locations.  Please double check the address before pressing OK. "
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 110, 55, 10, "City/State/Zip:"
  Text 140, 10, 70, 10, "HH Member Number:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone default screen
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(case_number)

'Shows dialog, checks for password prompt
DO
	DO
		DO
			DO
				DO
					DO
						DO
							DO
								DO
									DO
										DO
											Dialog SNAPET_dialog
											cancel_confirmation 'asks if they really want to cancel script	
											IF case_number = "" then MsgBox "You did not enter a case number. Please try again."
										LOOP UNTIL case_number <> ""
										If isdate(appointment_date) = FALSE then MsgBox "You did not enter a valid appointment date. Please try again."
									LOOP UNTIL isdate(appointment_date) = True
									IF member_number = "" then MsgBox "You did not specify a household member number.  Please try again."
								LOOP UNTIL isnumeric(member_number) = true
								IF SNAPET_name = "" then MsgBox "Please specify the agency name."
							LOOP UNTIL SNAPET_name <> ""
							IF SNAPET_address_01 = "" then MsgBox "Please enter the address for the SNAP ET agency."
						LOOP UNTIL SNAPET_address_01 <> ""
						IF appointment_time_prefix_editbox = "" then MsgBox "Please specify an appointment time."
					LOOP UNTIL appointment_time_prefix_editbox <> ""
					IF appointment_time_post_editbox = "" then MsgBox "Please specify an appointment time."
				LOOP UNTIL appointment_time_post_editbox <> ""	
				If AM_PM = "Select One..." THEN MsgBox "Please choose either a.m. or p.m."
			LOOP UNTIL AM_PM <> "Select One..."					
			IF worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
		LOOP UNTIL worker_signature <> ""
		IF SNAPET_contact = "" THEN MsgBox "You must specify the E&T contact name.  Please try again."
	LOOP UNTIL SNAPET_contact <> ""
	IF SNAPET_phone = "" THEN MsgBox "You must enter a contact phone number.  Please try again."
LOOP UNTIL SNAPET_phone <> ""	

transmit
Call maxis_check_function

'Logic for Hennepin County addresses only (currently)
county_FSET_offices = array("Select one", "Century Plaza", "Sabathani Community Center")
IF interview_location = "Century Plaza" THEN 
	SNAPET_name = "Century Plaza"
	SNAPET_address_01 = "330 South 12th Street #3650"
	SNAPET_address_02 = "Minneapolis, MN. 55404"
ElseIf interview_location = "Sabathani Community Center" THEN 
	SNAPET_name = "Sabathani Community Center"
	SNAPET_address_01 = "310 East 38th Street #120"
	SNAPET_address_02 = "Minneapolis, MN. 55409"
END IF

'Pulls the member name.
call navigate_to_MAXIS_screen("STAT", "MEMB")
EMWriteScreen member_number, 20, 76
transmit
EMReadScreen last_name, 24, 6, 30
EMReadScreen first_name, 11, 6, 63
last_name = trim(replace(last_name, "_", ""))
first_name = trim(replace(first_name, "_", ""))

'Navigates into SPEC/LETR
call navigate_to_MAXIS_screen("SPEC", "LETR") 

'Opens up the SNAP E&T Orientation LETR. If it's unable the script will stop.
EMWriteScreen "x", 8, 12
transmit
EMReadScreen LETR_check, 4, 2, 49
If LETR_check = "LETR" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")


'Writes the info into the LETR. 
EMWriteScreen first_name & " " & last_name, 4, 28
call create_MAXIS_friendly_date_three_spaces_between(appointment_date, 0, 6, 28) 
EMWriteScreen appointment_time_prefix_editbox, 7, 28
EMWriteScreen appointment_time_post_editbox, 7, 33
EMWriteScreen AM_PM, 7, 38
EMWriteScreen SNAPET_name, 9, 28
EMWriteScreen SNAPET_address_01, 10, 28
EMWriteScreen SNAPET_address_02, 11, 28
call create_MAXIS_friendly_phone_number(SNAPET_phone, 13, 28) 'takes out non-digits if listed in variable, and formats phone number for the field
EMWriteScreen SNAPET_contact, 16, 28
PF4		'saves and sends memo

'Navigates to a blank case note
call start_a_blank_CASE_NOTE

'Writes the case note
CALL write_new_line_in_case_note("***SNAP E&T Appointment Letter Sent***")
CALL write_bullet_and_variable_in_case_note("Appointment date", appointment_date)
CALL write_bullet_and_variable_in_case_note("Appointment time", appointment_time_prefix_editbox & ":" & appointment_time_post_editbox & " " & AM_PM)
CALL write_bullet_and_variable_in_case_note("Appointment location", SNAPET_name)
CALL write_new_line_in_case_note("---")
CALL write_new_line_in_case_note(worker_signature)

script_end_procedure("")
