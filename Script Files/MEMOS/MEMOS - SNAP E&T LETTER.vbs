'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - SNAP E&T LETTER"
start_time = timer

'Option Explicit

DIM beta_agency
DIM url, req, fso

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF



'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog SNAPE&T_dialog, 0, 0, 326, 210, "SNAP E&T Appointment Letter"
  ButtonGroup ButtonPressed
    OkButton 215, 195, 50, 15
    CancelButton 270, 195, 50, 15
	PushButton 290, 40, 35, 15, "refresh", refresh_button
  EditBox 85, 0, 55, 15, case_number
  EditBox 235, 0, 20, 15, member_number
  EditBox 85, 20, 55, 15, appointment_date
  EditBox 235, 20, 45, 15, appointment_time
  DropListBox 240, 40, 45, 15, county_office_list, interview_location
  EditBox 65, 60, 170, 15, SNAPET_name
  EditBox 65, 75, 170, 15, SNAPET_address_01
  EditBox 65, 90, 170, 15, SNAPET_address_02
  EditBox 65, 105, 65, 15, SNAPET_contact
  EditBox 210, 105, 60, 15, SNAPET_phone
  EditBox 65, 140, 65, 15, worker_signature
  Text 10, 5, 50, 10, "Case Number:"
  Text 150, 5, 70, 10, "HH Member Number:"
  Text 10, 20, 65, 15, "Appointment Date:"
  Text 160, 20, 75, 10, "Appointment Time:"
  Text 5, 40, 230, 15, "Location (select from dropdown and click refresh, or fill in manually)"
  Text 5, 60, 55, 10, "Provider Name:"
  Text 5, 75, 55, 10, "Address line 1:"
  Text 5, 90, 55, 10, "Address Line 2"
  Text 5, 110, 55, 10, "Contact Name:"
  Text 150, 110, 55, 10, "Contact Phone:"
  Text 5, 145, 60, 10, "Worker Signature:"
  Text 5, 160, 315, 25, "Please note: the dropdown above automatically fills in from your agency office/intake locations.  It may not match your SNAP E&T orientation locations.  Please double check the address before pressing OK. "
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone default screen
EMConnect ""

'Searches for a case number
row = 1
col = 1
EMSearch "Case Nbr: ", row, col
EMReadScreen case_number, 8, row, col + 10
case_number = trim(replace(case_number, "_", ""))
If isnumeric(case_number) = False then case_number = ""

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
											Dialog SNAPE&T_dialog
											IF ButtonPressed = 0 then stopscript
											IF buttonPressed = refresh_button then
												IF interview_location <> "" then 
													call assign_county_address_variables(county_address_line_01, county_address_line_02)
													SNAPET_address_01 = county_address_line_01
													SNAPET_address_02 = county_address_line_02
												END IF
											END IF
										LOOP UNTIL ButtonPressed = OK
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
				IF appointment_time = "" then MsgBox "Please specify an appointment time."
				LOOP UNTIL appointment_time <> ""
			IF worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
			LOOP UNTIL worker_signature <> ""
		IF SNAPET_contact = "" THEN MsgBox "You must specify the E&T contact name.  Please try again."
		LOOP UNTIL SNAPET_contact <> ""
	IF SNAPET_phone = "" THEN MsgBox "You must enter a contact phone number.  Please try again."
	LOOP UNTIL SNAPET_phone <> ""	
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be outside of MAXIS. You may be locked out of MAXIS, check your screen and try again."
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

  'Pulls the member name.
 call navigate_to_screen("STAT", "MEMB")
 EMWriteScreen member_number, 20, 76
 transmit
 EMReadScreen last_name, 24, 6, 30
 EMReadScreen first_name, 11, 6, 63
 last_name = trim(replace(last_name, "_", ""))
 first_name = trim(replace(first_name, "_", ""))
 
 'Navigates into SPEC/LETR
  call navigate_to_screen("SPEC", "LETR")
  
  'Checks to make sure we're past the SELF menu
  EMReadScreen still_self, 27, 2, 28 
  If still_self = "Select Function Menu (SELF)" then script_end_procedure("Unable to get past the SELF screen. Is your case in background?")
  
  'Opens up the SNAP E&T Orientation LETR. If it's unable the script will stop.
  EMWriteScreen "x", 8, 12
  transmit
  EMReadScreen LETR_check, 4, 2, 49
  If LETR_check = "LETR" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
 

 
 'Creates a Maxis friendly date for spacing of date fields on the FSET letter screen
 Function create_FSET_friendly_date(date_variable, variable_length, screen_row, screen_col) 
  var_month = datepart("m", dateadd("d", variable_length, date_variable))
  If len(var_month) = 1 then var_month = "0" & var_month
  EMWriteScreen var_month, screen_row, screen_col
  var_day = datepart("d", dateadd("d", variable_length, date_variable))
  If len(var_day) = 1 then var_day = "0" & var_day
  EMWriteScreen var_day, screen_row, screen_col + 5
  var_year = datepart("yyyy", dateadd("d", variable_length, date_variable))
  EMWriteScreen right(var_year, 2), screen_row, screen_col + 10
End function
'Cleaning up the phone number
With (New RegExp)
	.Global = True
	.Pattern = "\D"
	SNAPET_phone = .Replace(SNAPET_phone, "") 'Removes all non-digits
End With
SNAPET_phoneright = right(SNAPET_phone, 4)
SNAPET_phone = left(SNAPET_phone, 6)

'Cleaning up the appointment time
With (New RegExp)
	.Global = True
	.Pattern = "\D"
	appointment_time_fix = .Replace(appointment_time, "") 'Removes all non-digits
End With
IF LEN(appointment_time_fix) = 3 then appointment_time_fix = "0" & appointment_time_fix
 

  'Writes the info into the LETR. 
  EMWriteScreen first_name & " " & last_name, 4, 28
  call create_FSET_friendly_date(appointment_date, 0, 6, 28) 
  EMWriteScreen left(appointment_time_fix, 2), 7, 28
  EMWriteScreen right(appointment_time_fix, 2), 7, 33
  IF cint(left(appointment_time_fix, 2)) > 6 THEN 'Automatically determines AM / PM based on hour of appointment (no appointments expected before 6 Am / after 6 pm)
	EMWriteScreen "AM", 7, 38
  ELSE 
	EMWriteScreen "PM", 7, 38
  END IF
  EMWriteScreen SNAPET_name, 9, 28
  EMWriteScreen SNAPET_address_01, 10, 28
  EMWriteScreen SNAPET_address_02, 11, 28
  EMWriteScreen left(SNAPET_phone, 3), 13, 28
  EMWriteScreen right(SNAPET_phone, 3), 13, 34
  EMWriteScreen SNAPET_phoneright, 13, 40
  EMWriteScreen SNAPET_contact, 16, 28
  PF4
  'check to make sure memo sent
      
 'Navigates to a blank case note
  call navigate_to_screen("case", "note")
  PF9
    
 'Writes the case note
 CALL write_new_line_in_case_note("***SNAP E&T Appointment Letter Sent***")
 CALL write_bullet_and_variable_in_case_note("Appointment date:", appointment_date)
 CALL write_bullet_and_variable_in_case_note("Appointment time:", appointment_time)
 CALL write_bullet_and_variable_in_case_note("Appointment location:", SNAPET_name)
 
 CALL write_new_line_in_case_note("---")
 CALL write_new_line_in_case_note(worker_signature)

script_end_procedure("")

