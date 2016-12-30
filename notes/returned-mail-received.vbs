'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - RETURNED MAIL RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 360          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog case_number_dlg, 0, 0, 150, 80, "CASE NUMBER DIALOG"
  EditBox 75, 10, 70, 15, MAXIS_case_number
  EditBox 75, 30, 40, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 10, 55, 50, 15
    CancelButton 95, 55, 50, 15
  Text 20, 15, 50, 10, "Case Number:"
  Text 10, 35, 65, 10, "Worker Signature:"
EndDialog

BeginDialog RETURNED_MAIL, 0, 0, 185, 335, "RETURNED MAIL DIALOG"
  EditBox 110, 5, 65, 15, date_received
  EditBox 55, 25, 120, 15, from_ADDR
  EditBox 55, 45, 120, 15, from_CITY
  EditBox 55, 65, 20, 15, from_STATE
  EditBox 120, 65, 55, 15, from_ZIP
  DropListBox 115, 85, 35, 15, "No"+chr(9)+"Yes", forwarding_ADDR
  EditBox 55, 105, 120, 15, new_ADDR
  EditBox 55, 125, 120, 15, new_CITY
  EditBox 55, 145, 25, 15, new_STATE
  EditBox 135, 145, 40, 15, new_ZIP
  DropListBox 110, 165, 35, 15, "No"+chr(9)+"Yes", updated_ADDR
  EditBox 110, 180, 65, 15, new_COUNTY
  CheckBox 50, 200, 70, 10, "Sent DHS-2919A", verifA_sent_checkbox
  CheckBox 50, 210, 65, 10, "Sent DHS-2952", SHEL_form_sent_checkbox
  CheckBox 50, 220, 65, 10, "Sent DHS-2402", CRF_sent_checkbox
  DropListBox 120, 230, 30, 15, "No"+chr(9)+"Yes", returned_mail_resent_list
  DropListBox 105, 245, 45, 15, "Select"+chr(9)+"Yes"+chr(9)+"No", MNsure_active
  EditBox 100, 260, 75, 15, MNsure_number
  DropListBox 100, 278, 40, 10, "N/A"+chr(9)+"Yes"+chr(9)+"No", MNsure_ADDR
  EditBox 100, 295, 75, 15, misc_notes
  ButtonGroup ButtonPressed
    OkButton 5, 315, 50, 15
    CancelButton 130, 315, 50, 15
  Text 45, 235, 70, 10, "Returned mail resent:"
  Text 85, 150, 50, 10, "New Zip Code:"
  Text 10, 300, 90, 10, "Misc notes/Actions Taken:"
  Text 10, 30, 40, 10, "From ADDR:"
  Text 10, 110, 40, 10, "New ADDR:"
  Text 30, 70, 20, 10, "State:"
  Text 10, 150, 40, 10, "New State:"
  Text 20, 90, 90, 10, "Forwarding ADDR provided:"
  Text 35, 170, 65, 10, "Updated Address:"
  Text 85, 70, 35, 10, "Zip Code:"
  Text 10, 10, 95, 10, "Returned Mail Received on:"
  Text 15, 50, 35, 10, "From City:"
  Text 10, 130, 35, 10, "New City:"
  Text 35, 185, 70, 10, "County of new ADDR:"
  Text 10, 250, 90, 10, "Client have a MNsure case:"
  Text 20, 265, 75, 10, "MNsure case number:"
  Text 20, 280, 75, 10, "MNsure ADDR correct:"
EndDialog


'connects to BlueZone and brings it forward
EMConnect ""
EMFocus

'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
Call check_for_MAXIS(false)

'Finds the case number
Call MAXIS_case_number_finder(MAXIS_case_number)

'Finds the benefit month
EMReadScreen on_SELF, 4, 2, 50
IF on_SELF = "SELF" THEN
	CALL find_variable("Benefit Period (MM YY): ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" THEN CALL find_variable("Benefit Period (MM YY): " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
ELSE
	CALL find_variable("Month: ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" THEN CALL find_variable("Month: " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
END IF

' >>>>> GATHERING & CONFIRMING THE MAXIS CASE NUMBER <<<<<

DO
	err_msg = ""
	DIALOG case_number_dlg
		IF ButtonPressed = 0 THEN stopscript
		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
		'checks that the case note was signed
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "You must sign your case note!"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

CALL check_for_MAXIS(False)

'starts the EVF received case note dialog
DO
	err_msg = ""
	'starts the Returned Mail dialog
	Dialog RETURNED_MAIL
	'asks if you want to cancel and if "yes" is selected sends StopScript
	cancel_confirmation
	'checks that there is a date in the date received box
	IF IsDate (date_received) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a date in mm/dd/yy for date received."
	'checks if the ADDR rec'd from is filled in
	IF from_ADDR = "" THEN err_msg = err_msg & vbCr & "You must enter the ADDR rec'd from."
	'checks if the from City has been entered
	IF from_CITY = "" THEN err_msg = err_msg & vbCr & "You must enter the City rec'd from."
	'checks if the from State has been entered
	IF from_STATE = "" THEN err_msg = err_msg & vbCr & "You must enter the State rec'd from."
	'checks if the from Zip has been entered
	IF from_ZIP = "" THEN err_msg = err_msg & vbCr & "You must enter the Zip rec'd from."
	'checks if a forwarding ADDR has been provided and new ADDR entered if provided
	IF forwarding_ADDR = "Yes" and new_ADDR = "" THEN err_msg = err_msg & vbCr & "You must input the new ADDR."
	'checks if a forwarding ADDR has been provided and new State entered if provided
	IF forwarding_ADDR = "Yes" and new_CITY = "" THEN err_msg = err_msg & vbCr & "You must input the new City."
	'checks if a forwarding ADDR has been provided and new State entered if provided
	IF forwarding_ADDR = "Yes" and new_STATE = "" THEN err_msg = err_msg & vbCr & "You must input the new State."
	'checks if STAT/ADDR was updated when a forwarding ADDR was provided
	IF forwarding_ADDR = "Yes" and updated_ADDR = "No" THEN err_msg = err_msg & vbCr & "You must update the address when a forwarding address is rec'd."
	'checks if a forwarding ADDR has been provided
	IF forwarding_ADDR = "Yes" and new_ZIP = "" THEN err_msg = err_msg & vbCr & "You must input the new Zip Code."
	'checks if a forwarding ADDR has been provided
	IF forwarding_ADDR = "Yes" and new_COUNTY = "" THEN err_msg = err_msg & vbCr & "You must input the County of the new ADDR."
	'checks if client is active on MNsure question has been answered
	IF MNsure_active = "Select" THEN err_msg = err_msg & vbCr & "You must select if the client has a MNsure case or not."
	'checks if MNsure case number has been entered on a MNsure active case
	IF MNsure_active = "Yes" and MNsure_number = "" THEN err_msg = err_msg & vbCr & "You must enter the MNsure case number."
	'checks if MNsure ADDR updated
	IF MNsure_active = "Yes" and MNsure_ADDR = "N/A" THEN err_msg = err_msg & vbCr & "You must select if the MNsure ADDR is correct."
	'checks if notes/actions taken were entered
	IF misc_notes = "" THEN err_msg = err_msg & vbCr & "You must enter action taken/misc notes."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'assigns a value to the ADDR_status variable based on the value of the complete variable
IF forwarding_ADDR = "Yes" THEN ADDR_status = "a forwarding ADDR."
IF forwarding_ADDR = "No" THEN ADDR_status = "no forwarding ADDR."

'assigns a value to the MNsure variable based on the value of MNsure_active
IF MNsure_active = "Yes" THEN MNsure = "MNsure case"
IF MNsure_active = "No" THEN MNsure = "Non-MNsure"

'converts the old and new ADDR to all CAPS
from_ADDR = UCase(from_ADDR)
from_CITY = UCase(from_CITY)
from_STATE = UCase(from_STATE)
new_ADDR = UCase(new_ADDR)
new_CITY = UCase(new_CITY)
new_STATE = UCase(new_STATE)
new_COUNTY = UCase(new_COUNTY)

'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
call check_for_MAXIS (false)

'starts a blank case note
call start_a_blank_case_note

'this enters the actual case note info
call write_variable_in_CASE_NOTE("***Returned Mail received " & date_received & " with " & ADDR_status & "*** " & MNsure)
call write_bullet_and_variable_in_CASE_NOTE("From ADDR", from_ADDR)
call write_variable_in_CASE_NOTE("             " & from_CITY & ", " & from_STATE & " " & from_ZIP)
IF forwarding_ADDR = "Yes" THEN call write_variable_in_CASE_NOTE("* Address updated to: " & new_ADDR)
IF forwarding_ADDR = "Yes" THEN call write_variable_in_CASE_NOTE("                      " & new_CITY & ", " & new_STATE & " " & new_Zip)
IF forwarding_ADDR = "Yes" THEN call write_variable_in_CASE_NOTE("* New ADDR is in " & new_COUNTY & " COUNTY.")
If verifA_sent_checkbox = 1 then call write_variable_in_CASE_NOTE("* Verification Request Form A sent. **Auto TIKL set**")
If SHEL_form_sent_checkbox = 1 then call write_variable_in_CASE_NOTE("* Shelter Verification Form sent.")
If CRF_sent_checkbox = 1 then call write_variable_in_CASE_NOTE("* Change Report Form sent.")
call write_bullet_and_variable_in_CASE_NOTE("Returned Mail resent", mail_resent)
If returned_mail_resent_list = "Yes" then call write_variable_in_CASE_NOTE("* Returned Mail resent to client.")
If returned_mail_resent_list = "No" then call write_variable_in_CASE_NOTE("* Returned Mail not resent to client.")
call write_variable_in_CASE_NOTE("* " & MNsure & " " & MNsure_number)
call write_bullet_and_variable_in_CASE_NOTE("Misc notes/Actions Taken", misc_notes)
call write_variable_in_CASE_NOTE ("---")
call write_variable_in_CASE_NOTE(worker_signature)

'Checks if this is a MNsure case and pops up a message box with instructions if the ADDR is incorrect.
IF MNsure_active = "Yes" and MNsure_ADDR = "No" THEN MsgBox "Please update the MNsure ADDR if you are able to. If unable, please forward the new ADDR information to the correct area (i.e. HPU Case Manitenance - Action Needed Log)"

'creates a message box reminding the worker to review their case note prior to Auto-TIKLing.
IF verifA_sent_checkbox = 1 THEN MsgBox "Please review your case note for accuracy. When you click OK or press enter the script will enter an Auto-TIKL for you."

'Checks if a DHS2919A mailed and sets a TIKL for the return of the info.
IF verifA_sent_checkbox = 1 THEN
	call navigate_to_MAXIS_screen("dail", "writ")

	'The following will generate a TIKL formatted date for 10 days from now.
	call create_MAXIS_friendly_date(date, 10, 5, 18)

	'Writing in the rest of the TIKL.
	call write_variable_in_TIKL("ADDR verification requested via 2919A after returned mail being rec'd should have returned by now. If not received, take appropriate action. (TIKL auto-generated from script)." )
	transmit
	PF3

	'Success message
	MsgBox "Success! TIKL has been sent for 10 days from now for the ADDR verification requested via 2919A."

End if

script_end_procedure("")
