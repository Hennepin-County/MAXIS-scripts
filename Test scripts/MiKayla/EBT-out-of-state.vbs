'GATHERING STATS===========================================================================================
name_of_script = "NOTES - DEU-EBT OUT OF STATE.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 0
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
call changelog_update("11/30/2018", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================----------
'connecting to BlueZone and grabbing the case number
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
MEMB_number = "01"
BeginDialog EBT_dialog, 0, 0, 161, 85, "EBT OUT OF STATE "
  EditBox 55, 5, 50, 15, maxis_case_number
  EditBox 55, 25, 50, 15, out_of_state
  DropListBox 55, 45, 100, 15, "Select One:"+chr(9)+"Initial Review"+chr(9)+"Response Received"+chr(9)+"No Response Received"+chr(9)+"Other", action_taken
  ButtonGroup ButtonPressed
    OkButton 60, 65, 45, 15
    CancelButton 110, 65, 45, 15
  Text 25, 30, 30, 10, "State(s):"
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 50, 45, 10, "Action Taken:"
EndDialog

BeginDialog intial_review_dialog, 0, 0, 196, 105, "EBT Out of State Initial Review"
  EditBox 55, 5, 40, 15, date_received
  DropListBox 160, 5, 30, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9"+chr(9)+"10"+chr(9)+"11"+chr(9)+"12"+chr(9)+"YEAR plus", months_used
  CheckBox 10, 35, 75, 10, "Request for Contact", request_contact_checkbox
  CheckBox 10, 45, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
  CheckBox 105, 35, 70, 10, "Shelter Verification", EVF_checkbox
  CheckBox 105, 45, 80, 10, "Other (please specify)", other_checkbox
  EditBox 50, 65, 140, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 105, 85, 40, 15
    CancelButton 150, 85, 40, 15
  Text 5, 10, 50, 10, "Date reported:"
  Text 105, 10, 55, 10, "# Months Used: "
  GroupBox 5, 25, 185, 35, "Verification Requested: "
  Text 5, 70, 45, 10, "Other Notes: "
EndDialog




DO
	DO
		err_msg = ""
		Dialog EBT_dialog
		IF ButtonPressed = 0 THEN StopScript
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If out_of_state = ""  then err_msg = err_msg & vbNewLine & "* Enter the state(s) that the client has used benefits in."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false


BeginDialog EBT_dialog, 0, 0, 256, 105, "EBT OUT OF STATE "
  EditBox 60, 5, 50, 15, maxis_case_number
  EditBox 200, 5, 50, 15, bene_date
  EditBox 60, 25, 50, 15, state
  EditBox 60, 45, 50, 15, date_closed
  DropListBox 180, 25, 70, 15, "Select One:"+chr(9)+"Initial Review"+chr(9)+"Client Response to Request"+chr(9)+"No Response Received"+chr(9)+"Other", action_taken
  EditBox 60, 65, 195, 15, reason_closed
  ButtonGroup ButtonPressed
    OkButton 150, 85, 50, 15
    CancelButton 205, 85, 50, 15
  Text 115, 10, 80, 10, "Date accessing benefits:"
  Text 5, 70, 55, 10, "Closure Reason:"
  Text 30, 30, 30, 10, "State(s):"
  Text 10, 10, 50, 10, "Case Number:"
  Text 130, 30, 45, 10, "Action Taken:"
  Text 15, 50, 45, 10, "Date Closed:"
EndDialog


Do
	Do
        err_msg = ""
		Dialog
		cancel_confirmation
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF Isdate(bene_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the benefit start date."
		IF Isdate(date_closed) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the closed date."
		IF action_taken = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please enter action completed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
 	Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("----- EBT OUT OF STATE REVIEWED -----")
	Call write_variable_in_CASE_NOTE("----- EBT OUT OF STATE SHELTER FORM SENT -----")
    Call write_bullet_and_variable_in_CASE_NOTE("Client has been accessing benefits out of state since:", bene_date)
	Call write_bullet_and_variable_in_CASE_NOTE("State(s):", state)
	Call write_variable_in_CASE_NOTE("Request sent to client for explanation of benefits received in the other state and shelter request ")
    Call write_variable_in_CASE_NOTE("Client will need to verify residence when reapplying")
    Call write_variable_in_CASE_NOTE("Agency will need to verify benefits received in the other state prior to reopening case")
	Call write_bullet_and_variable_in_CASE_NOTE("Date case was closed", date_closed)
	Call write_bullet_and_variable_in_CASE_NOTE("Explanation of action to close the case", reason_closed)
	Call write_variable_in_CASE_NOTE("DEU will review for possible overpayment regarding out of state usage at a later date.")
	Call write_variable_in_CASE_NOTE("Clients have 10 days to return requested verifications")
	Call write_bullet_and_variable_in_CASE_NOTE("Date due", date_due)
	Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
	Call write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")

	IF  THEN
		Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
		two_weeks_from_now = DateAdd("d", 14, date)
		call create_MAXIS_friendly_date(two_weeks_from_now, 10, 5, 18)
		call write_variable_in_TIKL ("Review client's application for Unemployment and request an update if needed.")
		PF3
	END IF


script_end_procedure("EBT out of state case note complete.")
