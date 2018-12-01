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
Call MAXIS_case_number_finder(maxis_case_number)

'Initial dialog and do...loop
BeginDialog EBT_dialog, 0, 0, 256, 105, "EBT OUT OF STATE "
  EditBox 60, 5, 50, 15, maxis_case_number
  EditBox 200, 5, 50, 15, bene_date
  EditBox 60, 25, 50, 15, state
  EditBox 60, 45, 50, 15, date_closed
  DropListBox 180, 25, 70, 15, "Select One:"+chr(9)+"Initial Review"+chr(9)+"Respond to Request"+chr(9)+"Other", action_taken
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

	Call write_bullet_and_variable_in_CASE_NOTE("Date case was closed:", date_closed)
	Call write_bullet_and_variable_in_CASE_NOTE("Explanation of action to close the case:", reason_closed)
	Call write_variable_in_CASE_NOTE("Possible overpayments will be reviewed") 'do we want to add the claim referral?'
	Call write_variable_in_CASE_NOTE("Clients have 10 days to return requested verifications")
	Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
	Call write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")


	REQUEST TO CLIENT TO VERIFY WHY ACCESSING BENEFITS OUT OF STATE
	AND SHELTER VERIFICATION FORM.
	EBT OUT OF STATE USAGE WILL BE REVIEWED FOR POSSIBLE OVERPAYMENT OR NO OVERPAYMENT AT A LATER DATE.
	DUE:
	Debt Establishment Unit 612-348-4290 X111
	d.	 Clients have 10 days to respond.


script_end_procedure("EBT out of state case note complete.")
