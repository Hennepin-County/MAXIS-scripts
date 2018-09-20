'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - CLAIM REFERRAL TRACKING.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 200            'manual run time in seconds
STATS_denomination = "C"        'C is for each case
 'END OF stats block=========================================================================================================

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
call changelog_update("09/20/2018", "Updated dialog to match MAXIS panel.", "MiKayla Handley")
call changelog_update("06/26/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'-----------------------------------------------------------------------------DIALOG
BeginDialog Claim_Referral_Tracking, 0, 0, 266, 105, "Claim Referral Tracking"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 165, 5, 45, 15, Action_Date
  DropListBox 55, 25, 50, 15, "Select One:"+chr(9)+"SNAP"+chr(9)+"MFIP"+chr(9)+"SNAP/MFIP", Program_droplist
  DropListBox 165, 25, 95, 15, "Select One:"+chr(9)+"Initial Claim Referral"+chr(9)+"Claim Determination", Action_Taken
  CheckBox 55, 45, 110, 10, "Sent Request for Additional Info", Verif_Checkbox
  CheckBox 185, 45, 75, 10, "Overpayment Exists", Overpayment_Checkbox
  EditBox 50, 60, 210, 15, Other_Notes
  EditBox 50, 80, 90, 15, Worker_Signature
  ButtonGroup ButtonPressed
    OkButton 155, 80, 50, 15
    CancelButton 210, 80, 50, 15
  Text 5, 45, 45, 10, "Action Taken:"
  Text 5, 65, 45, 10, "Other Notes:"
  Text 5, 85, 40, 10, "Worker Sig:"
  Text 120, 10, 40, 10, "Action Date:"
  Text 15, 30, 40, 10, "Program(s):"
  Text 5, 10, 50, 10, "Case Number: "
  Text 120, 30, 40, 10, "Description:"
EndDialog

'----------------------------------------------------------------------------------------------------Thescript
EMCONNECT ""

Call MAXIS_case_number_finder(MAXIS_case_number)
Action_Date = Date & ""

Do
	Do
		err_msg = ""
		dialog Claim_Referral_Tracking
		IF buttonpressed = 0 then stopscript
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		IF isdate(action_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid action date."
		IF program_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select a program."
        IF Overpayment_Checkbox <> CHECKED and Verif_Checkbox <> CHECKED  THEN err_msg = err_msg & vbnewline & "Please select an action taken"
		IF Action_Taken = "Select One:" then err_msg = err_msg & vbnewline & "* Select an action."
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Going to the MISC panel
Call navigate_to_MAXIS_screen ("STAT", "MISC")
Row = 6
EmReadScreen panel_number, 1, 02, 78
If panel_number = "0" then
	EMWriteScreen "NN", 20,79
	TRANSMIT
ELSE
	Do
    	'Checking to see if the MISC panel is empty, if not it will find a new line'
    	EmReadScreen MISC_description, 25, row, 30
    	MISC_description = replace(MISC_description, "_", "")
    	If trim(MISC_description) = "" then
			PF9
    		EXIT DO
    	Else
            row = row + 1
    	End if
	Loop Until row = 17
    If row = 17 then script_end_procedure("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
End if

'writing in the action taken and date to the MISC panel
EMWriteScreen Action_Taken, Row, 30
EMWriteScreen Action_Date, Row, 66
PF3

'set TIKL------------------------------------------------------------------------------------------------------
If Verif_Checkbox = checked then
	If Action_Taken = "Claim Determination" THEN
		Msgbox "You identified your case is ready to process for overpayment follow procedures for claim entry.  A TIKL will NOT be made."
	ELSE
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(Action_Date, 10, 5, 18)
		Call write_variable_in_TIKL("Potential overpayment exists on case. Please review case for receipt of additional requested information.")
		PF3
	END IF
END IF
'The case note-------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_case_note("***Claim Referral Tracking - " & Action_Taken & "***")
Call write_bullet_and_variable_in_case_note("Program(s)", program_droplist)
Call write_bullet_and_variable_in_case_note("Action Date", Action_Date)
Call write_bullet_and_variable_in_case_note("Other Notes", Other_Notes)
Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
IF (Verif_Checkbox = checked and Action_Taken <> "Claim Determination") then write_variable_in_case_note("* Additional verifications requested, TIKL set for 10 day return.")
IF Overpayment_Checkbox = checked then write_variable_in_case_note("* Overpayment exists, collection process to follow.")
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)

IF Overpayment_Checkbox = checked then
	script_end_procedure("You have indicated that an overpayment exists. Please follow the agency's procedure(s) for claim entry.")
Else
	script_end_procedure("")
End if
