'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - SMRT.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 100           'manual run time in seconds
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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
Call changelog_update("06/13/2020", "Since there are updates to the requirements for submitting a SMRT Referral in ISDS, we are reviewing the functionality of this script. ##~## If there are specific changes, fields, information, or functionality that would make your work with this script easier, pleae contact us. ##~## ##~## Email us at HSPH.EWS.BlueZoneScripts@hennepin.us or submit an 'Error Report' at the end of the script run.##~##", "Casey Love, Hennepin County")
call changelog_update("01/19/2017", "Initial version.", "Ilse Ferris, Hennepin County")
call changelog_update("11/29/2017", "Update script for denials to remove start date.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'intial dialog for user to select a SMRT action
BeginDialog Dialog1 , 0, 0, 186, 65, "SMRT initial dialog"
  EditBox 85, 5, 60, 15, maxis_case_number
  DropListBox 85, 25, 95, 15, "Select one..."+chr(9)+"Initial request"+chr(9)+"ISDS referral completed"+chr(9)+"Determination received", SMRT_actions
  ButtonGroup ButtonPressed
    OkButton 75, 45, 50, 15
    CancelButton 130, 45, 50, 15
  Text 5, 30, 75, 10, "Select a SMRT action:"
  Text 30, 10, 45, 10, "Case number:"
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		if IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If SMRT_actions = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Select a SMRT action."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

If SMRT_actions = "Initial request" then
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1 , 0, 0, 326, 180, "Initial SMRT referral dialog"
      EditBox 80, 10, 75, 15, SMRT_member
      EditBox 270, 10, 50, 15, referral_date
      DropListBox 80, 35, 60, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", referred_exp
      EditBox 200, 35, 120, 15, expedited_reason
      EditBox 80, 60, 240, 15, referral_reason
      EditBox 80, 85, 50, 15, SMRT_start_date
      If UCASE(worker_county_code = "X127") then CheckBox 140, 90, 180, 10, "Check here if the ECF workflow has been completed.", ECF_workflow_checkbox
      EditBox 80, 110, 240, 15, other_notes
      EditBox 80, 135, 240, 15, action_taken
      EditBox 80, 160, 130, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 215, 160, 50, 15
        CancelButton 270, 160, 50, 15
      Text 15, 115, 65, 10, "Other SMRT notes:"
      Text 5, 40, 70, 10, "Is referral expedited?"
      Text 25, 140, 50, 10, " Actions taken:"
      Text 165, 15, 100, 10, "Date SMRT referral completed:"
      Text 5, 15, 70, 10, "SMRT requested for: "
      Text 20, 90, 60, 10, "SMRT start date:"
      Text 15, 165, 60, 10, "Worker Signature:"
      Text 155, 40, 45, 10, "If yes, why?:"
      Text 10, 65, 65, 10, "Reason for referral:"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		If SMRT_member = "" THEN err_msg = err_msg & vbNewLine & "* Enter the member info the SMRT referral."
    		If isdate(referral_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid referral date."
			If referred_exp = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Is the referral expedited?"
			If (referred_exp = "Yes" and trim(expedited_reason) = "") THEN err_msg = err_msg & vbNewLine & "* Enter the expedited reason."
			If trim(referral_reason) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the reason for the referral."
			If isdate(SMRT_start_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid SMRT start date."
			If trim(action_taken) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the actions taken."
			If trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False

	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("---Initial SMRT referral requested---")
	call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
	Call write_bullet_and_variable_in_CASE_NOTE("SMRT referral completed on", referral_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Is referral expedited", referred_exp)
	If referred_exp = "Yes" then Call write_bullet_and_variable_in_CASE_NOTE("Expedited reason", expedited_reason)
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for referral", referral_reason)
	Call write_bullet_and_variable_in_CASE_NOTE("SMRT start date", SMRT_start_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Other SMRT notes", other_notes)
	Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	If ECF_workflow_checkbox = 1 then call write_variable_in_CASE_NOTE("* Workflow process has been completed in case file system.")
	Call write_variable_in_CASE_NOTE ("---")
	call write_variable_in_CASE_NOTE(worker_signature)

	end_msg = "SMRT Action for Initial Request noted on Case."
END If

If SMRT_actions = "ISDS referral completed" then
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 326, 130, "ISDS referral completed for SMRT"
      EditBox 80, 10, 75, 15, SMRT_member
      EditBox 225, 10, 50, 15, referral_date
      EditBox 80, 35, 75, 15, prog_requested
      EditBox 225, 35, 50, 15, SMRT_start_date
      EditBox 80, 60, 240, 15, other_notes
      EditBox 80, 85, 240, 15, action_taken
      EditBox 80, 110, 130, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 215, 110, 50, 15
        CancelButton 270, 110, 50, 15
      Text 10, 65, 65, 10, "Other SMRT notes:"
      Text 10, 40, 65, 10, "Program requested:"
      Text 25, 90, 50, 10, " Actions taken:"
      Text 165, 15, 55, 10, "Completion date:"
      Text 5, 15, 70, 10, "SMRT requested for: "
      Text 165, 40, 60, 10, "SMRT start date:"
      Text 15, 115, 60, 10, "Worker Signature:"
    EndDialog
    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		If SMRT_member = "" THEN err_msg = err_msg & vbNewLine & "* Enter the member info the SMRT referral."
    		If isdate(referral_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid referral date."
    		If trim(prog_requested) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the program requested by the client."
    		If isdate(SMRT_start_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid SMRT start date."
			If trim(action_taken) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the actions taken."
    		If trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("---ISDS referral completed for SMRT---")
    call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT referral completed on", referral_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Program requested", prog_requested)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason for referral", referral_reason)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT start date", SMRT_start_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Other SMRT notes", other_notes)
    Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
    Call write_variable_in_CASE_NOTE ("---")
    call write_variable_in_CASE_NOTE(worker_signature)

	end_msg = "SMRT Action for ISDS Referral Completed noted on Case."
END If

If SMRT_actions = "Determination received" then
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 326, 140, "SMRT determination received"
      EditBox 80, 10, 75, 15, SMRT_member
      DropListBox 240, 10, 55, 15, "Select one..."+chr(9)+"Approved"+chr(9)+"Denied", SMRT_determination
      EditBox 80, 35, 75, 15, appd_progs
      EditBox 240, 35, 55, 15, SMRT_start_date
      EditBox 80, 60, 240, 15, other_notes
      EditBox 80, 85, 240, 15, action_taken
      CheckBox 80, 105, 60, 10, "MMIS updated", MMIS_checkbox
      EditBox 80, 120, 130, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 215, 120, 50, 15
        CancelButton 270, 120, 50, 15
      Text 25, 90, 50, 10, " Actions taken:"
      Text 165, 15, 70, 10, "SMRT determination:"
      Text 5, 15, 70, 10, "SMRT requested for: "
      Text 180, 40, 55, 10, "SMRT start date:"
      Text 15, 125, 60, 10, "Worker Signature:"
      Text 10, 65, 65, 10, "Other SMRT notes:"
      Text 10, 40, 70, 10, "Approved programs:"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		If SMRT_member = "" THEN err_msg = err_msg & vbNewLine & "* Enter the member info the SMRT referral."
    		If SMRT_determination = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Select the determination status."
    		If trim(appd_progs) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the approved programs."
    		If isdate(SMRT_start_date) = False and SMRT_determination = "Approved" THEN err_msg = err_msg & vbNewLine & "* Enter a valid SMRT start date."
			If trim(action_taken) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the actions taken."
    		If trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("---SMRT determination received: " & SMRT_determination & "---")
    call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
    Call write_bullet_and_variable_in_CASE_NOTE("Approved programs",appd_progs)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT start date", SMRT_start_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Other SMRT notes", other_notes)
    Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	If MMIS_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MMIS updated")
    Call write_variable_in_CASE_NOTE ("---")
    call write_variable_in_CASE_NOTE(worker_signature)

	end_msg = "SMRT Action for Determination Received noted on Case."
END If

Call script_end_procedure_with_error_report(end_msg)
