'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CS GOOD CAUSE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 240          'manual run time in seconds
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
call changelog_update("05/14/2018", "Updated per GC Committee requests.", "MiKayla Handley, Hennepin County")
call changelog_update("03/27/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The DIALOGS----------------------------------------------------------------------------------------------------
EMConnect ""

'Inserts Maxis Case number
CALL MAXIS_case_number_finder(MAXIS_case_number)

BeginDialog change_exemption_dialog, 0, 0, 216, 100, "Good cause change/exemption "
  EditBox 105, 5, 50, 15, change_reported_date
  EditBox 105, 25, 105, 15, change_reported
  EditBox 105, 45, 105, 15, maxis_updates
  CheckBox 10, 70, 75, 10, " No longer claiming ", no_longer_claiming_checkbox
  ButtonGroup ButtonPressed
    OkButton 105, 80, 50, 15
    CancelButton 160, 80, 50, 15
  Text 10, 10, 80, 10, "Date of change reported:"
  Text 10, 30, 90, 10, "What change was reported:"
  Text 10, 50, 95, 10, "What was updated in MAXIS:"
EndDialog

BeginDialog Good_cause_initial_dialog, 0, 0, 231, 210, "Good Cause"
  EditBox 60, 5, 45, 15, MAXIS_case_number
  EditBox 190, 5, 35, 15, claim_date
  EditBox 60, 25, 20, 15, MAXIS_footer_month
  EditBox 85, 25, 20, 15, MAXIS_footer_year
  EditBox 190, 25, 20, 15, memb_number
  DropListBox 110, 45, 115, 15, "Select One:"+chr(9)+"Application Review - Complete"+chr(9)+"Application Review -  Incomplete"+chr(9)+"Change/exemption ending"+chr(9)+"Determination"+chr(9)+"Recertification", good_cause_droplist
  DropListBox 110, 65, 115, 15, "Select One:"+chr(9)+"Potential phys harm/Child"+chr(9)+"Potential Emotnl harm/Child"+chr(9)+"Potential phys harm/Caregiver"+chr(9)+"Potential Emotnl harm/Caregiver"+chr(9)+"Cncptn Incest/Forced Rape"+chr(9)+"Legal adoption Before Court"+chr(9)+"Parent Gets Preadoptn Svc", reason_droplist
  EditBox 65, 90, 160, 15, verifs_req
  EditBox 65, 110, 160, 15, mets_info
  EditBox 65, 130, 160, 15, other_notes
  CheckBox 10, 160, 155, 10, "Sent copy of good cause statement to client", GC_sent_checkbox
  CheckBox 10, 150, 145, 10, "DHS-2338 is in ECF and completed in full", DHS_2338_complete
  CheckBox 10, 170, 90, 10, "Sent DHS 3632 to client", sent_3632_incomplete
  CheckBox 10, 180, 90, 10, "Sent DHS 3633 to client", sent_3633_incomplete
  ButtonGroup ButtonPressed
    OkButton 120, 190, 50, 15
    CancelButton 175, 190, 50, 15
  Text 10, 10, 45, 10, "Case number:"
  Text 110, 10, 80, 10, "Claim date MM/DD/YY:"
  Text 5, 30, 50, 10, "Footer MM/YY:"
  Text 130, 30, 55, 10, "Child's MEMB #:"
  Text 80, 50, 25, 10, "Action:"
  Text 75, 70, 30, 10, "Reason:"
  Text 5, 95, 55, 10, "Verifs requested:"
  Text 5, 115, 60, 10, "Mets Information:"
  Text 20, 135, 40, 10, "Other notes:"
EndDialog



'The script----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'Initial dialog giving the user the option to select the type of good cause action
Do
	Do
		err_msg = ""
		dialog Good_cause_initial_dialog
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		IF good_cause_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select a good cause option."
		If isnumeric(MAXIS_footer_month) = false then err_msg = err_msg & vbnewline & "* You must enter the footer month to begin good cause."
		If isnumeric(MAXIS_footer_year) = false then err_msg = err_msg & vbnewline & "* You must enter the footer year to begin good cause."
		If isdate(claim_date) = False then err_msg = err_msg & vbnewline & "* You must enter a valid good cause claim date."
		If len(claim_date) <> 8 then err_msg = err_msg & vbnewline & "* You must enter the date in MM/DD/YY."
		If reason_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select the Good Cause reason."
		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

	Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
	Call MAXIS_footer_month_confirmation			'function that confirms that the current footer month/year is the same as what was selected by the user. If not, it will navigate to correct footer month/year

    'grabbing the case name for the email
    Call navigate_to_MAXIS_screen("STAT", "MEMB")
    EMReadScreen last_name, 24, 6, 30
    EMReadScreen first_name, 12, 6, 63
    'cleaning up the name variable
    last_name = replace(last_name, "_", "")
    first_name = replace(first_name, "_", "")
    client_name = first_name & " " & last_name
    Call fix_case_for_name(client_name)

	'----------------------------------------------------------------------------------------------------ABPS panel
	Call navigate_to_MAXIS_screen("STAT", "ABPS")
	'Making sure we have the correct ABPS
	EMReadScreen panel_number, 1, 2, 73
	If panel_number = "0" then script_end_procedure("An ABPS panel does not exist. Please create the panel before running the script again. ")

	'If there is more than one panel, this part will grab employer info off of them and present it to the worker to decide which one to use.
	If panel_number <> "0" then
		Do
			EMReadScreen current_panel_number, 1, 2, 73
			ABPS_check = MsgBox("Is this the right ABPS?", vbYesNo +vbQuestion)
			If ABPS_check = vbYes then
				ABPS_found = True
				exit do
			END IF
			If (ABPS_check = vbNo AND current_panel_number = panel_number) then
				ABPS_found = False
				script_end_procedure("Unable to find another ABPS. Please review the case, and run the script again if applicable.")
			End if
			transmit
		Loop until current_panel_number = panel_number
	End if

	'Updating the ABPS panel
	PF9
	EMReadScreen error_check, 2, 24, 2	'making sure we can actually update this case.
	error_check = trim(error_check)
	If error_check <> "" then script_end_procedure("Unable to update this case. Please review case, and run the script again if applicable.")

	EMWriteScreen "Y", 4, 73			'Support Coop Y/N field
	IF good_cause_droplist <> "Determination" THEN EMWriteScreen "P", 5, 47			'Good Cause status field
	EMWriteScreen "N", 7, 47			'Sup evidence Y/N field (defaulted to N during this process)
	Call create_MAXIS_friendly_date(claim_date, 0, 5, 73)

	'converting the good cause reason from reason_droplist to the applicable MAXIS coding
	If reason_droplist = "Potential phys harm/Child"		then claim_reason = "1"
	If reason_droplist = "Potential Emotnl harm/Child"	 	then claim_reason = "2"
	If reason_droplist = "Potential phys harm/Caregiver" 	then claim_reason = "3"
	If reason_droplist = "Potential Emotnl harm/Caregiver" 	then claim_reason = "4"
	If reason_droplist = "Cncptn Incest/Forced Rape" 		then claim_reason = "5"
	If reason_droplist = "Legal adoption Before Court" 		then claim_reason = "6"
	If reason_droplist = "Parent Gets Preadoptn Svc" 		then claim_reason = "7"
	EMWriteScreen claim_reason, 6, 47
	PF3
	PF3	'to move past non-inhibiting warning messages on ABPS
	EMReadScreen ABPS_screen, 4, 2, 46		'if inhibiting error exists, this will catch it and instruct the user to update ABPS
	msgbox ABPS_screen
	If ABPS_screen = "ABPS" then script_end_procedure("An error occurred on the ABPS panel. Please update the panel before using the script again.")

		'-----------------------------------------------------------------------------------------------------Case note & email sending
	start_a_blank_CASE_NOTE
  IF GC_sent_checkbox = CHECKED Call write_bullet_and_variable_in_case_note("Sent copy of good cause statement to client")
  Call write_bullet_and_variable_in_case_note("Send DHS 3632 to client")
  Call write_bullet_and_variable_in_case_note("Child member #’s")
  Call write_bullet_and_variable_in_case_note("ABPS name", )
  Call write_bullet_and_variable_in_case_note("Good cause claim date", claim_date)
  Call write_bullet_and_variable_in_case_note("Which programs are active METS active? Mets information field.")
  Call write_bullet_and_variable_in_case_note("What is GC incomplete for?", mandatory field).
  Call write_bullet_and_variable_in_case_note("Other information field")
  Call write_bullet_and_variable_in_case_note("Good cause claim date", claim_date)Email Tina to add to the next month’s agenda
	IF good_cause_droplist = "Application Review" THEN Call write_variable_in_case_note("***Good Cause Application Review - Incomplete***")
	Call write_variable_in_case_note("Good Cause Application Review - Complete")
	Call write_bullet_and_variable_in_case_note("Good cause claim date", claim_date)
	Call write_bullet_and_variable_in_case_note("Reason for claiming good cause", reason_droplist)
	Call write_variable_in_case_note("*DHS-2338 is in ECF, and fully completed by parent/caregiver.")
  Call write_bullet_and_variable_in_case_note("Verifs requested", verifs_req)
	Call write_bullet_and_variable_in_case_note("Other notes", other_notes)
	Call write_variable_in_case_note("---")
	Call write_variable_in_case_note(worker_signature)
script_end_procedure("Success! MAXIS has been updated, and an email has been sent to the Good Cause Committee.")
END IF
