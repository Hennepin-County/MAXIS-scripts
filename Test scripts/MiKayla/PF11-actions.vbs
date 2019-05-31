'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PF11 ACTIONS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 120                	'manual run time in seconds
STATS_denomination = "C"       		'M is for each MEMBER
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
call changelog_update("05/13/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'FUNCTIONS==================================================================================================================
Function Generate_Client_List(list_for_dropdown)

	memb_row = 5

	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	Do
		EMReadScreen ref_numb, 2, memb_row, 3
		If ref_numb = "  " Then Exit Do
		EMWriteScreen ref_numb, 20, 76
		TRANSMIT
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		client_info = client_info & "~" & ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
		memb_row = memb_row + 1
	Loop until memb_row = 20

	client_info = right(client_info, len(client_info) - 1)
	client_list_array = split(client_info, "~") 'this is where the tilday goeas away'

	For each person in client_list_array
		list_for_dropdown = list_for_dropdown & chr(9) & person
	Next

End Function
'THE SCRIPT=================================================================================================================
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call check_for_maxis(FALSE) 'checking for passord out, brings up dialog'
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


If MAXIS_case_number <> "" Then 		'If a case number is found the script will get the list of
	Call Generate_Client_List(HH_Memb_DropDown)
End If

'Running the dialog for case number and client
Do
	err_msg = ""
	'intial dialog for user to select a SMRT action
	BeginDialog PF11_actions_dialog, 0, 0, 196, 130, "PF11 Action"
	  EditBox 55, 5, 40, 15, maxis_case_number
	  DropListBox 75, 25, 115, 15, "Select One:" & HH_Memb_DropDown, clt_to_update
	  DropListBox 75, 45, 115, 15, "Select One:"+chr(9)+"PMI merge request"+chr(9)+"Non-actionable DAIL removal"+chr(9)+"Case note removal request"+chr(9)+"MFIP New Spouse Income"+chr(9)+"Other", PF11_actions
	  Text 5, 85, 185, 20, "The system being down, issuance problems, or any type     of emergency should NOT be reported via a PF11."
	  EditBox 75, 65, 115, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 85, 110, 50, 15
	    CancelButton 140, 110, 50, 15
	    PushButton 105, 5, 85, 15, "HH MEMB SEARCH", search_button
	  Text 5, 10, 45, 10, "Case number:"
	  Text 5, 30, 70, 10, "Household member:"
	  Text 5, 50, 65, 10, "Select PF11 action:"
	  Text 5, 70, 60, 10, "Worker signature:"
	EndDialog


	Dialog PF11_actions_dialog
	If ButtonPressed = cancel Then StopScript
	If ButtonPressed = search_button Then
		If MAXIS_case_number = "" Then
			MsgBox "Cannot search without a case number, please try again."
		Else
			HH_Memb_DropDown = ""
			Call Generate_Client_List(HH_Memb_DropDown)
			err_msg = err_msg & "Start Over"
		End If
	End If
	If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "You must enter a valid case number."
	If clt_to_update = "Select One:" Then err_msg = err_msg & vbNewLine & "Please pick a client whose EMPS panel you need to update."
	If err_msg <> "" AND left(err_msg, 10) <> "Start Over" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
Loop until err_msg = ""

IF PF11_actions = "Other" THEN
    BeginDialog other_dialog, 0, 0, 326, 90, "Other"
      EditBox 70, 10, 240, 15, request_reason
      EditBox 70, 30, 240, 15, other_notes
      EditBox 70, 50, 240, 15, action_taken
      ButtonGroup ButtonPressed
    	OkButton 205, 70, 50, 15
    	CancelButton 260, 70, 50, 15
      Text 25, 35, 45, 10, "Other Notes:"
      Text 15, 55, 50, 10, " Actions Taken:"
      Text 5, 15, 60, 10, "Describe Problem:"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog other_dialog
    		if ButtonPressed = 0 then StopScript
    		if IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
    		If SMRT_actions = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Select a PF11 action."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

	'Do
		Call Navigate_to_MAXIS_screen ("STAT", "MEMB")
		'redefine ref_numb'
		'MEMB_number = left(ref_numb, len(client_info) - 2)
		MEMB_number = left(clt_to_update, 2)	'Settin the reference number
		'msgbox MEMB_number
		EMWriteScreen MEMB_number, 20, 76
		TRANSMIT
		EMReadScreen client_first_name, 12, 6, 63
		'replace(client_first_name, "_", "")
		EMReadScreen client_last_name, 25, 6, 30
		'replace(client_last_name, "_", "")
		EMReadScreen client_DOB_month, 02, 08, 42
		EMReadScreen client_DOB_date, 02, 08, 45
		EMReadScreen client_DOB_year, 04, 08, 48

		PF2 'going to PERS'
		EMReadScreen nav_check, 4, 2, 47
		'msgbox nav_check
		EMWriteScreen client_last_name, 04, 36
		client_last_name = trim(client_last_name)
		client_last_name = replace(client_last_name, "_", "")
		'msgbox client_last_name
		EMWriteScreen client_first_name, 10, 36
		client_first_name = trim(client_first_name)
		client_first_name = replace(client_first_name, "_", "")
		'MsgBox client_first_name
		EMWriteScreen client_DOB_month, 11, 53
		EMWriteScreen client_DOB_date, 11, 56
		EMWriteScreen client_DOB_year, 11, 59

		TRANSMIT

	'msgbox "where are we"
	'' PMI NBR ASSIGNED THRU SMI OR PMIN - NO MAXIS CASE EXISTS
	If PF11_actions = "PMI merge request" then
	BeginDialog PMI_merge_dialog, 0, 0, 276, 125, "PMI merge requested"
	  EditBox 80, 5, 50, 15, PMI_number
	  EditBox 80, 25, 50, 15, PMI_number_two
	  EditBox 220, 25, 50, 15, second_case_number
	  DropListBox 80, 45, 190, 15, "Select One:"+chr(9)+"METS case opened"+chr(9)+"PMI number assigned thru SMI or PMIN"+chr(9)+"Incorrect information on case", reason_request_dropdown
	  EditBox 80, 65, 190, 15, action_taken
	  EditBox 80, 85, 190, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 165, 105, 50, 15
	    CancelButton 220, 105, 50, 15
	  Text 5, 70, 50, 10, " Actions taken:"
	  Text 5, 10, 65, 10, "PMI on this case:"
	  Text 5, 110, 160, 10, "**If additional PMI(s) are found add to other notes"
	  Text 5, 90, 45, 10, "Other notes:"
	  Text 5, 30, 60, 10, "Duplicate PMI(s):"
	  Text 5, 50, 65, 10, "Reason for request:"
	  Text 150, 30, 65, 10, "Other case number:"
	EndDialog
		Do
			Do
				err_msg = ""
				Dialog PMI_merge_dialog
				cancel_confirmation
				If PMI_number = "" THEN err_msg = err_msg & vbNewLine & "* Enter the PMI on this case."
				If trim(second_case_number) = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the second case number, if none enter N/A."
				'If trim(action_taken) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the actions taken."
				If trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
			Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
		LOOP UNTIL check_for_password(are_we_passworded_out) = False
	END IF

	'Loop until nav_check = "PERS"
	'now we are at PERS shoing all the names'
	PF11

	'Problem.Reporting
	EMReadScreen nav_check, 4, 1, 27
 	IF nav_check = "Prob" THEN
		EMWriteScreen "PMI merge request for case number: " & maxis_case_number, 05, 07
		'msgbox "are we writing"
		'EMWriteScreen client_SSN, 06, 07
		EMWriteScreen "Current case PMI number: " & PMI_number, 06, 07
		IF PMI_number_two <> "" THEN EMWriteScreen "Duplicate PMI number: " & PMI_number_two, 07, 07
		'msgbox PMI_number
		EMWriteScreen "Reason for request: " & reason_request_dropdown, 08, 07
		'msgbox reason_request_dropdown
		IF second_case_number <> "" THEN EMWriteScreen "Other case number: " & second_case_number, 08, 07
		TRANSMIT
		EMReadScreen task_number, 7, 3, 27
		'msgbox task_number
		TRANSMIT
		'back_to_self
		PF3 ''-self'
		PF3 '- MEMB'
	ELSE
		MsgBox "Could not reach PF11."
	END IF

	reminder_date = dateadd("d", 5, date)

	CALL start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	CALL write_variable_in_case_note("---PF11 requested " & PF11_actions & "---")
	CALL write_bullet_and_variable_in_CASE_NOTE("Reason for request", reason_request_dropdown)
	CALL write_bullet_and_variable_in_CASE_NOTE("Task number", task_number)
	CALL write_bullet_and_variable_in_CASE_NOTE("Requested for", client_info)
	If PMI_number <> "" THEN Call write_bullet_and_variable_in_CASE_NOTE("PMI number", PMI_number)
	If PMI_number_two <> "" then Call write_bullet_and_variable_in_CASE_NOTE("Duplicate PMI number", PMI_number_two)
	CALL write_bullet_and_variable_in_CASE_NOTE("Associated case number", second_case_number)
	CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	If outlook_reminder = True then call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
	CALL write_variable_in_CASE_NOTE ("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
	PF3

	'Outlook appointment is created in prior to the case note being created
	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "PF11 check: " & PF11_actions & " for " & MAXIS_case_number, "", "", TRUE, 5, "")
	outlook_reminder = True

script_end_procedure("It worked!")
