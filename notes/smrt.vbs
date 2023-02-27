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
BeginDialog Dialog1, 0, 0, 186, 85, "SMRT initial dialog"
  EditBox 85, 5, 60, 15, maxis_case_number
  DropListBox 85, 25, 95, 15, "Select one..."+chr(9)+"Initial request"+chr(9)+"ISDS referral completed"+chr(9)+"Determination received", SMRT_actions
  EditBox 85, 45, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 65, 50, 15
    CancelButton 130, 65, 50, 15
  Text 30, 10, 45, 10, "Case number:"
  Text 5, 30, 75, 10, "Select a SMRT action:"
  Text 15, 50, 65, 10, "Worker Signature:"
EndDialog

Do
    Do
        err_msg = ""
        Dialog Dialog1
        cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg,"*")
        If SMRT_actions = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Select a SMRT action."
        If trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = False

Call generate_client_list(all_the_clients, "Select or Type")
initial_request_note_found = False

If SMRT_actions <> "Initial request" then
	Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
	too_old_date = DateAdd("M", -3, date)              'We don't need to read notes from before the CAF date

	note_row = 5
	Do
		EMReadScreen note_date, 8, note_row, 6                  'reading the note date

		EMReadScreen note_title, 55, note_row, 25               'reading the note header
		note_title = trim(note_title)

		If note_title = "---Initial SMRT referral requested---" Then

			initial_request_note_found = True
			Call write_value_and_transmit("X", note_row, 3)

			in_note_row = 4
			Do
				EMReadScreen note_line, 78, in_note_row, 3
				note_line = trim(note_line)

				If left(note_line, 20) = "* SMRT requested for" Then
					If InStr(note_line, "MEMB") <> 0 Then
						memb_numb_pos = InStr(note_line, "MEMB")
						memb_numb_pos = memb_numb_pos + 5
						memb_number = mid(note_line, memb_numb_pos, 2)
						memb_name_pos = memb_numb_pos + 3
						memb_name_len = len(note_line) - memb_name_pos
						SMRT_member_name = right(note_line, memb_name_len)
						SMRT_member = memb_number & " - " & SMRT_member_name
					End If
					If InStr(note_line, "MEMB") = 0 Then
						SMRT_member = right(note_line, len(note_line)-22)
					End If
				End If
				If left(note_line, 5) = "* Age" Then memb_age = right(note_line, len(note_line)-7)
				If left(note_line, 28) = "* SMRT referral completed on" Then referral_date = right(note_line, len(note_line)-30)
				If left(note_line, 23) = "* Is referral expedited" Then referred_exp = right(note_line, len(note_line)-25)
				If left(note_line, 17) = "* SMRT start date" Then SMRT_start_date = right(note_line, len(note_line)-19)
				If left(note_line, 18) = "* Expedited reason" Then
					expedited_reason = right(note_line, len(note_line)-20)
					EMReadScreen next_note_line, 78, in_note_row+1, 3
					If left(next_note_line, 10) = "          " Then
						expedited_reason = expedited_reason & " " & trim(next_note_line)
						in_note_row = in_note_row + 1
					End If
				End If
				If left(note_line, 21) = "* Reason for referral" Then
					referral_reason = right(note_line, len(note_line)-23)
					EMReadScreen next_note_line, 78, in_note_row+1, 3
					If left(next_note_line, 10) = "          " Then
						referral_reason = referral_reason & " " & trim(next_note_line)
						in_note_row = in_note_row + 1
					End If
				End If
				' If left(note_line, 1) =  1234567890123456789012345678

				in_note_row = in_note_row + 1
				If in_note_row = 18 Then
					PF8
					in_note_row = 4
					EMReadScreen end_of_note, 9, 24, 14
					If end_of_note = "LAST PAGE" Then Exit Do
				End If
			Loop until note_line = ""
			PF3
			Exit Do
		End If

		if note_date = "        " then Exit Do                                      'if we are at the end of the list of notes - we can't read any more

		note_row = note_row + 1
		if note_row = 19 then
			note_row = 5
			PF8
			EMReadScreen check_for_last_page, 9, 24, 14
			If check_for_last_page = "LAST PAGE" Then Exit Do
		End If
		EMReadScreen next_note_date, 8, note_row, 6
		if next_note_date = "        " then Exit Do
	Loop until DateDiff("d", too_old_date, next_note_date) <= 0
End If

If SMRT_actions = "Initial request" or initial_request_note_found = False then
    '-------------------------------------------------------------------------------------------------DIALOG
    referral_date = date & ""
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 301, 75, "Initial SMRT referral dialog"
      ComboBox 80, 10, 215, 45, all_the_clients+chr(9)+SMRT_member, SMRT_member
      EditBox 80, 30, 50, 15, SMRT_start_date
      DropListBox 230, 30, 65, 15, "Select one..."+chr(9)+"No"+chr(9)+"Yes", referred_exp
      EditBox 110, 50, 50, 15, referral_date
      ButtonGroup ButtonPressed
        OkButton 190, 50, 50, 15
        CancelButton 245, 50, 50, 15
      Text 5, 15, 70, 10, "SMRT requested for: "
      Text 20, 35, 60, 10, "SMRT start date:"
      Text 155, 35, 70, 10, "Is referral expedited?"
      Text 5, 55, 100, 10, "Date SMRT referral completed:"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		If SMRT_member = "Select or Type" or trim(SMRT_member) = "" THEN err_msg = err_msg & vbNewLine & "* Select or Enter the member name the SMRT referral is for."
    		If isdate(referral_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid referral date."
			If referred_exp = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Is the referral expedited?"
			If isdate(SMRT_start_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid SMRT start date."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = False

    memb_number = left(SMRT_member, 2)
    If IsNumeric(memb_number) = true Then
		SMRT_member_name = right(SMRT_member, len(SMRT_member)-5)
		Call Navigate_to_MAXIS_screen("STAT", "MEMB")
		Call write_value_and_transmit(memb_number, 27, 76)
		EMReadScreen memb_age, 3, 8, 76
		memb_age = trim(memb_age)
		If memb_age = "" Then memb_age = 0
		memb_age = memb_age * 1
    Else
		SMRT_member_name = SMRT_member
		memb_number = ""
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 336, 45, "SMRT Member Age"
		  EditBox 150, 25, 60, 15, memb_age
		  ButtonGroup ButtonPressed
			OkButton 280, 25, 50, 15
		  Text 10, 10, 320, 10, "What is the age of "& SMRT_member &", the person the SMRT referral is for?"
		  Text 125, 30, 20, 10, "Age:"
		EndDialog

		Do
			Do
				err_msg = ""
				Dialog Dialog1
				cancel_without_confirmation
				If IsNumeric(memb_age) = False Then err_msg = err_msg & "* Enter the persons age as a number."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
			Loop until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		LOOP UNTIL are_we_passworded_out = False
		memb_age = memb_age * 1
    End If

    list_of_referral_reasons = "Select One..."
    list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Needs services under a home and community-based waiver program"
    list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks a managed-care exclusion due to a disability"
    If memb_age < 19 Then
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks a Family Support Grant (FSG)"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks coverage under the TEFRA (Tax Equity and Fiscal Responsibility Act) Option"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks coverage under Medical Assistance for Employed Persons with Disabilities (MA-EPD) "
    ElseIf memb_age = 19 or memb_age = 20 Then
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks a Family Support Grant (FSG)"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks Medical Assistance for Employed Persons with Disabilities (MA-EPD)"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks enrollment in Special Needs Basic Care (SNBC)"
    ElseIf memb_age > 20 Then
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks Medical Assistance for Employed Persons with Disabilities (MA-EPD)"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks enrollment in Special Needs Basic Care (SNBC)"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks MA with a spenddown and is without children"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Requires a continuing disability review at age 65 for MA-EPD"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Is 65 years old or older and setting up a pooled trust"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks disability for a trust beneficiary (non-MA applicant or enrollee)"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks disability for a child of any age to establish an asset transfer penalty exception"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Applicant is deceased and needs a disability determination for MA eligibility"
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Other"
    End If

    list_of_expedited_reasons = "Select One..."
    list_of_expedited_reasons = list_of_expedited_reasons+chr(9)+"The person has a condition that appears on the SSA Compassionate Allowance Listing (CAL)."
    list_of_expedited_reasons = list_of_expedited_reasons+chr(9)+"The person is awaiting discharge from a facility and can be discharged immediately if MA is approved."
    list_of_expedited_reasons = list_of_expedited_reasons+chr(9)+"The person has a potentially life-threatening situation and requires immediate treatment or medication."
    list_of_expedited_reasons = list_of_expedited_reasons+chr(9)+"Other circumstances that may jeopardize a resident's benefits. The circumstance is reviewed and accepted on a case-by-case basis."

    Dialog1 = "" 'Blanking out previous dialog detail
    dlg_len = 125
    If referred_exp = "Yes" Then dlg_len = 155
    y_pos = dlg_len - 50
    BeginDialog Dialog1, 0, 0, 446, dlg_len, "Initial SMRT referral dialog"
      Text 5, 10, 195, 10, "SMRT requested for: " & SMRT_member
      Text 5, 20, 175, 10, "Date SMRT referral completed: " & referral_date
      Text 5, 30, 105, 10, "SMRT start date: " & SMRT_start_date
      Text 5, 45, 65, 10, "Reason for referral:"
      DropListBox 5, 55, 435, 45, list_of_referral_reasons+chr(9)+referral_reason, referral_reason
      If referred_exp = "Yes" Then
        Text 5, 75, 110, 10, "EXPEDTIED REFERRAL Reason:"
        DropListBox 5, 85, 435, 45, list_of_expedited_reasons+chr(9)+expedited_reason, expedited_reason
      End If
      Text 5, y_pos, 80, 10, "Additional SMRT Notes"
      EditBox 5, y_pos + 10, 435, 15, other_notes
      Text 5, y_pos + 35, 90, 10, "ECF Workflow Completed?"
      DropListBox 95, y_pos + 30, 75, 45, "Select"+chr(9)+"Yes"+chr(9)+"No", ecf_workflow_done
      ButtonGroup ButtonPressed
        OkButton 335, y_pos + 30, 50, 15
        CancelButton 390, y_pos + 30, 50, 15
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation

			If (referred_exp = "Yes" and expedited_reason = "Select One...") THEN err_msg = err_msg & vbNewLine & "* Enter the expedited reason."
			If referral_reason = "Select One..." THEN err_msg = err_msg & vbNewLine & "* Enter the reason for the referral."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = False

	If SMRT_actions = "Initial request" Then
		start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
		Call write_variable_in_CASE_NOTE("---Initial SMRT referral requested---")
		If memb_number = "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
		If memb_number <> "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", "MEMB " & memb_number & " - " & SMRT_member_name)
		Call write_bullet_and_variable_in_CASE_NOTE("Age", memb_age)
		Call write_bullet_and_variable_in_CASE_NOTE("SMRT referral completed on", referral_date)
		Call write_bullet_and_variable_in_CASE_NOTE("Is referral expedited", referred_exp)
		If referred_exp = "Yes" then Call write_bullet_and_variable_in_CASE_NOTE("Expedited reason", expedited_reason)
		Call write_bullet_and_variable_in_CASE_NOTE("Reason for referral", referral_reason)
		Call write_bullet_and_variable_in_CASE_NOTE("SMRT start date", SMRT_start_date)
		Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
		If ecf_workflow_done = "Yes" then call write_variable_in_CASE_NOTE("* Workflow process has been completed in case file system.")
		Call write_variable_in_CASE_NOTE ("---")
		call write_variable_in_CASE_NOTE(worker_signature)

		end_msg = "SMRT Action for Initial Request noted on Case."
	End If
END If

initial_SMRT_member = SMRT_member
initial_memb_number = memb_number
initial_SMRT_member_name = SMRT_member_name
initial_SMRT_start_date = SMRT_start_date
initial_referral_date = referral_date
initial_referred_exp = referred_exp
initial_expedited_reason = expedited_reason
initial_referral_reason = referral_reason


If SMRT_actions = "ISDS referral completed" then
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
	' BeginDialog Dialog1, 0, 0, 326, 130, "ISDS referral completed for SMRT"
	' 	EditBox 80, 10, 75, 15, SMRT_member
	' 	EditBox 225, 10, 50, 15, referral_date
	' 	EditBox 80, 35, 75, 15, prog_requested
	' 	EditBox 225, 35, 50, 15, SMRT_start_date
	' 	EditBox 80, 60, 240, 15, other_notes
	' 	EditBox 80, 85, 240, 15, action_taken
	' 	EditBox 80, 110, 130, 15, worker_signature
	' 	ButtonGroup ButtonPressed
	' 	OkButton 215, 110, 50, 15
	' 	CancelButton 270, 110, 50, 15
	' 	Text 10, 65, 65, 10, "Other SMRT notes:"
	' 	Text 10, 40, 65, 10, "Program requested:"
	' 	Text 25, 90, 50, 10, " Actions taken:"
	' 	Text 165, 15, 55, 10, "Completion date:"
	' 	Text 5, 15, 70, 10, "SMRT requested for: "
	' 	Text 165, 40, 60, 10, "SMRT start date:"
	' 	Text 15, 115, 60, 10, "Worker Signature:"
	' EndDialog

	BeginDialog Dialog1, 0, 0, 446, 175, "Initial SMRT referral dialog"
		EditBox 130, 105, 50, 15, isds_referral_date
		EditBox 5, 135, 435, 15, other_notes
		ButtonGroup ButtonPressed
			OkButton 335, 155, 50, 15
			CancelButton 390, 155, 50, 15
			PushButton 305, 10, 135, 15, "The Request Information is Incorrect", change_details_btn
		Text 5, 10, 195, 10, "SMRT requested for: " & SMRT_member
		Text 5, 20, 175, 10, "Date SMRT referral completed: " & referral_date
		Text 5, 30, 125, 10, "SMRT start date: " & SMRT_start_date
		Text 5, 45, 70, 10, "Reason for referral:"
		Text 5, 55, 435, 10, referral_reason
		If referred_exp = "Yes" Then
			Text 5, 75, 135, 10, "Expedited Referral Requested. Reason:"
			Text 5, 85, 435, 10, expedited_reason
		Else
			Text 5, 75, 135, 10, "Expedited Referral was NOT Requested."
		End If
		Text 5, 110, 120, 10, "SMRT Referral Submitted to ISDS on "
		Text 5, 125, 80, 10, "Additional SMRT Notes"
		Text 185, 110, 50, 10, "(date)"
	EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		If isdate(isds_referral_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter the date the ISDS referral was submitted as a valid date."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = False

	memb_number = left(SMRT_member, 2)
	If IsNumeric(memb_number) = true Then
		SMRT_member_name = right(SMRT_member, len(SMRT_member)-5)
    Else
		SMRT_member_name = SMRT_member
		memb_number = ""
	End If

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
    LOOP UNTIL are_we_passworded_out = False

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
