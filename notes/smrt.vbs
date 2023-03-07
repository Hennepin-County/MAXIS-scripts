'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - SMRT.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180           'manual run time in seconds
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
call changelog_update("03/07/2023", "The SMRT Script has been updated to better support the gathering of details for each step of the SMRT process.##~####~##For Initial request, the script will pull the allowable reasons based on person age to reduce manual entry of information and reduce errors and miscommunication.##~## ##~##For workers entering the SMRT request to ISDS the script will attempt to read the information from CASE/NOTE to autofill the information about the request. ##~##Ensuring use of this script at Initial Request will help support each step of this process.##~####~##When a determination is received, the script will also attempt to read DISA to autofill the certification dates.##~####~##An additional function has been added to create a CASE/NOTE in the instance that an Initial Request cannot be submitted to ISDS, so case actions are clear.##~##", "Casey Love, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
Call changelog_update("06/13/2020", "Since there are updates to the requirements for submitting a SMRT Referral in ISDS, we are reviewing the functionality of this script. ##~## If there are specific changes, fields, information, or functionality that would make your work with this script easier, pleae contact us. ##~## ##~## Email us at HSPH.EWS.BlueZoneScripts@hennepin.us or submit an 'Error Report' at the end of the script run.##~##", "Casey Love, Hennepin County")
call changelog_update("01/19/2017", "Initial version.", "Ilse Ferris, Hennepin County")
call changelog_update("11/29/2017", "Update script for denials to remove start date.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function gather_SMRT_request_info()
	'This is the functionality that will collect the name of the person requesting SMRT and all request details.
	'This is in a function because this might need to be called from any step in the process and from any dialog.

	'this is a dialog to select the person, start date, aand expedited information
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 301, 75, "Initial SMRT referral dialog"
      ComboBox 80, 10, 215, 45, all_the_clients+chr(9)+SMRT_member, SMRT_member
      EditBox 80, 30, 50, 15, SMRT_start_date
      DropListBox 230, 30, 65, 15, "Select one..."+chr(9)+"No"+chr(9)+"Yes", referred_exp
      EditBox 110, 50, 50, 15, referral_request_date
      ButtonGroup ButtonPressed
        OkButton 190, 50, 50, 15
        CancelButton 245, 50, 50, 15
      Text 5, 15, 70, 10, "SMRT requested for: "
      Text 20, 35, 60, 10, "SMRT start date:"
      Text 155, 35, 70, 10, "Is referral expedited?"
      Text 5, 55, 100, 10, "Date SMRT referral requested:"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		If SMRT_member = "Select or Type" or trim(SMRT_member) = "" THEN err_msg = err_msg & vbNewLine & "* Select or Enter the member name the SMRT referral is for."
    		If isdate(referral_request_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid referral date."
			If referred_exp = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Is the referral expedited?"
			If isdate(SMRT_start_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid SMRT start date."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = False

	'trying to find the name of the person from MEMB if selected as one of the household members or as a dialog input
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
	'these are the options for referral reason and expedited reasons for selection in a dialog.
    list_of_referral_reasons = "Select One..."
    list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Needs services under a home and community-based waiver program."
    list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks a managed-care exclusion due to a disability."
    If memb_age < 19 Then
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks a Family Support Grant (FSG)."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks coverage under the TEFRA (Tax Equity and Fiscal Responsibility Act) Option."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks coverage under Medical Assistance for Employed Persons with Disabilities (MA-EPD)."
    ElseIf memb_age = 19 or memb_age = 20 Then
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks a Family Support Grant (FSG)."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks Medical Assistance for Employed Persons with Disabilities (MA-EPD)."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks enrollment in Special Needs Basic Care (SNBC)."
    ElseIf memb_age > 20 Then
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks Medical Assistance for Employed Persons with Disabilities (MA-EPD)."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks enrollment in Special Needs Basic Care (SNBC)."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks MA with a spenddown and is without children."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Requires a continuing disability review at age 65 for MA-EPD."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Is 65 years old or older and setting up a pooled trust."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks disability for a trust beneficiary (non-MA applicant or enrollee)."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Seeks disability for a child of any age to establish an asset transfer penalty exception."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Applicant is deceased and needs a disability determination for MA eligibility."
		list_of_referral_reasons = list_of_referral_reasons+chr(9)+"Other"
    End If
    list_of_expedited_reasons = "Select One..."
    list_of_expedited_reasons = list_of_expedited_reasons+chr(9)+"The person has a condition that appears on the SSA Compassionate Allowance Listing (CAL)."
    list_of_expedited_reasons = list_of_expedited_reasons+chr(9)+"The person is awaiting discharge from a facility and can be discharged immediately if MA is approved."
    list_of_expedited_reasons = list_of_expedited_reasons+chr(9)+"The person has a potentially life-threatening situation and requires immediate treatment or medication."
    list_of_expedited_reasons = list_of_expedited_reasons+chr(9)+"Other circumstances that may jeopardize a resident's benefits. The circumstance is reviewed and accepted on a case-by-case basis."

	'this dialog is to select the reason details from the droplist options
    Dialog1 = "" 'Blanking out previous dialog detail
    dlg_len = 125
    If referred_exp = "Yes" Then dlg_len = 155
    y_pos = dlg_len - 50
    BeginDialog Dialog1, 0, 0, 446, dlg_len, "Initial SMRT referral dialog"
      Text 5, 10, 195, 10, "SMRT requested for: " & SMRT_member
      Text 5, 20, 175, 10, "Date SMRT referral requested: " & referral_request_date
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
end function

'these variables are dimmed here because we define them within the functions and this keeps us from having to pass the variables through
Dim all_the_clients, SMRT_member, SMRT_start_date, referred_exp, referral_request_date, memb_number, SMRT_member_name
Dim memb_age, list_of_referral_reasons, list_of_expedited_reasons, referral_reason, expedited_reason, other_notes, ecf_workflow_done

'Start of the Script ===================================================================================
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)		'collecting the case number
end_msg = "SMRT NOTE Script completed, SMRT informmation has been entered into the CASE/NOTE."		'setting the end message
smrt_request_info_changed = False																	'defaulting the information about SMRT info being changed
initial_request_note_found = False																	'defaulting if the requesting information was found from the note

'intial dialog for user to select a SMRT action, confirm Case Number and the Worker Signature
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 256, 85, "SMRT initial dialog"
  EditBox 85, 5, 60, 15, maxis_case_number
  DropListBox 85, 25, 165, 15, "Select one..."+chr(9)+"Initial request"+chr(9)+"ISDS referral completed"+chr(9)+"SMRT Referral NOT Submitted"+chr(9)+"Determination received", SMRT_actions
  EditBox 85, 45, 165, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 145, 65, 50, 15
    CancelButton 200, 65, 50, 15
	PushButton 160, 7, 90, 13, "HSR Manual Instructions", hsr_manual_information_btn
  Text 30, 10, 45, 10, "Case number:"
  Text 5, 30, 75, 10, "Select a SMRT action:"
  Text 15, 50, 65, 10, "Worker Signature:"
EndDialog

Do
    Do
        err_msg = ""
        Dialog Dialog1
        cancel_without_confirmation
		If ButtonPressed = hsr_manual_information_btn Then		'this runs the dialogs to change the referral information using the dialogs in function
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/State_Medical_Review_Team_(SMRT).aspx"
			err_msg = "LOOP"
		End If
        Call validate_MAXIS_case_number(err_msg,"*")
        If SMRT_actions = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Select a SMRT action."
        If trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
    	IF err_msg <> "" and left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = False

'This gathers all the client information from MAXIS
Call generate_client_list(all_the_clients, "Select or Type")

'If requested action is not the initial request, we are going to check CASE/NOTE to see if we can find a NOTE from a previous run to autofill the information
If SMRT_actions <> "Initial request" then
	referred_exp = "No"				'defaulting the expedited information
	ecf_workflow_done = "Yes"		'defaulting if the workflow was completed
	Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
	too_old_date = DateAdd("M", -3, date)              			'We will look back to  months

	note_row = 5												'this is the beginning of the list of case notes
	Do
		EMReadScreen note_date, 8, note_row, 6                  'reading the note date

		EMReadScreen note_title, 55, note_row, 25               'reading the note header
		note_title = trim(note_title)
		If note_title = "---SMRT NOT submitted to ISDS---"	Then Exit Do		'if the most recent note was to NOT submit the referral, we don't want to autofill any details
		'if the note is found from the initial request or the completed referral, it will read the details from that note to autofil the referral information.
		If note_title = "---Initial SMRT referral requested---" or note_title = "---ISDS referral completed for SMRT---" Then
			STATS_manualtime = STATS_manualtime + 200			'adding to manualtime information because autofilling will save additional time
			initial_request_note_found = True					'setting this boolean to true if a note was found
			Call write_value_and_transmit("X", note_row, 3)		'opening the CASE/NOTE

			in_note_row = 4										'this is the beginning of the CASE/NOTE information
			Do
				EMReadScreen note_line, 78, in_note_row, 3		'reading the line of the notes
				note_line = trim(note_line)						'trimming the note line

				If left(note_line, 20) = "* SMRT requested for" Then				'If the line is for the person information we will read the detail and format it.
					If InStr(note_line, "MEMB") <> 0 Then
						memb_numb_pos = InStr(note_line, "MEMB")
						memb_numb_pos = memb_numb_pos + 5
						memb_number = mid(note_line, memb_numb_pos, 2)
						memb_name_pos = memb_numb_pos + 3
						memb_name_len = len(note_line) - memb_name_pos
						SMRT_member_name = right(note_line, memb_name_len)
						SMRT_member_name = trim(SMRT_member_name)
						SMRT_member = memb_number & " - " & SMRT_member_name
					End If
					If InStr(note_line, "MEMB") = 0 Then
						SMRT_member = right(note_line, len(note_line)-22)
					End If
				End If
				If left(note_line, 5) = "* Age" Then memb_age = right(note_line, len(note_line)-7)											'information about resident age
				If left(note_line, 28) = "* SMRT referral requested on" Then referral_request_date = right(note_line, len(note_line)-30)	'ifnormation about requested date
				If left(note_line, 23) = "* Is referral expedited" Then referred_exp = right(note_line, len(note_line)-25)					'information about expedited status
				If left(note_line, 27) = "* SMRT requested start date" Then SMRT_start_date = right(note_line, len(note_line)-29)			'information about the SMRT requested start date
				If left(note_line, 28) = "* ISDS Referral Submitted on" Then isds_referral_date = right(note_line, len(note_line)-30)		'information about the referral submitted date
				If left(note_line, 18) = "* Expedited reason" Then										'if the line is about the expedited reason, we might need to read more that one line
					expedited_reason = right(note_line, len(note_line)-20)
					If in_note_row = 17 Then
						PF8
						in_note_row = 3
						EMReadScreen end_of_note, 9, 24, 14
						If end_of_note = "LAST PAGE" Then Exit Do
					End If
					EMReadScreen next_note_line, 78, in_note_row+1, 3
					If left(next_note_line, 10) = "          " Then
						expedited_reason = expedited_reason & " " & trim(next_note_line)
						in_note_row = in_note_row + 1
						EMReadScreen third_note_line, 78, in_note_row+1, 3
						If left(third_note_line, 10) = "          " Then
							expedited_reason = expedited_reason & " " & trim(third_note_line)
							in_note_row = in_note_row + 1
						End If
					End If
				End If
				If left(note_line, 21) = "* Reason for referral" Then									'if the line is about the referral reason, we might need to read more that one line
					referral_reason = right(note_line, len(note_line)-23)
					If in_note_row = 17 Then
						PF8
						in_note_row = 3
						EMReadScreen end_of_note, 9, 24, 14
						If end_of_note = "LAST PAGE" Then Exit Do
					End If
					EMReadScreen next_note_line, 78, in_note_row+1, 3
					' MsgBox "~" & next_note_line & "~" & vbCr & "in_note_row - " & in_note_row
					If left(next_note_line, 10) = "          " Then
						referral_reason = referral_reason & " " & trim(next_note_line)
						in_note_row = in_note_row + 1
						EMReadScreen third_note_line, 78, in_note_row+1, 3
						If left(third_note_line, 10) = "          " Then
							expedited_reason = expedited_reason & " " & trim(third_note_line)
							in_note_row = in_note_row + 1
						End If
					End If
				End If

				in_note_row = in_note_row + 1			'going to the next note line
				If in_note_row = 18 Then				'if we are at line 18, we need to go to the next page of the note
					PF8
					in_note_row = 4
					EMReadScreen end_of_note, 9, 24, 14
					If end_of_note = "LAST PAGE" Then Exit Do
				End If
			Loop until note_line = ""					'going until the line is blank
			PF3			'eaving the note
			Exit Do		'we do not need to read any more notes if the information was found - this leaves the loop for reading each note
		End If

		if note_date = "        " then Exit Do                                      'if we are at the end of the list of notes - we can't read any more

		note_row = note_row + 1					'reading the next line of CASE/NOTEs to find the note we need
		if note_row = 19 then
			note_row = 5
			PF8
			EMReadScreen check_for_last_page, 9, 24, 14
			If check_for_last_page = "LAST PAGE" Then Exit Do
		End If
		EMReadScreen next_note_date, 8, note_row, 6
		if next_note_date = "        " then Exit Do
	Loop until DateDiff("d", too_old_date, next_note_date) <= 0
	Call back_to_SELF
	If initial_request_note_found = True Then						'If the CASE/NOTE was found, we need to make sure all necessary information was read from the NOTE
		If SMRT_member = "" Then initial_request_note_found = False
		If SMRT_start_date = "" Then initial_request_note_found = False
		If referral_request_date = "" Then initial_request_note_found = False
		If referral_reason = "" Then initial_request_note_found = False
		If referred_exp = "Yes" and expedited_reason = "" Then initial_request_note_found = False
	End If
End If

'If the NOTE was not found or the initial reuest is the option, we need to run the function to gather the information from a dialog
If SMRT_actions = "Initial request" or initial_request_note_found = False then
    referral_request_date = date			'defaulting the referral date
	referral_request_date = referral_request_date & ""
	Call gather_SMRT_request_info()

	'If the intital request was the option, we will CASE/NOTE the detail
	If SMRT_actions = "Initial request" Then
		start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
		Call write_variable_in_CASE_NOTE("---Initial SMRT referral requested---")
		If memb_number = "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
		If memb_number <> "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", "MEMB " & memb_number & " - " & SMRT_member_name)
		Call write_bullet_and_variable_in_CASE_NOTE("Age", memb_age)
		Call write_bullet_and_variable_in_CASE_NOTE("SMRT referral requested on", referral_request_date)
		Call write_bullet_and_variable_in_CASE_NOTE("Is referral expedited", referred_exp)
		If referred_exp = "Yes" then Call write_bullet_and_variable_in_CASE_NOTE("Expedited reason", expedited_reason)
		Call write_bullet_and_variable_in_CASE_NOTE("Reason for referral", referral_reason)
		Call write_bullet_and_variable_in_CASE_NOTE("SMRT requested start date", SMRT_start_date)
		Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_notes)
		If ecf_workflow_done = "Yes" then call write_variable_in_CASE_NOTE("* Workflow process has been completed in case file system.")
		Call write_variable_in_CASE_NOTE ("---")
		call write_variable_in_CASE_NOTE(worker_signature)

		end_msg = "SMRT Action for Initial Request noted on Case."
	End If
END If

'we are saving the initial information gathered to identify if there is a change later in the scipt run.
initial_SMRT_member = SMRT_member
initial_memb_number = memb_number
initial_SMRT_member_name = SMRT_member_name
initial_SMRT_start_date = SMRT_start_date
initial_referral_date = referral_request_date
initial_referred_exp = referred_exp
initial_expedited_reason = expedited_reason
initial_referral_reason = referral_reason

'For the option when the ISDS Referral is done.
If SMRT_actions = "ISDS referral completed" then
	If isds_referral_date = "" Then isds_referral_date = date		'defaulted the referral date if not known.
	isds_referral_date = isds_referral_date & ""

    Do
		Do
			err_msg = ""
			Dialog1 = "" 'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 446, 175, "Initial SMRT referral dialog"
				EditBox 130, 105, 50, 15, isds_referral_date
				EditBox 5, 135, 435, 15, other_notes
				ButtonGroup ButtonPressed
					OkButton 335, 155, 50, 15
					CancelButton 390, 155, 50, 15
					PushButton 305, 10, 135, 15, "The Request Information is Incorrect", change_details_btn
				Text 5, 10, 195, 10, "SMRT requested for: " & SMRT_member
				Text 5, 20, 175, 10, "Date SMRT referral completed: " & referral_request_date
				Text 5, 30, 125, 10, "SMRT start date: " & SMRT_start_date
				Text 5, 45, 70, 10, "Reason for referral:"
				Text 5, 55, 435, 10, referral_reason
				If referred_exp = "Yes" Then
					Text 5, 75, 135, 10, "Expedited Referral Requested Reason:"
					Text 5, 85, 435, 10, expedited_reason
				Else
					Text 5, 75, 135, 10, "Expedited Referral was NOT Requested."
				End If
				Text 5, 110, 120, 10, "SMRT Referral Submitted to ISDS on "
				Text 5, 125, 80, 10, "Additional SMRT Notes"
				Text 185, 110, 50, 10, "(date)"
			EndDialog

    		Dialog Dialog1
    		cancel_without_confirmation
			If ButtonPressed = change_details_btn Then		'this runs the dialogs to change the referral information using the dialogs in function
				Call gather_SMRT_request_info()
				err_msg = "LOOP"
			End If

    		If isdate(isds_referral_date) = False THEN err_msg = err_msg & vbNewLine & "* Enter the date the ISDS referral was submitted as a valid date."
    		IF err_msg <> "" and left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = False

	If memb_number = "" Then		'finding the member number if needed
		memb_number = left(SMRT_member, 2)
		If IsNumeric(memb_number) = true Then
			SMRT_member_name = right(SMRT_member, len(SMRT_member)-5)
		Else
			SMRT_member_name = SMRT_member
			memb_number = ""
		End If
	End If

	'identifying if the request information has been changed for details in CASE/NOTE
	If initial_SMRT_member <> SMRT_member Then smrt_request_info_changed = True
	If initial_memb_number <> memb_number Then smrt_request_info_changed = True
	If initial_SMRT_member_name <> SMRT_member_name Then smrt_request_info_changed = True
	If initial_SMRT_start_date <> SMRT_start_date Then smrt_request_info_changed = True
	If initial_referral_date <> referral_request_date Then smrt_request_info_changed = True
	If initial_referred_exp <> referred_exp Then smrt_request_info_changed = True
	If initial_expedited_reason <> expedited_reason Then smrt_request_info_changed = True
	If initial_referral_reason <> referral_reason Then smrt_request_info_changed = True


	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("---ISDS referral completed for SMRT---")
	call write_variable_in_CASE_NOTE("SMRT referral has been submitted to ISDS for review by the State Medical Review Team for determination of disability status.")
    call write_bullet_and_variable_in_CASE_NOTE("ISDS Referral Submitted on", isds_referral_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Other SMRT notes", other_notes)
	If smrt_request_info_changed = True Then call write_variable_in_CASE_NOTE("Referral information has been updated.")
	call write_variable_in_CASE_NOTE("Referral details:")
	If memb_number = "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
	If memb_number <> "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", "MEMB " & memb_number & " - " & SMRT_member_name)
	Call write_bullet_and_variable_in_CASE_NOTE("Age", memb_age)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT referral requested on", referral_request_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Is referral expedited", referred_exp)
	If referred_exp = "Yes" then Call write_bullet_and_variable_in_CASE_NOTE("Expedited reason", expedited_reason)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason for referral", referral_reason)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT requested start date", SMRT_start_date)
    Call write_variable_in_CASE_NOTE ("---")
    call write_variable_in_CASE_NOTE(worker_signature)

	end_msg = end_msg & vbCr & vbCr & "Request for SMRT disability determination submitted to ISDS on " & isds_referral_date
END If

'If the option was selected that the referral is not submitted, this dialog will run.
If SMRT_actions = "SMRT Referral NOT Submitted" Then
	Do
    	Do
    		err_msg = ""
			Dialog1 = "" 'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 446, 200, "Initial SMRT referral dialog"
				EditBox 5, 130, 435, 15, isds_referral_reject_reason
				EditBox 5, 160, 435, 15, other_notes
				ButtonGroup ButtonPressed
					OkButton 335, 180, 50, 15
					CancelButton 390, 180, 50, 15
					PushButton 305, 10, 135, 15, "The Request Information is Incorrect", change_details_btn
				Text 5, 10, 195, 10, "SMRT requested for: " & SMRT_member
				Text 5, 20, 175, 10, "Date SMRT referral requested: " & referral_request_date
				Text 5, 30, 125, 10, "SMRT start date: " & SMRT_start_date
				Text 5, 45, 70, 10, "Reason for referral:"
				Text 5, 55, 435, 10, referral_reason
				If referred_exp = "Yes" Then
					Text 5, 75, 135, 10, "Expedited Referral Requested Reason:"
					Text 5, 85, 435, 10, expedited_reason
				Else
					Text 5, 75, 135, 10, "Expedited Referral was NOT Requested."
				End If
				Text 5, 105, 155, 10, "SMRT Referral will NOT be submitted to ISDS."
				Text 5, 150, 80, 10, "Additional SMRT Notes"
				Text 5, 120, 120, 10, "Reason SMRT cannot be submitted:"
			EndDialog

    		Dialog Dialog1
    		cancel_without_confirmation

			If ButtonPressed = change_details_btn Then		'this runs the dialogs to change the referral information using the dialogs in function
				Call gather_SMRT_request_info()
				err_msg = "LOOP"
			End If
			isds_referral_reject_reason = trim(isds_referral_reject_reason)
			If isds_referral_reject_reason = "" Then err_msg = err_msg & vbNewLine & "* Enter the reason the SMRT referral cannot be sent to ISDS."
    		If len(isds_referral_reject_reason) < 20 THEN err_msg = err_msg & vbNewLine & "* The reason for not submitting the SMRT referral needs more detail in the explaination. Add additonal information to the reason the SMRT referral cannot be submitted."
    		IF err_msg <> "" and left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	Call check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = False

	If memb_number = "" Then		'finding the member number if needed
		memb_number = left(SMRT_member, 2)
		If IsNumeric(memb_number) = true Then
			SMRT_member_name = right(SMRT_member, len(SMRT_member)-5)
		Else
			SMRT_member_name = SMRT_member
			memb_number = ""
		End If
	End If

	'identifying if the request information has been changed for details in CASE/NOTE
	If initial_SMRT_member <> SMRT_member Then smrt_request_info_changed = True
	If initial_memb_number <> memb_number Then smrt_request_info_changed = True
	If initial_SMRT_member_name <> SMRT_member_name Then smrt_request_info_changed = True
	If initial_SMRT_start_date <> SMRT_start_date Then smrt_request_info_changed = True
	If initial_referral_date <> referral_request_date Then smrt_request_info_changed = True
	If initial_referred_exp <> referred_exp Then smrt_request_info_changed = True
	If initial_expedited_reason <> expedited_reason Then smrt_request_info_changed = True
	If initial_referral_reason <> referral_reason Then smrt_request_info_changed = True

	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("---SMRT NOT submitted to ISDS---")
	call write_variable_in_CASE_NOTE("This SMRT referral could not be sent to ISDS")
    call write_bullet_and_variable_in_CASE_NOTE("Reason referral not submitted", isds_referral_reject_reason)
	call write_variable_in_CASE_NOTE("* SMRT Referral Request returned to originating worker.")
   	Call write_bullet_and_variable_in_CASE_NOTE("Other SMRT notes", other_notes)
	If smrt_request_info_changed = True Then call write_variable_in_CASE_NOTE("Referral information has been updated.")
	call write_variable_in_CASE_NOTE("Referral details:")
	If memb_number = "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
	If memb_number <> "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", "MEMB " & memb_number & " - " & SMRT_member_name)
	Call write_bullet_and_variable_in_CASE_NOTE("Age", memb_age)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT referral requested on", referral_request_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason for referral", referral_reason)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT requested start date", SMRT_start_date)
    Call write_variable_in_CASE_NOTE ("---")
    call write_variable_in_CASE_NOTE(worker_signature)

	end_msg = end_msg & vbCr & vbCr & "Request for SMRT disability determination was not submitted to ISDS."
	end_msg = end_msg & vbCr & "Reason: " & isds_referral_reject_reason
End If

'Once the decision is received, this functionality will detail the information.
If SMRT_actions = "Determination received" then

	If memb_number <> "" Then			'if the MEMBER Number is known, we will try to read DISA
		call navigate_to_MAXIS_screen("STAT", "DISA")
		call write_value_and_transmit(memb_number, 20, 76)

		EMReadScreen disa_version, 1, 2, 73
		If disa_version = "1" Then						'if a DISA version is found, we will pull information from the panel
			EMReadScreen disa_begin_date, 10, 6, 47
			EMReadScreen disa_end_date, 10, 6, 69
			EMReadScreen cert_begin_date, 10, 7, 47
			EMReadScreen cert_end_date, 10, 7, 69

			EMReadScreen cash_disa_status, 2, 11, 59
			EMReadScreen cash_disa_verif, 1, 11, 69
			EMReadScreen snap_disa_status, 2, 12, 59
			EMReadScreen snap_disa_verif, 1, 12, 69
			EMReadScreen hc_disa_status, 2, 13, 59
			EMReadScreen hc_disa_verif, 1, 13, 69
			If cert_begin_date = "__ __ ____" Then cert_begin_date = ""
			If cert_end_date = "__ __ ____" Then cert_end_date = ""

			If hc_disa_verif = "2" Then					'if HC verification is SMRT - we will default the determination to Approved and the certification dates
				SMRT_determination = "Approved"
				SMRT_cert_start_date = 	replace(cert_begin_date, " ", "/")
				SMRT_cert_end_date = replace(cert_end_date, " ", "/")
			End If
		End If
	End If

	'this dialog will detail the decision information
	Do
    	Do
    		err_msg = ""
			Dialog1 = "" 'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 446, 215, "Initial SMRT referral dialog"
				DropListBox 80, 105, 75, 45, "Select one..."+chr(9)+"Approved"+chr(9)+"Denied", SMRT_determination
				EditBox 105, 125, 50, 15, SMRT_cert_start_date
				EditBox 105, 145, 50, 15, SMRT_cert_end_date
				EditBox 5, 175, 435, 15, other_notes
				ButtonGroup ButtonPressed
					PushButton 305, 10, 135, 15, "The Request Information is Incorrect", change_details_btn
					OkButton 335, 195, 50, 15
					CancelButton 390, 195, 50, 15
				Text 5, 10, 195, 10, "SMRT requested for: " & SMRT_member
				Text 5, 20, 175, 10, "Date SMRT referral requested: " & referral_request_date
				Text 5, 30, 140, 10, "SMRT requested start date: " & SMRT_start_date
				Text 5, 40, 155, 10, "SMRT Submitted to ISDS date: " & isds_referral_date
				Text 5, 55, 70, 10, "Reason for referral:"
				Text 5, 65, 435, 10, referral_reason
				If referred_exp = "Yes" Then
					Text 5, 80, 135, 10, "Expedited Referral Requested Reason:"
					Text 5, 90, 435, 10, expedited_reason
				Else
					Text 5, 80, 135, 10, "Expedited Referral was NOT Requested."
				End If
				Text 5, 110, 75, 10, "SMRT Determination: "
				If disa_version = "1" Then Text 165, 110, 70, 10, "DISA Panel Exists"
				If disa_version = "1" and hc_disa_verif <> "2" Then Text 235, 110, 120, 20, "DISA exists but was not listed with SMRT verification for Health Care."
				Text 5, 130, 100, 10, "SMRT Certification Start Date:"
				Text 5, 150, 95, 10, "SMRT Certification End Date:"
				Text 5, 165, 80, 10, "Additional SMRT Notes"
			EndDialog

    		Dialog Dialog1
    		cancel_without_confirmation

			If ButtonPressed = change_details_btn Then		'this runs the dialogs to change the referral information using the dialogs in function
				Call gather_SMRT_request_info()
				err_msg = "LOOP"
			End If
    		If SMRT_determination = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Select the determination status."
    		If SMRT_determination = "Approved" Then
				If IsDate(SMRT_cert_start_date) = False Then err_msg = err_msg & vbNewLine & "* Enter the SMRT certification start date as a valid date."
			End If
			IF err_msg <> "" and left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
   		Call check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = False

	'identifying if the request information has been changed for details in CASE/NOTE
	If initial_SMRT_member <> SMRT_member Then smrt_request_info_changed = True
	If initial_memb_number <> memb_number Then smrt_request_info_changed = True
	If initial_SMRT_member_name <> SMRT_member_name Then smrt_request_info_changed = True
	If initial_SMRT_start_date <> SMRT_start_date Then smrt_request_info_changed = True
	If initial_referral_date <> referral_request_date Then smrt_request_info_changed = True
	If initial_referred_exp <> referred_exp Then smrt_request_info_changed = True
	If initial_expedited_reason <> expedited_reason Then smrt_request_info_changed = True
	If initial_referral_reason <> referral_reason Then smrt_request_info_changed = True

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("---SMRT determination received: " & SMRT_determination & "---")
	Call write_bullet_and_variable_in_CASE_NOTE("SMRT Certification Start Date", SMRT_cert_start_date)
	Call write_bullet_and_variable_in_CASE_NOTE("SMRT Certification End Date", SMRT_cert_end_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Other SMRT notes", other_notes)
	If smrt_request_info_changed = True Then call write_variable_in_CASE_NOTE("Referral information has been updated.")
	call write_variable_in_CASE_NOTE("Referral details:")
	If memb_number = "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", SMRT_member)
	If memb_number <> "" Then call write_bullet_and_variable_in_CASE_NOTE("SMRT requested for", "MEMB " & memb_number & " - " & SMRT_member_name)
	Call write_bullet_and_variable_in_CASE_NOTE("Age", memb_age)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT referral requested on", referral_request_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason for referral", referral_reason)
    Call write_bullet_and_variable_in_CASE_NOTE("SMRT requested start date", SMRT_start_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Was referral expedited", referred_exp)
	If referred_exp = "Yes" then Call write_bullet_and_variable_in_CASE_NOTE("Expedited reason", expedited_reason)
    Call write_variable_in_CASE_NOTE ("---")
    call write_variable_in_CASE_NOTE(worker_signature)

	'additional information in the end message for the determination functionality to advise on possible next steps
	end_msg = end_msg & vbCr & vbCr & "The State Medical Review Team has completed the review for " & SMRT_member
	If SMRT_determination = "Denied" Then
		end_msg = end_msg & vbCr & "The SMRT Disability Determination has been denied."
	End If
	If SMRT_determination = "Approved" Then
		end_msg = end_msg & vbCr & "SMRT Disability Determination has been approved."
		If disa_version = "1" Then
			end_msg = end_msg & vbCr & vbCr & "The DISA panel exists for this person, review the coding on DISA and ensure all programs are processed with this disability determination."
			If hc_disa_verif <> "2" Then end_msg = end_msg & vbCr & " - Health Care verification was not listed as SMRT and should be reviewed."
		ElseIf disa_version <> "1" Then
			end_msg = end_msg & vbCr & vbCr & "There is no DISA panel for this person."
			end_msg = end_msg & vbCr & "Please update the case to correctly document the disability determination and take proper action on the case."
		End If
	End If
END If

Call script_end_procedure_with_error_report(end_msg)
