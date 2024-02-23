'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - INDIVIDUAL ON DEMAND NOTICES.vbs"
start_time = timer
STATS_counter = 1			 'sets the stats counter at one
STATS_manualtime = 180			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

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
call changelog_update("02/22/2024", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function confirm_memo_created_today(notc_confirm)
'function to check SPEC/MEMO to ensure that the notice was created and is in a waiting status.
	'Requires the current date to be in a string format that matches the display on SPEC/MEMO
	today_mo = DatePart("m", date)
	today_mo = right("00" & today_mo, 2)

	today_day = DatePart("d", date)
	today_day = right("00" & today_day, 2)

	today_yr = DatePart("yyyy", date)
	today_yr = right(today_yr, 2)
	today_date = today_mo & "/" & today_day & "/" & today_yr

	'reading the notices on SPEC/MEMO to find if a notice was created
	memo_row = 7                                            'Setting the row for the loop to read MEMOs
	notc_confirm = FALSE         'Defaulting this to 'N'
	Do
		EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
		EMReadScreen print_status, 7, memo_row, 67
		If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
			notc_confirm = TRUE             'If we've found this then no reason to keep looking.
			successful_notices = successful_notices + 1                 'For statistical purposes
			Exit Do
		End If

		memo_row = memo_row + 1           'Looking at next row'
	Loop Until create_date = "        "
end function

'THE SCRIPT ================================================================================================================
EMConnect ""													'Connects to BlueZone
Call check_for_MAXIS(TRUE)										'Makes sure we are in MAXIS
Call MAXIS_case_number_finder(MAXIS_case_number)				'grabs case number
MAXIS_footer_month = CM_mo										'Setting the footer month
MAXIS_footer_year = CM_yr

'Identifying if a member of the A&I Team is running the script
bz_writer = False
If user_ID_for_validation = "CALO001" Then bz_writer = True
If user_ID_for_validation = "ILFE001" Then bz_writer = True
If user_ID_for_validation = "MEGE001" Then bz_writer = True
If user_ID_for_validation = "MARI001" Then bz_writer = True
If user_ID_for_validation = "DACO003" Then bz_writer = True

'Creating a list of the options for notices
memo_list = "Select One..."									'the ones for applications are available to anyone with QI access to the scripts
memo_list = memo_list+chr(9)+"APPL - Appt Notice"
memo_list = memo_list+chr(9)+"APPL - NOMI"
If bz_writer = True Then
	memo_list = memo_list+chr(9)+"RECERT - APPT Notice"		'only script writers should be able to access the recert notices
	memo_list = memo_list+chr(9)+"RECERT - NOMI"
End If

'Capturing CASE Number and Worker Signature
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 176, 65, "Case number"
  EditBox 70, 5, 50, 15, MAXIS_case_number
  EditBox 70, 25, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 45, 50, 15
    CancelButton 125, 45, 45, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 60, 10, "Worker signature:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
Call check_for_MAXIS(False)									'Makes sure we are in MAXIS

'Here we do some discovery about the case to make an action determination and autofill some detail.

'Finding the program and case status
Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_PRIV)
If is_this_PRIV = True Then Call script_end_procedure("This case appears privileged, the script will now end as you do not have access to this case.")
EMReadScreen county_code, 2, 21, 16
If county_code <> "27" Then Call script_end_procedure("This case is not in Hennepin County and notices cannot be sent, the script will now end.")
Call  determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

application_notice_required = False										'determining if the case is pending a program that requires an application notice
If ga_status = "PENDING" Then application_notice_required = True
If msa_status = "PENDING" Then application_notice_required = True
If mfip_status = "PENDING" Then application_notice_required = True
If dwp_status = "PENDING" Then application_notice_required = True
If grh_status = "PENDING" Then application_notice_required = True
If snap_status = "PENDING" Then application_notice_required = True
If emer_status = "PENDING" Then application_notice_required = True

'The script will end if this is NOT an application notice and it is NOT being run by a BZ Script Writer
If application_notice_required = False and bz_writer = False Then
	end_early_msg = "This script is to send Appointment Notices and NOMIs for pending cases to support the On Demand Notice Requirements."
	end_early_msg = end_early_msg & vbCr & vbCr & "This case does not appear to be pending and so does not require an On Demand Appointment Notice or NOMI."
	end_early_msg = end_early_msg & vbCr & vbCr & MAXIS_case_number & " Case Status: "
	end_early_msg = end_early_msg & vbCr & "Case: " & case_status

	end_early_msg = end_early_msg & vbCr & "Active Programs: " & list_active_programs
	end_early_msg = end_early_msg & vbCr & "Pending Programs: " & list_pending_programs
	end_early_msg = end_early_msg & vbCr & vbCr & "The script will now end."

	Call script_end_procedure_with_error_report(end_early_msg)
End If

'Finding some renewal details if the case is not pending
date_of_app = ""
If application_notice_required = False Then
	memo_to_send = "RECERT - APPT Notice"							'defaulting the notice to RECERT - Appt notice

    MAXIS_footer_month = CM_plus_1_mo       						'need to look at stat for next month to see if app is received.
    MAXIS_footer_year = CM_plus_1_yr

    Call navigate_to_MAXIS_screen("STAT", "REVW")					'go to review
	EMReadScreen cash_revw_code, 1, 7, 40							'read the code for SNAP or CASH review
	EMReadScreen snap_revw_code, 1, 7, 60

	'If either the snap or cash code are not blank, then this month has a REVW and we are dealing with a RECERT - NOMI
	If snap_revw_code <> "_" OR cash_revw_code <> "_" Then
		memo_to_send = "RECERT - NOMI"								'default the notice to RECERT - NOMI

		EmReadscreen caf_recvd_date, 8, 13, 37						'looking to see if the application form was received and entered to STAT/REVW
		caf_recvd_date = replace(caf_recvd_date, " ", "/")
		If caf_recvd_date = "__/__/__" Then
			date_of_app = ""
		Else
			date_of_app = caf_recvd_date
		End If
		date_of_app = date_of_app & ""								'making this a string for display
	End If
End If

'For cases that do appear pending, we will check for application information
If application_notice_required = True Then
	Call autofill_editbox_from_MAXIS(HH_member_array, "PROG", application_date)			'finding the application date from PROG - this will pull the most recent application date

	'goind to CASE/NOTE to read if an appointment notice was sent.
	Call navigate_to_MAXIS_screen("CASE", "NOTE")
	note_row = 5            															'setting the variables on the loop
	day_before_app = DateAdd("d", -1, application_date)									'We only need to look back to the date of application
	appt_date = ""																		'blanking the appointment date variable.

	Do
		EMReadScreen note_date, 8, note_row, 6      									'reading the note date
		EMReadScreen note_title, 55, note_row, 25   									'reading the note header
		note_title = trim(note_title)

		IF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then			'This is the the header for the appointment notice case note
			EMReadScreen appt_date, 10, note_row, 63									'Reading the appointment date and formatting
			appt_date = replace(appt_date, "~", "")
			appt_date = trim(appt_date)
			Exit Do
		END IF

		'This is how we move through the notes and leave when we are done
		IF note_date = "        " then Exit Do
		note_row = note_row + 1
		IF note_row = 19 THEN
			PF8
			note_row = 5
		END IF
		EMReadScreen next_note_date, 8, note_row, 6
		IF next_note_date = "        " then Exit Do
	Loop until datevalue(next_note_date) < day_before_app 								'looking ahead at the next case note kicking out the dates before app'
End If

If application_notice_required = True Then												'defaulting the notice to send based on found information
	memo_to_send = "APPL - Appt Notice"
	If appt_date <> "" Then memo_to_send = "APPL - NOMI"
End If

'Dialog to select which notice to send
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 246, 55, "Select the Notice to Send"
  DropListBox 60, 10, 100, 45, memo_list, memo_to_send
  DropListBox 60, 30, 100, 45, "English", select_language				'Note that this doesn't do anything right now BUT can be used when we are ready to add language supports
  ButtonGroup ButtonPressed
    OkButton 185, 10, 50, 15
    CancelButton 185, 30, 50, 15
  Text 5, 15, 50, 10, "Notice to Send:"
  Text 20, 35, 35, 10, "Language:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation

		If memo_to_send = "Select One..." Then err_msg = err_msg & vbCr & "* Select which type of notice to send."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Here we need to get some information ready since we know which notice is being sent.
If memo_to_send = "RECERT - APPT Notice" Then
	month_plus_one = CM_plus_2_mo & "/01/" & CM_plus_2_yr				'Finding the ER month
	last_day_of_recert = DateAdd("d", -1, month_plus_one)				'Determining the end of the cert period
    interview_end_date = CM_plus_1_mo & "/15/" & CM_plus_1_yr			'Creating a Interview Due / Appointment Date

	last_day_of_recert = last_day_of_recert & ""						'format as string
    interview_end_date = interview_end_date & ""

	if MFIP_case = TRUE then           									'setting the language for the notices - MFIP or SNAP
		if SNAP_case = TRUE then
			intvw_programs = "MFIP/SNAP"
		else
			intvw_programs = "MFIP"
		end if
	else
		intvw_programs = "SNAP"
	end if

	programs = intvw_programs											'Creating a list of all programs
	If GA_case = True Then programs = programs & "/GA"
	If MSA_case = True Then programs = programs & "/MSA"
	If GRH_case = True Then programs = programs & "/GRH"
	If left(programs, 1) = "/" Then programs = right(programs, len(programs)-1)
End If

If memo_to_send = "RECERT - NOMI" Then
	month_plus_one = CM_plus_1_mo & "/01/" & CM_plus_1_yr				'Finding the ER month
	last_day_of_recert = DateAdd("d", -1, month_plus_one)				'Determining the end of the cert period

	last_day_of_recert = last_day_of_recert & ""						'format as string
End If

If memo_to_send = "APPL - Appt Notice" Then
	interview_date = dateadd("d", 5, application_date)					'Using the established logic to create the interview / appointment date
	If interview_date <= date then interview_date = dateadd("d", 5, date)
	Call change_date_to_soonest_working_day(interview_date, "FORWARD")

	application_date = application_date & ""						'format as string
	interview_date = interview_date & ""							'format as string
End If

If memo_to_send = "APPL - NOMI" Then
	application_date = application_date & ""						'format as string
End If

'Dialog to enter dates relevant to the specific notice being sent
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 130, "On Demand Notice Details"
	If memo_to_send = "RECERT - APPT Notice" Then
		EditBox 110, 45, 65, 15, interview_end_date
		EditBox 110, 65, 65, 15, last_day_of_recert
		Text 15, 50, 90, 10, "Date of Interview Deadline:"
		Text 15, 70, 90, 10, "Last Day of Cert Pd:"
		Text 10, 90, 250, 10, "Sending a APPT Notc for the " & CM_plus_2_mo & "/" & CM_plus_2_yr & " ER in " & select_language
	End If
	If memo_to_send = "RECERT - NOMI" Then
		EditBox 110, 45, 65, 15, date_of_app
		EditBox 110, 65, 65, 15, last_day_of_recert
		Text 15, 50, 90, 10, "Date Application Received:"
		Text 15, 70, 90, 10, "Last Day of Cert Pd:"
		Text 10, 90, 250, 10, "Sending a NOMI for the " & CM_plus_1_mo & "/" & CM_plus_1_yr & " ER in " & select_language
		Text 10, 110, 250, 10, "If an application has not been received, leave this field blank."
	End If
	If memo_to_send = "APPL - Appt Notice" Then
		EditBox 80, 45, 65, 15, application_date
		EditBox 80, 65, 65, 15, interview_date
		Text 15, 50, 60, 10, "Application date:"
		Text 15, 70, 60, 10, "Appointment date:"
		Text 10, 90, 250, 10, "Default appointment date aligns with standard notice process."
	End If
	If memo_to_send = "APPL - NOMI" Then
		EditBox 100, 45, 65, 15, application_date
		EditBox 100, 65, 65, 15, appt_date
		Text 15, 50, 80, 10, "Application date:"
		Text 15, 70, 80, 10, "Missed Interview date:"
		Text 10, 90, 250, 10, "Check the CASE/NOTE or SPEC/MEMO to find the 'Missed Interview Date'."
	End If

	ButtonGroup ButtonPressed
		OkButton 155, 110, 50, 15
		CancelButton 210, 110, 50, 15
	Text 10, 10, 225, 10, "Enter/Update the associated dates for the notice being sent."
	Text 10, 25, 150, 10, "Notice to send: " & memo_to_send
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation

		If memo_to_send = "RECERT - APPT Notice" Then
			If IsDate(interview_end_date) = False Then err_msg = err_msg & vbCr & "* Enter the date the interview is due by for the Notice to the resident."
			If IsDate(last_day_of_recert) = False Then err_msg = err_msg & vbCr & "* Enter the last day of the certification period."
		End If
		If memo_to_send = "RECERT - NOMI" Then
			If IsDate(interview_end_date) = False Then err_msg = err_msg & vbCr & "* Enter the date the interview is due by for the Notice to the resident."
			If date_of_app <> "" Then
				If IsDate(date_of_app) = False Then err_msg = err_msg & vbCr & "* Enter the day the recertification form was received in the agency. This can be blank if not received."
			End If
		End If
		If memo_to_send = "APPL - Appt Notice" Then
			If IsDate(application_date) = False Then err_msg = err_msg & vbCr & "* Enter the date of application."
			If IsDate(interview_date) = False Then err_msg = err_msg & vbCr & "* Enter the date the interview is due by for the Notice to the resident."
		End If
		If memo_to_send = "APPL - NOMI" Then
			If IsDate(application_date) = False Then err_msg = err_msg & vbCr & "* Enter the date of application."
			If IsDate(appt_date) = False Then err_msg = err_msg & vbCr & "* Enter the date the interview is due by for the Notice to the resident."
		End If

		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Checking the Appointment date and ending the script if a NOMI is selected BEFORE the Appointment date occurs
If memo_to_send = "RECERT - NOMI" Then
	nomi_date = CM_mo & "/15/" & CM_yr
	nomi_date = DateAdd("d", 0, nomi_date)
	If DateDiff("d", date, nomi_date) > 0 Then Call script_end_procedure_with_error_report("NOMI cannot be send until the Interview deadline has been reached.")
End If
If memo_to_send = "APPL - NOMI" Then
	appt_date = DateAdd("d", 0, appt_date)
	If DateDiff("d", date, appt_date) > 0 Then Call script_end_procedure_with_error_report("It does not appear the interview due date has hapened yet. The NOMI should not be sent until the interview is due. Interview due on " & appt_date & ".")
	appt_date = appt_date & ""
End If

Call check_for_MAXIS(False)									'Makes sure we are in MAXIS
'Open a MEMO
Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)

'Each notice has a function call and CASE/NOTE and will establish the details of the end message display
If memo_to_send = "RECERT - APPT Notice" Then
	end_msg = "Recertification - Appointment Notice" & vbCr & CM_plus_2_mo & "/" & CM_plus_2_yr & " ER in " & select_language & vbCr & "Programs: " & programs & vbCr & "Interview Due: " & interview_end_date & vbCr & "Cert Pd End :" & last_day_of_recert
	CALL create_appointment_letter_notice_recertification(programs, intvw_programs, interview_end_date, last_day_of_recert)
	Call confirm_memo_created_today(notc_confirm)

	If notc_confirm = True Then
		'Appointment Notice Recerts have an additional MEMO with renewal guidance
		Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)
		CALL write_variable_in_SPEC_MEMO("As a part of the Renewal Process we must receive recent verification of your information. To speed the renewal process, please send proofs with your renewal paperwork.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, employer statement,")
		CALL write_variable_in_SPEC_MEMO("   income reports, business ledgers, income tax forms, etc.")
		CALL write_variable_in_SPEC_MEMO("   *If a job has ended, send proof of the end of employment")
		CALL write_variable_in_SPEC_MEMO("   and last pay.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house")
		CALL write_variable_in_SPEC_MEMO("   payment receipt, mortgage, lease, subsidy, etc.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed):")
		CALL write_variable_in_SPEC_MEMO("   prescription and medical bills, etc.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")

		start_a_blank_case_note
		CALL write_variable_in_CASE_NOTE("*** Notice of " & programs & " Recertification Interview Sent ***")
		CALL write_variable_in_case_note("* A notice has been sent to client with detail about how to call in for an interview.")
		CALL write_variable_in_case_note("* Client must submit paperwork and call 612-596-1300 to complete interview.")
		If forms_to_arep = "Y" then call write_variable_in_case_note("* Copy of notice sent to AREP.")
		If forms_to_swkr = "Y" then call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
		call write_variable_in_case_note("---")
		CALL write_variable_in_case_note("Link to Domestic Violence Brochure sent to client in SPEC/MEMO as a part of interview notice.")
		call write_variable_in_case_note("---")
		call write_variable_in_case_note(worker_signature)

	End If
End If

If memo_to_send = "RECERT - NOMI" Then
	end_msg = "Recertification - NOMI" & vbCr & CM_plus_1_mo & "/" & CM_plus_1_yr & " ER in " & select_language & vbCr & "Form Received: " & date_of_app & vbCr & "Cert Pd End :" & last_day_of_recert
	CALL create_NOMI_recertification(date_of_app, last_day_of_recert)

	If date_of_app = "" Then recvd_appl = False
	If date_of_app <> "" Then recvd_appl = True
	Call confirm_memo_created_today(notc_confirm)
	If notc_confirm = True Then
		start_a_blank_case_note
        CALL write_variable_in_CASE_NOTE("*** NOMI Sent for SNAP Recertification***")
        if recvd_appl = TRUE then CALL write_variable_in_CASE_NOTE("* Recertification app received on " & date_of_app)
        if recvd_appl = FALSE then CALL write_variable_in_CASE_NOTE("* Recertification app has NOT been received. Client must submit paperwork.")
        CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about how to call in for an interview.")
        CALL write_variable_in_CASE_NOTE("* Client must call 612-596-1300 to complete interview.")
        If forms_to_arep = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")
        If forms_to_swkr = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")
        call write_variable_in_case_note("---")
        call write_variable_in_case_note(worker_signature)
	End If
End If

If memo_to_send = "APPL - Appt Notice" Then
	last_contact_day = dateadd("d", 30, application_date)
	If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date

	end_msg = "Application - Appointment Notice" & vbCr & "Application Date: " & application_date & vbCr & "Interview Due: " & interview_date & vbCr & "Last Contact Day :" & last_contact_day
	Call create_appointment_letter_notice_application(application_date, interview_date, last_contact_day)

	Call confirm_memo_created_today(notc_confirm)
	MsgBox "notc_confirm - " & notc_confirm
	If notc_confirm = True Then
		start_a_blank_case_note
		Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & interview_date & " ~")
		Call write_variable_in_CASE_NOTE("A notice has been sent via SPEC/MEMO informing the client of needed interview.")
		Call write_variable_in_CASE_NOTE("Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
		Call write_variable_in_CASE_NOTE("A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE(worker_signature)
	End If
End If

If memo_to_send = "APPL - NOMI" Then
	last_contact_day = dateadd("d", 30, application_date)
	If DateDiff("d", appt_date, last_contact_day) < 1 then last_contact_day = appt_date

	end_msg = "Application - NOMI" & vbCr & "Application Date: " & application_date & vbCr & "Missed Interview Date: " & appt_date & vbCr & "Last Contact Day :" & last_contact_day
	CALL create_NOMI_application(application_date, appt_date, last_contact_day)

	Call confirm_memo_created_today(notc_confirm)
	If notc_confirm = True Then
		start_a_blank_case_note
		Call write_variable_in_CASE_NOTE("~ Client has not completed application interview, NOMI sent ~ ")
		Call write_variable_in_CASE_NOTE("A notice was previously sent to client with detail about completing an interview. ")
		Call write_variable_in_CASE_NOTE("Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
		Call write_variable_in_CASE_NOTE("A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE(worker_signature)
	End If
End If

'This will alert the worker if the Notice appears to have failed - additionally a CASE/NOTE will not be created.
If notc_confirm = False Then end_msg = "NOTICE HAS FAILED" & vbCr & vbCr & end_msg

Call script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------02/23/2024
'--Tab orders reviewed & confirmed----------------------------------------------02/23/2024
'--Mandatory fields all present & Reviewed--------------------------------------02/23/2024
'--All variables in dialog match mandatory fields-------------------------------02/23/2024
'Review dialog names for content and content fit in dialog----------------------02/23/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------02/23/2024					The CASE/NOTEs align with BULK script to send these notices
'--CASE:NOTE Header doesn't look funky------------------------------------------02/23/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------02/23/2024
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------02/23/2024
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------02/23/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------02/23/2024
'--PRIV Case handling reviewed -------------------------------------------------02/23/2024
'--Out-of-County handling reviewed----------------------------------------------02/23/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------02/23/2024
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------02/23/2024
'--Incrementors reviewed (if necessary)-----------------------------------------02/23/2024
'--Denomination reviewed -------------------------------------------------------02/23/2024
'--Script name reviewed---------------------------------------------------------02/23/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------02/23/2024
'--comment Code-----------------------------------------------------------------02/23/2024
'--Update Changelog for release/update------------------------------------------02/23/2024
'--Remove testing message boxes-------------------------------------------------02/23/2024
'--Remove testing code/unnecessary code-----------------------------------------02/23/2024
'--Review/update SharePoint instructions----------------------------------------02/23/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------02/23/2024					Checked QI On Demand Instructions and there is no specific script reference
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------To Do
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------To Do
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
