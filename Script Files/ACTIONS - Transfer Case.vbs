'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("Q:\Blue Zone Scripts\Public assistance script files\Script files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'----------DIALOGS----------
BeginDialog xfer_menu_dialog, 0, 0, 156, 80, "Case XFER"
  OptionGroup XFERRadioGroup
    RadioButton 30, 25, 115, 10, "Within the County/Agency", within_agency_check
    RadioButton 30, 40, 120, 10, "Out-of-County/Agency", out_of_agency_check
  ButtonGroup ButtonPressed
    OkButton 30, 60, 50, 15
    CancelButton 80, 60, 50, 15
  Text 10, 10, 125, 10, "Please select transfer type..."
EndDialog

BeginDialog out_of_county_dlg, 0, 0, 206, 270, "Case Transfer"
  EditBox 60, 5, 50, 15, case_number
  EditBox 125, 30, 65, 15, transfer_to
  CheckBox 5, 55, 190, 10, "Check here if the client is active on HC through MNSure", mnsure_active_check
  CheckBox 5, 70, 190, 10, "Check here if the client is active on Minnesota Care", mcre_active_check
  CheckBox 5, 85, 200, 10, "Check here if the client has a pending MNSure application", mnsure_pend_check
  CheckBox 5, 100, 200, 10, "Check here if you have sent the DHS 3195 transfer form", transfer_form_check
  CheckBox 5, 115, 175, 10, "Check here if you sent the Change Report Form", crf_sent_check
  CheckBox 5, 130, 195, 10, "Check here to add a programs closure date to the MEMO", closure_date_check
  EditBox 80, 145, 45, 15, cl_move_date
  DropListBox 65, 165, 40, 15, "No"+chr(9)+"Yes", excluded_time
  EditBox 155, 165, 45, 15, excl_date
  EditBox 75, 185, 120, 15, Transfer_reason
  EditBox 70, 205, 125, 15, Action_to_be_taken
  EditBox 70, 225, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 50, 250, 50, 15
    CancelButton 100, 250, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 110, 20, "Worker number/Agency you're transferring to  (x###### format)"
  Text 5, 150, 60, 10, "Client Move Date"
  Text 5, 170, 55, 10, "Excluded time?"
  Text 115, 170, 40, 10, "Begin Date"
  Text 5, 190, 70, 10, "Reason for Transfer:"
  Text 5, 210, 65, 10, "Actions to be taken:"
  Text 5, 230, 60, 10, "Worker Signature:"
EndDialog

BeginDialog within_county_dlg, 0, 0, 216, 240, "Case Transfer"
  EditBox 65, 10, 50, 15, case_number
  DropListBox 60, 30, 70, 15, "Adult"+chr(9)+"Family", unit_drop_down
  EditBox 130, 50, 65, 15, worker_to_transfer_to
  CheckBox 20, 90, 30, 10, "Cash", cash_active_check
  CheckBox 55, 90, 30, 10, "SNAP", SNAP_active_check
  CheckBox 95, 90, 20, 10, "HC", HC_active_check
  CheckBox 125, 90, 80, 10, "MCRE Or Mnsure", Mcremnsure_active_check
  CheckBox 20, 125, 30, 10, "Cash", Cash_pend_check
  CheckBox 55, 125, 30, 10, "SNAP", SNAP_pend_check
  CheckBox 95, 125, 20, 10, "HC", HC_pend_check
  CheckBox 125, 125, 75, 10, "MCRE Or Mnsure", Mcremnsure_pend_check
  DropListBox 100, 140, 65, 10, ""+chr(9)+"Yes"+chr(9)+"No", preg_y_n
  EditBox 85, 160, 120, 15, Transfer_reason
  EditBox 80, 180, 125, 15, Action_to_be_taken
  EditBox 80, 200, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 100, 220, 50, 15
    CancelButton 155, 220, 50, 15
  Text 15, 185, 65, 10, "Actions to be taken:"
  Text 15, 35, 40, 10, "Transfer to: "
  Text 15, 165, 70, 10, "Reason for Transfer:"
  Text 15, 145, 85, 10, "Pregnancy verif received:"
  Text 15, 15, 50, 10, "Case Number:"
  Text 15, 75, 80, 10, "Active On:"
  Text 15, 110, 60, 10, "Pending On:"
  Text 15, 205, 60, 10, "Worker Signature:"
  Text 15, 50, 105, 20, "Worker number you're transferring to  (x102XXX format)"
EndDialog

'----------THE SCRIPT----------
EMConnect ""

maxis_check_function

DIALOG xfer_menu_dialog
	IF ButtonPressed = 0 THEN stopscript

	IF XFERRadioGroup = 0 THEN
			'Finds the case number
		call find_variable("Case Nbr: ", case_number, 8)
		case_number = trim(case_number)
		case_number = replace(case_number, "_", "")
		If IsNumeric(case_number) = False then case_number = ""

		'Displays the dialog and navigates to case note
		Do
		  Do
			Do
			  Dialog within_county_dlg
			  If buttonpressed = 0 then stopscript
			  If case_number = "" then MsgBox "You must have a case number to continue!"
			IF len(worker_to_transfer_to) <> 7 then Msgbox "Please include X102 in the worker number"
			Loop until case_number <> "" and len(worker_to_transfer_to) = 7
			transmit
			EMReadScreen MAXIS_check, 5, 1, 39
			If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
		  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
		  call navigate_to_screen("case", "note")
		  PF9
		  EMReadScreen mode_check, 7, 20, 3
		  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
		Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

		'Cleaning up for case note
		  IF SNAP_active_check = 1 THEN active_programs = active_programs & "SNAP/"
		  IF hc_active_check = 1 THEN active_programs = active_programs & "HC/"
		  IF cash_active_check = 1 THEN active_programs = active_programs & "CASH/"
		  IF Mcremnsure_active_check = 1 THEN active_programs = active_programs & "Mcre,Mnsure/"

		  IF SNAP_pend_check = 1 THEN pend_programs = pend_programs & "SNAP/"
		  IF hc_pend_check = 1 THEN pend_programs = pend_programs & "HC/"
		  IF cash_pend_check = 1 THEN pend_programs = pend_programs & "CASH/"
		  IF Mcremnsure_pend_check = 1 THEN pend_programs = pend_programs & "Mcre,Mnsure/"

		'Case notes
		EMSendKey "***Transfer within county***" & "<newline>"
		call write_new_line_in_case_note("* Transfer to: " & unit_drop_down) 
		IF active_programs <> "" THEN
		  EMSendKey "* Active Programs: " & active_programs & "<backspace>"
		  EMSendKey "<newline>"
		END IF
		IF pend_programs <> "" THEN 
		  EMSendKey "* Pending Programs: " & pend_programs & "<backspace>"
		  EMSendKey "<newline>"
		END IF
		IF preg_y_n <> "" THEN call write_new_line_in_case_note("* Pregnancy verification rec'd: " & preg_y_n)
		call write_editbox_in_case_note("Reason for transfer", Transfer_reason, 6) 
		call write_editbox_in_case_note("Actions to be taken", Action_to_be_taken, 6) 
		call write_new_line_in_case_note("---")
		call write_new_line_in_case_note(worker_signature)

		'Transfers case
		back_to_self
		EMWriteScreen "spec", 16, 43
		EMWriteScreen "________", 18, 43
		EMWriteScreen case_number, 18, 43
		EMWriteScreen "xfer", 21, 70
		transmit
		EMWriteScreen "x", 7, 16
		transmit
		PF9
		EMWriteScreen worker_to_transfer_to, 18, 61
		transmit

		script_end_procedure("")
	
	ELSEIF XFERRadioGroup = 1 THEN

	DO
		DO
			DO
				DIALOG out_of_county_dlg
					IF ButtonPressed = 0 THEN stopscript

				'----------Goes to STAT/PROG to pull active/pending case information----------
				call navigate_to_screen("STAT", "PROG")
					EMReadScreen cash_one_status, 4, 6, 74
					EMReadScreen cash_two_status, 4, 7, 74
					EMReadScreen snap_status, 4, 10, 74
					EMReadScreen hc_status, 4, 12, 74

					IF excluded_time = "Yes" AND isdate(excl_date) = FALSE THEN MsgBox "Please enter a valid date for the start of excluded time or double check that the client's absense is due to excluded time."
					IF isdate(cl_move_date) = FALSE THEN MsgBox "Please enter a valid date for client move."
					IF ucase(left(transfer_to, 4)) = ucase(worker_county_code) THEN MsgBox "You must use the ''Within the Agency'' script to transfer the case within the agency. The Worker/Agency you have selected indicates you are trying to transfer within your agency."
					IF (hc_status = "ACTV" AND excluded_time = "") THEN MsgBox "Please select whether the client is on excluded time."
					IF len(transfer_to) <> 7 THEN MsgBox "Please select a valid worker or agency to receive the case (proper format = x######)."	

			LOOP UNTIL ((HC_status <> "ACTV") OR (hc_status = "ACTV" AND excluded_time = "No") OR (HC_status = "ACTV" AND excluded_time = "Yes" AND isdate(excl_date) = TRUE)) AND _
			(isdate(cl_move_date) = TRUE) AND _
			(len(transfer_to) = 7) AND _
			(ucase(left(transfer_to, 4)) <> ucase(worker_county_code))

			'----------Checks that the worker or agency is valid----------
			call navigate_to_screen("REPT", "USER")
			EMWriteScreen transfer_to, 21, 12
			transmit
			EMReadScreen worker_found, 15, 24, 2
			EMReadScreen inactive_worker, 8, 7, 38
				IF inactive_worker = "INACTIVE" THEN MsgBox "The worker or agency selected is not active. Please try again."
				IF worker_found = "NO WORKER FOUND" THEN MsgBox "The worker or agency selected does not exist. Please try again."
		LOOP UNTIL inactive_worker <> "INACTIVE" AND worker_found <> "NO WORKER FOUND"

		IF closure_date_check = checked THEN
			closure_date = dateadd("m", 1, date)
			closure_date = datepart("M", closure_date) & "/01/" & datepart("YYYY", closure_date)
			IF datediff("d", date, closure_date) < 10 THEN 
				closure_date = dateadd("m", 1, closure_date)
			END IF
			closure_date = dateadd("d", -1, closure_date)
		END IF


		EMWriteScreen "X", 7, 3
		transmit

		EMReadScreen worker_agency_name, 43, 8, 27
			worker_agency_name = trim(worker_agency_name)
		EMReadScreen mail_addr_line_one, 43, 9, 27
			mail_addr_line_one = trim(mail_addr_line_one)
		EMReadScreen mail_addr_line_two, 43, 10, 27
			mail_addr_line_two = trim(mail_addr_line_two)
		EMReadScreen mail_addr_line_three, 43, 11, 27
			mail_addr_line_three = trim(mail_addr_line_three)
		EMReadScreen mail_addr_line_four, 43, 12, 27
			mail_addr_line_four = trim(mail_addr_line_four)
		EMReadScreen worker_agency_phone, 14, 13, 27
		

		back_to_SELF

		'----------This converts the dates into a format that can be entered into SPEC/XFER/XCTY----------
		cl_move_dt = cdate(cl_move_date)
		cl_move_dt = replace(cl_move_dt, "/", "")
			cl_move_month = left(cl_move_dt, 2)
			cl_move_day = right(left(cl_move_dt, 4), 2)
			cl_move_year = right(cl_move_dt, 2)

		IF crf_sent_check = checked THEN
			crf_sent_month = datepart("M", date)
				IF len(crf_sent_month) = 1 THEN crf_sent_month = "0" & crf_sent_month
			crf_sent_day = datepart("D", date)
				IF len(crf_sent_day) = 1 THEN crf_sent_day = "0" & crf_sent_day
			crf_sent_year = right(datepart("YYYY", date), 2)
		END IF

		IF excluded_time = "Yes" THEN
			excl_dt = cdate(excl_date)
			excl_dt = replace(excl_dt, "/", "")
			excl_month = left(excl_dt, 2)
			excl_day = right(left(excl_dt, 4), 2)
			excl_year = right(excl_dt, 2)
		END IF

		call navigate_to_screen("CASE", "NOTE")
		PF9
		EMReadScreen write_access, 9, 24, 12
		IF write_access = "READ ONLY" THEN MsgBox("You do not have access to modify this case. Please double check your case number and try again." & chr(13) & chr(13) & "Alternatively, you may be in INQUIRY MODE.")
	LOOP UNTIL write_access <> "READ ONLY"

	PF3

	'----------Sending the CL a SPEC/MEMO notifying them of the details of the transfer----------
	call navigate_to_screen("SPEC", "MEMO")
	PF5
	EMWriteScreen "X", 5, 10
	transmit

	EMWriteScreen "Your case has been transferred. Your new worker/agency is:", 3, 15
	memo_line = 5
	IF worker_agency_name <> "" THEN
		EMWriteScreen ("  " & worker_agency_name), memo_line, 15
		memo_line = memo_line + 1
	END IF
	IF mail_addr_line_one <> "" THEN
		EMWriteScreen ("  " & mail_addr_line_one), memo_line, 15
		memo_line = memo_line + 1
	END IF
	IF mail_addr_line_two <> "" THEN
		EMWriteScreen ("  " & mail_addr_line_two), memo_line, 15
		memo_line = memo_line + 1
	END IF
	IF mail_addr_line_three <> "" THEN
		EMWriteScreen ("  " & mail_addr_line_three), memo_line, 15
		memo_line = memo_line + 1
	END IF
	IF mail_addr_line_four <> "" THEN
		EMWriteScreen ("  " & mail_addr_line_four), memo_line, 15
		memo_line = memo_line + 1
	END IF
	EMWriteScreen ("  " & worker_agency_phone), memo_line, 15
	memo_line = memo_line + 2
	EMWriteScreen "If you have any questions, or to send in requested proofs,", memo_line, 15
	memo_line = memo_line + 1
	EMWriteScreen "please direct all communications to the worker or agency", memo_line, 15
	memo_line = memo_line + 1
	EMWriteScreen "listed above.", memo_line, 15
	IF closure_date_check = checked THEN
		memo_line = memo_line + 2
		EMWriteScreen "If you fail to provide required proofs to your new worker", memo_line, 15
		memo_line = memo_line + 1
		EMWriteScreen ("or agency then your benefits will close on " & closure_date & "."), memo_line, 15
	END IF

	PF4

	'----------The case note of the reason for the XFER----------
		call navigate_to_screen("CASE", "NOTE")
		PF9
		IF SNAP_status = "ACTV" THEN active_programs = active_programs & "SNAP/"
		IF hc_status = "ACTV" THEN active_programs = active_programs & "HC/"
		IF cash_one_status = "ACTV" OR cash_two_status = "ACTV" THEN active_programs = active_programs & "CASH/"

		IF SNAP_status = "PEND" THEN pend_programs = pend_programs & "SNAP/"
		IF hc_status = "PEND" THEN pend_programs = pend_programs & "HC/"
		IF cash_pend_check = checked THEN pend_programs = pend_programs & "CASH/"

	call write_new_line_in_case_note("**CASE TRANSFER TO " & ucase(transfer_to) & "**")
	IF active_programs <> "" THEN call write_editbox_in_case_note("Active Programs", (left(active_programs, (len(active_programs) - 1))), 6)
	IF mnsure_active_check = checked THEN call write_new_line_in_case_note("* CL is active on HC through MNSure")
	IF mcre_active_check = checked THEN call write_new_line_in_case_note("* CL is active on HC through Minnesota Care")
	IF pend_programs <> "" THEN call write_editbox_in_case_note("Pending Programs", (left(active_programs, (len(active_programs) - 1))), 6)
	IF mnsure_pend_check = checked THEN call write_new_line_in_case_note("* CL has pending HC application through MNSure")
	call write_editbox_in_case_note("CL Move Date", cl_move_date, 6)
	IF crf_sent_check = checked THEN call write_editbox_in_case_note("Change Report Sent", crf_sent_date, 6)
	IF excluded_time = "Yes" THEN
		excluded_time = excluded_time & ", Begins " & excl_date
		call write_editbox_in_case_note("Excluded Time" , excluded_time, 6)
		excluded_time = "Yes"
	ELSEIF excluded_time = "No" THEN
		call write_editbox_in_case_note("Excluded Time", excluded_time, 6)
	ELSEIF excluded_time = "" THEN
		call write_new_line_in_case_note("* Excluded Time: N/A")
	END IF
	IF excluded_time = "No" THEN 
		cfr_change_date = dateadd("M", 3, cl_move_date)
		cfr_change_month = datepart("M", cfr_change_date)
		IF len(cfr_change_month) = 1 THEN cfr_change_month = "0" & cfr_change_month
		cfr_change_year = datepart("YYYY", cfr_change_date)
		IF hc_active_check = 1 OR cash_active_check = 1 THEN call write_editbox_in_case_note("Cty of Financial Responsibility Changes", (cfr_change_month & "/" & cfr_change_year), 6)
	END IF
	IF closure_date_check = checked THEN call write_new_line_in_case_note("* CL has until " & closure_date & " to provide required proofs or the case will close.")
	call write_editbox_in_case_note("Reason for Transfer", Transfer_reason, 6)
	call write_editbox_in_case_note("Actions to be Taken", Action_to_be_taken, 6)
	If transfer_form_check = checked THEN call write_new_line_in_case_note("DHS 3195 Inter Agency Case Transfer Form completed and sent.")
	call write_new_line_in_case_note("---")
	call write_new_line_in_case_note(worker_signature)
	PF3	


	'----------The business end of the script (DON'T POINT THIS SCRIPT AT ANYTHING YOU DON'T WANT TRANSFERRED!!)----------
	call navigate_to_screen("SPEC", "XFER")
	EMWriteScreen "X", 9, 16
	transmit
	PF9

	EMWriteScreen cl_move_month, 4, 28
	EMWriteScreen cl_move_day, 4, 31
	EMWriteScreen cl_move_year, 4, 34
	IF crf_sent_check = checked THEN
		EMWriteScreen crf_sent_month, 4, 61
		EMWriteScreen crf_sent_day, 4, 64
		EMWriteScreen crf_sent_year, 4, 67
	END IF
	EMWriteScreen left(excluded_time, 1), 5, 28
	IF excluded_time = "Yes" THEN
		EMWriteScreen excl_month, 6, 28
		EMWriteScreen excl_day, 6, 31
		EMWriteScreen excl_year, 6, 34
		EMWriteScreen right(worker_county_code, 2), 15, 39
	ELSEIF excluded_time = "" THEN
		EMWriteScreen "N", 5, 28
	END IF

	IF hc_status = "ACTV" THEN
		EMWriteScreen right(worker_county_code, 2), 14, 39
		EMWriteScreen cfr_change_month, 14, 53
		EMWriteScreen cfr_change_year, 14, 59
	END IF

	IF cash_one_status = "ACTV" THEN
		EMWriteScreen right(worker_county_code, 2), 11, 39
		EMWriteScreen cfr_change_month, 11, 53
		EMWriteScreen right(cfr_change_year, 2), 11, 59
	END IF

	IF cash_two_status = "ACTV" THEN
		EMWriteScreen right(worker_county_code, 2), 12, 39
		EMWriteScreen cfr_change_month, 12, 53
		EMWriteScreen right(cfr_change_year, 2), 12, 59
	END IF

	EMReadScreen primary_worker, 7, 21, 16
	EMWriteScreen primary_worker, 18, 28
	EMWriteScreen transfer_to, 18, 61

	transmit
	IF crf_sent_check = unchecked THEN transmit

	script_end_procedure("")
	
	END IF