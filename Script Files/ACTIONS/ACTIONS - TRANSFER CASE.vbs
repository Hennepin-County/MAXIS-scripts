'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - TRANSFER CASE.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'VARIABLES TO DECLARE----------------------------------------------------------------------------
SPEC_MEMO_check = checked		'Should default to checked, as we usually want to send a new worker memo

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

BeginDialog out_of_county_dlg, 0, 0, 236, 385, "Case Transfer"
  EditBox 60, 5, 50, 15, case_number
  EditBox 115, 30, 65, 15, transfer_to
  CheckBox 5, 55, 190, 10, "Check here if the client is active on HC through MNSure.", mnsure_active_check
  CheckBox 5, 70, 190, 10, "Check here if the client is active on Minnesota Care.", mcre_active_check
  CheckBox 5, 85, 200, 10, "Check here if the client has a pending MNSure application.", mnsure_pend_check
  CheckBox 5, 100, 200, 10, "Check here if you have sent the DHS 3195 transfer form.", transfer_form_check
  CheckBox 5, 115, 175, 10, "Check here if you sent the Change Report Form.", crf_sent_check
  CheckBox 5, 130, 230, 10, "Check here to send a SPEC/MEMO to the client with new worker info.", SPEC_MEMO_check
  CheckBox 5, 145, 195, 10, "Check here to add a programs closure date to the MEMO.", closure_date_check
  EditBox 80, 160, 45, 15, cl_move_date
  DropListBox 65, 180, 40, 15, "No"+chr(9)+"Yes", excluded_time
  EditBox 155, 180, 45, 15, excl_date
  CheckBox 5, 200, 220, 10, "Check here to manually set the CFR and change date for CASH.", manual_cfr_cash_check
  CheckBox 5, 215, 215, 10, "Check here if CASH CFR is not changing.", cash_cfr_no_change_check
  EditBox 60, 230, 20, 15, cash_cfr
  EditBox 165, 230, 20, 15, cash_cfr_month
  EditBox 190, 230, 20, 15, cash_cfr_year
  CheckBox 5, 250, 220, 10, "Check here to manually set the CFR and change date for HC.", manual_cfr_hc_check
  CheckBox 5, 265, 225, 10, "Check here if the HC CFR is not changing.", hc_cfr_no_change_check
  EditBox 60, 280, 20, 15, hc_cfr
  EditBox 165, 280, 20, 15, hc_cfr_month
  EditBox 190, 280, 20, 15, hc_cfr_year
  EditBox 75, 300, 155, 15, Transfer_reason
  EditBox 70, 320, 160, 15, Action_to_be_taken
  EditBox 70, 340, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 360, 50, 15
    CancelButton 120, 360, 50, 15
    PushButton 155, 5, 70, 15, "NAV to SPEC/XFER", nav_to_xfer_button
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 110, 20, "Worker number/Agency you're transferring to  (x###### format)"
  Text 5, 165, 60, 10, "Client Move Date"
  Text 5, 185, 55, 10, "Excluded time?"
  Text 115, 185, 40, 10, "Begin Date"
  Text 10, 235, 45, 10, "Current CFR:"
  Text 85, 235, 75, 10, "Change Date (MM YY)"
  Text 10, 285, 45, 10, "Current CFR:"
  Text 85, 285, 75, 10, "Change Date (MM YY)"
  Text 5, 305, 70, 10, "Reason for Transfer:"
  Text 5, 325, 65, 10, "Actions to be taken:"
  Text 5, 345, 60, 10, "Worker Signature:"
EndDialog


BeginDialog within_county_dlg, 0, 0, 216, 240, "Case Transfer"
  EditBox 65, 10, 50, 15, case_number
  ComboBox 80, 30, 75, 15, "Select one..."+chr(9)+"N/A"+chr(9)+"Adult"+chr(9)+"Family"+chr(9)+"Cash"+chr(9)+"GRH"+chr(9)+"LTC", unit_drop_down
  EditBox 130, 50, 65, 15, worker_to_transfer_to
  CheckBox 20, 90, 30, 10, "Cash", cash_active_check
  CheckBox 55, 90, 30, 10, "SNAP", SNAP_active_check
  CheckBox 95, 90, 20, 10, "HC", HC_active_check
  CheckBox 125, 90, 35, 10, "MNsure", mnsure_active_check
  CheckBox 170, 90, 35, 10, "EMER", EMER_active_check
  CheckBox 20, 125, 30, 10, "Cash", Cash_pend_check
  CheckBox 55, 125, 30, 10, "SNAP", SNAP_pend_check
  CheckBox 95, 125, 20, 10, "HC", HC_pend_check
  CheckBox 125, 125, 35, 10, "MNsure", mnsure_pend_check
  CheckBox 170, 125, 40, 10, "EMER", EMER_pend_check
  DropListBox 100, 140, 65, 10, "Select one..."+chr(9)+"N/A"+chr(9)+"Yes"+chr(9)+"No", preg_y_n
  EditBox 85, 160, 120, 15, Transfer_reason
  EditBox 80, 180, 125, 15, Action_to_be_taken
  EditBox 80, 200, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 100, 220, 50, 15
    CancelButton 155, 220, 50, 15
  Text 15, 185, 65, 10, "Actions to be taken:"
  Text 15, 35, 60, 10, "Unit to transfer to: "
  Text 15, 165, 70, 10, "Reason for Transfer:"
  Text 15, 145, 85, 10, "Pregnancy verif received:"
  Text 15, 15, 50, 10, "Case Number:"
  Text 15, 75, 80, 10, "Active On:"
  Text 15, 110, 60, 10, "Pending On:"
  Text 15, 205, 60, 10, "Worker Signature:"
  Text 15, 50, 110, 20, "Worker number you're transferring to  (x102XXX format)"
EndDialog



'----------THE SCRIPT----------
EMConnect ""

maxis_check_function

DIALOG xfer_menu_dialog
	IF ButtonPressed = 0 THEN stopscript

call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""
	
IF XFERRadioGroup = 0 THEN
		'Displays the dialog and navigates to case note
		Do
		  Do
			Do
			  Dialog within_county_dlg
			  If buttonpressed = 0 then stopscript
			  If case_number = "" then MsgBox "You must have a case number to continue."
			  IF len(worker_to_transfer_to) <> 7 then Msgbox "Please include X1## in the worker number"
			  IF preg_y_n = "Select one..." THEN MsgBox "Please indicate if a pregnancy verification was submitted or N/A if that is not applicable."
			  IF unit_drop_down = "Select one..." THEN MsgBox "Please indicate the unit to which the case is being transferred or N/A if that is not applicable."
			Loop until case_number <> "" and len(worker_to_transfer_to) = 7 AND preg_y_n <> "Select one..." AND unit_drop_down <> "Select one..."
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
		  IF mnsure_active_check = 1 THEN active_programs = active_programs & "Mnsure/"
		  IF EMER_active_check = 1 THEN active_programs = active_programs & "EMER/"

		  IF SNAP_pend_check = 1 THEN pend_programs = pend_programs & "SNAP/"
		  IF hc_pend_check = 1 THEN pend_programs = pend_programs & "HC/"
		  IF cash_pend_check = 1 THEN pend_programs = pend_programs & "CASH/"
		  IF mnsure_pend_check = 1 THEN pend_programs = pend_programs & "Mnsure/"
		  IF EMER_pend_check = 1 THEN pend_programs = pend_programs & "EMER/"

		'Case notes
		EMSendKey "***Transfer within county***" & "<newline>"
		IF unit_drop_down <> "N/A" THEN call write_variable_in_case_note("* Transfer to: " & unit_drop_down) 
		IF active_programs <> "" THEN
		  EMSendKey "* Active Programs: " & active_programs & "<backspace>"
		  EMSendKey "<newline>"
		END IF
		IF pend_programs <> "" THEN 
		  EMSendKey "* Pending Programs: " & pend_programs & "<backspace>"
		  EMSendKey "<newline>"
		END IF
		IF preg_y_n <> "N/A" THEN call write_variable_in_case_note("* Pregnancy verification rec'd: " & preg_y_n)
		call write_bullet_and_variable_in_case_note("Reason for transfer", Transfer_reason) 
		call write_bullet_and_variable_in_case_note("Actions to be taken", Action_to_be_taken) 
		call write_variable_in_case_note("---")
		call write_variable_in_case_note(worker_signature)

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
				DO
					DO
						DIALOG out_of_county_dlg
							IF ButtonPressed = 0 THEN stopscript
							IF ButtonPressed = nav_to_xfer_button THEN 
								CALL navigate_to_screen("SPEC", "XFER")
								EMWriteScreen "X", 9, 16
								transmit
							END IF
					LOOP UNTIL ButtonPressed = -1	
						last_chance = MsgBox("Do you want to continue? NOTE: You will get a chance to review SPEC/XFER before transmitting to transfer.", vbYesNo)
				LOOP UNTIL last_chance = vbYes

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
					IF manual_cfr_hc_check = 1 AND (hc_cfr = "" OR len(hc_cfr) <> 2 OR hc_cfr_month = "" OR hc_cfr_year = "" OR len(hc_cfr_month) <> 2 OR len(hc_cfr_year) <> 2) THEN MsgBox ("You indicated you wish to manually determine the Health Care County of Financial Responsibility and CFR Change Date. There is an error because you either:" & vbCr & vbCr & "1. Did not enter a County of Financial Responsibility, and/or" & vbCr & "2. Your County of Financial Responsibility is not in the correct 2-digit format, and/or" & vbCr & "3. You did not enter a date for the CFR to change, and/or" & vbCr & "4. You did not correctly format the month and/or year for the change." & vbCr & vbCr & "Please review your input and try again.")
					IF manual_cfr_cash_check = 1 AND cash_cfr_no_change_check = 1 THEN MsgBox ("Please select whether the CFR for CASH is changing or not. Review input.")
					IF manual_cfr_cash_check = 1 AND (cash_cfr = "" OR len(cash_cfr) <> 2 OR cash_cfr_month = "" OR cash_cfr_year = "" OR len(cash_cfr_month) <> 2 OR len(cash_cfr_year) <> 2) THEN MsgBox ("You indicated you wish to manually determine the CASH County of Financial Responsibility and CFR Change Date. There is an error because you either:" & vbCr & vbCr & "1. Did not enter a County of Financial Responsibility, and/or" & vbCr & "2. Your County of Financial Responsibility is not in the correct 2-digit format, and/or" & vbCr & "3. You did not enter a date for the CFR to change, and/or" & vbCr & "4. You did not correctly format the month and/or year for the change." & vbCr & vbCr & "Please review your input and try again.")
					IF manual_cfr_hc_check = 1 AND hc_cfr_no_change_check = 1 THEN MsgBOx ("Please select whether the CFR for HC is changing or not. Review input.")
					
			LOOP UNTIL (((HC_status <> "ACTV") OR (hc_status = "ACTV" AND excluded_time = "No") OR (HC_status = "ACTV" AND excluded_time = "Yes" AND isdate(excl_date) = TRUE))) AND _
			(isdate(cl_move_date) = TRUE) AND _
			(len(transfer_to) = 7) AND _
			(ucase(left(transfer_to, 4)) <> ucase(worker_county_code)) AND _ 
			(manual_cfr_hc_check = 0 OR (manual_cfr_hc_check = 1 AND hc_cfr <> "" AND len(hc_cfr) = 2 AND hc_cfr_month <> "" AND hc_cfr_year <> "" AND len(hc_cfr_month) = 2 AND len(hc_cfr_year) = 2)) AND _
			(manual_cfr_cash_check = 0 OR (manual_cfr_cash_check = 1 AND cash_cfr <> "" AND len(cash_cfr) = 2 AND cash_cfr_month <> "" AND cash_cfr_year <> "" AND len(cash_cfr_month) = 2 AND len(cash_cfr_year) = 2)) AND _
			((manual_cfr_cash_check = 0 AND cash_cfr_no_change_check = 0) OR (manual_cfr_cash_check = 1 AND cash_cfr_no_change_check = 0) OR (manual_cfr_cash_check = 0 AND cash_cfr_no_change_check = 1))  AND _
			((manual_cfr_hc_check = 0 AND hc_cfr_no_change_check = 0) OR (manual_cfr_hc_check = 1 AND hc_cfr_no_change_check = 0) OR (manual_cfr_hc_check = 0 AND hc_cfr_no_change_check = 1))			

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

		IF ucase(transfer_to) = "X120ICT" OR ucase(transfer_to) = "X181ICT" THEN transfer_to = "X174ICT"

		'Using move date to determine CRF change date.
		IF manual_cfr_cash_check = 0 AND cash_cfr_no_change_check = 0 THEN
			cash_cfr = right(worker_county_code, 2)
			cash_cfr_date = dateadd("M", 1, cl_move_date)
			cash_cfr_date = datepart("M", cash_cfr_date) & "/01/" & datepart("YYYY", cash_cfr_date)
			cash_cfr_date = dateadd("M", 2, cash_cfr_date)
			cash_cfr_month = datepart("M", cash_cfr_date)
			IF len(cash_cfr_month) <> 2 THEN cash_cfr_month = "0" & cash_cfr_month
			cash_cfr_year = datepart("YYYY", cash_cfr_date)
			cash_cfr_year = right(cash_cfr_year, 2)
		END IF
		IF manual_cfr_hc_check = 0 AND hc_cfr_no_change_check = 0 THEN
			hc_cfr = right(worker_county_code, 2)
			hc_cfr_date = dateadd("M", 1, cl_move_date)
			hc_cfr_date = datepart("M", hc_cfr_date) & "/01/" & datepart("YYYY", hc_cfr_date)
			hc_cfr_date = dateadd("M", 2, hc_cfr_date)
			hc_cfr_month = datepart("M", hc_cfr_date)
			IF len(hc_cfr_month) <> 2 THEN hc_cfr_month = "0" & hc_cfr_month
			hc_cfr_year = datepart("YYYY", hc_cfr_date)
			hc_cfr_year = right(hc_cfr_year, 2)
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

		call navigate_to_screen("CASE", "NOTE")
		PF9
		EMReadScreen write_access, 9, 24, 12
		IF write_access = "READ ONLY" THEN MsgBox("You do not have access to modify this case. Please double check your case number and try again." & chr(13) & chr(13) & "Alternatively, you may be in INQUIRY MODE.")
	LOOP UNTIL write_access <> "READ ONLY"
	PF3

	'----------Sending the CL a SPEC/MEMO notifying them of the details of the transfer----------
	If SPEC_MEMO_check = checked then
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
	End if

	'----------The case note of the reason for the XFER----------
		call navigate_to_screen("CASE", "NOTE")
		PF9
		IF SNAP_status = "ACTV" THEN active_programs = active_programs & "SNAP/"
		IF hc_status = "ACTV" THEN active_programs = active_programs & "HC/"
		IF cash_one_status = "ACTV" OR cash_two_status = "ACTV" THEN active_programs = active_programs & "CASH/"

		IF SNAP_status = "PEND" THEN pend_programs = pend_programs & "SNAP/"
		IF hc_status = "PEND" THEN pend_programs = pend_programs & "HC/"
		IF cash_pend_check = checked THEN pend_programs = pend_programs & "CASH/"

	call write_variable_in_case_note("**CASE TRANSFER TO " & ucase(transfer_to) & "**")
	IF active_programs <> "" THEN call write_bullet_and_variable_in_case_note("Active Programs", (left(active_programs, (len(active_programs) - 1))))
	IF mnsure_active_check = checked THEN call write_variable_in_case_note("* CL is active on HC through MNSure")
	IF mcre_active_check = checked THEN call write_variable_in_case_note("* CL is active on HC through Minnesota Care")
	IF pend_programs <> "" THEN call write_bullet_and_variable_in_case_note("Pending Programs", (left(pend_programs, (len(pend_programs) - 1))))
	IF mnsure_pend_check = checked THEN call write_variable_in_case_note("* CL has pending HC application through MNSure")
	call write_bullet_and_variable_in_case_note("CL Move Date", cl_move_date)
	IF crf_sent_check = checked THEN 
		crf_sent_date = date
		call write_bullet_and_variable_in_case_note("Change Report Sent", crf_sent_date)
	END IF
	IF excluded_time = "Yes" THEN
		excluded_time = excluded_time & ", Begins " & excl_date
		call write_bullet_and_variable_in_case_note("Excluded Time" , excluded_time)
		excluded_time = "Yes"
	ELSEIF excluded_time = "No" THEN
		call write_bullet_and_variable_in_case_note("Excluded Time", excluded_time)
	ELSEIF excluded_time = "" THEN
		call write_variable_in_case_note("* Excluded Time: N/A")
	END IF
	IF hc_status = "ACTV" THEN 
		CALL write_bullet_and_variable_in_case_note("HC County of Financial Responsibility", hc_cfr)
		IF hc_cfr_no_change_check = 0 THEN
			CALL write_bullet_and_variable_in_case_note("HC CFR Change Date", (hc_cfr_month & "/" & hc_cfr_year))
		ELSE
			CALL write_bullet_and_variable_in_case_note("HC CFR", "Not changing")
		END IF
	END IF
	IF cash_one_status = "ACTV" OR cash_two_status = "ACTV" THEN
		CALL write_bullet_and_variable_in_case_note("CASH County of Financial Responsibility", cash_cfr)
		IF cash_cfr_no_change_check = 0 THEN
			CALL write_bullet_and_variable_in_case_note("CASH CFR Change Date", (cash_cfr_month & "/" & cash_cfr_year))
		ELSE
			CALL write_bullet_and_variable_in_case_note("CASH CFR", "Not changing")
		END IF
	END IF
	IF closure_date_check = checked THEN call write_variable_in_case_note("* CL has until " & closure_date & " to provide required proofs or the case will close.")
	If transfer_form_check = checked THEN call write_variable_in_case_note("* DHS 3195 Inter Agency Case Transfer Form completed and sent.")
	IF Transfer_reason <> "" THEN CALL write_bullet_and_variable_in_case_note("Reason for Transfer", Transfer_reason)
	IF Action_to_be_taken <> "" THEN call write_bullet_and_variable_in_case_note("Actions to be Taken", Action_to_be_taken)
	call write_variable_in_case_note("---")
	call write_variable_in_case_note(worker_signature)
	PF3	

	'----------The business end of the script (DON'T POINT THIS SCRIPT AT ANYTHING YOU DON'T WANT TRANSFERRED!!)----------
	call navigate_to_screen("SPEC", "XFER")
	EMWriteScreen "X", 9, 16
	transmit
	PF9
	
	'Writing client move date
	call create_MAXIS_friendly_date(cl_move_date, 0, 4, 28)
	
	'Writing CRF date, if CRF check is checked
	IF crf_sent_check = checked THEN call create_MAXIS_friendly_date(crf_sent_date, 0, 4, 61)
	
	'Writes the excluded time info. Only need the left character (it's a dropdown)
	EMWriteScreen left(excluded_time, 1), 5, 28
	
	'If there's excluded time, need to write the info
	
	IF excluded_time = "Yes" THEN 
		call create_MAXIS_friendly_date(excl_date, 0, 6, 28)
		EMWriteScreen hc_cfr, 15, 39
	END IF
	
	IF excl_date = "" AND excluded_time = "No" THEN
		EMWriteScreen "__", 6, 28
		EMWriteScreen "__", 6, 31
		EMWriteScreen "__", 6, 34
	END IF

	IF hc_status = "ACTV" AND hc_cfr_no_change_check = 0 THEN
		EMWriteScreen hc_cfr, 14, 39
		EMWriteScreen hc_cfr_month, 14, 53
		EMWriteScreen hc_cfr_year, 14, 59
	END IF

	IF cash_one_status = "ACTV" AND cash_cfr_no_change_check = 0 THEN
		EMWriteScreen cash_cfr, 11, 39
		EMWriteScreen cash_cfr_month, 11, 53
		EMWriteScreen cash_cfr_year, 11, 59
	END IF

	IF cash_two_status = "ACTV" AND cash_cfr_no_change_check = 0 THEN
		EMWriteScreen cash_cfr, 12, 39
		EMWriteScreen cash_cfr_month, 12, 53
		EMWriteScreen cash_cfr_year, 12, 59
	END IF

	EMReadScreen primary_worker, 7, 21, 16
	EMWriteScreen primary_worker, 18, 28
	EMWriteScreen transfer_to, 18, 61

	script_end_procedure("The script has added a case note, created any requested memos, and has updated SPEC/XFER. Please review the information on SPEC/XFER and transfer the case.")
	
END IF
