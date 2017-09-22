'Required for statistical purposes==========================================================================================
name_of_script = "CA-Out of County Transfer.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 229                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("08/28/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'VARIABLES TO DECLARE----------------------------------------------------------------------------
SPEC_MEMO_check = checked		'Should default to checked, as we usually want to send a new worker memo

BeginDialog out_of_county_dlg, 0, 0, 206, 105, "Out of County Case Transfer"
  ButtonGroup ButtonPressed
    PushButton 125, 15, 75, 15, "Geocoder", Geo_coder_button
  EditBox 60, 5, 55, 15, MAXIS_case_number
  EditBox 60, 25, 55, 15, transfer_to
  EditBox 80, 45, 120, 15, Transfer_reason
  EditBox 80, 65, 120, 15, Action_to_be_taken
  ButtonGroup ButtonPressed
    OkButton 95, 85, 50, 15
    CancelButton 150, 85, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 30, 50, 10, "Transferring to:"
  Text 5, 50, 70, 10, "Reason for Transfer:"
  Text 5, 70, 65, 10, "Actions to be taken:"
EndDialog

'----------THE SCRIPT----------
EMConnect ""

call MAXIS_case_number_finder(MAXIS_case_number)


	DO
		DO
			DO
				DO
					DO
						DIALOG out_of_county_dlg
							cancel_confirmation
							Dialog initial_dialog
							IF buttonpressed = cancel THEN stopscript
							IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
							IF buttonpressed = Geo_coder_button THEN CreateObject("WScript.Shell").Run("https://hcgis.hennepin.us/agsinteractivegeocoder/default.aspx")
							IF ButtonPressed = nav_to_xfer_button THEN
								CALL navigate_to_MAXIS_screen("SPEC", "XFER")
								EMWriteScreen "X", 20, 22 'this is out of county'								
								transmit
							END IF
					LOOP UNTIL ButtonPressed = -1
						last_chance = MsgBox("Do you want to continue? NOTE: You will get a chance to review SPEC/XFER before transmitting to transfer.", vbYesNo)
				LOOP UNTIL last_chance = vbYes

				'----------Goes to STAT/PROG to pull active/pending case information----------
				call navigate_to_MAXIS_screen("STAT", "PROG")
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
			call navigate_to_MAXIS_screen("REPT", "USER")
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

		call navigate_to_MAXIS_screen("CASE", "NOTE")
		PF9
		EMReadScreen write_access, 9, 24, 12
		IF write_access = "READ ONLY" THEN MsgBox("You do not have access to modify this case. Please double check your case number and try again." & chr(13) & chr(13) & "Alternatively, you may be in INQUIRY MODE.")
	LOOP UNTIL write_access <> "READ ONLY"
	PF3

	'----------Sending the CL a SPEC/MEMO notifying them of the details of the transfer----------
	If SPEC_MEMO_check = checked then
		call navigate_to_MAXIS_screen("SPEC", "MEMO")
		'Creates a new MEMO. If it's unable the script will stop.
		PF5
		EMReadScreen memo_display_check, 12, 2, 33
		If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. You may not have the proper authorization to send a Spec/Memo or you may be in inquiry by mistake. Please try again in production first then contact your supervisor about obtaining permissions if still unable to access.")

		'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
		row = 4                             'Defining row and col for the search feature.
		col = 1
		EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
		IF row > 4 THEN                     'If it isn't 4, that means it was found.
    	arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
    	call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
    	EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
    	call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
    	PF5                                                     'PF5s again to initiate the new memo process
		END IF
		'Checking for SWKR
		row = 4                             'Defining row and col for the search feature.
		col = 1
		EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
		IF row > 4 THEN                     'If it isn't 4, that means it was found.
    	swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
    	call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
    	EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
    	call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
    	PF5                                           'PF5s again to initiate the new memo process
		END IF
		EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
		IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		transmit                                                        'Transmits to start the memo writing process

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
		call navigate_to_MAXIS_screen("CASE", "NOTE")
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
	IF SPEC_MEMO_check = checked THEN call write_variable_in_case_note("SPEC / MEMO sent to client with new worker information.")
	IF forms_to_arep = "Y" THEN call write_variable_in_case_note("Copy of SPEC/MEMO sent to AREP.")
	IF forms_to_swkr = "Y" THEN call write_variable_in_case_note("Copy of SPEC/MEMO sent to social worker.")
	call write_variable_in_case_note("---")
	call write_variable_in_case_note(worker_signature)
	PF3

	'----------The business end of the script (DON'T POINT THIS SCRIPT AT ANYTHING YOU DON'T WANT TRANSFERRED!!)----------
	call navigate_to_MAXIS_screen("SPEC", "XFER")
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


script_end_procedure("Success! The script has added a case note, created any requested memos, and has updated SPEC/XFER. Please review the information on SPEC/XFER, send the case file to the next county, and transfer the case.")


