'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - REVIEW REPORT.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	100			 'manual run time in seconds
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
call changelog_update("10/15/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function add_autoclose_case_note(revw_status_cash, revw_status_snap, revw_status_hc, revw_form_date, revw_intvw_date)
	If add_case_note = True Then		'only run the details here if we are running the 'end of month processing'
		If revw_status_cash = "T" OR revw_status_cash = "I" OR revw_status_cash = "U" OR revw_status_snap = "T" OR revw_status_snap = "I" OR revw_status_snap = "U" Then
		'We only care if the review has been terminated by the system, which is only indicated if the REVW status is T, I or U'
			Call navigate_to_MAXIS_screen("CASE", "NOTE")						'navigating to CASE:NOTE now
			EMReadScreen pw_county, 2, 21, 16									'reading to make sure this is still in Hennepin Country
			If pw_county = "27" Then
				autoclosed_programs = ""										'resetting variables
				cash_1_autoclosed = ""
				cash_2_autoclosed = ""
				snap_autoclosed = ""
				hc_autoclosed = ""
				n_code_programs = ""

				Call read_boolean_from_excel(objExcel.cells(excel_row,  6).value, MFIP_status)		'reading the program status information from the Review Report information
				Call read_boolean_from_excel(objExcel.cells(excel_row,  7).value, DWP_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row,  8).value, GA_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row,  9).value, MSA_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row, 10).value, GRH_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row, 13).value, SNAP_status)

				REPT_full = REPT_month & "/" & REPT_year						'creating a string from the review month and year for comparing the information in the REVW columns of the Review Report
				CASH_SR_Info = trim(objExcel.cells(excel_row, 11).value)
				CASH_ER_Info = trim(objExcel.cells(excel_row, 12).value)
				SNAP_SR_Info = trim(objExcel.cells(excel_row, 14).value)
				SNAP_ER_Info = trim(objExcel.cells(excel_row, 15).value)

				'CASH is first - if the status is T, I, or U - we are going to look at program status to firgure out which program of cash that it actually is
				If revw_status_cash = "T" OR revw_status_cash = "I" OR revw_status_cash = "U" Then
					If CASH_ER_Info = REPT_full then		'These programs are only ER
						If MFIP_status = True Then
							autoclosed_programs = autoclosed_programs & "/MFIP"
							If cash_1_autoclosed = "" Then cash_1_autoclosed = "MFIP ER"
							If cash_1_autoclosed <> "" Then cash_2_autoclosed = "MFIP ER"
						End If
						If DWP_status = True Then
							autoclosed_programs = autoclosed_programs & "/DWP"
							If cash_1_autoclosed = "" Then cash_1_autoclosed = "DWP ER"
							If cash_1_autoclosed <> "" Then cash_2_autoclosed = "DWP ER"
						End If
						If GA_status = True Then
							autoclosed_programs = autoclosed_programs & "/GA"
							If cash_1_autoclosed = "" Then cash_1_autoclosed = "GA ER"
							If cash_1_autoclosed <> "" Then cash_2_autoclosed = "GA ER"
						End If
						If MSA_status = True Then
							autoclosed_programs = autoclosed_programs & "/MSA"
							If cash_1_autoclosed = "" Then cash_1_autoclosed = "MSA ER"
							If cash_1_autoclosed <> "" Then cash_2_autoclosed = "MSA ER"
						End If
					End If
					If GRH_status = True AND (CASH_SR_Info = REPT_full OR CASH_ER_Info = REPT_full) Then		'GRH could be ER or SR
						If GRH_status = True Then
							autoclosed_programs = autoclosed_programs & "/GRH"
							If CASH_SR_Info = REPT_full Then
								If cash_1_autoclosed = "" Then cash_1_autoclosed = "GRH SR"
								If cash_1_autoclosed <> "" Then cash_2_autoclosed = "GRH SR"
							End If
							If CASH_ER_Info = REPT_full Then
								If cash_1_autoclosed = "" Then cash_1_autoclosed = "GRH ER"						'ER will overwrite SR because it is higher on the heirarchy
								If cash_1_autoclosed <> "" Then cash_2_autoclosed = "GRH ER"
							End If
						End If
					End If
				End If
				'Now looking at SNAP if the review status is T, I, or '
				If revw_status_snap = "T" OR revw_status_snap = "I" OR revw_status_snap = "U" Then
					If SNAP_status = True AND (SNAP_SR_Info = REPT_full OR SNAP_ER_Info = REPT_full) Then
						autoclosed_programs = autoclosed_programs & "/SNAP"
						If SNAP_SR_Info = REPT_full Then snap_autoclosed = "SNAP SR"
						If SNAP_ER_Info = REPT_full Then snap_autoclosed = "SNAP ER"							'ER will overwirte SR
					End If
				End If
				' 'HC Cases not set up yet as no REVWs and we cannot test
				' If revw_status_hc = "T" Then
				' End If

				'Now we check for any programs that have an 'N' as the review status so we can add a line about a program that may not be active.
				If revw_status_cash = "N" Then
					If MFIP_status = True Then n_code_programs = n_code_programs & "/MFIP"
					If DWP_status = True Then n_code_programs = n_code_programs & "/DWP"
					If GA_status = True Then n_code_programs = n_code_programs & "/GA"
					If MSA_status = True Then n_code_programs = n_code_programs & "/MSA"
					If GRH_status = True Then n_code_programs = n_code_programs & "/GRH"
				End If
				If revw_status_snap = "N" Then n_code_programs = n_code_programs & "/SNAP"

				'If there is at least one 'autoclosed' program, we are going to enter the note.
				If autoclosed_programs <> "" Then
					If left(autoclosed_programs, 1) = "/" Then autoclosed_programs = right(autoclosed_programs, len(autoclosed_programs)-1)
					If left(n_code_programs, 1) = "/" Then n_code_programs = right(n_code_programs, len(n_code_programs)-1)
					Call start_a_blank_case_note

					Call write_variable_in_CASE_NOTE(autoclosed_programs & " AUTOCLOSED eff " & REPT_month & "/" & REPT_year & " for Incomplete REVW")
					Call write_variable_in_CASE_NOTE("Renewals Terminated:")
					If cash_1_autoclosed <> "" Then Call write_variable_in_CASE_NOTE("    " & REPT_month & "/" & REPT_year & " " & cash_1_autoclosed)
					If cash_2_autoclosed <> "" Then Call write_variable_in_CASE_NOTE("    " & REPT_month & "/" & REPT_year & " " & cash_2_autoclosed)
					If snap_autoclosed <> "" Then Call write_variable_in_CASE_NOTE("    " & REPT_month & "/" & REPT_year & " " & snap_autoclosed)
					If hc_autoclosed <> "" Then Call write_variable_in_CASE_NOTE("    " & REPT_month & "/" & REPT_year & " " & hc_autoclosed)
					If revw_form_date <> "" Then Call write_variable_in_CASE_NOTE("Renewal Form Received on " & revw_form_date)
					' If revw_intvw_date <> "" Then Call write_variable_in_CASE_NOTE("Interview Completed on " & revw_intvw_date )				'Taking this out during the Interview Waiver time.
					Call write_variable_in_CASE_NOTE("Review case to determine additional actions to be taken.")
					If n_code_programs <> "" Then Call write_variable_in_CASE_NOTE("Check previous CASE:NOTE information for status about: " & n_code_programs)
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE("This is an automated process to NOTE a system action and no manual review of the case was completed. The programs autoclosed because the renewal process was incomplete, no action taken at county level.")
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature)
					' MsgBox "Look here"
					PF3															'saving the CASE:NOTE

					' Adding the note informaiton to the Review Report Excel
					ObjExcel.Cells(excel_row, closure_note_col) = "Yes"
					ObjExcel.Cells(excel_row, closure_progs_col) = autoclosed_programs
					' ObjExcel.Cells(excel_row, closure_progs_col+1) = n_code_programs
					' Msgbox "check"
				End If
			End If
			Call back_to_SELF
		End If
	End If
end function

'defining this function here because it needs to not end the script if a MEMO fails.
function start_a_new_spec_memo_and_continue(success_var)
'--- This function navigates user to SPEC/MEMO and starts a new SPEC/MEMO, selecting client, AREP, and SWKR if appropriate
'===== Keywords: MAXIS, notice, navigate, edit
    success_var = True
	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO

	PF5															'Creates a new MEMO. If it's unable the script will stop.
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then success_var = False

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
		EMReadScreen this_is_it, 60, row, col
	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	    PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 12                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit                                                        'Transmits to start the memo writing process
end function

'This is a script specific function and will not work outside of this script.
function read_case_details_for_review_report(incrementor_var)
	Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv) 'function to check PRIV status
	If is_this_priv = True then
		review_array(notes_const, incrementor_var) = "PRIV Case."
		review_array(interview_const, incrementor_var) = ""
		review_array(no_interview_const, incrementor_var) = ""
		review_array(current_SR_const, incrementor_var) = ""
	Else
		EmReadscreen worker_prefix, 4, 21, 14
		If worker_prefix <> "X127" then
			review_array(notes_const, i) = "Out-of-County: " & right(worker_prefix, 2)
			review_array(notes_const, incrementor_var) = "PRIV Case."
			review_array(interview_const, incrementor_var) = ""
			review_array(no_interview_const, incrementor_var) = ""
			review_array(current_SR_const, incrementor_var) = ""
		Else
			'function to determine programs and the program's status---Yay Casey!
			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)

			If case_active = False then
				review_array(notes_const, incrementor_var) = "Case Not Active."
			Else
				'valuing the array variables from the inforamtion gathered in from CASE/CURR
				review_array(MFIP_status_const, incrementor_var) = mfip_case
				review_array(DWP_status_const,  incrementor_var) = dwp_case
				review_array(GA_status_const,   incrementor_var) = ga_case
				review_array(MSA_status_const,  incrementor_var) = msa_case
				review_array(GRH_status_const,  incrementor_var) = grh_case
				review_array(SNAP_status_const, incrementor_var) = snap_case
				review_array(MA_status_const,   incrementor_var) = ma_case
				review_array(MSP_status_const,  incrementor_var) = msp_case
				'----------------------------------------------------------------------------------------------------STAT/REVW
				CALL navigate_to_MAXIS_screen("STAT", "REVW")

				If family_cash_case = True or adult_cash_case = True or grh_case = True then
					'read the CASH review information
					Call write_value_and_transmit("X", 5, 35) 'CASH Review Information
					EmReadscreen cash_review_popup, 11, 5, 35
					If cash_review_popup = "GRH Reports" then
					'The script will now read the CSR MO/YR and the Recert MO/YR
						EMReadScreen CSR_mo, 2, 9, 26
						EMReadScreen CSR_yr, 2, 9, 32
						EMReadScreen recert_mo, 2, 9, 64
						EMReadScreen recert_yr, 2, 9, 70

						CASH_CSR_date = CSR_mo & "/" & CSR_yr
						If CASH_CSR_date = "__/__" then CASH_CSR_date = ""

						CASH_ER_date = recert_mo & "/" & recert_yr
						If CASH_ER_date = "__/__" then CASH_ER_date = ""

						'Comparing CSR dates to the month of REVS review
						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN review_array(current_SR_const, incrementor_var) = True

						'Determining if a case is ER, and if it meets interview requirement
						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then
							If mfip_case = True then review_array(interview_const, incrementor_var) = True             'MFIP interview requirement
							IF adult_cash_case = True or grh_case = True then review_array(no_interview_const, incrementor_var) = True    'Adult CASH programs do not meet interview requirement
						End if

						'Next CASH ER and SR dates
						review_array(CASH_next_SR_const, incrementor_var) = CASH_CSR_date
						review_array(CASH_next_ER_const, incrementor_var) = CASH_ER_date
					Else
						review_array(notes_const, incrementor_var) = "Unable to Access CASH Review Information."
					End if
					Transmit 'to exit out of the pop-up screen
				End if

				If snap_case = True then
					'read the SNAP review information
					Call write_value_and_transmit("X", 5, 58) 'SNAP Review Information
					EmReadscreen food_review_popup, 20, 5, 30
					If food_review_popup = "Food Support Reports" then
					'The script will now read the CSR MO/YR and the Recert MO/YR
						EMReadScreen CSR_mo, 2, 9, 26
						EMReadScreen CSR_yr, 2, 9, 32
						EMReadScreen recert_mo, 2, 9, 64
						EMReadScreen recert_yr, 2, 9, 70

						SNAP_CSR_date = CSR_mo & "/" & CSR_yr
						If SNAP_CSR_date = "__/__" then SNAP_CSR_date = ""

						SNAP_ER_date = recert_mo & "/" & recert_yr
						If SNAP_ER_date = "__/__" then SNAP_ER_date = ""

						'Comparing CSR and ER daates to the month of REVS review
						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN review_array(current_SR_const, incrementor_var) = True

						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then review_array(interview_const, incrementor_var) = True

						'Next SNAP ER and SR dates
						review_array(SNAP_next_SR_const, incrementor_var) = SNAP_CSR_date
						review_array(SNAP_next_ER_const, incrementor_var) = SNAP_ER_date
					Else
						review_array(notes_const, incrementor_var) = "Unable to Access FS Review Information."
					End if
					Transmit 'to exit out of the pop-up screen
				End if

				If ma_case = True or msp_case = True then
					'read the HC review information
					Call write_value_and_transmit("X", 5, 71) 'HC Review Information
					EmReadscreen HC_review_popup, 20, 4, 32
					If HC_review_popup = "HEALTH CARE RENEWALS" then
					'The script will now read the CSR MO/YR and the Recert MO/YR
						EMReadScreen CSR_mo, 2, 8, 27   'IR dates
						EMReadScreen CSR_yr, 2, 8, 33
						If CSR_mo = "__" or CSR_yr = "__" then
							EMReadScreen CSR_mo, 2, 8, 71   'IR/AR dates
							EMReadScreen CSR_yr, 2, 8, 77
						End if
						EMReadScreen recert_mo, 2, 9, 27
						EMReadScreen recert_yr, 2, 9, 33

						HC_CSR_date = CSR_mo & "/" & CSR_yr
						If HC_CSR_date = "__/__" then HC_CSR_date = ""

						HC_ER_date = recert_mo & "/" & recert_yr
						If HC_ER_date = "__/__" then HC_ER_date = ""

						'Comparing CSR and ER daates to the month of REVS review
						IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN review_array(current_SR_const, incrementor_var) = True

						IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then review_array(no_interview_const, incrementor_var) = True

						'Next HC ER and SR dates
						review_array(HC_next_SR_const, incrementor_var) = HC_CSR_date
						review_array(HC_next_ER_const, incrementor_var) = HC_ER_date

						Transmit 'to exit out of the pop-up screen
					Else
						Transmit 'to exit out of the pop-up screen
						review_array(notes_const, i) = "Unable to Access HC Review Information."
					End if
				End if
			End if

			'----------------------------------------------------------------------------------------------------language and Contact Information
			'Gathering the phone numbers
			Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_number_one, phone_number_two, phone_number_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
			Call reformat_phone_number(phone_number_one, "( 111 ) 222 - 3333")
			Call reformat_phone_number(phone_number_two, "( 111 ) 222 - 3333")
			Call reformat_phone_number(phone_number_three, "( 111 ) 222 - 3333")
			review_array(phone_1_const, incrementor_var) = phone_number_one
			review_array(phone_2_const, incrementor_var) = phone_number_two
			review_array(phone_3_const, incrementor_var) = phone_number_three

			'Going to STAT/MEMB for Language Information
			CALL navigate_to_MAXIS_screen("STAT", "MEMB")
			EMReadScreen interpreter_code, 1, 14, 68
			EMReadScreen language_coded, 16, 12, 46
			language_coded = replace(language_coded, "_", "")
			If trim(language_coded) = "" then
				EMReadScreen lang_ID, 2, 12, 42
				If lang_ID = "99" then lang_ID = "English"
				language_coded = lang_ID
			End if

			review_array(Interpreter_const, incrementor_var) = interpreter_code
			review_array(Language_const, incrementor_var) = language_coded
		End if
	End if
end function

'This is a script specific function and will not work outside of this script.
function enter_excel_headers(ObjExcel)
	ObjExcel.Cells(t1_row_titles, t1_REVW_type).Value = "REVW Type"
	ObjExcel.Cells(t1_er_w_intv, t1_REVW_type).Value = "ER with Interview"
	ObjExcel.Cells(t1_er_no_intv, t1_REVW_type).Value = "ER - No Interview"
	ObjExcel.Cells(t1_csr, t1_REVW_type).Value = "CSR"
	ObjExcel.Cells(t1_priv, t1_REVW_type).Value = "PRIV"
	ObjExcel.Cells(t1_total, t1_REVW_type).Value = "Total"

	ObjExcel.Cells(t1_row_titles, t1_all_count).Value = "All"
	ObjExcel.Cells(t1_row_titles, t1_apps_revd_count).Value = "Apps Received"
	ObjExcel.Cells(t1_row_titles, t1_no_app_count).Value = "No App in MX"
	ObjExcel.Cells(t1_row_titles, t1_app_percent_count).Value = "Percent Received"
	ObjExcel.Cells(t1_row_titles, t1_interview_count).Value = "Interview Completed"
	ObjExcel.Cells(t1_row_titles, t1_no_intvw_count).Value = "No Interview"
	ObjExcel.Cells(t1_row_titles, t1_intvw_percent).Value = "Percent of Interviews Done"
	For i = t1_row_titles to t1_total
		ObjExcel.Cells(i, 1).Font.Bold = TRUE
	Next

	ObjExcel.Cells(1, t2_dates).Value = "Dates"

	ObjExcel.Cells(1, t2_er_w_intv_intv_count).Value = "Intvw ER - Intvw Count"
	ObjExcel.Cells(1, t2_er_w_intv_app_count).Value = "Intvw ER - App Recvd Count"
	ObjExcel.Cells(1, t2_er_no_intv_app_count).Value = "ER no Intvw - App Recvd Count"
	ObjExcel.Cells(1, t2_csr_app_count).Value = "CSR - App Recvd Count"
	ObjExcel.Cells(1, t2_priv_intv_count).Value = "PRIV - Intvw Count"
	ObjExcel.Cells(1, t2_priv_app_count).Value = "PRIV - App Recvd Count"
	ObjExcel.Cells(1, t2_total_intv_count).Value = "Total - Intvw Count"
	ObjExcel.Cells(1, t2_total_app_count).Value = "Total - App Recvd Count"

	ObjExcel.Cells(1, t3_apps_recvd_count).Value = "Apps Recvd Count"
	ObjExcel.Cells(1, t3_apps_recvd_percent).Value = "Apps Recvd %"

	ObjExcel.Cells(1, t3_intvs_count).Value = "Intvws Count"
	ObjExcel.Cells(1, t3_intvs_percent).Value = "Intvws %"
	ObjExcel.Cells(1, t3_revw_i_count).Value = "REVW - I Count"
	ObjExcel.Cells(1, t3_revw_i_percent).Value = "REVW - I %"
	ObjExcel.Cells(1, t3_revw_u_count).Value = "REVW - U Count"
	ObjExcel.Cells(1, t3_revw_u_percent).Value = "REVW -  U %"
	ObjExcel.Cells(1, t3_revw_n_count).Value = "REVW - N Count"
	ObjExcel.Cells(1, t3_revw_n_percent).Value = "REVW - N %"
	ObjExcel.Cells(1, t3_revw_a_count).Value = "REVW - A Count"
	ObjExcel.Cells(1, t3_revw_a_percent).Value = "REVW - A %"
	ObjExcel.Cells(1, t3_revw_o_count).Value = "REVW - O Count"
	ObjExcel.Cells(1, t3_revw_o_percent).Value = "REVW - O %"
	ObjExcel.Cells(1, t3_revw_t_count).Value = "REVW - T Count"
	ObjExcel.Cells(1, t3_revw_t_percent).Value = "REVW - T %"
	ObjExcel.Cells(1, t3_revw_d_count).Value = "REVW - D Count"
	ObjExcel.Cells(1, t3_revw_d_percent).Value = "REVW - D %"

	ObjExcel.Cells(1, t3_totals_count).Value = "Totals"
	ObjExcel.Range("A1").EntireRow.Font.Size = "14"

	ObjExcel.Cells(1,  t3_progs).Value = "REVW Type"

	ObjExcel.Cells(2,  t3_progs).Value = "ER w Intvw - All"
	ObjExcel.Cells(3,  t3_progs).Value = "ER w Intvw - Cash"
	ObjExcel.Cells(4,  t3_progs).Value = "ER w Intvw - SNAP"

	ObjExcel.Cells(5,  t3_progs).Value = "ER no Intvw - All"
	ObjExcel.Cells(6,  t3_progs).Value = "ER no Intvw - Cash"
	ObjExcel.Cells(7,  t3_progs).Value = "ER no Intvw - SNAP"

	ObjExcel.Cells(8,  t3_progs).Value = "CSR - All"
	ObjExcel.Cells(9,  t3_progs).Value = "CSR - GRH"
	ObjExcel.Cells(10,  t3_progs).Value = "CSR - SNAP"

	ObjExcel.Cells(11,  t3_progs).Value = "PRIV - All"
	ObjExcel.Cells(12,  t3_progs).Value = "PRIV - Cash"
	ObjExcel.Cells(13,  t3_progs).Value = "PRIV - SNAP"

	ObjExcel.Cells(14,  t3_progs).Value = "Totals - All"
	ObjExcel.Cells(15,  t3_progs).Value = "Totals - Cash"
	ObjExcel.Cells(16,  t3_progs).Value = "Totals - SNAP"
	for i = 2 to 16
		ObjExcel.Cells(i, t3_revw_types).Font.Bold = True
		ObjExcel.Cells(i, t3_progs).Font.Bold = True
	next

	first_of_rept_month = REPT_month & "/1/" & REPT_year
	first_of_rept_month = DateAdd("d", 0, first_of_rept_month)
	search_start = DateAdd("m", -2, first_of_rept_month)
	search_start = DateAdd("d", 15, search_start)

	date_row = 2
	the_date = search_start
	Do
		ObjExcel.Cells(date_row, t2_dates).Value = the_date
		the_date = DateAdd("d", 1, the_date)
		date_row = date_row + 1
	Loop until DateDiff("d", the_date, date) < 0
end function

'DECLARATIONS TO MAKE ===============================================================================================

'Row numbers for EXCEL for the STATISTICS Gathering - Numbers and Letters if needed
t1_row_titles			= 1
t1_row_titles_letter = convert_digit_to_excel_column(t1_row_titles)
t1_er_w_intv			= 2
t1_er_w_intv_letter = convert_digit_to_excel_column(t1_er_w_intv)
t1_er_no_intv			= 3
t1_er_no_intv_letter = convert_digit_to_excel_column(t1_er_no_intv)
t1_csr 					= 4
t1_csr_letter = convert_digit_to_excel_column(t1_csr)
t1_priv					= 5
t1_priv_letter = convert_digit_to_excel_column(t1_priv)
t1_total 				= 6
t1_total_letter = convert_digit_to_excel_column(t1_total)


'Column numbers for EXCEL for the STATISTICS Gathering - Numbers and Letters if needed
t1_REVW_type		= 1
t1_REVW_type_letter = convert_digit_to_excel_column(t1_REVW_type)
t1_all_count 		= 2
t1_all_count_letter = convert_digit_to_excel_column(t1_all_count)
t1_apps_revd_count	= 3
t1_apps_revd_count_letter = convert_digit_to_excel_column(t1_apps_revd_count)
t1_no_app_count 	= 4
t1_no_app_count_letter = convert_digit_to_excel_column(t1_no_app_count)
t1_app_percent_count = 5
t1_app_percent_count_letter = convert_digit_to_excel_column(t1_app_percent_count)
t1_interview_count	= 6
t1_interview_count_letter = convert_digit_to_excel_column(t1_interview_count)
t1_no_intvw_count	= 7
t1_no_intvw_count_letter = convert_digit_to_excel_column(t1_no_intvw_count)
t1_intvw_percent	= 8
t1_intvw_percent_letter = convert_digit_to_excel_column(t1_intvw_percent)



t2_dates				= 10
t2_dates_letter = convert_digit_to_excel_column(t2_dates)
t2_er_w_intv_intv_count	= 11
t2_er_w_intv_app_count	= 12
t2_er_no_intv_app_count	= 13
t2_csr_app_count		= 14
t2_priv_intv_count		= 15
t2_priv_app_count		= 16
t2_total_intv_count		= 17
t2_total_app_count		= 18
t2_total_app_count_letter = convert_digit_to_excel_column(t2_total_app_count)

t3_revw_types			= 20
t3_progs				= 21
t3_progs_letter = convert_digit_to_excel_column(t3_progs)
t3_apps_recvd_count		= 22
t3_apps_recvd_count_letter = convert_digit_to_excel_column(t3_apps_recvd_count)
t3_apps_recvd_percent	= 23
t3_iapps_recvd_percent_letter = convert_digit_to_excel_column(t3_apps_recvd_percent)
t3_intvs_count			= 24
t3_intvs_count_letter = convert_digit_to_excel_column(t3_intvs_count)
t3_intvs_percent		= 25
t3_intvs_percent_letter = convert_digit_to_excel_column(t3_intvs_percent)
t3_revw_i_count			= 26
t3_revw_i_count_letter = convert_digit_to_excel_column(t3_revw_i_count)
t3_revw_i_percent		= 27
t3_revw_i_percent_letter = convert_digit_to_excel_column(t3_revw_i_percent)
t3_revw_u_count			= 28
t3_revw_u_count_letter = convert_digit_to_excel_column(t3_revw_u_count)
t3_revw_u_percent		= 29
t3_revw_n_count			= 30
t3_revw_n_count_letter = convert_digit_to_excel_column(t3_revw_n_count)
t3_revw_n_percent		= 31
t3_revw_a_count			= 32
t3_revw_a_count_letter = convert_digit_to_excel_column(t3_revw_a_count)
t3_revw_a_percent		= 33
t3_revw_a_percent_letter = convert_digit_to_excel_column(t3_revw_a_percent)
t3_revw_o_count			= 34
t3_revw_o_count_letter = convert_digit_to_excel_column(t3_revw_o_count)
t3_revw_o_percent		= 35
t3_revw_t_count			= 36
t3_revw_t_count_letter = convert_digit_to_excel_column(t3_revw_t_count)
t3_revw_t_percent		= 37
t3_revw_d_count			= 38
t3_revw_d_count_letter = convert_digit_to_excel_column(t3_revw_d_count)
t3_revw_d_percent		= 39
t3_totals_count			= 40
t3_totals_count_letter = convert_digit_to_excel_column(t3_totals_count)

output_headers			= 42
output_data				= 43

'strings for use in Excel formulas
is_not_blank = chr(34) & "<>" & chr(34)
is_blank = chr(34) & chr(34)
is_true = chr(34)&"TRUE"&chr(34)
is_false = chr(34)&"FALSE"&chr(34)

'These are the constants that we need to create tables in Excel
Const xlSrcRange = 1
Const xlYes = 1

'constants for review_array
const worker_const          = 0
const case_number_const     = 1
const interview_const       = 2
const no_interview_const    = 3
const current_SR_const      = 4
const MFIP_status_const     = 5
const DWP_status_const      = 6
const GA_status_const       = 7
const MSA_status_const      = 8
const GRH_status_const      = 9
const CASH_next_SR_const    = 10
const CASH_next_ER_const    = 11
const SNAP_status_const     = 12
const SNAP_SR_status_const  = 13
const SNAP_next_SR_const	= 14
const SNAP_next_ER_const    = 15
const MA_status_const       = 16
const MSP_status_const      = 17
const HC_SR_status_const	= 18
const HC_ER_status_const	= 19
const HC_next_SR_const      = 20
const HC_next_ER_const      = 21
const Language_const        = 22
const Interpreter_const     = 23
const phone_1_const         = 24
const phone_2_const         = 25
const phone_3_const         = 26
const CASH_revw_status_const= 27
const SNAP_revw_status_const= 28
const HC_revw_status_const	= 29
const HC_MAGI_code_const	= 30
const review_recvd_const	= 31
const interview_date_const	= 32
const saved_to_excel_const	= 33
const notes_const           = 34

DIM review_array()              'declaring the array
ReDim review_array(notes_const, 0)       're-establihing size of array.

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
all_workers_check = 1		'defaulting the check box to checked
CM_plus_two_checkbox = 1    'defaulting the check box to checked
today_day = DatePart("d", date)
If today_day < 16 Then CM_plus_two_checkbox = unchecked
add_case_note = False
If Weekday(date) = 2 Then renewal_option = "Collect Statistics"
If today_day = 16 Then renewal_option = "Create Renewal Report"
If today_day = 18 Then renewal_option = "Send Appointment Letters"
If today_day = 15 Then renewal_option = "Send NOMIs"
If today_day = 1 Then renewal_option = "End of Processing Month"

'DISPLAYS DIALOG
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 85, "Review Report"
  DropListBox 90, 35, 90, 15, "Select one..."+chr(9)+"Create Renewal Report"+chr(9)+"Discrepancy Run"+chr(9)+"Collect Statistics"+chr(9)+"Send Appointment Letters"+chr(9)+"Send NOMIs"+chr(9)+"End of Processing Month"+chr(9)+"Create Worklist", renewal_option
  CheckBox 5, 55, 70, 10, "Select all agency.", all_workers_check
  CheckBox 5, 70, 70, 10, "Select for CM + 2.", CM_plus_two_checkbox
  ButtonGroup ButtonPressed
    OkButton 95, 65, 40, 15
    CancelButton 140, 65, 40, 15
  EditBox 70, 5, 110, 15, worker_number
  Text 5, 20, 175, 10, "Enter the fulll 7-digit worker #(s), comma separated."
  Text 5, 40, 85, 10, "Select a reporting option:"
  Text 5, 10, 60, 10, "Worker number(s):"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
        If renewal_option = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a renewal option."
        If worker_number = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Enter a valid worker number."
		If worker_number <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Enter a worker number OR select the entire agency, not both."
		If (CM_plus_two_checkbox = 1 and datePart("d", date) < 16) then err_msg = err_msg & VbNewLine & "* This is not a valid time period for REPT/REVS until the 16th of the month. Please select a new time period."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer
If CM_plus_two_checkbox = 1 then
    REPT_month = CM_plus_2_mo
    REPT_year  = CM_plus_2_yr
Else
    REPT_month = CM_plus_1_mo
    REPT_year  = CM_plus_1_yr
End if

'The End of Processing Month option is mostly 'collecting statistics' just with adding a CNOTE
If renewal_option = "End of Processing Month" Then
	renewal_option = "Collect Statistics"
	add_case_note = True
	last_day_checkbox = checked
	REPT_month = CM_mo
	REPT_year  = CM_yr
	call back_to_self
	EMReadScreen mx_region, 10, 22, 48

	If mx_region = "INQUIRY DB" Then
		continue_in_inquiry = MsgBox("It appears you are attempting to have the script send notices for these cases." & vbNewLine & vbNewLine & "However, you appear to be in MAXIS Inquiry." &vbNewLine & "*************************" & vbNewLine & "Do you want to continue?", vbQuestion + vbYesNo, "Confirm Inquiry")
		If continue_in_inquiry = vbNo Then script_end_procedure("Live script run was attempted in Inquiry and aborted.")
	End If
End If

report_date = REPT_month & "-" & REPT_year  'establishing review date

open_existing_review_report = FALSE
If renewal_option <> "Create Renewal Report" Then open_existing_review_report = TRUE

If open_existing_review_report = TRUE Then

	'If we are collecting statistics, we may be running on a current or past month, we need to clarify which month we are looking at.'
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 115, 55, "Select REVW Month for Information"
	  EditBox 75, 10, 15, 15, REPT_month
	  EditBox 95, 10, 15, 15, REPT_year
	  Text 10, 10, 60, 20, "Which REVW Month?"
	  ButtonGroup ButtonPressed
		OkButton 25, 35, 40, 15
		CancelButton 70, 35, 40, 15
	EndDialog

	Do
		Do
			err_msg = ""

			dialog Dialog1
			cancel_without_confirmation

		Loop Until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	report_date = REPT_month & "-" & REPT_year  'establishing review date

	'This is where the review report is currently saved.
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\" & report_date & " Review Report.xlsx"


	'Initial Dialog which requests a file path for the excel file
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 361, 65, "On Demand Recertifications - Send Appointment Notices"
	  EditBox 130, 20, 175, 15, excel_file_path
	  ButtonGroup ButtonPressed
		PushButton 310, 20, 45, 15, "Browse...", select_a_file_button
		If renewal_option = "Collect Statistics" Then CheckBox 10, 45, 205, 10, "Check here if this is the END of the processing month.", last_day_checkbox
		OkButton 250, 45, 50, 15
		CancelButton 305, 45, 50, 15
	  Text 10, 10, 170, 10, "Select the recert fle from the Review Report original run"
	  Text 10, 25, 120, 10, "Select an Excel file for recert cases:"
	EndDialog

	'Show file path dialog
	Do
		Dialog Dialog1
		cancel_confirmation
		If ButtonPressed = select_a_file_button then call file_selection_system_dialog(excel_file_path, ".xlsx")
	Loop until ButtonPressed = OK and excel_file_path <> ""

	'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
	call excel_open(excel_file_path, True, True, ObjExcel, objWorkbook)

	'Finding all of the worksheets available in the file. We will likely open up the main 'Review Report' so the script will default to that one.
	For Each objWorkSheet In objWorkbook.Worksheets
		If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
	Next
	scenario_dropdown = report_date & " Review Report"

	'Dialog to select worksheet
	'DIALOG is defined here so that the dropdown can be populated with the above code
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 151, 75, "Select the Worksheet"
	  DropListBox 5, 35, 140, 15, "Select One..." & scenario_list, scenario_dropdown
	  ButtonGroup ButtonPressed
		OkButton 40, 55, 50, 15
		CancelButton 95, 55, 50, 15
	  Text 5, 10, 130, 20, "Select the correct worksheet to run for review statistics:"
	EndDialog

	'Shows the dialog to select the correct worksheet
	Do
		Do
			Dialog Dialog1
			cancel_without_confirmation
		Loop until scenario_dropdown <> "Select One..."
		call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	'Activates worksheet based on user selection
	objExcel.worksheets(scenario_dropdown).Activate
End If

'Stats option ignores the 'list of workers' since it works off of an existing Excel, it needs to pull all of the workers
If renewal_option <> "Send Appointment Letters" Then
	If all_workers_check = checked then
		call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)

		new_array_string = ""
		For each worker in worker_array
			save_worker_numb = TRUE
			If worker = "X127V83" Then save_worker_numb = FALSE
			If worker = "X127VS2" Then save_worker_numb = FALSE
			If worker = "X127V51" Then save_worker_numb = FALSE
			If save_worker_numb = TRUE Then new_array_string = new_array_string & " " & worker
		Next
		new_array_string = trim(new_array_string)
		worker_array = split(new_array_string, " ")
	Else
		x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas
		'formatting array
		For each x1_number in x1s_from_dialog
			If worker_array = "" then
				worker_array = trim(x1_number)		'replaces worker_county_code if found in the typed x1 number
			Else
				worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
			End if
		Next
		'Split worker_array
		worker_array = split(worker_array, ", ")
	End if
End If

If renewal_option = "Send NOMIs" then

	call back_to_self
	EMReadScreen mx_region, 10, 22, 48

	If mx_region = "INQUIRY DB" Then
		continue_in_inquiry = MsgBox("It appears you are attempting to have the script send notices for these cases." & vbNewLine & vbNewLine & "However, you appear to be in MAXIS Inquiry." &vbNewLine & "*************************" & vbNewLine & "Do you want to continue?", vbQuestion + vbYesNo, "Confirm Inquiry")
		If continue_in_inquiry = vbNo Then script_end_procedure("Live script run was attempted in Inquiry and aborted.")
	End If

	date_month = DatePart("m", date)		'Creating a variable to enter in the column headers
	date_day = DatePart("d", date)
	date_header = date_month & "/" & date_day

	col_to_use = 1
	Do
		col_header = trim(ObjExcel.Cells(1, col_to_use).Value)
		If col_header = "CASH (" & date_header & ")" Then cash_stat_excel_col = col_to_use
		If col_header = "SNAP (" & date_header & ")" Then snap_stat_excel_col = col_to_use
		If col_header = "HC (" & date_header & ")" Then hc_stat_excel_col = col_to_use
		If col_header = "MAGI (" & date_header & ")" Then magi_stat_excel_col = col_to_use
		If col_header = "CAF Date (" & date_header & ")" Then recvd_date_excel_col = col_to_use
		If col_header = "Intvw Date (" & date_header & ")" Then intvw_date_excel_col = col_to_use

		col_to_use = col_to_use + 1
	Loop until col_header = ""

	If cash_stat_excel_col = "" Then
		renewal_option = "Collect Statistics"
		original_renewal_option = "Send NOMIs"
	End If
End If

If renewal_option = "Create Renewal Report" then
	review_report_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\" & report_date & " Review Report.xlsx"

	Set fso = CreateObject("Scripting.FileSystemObject")

	If (fso.FileExists(review_report_file_path)) Then
		'Opens Excel file since it exists
		call excel_open(review_report_file_path, True, True, ObjExcel, objWorkbook)

		'look through the rows to find the last one'
		excel_restart_line = 1
		Do
			excel_restart_line = excel_restart_line + 1
			review_cell_one = trim(ObjExcel.Cells(excel_restart_line, 3).Value)
			review_cell_two = trim(ObjExcel.Cells(excel_restart_line, 25).Value)
		Loop until review_cell_one = "" AND review_cell_two = ""
		excel_restart_line = excel_restart_line & ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 206, 115, "Restart Previous Run"
		  OptionGroup RadioGroup1
		    RadioButton 25, 50, 105, 10, "Yes! Restart from Excel Line ", restart_run_radio
			RadioButton 25, 70, 85, 10, "No, start a new report.", new_run_radio
		  EditBox 135, 45, 35, 15, excel_restart_line
		  ButtonGroup ButtonPressed
		    OkButton 100, 95, 50, 15
		    CancelButton 150, 95, 50, 15
		  Text 10, 10, 190, 10, "It appears this Review Report has already been created."
		  GroupBox 10, 30, 170, 55, "Do you need to RESTART a Report Creation?"
		EndDialog

		Do
			Do
				err_msg = ""

				dialog Dialog1
				cancel_without_confirmation

			Loop Until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE

		If new_run_radio = checked then
			objExcel.ActiveWorkbook.Close
			objExcel.Application.Quit
			objExcel.Quit
		Else
			excel_restart_line = excel_restart_line * 1
		End If
	Else
		new_run_radio = checked
	End If

	If new_run_radio = checked then
	    'Opening the Excel file, (now that the dialog is done)
	    Set objExcel = CreateObject("Excel.Application")
	    objExcel.Visible = True
	    Set objWorkbook = objExcel.Workbooks.Add()
	    objExcel.DisplayAlerts = True

	    'Changes name of Excel sheet to "Case information"
	    ObjExcel.ActiveSheet.Name = report_date & " Review Report"

	    'formatting excel file with columns for case number and interview date/time
	    objExcel.cells(1,  1).value = "X number"
	    objExcel.cells(1,  2).value = "Case number"
	    objExcel.cells(1,  3).value = "Interview ER"
	    objExcel.cells(1,  4).value = "No Interview ER"
	    objExcel.cells(1,  5).value = "Current SR"
	    objExcel.cells(1,  6).value = "MFIP Status"
	    objExcel.cells(1,  7).value = "DWP Status"
	    objExcel.cells(1,  8).value = "GA Status"
	    objExcel.cells(1,  9).value = "MSA Status"
	    objExcel.cells(1, 10).value = "HS/GRH Status"
	    objExcel.cells(1, 11).value = "CASH Next SR"
	    objExcel.cells(1, 12).value = "CASH Next ER"
	    objExcel.cells(1, 13).value = "SNAP Status"
	    objExcel.cells(1, 14).value = "Next SNAP SR"
	    objExcel.cells(1, 15).value = "Next SNAP ER"
	    objExcel.cells(1, 16).value = "MA Status"
	    objExcel.cells(1, 17).value = "MSP Status"
	    objExcel.cells(1, 18).value = "Next HC SR"
	    objExcel.cells(1, 19).value = "Next HC ER"
	    objExcel.cells(1, 20).value = "Case Language"
	    objExcel.Cells(1, 21).value = "Interpreter"
	    objExcel.cells(1, 22).value = "Phone # One"
	    objExcel.cells(1, 23).value = "Phone # Two"
	    objExcel.Cells(1, 24).value = "Phone # Three"
	    objExcel.Cells(1, 25).value = "Notes"

	    FOR i = 1 to 25									'formatting the cells'
	    	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	    	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	        objExcel.Columns(i).AutoFit()				'sizing the columns'
	    NEXT

	    excel_row = 2

	    back_to_self    'We need to get back to SELF and manually update the footer month
	    Call navigate_to_MAXIS_screen("REPT", "REVS")
	    EMWriteScreen REPT_month, 20, 55
	    EMWriteScreen REPT_year, 20, 58
	    transmit

	    'start of the FOR...next loop
	    For each worker in worker_array
	    	worker = trim(worker)
	        If worker = "" then exit for
	    	Call write_value_and_transmit(worker, 21, 6)   'writing in the worker number in the correct col

	        'Grabbing case numbers from REVS for requested worker
	    	DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
	    		row = 7	'Setting or resetting this to look at the top of the list
	    		DO		'All of this loops until row = 19
	    			'Reading case information (case number, SNAP status, and cash status)
	    			EMReadScreen MAXIS_case_number, 8, row, 6
	    			MAXIS_case_number = trim(MAXIS_case_number)
	    			EMReadScreen SNAP_status, 1, row, 45
	    			EMReadScreen cash_status, 1, row, 39
	                EmReadscreen HC_status, 1, row, 49

	    			'Navigates though until it runs out of case numbers to read
	    			IF MAXIS_case_number = "" then exit do

	    			'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
	    			If cash_status = "-" 	then cash_status = ""
	    			If SNAP_status = "-" 	then SNAP_status = ""
	    			If HC_status = "-" 		then HC_status = ""

	    			'Using if...thens to decide if a case should be added (status isn't blank)
	    			If ( ( trim(SNAP_status) = "N" or trim(SNAP_status) = "I" or trim(SNAP_status) = "U" or trim(SNAP_status) = "A" or trim(SNAP_status) = "O" or trim(SNAP_status) = "D" or trim(SNAP_status) = "T" )_
					or ( trim(cash_status) = "N" or trim(cash_status) = "I" or trim(cash_status) = "U" or trim(cash_status) = "A" or trim(cash_status) = "O" or trim(cash_status) = "D" or trim(cash_status) = "T" ) _
	                or ( trim(HC_status) = "N" or trim(HC_status) = "I" or trim(HC_status) = "U"  or trim(HC_status) = "A" or trim(HC_status) = "O" or trim(HC_status) = "D" or trim(HC_status) = "T" ) ) then
	                    'Adding the case information to Excel
	                    ObjExcel.Cells(excel_row, 1).value  = worker
	                    ObjExcel.Cells(excel_row, 2).value  = trim(MAXIS_case_number)
	                    excel_row = excel_row + 1
	                End if

	    			row = row + 1    'On the next loop it must look to the next row
	    			MAXIS_case_number = "" 'Clearing variables before next loop
	    		Loop until row = 19		'Last row in REPT/REVS
	    		'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
	    		PF8
	    		EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
	            'if max reviews are reached, the goes to next worker is applicable
	    	Loop until last_page_check = "THIS IS THE LAST PAGE"
	    next

	    'Saves and closes the most the main spreadsheet before continuing
	    objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\" & report_date & " Review Report.xlsx"
	End If

    'Establish the reviews array
    recert_cases = 0	            'incrementor for the array

    objExcel.worksheets(report_date & " Review Report").Activate  'Activates the review worksheet
    excel_row = 2   'Excel start row reading the case information for the array

    Do
        MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value 'reading case number
        MAXIS_case_number = trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit do

        worker = ObjExcel.Cells(excel_row, 1).Value

        ReDim Preserve review_array(notes_const, recert_cases)	'This resizes the array based on if master notes were found or not
        review_array(worker_const,          recert_cases) = trim(worker)
        review_array(case_number_const,     recert_cases) = MAXIS_case_number
        review_array(interview_const,       recert_cases) = False   'values defaulted to False
        review_array(no_interview_const,    recert_cases) = False
        review_array(current_SR_const,      recert_cases) = False
        review_array(MFIP_status_const,     recert_cases) = ""      'values start at blank
        review_array(DWP_status_const,      recert_cases) = ""
        review_array(GA_status_const,       recert_cases) = ""
        review_array(MSA_status_const,      recert_cases) = ""
        review_array(GRH_status_const,      recert_cases) = ""
        review_array(CASH_next_SR_const,    recert_cases) = ""
        review_array(CASH_next_ER_const,    recert_cases) = ""
        review_array(SNAP_status_const,     recert_cases) = ""
        review_array(SNAP_next_SR_const,    recert_cases) = ""
        review_array(SNAP_next_ER_const,    recert_cases) = ""
        review_array(MA_status_const,       recert_cases) = ""
        review_array(MSP_status_const,      recert_cases) = ""
        review_array(HC_SR_status_const,    recert_cases) = ""
        review_array(HC_ER_status_const,    recert_cases) = ""
        review_array(Language_const,        recert_cases) = ""
        review_array(Interpreter_const,     recert_cases) = ""
        review_array(phone_1_const,         recert_cases) = ""
        review_array(phone_2_const,         recert_cases) = ""
        review_array(phone_3_const,         recert_cases) = ""
        review_array(notes_const,           recert_cases) = ""
		If restart_run_radio = checked AND IsNumeric(excel_restart_line) = TRUE Then
		 	If excel_row = excel_restart_line Then starting_array_position = recert_cases
		End If

        'Incremented variables
        recert_cases = recert_cases + 1                 'array incrementor
        STATS_counter = STATS_counter + 1               'stats incrementor
        excel_row = excel_row + 1                       'Excel row incrementor
    LOOP

    '----------------------------------------------------------------------------------------------------MAXIS TIME
    back_to_SELF
    MAXIS_footer_month = CM_plus_1_mo
    MAXIS_footer_year = CM_plus_1_yr
    Call MAXIS_footer_month_confirmation

    total_cases_review = 0  'for total recert counts for stats
    excel_row = 2          'resetting excel_row to output the array information

    'DO 'Loops until there are no more cases in the Excel list
    For item = 0 to Ubound(review_array, 2)
    	MAXIS_case_number = review_array(case_number_const, item)

		If new_run_radio = checked Then
			Call read_case_details_for_review_report(item)

	        '----------------------------------------------------------------------------------------------------Excel Output
	        ObjExcel.Cells(excel_row,  3).value = review_array(interview_const,       item)     'COL C
	        ObjExcel.Cells(excel_row,  4).value = review_array(no_interview_const,    item)     'COL D
	        ObjExcel.Cells(excel_row,  5).value = review_array(current_SR_const,      item)     'COL E
	        ObjExcel.Cells(excel_row,  6).value = review_array(MFIP_status_const,     item)     'COL F
	        ObjExcel.Cells(excel_row,  7).value = review_array(DWP_status_const,      item)     'COL G
	        ObjExcel.Cells(excel_row,  8).value = review_array(GA_status_const,       item)     'COL H
	        ObjExcel.Cells(excel_row,  9).value = review_array(MSA_status_const,      item)     'COL I
	        ObjExcel.Cells(excel_row, 10).value = review_array(GRH_status_const,      item)     'COL J
	        ObjExcel.Cells(excel_row, 11).value = review_array(CASH_next_SR_const,    item)     'COL K
	        ObjExcel.Cells(excel_row, 12).value = review_array(CASH_next_ER_const,    item)     'COL L
	        ObjExcel.Cells(excel_row, 13).value = review_array(SNAP_status_const,     item)     'COL M
	        ObjExcel.Cells(excel_row, 14).value = review_array(SNAP_next_SR_const,    item)     'COL N
	        ObjExcel.Cells(excel_row, 15).value = review_array(SNAP_next_ER_const,    item)     'COL O
	        ObjExcel.Cells(excel_row, 16).value = review_array(MA_status_const,       item)     'COL P
	        ObjExcel.Cells(excel_row, 17).value = review_array(MSP_status_const,      item)     'COL Q
	        ObjExcel.Cells(excel_row, 18).value = review_array(HC_next_SR_const,      item)     'COL R
	        ObjExcel.Cells(excel_row, 19).value = review_array(HC_next_ER_const,      item)     'COL S
	        ObjExcel.Cells(excel_row, 20).value = review_array(Language_const,        item)     'COL T
	        ObjExcel.Cells(excel_row, 21).value = review_array(Interpreter_const,     item)     'COL U
	        ObjExcel.Cells(excel_row, 22).value = review_array(phone_1_const,         item)     'COL V
	        ObjExcel.Cells(excel_row, 23).value = review_array(phone_2_const,         item)     'COL W
	        ObjExcel.Cells(excel_row, 24).value = review_array(phone_3_const,         item)     'COL X
	        ObjExcel.Cells(excel_row, 25).value = review_array(notes_const,           item)     'COL Y
		End If

		If restart_run_radio = checked Then
			If item < starting_array_position Then
				'----------------------------------------------------------------------------------------------------Excel Output
				review_array(interview_const,       item) = ObjExcel.Cells(excel_row,  3).value      'COL C
				review_array(no_interview_const,    item) = ObjExcel.Cells(excel_row,  4).value      'COL D
				review_array(current_SR_const,      item) = ObjExcel.Cells(excel_row,  5).value      'COL E
				review_array(MFIP_status_const,     item) = ObjExcel.Cells(excel_row,  6).value      'COL F
				review_array(DWP_status_const,      item) = ObjExcel.Cells(excel_row,  7).value      'COL G
				review_array(GA_status_const,       item) = ObjExcel.Cells(excel_row,  8).value      'COL H
				review_array(MSA_status_const,      item) = ObjExcel.Cells(excel_row,  9).value      'COL I
				review_array(GRH_status_const,      item) = ObjExcel.Cells(excel_row, 10).value      'COL J
				review_array(CASH_next_SR_const,    item) = ObjExcel.Cells(excel_row, 11).value      'COL K
				review_array(CASH_next_ER_const,    item) = ObjExcel.Cells(excel_row, 12).value      'COL L
				review_array(SNAP_status_const,     item) = ObjExcel.Cells(excel_row, 13).value      'COL M
				review_array(SNAP_next_SR_const,    item) = ObjExcel.Cells(excel_row, 14).value      'COL N
				review_array(SNAP_next_ER_const,    item) = ObjExcel.Cells(excel_row, 15).value      'COL O
				review_array(MA_status_const,       item) = ObjExcel.Cells(excel_row, 16).value      'COL P
				review_array(MSP_status_const,      item) = ObjExcel.Cells(excel_row, 17).value      'COL Q
				review_array(HC_next_SR_const,      item) = ObjExcel.Cells(excel_row, 18).value      'COL R
				review_array(HC_next_ER_const,      item) = ObjExcel.Cells(excel_row, 19).value      'COL S
				review_array(Language_const,        item) = ObjExcel.Cells(excel_row, 20).value      'COL T
				review_array(Interpreter_const,     item) = ObjExcel.Cells(excel_row, 21).value      'COL U
				review_array(phone_1_const,         item) = ObjExcel.Cells(excel_row, 22).value      'COL V
				review_array(phone_2_const,         item) = ObjExcel.Cells(excel_row, 23).value      'COL W
				review_array(phone_3_const,         item) = ObjExcel.Cells(excel_row, 24).value      'COL X
				review_array(notes_const,           item) = ObjExcel.Cells(excel_row, 25).value      'COL Y
			Else
				Call check_for_MAXIS(FALSE)		'making sure we haven't passworded out
				Call read_case_details_for_review_report(item)

		        '----------------------------------------------------------------------------------------------------Excel Output
		        ObjExcel.Cells(excel_row,  3).value = review_array(interview_const,       item)     'COL C
		        ObjExcel.Cells(excel_row,  4).value = review_array(no_interview_const,    item)     'COL D
		        ObjExcel.Cells(excel_row,  5).value = review_array(current_SR_const,      item)     'COL E
		        ObjExcel.Cells(excel_row,  6).value = review_array(MFIP_status_const,     item)     'COL F
		        ObjExcel.Cells(excel_row,  7).value = review_array(DWP_status_const,      item)     'COL G
		        ObjExcel.Cells(excel_row,  8).value = review_array(GA_status_const,       item)     'COL H
		        ObjExcel.Cells(excel_row,  9).value = review_array(MSA_status_const,      item)     'COL I
		        ObjExcel.Cells(excel_row, 10).value = review_array(GRH_status_const,      item)     'COL J
		        ObjExcel.Cells(excel_row, 11).value = review_array(CASH_next_SR_const,    item)     'COL K
		        ObjExcel.Cells(excel_row, 12).value = review_array(CASH_next_ER_const,    item)     'COL L
		        ObjExcel.Cells(excel_row, 13).value = review_array(SNAP_status_const,     item)     'COL M
		        ObjExcel.Cells(excel_row, 14).value = review_array(SNAP_next_SR_const,    item)     'COL N
		        ObjExcel.Cells(excel_row, 15).value = review_array(SNAP_next_ER_const,    item)     'COL O
		        ObjExcel.Cells(excel_row, 16).value = review_array(MA_status_const,       item)     'COL P
		        ObjExcel.Cells(excel_row, 17).value = review_array(MSP_status_const,      item)     'COL Q
		        ObjExcel.Cells(excel_row, 18).value = review_array(HC_next_SR_const,      item)     'COL R
		        ObjExcel.Cells(excel_row, 19).value = review_array(HC_next_ER_const,      item)     'COL S
		        ObjExcel.Cells(excel_row, 20).value = review_array(Language_const,        item)     'COL T
		        ObjExcel.Cells(excel_row, 21).value = review_array(Interpreter_const,     item)     'COL U
		        ObjExcel.Cells(excel_row, 22).value = review_array(phone_1_const,         item)     'COL V
		        ObjExcel.Cells(excel_row, 23).value = review_array(phone_2_const,         item)     'COL W
		        ObjExcel.Cells(excel_row, 24).value = review_array(phone_3_const,         item)     'COL X
		        ObjExcel.Cells(excel_row, 25).value = review_array(notes_const,           item)     'COL Y
			End If
		End If
		excel_row = excel_row + 1
		total_cases_review = total_cases_review + 1
		STATS_counter = STATS_counter + 1						'adds one instance to the stats counter
		MAXIS_case_number = ""
    Next

    'Formatting the columns to autofit after they are all finished being created.
    FOR i = 1 to 25
    	objExcel.Columns(i).autofit()
    Next

    'Saves and closes the main reivew report
    objWorkbook.Save()
    objExcel.ActiveWorkbook.Close
    objExcel.Application.Quit
    objExcel.Quit

    '----------------------------------------------------------------------------------------------------Creating the Interview Required Excel List for the auto-dialer and notices
    'Opening the Excel file, (now that the dialog is done)
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add()
    objExcel.DisplayAlerts = True

    'Changes name of Excel sheet to "Case information"
    ObjExcel.ActiveSheet.Name = "ER cases " & REPT_month & "-" & REPT_year

    'formatting excel file with columns for case number and interview date/time
    objExcel.cells(1, 1).value 	= "X number"
    objExcel.cells(1, 2).value 	= "Case Number"
    objExcel.cells(1, 3).value 	= "Programs"
    objExcel.cells(1, 4).value 	= "Case language"
    objExcel.Cells(1, 5).value 	= "Interpreter"
    objExcel.cells(1, 6).value 	= "Phone # One"
    objExcel.cells(1, 7).value 	= "Phone # Two"
    objExcel.Cells(1, 8).value 	= "Phone # Three"

    FOR i = 1 to 8									'formatting the cells'
    	objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
    	objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    excel_row = 2 'Adding the case information to Excel
	recert_cases = 0

    For item = 0 to UBound(review_array, 2)
        If review_array(interview_const, item) = True then
            'determining the programs list
            If ( review_array(SNAP_status_const, item) = True and review_array(MFIP_status_const, item) = True ) then
                programs_list = "SNAP & MFIP"
            elseif review_array(SNAP_status_const, item) = True then
                programs_list = "SNAP"
            elseif review_array(MFIP_status_const, item) = True then
                programs_list = "MFIP"
            End if
            'Excel output of Interview Required case information
            If review_array(notes_const, item) <> "PRIV Case." then
    	        ObjExcel.Cells(excel_row, 1).value = review_array(worker_const,       item)
    	        ObjExcel.Cells(excel_row, 2).value = review_array(case_number_const,  item)
    	        ObjExcel.Cells(excel_row, 3).value = programs_list
    	        ObjExcel.Cells(excel_row, 4).value = review_array(Language_const,     item)
    	        ObjExcel.Cells(excel_row, 5).value = review_array(Interpreter_const,  item)
    	        ObjExcel.Cells(excel_row, 6).value = review_array( phone_1_const,     item)
    	        ObjExcel.Cells(excel_row, 7).value = review_array( phone_2_const,     item)
    	        ObjExcel.Cells(excel_row, 8).value = review_array( phone_3_const,     item)
				recert_cases = recert_cases + 1
                excel_row = excel_row + 1
            End if
        End if
    Next

    'Query date/time/runtime info
    objExcel.Cells(1, 11).Font.Bold = TRUE
    objExcel.Cells(2, 11).Font.Bold = TRUE
    objExcel.Cells(3, 11).Font.Bold = TRUE
    objExcel.Cells(4, 11).Font.Bold = TRUE
    ObjExcel.Cells(1, 11).Value = "Query date and time:"
    ObjExcel.Cells(2, 11).Value = "Query runtime (in seconds):"
    ObjExcel.Cells(3, 11).Value = "Total reviews:"
    ObjExcel.Cells(4, 11).Value = "Interview required:"
    ObjExcel.Cells(1, 12).Value = now
    ObjExcel.Cells(2, 12).Value = timer - query_start_time
    ObjExcel.Cells(3, 12).Value = total_cases_review
    ObjExcel.Cells(4, 12).Value = recert_cases

    'Formatting the columns to autofit after they are all finished being created.
    FOR i = 1 to 12
    	objExcel.Columns(i).autofit()
    Next

    ObjExcel.Worksheets.Add().Name = "Priviliged Cases"

    'adding information to the Excel list from PND2
    ObjExcel.Cells(1, 1).Value = "Worker #"
    ObjExcel.Cells(1, 2).Value = "Case number"

    FOR i = 1 to 2								'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    excel_row = 2   'Adding the case information to Excel

    For item = 0 to UBound(review_array, 2)
        'Excel output of Interview Required case information
        If review_array(notes_const, item) = "PRIV Case." then
            ObjExcel.Cells(excel_row, 1).value = review_array(worker_const,       item)
            ObjExcel.Cells(excel_row, 2).value = review_array(case_number_const,  item)
            excel_row = excel_row + 1
        End if
    Next

    'Formatting the columns to autofit after they are all finished being created.
    FOR i = 1 to 2
    	objExcel.Columns(i).autofit()
    Next

	end_msg = "Success! The review report is ready."
ElseIf renewal_option = "Collect Statistics" Then			'This option is used when we are ready to collect statistics about review cases.
	If REPT_month = CM_plus_2_mo AND REPT_year = CM_plus_2_yr Then
		MAXIS_footer_month = CM_plus_1_mo							'Setting the footer month and year based on the review month. We do not run statistics in CM + 2
		MAXIS_footer_year = CM_plus_1_yr
	Else
		MAXIS_footer_month = REPT_month							'Setting the footer month and year based on the review month. We do not run statistics in CM + 2
		MAXIS_footer_year = REPT_year
	End If

	info_sheet_name = replace(excel_file_path, t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\", "")

	'Finding the last column that has something in it so we can add to the end.
	col_to_use = 0
	Do
		col_to_use = col_to_use + 1
		col_header = trim(ObjExcel.Cells(1, col_to_use).Value)
	Loop until col_header = ""
	last_col_letter = convert_digit_to_excel_column(col_to_use)

	'Insert columns in excel for additional information to be added
	column_end = last_col_letter & "1"
	Set objRange = objExcel.Range(column_end).EntireColumn

	objRange.Insert(xlShiftToRight)			'We neeed six more columns
	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)

	cash_stat_excel_col = col_to_use		'Setting the columns to individual variables so we enter the found information in the right place
	snap_stat_excel_col = col_to_use + 1
	hc_stat_excel_col = col_to_use + 2
	magi_stat_excel_col = col_to_use + 3
	recvd_date_excel_col = col_to_use + 4
	intvw_date_excel_col = col_to_use + 5

	date_month = DatePart("m", date)		'Creating a variable to enter in the column headers
	date_day = DatePart("d", date)
	date_header = date_month & "/" & date_day

	ObjExcel.Cells(1, cash_stat_excel_col).Value = "CASH (" & date_header & ")"			'creating the column headers for the statistics information for the day of the run.
	ObjExcel.Cells(1, snap_stat_excel_col).Value = "SNAP (" & date_header & ")"
	ObjExcel.Cells(1, hc_stat_excel_col).Value = "HC (" & date_header & ")"
	ObjExcel.Cells(1, magi_stat_excel_col).Value = "MAGI (" & date_header & ")"
	ObjExcel.Cells(1, recvd_date_excel_col).Value = "CAF Date (" & date_header & ")"
	ObjExcel.Cells(1, intvw_date_excel_col).Value = "Intvw Date (" & date_header & ")"

	FOR i = col_to_use to col_to_use + 5									'formatting the cells'
		objExcel.Cells(1, i).Font.Bold = True		'bold font'
		ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
		objExcel.Columns(i).AutoFit()				'sizing the columns'
	NEXT

	'This is for th end of processing month option - it needs 2 additional columns.
	If add_case_note = True Then
		closure_note_col = col_to_use + 6
		ObjExcel.Cells(1, closure_note_col).Value = "Close Note"

		closure_progs_col = col_to_use + 7
		ObjExcel.Cells(1, closure_progs_col).Value = "Progs Closed"
	End If

	recert_cases = 0	            'incrementor for the array

	back_to_self    'We need to get back to SELF and manually update the footer month
    Call navigate_to_MAXIS_screen("REPT", "REVS")		'going to REPT REVS where all the information is displayed'
    EMWriteScreen REPT_month, 20, 55					'going to the right month
    EMWriteScreen REPT_year, 20, 58
    transmit

    'We are going to look at REPT/REVS for each worker in Hennepin County
    For each worker in worker_array
    	worker = trim(worker)				'get to the right worker
        If worker = "" then exit for
    	Call write_value_and_transmit(worker, 21, 6)   'writing in the worker number in the correct col

        'Grabbing case numbers from REVS for requested worker
    	DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
    		row = 7	'Setting or resetting this to look at the top of the list
    		DO		'All of this loops until row = 19
    			'Reading case information (case number, SNAP status, and cash status)
    			EMReadScreen MAXIS_case_number, 8, row, 6
    			MAXIS_case_number = trim(MAXIS_case_number)
    			EMReadScreen SNAP_status, 1, row, 45
    			EMReadScreen cash_status, 1, row, 39
                EmReadscreen HC_status, 1, row, 49
				EMReadScreen MAGI_status, 4, row, 55
				EMReadScreen recvd_date, 8, row, 62
				EMReadScreen intvw_date, 8, row, 72

    			'Navigates though until it runs out of case numbers to read
    			IF MAXIS_case_number = "" then exit do

    			'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
    			If cash_status = "-" 	then cash_status = ""
    			If SNAP_status = "-" 	then SNAP_status = ""
    			If HC_status = "-" 		then HC_status = ""

				ReDim Preserve review_array(notes_const, recert_cases)		'resizing the array

				'Adding the case information to the array
				review_array(worker_const, recert_cases) = worker
				review_array(case_number_const, recert_cases) = trim(MAXIS_case_number)
				review_array(CASH_revw_status_const, recert_cases) = cash_status
				review_array(SNAP_revw_status_const, recert_cases) = SNAP_status
				review_array(HC_revw_status_const, recert_cases) = HC_status
				review_array(HC_MAGI_code_const, recert_cases) = trim(MAGI_status)
				review_array(review_recvd_const, recert_cases) = replace(recvd_date, " ", "/")
				If review_array(review_recvd_const, recert_cases) = "__/__/__" Then review_array(review_recvd_const, recert_cases) = ""
				review_array(interview_date_const, recert_cases) = replace(intvw_date, " ", "/")
				If review_array(interview_date_const, recert_cases) = "__/__/__" Then review_array(interview_date_const, recert_cases) = ""
				review_array(saved_to_excel_const, recert_cases) = FALSE

                recert_cases = recert_cases + 1
				STATS_counter = STATS_counter + 1						'adds one instance to the stats counter

    			row = row + 1    'On the next loop it must look to the next row
    			MAXIS_case_number = "" 'Clearing variables before next loop
    		Loop until row = 19		'Last row in REPT/REVS
    		'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
    		PF8
    		EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
            'if max reviews are reached, the goes to next worker is applicable
    	Loop until last_page_check = "THIS IS THE LAST PAGE"
    next
	Call back_to_SELF


	'Now we are going to look at the Excel spreadsheet that has all of the reviews saved.
	excel_row = "2"		'starts at row 2'
	Do
		case_number_to_check = trim(ObjExcel.Cells(excel_row, 2).Value)			'getting the case number from the spreadsheet
		found_in_array = FALSE													'variale to identify if we have found this case in our array
		MAXIS_case_number = case_number_to_check		'setting the case number for NAV functions
		'Here we look through the entire array until we find a match
		For revs_item = 0 to UBound(review_array, 2)
			If review_array(saved_to_excel_const, revs_item) = FALSE Then
				If case_number_to_check = review_array(case_number_const, revs_item) Then		'if the case numbers match we have found our case.
					'Entering information from the array into the excel spreadsheet
					If review_array(CASH_revw_status_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, cash_stat_excel_col).Value = review_array(CASH_revw_status_const, revs_item)
					If review_array(SNAP_revw_status_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, snap_stat_excel_col).Value = review_array(SNAP_revw_status_const, revs_item)
					If review_array(HC_revw_status_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, hc_stat_excel_col).Value = review_array(HC_revw_status_const, revs_item)
					If review_array(HC_MAGI_code_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, magi_stat_excel_col).Value = review_array(HC_MAGI_code_const, revs_item)
					If review_array(review_recvd_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, recvd_date_excel_col).Value = review_array(review_recvd_const, revs_item)
					If review_array(interview_date_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, intvw_date_excel_col).Value = review_array(interview_date_const, revs_item)
					found_in_array = TRUE			'this lets the script know that this case was found in the array
					review_array(saved_to_excel_const, revs_item) = TRUE

					Call add_autoclose_case_note(review_array(CASH_revw_status_const, revs_item), review_array(SNAP_revw_status_const, revs_item), review_array(HC_revw_status_const, revs_item), review_array(review_recvd_const, revs_item), review_array(interview_date_const, revs_item))

					Exit For						'if we found a match, we should stop looking
				End If
			End If
		Next
		'if the case was not found in the array, we need to look in STAT for the information
		If found_in_array = FALSE AND case_number_to_check <> "" Then
			Call check_for_MAXIS(FALSE)		'making sure we haven't passworded out

			' MAXIS_case_number = case_number_to_check		'setting the case number for NAV functions
			call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)		'Go to STAT REVW and be sure the case is not privleged.
			If is_this_priv = FALSE Then
				EMReadScreen recvd_date, 8, 13, 37										'Reading the CAF Received Date and format
				recvd_date = replace(recvd_date, " ", "/")
				if recvd_date = "__/__/__" then recvd_date = ""

				EMReadScreen interview_date, 8, 15, 37									'Reading the interview date and format
				interview_date = replace(interview_date, " ", "/")
				if interview_date = "__/__/__" then interview_date = ""

				EMReadScreen cash_review_status, 1, 7, 40								'Reading the review status and format
				EMReadScreen snap_review_status, 1, 7, 60
				EMReadScreen hc_review_status, 1, 7, 73
				If cash_review_status = "_" Then cash_review_status = ""
				If snap_review_status = "_" Then snap_review_status = ""
				If hc_review_status = "_" Then hc_review_status = ""

				If cash_review_status <> "" Then ObjExcel.Cells(excel_row, cash_stat_excel_col).Value = cash_review_status		'Enter all the information into Excel
				If snap_review_status <> "" Then ObjExcel.Cells(excel_row, snap_stat_excel_col).Value = snap_review_status
				If hc_review_status <> "" Then ObjExcel.Cells(excel_row, hc_stat_excel_col).Value = hc_review_status
				If recvd_date <> "" Then ObjExcel.Cells(excel_row, recvd_date_excel_col).Value = recvd_date
				If interview_date <> "" Then ObjExcel.Cells(excel_row, intvw_date_excel_col).Value = interview_date

				Call add_autoclose_case_note(cash_review_status, snap_review_status, hc_review_status, recvd_date, interview_date)

			End If

			Call back_to_SELF		'Back out in case we need to look into another case.
		End If
		excel_row = excel_row + 1		'going to the next excel
	Loop until case_number_to_check = ""
	excel_row = excel_row - 1
	'Now we will check for any cases that have been ADDED to REVS since we created the report or last ran statistics.
	For revs_item = 0 to UBound(review_array, 2)
		If review_array(saved_to_excel_const, revs_item) = FALSE Then
			Call check_for_MAXIS(FALSE)		'making sure we haven't passworded out
			MAXIS_case_number = review_array(case_number_const, revs_item)

			review_array(interview_const,       revs_item) = False   'values defaulted to False
			review_array(no_interview_const,    revs_item) = False
			review_array(current_SR_const,      revs_item) = False

			Call read_case_details_for_review_report(revs_item)

			'----------------------------------------------------------------------------------------------------Excel Output
			ObjExcel.Cells(excel_row,  1).value = review_array(worker_const,      	  revs_item)     'COL A
			ObjExcel.Cells(excel_row,  2).value = review_array(case_number_const,     revs_item)     'COL B

			ObjExcel.Cells(excel_row,  3).value = review_array(interview_const,       revs_item)     'COL C
			ObjExcel.Cells(excel_row,  4).value = review_array(no_interview_const,    revs_item)     'COL D
			ObjExcel.Cells(excel_row,  5).value = review_array(current_SR_const,      revs_item)     'COL E
			ObjExcel.Cells(excel_row,  6).value = review_array(MFIP_status_const,     revs_item)     'COL F
			ObjExcel.Cells(excel_row,  7).value = review_array(DWP_status_const,      revs_item)     'COL G
			ObjExcel.Cells(excel_row,  8).value = review_array(GA_status_const,       revs_item)     'COL H
			ObjExcel.Cells(excel_row,  9).value = review_array(MSA_status_const,      revs_item)     'COL I
			ObjExcel.Cells(excel_row, 10).value = review_array(GRH_status_const,      revs_item)     'COL J
			ObjExcel.Cells(excel_row, 11).value = review_array(CASH_next_SR_const,    revs_item)     'COL K
			ObjExcel.Cells(excel_row, 12).value = review_array(CASH_next_ER_const,    revs_item)     'COL L
			ObjExcel.Cells(excel_row, 13).value = review_array(SNAP_status_const,     revs_item)     'COL M
			ObjExcel.Cells(excel_row, 14).value = review_array(SNAP_next_SR_const,    revs_item)     'COL N
			ObjExcel.Cells(excel_row, 15).value = review_array(SNAP_next_ER_const,    revs_item)     'COL O
			ObjExcel.Cells(excel_row, 16).value = review_array(MA_status_const,       revs_item)     'COL P
			ObjExcel.Cells(excel_row, 17).value = review_array(MSP_status_const,      revs_item)     'COL Q
			ObjExcel.Cells(excel_row, 18).value = review_array(HC_next_SR_const,      revs_item)     'COL R
			ObjExcel.Cells(excel_row, 19).value = review_array(HC_next_ER_const,      revs_item)     'COL S
			ObjExcel.Cells(excel_row, 20).value = review_array(Language_const,        revs_item)     'COL T
			ObjExcel.Cells(excel_row, 21).value = review_array(Interpreter_const,     revs_item)     'COL U
			ObjExcel.Cells(excel_row, 22).value = review_array(phone_1_const,         revs_item)     'COL V
			ObjExcel.Cells(excel_row, 23).value = review_array(phone_2_const,         revs_item)     'COL W
			ObjExcel.Cells(excel_row, 24).value = review_array(phone_3_const,         revs_item)     'COL X
			ObjExcel.Cells(excel_row, 25).value = review_array(notes_const,           revs_item)     'COL Y

			'Entering information from the array into the excel spreadsheet
			If review_array(CASH_revw_status_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, cash_stat_excel_col).Value = review_array(CASH_revw_status_const, revs_item)
			If review_array(SNAP_revw_status_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, snap_stat_excel_col).Value = review_array(SNAP_revw_status_const, revs_item)
			If review_array(HC_revw_status_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, hc_stat_excel_col).Value = review_array(HC_revw_status_const, revs_item)
			If review_array(HC_MAGI_code_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, magi_stat_excel_col).Value = review_array(HC_MAGI_code_const, revs_item)
			If review_array(review_recvd_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, recvd_date_excel_col).Value = review_array(review_recvd_const, revs_item)
			If review_array(interview_date_const, revs_item) <> "" Then ObjExcel.Cells(excel_row, intvw_date_excel_col).Value = review_array(interview_date_const, revs_item)
			ObjExcel.Range(ObjExcel.Cells(excel_row, 1), ObjExcel.Cells(excel_row, intvw_date_excel_col)).Interior.ColorIndex = 6
			Call back_to_SELF		'Back out in case we need to look into another case.

			Call add_autoclose_case_note(review_array(CASH_revw_status_const, revs_item), review_array(SNAP_revw_status_const, revs_item), review_array(HC_revw_status_const, revs_item), review_array(review_recvd_const, revs_item), review_array(interview_date_const, revs_item))

			excel_row = excel_row + 1		'going to the next excel
		End If
	Next

	'Going to another sheet, to enter statistics for the entire Renewal load
	sheet_name = "Statistics from " & date_month & "-" & date_day
	ObjExcel.Worksheets.Add().Name = sheet_name

	'HERE WE open a new Excel document so we can save the stats counts for uyse in report outs
	'The Review Report Excel will use Excel formulas and the data on the main Review Report Table to calculate the counts and percentages
	'The new Excel - the Stats Excel will just pull the value from the calculation that Excel completes

	'Opening the Excel file
	Set objStatsExcel = CreateObject("Excel.Application")
	objStatsExcel.Visible = True
	Set objStatsWorkbook = objStatsExcel.Workbooks.Add()
	objStatsExcel.DisplayAlerts = FALSE		'This is off because we are saving over an existing file and Excel will alert and we need those to be off

	'This function just enters the header information - which is exactly the same for each sheet
	Call enter_excel_headers(ObjExcel)
	Call enter_excel_headers(objStatsExcel)

	'Now we enter all theformulas and counts in the stats areas
	ObjExcel.Cells(t1_er_w_intv, t1_all_count).Value 				= "=COUNTIFS(Table1[Interview ER],"&is_true&")"
	ObjExcel.Cells(t1_er_w_intv, t1_apps_revd_count).Value 			= "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(t1_er_w_intv, t1_no_app_count).Value 			= "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(t1_er_w_intv, t1_app_percent_count).Value 		= "=" & t1_apps_revd_count_letter & t1_er_w_intv &"/" & t1_all_count_letter & t1_er_w_intv
	ObjExcel.Cells(t1_er_w_intv, t1_app_percent_count).NumberFormat = "0.00%"
	ObjExcel.Cells(t1_er_w_intv, t1_interview_count).Value 			= "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[Intvw Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(t1_er_w_intv, t1_no_intvw_count).Value 			= "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[Intvw Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(t1_er_w_intv, t1_intvw_percent).Value 			= "=" & t1_interview_count_letter & t1_er_w_intv &"/" & t1_all_count_letter & t1_er_w_intv
	ObjExcel.Cells(t1_er_w_intv, t1_intvw_percent).NumberFormat 	= "0.00%"

	ObjExcel.Cells(t1_er_no_intv, t1_all_count).Value 				= "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&")"
	ObjExcel.Cells(t1_er_no_intv, t1_apps_revd_count).Value 		= "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(t1_er_no_intv, t1_no_app_count).Value 			= "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(t1_er_no_intv, t1_app_percent_count).Value 		= "=" & t1_apps_revd_count_letter & t1_er_no_intv &"/" & t1_all_count_letter & t1_er_no_intv
	ObjExcel.Cells(t1_er_no_intv, t1_app_percent_count).NumberFormat = "0.00%"

	ObjExcel.Cells(t1_csr, t1_all_count).Value 						= "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&")"
	ObjExcel.Cells(t1_csr, t1_apps_revd_count).Value 				= "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(t1_csr, t1_no_app_count).Value 					= "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(t1_csr, t1_app_percent_count).Value 				= "=" & t1_apps_revd_count_letter & t1_csr &"/" & t1_all_count_letter & t1_csr
	ObjExcel.Cells(t1_csr, t1_app_percent_count).NumberFormat 		= "0.00%"

	ObjExcel.Cells(t1_priv, t1_all_count).Value 					= "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&")"
	ObjExcel.Cells(t1_priv, t1_apps_revd_count).Value 				= "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CAF Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(t1_priv, t1_no_app_count).Value 					= "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CAF Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(t1_priv, t1_app_percent_count).Value 			= "=" & t1_apps_revd_count_letter & t1_priv &"/" & t1_all_count_letter & t1_priv
	ObjExcel.Cells(t1_priv, t1_app_percent_count).NumberFormat 		= "0.00%"
	ObjExcel.Cells(t1_priv, t1_interview_count).Value 				= "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[Intvw Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(t1_priv, t1_no_intvw_count).Value 				= "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[Intvw Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(t1_priv, t1_intvw_percent).Value 				= "=" & t1_interview_count_letter & t1_priv &"/" & t1_all_count_letter & t1_priv
	ObjExcel.Cells(t1_priv, t1_intvw_percent).NumberFormat 			= "0.00%"

	ObjExcel.Cells(t1_total, t1_all_count).Value 					= "=COUNTA(Table1[Case number])"
	ObjExcel.Cells(t1_total, t1_apps_revd_count).Value 				= "=COUNTIFS(Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(t1_total, t1_no_app_count).Value 				= "=COUNTIFS(Table1[CAF Date ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(t1_total, t1_app_percent_count).Value 			= "=" & t1_apps_revd_count_letter & t1_total &"/" & t1_all_count_letter & t1_total
	ObjExcel.Cells(t1_total, t1_app_percent_count).NumberFormat 	= "0.00%"
	ObjExcel.Cells(t1_total, t1_interview_count).Value 				= "=COUNTIFS(Table1[Intvw Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(t1_total, t1_no_intvw_count).Value 				= "=COUNTIFS(Table1[Intvw Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(t1_total, t1_intvw_percent).Value 				= "=" & t1_interview_count_letter & t1_total &"/" & t1_all_count_letter & t1_total
	ObjExcel.Cells(t1_total, t1_intvw_percent).NumberFormat 		= "0.00%"

	objStatsExcel.Cells(t1_er_w_intv, t1_all_count).Value 					= ObjExcel.Cells(t1_er_w_intv, t1_all_count).Value
	objStatsExcel.Cells(t1_er_w_intv, t1_apps_revd_count).Value 			= ObjExcel.Cells(t1_er_w_intv, t1_apps_revd_count).Value
	objStatsExcel.Cells(t1_er_w_intv, t1_no_app_count).Value 				= ObjExcel.Cells(t1_er_w_intv, t1_no_app_count).Value
	objStatsExcel.Cells(t1_er_w_intv, t1_app_percent_count).Value 			= ObjExcel.Cells(t1_er_w_intv, t1_app_percent_count).Value
	objStatsExcel.Cells(t1_er_w_intv, t1_app_percent_count).NumberFormat 	= "0.00%"
	objStatsExcel.Cells(t1_er_w_intv, t1_interview_count).Value 			= ObjExcel.Cells(t1_er_w_intv, t1_interview_count).Value
	objStatsExcel.Cells(t1_er_w_intv, t1_no_intvw_count).Value 				= ObjExcel.Cells(t1_er_w_intv, t1_no_intvw_count).Value
	objStatsExcel.Cells(t1_er_w_intv, t1_intvw_percent).Value 				= ObjExcel.Cells(t1_er_w_intv, t1_intvw_percent).Value
	objStatsExcel.Cells(t1_er_w_intv, t1_intvw_percent).NumberFormat 		= "0.00%"

	objStatsExcel.Cells(t1_er_no_intv, t1_all_count).Value 					= ObjExcel.Cells(t1_er_no_intv, t1_all_count).Value
	objStatsExcel.Cells(t1_er_no_intv, t1_apps_revd_count).Value 			= ObjExcel.Cells(t1_er_no_intv, t1_apps_revd_count).Value
	objStatsExcel.Cells(t1_er_no_intv, t1_no_app_count).Value 				= ObjExcel.Cells(t1_er_no_intv, t1_no_app_count).Value
	objStatsExcel.Cells(t1_er_no_intv, t1_app_percent_count).Value 			= ObjExcel.Cells(t1_er_no_intv, t1_app_percent_count).Value
	objStatsExcel.Cells(t1_er_no_intv, t1_app_percent_count).NumberFormat 	= "0.00%"

	objStatsExcel.Cells(t1_csr, t1_all_count).Value 						= ObjExcel.Cells(t1_csr, t1_all_count).Value
	objStatsExcel.Cells(t1_csr, t1_apps_revd_count).Value 					= ObjExcel.Cells(t1_csr, t1_apps_revd_count).Value
	objStatsExcel.Cells(t1_csr, t1_no_app_count).Value 						= ObjExcel.Cells(t1_csr, t1_no_app_count).Value
	objStatsExcel.Cells(t1_csr, t1_app_percent_count).Value 				= ObjExcel.Cells(t1_csr, t1_app_percent_count).Value
	objStatsExcel.Cells(t1_csr, t1_app_percent_count).NumberFormat 			= "0.00%"

	objStatsExcel.Cells(t1_priv, t1_all_count).Value 						= ObjExcel.Cells(t1_priv, t1_all_count).Value
	objStatsExcel.Cells(t1_priv, t1_apps_revd_count).Value 					= ObjExcel.Cells(t1_priv, t1_apps_revd_count).Value
	objStatsExcel.Cells(t1_priv, t1_no_app_count).Value 					= ObjExcel.Cells(t1_priv, t1_no_app_count).Value
	objStatsExcel.Cells(t1_priv, t1_app_percent_count).Value 				= ObjExcel.Cells(t1_priv, t1_app_percent_count).Value
	objStatsExcel.Cells(t1_priv, t1_app_percent_count).NumberFormat 		= "0.00%"
	objStatsExcel.Cells(t1_priv, t1_interview_count).Value 					= ObjExcel.Cells(t1_priv, t1_interview_count).Value
	objStatsExcel.Cells(t1_priv, t1_no_intvw_count).Value 					= ObjExcel.Cells(t1_priv, t1_no_intvw_count).Value
	objStatsExcel.Cells(t1_priv, t1_intvw_percent).Value 					= ObjExcel.Cells(t1_priv, t1_intvw_percent).Value
	objStatsExcel.Cells(t1_priv, t1_intvw_percent).NumberFormat 			= "0.00%"

	objStatsExcel.Cells(t1_total, t1_all_count).Value 						= ObjExcel.Cells(t1_total, t1_all_count).Value
	objStatsExcel.Cells(t1_total, t1_apps_revd_count).Value 				= ObjExcel.Cells(t1_total, t1_apps_revd_count).Value
	objStatsExcel.Cells(t1_total, t1_no_app_count).Value 					= ObjExcel.Cells(t1_total, t1_no_app_count).Value
	objStatsExcel.Cells(t1_total, t1_app_percent_count).Value 				= ObjExcel.Cells(t1_total, t1_app_percent_count).Value
	objStatsExcel.Cells(t1_total, t1_app_percent_count).NumberFormat 		= "0.00%"
	objStatsExcel.Cells(t1_total, t1_interview_count).Value 				= ObjExcel.Cells(t1_total, t1_interview_count).Value
	objStatsExcel.Cells(t1_total, t1_no_intvw_count).Value 					= ObjExcel.Cells(t1_total, t1_no_intvw_count).Value
	objStatsExcel.Cells(t1_total, t1_intvw_percent).Value 					= ObjExcel.Cells(t1_total, t1_intvw_percent).Value
	objStatsExcel.Cells(t1_total, t1_intvw_percent).NumberFormat 			= "0.00%"

	stats_row = 2
	Do
		ObjExcel.Cells(stats_row, t2_er_w_intv_app_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[Intvw Date ("&date_header&")], " & t2_dates_letter &stats_row&")"
		ObjExcel.Cells(stats_row, t2_er_w_intv_intv_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CAF Date ("&date_header&")], " & t2_dates_letter &stats_row&")"
		ObjExcel.Cells(stats_row, t2_er_no_intv_app_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_false&",Table1[No Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")], " & t2_dates_letter &stats_row&")"
		ObjExcel.Cells(stats_row, t2_csr_app_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CAF Date ("&date_header&")], " & t2_dates_letter &stats_row&")"
		ObjExcel.Cells(stats_row, t2_priv_intv_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[Intvw Date ("&date_header&")], " & t2_dates_letter &stats_row&")"
		ObjExcel.Cells(stats_row, t2_priv_app_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CAF Date ("&date_header&")], " & t2_dates_letter &stats_row&")"
		ObjExcel.Cells(stats_row, t2_total_intv_count).Value = "=COUNTIFS(Table1[Intvw Date ("&date_header&")], " & t2_dates_letter &stats_row&")"
		ObjExcel.Cells(stats_row, t2_total_app_count).Value = "=COUNTIFS(Table1[CAF Date ("&date_header&")], " & t2_dates_letter &stats_row&")"
		stats_row = stats_row + 1
		next_row_date = ObjExcel.Cells(stats_row, t2_dates).Value
	Loop until next_row_date = ""
	ObjExcel.columns(t2_dates).AutoFit()
	last_row = stats_row - 1

	stats_row = 2
	Do
		objStatsExcel.Cells(stats_row, t2_er_w_intv_app_count).Value = ObjExcel.Cells(stats_row, t2_er_w_intv_app_count).Value
		objStatsExcel.Cells(stats_row, t2_er_w_intv_intv_count).Value = ObjExcel.Cells(stats_row, t2_er_w_intv_intv_count).Value
		objStatsExcel.Cells(stats_row, t2_er_no_intv_app_count).Value = ObjExcel.Cells(stats_row, t2_er_no_intv_app_count).Value
		objStatsExcel.Cells(stats_row, t2_csr_app_count).Value = ObjExcel.Cells(stats_row, t2_csr_app_count).Value
		objStatsExcel.Cells(stats_row, t2_priv_intv_count).Value = ObjExcel.Cells(stats_row, t2_priv_intv_count).Value
		objStatsExcel.Cells(stats_row, t2_priv_app_count).Value = ObjExcel.Cells(stats_row, t2_priv_app_count).Value
		objStatsExcel.Cells(stats_row, t2_total_intv_count).Value = ObjExcel.Cells(stats_row, t2_total_intv_count).Value
		objStatsExcel.Cells(stats_row, t2_total_app_count).Value = ObjExcel.Cells(stats_row, t2_total_app_count).Value
		stats_row = stats_row + 1
		next_row_date = ObjExcel.Cells(stats_row, t2_dates).Value
	Loop until next_row_date = ""
	ObjExcel.columns(t2_dates).AutoFit()

	ObjExcel.Cells(2, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&")"
	ObjExcel.Cells(3, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(4, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(2, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(3, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(4, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(2, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(2, t3_intvs_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(3, t3_intvs_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(4, t3_intvs_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(2, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(2, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(3, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(4, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(2, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(2, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(3, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(4, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(2, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(2, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(3, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(4, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(2, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(2, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(3, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(4, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(2, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(2, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(3, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(4, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(2, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(2, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(3, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(4, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(2, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(2, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(3, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(4, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(2, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "2/" & t3_totals_count_letter & "2"
	ObjExcel.Cells(3, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "3/" & t3_totals_count_letter & "3"
	ObjExcel.Cells(4, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "4/" & t3_totals_count_letter & "4"
	ObjExcel.Cells(2, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(3, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(4, t3_revw_d_percent).NumberFormat = "0.00%"

	ObjExcel.Cells(5, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&")"
	ObjExcel.Cells(6, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(7, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(5, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(6, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(7, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(5, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "5/" & t3_totals_count_letter & "5"
	ObjExcel.Cells(6, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "6/" & t3_totals_count_letter & "6"
	ObjExcel.Cells(7, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "7/" & t3_totals_count_letter & "7"
	ObjExcel.Cells(5, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(6, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(7, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(5, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(6, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(7, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(5, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "5/" & t3_totals_count_letter & "5"
	ObjExcel.Cells(6, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "6/" & t3_totals_count_letter & "6"
	ObjExcel.Cells(7, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "7/" & t3_totals_count_letter & "7"
	ObjExcel.Cells(5, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(6, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(7, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(5, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(6, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(7, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(5, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "5/" & t3_totals_count_letter & "5"
	ObjExcel.Cells(6, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "6/" & t3_totals_count_letter & "6"
	ObjExcel.Cells(7, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "7/" & t3_totals_count_letter & "7"
	ObjExcel.Cells(5, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(6, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(7, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(5, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(6, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(7, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(5, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "5/" & t3_totals_count_letter & "5"
	ObjExcel.Cells(6, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "6/" & t3_totals_count_letter & "6"
	ObjExcel.Cells(7, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "7/" & t3_totals_count_letter & "7"
	ObjExcel.Cells(5, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(6, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(7, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(5, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(6, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(7, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(5, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "5/" & t3_totals_count_letter & "5"
	ObjExcel.Cells(6, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "6/" & t3_totals_count_letter & "6"
	ObjExcel.Cells(7, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "7/" & t3_totals_count_letter & "7"
	ObjExcel.Cells(5, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(6, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(7, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(5, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(6, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(7, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(5, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "5/" & t3_totals_count_letter & "5"
	ObjExcel.Cells(6, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "6/" & t3_totals_count_letter & "6"
	ObjExcel.Cells(7, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "7/" & t3_totals_count_letter & "7"
	ObjExcel.Cells(5, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(6, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(7, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(5, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(6, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(7, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(5, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "5/" & t3_totals_count_letter & "5"
	ObjExcel.Cells(6, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "6/" & t3_totals_count_letter & "6"
	ObjExcel.Cells(7, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "7/" & t3_totals_count_letter & "7"
	ObjExcel.Cells(5, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(6, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(7, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(5, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(6, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(7, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(5, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "5/" & t3_totals_count_letter & "5"
	ObjExcel.Cells(6, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "6/" & t3_totals_count_letter & "6"
	ObjExcel.Cells(7, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "7/" & t3_totals_count_letter & "7"
	ObjExcel.Cells(5, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(6, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(7, t3_revw_d_percent).NumberFormat = "0.00%"

	ObjExcel.Cells(8, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&")"
	ObjExcel.Cells(9, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(10, t3_totals_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(8, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(9, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(10, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(8, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "8/" & t3_totals_count_letter & "8"
	ObjExcel.Cells(9, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "9/" & t3_totals_count_letter & "9"
	ObjExcel.Cells(10, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "10/" & t3_totals_count_letter & "10"
	ObjExcel.Cells(8, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(9, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(10, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(8, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&_
	",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(9, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(10, t3_revw_i_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(8, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "8/" & t3_totals_count_letter & "8"
	ObjExcel.Cells(9, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "9/" & t3_totals_count_letter & "9"
	ObjExcel.Cells(10, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "10/" & t3_totals_count_letter & "10"
	ObjExcel.Cells(8, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(9, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(10, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(8, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(9, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(10, t3_revw_u_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(8, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "8/" & t3_totals_count_letter & "8"
	ObjExcel.Cells(9, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "9/" & t3_totals_count_letter & "9"
	ObjExcel.Cells(10, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "10/" & t3_totals_count_letter & "10"
	ObjExcel.Cells(8, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(9, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(10, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(8, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(9, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(10, t3_revw_n_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(8, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "8/" & t3_totals_count_letter & "8"
	ObjExcel.Cells(9, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "9/" & t3_totals_count_letter & "9"
	ObjExcel.Cells(10, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "10/" & t3_totals_count_letter & "10"
	ObjExcel.Cells(8, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(9, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(10, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(8, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(9, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(10, t3_revw_a_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(8, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "8/" & t3_totals_count_letter & "8"
	ObjExcel.Cells(9, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "9/" & t3_totals_count_letter & "9"
	ObjExcel.Cells(10, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "10/" & t3_totals_count_letter & "10"
	ObjExcel.Cells(8, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(9, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(10, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(8, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(9, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(10, t3_revw_o_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(8, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "8/" & t3_totals_count_letter & "8"
	ObjExcel.Cells(9, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "9/" & t3_totals_count_letter & "9"
	ObjExcel.Cells(10, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "10/" & t3_totals_count_letter & "10"
	ObjExcel.Cells(8, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(9, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(10, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(8, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(9, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(10, t3_revw_t_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(8, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "8/" & t3_totals_count_letter & "8"
	ObjExcel.Cells(9, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "9/" & t3_totals_count_letter & "9"
	ObjExcel.Cells(10, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "10/" & t3_totals_count_letter & "10"
	ObjExcel.Cells(8, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(9, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(10, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(8, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(9, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(10, t3_revw_d_count).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(8, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "8/" & t3_totals_count_letter & "8"
	ObjExcel.Cells(9, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "9/" & t3_totals_count_letter & "9"
	ObjExcel.Cells(10, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "10/" & t3_totals_count_letter & "10"
	ObjExcel.Cells(8, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(9, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(10, t3_revw_d_percent).NumberFormat = "0.00%"

	ObjExcel.Cells(11, t3_totals_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&")"
	ObjExcel.Cells(12, t3_totals_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(13, t3_totals_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(11, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(12, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(13, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(11, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(11, t3_intvs_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(12, t3_intvs_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(13, t3_intvs_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(11, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(11, t3_revw_i_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(12, t3_revw_i_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(13, t3_revw_i_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(11, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(11, t3_revw_u_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(12, t3_revw_u_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(13, t3_revw_u_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(11, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(11, t3_revw_n_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(12, t3_revw_n_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(13, t3_revw_n_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(11, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(11, t3_revw_a_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(12, t3_revw_a_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(13, t3_revw_a_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(11, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(11, t3_revw_o_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(12, t3_revw_o_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(13, t3_revw_o_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(11, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(11, t3_revw_t_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(12, t3_revw_t_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(13, t3_revw_t_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(11, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(11, t3_revw_d_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(12, t3_revw_d_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(13, t3_revw_d_count).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(11, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "11/" & t3_totals_count_letter & "11"
	ObjExcel.Cells(12, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "12/" & t3_totals_count_letter & "12"
	ObjExcel.Cells(13, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "13/" & t3_totals_count_letter & "13"
	ObjExcel.Cells(11, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(12, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(13, t3_revw_d_percent).NumberFormat = "0.00%"

	ObjExcel.Cells(14, t3_totals_count).Value = "=COUNTA(Table1[Case number])"
	ObjExcel.Cells(15, t3_totals_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(16, t3_totals_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(14, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(15, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(16, t3_apps_recvd_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(14, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_apps_recvd_percent).Value = "=" & t3_apps_recvd_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_apps_recvd_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(14, t3_intvs_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(15, t3_intvs_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(16, t3_intvs_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(14, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_intvs_percent).Value = "=" & t3_intvs_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_intvs_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(14, t3_revw_i_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(15, t3_revw_i_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(16, t3_revw_i_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(14, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_revw_i_percent).Value = "=" & t3_revw_i_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_revw_i_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(14, t3_revw_u_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(15, t3_revw_u_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(16, t3_revw_u_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(14, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_revw_u_percent).Value = "=" & t3_revw_u_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_revw_u_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(14, t3_revw_n_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(15, t3_revw_n_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(16, t3_revw_n_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(14, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_revw_n_percent).Value = "=" & t3_revw_n_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_revw_n_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(14, t3_revw_a_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(15, t3_revw_a_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(16, t3_revw_a_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(14, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_revw_a_percent).Value = "=" & t3_revw_a_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_revw_a_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(14, t3_revw_o_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(15, t3_revw_o_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(16, t3_revw_o_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(14, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_revw_o_percent).Value = "=" & t3_revw_o_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_revw_o_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(14, t3_revw_t_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(15, t3_revw_t_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(16, t3_revw_t_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(14, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_revw_t_percent).Value = "=" & t3_revw_t_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_revw_t_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(14, t3_revw_d_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(15, t3_revw_d_count).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(16, t3_revw_d_count).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(14, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "14/" & t3_totals_count_letter & "14"
	ObjExcel.Cells(15, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "15/" & t3_totals_count_letter & "15"
	ObjExcel.Cells(16, t3_revw_d_percent).Value = "=" & t3_revw_d_count_letter & "16/" & t3_totals_count_letter & "16"
	ObjExcel.Cells(14, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(15, t3_revw_d_percent).NumberFormat = "0.00%"
	ObjExcel.Cells(16, t3_revw_d_percent).NumberFormat = "0.00%"

	objStatsExcel.Cells(2, t3_totals_count).Value = ObjExcel.Cells(2, t3_totals_count).Value
	objStatsExcel.Cells(3, t3_totals_count).Value = ObjExcel.Cells(3, t3_totals_count).Value
	objStatsExcel.Cells(4, t3_totals_count).Value = ObjExcel.Cells(4, t3_totals_count).Value

	objStatsExcel.Cells(2, t3_apps_recvd_count).Value = ObjExcel.Cells(2, t3_apps_recvd_count).Value
	objStatsExcel.Cells(3, t3_apps_recvd_count).Value = ObjExcel.Cells(3, t3_apps_recvd_count).Value
	objStatsExcel.Cells(4, t3_apps_recvd_count).Value = ObjExcel.Cells(4, t3_apps_recvd_count).Value
	objStatsExcel.Cells(2, t3_apps_recvd_percent).Value = ObjExcel.Cells(2, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(3, t3_apps_recvd_percent).Value = ObjExcel.Cells(3, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(4, t3_apps_recvd_percent).Value = ObjExcel.Cells(4, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(2, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(2, t3_intvs_count).Value = ObjExcel.Cells(2, t3_intvs_count).Value
	objStatsExcel.Cells(3, t3_intvs_count).Value = ObjExcel.Cells(3, t3_intvs_count).Value
	objStatsExcel.Cells(4, t3_intvs_count).Value = ObjExcel.Cells(4, t3_intvs_count).Value
	objStatsExcel.Cells(2, t3_intvs_percent).Value = ObjExcel.Cells(2, t3_intvs_percent).Value
	objStatsExcel.Cells(3, t3_intvs_percent).Value = ObjExcel.Cells(3, t3_intvs_percent).Value
	objStatsExcel.Cells(4, t3_intvs_percent).Value = ObjExcel.Cells(4, t3_intvs_percent).Value
	objStatsExcel.Cells(2, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(2, t3_revw_i_count).Value = ObjExcel.Cells(2, t3_revw_i_count).Value
	objStatsExcel.Cells(3, t3_revw_i_count).Value = ObjExcel.Cells(3, t3_revw_i_count).Value
	objStatsExcel.Cells(4, t3_revw_i_count).Value = ObjExcel.Cells(4, t3_revw_i_count).Value
	objStatsExcel.Cells(2, t3_revw_i_percent).Value = ObjExcel.Cells(2, t3_revw_i_percent).Value
	objStatsExcel.Cells(3, t3_revw_i_percent).Value = ObjExcel.Cells(3, t3_revw_i_percent).Value
	objStatsExcel.Cells(4, t3_revw_i_percent).Value = ObjExcel.Cells(4, t3_revw_i_percent).Value
	objStatsExcel.Cells(2, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(2, t3_revw_u_count).Value = ObjExcel.Cells(2, t3_revw_u_count).Value
	objStatsExcel.Cells(3, t3_revw_u_count).Value = ObjExcel.Cells(3, t3_revw_u_count).Value
	objStatsExcel.Cells(4, t3_revw_u_count).Value = ObjExcel.Cells(4, t3_revw_u_count).Value
	objStatsExcel.Cells(2, t3_revw_u_percent).Value = ObjExcel.Cells(2, t3_revw_u_percent).Value
	objStatsExcel.Cells(3, t3_revw_u_percent).Value = ObjExcel.Cells(3, t3_revw_u_percent).Value
	objStatsExcel.Cells(4, t3_revw_u_percent).Value = ObjExcel.Cells(4, t3_revw_u_percent).Value
	objStatsExcel.Cells(2, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(2, t3_revw_n_count).Value = ObjExcel.Cells(2, t3_revw_n_count).Value
	objStatsExcel.Cells(3, t3_revw_n_count).Value = ObjExcel.Cells(3, t3_revw_n_count).Value
	objStatsExcel.Cells(4, t3_revw_n_count).Value = ObjExcel.Cells(4, t3_revw_n_count).Value
	objStatsExcel.Cells(2, t3_revw_n_percent).Value = ObjExcel.Cells(2, t3_revw_n_percent).Value
	objStatsExcel.Cells(3, t3_revw_n_percent).Value = ObjExcel.Cells(3, t3_revw_n_percent).Value
	objStatsExcel.Cells(4, t3_revw_n_percent).Value = ObjExcel.Cells(4, t3_revw_n_percent).Value
	objStatsExcel.Cells(2, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(2, t3_revw_a_count).Value = ObjExcel.Cells(2, t3_revw_a_count).Value
	objStatsExcel.Cells(3, t3_revw_a_count).Value = ObjExcel.Cells(3, t3_revw_a_count).Value
	objStatsExcel.Cells(4, t3_revw_a_count).Value = ObjExcel.Cells(4, t3_revw_a_count).Value
	objStatsExcel.Cells(2, t3_revw_a_percent).Value = ObjExcel.Cells(2, t3_revw_a_percent).Value
	objStatsExcel.Cells(3, t3_revw_a_percent).Value = ObjExcel.Cells(3, t3_revw_a_percent).Value
	objStatsExcel.Cells(4, t3_revw_a_percent).Value = ObjExcel.Cells(4, t3_revw_a_percent).Value
	objStatsExcel.Cells(2, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(2, t3_revw_o_count).Value = ObjExcel.Cells(2, t3_revw_o_count).Value
	objStatsExcel.Cells(3, t3_revw_o_count).Value = ObjExcel.Cells(3, t3_revw_o_count).Value
	objStatsExcel.Cells(4, t3_revw_o_count).Value = ObjExcel.Cells(4, t3_revw_o_count).Value
	objStatsExcel.Cells(2, t3_revw_o_percent).Value = ObjExcel.Cells(2, t3_revw_o_percent).Value
	objStatsExcel.Cells(3, t3_revw_o_percent).Value = ObjExcel.Cells(3, t3_revw_o_percent).Value
	objStatsExcel.Cells(4, t3_revw_o_percent).Value = ObjExcel.Cells(4, t3_revw_o_percent).Value
	objStatsExcel.Cells(2, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(2, t3_revw_t_count).Value = ObjExcel.Cells(2, t3_revw_t_count).Value
	objStatsExcel.Cells(3, t3_revw_t_count).Value = ObjExcel.Cells(3, t3_revw_t_count).Value
	objStatsExcel.Cells(4, t3_revw_t_count).Value = ObjExcel.Cells(4, t3_revw_t_count).Value
	objStatsExcel.Cells(2, t3_revw_t_percent).Value = ObjExcel.Cells(2, t3_revw_t_percent).Value
	objStatsExcel.Cells(3, t3_revw_t_percent).Value = ObjExcel.Cells(3, t3_revw_t_percent).Value
	objStatsExcel.Cells(4, t3_revw_t_percent).Value = ObjExcel.Cells(4, t3_revw_t_percent).Value
	objStatsExcel.Cells(2, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(2, t3_revw_d_count).Value = ObjExcel.Cells(2, t3_revw_d_count).Value
	objStatsExcel.Cells(3, t3_revw_d_count).Value = ObjExcel.Cells(3, t3_revw_d_count).Value
	objStatsExcel.Cells(4, t3_revw_d_count).Value = ObjExcel.Cells(4, t3_revw_d_count).Value
	objStatsExcel.Cells(2, t3_revw_d_percent).Value = ObjExcel.Cells(2, t3_revw_d_percent).Value
	objStatsExcel.Cells(3, t3_revw_d_percent).Value = ObjExcel.Cells(3, t3_revw_d_percent).Value
	objStatsExcel.Cells(4, t3_revw_d_percent).Value = ObjExcel.Cells(4, t3_revw_d_percent).Value
	objStatsExcel.Cells(2, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(3, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(4, t3_revw_d_percent).NumberFormat = "0.00%"


	objStatsExcel.Cells(5, t3_totals_count).Value = ObjExcel.Cells(5, t3_totals_count).Value
	objStatsExcel.Cells(6, t3_totals_count).Value = ObjExcel.Cells(6, t3_totals_count).Value
	objStatsExcel.Cells(7, t3_totals_count).Value = ObjExcel.Cells(7, t3_totals_count).Value

	objStatsExcel.Cells(5, t3_apps_recvd_count).Value = ObjExcel.Cells(5, t3_apps_recvd_count).Value
	objStatsExcel.Cells(6, t3_apps_recvd_count).Value = ObjExcel.Cells(6, t3_apps_recvd_count).Value
	objStatsExcel.Cells(7, t3_apps_recvd_count).Value = ObjExcel.Cells(7, t3_apps_recvd_count).Value
	objStatsExcel.Cells(5, t3_apps_recvd_percent).Value = ObjExcel.Cells(5, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(6, t3_apps_recvd_percent).Value = ObjExcel.Cells(6, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(7, t3_apps_recvd_percent).Value = ObjExcel.Cells(7, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(5, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(6, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(7, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(5, t3_revw_i_count).Value = ObjExcel.Cells(5, t3_revw_i_count).Value
	objStatsExcel.Cells(6, t3_revw_i_count).Value = ObjExcel.Cells(6, t3_revw_i_count).Value
	objStatsExcel.Cells(7, t3_revw_i_count).Value = ObjExcel.Cells(7, t3_revw_i_count).Value
	objStatsExcel.Cells(5, t3_revw_i_percent).Value = ObjExcel.Cells(5, t3_revw_i_percent).Value
	objStatsExcel.Cells(6, t3_revw_i_percent).Value = ObjExcel.Cells(6, t3_revw_i_percent).Value
	objStatsExcel.Cells(7, t3_revw_i_percent).Value = ObjExcel.Cells(7, t3_revw_i_percent).Value
	objStatsExcel.Cells(5, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(6, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(7, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(5, t3_revw_u_count).Value = ObjExcel.Cells(5, t3_revw_u_count).Value
	objStatsExcel.Cells(6, t3_revw_u_count).Value = ObjExcel.Cells(6, t3_revw_u_count).Value
	objStatsExcel.Cells(7, t3_revw_u_count).Value = ObjExcel.Cells(7, t3_revw_u_count).Value
	objStatsExcel.Cells(5, t3_revw_u_percent).Value = ObjExcel.Cells(5, t3_revw_u_percent).Value
	objStatsExcel.Cells(6, t3_revw_u_percent).Value = ObjExcel.Cells(6, t3_revw_u_percent).Value
	objStatsExcel.Cells(7, t3_revw_u_percent).Value = ObjExcel.Cells(7, t3_revw_u_percent).Value
	objStatsExcel.Cells(5, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(6, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(7, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(5, t3_revw_n_count).Value = ObjExcel.Cells(5, t3_revw_n_count).Value
	objStatsExcel.Cells(6, t3_revw_n_count).Value = ObjExcel.Cells(6, t3_revw_n_count).Value
	objStatsExcel.Cells(7, t3_revw_n_count).Value = ObjExcel.Cells(7, t3_revw_n_count).Value
	objStatsExcel.Cells(5, t3_revw_n_percent).Value = ObjExcel.Cells(5, t3_revw_n_percent).Value
	objStatsExcel.Cells(6, t3_revw_n_percent).Value = ObjExcel.Cells(6, t3_revw_n_percent).Value
	objStatsExcel.Cells(7, t3_revw_n_percent).Value = ObjExcel.Cells(7, t3_revw_n_percent).Value
	objStatsExcel.Cells(5, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(6, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(7, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(5, t3_revw_a_count).Value = ObjExcel.Cells(5, t3_revw_a_count).Value
	objStatsExcel.Cells(6, t3_revw_a_count).Value = ObjExcel.Cells(6, t3_revw_a_count).Value
	objStatsExcel.Cells(7, t3_revw_a_count).Value = ObjExcel.Cells(7, t3_revw_a_count).Value
	objStatsExcel.Cells(5, t3_revw_a_percent).Value = ObjExcel.Cells(5, t3_revw_a_percent).Value
	objStatsExcel.Cells(6, t3_revw_a_percent).Value = ObjExcel.Cells(6, t3_revw_a_percent).Value
	objStatsExcel.Cells(7, t3_revw_a_percent).Value = ObjExcel.Cells(7, t3_revw_a_percent).Value
	objStatsExcel.Cells(5, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(6, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(7, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(5, t3_revw_o_count).Value = ObjExcel.Cells(5, t3_revw_o_count).Value
	objStatsExcel.Cells(6, t3_revw_o_count).Value = ObjExcel.Cells(6, t3_revw_o_count).Value
	objStatsExcel.Cells(7, t3_revw_o_count).Value = ObjExcel.Cells(7, t3_revw_o_count).Value
	objStatsExcel.Cells(5, t3_revw_o_percent).Value = ObjExcel.Cells(5, t3_revw_o_percent).Value
	objStatsExcel.Cells(6, t3_revw_o_percent).Value = ObjExcel.Cells(6, t3_revw_o_percent).Value
	objStatsExcel.Cells(7, t3_revw_o_percent).Value = ObjExcel.Cells(7, t3_revw_o_percent).Value
	objStatsExcel.Cells(5, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(6, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(7, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(5, t3_revw_t_count).Value = ObjExcel.Cells(5, t3_revw_t_count).Value
	objStatsExcel.Cells(6, t3_revw_t_count).Value = ObjExcel.Cells(6, t3_revw_t_count).Value
	objStatsExcel.Cells(7, t3_revw_t_count).Value = ObjExcel.Cells(7, t3_revw_t_count).Value
	objStatsExcel.Cells(5, t3_revw_t_percent).Value = ObjExcel.Cells(5, t3_revw_t_percent).Value
	objStatsExcel.Cells(6, t3_revw_t_percent).Value = ObjExcel.Cells(6, t3_revw_t_percent).Value
	objStatsExcel.Cells(7, t3_revw_t_percent).Value = ObjExcel.Cells(7, t3_revw_t_percent).Value
	objStatsExcel.Cells(5, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(6, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(7, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(5, t3_revw_d_count).Value = ObjExcel.Cells(5, t3_revw_d_count).Value
	objStatsExcel.Cells(6, t3_revw_d_count).Value = ObjExcel.Cells(6, t3_revw_d_count).Value
	objStatsExcel.Cells(7, t3_revw_d_count).Value = ObjExcel.Cells(7, t3_revw_d_count).Value
	objStatsExcel.Cells(5, t3_revw_d_percent).Value = ObjExcel.Cells(5, t3_revw_d_percent).Value
	objStatsExcel.Cells(6, t3_revw_d_percent).Value = ObjExcel.Cells(6, t3_revw_d_percent).Value
	objStatsExcel.Cells(7, t3_revw_d_percent).Value = ObjExcel.Cells(7, t3_revw_d_percent).Value
	objStatsExcel.Cells(5, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(6, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(7, t3_revw_d_percent).NumberFormat = "0.00%"

	objStatsExcel.Cells(8, t3_totals_count).Value = ObjExcel.Cells(8, t3_totals_count).Value
	objStatsExcel.Cells(9, t3_totals_count).Value = ObjExcel.Cells(9, t3_totals_count).Value
	objStatsExcel.Cells(10, t3_totals_count).Value = ObjExcel.Cells(10, t3_totals_count).Value

	objStatsExcel.Cells(8, t3_apps_recvd_count).Value = ObjExcel.Cells(8, t3_apps_recvd_count).Value
	objStatsExcel.Cells(9, t3_apps_recvd_count).Value = ObjExcel.Cells(9, t3_apps_recvd_count).Value
	objStatsExcel.Cells(10, t3_apps_recvd_count).Value = ObjExcel.Cells(10, t3_apps_recvd_count).Value
	objStatsExcel.Cells(8, t3_apps_recvd_percent).Value = ObjExcel.Cells(8, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(9, t3_apps_recvd_percent).Value = ObjExcel.Cells(9, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(10, t3_apps_recvd_percent).Value = ObjExcel.Cells(10, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(8, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(9, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(10, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(8, t3_revw_i_count).Value = ObjExcel.Cells(8, t3_revw_i_count).Value
	objStatsExcel.Cells(9, t3_revw_i_count).Value = ObjExcel.Cells(9, t3_revw_i_count).Value
	objStatsExcel.Cells(10, t3_revw_i_count).Value = ObjExcel.Cells(10, t3_revw_i_count).Value
	objStatsExcel.Cells(8, t3_revw_i_percent).Value = ObjExcel.Cells(8, t3_revw_i_percent).Value
	objStatsExcel.Cells(9, t3_revw_i_percent).Value = ObjExcel.Cells(9, t3_revw_i_percent).Value
	objStatsExcel.Cells(10, t3_revw_i_percent).Value = ObjExcel.Cells(10, t3_revw_i_percent).Value
	objStatsExcel.Cells(8, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(9, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(10, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(8, t3_revw_u_count).Value = ObjExcel.Cells(8, t3_revw_u_count).Value
	objStatsExcel.Cells(9, t3_revw_u_count).Value = ObjExcel.Cells(9, t3_revw_u_count).Value
	objStatsExcel.Cells(10, t3_revw_u_count).Value = ObjExcel.Cells(10, t3_revw_u_count).Value
	objStatsExcel.Cells(8, t3_revw_u_percent).Value = ObjExcel.Cells(8, t3_revw_u_percent).Value
	objStatsExcel.Cells(9, t3_revw_u_percent).Value = ObjExcel.Cells(9, t3_revw_u_percent).Value
	objStatsExcel.Cells(10, t3_revw_u_percent).Value = ObjExcel.Cells(10, t3_revw_u_percent).Value
	objStatsExcel.Cells(8, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(9, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(10, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(8, t3_revw_n_count).Value = ObjExcel.Cells(8, t3_revw_n_count).Value
	objStatsExcel.Cells(9, t3_revw_n_count).Value = ObjExcel.Cells(9, t3_revw_n_count).Value
	objStatsExcel.Cells(10, t3_revw_n_count).Value = ObjExcel.Cells(10, t3_revw_n_count).Value
	objStatsExcel.Cells(8, t3_revw_n_percent).Value = ObjExcel.Cells(8, t3_revw_n_percent).Value
	objStatsExcel.Cells(9, t3_revw_n_percent).Value = ObjExcel.Cells(9, t3_revw_n_percent).Value
	objStatsExcel.Cells(10, t3_revw_n_percent).Value = ObjExcel.Cells(10, t3_revw_n_percent).Value
	objStatsExcel.Cells(8, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(9, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(10, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(8, t3_revw_a_count).Value = ObjExcel.Cells(8, t3_revw_a_count).Value
	objStatsExcel.Cells(9, t3_revw_a_count).Value = ObjExcel.Cells(9, t3_revw_a_count).Value
	objStatsExcel.Cells(10, t3_revw_a_count).Value = ObjExcel.Cells(10, t3_revw_a_count).Value
	objStatsExcel.Cells(8, t3_revw_a_percent).Value = ObjExcel.Cells(8, t3_revw_a_percent).Value
	objStatsExcel.Cells(9, t3_revw_a_percent).Value = ObjExcel.Cells(9, t3_revw_a_percent).Value
	objStatsExcel.Cells(10, t3_revw_a_percent).Value = ObjExcel.Cells(10, t3_revw_a_percent).Value
	objStatsExcel.Cells(8, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(9, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(10, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(8, t3_revw_o_count).Value = ObjExcel.Cells(8, t3_revw_o_count).Value
	objStatsExcel.Cells(9, t3_revw_o_count).Value = ObjExcel.Cells(9, t3_revw_o_count).Value
	objStatsExcel.Cells(10, t3_revw_o_count).Value = ObjExcel.Cells(10, t3_revw_o_count).Value
	objStatsExcel.Cells(8, t3_revw_o_percent).Value = ObjExcel.Cells(8, t3_revw_o_percent).Value
	objStatsExcel.Cells(9, t3_revw_o_percent).Value = ObjExcel.Cells(9, t3_revw_o_percent).Value
	objStatsExcel.Cells(10, t3_revw_o_percent).Value = ObjExcel.Cells(10, t3_revw_o_percent).Value
	objStatsExcel.Cells(8, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(9, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(10, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(8, t3_revw_t_count).Value = ObjExcel.Cells(8, t3_revw_t_count).Value
	objStatsExcel.Cells(9, t3_revw_t_count).Value = ObjExcel.Cells(9, t3_revw_t_count).Value
	objStatsExcel.Cells(10, t3_revw_t_count).Value = ObjExcel.Cells(10, t3_revw_t_count).Value
	objStatsExcel.Cells(8, t3_revw_t_percent).Value = ObjExcel.Cells(8, t3_revw_t_percent).Value
	objStatsExcel.Cells(9, t3_revw_t_percent).Value = ObjExcel.Cells(9, t3_revw_t_percent).Value
	objStatsExcel.Cells(10, t3_revw_t_percent).Value = ObjExcel.Cells(10, t3_revw_t_percent).Value
	objStatsExcel.Cells(8, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(9, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(10, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(8, t3_revw_d_count).Value = ObjExcel.Cells(8, t3_revw_d_count).Value
	objStatsExcel.Cells(9, t3_revw_d_count).Value = ObjExcel.Cells(9, t3_revw_d_count).Value
	objStatsExcel.Cells(10, t3_revw_d_count).Value = ObjExcel.Cells(10, t3_revw_d_count).Value
	objStatsExcel.Cells(8, t3_revw_d_percent).Value = ObjExcel.Cells(8, t3_revw_d_percent).Value
	objStatsExcel.Cells(9, t3_revw_d_percent).Value = ObjExcel.Cells(9, t3_revw_d_percent).Value
	objStatsExcel.Cells(10, t3_revw_d_percent).Value = ObjExcel.Cells(10, t3_revw_d_percent).Value
	objStatsExcel.Cells(8, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(9, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(10, t3_revw_d_percent).NumberFormat = "0.00%"

	objStatsExcel.Cells(11, t3_totals_count).Value = ObjExcel.Cells(11, t3_totals_count).Value
	objStatsExcel.Cells(12, t3_totals_count).Value = ObjExcel.Cells(12, t3_totals_count).Value
	objStatsExcel.Cells(13, t3_totals_count).Value = ObjExcel.Cells(13, t3_totals_count).Value

	objStatsExcel.Cells(11, t3_apps_recvd_count).Value = ObjExcel.Cells(11, t3_apps_recvd_count).Value
	objStatsExcel.Cells(12, t3_apps_recvd_count).Value = ObjExcel.Cells(12, t3_apps_recvd_count).Value
	objStatsExcel.Cells(13, t3_apps_recvd_count).Value = ObjExcel.Cells(13, t3_apps_recvd_count).Value
	objStatsExcel.Cells(11, t3_apps_recvd_percent).Value = ObjExcel.Cells(11, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(12, t3_apps_recvd_percent).Value = ObjExcel.Cells(12, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(13, t3_apps_recvd_percent).Value = ObjExcel.Cells(13, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(11, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(11, t3_intvs_count).Value = ObjExcel.Cells(11, t3_intvs_count).Value
	objStatsExcel.Cells(12, t3_intvs_count).Value = ObjExcel.Cells(12, t3_intvs_count).Value
	objStatsExcel.Cells(13, t3_intvs_count).Value = ObjExcel.Cells(13, t3_intvs_count).Value
	objStatsExcel.Cells(11, t3_intvs_percent).Value = ObjExcel.Cells(11, t3_intvs_percent).Value
	objStatsExcel.Cells(12, t3_intvs_percent).Value = ObjExcel.Cells(12, t3_intvs_percent).Value
	objStatsExcel.Cells(13, t3_intvs_percent).Value = ObjExcel.Cells(13, t3_intvs_percent).Value
	objStatsExcel.Cells(11, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(11, t3_revw_i_count).Value = ObjExcel.Cells(11, t3_revw_i_count).Value
	objStatsExcel.Cells(12, t3_revw_i_count).Value = ObjExcel.Cells(12, t3_revw_i_count).Value
	objStatsExcel.Cells(13, t3_revw_i_count).Value = ObjExcel.Cells(13, t3_revw_i_count).Value
	objStatsExcel.Cells(11, t3_revw_i_percent).Value = ObjExcel.Cells(11, t3_revw_i_percent).Value
	objStatsExcel.Cells(12, t3_revw_i_percent).Value = ObjExcel.Cells(12, t3_revw_i_percent).Value
	objStatsExcel.Cells(13, t3_revw_i_percent).Value = ObjExcel.Cells(13, t3_revw_i_percent).Value
	objStatsExcel.Cells(11, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(11, t3_revw_u_count).Value = ObjExcel.Cells(11, t3_revw_u_count).Value
	objStatsExcel.Cells(12, t3_revw_u_count).Value = ObjExcel.Cells(12, t3_revw_u_count).Value
	objStatsExcel.Cells(13, t3_revw_u_count).Value = ObjExcel.Cells(13, t3_revw_u_count).Value
	objStatsExcel.Cells(11, t3_revw_u_percent).Value = ObjExcel.Cells(11, t3_revw_u_percent).Value
	objStatsExcel.Cells(12, t3_revw_u_percent).Value = ObjExcel.Cells(12, t3_revw_u_percent).Value
	objStatsExcel.Cells(13, t3_revw_u_percent).Value = ObjExcel.Cells(13, t3_revw_u_percent).Value
	objStatsExcel.Cells(11, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(11, t3_revw_n_count).Value = ObjExcel.Cells(11, t3_revw_n_count).Value
	objStatsExcel.Cells(12, t3_revw_n_count).Value = ObjExcel.Cells(12, t3_revw_n_count).Value
	objStatsExcel.Cells(13, t3_revw_n_count).Value = ObjExcel.Cells(13, t3_revw_n_count).Value
	objStatsExcel.Cells(11, t3_revw_n_percent).Value = ObjExcel.Cells(11, t3_revw_n_percent).Value
	objStatsExcel.Cells(12, t3_revw_n_percent).Value = ObjExcel.Cells(12, t3_revw_n_percent).Value
	objStatsExcel.Cells(13, t3_revw_n_percent).Value = ObjExcel.Cells(13, t3_revw_n_percent).Value
	objStatsExcel.Cells(11, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(11, t3_revw_a_count).Value = ObjExcel.Cells(11, t3_revw_a_count).Value
	objStatsExcel.Cells(12, t3_revw_a_count).Value = ObjExcel.Cells(12, t3_revw_a_count).Value
	objStatsExcel.Cells(13, t3_revw_a_count).Value = ObjExcel.Cells(13, t3_revw_a_count).Value
	objStatsExcel.Cells(11, t3_revw_a_percent).Value = ObjExcel.Cells(11, t3_revw_a_percent).Value
	objStatsExcel.Cells(12, t3_revw_a_percent).Value = ObjExcel.Cells(12, t3_revw_a_percent).Value
	objStatsExcel.Cells(13, t3_revw_a_percent).Value = ObjExcel.Cells(13, t3_revw_a_percent).Value
	objStatsExcel.Cells(11, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(11, t3_revw_o_count).Value = ObjExcel.Cells(11, t3_revw_o_count).Value
	objStatsExcel.Cells(12, t3_revw_o_count).Value = ObjExcel.Cells(12, t3_revw_o_count).Value
	objStatsExcel.Cells(13, t3_revw_o_count).Value = ObjExcel.Cells(13, t3_revw_o_count).Value
	objStatsExcel.Cells(11, t3_revw_o_percent).Value = ObjExcel.Cells(11, t3_revw_o_percent).Value
	objStatsExcel.Cells(12, t3_revw_o_percent).Value = ObjExcel.Cells(12, t3_revw_o_percent).Value
	objStatsExcel.Cells(13, t3_revw_o_percent).Value = ObjExcel.Cells(13, t3_revw_o_percent).Value
	objStatsExcel.Cells(11, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(11, t3_revw_t_count).Value = ObjExcel.Cells(11, t3_revw_t_count).Value
	objStatsExcel.Cells(12, t3_revw_t_count).Value = ObjExcel.Cells(12, t3_revw_t_count).Value
	objStatsExcel.Cells(13, t3_revw_t_count).Value = ObjExcel.Cells(13, t3_revw_t_count).Value
	objStatsExcel.Cells(11, t3_revw_t_percent).Value = ObjExcel.Cells(11, t3_revw_t_percent).Value
	objStatsExcel.Cells(12, t3_revw_t_percent).Value = ObjExcel.Cells(12, t3_revw_t_percent).Value
	objStatsExcel.Cells(13, t3_revw_t_percent).Value = ObjExcel.Cells(13, t3_revw_t_percent).Value
	objStatsExcel.Cells(11, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(11, t3_revw_d_count).Value = ObjExcel.Cells(11, t3_revw_d_count).Value
	objStatsExcel.Cells(12, t3_revw_d_count).Value = ObjExcel.Cells(12, t3_revw_d_count).Value
	objStatsExcel.Cells(13, t3_revw_d_count).Value = ObjExcel.Cells(13, t3_revw_d_count).Value
	objStatsExcel.Cells(11, t3_revw_d_percent).Value = ObjExcel.Cells(11, t3_revw_d_percent).Value
	objStatsExcel.Cells(12, t3_revw_d_percent).Value = ObjExcel.Cells(12, t3_revw_d_percent).Value
	objStatsExcel.Cells(13, t3_revw_d_percent).Value = ObjExcel.Cells(13, t3_revw_d_percent).Value
	objStatsExcel.Cells(11, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(12, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(13, t3_revw_d_percent).NumberFormat = "0.00%"

	objStatsExcel.Cells(14, t3_totals_count).Value = ObjExcel.Cells(14, t3_totals_count).Value
	objStatsExcel.Cells(15, t3_totals_count).Value = ObjExcel.Cells(15, t3_totals_count).Value
	objStatsExcel.Cells(16, t3_totals_count).Value = ObjExcel.Cells(16, t3_totals_count).Value

	objStatsExcel.Cells(14, t3_apps_recvd_count).Value = ObjExcel.Cells(14, t3_apps_recvd_count).Value
	objStatsExcel.Cells(15, t3_apps_recvd_count).Value = ObjExcel.Cells(15, t3_apps_recvd_count).Value
	objStatsExcel.Cells(16, t3_apps_recvd_count).Value = ObjExcel.Cells(16, t3_apps_recvd_count).Value
	objStatsExcel.Cells(14, t3_apps_recvd_percent).Value = ObjExcel.Cells(14, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(15, t3_apps_recvd_percent).Value = ObjExcel.Cells(15, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(16, t3_apps_recvd_percent).Value = ObjExcel.Cells(16, t3_apps_recvd_percent).Value
	objStatsExcel.Cells(14, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_apps_recvd_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(14, t3_intvs_count).Value = ObjExcel.Cells(14, t3_intvs_count).Value
	objStatsExcel.Cells(15, t3_intvs_count).Value = ObjExcel.Cells(15, t3_intvs_count).Value
	objStatsExcel.Cells(16, t3_intvs_count).Value = ObjExcel.Cells(16, t3_intvs_count).Value
	objStatsExcel.Cells(14, t3_intvs_percent).Value = ObjExcel.Cells(14, t3_intvs_percent).Value
	objStatsExcel.Cells(15, t3_intvs_percent).Value = ObjExcel.Cells(15, t3_intvs_percent).Value
	objStatsExcel.Cells(16, t3_intvs_percent).Value = ObjExcel.Cells(16, t3_intvs_percent).Value
	objStatsExcel.Cells(14, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_intvs_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(14, t3_revw_i_count).Value = ObjExcel.Cells(14, t3_revw_i_count).Value
	objStatsExcel.Cells(15, t3_revw_i_count).Value = ObjExcel.Cells(15, t3_revw_i_count).Value
	objStatsExcel.Cells(16, t3_revw_i_count).Value = ObjExcel.Cells(16, t3_revw_i_count).Value
	objStatsExcel.Cells(14, t3_revw_i_percent).Value = ObjExcel.Cells(14, t3_revw_i_percent).Value
	objStatsExcel.Cells(15, t3_revw_i_percent).Value = ObjExcel.Cells(15, t3_revw_i_percent).Value
	objStatsExcel.Cells(16, t3_revw_i_percent).Value = ObjExcel.Cells(16, t3_revw_i_percent).Value
	objStatsExcel.Cells(14, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_revw_i_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(14, t3_revw_u_count).Value = ObjExcel.Cells(14, t3_revw_u_count).Value
	objStatsExcel.Cells(15, t3_revw_u_count).Value = ObjExcel.Cells(15, t3_revw_u_count).Value
	objStatsExcel.Cells(16, t3_revw_u_count).Value = ObjExcel.Cells(16, t3_revw_u_count).Value
	objStatsExcel.Cells(14, t3_revw_u_percent).Value = ObjExcel.Cells(14, t3_revw_u_percent).Value
	objStatsExcel.Cells(15, t3_revw_u_percent).Value = ObjExcel.Cells(15, t3_revw_u_percent).Value
	objStatsExcel.Cells(16, t3_revw_u_percent).Value = ObjExcel.Cells(16, t3_revw_u_percent).Value
	objStatsExcel.Cells(14, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_revw_u_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(14, t3_revw_n_count).Value = ObjExcel.Cells(14, t3_revw_n_count).Value
	objStatsExcel.Cells(15, t3_revw_n_count).Value = ObjExcel.Cells(15, t3_revw_n_count).Value
	objStatsExcel.Cells(16, t3_revw_n_count).Value = ObjExcel.Cells(16, t3_revw_n_count).Value
	objStatsExcel.Cells(14, t3_revw_n_percent).Value = ObjExcel.Cells(14, t3_revw_n_percent).Value
	objStatsExcel.Cells(15, t3_revw_n_percent).Value = ObjExcel.Cells(15, t3_revw_n_percent).Value
	objStatsExcel.Cells(16, t3_revw_n_percent).Value = ObjExcel.Cells(16, t3_revw_n_percent).Value
	objStatsExcel.Cells(14, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_revw_n_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(14, t3_revw_a_count).Value = ObjExcel.Cells(14, t3_revw_a_count).Value
	objStatsExcel.Cells(15, t3_revw_a_count).Value = ObjExcel.Cells(15, t3_revw_a_count).Value
	objStatsExcel.Cells(16, t3_revw_a_count).Value = ObjExcel.Cells(16, t3_revw_a_count).Value
	objStatsExcel.Cells(14, t3_revw_a_percent).Value = ObjExcel.Cells(14, t3_revw_a_percent).Value
	objStatsExcel.Cells(15, t3_revw_a_percent).Value = ObjExcel.Cells(15, t3_revw_a_percent).Value
	objStatsExcel.Cells(16, t3_revw_a_percent).Value = ObjExcel.Cells(16, t3_revw_a_percent).Value
	objStatsExcel.Cells(14, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_revw_a_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(14, t3_revw_o_count).Value = ObjExcel.Cells(14, t3_revw_o_count).Value
	objStatsExcel.Cells(15, t3_revw_o_count).Value = ObjExcel.Cells(15, t3_revw_o_count).Value
	objStatsExcel.Cells(16, t3_revw_o_count).Value = ObjExcel.Cells(16, t3_revw_o_count).Value
	objStatsExcel.Cells(14, t3_revw_o_percent).Value = ObjExcel.Cells(14, t3_revw_o_percent).Value
	objStatsExcel.Cells(15, t3_revw_o_percent).Value = ObjExcel.Cells(15, t3_revw_o_percent).Value
	objStatsExcel.Cells(16, t3_revw_o_percent).Value = ObjExcel.Cells(16, t3_revw_o_percent).Value
	objStatsExcel.Cells(14, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_revw_o_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(14, t3_revw_t_count).Value = ObjExcel.Cells(14, t3_revw_t_count).Value
	objStatsExcel.Cells(15, t3_revw_t_count).Value = ObjExcel.Cells(15, t3_revw_t_count).Value
	objStatsExcel.Cells(16, t3_revw_t_count).Value = ObjExcel.Cells(16, t3_revw_t_count).Value
	objStatsExcel.Cells(14, t3_revw_t_percent).Value = ObjExcel.Cells(14, t3_revw_t_percent).Value
	objStatsExcel.Cells(15, t3_revw_t_percent).Value = ObjExcel.Cells(15, t3_revw_t_percent).Value
	objStatsExcel.Cells(16, t3_revw_t_percent).Value = ObjExcel.Cells(16, t3_revw_t_percent).Value
	objStatsExcel.Cells(14, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_revw_t_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(14, t3_revw_d_count).Value = ObjExcel.Cells(14, t3_revw_d_count).Value
	objStatsExcel.Cells(15, t3_revw_d_count).Value = ObjExcel.Cells(15, t3_revw_d_count).Value
	objStatsExcel.Cells(16, t3_revw_d_count).Value = ObjExcel.Cells(16, t3_revw_d_count).Value
	objStatsExcel.Cells(14, t3_revw_d_percent).Value = ObjExcel.Cells(14, t3_revw_d_percent).Value
	objStatsExcel.Cells(15, t3_revw_d_percent).Value = ObjExcel.Cells(15, t3_revw_d_percent).Value
	objStatsExcel.Cells(16, t3_revw_d_percent).Value = ObjExcel.Cells(16, t3_revw_d_percent).Value
	objStatsExcel.Cells(14, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(15, t3_revw_d_percent).NumberFormat = "0.00%"
	objStatsExcel.Cells(16, t3_revw_d_percent).NumberFormat = "0.00%"

	'Resizing all of the columns to autofit
	For xl_col = 1 to t3_totals_count
		ObjExcel.columns(xl_col).AutoFit()
	Next
	For xl_col = 1 to t3_totals_count
		objStatsExcel.columns(xl_col).AutoFit()
	Next

	'Setting the Ranges for each of the 3 tables
	table1Range = t1_REVW_type_letter & "1:" & t1_intvw_percent_letter & "6"
	table2Range = t2_dates_letter & "1:" & t2_total_app_count_letter & last_row
	table3Range = t3_progs_letter & "1:" & t3_totals_count_letter & "16"

	'This is how we make a range of cells into a table.
	'The ADD Method creates the table and Full Detail can be found here - https://docs.microsoft.com/en-us/office/vba/api/excel.listobjects.add
	'This is a fairly simple use of this functionality and probably has a lot more to it than is exemplified here and the methods/properties can be found:
		'https://docs.microsoft.com/en-us/office/vba/api/excel.listobject
	ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table1Range, xlYes).Name = "AppsANDIntvwTable"
	ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table2Range, xlYes).Name = "DateCountTable"
	ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table3Range, xlYes).Name = "REVWStatusTable"

	'xlSrcRange is a constant that needs to be set in the script and is '1' which sets the source of the information to a range of cells for defining the table
		'Other documentation and information about this option can be found here - https://docs.microsoft.com/en-us/office/vba/api/excel.xllistobjectsourcetype
		' XlListObjectSourceType enumeration (Excel)
		' NAME				VALUE		DESCRIPTION
		' xlSrcExternal		0			External data source (Microsoft SharePoint Foundation site).
		' xlSrcModel		4			PowerPivot Model
		' xlSrcQuery		3			Query
		' xlSrcRange		1			Range
		' xlSrcXml			2			XML
	'xlYes is a constant that needs to be set in the script and is '1' which inidcates that the data source has headers.
		'Other documentation and information about this option can be found here - https://docs.microsoft.com/en-us/office/vba/api/excel.xlyesnoguess
		' XLYESNOGUESS ENUMERATION (EXCEL)
		' NAME		VALUE	DESCRIPTION
		' xlGuess	0		Excel determines whether there is a header, and where it is, if there is one.
		' xlNo		2		Default. The entire range should be sorted.
		' xlYes		1		The entire range should not be sorted.

	'This is how you can use the property 'TableStyle' to set the visuals of the table. The words here can be found IN EXCEL by hovering over the table styles to see how they are called - these 3 below work.
	' ObjExcel.ActiveSheet.ListObjects("AppsANDIntvwTable").TableStyle = "TableStyleDark9"
	' ObjExcel.ActiveSheet.ListObjects("DateCountTable").TableStyle = "TableStyleDark10"
	' ObjExcel.ActiveSheet.ListObjects("REVWStatusTable").TableStyle = "TableStyleDark11"

	'Now we make more tables for the other excel
	objStatsExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table1Range, xlYes).Name = "Table1"
	objStatsExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table2Range, xlYes).Name = "Table2"
	objStatsExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table3Range, xlYes).Name = "Table3"

	' 'https://docs.microsoft.com/en-us/office/vba/api/excel.xlcolorindex - This is where you find all the excel numbers to use are.
	ObjExcel.Range(t3_apps_recvd_count_letter 	& "2:" & t3_intvs_percent_letter 		& "2").Interior.ColorIndex = 6
	ObjExcel.Range(t3_revw_a_count_letter 		& "2:" & t3_revw_a_percent_letter 		& "2").Interior.ColorIndex = 6
	ObjExcel.Range(t3_totals_count_letter 		& "2:" & t3_totals_count_letter 		& "2").Interior.ColorIndex = 6
	ObjExcel.Range(t3_apps_recvd_count_letter 	& "5:" & t3_iapps_recvd_percent_letter 	& "5").Interior.ColorIndex = 6
	ObjExcel.Range(t3_revw_i_count_letter 		& "5:" & t3_revw_i_percent_letter 		& "5").Interior.ColorIndex = 6
	ObjExcel.Range(t3_revw_a_count_letter 		& "5:" & t3_revw_a_percent_letter 		& "5").Interior.ColorIndex = 6
	ObjExcel.Range(t3_totals_count_letter 		& "5:" & t3_totals_count_letter 		& "5").Interior.ColorIndex = 6
	ObjExcel.Range(t3_apps_recvd_count_letter 	& "8:" & t3_iapps_recvd_percent_letter 	& "8").Interior.ColorIndex = 6
	ObjExcel.Range(t3_revw_i_count_letter 		& "8:" & t3_revw_i_percent_letter 		& "8").Interior.ColorIndex = 6
	ObjExcel.Range(t3_revw_a_count_letter 		& "8:" & t3_revw_a_percent_letter 		& "8").Interior.ColorIndex = 6
	ObjExcel.Range(t3_totals_count_letter 		& "8:" & t3_totals_count_letter 		& "8").Interior.ColorIndex = 6
	ObjExcel.Range(t3_revw_a_count_letter 		& "14:" & t3_revw_a_percent_letter 		& "14").Interior.ColorIndex = 6

	objStatsExcel.Range(t3_apps_recvd_count_letter 	& "2:" & t3_intvs_percent_letter 		& "2").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_revw_a_count_letter 		& "2:" & t3_revw_a_percent_letter 		& "2").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_totals_count_letter 		& "2:" & t3_totals_count_letter 		& "2").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_apps_recvd_count_letter 	& "5:" & t3_iapps_recvd_percent_letter 	& "5").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_revw_i_count_letter 		& "5:" & t3_revw_i_percent_letter 		& "5").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_revw_a_count_letter 		& "5:" & t3_revw_a_percent_letter 		& "5").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_totals_count_letter 		& "5:" & t3_totals_count_letter 		& "5").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_apps_recvd_count_letter 	& "8:" & t3_iapps_recvd_percent_letter 	& "8").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_revw_i_count_letter 		& "8:" & t3_revw_i_percent_letter 		& "8").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_revw_a_count_letter 		& "8:" & t3_revw_a_percent_letter 		& "8").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_totals_count_letter 		& "8:" & t3_totals_count_letter 		& "8").Interior.ColorIndex = 6
	objStatsExcel.Range(t3_revw_a_count_letter 		& "14:" & t3_revw_a_percent_letter 		& "14").Interior.ColorIndex = 6

	'Query date/time/runtime info
	ObjExcel.Cells(1, output_headers).Font.Bold = TRUE
	ObjExcel.Cells(2, output_headers).Font.Bold = TRUE
	ObjExcel.Cells(1, output_headers).Value = "Query date and time:"
	ObjExcel.Cells(2, output_headers).Value = "Query runtime (in seconds):"
	ObjExcel.Cells(1, output_data).Value = now
	ObjExcel.Cells(2, output_data).Value = timer - query_start_time

	ObjExcel.columns(output_headers).AutoFit()
	ObjExcel.columns(output_data).AutoFit()

	objWorkbook.Save

	'Query date/time/runtime info
	objStatsExcel.Cells(1, output_headers).Font.Bold = TRUE
	objStatsExcel.Cells(2, output_headers).Font.Bold = TRUE
	objStatsExcel.Cells(1, output_headers).Value = "Query date and time:"
	objStatsExcel.Cells(2, output_headers).Value = "Query runtime (in seconds):"
	objStatsExcel.Cells(1, output_data).Value = now
	objStatsExcel.Cells(2, output_data).Value = timer - query_start_time

	objStatsExcel.columns(output_headers).AutoFit()
	objStatsExcel.columns(output_data).AutoFit()

	objStatsExcel.ActiveWorkbook.SaveAs t_drive & "\IPA\Restricted\DMA\PowerBIData\ES QI\Recertification Statistics\" & report_date & " Renewal Data.xlsx"
	If last_day_checkbox = checked Then objStatsExcel.ActiveWorkbook.SaveAs t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Recertification Statistics\Data\Last Processing Day for " & report_date & " Renewal Data.xlsx"

	'HERE WE NEED TO CLOSE THE NEW SHEETS
	ObjStatsExcel.Quit

	' WE AREN'T USING THIS BUT I AM SAVING THIS CODE BECAUSE IT IS THE ONLY PLACE WE HAVE THE BORDER FUNCTIONS INFORMATION
	' border_array = array("B1:C"&last_row, "D1:D"&last_row, "E1:E"&last_row, "F1:G"&last_row, "H1:I"&last_row, "A2:I2", "A3:I5", "A6:I8", "B10:I10", "K3:AE5", "K6:AE8", "K9:AE11", "K12:AE14", "K15:AE17", "M1:N17", "O1:P17",_
	'  					 "Q1:R17", "S1:T17", "U1:V17", "W1:X17", "Y1:Z17", "AA1:AB17", "AC1:AD17", "AE1:AE17")
	'
	' For each group in border_array
	' 	With ObjExcel.ActiveSheet.Range(group)
	' 		With .Borders(7)	'left'
	' 			.LineStyle = 1
	' 			.Weight = 2
	' 			.ColorIndex = -4105
	' 		End With
	' 		With .Borders(8)	'Top'
	' 			.LineStyle = 1
	' 			.Weight = 2
	' 			.ColorIndex = -4105
	' 		End With
	' 		With .Borders(9)	'Bottom'
	' 			.LineStyle = 1
	' 			.Weight = 2
	' 			.ColorIndex = -4105
	' 		End With
	' 		With .Borders(10)	'Right'
	' 			.LineStyle = 1
	' 			.Weight = 2
	' 			.ColorIndex = -4105
	' 		End With
	' 	End With
	' Next

	run_time = timer - query_start_time
	end_msg = "Case details have been added to the Review Report" & vbCr & vbCr & "Run time: " & run_time & " seconds."

	If original_renewal_option = "Send NOMIs" Then renewal_option = "Send NOMIs"
	If original_renewal_option <> "Send NOMIs" Then ObjExcel.Quit

ElseIf renewal_option = "Send Appointment Letters" Then
	MAXIS_footer_month = CM_mo							'Setting the footer month and year based on the review month.
	MAXIS_footer_year = CM_yr

	check_REVW = False
	If REPT_month = CM_plus_1_mo AND REPT_year = CM_plus_1_yr Then
		check_REVW = True
		MAXIS_footer_month = CM_plus_1_mo							'Setting the footer month and year based on the review month.
		MAXIS_footer_year = CM_plus_1_yr
	End If
	' MsgBox check_REVW
	'Finding the last column that has something in it so we can add to the end.
	col_to_use = 0
	Do
		col_to_use = col_to_use + 1
		col_header = trim(ObjExcel.Cells(1, col_to_use).Value)
		If col_header = "APPT NOTC Sent" Then notc_col = col_to_use
		If col_header = "APPT NOTC Date" Then notc_date_col = col_to_use
	Loop until col_header = ""

	' MsgBox "NOTC Col - " & notc_col & vbCr & "NOTC Date Col - " & notc_date_col
	If notc_col = "" OR notc_date_col = "" Then
		last_col_letter = convert_digit_to_excel_column(col_to_use)

		'Insert columns in excel for additional information to be added
		column_end = last_col_letter & "1"
		Set objRange = objExcel.Range(column_end).EntireColumn

		If notc_col = "" Then
			objRange.Insert(xlShiftToRight)			'We neeed one more columns
			notc_col = col_to_use		'Setting the column to individual variables so we enter the found information in the right place
			col_to_use = col_to_use + 1

			ObjExcel.Cells(1, notc_col).Value = "APPT NOTC Sent"			'creating the column headers for the statistics information for the day of the run.
			objExcel.Cells(1, notc_col).Font.Bold = True		'bold font'
			ObjExcel.columns(notc_col).NumberFormat = "@" 		'formatting as text
			ObjExcel.columns(notc_col).AutoFit() 				'sizing the columns'
		End If
		If notc_date_col = "" Then
			objRange.Insert(xlShiftToRight)			'We neeed one more columns
			notc_date_col = col_to_use		'Setting the column to individual variables so we enter the found information in the right place

			ObjExcel.Cells(1, notc_date_col).Value = "APPT NOTC Date"			'creating the column headers for the statistics information for the day of the run.
			objExcel.Cells(1, notc_date_col).Font.Bold = True		'bold font'
			ObjExcel.columns(notc_date_col).NumberFormat = "m/d/yy" 		'formatting as text
			ObjExcel.columns(notc_date_col).AutoFit() 						'fsizing the columns'
		End If
	End If

	today_mo = DatePart("m", date)
	today_mo = right("00" & today_mo, 2)

	today_day = DatePart("d", date)
	today_day = right("00" & today_day, 2)

	today_yr = DatePart("yyyy", date)
	today_yr = right(today_yr, 2)
	today_date = today_mo & "/" & today_day & "/" & today_yr
	call back_to_SELF

	'Now we loop through the whole Excel List and sending notices on the right cases
	excel_row = "2"		'starts at row 2'
	Do
		notc_col_info = trim(ObjExcel.Cells(excel_row, notc_col).Value)
		MAXIS_case_number 	= trim(ObjExcel.Cells(excel_row,  2).Value)			'getting the case number from the spreadsheet
		' MsgBox "row - " & excel_row & vbCr & "col - " & notc_col & vbCr & "val - *" & notc_col_info & "*"
		If notc_col_info = "" AND MAXIS_case_number <> "" Then
			send_appt_notc = True
			If check_REVW = True Then
				send_appt_notc = False
				Call check_for_MAXIS(FALSE)		'making sure we haven't passworded out

				' MAXIS_case_number = case_number_to_check		'setting the case number for NAV functions
				call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)		'Go to STAT REVW and be sure the case is not privleged.
				If is_this_priv = FALSE Then
					EMReadScreen recvd_date, 8, 13, 37										'Reading the CAF Received Date and format
					recvd_date = replace(recvd_date, " ", "/")
					if recvd_date = "__/__/__" then recvd_date = ""

					EMReadScreen interview_date, 8, 15, 37									'Reading the interview date and format
					interview_date = replace(interview_date, " ", "/")
					if interview_date = "__/__/__" then interview_date = ""

					EMReadScreen cash_review_status, 1, 7, 40								'Reading the review status and format
					EMReadScreen snap_review_status, 1, 7, 60
					EMReadScreen hc_review_status, 1, 7, 73
					If cash_review_status = "_" Then cash_review_status = ""
					If snap_review_status = "_" Then snap_review_status = ""
					If hc_review_status = "_" Then hc_review_status = ""

					revw_status_all_n = True
					If cash_review_status = "I" OR cash_review_status = "U" OR cash_review_status = "A" Then revw_status_all_n = False
					If snap_review_status = "I" OR snap_review_status = "U" OR snap_review_status = "A" Then revw_status_all_n = False
					If hc_review_status = "I" OR hc_review_status = "U" OR hc_review_status = "A" Then revw_status_all_n = False

					If recvd_date = "" AND revw_status_all_n = True Then send_appt_notc = True
				End If
			End If
			If send_appt_notc = False Then ObjExcel.Cells(excel_row, notc_col).Value = "N/A"

			If send_appt_notc = True or check_REVW = False Then
				' MsgBox excel_row

				forms_to_arep = ""
				forms_to_swkr = ""
				programs = ""
				intvw_programs = ""
				renewal_guidance_needed = False
				renewal_guidance_confirmed = False

				Call read_boolean_from_excel(ObjExcel.Cells(excel_row,  3).Value, er_with_intherview)
				Call read_boolean_from_excel(objExcel.cells(excel_row,  6).value, MFIP_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row,  7).value, DWP_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row,  8).value, GA_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row,  9).value, MSA_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row, 10).value, GRH_status)
				Call read_boolean_from_excel(objExcel.cells(excel_row, 13).value, SNAP_status)

				REPT_full = REPT_month & "/" & REPT_year
				CASH_SR_Info = trim(objExcel.cells(excel_row, 11).value)
				CASH_ER_Info = trim(objExcel.cells(excel_row, 12).value)
				SNAP_SR_Info = trim(objExcel.cells(excel_row, 14).value)
				SNAP_ER_Info = trim(objExcel.cells(excel_row, 15).value)

				If MFIP_status = True and SNAP_status = True Then
					intvw_programs = "MFIP/SNAP"
				ElseIf MFIP_status = True Then
					intvw_programs = "MFIP"
				ElseIf SNAP_status = True Then
					intvw_programs = "SNAP"
				End If
				If CASH_ER_Info = REPT_full then
					If MFIP_status = True Then programs = programs & "/MFIP"
					If DWP_status = True Then programs = programs & "/DWP"
					If GA_status = True Then programs = programs & "/GA"
					If MSA_status = True Then programs = programs & "/MSA"
				End If
				If CASH_SR_Info = REPT_full OR CASH_ER_Info = REPT_full then
					If GRH_status = True Then programs = programs & "/GRH"
				End If
				If SNAP_SR_Info = REPT_full OR SNAP_ER_Info = REPT_full then
					If SNAP_status = True Then programs = programs & "/SNAP"
				End If
				If left(programs, 1) = "/" Then programs = right(programs, len(programs)-1)
				' MsgBox "Sending NOTC - " & MAXIS_case_number & " - excel row - " & excel_row & " Programs - " & programs

				interview_end_date = CM_plus_1_mo & "/15/" & CM_plus_1_yr
				last_day_of_recert = CM_plus_2_mo & "/01/" & CM_plus_2_yr
			    last_day_of_recert = dateadd("D", -1, last_day_of_recert)

				notes_info = Trim(ObjExcel.cells(excel_row, 25).value)

				' If er_with_intherview = True Then

				If er_with_intherview = True AND MFIP_status = True Then
					'Writing the SPEC MEMO - dates will be input from the determination made earlier.
					' MsgBox "We're writing a MEMO here"
					Call start_a_new_spec_memo_and_continue(memo_started)

					IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY

						CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case.")
						CALL write_variable_in_SPEC_MEMO("")
						' CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You must also complete an interview for your " & intvw_programs & " case to continue.")
						CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You may need to complete an interview for your " & intvw_programs & " case to continue.")
						CALL write_variable_in_SPEC_MEMO("")
						' Call write_variable_in_SPEC_MEMO("  *** Please complete your interview by " & interview_end_date & ". ***")
						Call write_variable_in_SPEC_MEMO("  *** If required, complete your interview by " & interview_end_date & ". ***")
						Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
						Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
						CALL write_variable_in_SPEC_MEMO("")
						CALL write_variable_in_SPEC_MEMO("**  Your " & programs & " case will close on " & last_day_of_recert & " unless    **")
						CALL write_variable_in_SPEC_MEMO("** we receive your paperwork and complete the interview. **")
						CALL write_variable_in_SPEC_MEMO("")
						'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
						' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
						' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
						' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
						' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
						' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
						' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
						' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
						' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
						' Call write_variable_in_SPEC_MEMO(" ")
						CALL write_variable_in_SPEC_MEMO("** You can submit documents Online at www.MNBenefits.org **")
						CALL write_variable_in_SPEC_MEMO("Other options for submitting documents to Hennepin County:")
						CALL write_variable_in_SPEC_MEMO(" - Mail, Fax, or Drop Boxes at service centers")
						CALL write_variable_in_SPEC_MEMO(" - Email with document attachment.EMAIL: hhsews@hennepin.us")
						CALL write_variable_in_SPEC_MEMO("   (Only attach PNG, JPG, TIF, DOC, PDF, or HTM file types)")
						Call write_variable_in_SPEC_MEMO(" ")
						CALL write_variable_in_SPEC_MEMO("Domestic violence brochures are available at this website: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG. You can also request a paper copy.")

						PF4         'Submit the MEMO

						memo_row = 7                                            'Setting the row for the loop to read MEMOs
						ObjExcel.Cells(excel_row, notc_col).Value = "N"         'Defaulting this to 'N'
						Do
							EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
							EMReadScreen print_status, 7, memo_row, 67
							If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
								ObjExcel.Cells(excel_row, notc_col).Value = "Y"             'If we've found this then no reason to keep looking.
								ObjExcel.Cells(excel_row, notc_date_col).Value = today_date        'If we've found this then no reason to keep looking.
								successful_notices = successful_notices + 1                 'For statistical purposes
								Exit Do
							End If

							memo_row = memo_row + 1           'Looking at next row'
						Loop Until create_date = "        "

					ELSE
						ObjExcel.Cells(excel_row, notc_col).Value = "N"         'Setting this as N if the MEMO failed
						call back_to_SELF
					END IF
					' ObjExcel.Cells(excel_row, 25).Value = "All progs - " & programs & " : INTVW Progs - " & intvw_programs
				Else
					ObjExcel.Cells(excel_row, notc_col).Value = "N/A"
					renewal_guidance_needed = True
					If notes_info = "PRIV Case." then renewal_guidance_needed = False
				End If

				If ObjExcel.Cells(excel_row, notc_col).Value = "Y" OR renewal_guidance_needed = True Then

					Call start_a_new_spec_memo_and_continue(memo_started)   'Starting a MEMO to send information about verifications

					IF memo_started = True THEN
						If renewal_guidance_needed = True Then
							CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case and is due by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ".")
							' CALL write_variable_in_SPEC_MEMO("")
						End If

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
						If renewal_guidance_needed = False Then CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")
						If renewal_guidance_needed = True Then
							CALL write_variable_in_SPEC_MEMO("** You can submit documents Online at www.MNBenefits.org **")
							CALL write_variable_in_SPEC_MEMO("Other options for submitting documents to Hennepin County:")
							CALL write_variable_in_SPEC_MEMO(" - Mail, Fax, or Drop Boxes at service centers")
							CALL write_variable_in_SPEC_MEMO(" - Email with document attachment.EMAIL: hhsews@hennepin.us")
							CALL write_variable_in_SPEC_MEMO("   (Only attach PNG, JPG, TIF, DOC, PDF, or HTM file types)")
							CALL write_variable_in_SPEC_MEMO("")
							CALL write_variable_in_SPEC_MEMO("Once we receive and process your renewal paperwork, you will receive information BY MAIL with possible follow up or actions taken on your case. Call 612-596-1300 if you have additional questions.")
						End If

						PF4 'Submit the MEMO'

						If renewal_guidance_needed = True Then
							memo_row = 7                                            'Setting the row for the loop to read MEMOs
							Do
								EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
								EMReadScreen print_status, 7, memo_row, 67
								If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
									renewal_guidance_confirmed = True
									ObjExcel.Cells(excel_row, notc_col).Value = "RG"
									' ObjExcel.Cells(excel_row, 25).Value = "All progs - " & programs & " : INTVW Progs - " & intvw_programs
									Exit Do
								End If

								memo_row = memo_row + 1           'Looking at next row'
							Loop Until create_date = "        "
						End If
					End If

					If ObjExcel.Cells(excel_row, notc_col).Value = "Y" Then
						start_a_blank_case_note
						CALL write_variable_in_CASE_NOTE("*** Notice of " & intvw_programs & " Recertification Interview Sent ***")
						CALL write_variable_in_case_note("* A notice has been sent to client with detail about how to call in for an interview.")
						CALL write_variable_in_case_note("* Client must submit paperwork and call 612-596-1300 to complete interview.")
						If forms_to_arep = "Y" then call write_variable_in_case_note("* Copy of notice sent to AREP.")
						If forms_to_swkr = "Y" then call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
						call write_variable_in_case_note("---")
						CALL write_variable_in_case_note("Link to Domestic Violence Brochure sent to client in SPEC/MEMO as a part of interview notice.")
						call write_variable_in_case_note("---")
						call write_variable_in_case_note(worker_signature)
					ElseIf renewal_guidance_confirmed = True Then
						start_a_blank_case_note
						CALL write_variable_in_CASE_NOTE("Notice sent for " & programs & " Renewal Guidance")
						Call write_variable_in_case_note("* A renewal is due for this case for " & REPT_month & "/" & REPT_year)
						Call write_variable_in_case_note("* Reminder notice sent with forms due date and verification options.")
						Call write_variable_in_case_note("  -This is NOT an official verification request.-")
						Call write_variable_in_case_note("---")
						Call write_variable_in_case_note(worker_signature)
					End If
					PF3
				End If
			End If
		End If
		excel_row = excel_row + 1
	Loop until MAXIS_case_number = ""

	worksheet_found = FALSE
	'Finding all of the worksheets available in the file. We will likely open up the main 'Review Report' so the script will default to that one.
	For Each objWorkSheet In objWorkbook.Worksheets
		If instr(objWorkSheet.Name, "NOTICES") <> 0 Then
			objWorkSheet.Activate
			worksheet_found = TRUE
		End If
	Next

	If worksheet_found = FALSE Then
		'Going to another sheet, to enter worker-specific statistics and naming it
		sheet_name = "NOTICES"
		ObjExcel.Worksheets.Add().Name = sheet_name

		entry_row = 1

		objExcel.Cells(entry_row, 1).Value      = "Appointment Notices run on:"     'Date and time the script was completed
	    objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
	    objExcel.Cells(entry_row, 2).Value      = now
	    entry_row = entry_row + 1

	    objExcel.Cells(entry_row, 1).Value      = "Runtime (in seconds)"            'Enters the amount of time it took the script to run
	    objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
	    objExcel.Cells(entry_row, 2).Value      = timer - query_start_time
	    entry_row = entry_row + 1

	    objExcel.Cells(entry_row, 1).Value      = "Total Cases assesed"             'All cases from the spreadsheet
	    objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
	    objExcel.Cells(entry_row, 2).Value    	= excel_row - 2
	    entry_row = entry_row + 1

		objExcel.Cells(entry_row, 1).Value      = "Total Cases with ER Interview"             'All cases from the spreadsheet
		objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
		objExcel.Cells(entry_row, 2).Value      = "=COUNTIFS(Table1[Interview ER],"&is_true&")"
		' total_row = entry_row
		entry_row = entry_row + 1

		objExcel.Cells(entry_row, 1).Value      = "Total MFIP Cases with ER Interview"             'All cases from the spreadsheet
		objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
		objExcel.Cells(entry_row, 2).Value      = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[MFIP Status],"&is_true&")"
		total_row = entry_row
		entry_row = entry_row + 1

	    if successful_notices = "" then successful_notices = 0
	    objExcel.Cells(entry_row, 1).Value      = "Appointment Notices Sent"        'number of notices that were successful
	    objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
		objExcel.Cells(entry_row, 2).Value      = "=COUNTIFS(Table1[APPT NOTC Sent]," & Chr(34) & "Y" & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
	    appt_row = entry_row
	    entry_row = entry_row + 1

	    objExcel.Cells(entry_row, 1).Value      = "Percentage successful"           'calculation of the percent of successful notices
	    objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
	    objExcel.Cells(entry_row, 2).Value      = "=B" & appt_row & "/B" & total_row
	    objExcel.Cells(entry_row, 2).NumberFormat = "0.00%"		'Formula should be percent
	    entry_row = entry_row + 1

		objExcel.Cells(entry_row, 1).Value      = "Renewal Guidance Notices Sent"           'calculation of the percent of successful notices
		objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
		objExcel.Cells(entry_row, 2).Value      = "=COUNTIFS(Table1[APPT NOTC Sent]," & Chr(34) & "RG" & Chr(34) & ")"
		' objExcel.Cells(entry_row, 2).NumberFormat = "0.00%"		'Formula should be percent
		entry_row = entry_row + 1
	End If

	date_stats_row = 0
	Do
		date_stats_row = date_stats_row + 1
		in_the_cell = trim(objExcel.Cells(date_stats_row, 1).Value)
	Loop until in_the_cell = ""

	objExcel.Cells(date_stats_row, 1).Value      = "Appointment Notices Sent on " & today_date        'number of notices that were successful
	objExcel.Cells(date_stats_row, 1).Font.Bold 	= TRUE
	objExcel.Cells(date_stats_row, 2).Value      = "=COUNTIFS(Table1[APPT NOTC Date]," & Chr(34) & today_date & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written

	objExcel.Columns(1).AutoFit()
	objExcel.Columns(2).AutoFit()

	end_msg = "NOTICES have been sent on " & successful_notices & " cases today. Information added to the Review Report Excel document."

ElseIf renewal_option = "Create Worklist" Then

	MAXIS_footer_month = REPT_month							'Setting the footer month and year based on the review month. We do not run statistics in CM + 2
	MAXIS_footer_year = REPT_year

	recert_cases = 0	            'incrementor for the array

	back_to_self    'We need to get back to SELF and manually update the footer month
	Call navigate_to_MAXIS_screen("REPT", "REVS")		'going to REPT REVS where all the information is displayed'
	EMWriteScreen REPT_month, 20, 55					'going to the right month
	EMWriteScreen REPT_year, 20, 58
	transmit

	'We are going to look at REPT/REVS for each worker in Hennepin County
	For each worker in worker_array
		worker = trim(worker)				'get to the right worker
		If worker = "" then exit for
		Call write_value_and_transmit(worker, 21, 6)   'writing in the worker number in the correct col

		'Grabbing case numbers from REVS for requested worker
		DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
			row = 7	'Setting or resetting this to look at the top of the list
			DO		'All of this loops until row = 19
				'Reading case information (case number, SNAP status, and cash status)
				EMReadScreen MAXIS_case_number, 8, row, 6
				MAXIS_case_number = trim(MAXIS_case_number)
				EMReadScreen SNAP_status, 1, row, 45
				EMReadScreen cash_status, 1, row, 39
				EmReadscreen HC_status, 1, row, 49
				EMReadScreen MAGI_status, 4, row, 55
				EMReadScreen recvd_date, 8, row, 62
				EMReadScreen intvw_date, 8, row, 72

				'Navigates though until it runs out of case numbers to read
				IF MAXIS_case_number = "" then exit do

				'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
				If cash_status = "-" 	then cash_status = ""
				If SNAP_status = "-" 	then SNAP_status = ""
				If HC_status = "-" 		then HC_status = ""

				ReDim Preserve review_array(notes_const, recert_cases)		'resizing the array

				'Adding the case information to the array
				review_array(worker_const, recert_cases) = worker
				review_array(case_number_const, recert_cases) = trim(MAXIS_case_number)
				review_array(CASH_revw_status_const, recert_cases) = cash_status
				review_array(SNAP_revw_status_const, recert_cases) = SNAP_status
				review_array(HC_revw_status_const, recert_cases) = HC_status
				review_array(HC_MAGI_code_const, recert_cases) = trim(MAGI_status)
				review_array(review_recvd_const, recert_cases) = replace(recvd_date, " ", "/")
				If review_array(review_recvd_const, recert_cases) = "__/__/__" Then review_array(review_recvd_const, recert_cases) = ""
				review_array(interview_date_const, recert_cases) = replace(intvw_date, " ", "/")
				If review_array(interview_date_const, recert_cases) = "__/__/__" Then review_array(interview_date_const, recert_cases) = ""
				review_array(saved_to_excel_const, recert_cases) = FALSE

				recert_cases = recert_cases + 1
				STATS_counter = STATS_counter + 1						'adds one instance to the stats counter

				row = row + 1    'On the next loop it must look to the next row
				MAXIS_case_number = "" 'Clearing variables before next loop
			Loop until row = 19		'Last row in REPT/REVS
			'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
			'if max reviews are reached, the goes to next worker is applicable
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	next
	Call back_to_SELF
	had_to_check_STAT = 0

	'Now we are going to look at the Excel spreadsheet that has all of the reviews saved.
	excel_row = "2"		'starts at row 2'
	Do
		case_number_to_check = trim(ObjExcel.Cells(excel_row, 2).Value)			'getting the case number from the spreadsheet
		found_in_array = FALSE													'variale to identify if we have found this case in our array

		'Here we look through the entire array until we find a match
		For revs_item = 0 to UBound(review_array, 2)
			If case_number_to_check = review_array(case_number_const, revs_item) Then		'if the case numbers match we have found our case.

				'Saving the information from the excel into the array
				review_array(interview_const,       revs_item) = ObjExcel.Cells(excel_row,  3).value     'COL C
				review_array(no_interview_const,    revs_item) = ObjExcel.Cells(excel_row,  4).value     'COL D
				review_array(current_SR_const,      revs_item) = ObjExcel.Cells(excel_row,  5).value     'COL E
				review_array(MFIP_status_const,     revs_item) = ObjExcel.Cells(excel_row,  6).value     'COL F
				review_array(DWP_status_const,      revs_item) = ObjExcel.Cells(excel_row,  7).value     'COL G
				review_array(GA_status_const,       revs_item) = ObjExcel.Cells(excel_row,  8).value     'COL H
				review_array(MSA_status_const,      revs_item) = ObjExcel.Cells(excel_row,  9).value     'COL I
				review_array(GRH_status_const,      revs_item) = ObjExcel.Cells(excel_row, 10).value     'COL J
				review_array(CASH_next_SR_const,    revs_item) = ObjExcel.Cells(excel_row, 11).value     'COL K
				review_array(CASH_next_ER_const,    revs_item) = ObjExcel.Cells(excel_row, 12).value     'COL L
				review_array(SNAP_status_const,     revs_item) = ObjExcel.Cells(excel_row, 13).value     'COL M
				review_array(SNAP_next_SR_const,    revs_item) = ObjExcel.Cells(excel_row, 14).value     'COL N
				review_array(SNAP_next_ER_const,    revs_item) = ObjExcel.Cells(excel_row, 15).value     'COL O
				review_array(MA_status_const,       revs_item) = ObjExcel.Cells(excel_row, 16).value     'COL P
				review_array(MSP_status_const,      revs_item) = ObjExcel.Cells(excel_row, 17).value     'COL Q
				review_array(HC_next_SR_const,      revs_item) = ObjExcel.Cells(excel_row, 18).value     'COL R
				review_array(HC_next_ER_const,      revs_item) = ObjExcel.Cells(excel_row, 19).value     'COL S
				review_array(Language_const,        revs_item) = ObjExcel.Cells(excel_row, 20).value     'COL T
				review_array(Interpreter_const,     revs_item) = ObjExcel.Cells(excel_row, 21).value     'COL U
				review_array(phone_1_const,         revs_item) = ObjExcel.Cells(excel_row, 22).value     'COL V
				review_array(phone_2_const,         revs_item) = ObjExcel.Cells(excel_row, 23).value     'COL W
				review_array(phone_3_const,         revs_item) = ObjExcel.Cells(excel_row, 24).value     'COL X
				review_array(notes_const,           revs_item) = ObjExcel.Cells(excel_row, 25).value     'COL Y

				found_in_array = TRUE			'this lets the script know that this case was found in the array
				review_array(saved_to_excel_const, revs_item) = TRUE
				Exit For						'if we found a match, we should stop looking
			End If
		Next
		'if the case was not found in the array, we need to look in STAT for the information
		If found_in_array = FALSE AND case_number_to_check <> "" Then
			Call check_for_MAXIS(FALSE)		'making sure we haven't passworded out
			had_to_check_STAT = had_to_check_STAT + 1

			MAXIS_case_number = case_number_to_check		'setting the case number for NAV functions
			call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)		'Go to STAT REVW and be sure the case is not privleged.
			If is_this_priv = FALSE Then
				EMReadScreen recvd_date, 8, 13, 37										'Reading the CAF Received Date and format
				recvd_date = replace(recvd_date, " ", "/")
				if recvd_date = "__/__/__" then recvd_date = ""

				EMReadScreen interview_date, 8, 15, 37									'Reading the interview date and format
				interview_date = replace(interview_date, " ", "/")
				if interview_date = "__/__/__" then interview_date = ""

				EMReadScreen cash_review_status, 1, 7, 40								'Reading the review status and format
				EMReadScreen snap_review_status, 1, 7, 60
				EMReadScreen hc_review_status, 1, 7, 73
				If cash_review_status = "_" Then cash_review_status = ""
				If snap_review_status = "_" Then snap_review_status = ""
				If hc_review_status = "_" Then hc_review_status = ""

				' If cash_review_status <> "" Then ObjExcel.Cells(excel_row, cash_stat_excel_col).Value = cash_review_status		'Enter all the information into Excel
				' If snap_review_status <> "" Then ObjExcel.Cells(excel_row, snap_stat_excel_col).Value = snap_review_status
				' If hc_review_status <> "" Then ObjExcel.Cells(excel_row, hc_stat_excel_col).Value = hc_review_status
				' If recvd_date <> "" Then ObjExcel.Cells(excel_row, recvd_date_excel_col).Value = recvd_date
				' If interview_date <> "" Then ObjExcel.Cells(excel_row, intvw_date_excel_col).Value = interview_date
				ReDim Preserve review_array(notes_const, recert_cases)		'resizing the array

				review_array(CASH_revw_status_const, 	recert_cases) = cash_review_status
				review_array(SNAP_revw_status_const, 	recert_cases) = snap_review_status
				review_array(HC_revw_status_const, 		recert_cases) = hc_review_status
				review_array(review_recvd_const, 		recert_cases) = trim(recvd_date)
				review_array(interview_date_const, 		recert_cases) = interview_date

				'Saving the information from the excel into the array
				' review_array(interview_const,       recert_cases) = ObjExcel.Cells(excel_row,  3).value     'COL C
				If  trim(ObjExcel.Cells(excel_row,  3).value) = "TRUE" Then review_array(interview_const,       recert_cases) = TRUE
				If  trim(ObjExcel.Cells(excel_row,  3).value) = "FALSE" Then review_array(interview_const,       recert_cases) = FALSE
				review_array(no_interview_const,    recert_cases) = ObjExcel.Cells(excel_row,  4).value     'COL D
				review_array(current_SR_const,      recert_cases) = ObjExcel.Cells(excel_row,  5).value     'COL E
				If  trim(ObjExcel.Cells(excel_row,  5).value) = "TRUE" Then review_array(current_SR_const,       recert_cases) = TRUE
				If  trim(ObjExcel.Cells(excel_row,  5).value) = "FALSE" Then review_array(current_SR_const,       recert_cases) = FALSE
				review_array(MFIP_status_const,     recert_cases) = ObjExcel.Cells(excel_row,  6).value     'COL F
				review_array(DWP_status_const,      recert_cases) = ObjExcel.Cells(excel_row,  7).value     'COL G
				review_array(GA_status_const,       recert_cases) = ObjExcel.Cells(excel_row,  8).value     'COL H
				review_array(MSA_status_const,      recert_cases) = ObjExcel.Cells(excel_row,  9).value     'COL I
				review_array(GRH_status_const,      recert_cases) = ObjExcel.Cells(excel_row, 10).value     'COL J
				review_array(CASH_next_SR_const,    recert_cases) = ObjExcel.Cells(excel_row, 11).value     'COL K
				review_array(CASH_next_ER_const,    recert_cases) = ObjExcel.Cells(excel_row, 12).value     'COL L
				review_array(SNAP_status_const,     recert_cases) = ObjExcel.Cells(excel_row, 13).value     'COL M
				review_array(SNAP_next_SR_const,    recert_cases) = ObjExcel.Cells(excel_row, 14).value     'COL N
				review_array(SNAP_next_ER_const,    recert_cases) = ObjExcel.Cells(excel_row, 15).value     'COL O
				review_array(MA_status_const,       recert_cases) = ObjExcel.Cells(excel_row, 16).value     'COL P
				review_array(MSP_status_const,      recert_cases) = ObjExcel.Cells(excel_row, 17).value     'COL Q
				review_array(HC_next_SR_const,      recert_cases) = ObjExcel.Cells(excel_row, 18).value     'COL R
				review_array(HC_next_ER_const,      recert_cases) = ObjExcel.Cells(excel_row, 19).value     'COL S
				review_array(Language_const,        recert_cases) = ObjExcel.Cells(excel_row, 20).value     'COL T
				review_array(Interpreter_const,     recert_cases) = ObjExcel.Cells(excel_row, 21).value     'COL U
				review_array(phone_1_const,         recert_cases) = ObjExcel.Cells(excel_row, 22).value     'COL V
				review_array(phone_2_const,         recert_cases) = ObjExcel.Cells(excel_row, 23).value     'COL W
				review_array(phone_3_const,         recert_cases) = ObjExcel.Cells(excel_row, 24).value     'COL X
				review_array(notes_const,           recert_cases) = ObjExcel.Cells(excel_row, 25).value     'COL Y
				review_array(saved_to_excel_const, 	recert_cases) = TRUE

				recert_cases = recert_cases + 1

			End If

			Call back_to_SELF		'Back out in case we need to look into another case.
		End If
		excel_row = excel_row + 1		'going to the next excel
	Loop until case_number_to_check = ""

	ObjExcel.ActiveWorkbook.Close
	ObjExcel.Application.Quit
	ObjExcel.Quit

	' MsgBox "Had to check STAT on " & had_to_check_STAT & " cases."

	'Opening the Excel file, (now that the dialog is done)
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	Set objWorkbook = objExcel.Workbooks.Add()
	objExcel.DisplayAlerts = True

	'Changes name of Excel sheet to "Case information"
	ObjExcel.ActiveSheet.Name = "ER cases " & REPT_month & "-" & REPT_year & " "

	'formatting excel file with columns for case number and interview date/time
	objExcel.cells(1, 1).value 	= "X number"
	objExcel.cells(1, 2).value 	= "Case Number"
	objExcel.cells(1, 3).value 	= "Programs"
	objExcel.cells(1, 4).value 	= "Case language"
	objExcel.Cells(1, 5).value 	= "Interpreter"
	objExcel.cells(1, 6).value 	= "Phone # One"
	objExcel.cells(1, 7).value 	= "Phone # Two"
	objExcel.Cells(1, 8).value 	= "Phone # Three"

	FOR i = 1 to 8									'formatting the cells'
		objExcel.Cells(1, i).Font.Bold = True		'bold font'
		ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
		objExcel.Columns(i).AutoFit()				'sizing the columns'
	NEXT

	excel_row = 2
	er_case_to_work = 0

	For revs_item = 0 to UBound(review_array, 2)
		If review_array(saved_to_excel_const, revs_item) = TRUE Then
			If review_array(interview_const, revs_item) = TRUE AND review_array(review_recvd_const, revs_item) = "" Then
				'determining the programs list
				If ( review_array(SNAP_status_const, revs_item) = "TRUE" and review_array(MFIP_status_const, revs_item) = "TRUE" ) then
					programs_list = "SNAP & MFIP"
				elseif review_array(SNAP_status_const, revs_item) = "TRUE" then
					programs_list = "SNAP"
				elseif review_array(MFIP_status_const, revs_item) = "TRUE" then
					programs_list = "MFIP"
				End if
				'Excel output of Interview Required case information
	            If review_array(notes_const, revs_item) <> "PRIV Case." then
	    	        ObjExcel.Cells(excel_row, 1).value = review_array(worker_const,       revs_item)
	    	        ObjExcel.Cells(excel_row, 2).value = review_array(case_number_const,  revs_item)
	    	        ObjExcel.Cells(excel_row, 3).value = programs_list
	    	        ObjExcel.Cells(excel_row, 4).value = review_array(Language_const,     revs_item)
	    	        ObjExcel.Cells(excel_row, 5).value = review_array(Interpreter_const,  revs_item)
	    	        ObjExcel.Cells(excel_row, 6).value = review_array( phone_1_const,     revs_item)
	    	        ObjExcel.Cells(excel_row, 7).value = review_array( phone_2_const,     revs_item)
	    	        ObjExcel.Cells(excel_row, 8).value = review_array( phone_3_const,     revs_item)
					er_case_to_work = er_case_to_work + 1
	                excel_row = excel_row + 1
	            End if
			End If
		End If
	Next

	'Query date/time/runtime info
	objExcel.Cells(1, 11).Font.Bold = TRUE
	objExcel.Cells(2, 11).Font.Bold = TRUE
	objExcel.Cells(3, 11).Font.Bold = TRUE
	objExcel.Cells(4, 11).Font.Bold = TRUE
	ObjExcel.Cells(1, 11).Value = "Query date and time:"
	ObjExcel.Cells(2, 11).Value = "Query runtime (in seconds):"
	ObjExcel.Cells(3, 11).Value = "Total reviews:"
	ObjExcel.Cells(4, 11).Value = "Interview ERs with no CAF:"
	ObjExcel.Cells(1, 12).Value = now
	ObjExcel.Cells(2, 12).Value = timer - query_start_time
	ObjExcel.Cells(3, 12).Value = UBound(review_array, 2)
	ObjExcel.Cells(4, 12).Value = er_case_to_work

	'Formatting the columns to autofit after they are all finished being created.
	FOR i = 1 to 12
		objExcel.Columns(i).autofit()
	Next



	ObjExcel.Worksheets.Add().Name = "SR cases " & REPT_month & "-" & REPT_year & " "

	'formatting excel file with columns for case number and interview date/time
	objExcel.cells(1, 1).value 	= "X number"
	objExcel.cells(1, 2).value 	= "Case Number"
	objExcel.cells(1, 3).value 	= "Programs"
	objExcel.cells(1, 4).value 	= "Case language"
	objExcel.Cells(1, 5).value 	= "Interpreter"
	objExcel.cells(1, 6).value 	= "Phone # One"
	objExcel.cells(1, 7).value 	= "Phone # Two"
	objExcel.Cells(1, 8).value 	= "Phone # Three"

	FOR i = 1 to 8									'formatting the cells'
		objExcel.Cells(1, i).Font.Bold = True		'bold font'
		ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
		objExcel.Columns(i).AutoFit()				'sizing the columns'
	NEXT

	excel_row = 2
	sr_case_to_work = 0

	For revs_item = 0 to UBound(review_array, 2)
		If review_array(saved_to_excel_const, revs_item) = TRUE Then
			If review_array(interview_const, revs_item) = FALSE AND review_array(current_SR_const, revs_item) = TRUE Then
				'determining the programs list
				If ( review_array(SNAP_status_const, revs_item) = "TRUE" and review_array(MFIP_status_const, revs_item) = "TRUE" ) then
					programs_list = "SNAP & MFIP"
				elseif review_array(SNAP_status_const, revs_item) = "TRUE" then
					programs_list = "SNAP"
				elseif review_array(MFIP_status_const, revs_item) = "TRUE" then
					programs_list = "MFIP"
				End if
				'Excel output of Interview Required case information
				If review_array(notes_const, revs_item) <> "PRIV Case." then
					ObjExcel.Cells(excel_row, 1).value = review_array(worker_const,       revs_item)
					ObjExcel.Cells(excel_row, 2).value = review_array(case_number_const,  revs_item)
					ObjExcel.Cells(excel_row, 3).value = programs_list
					ObjExcel.Cells(excel_row, 4).value = review_array(Language_const,     revs_item)
					ObjExcel.Cells(excel_row, 5).value = review_array(Interpreter_const,  revs_item)
					ObjExcel.Cells(excel_row, 6).value = review_array( phone_1_const,     revs_item)
					ObjExcel.Cells(excel_row, 7).value = review_array( phone_2_const,     revs_item)
					ObjExcel.Cells(excel_row, 8).value = review_array( phone_3_const,     revs_item)
					sr_case_to_work = sr_case_to_work + 1
					excel_row = excel_row + 1
				End if
			End If
		End If
	Next

	'Query date/time/runtime info
	objExcel.Cells(1, 11).Font.Bold = TRUE
	objExcel.Cells(2, 11).Font.Bold = TRUE
	objExcel.Cells(3, 11).Font.Bold = TRUE
	objExcel.Cells(4, 11).Font.Bold = TRUE
	ObjExcel.Cells(1, 11).Value = "Query date and time:"
	ObjExcel.Cells(2, 11).Value = "Query runtime (in seconds):"
	ObjExcel.Cells(3, 11).Value = "Total reviews:"
	ObjExcel.Cells(4, 11).Value = "SRs with no Form:"
	ObjExcel.Cells(1, 12).Value = now
	ObjExcel.Cells(2, 12).Value = timer - query_start_time
	ObjExcel.Cells(3, 12).Value = UBound(review_array, 2)
	ObjExcel.Cells(4, 12).Value = sr_case_to_work

	'Formatting the columns to autofit after they are all finished being created.
	FOR i = 1 to 12
		objExcel.Columns(i).autofit()
	Next

	end_msg = "An Excel Workbook has been created with two lists of work:" & vbCr & vbCr & "The script found:" & vbCr & "  - " & er_case_to_work &" ER cases with no CAF entered in MAXIS" & vbCr & "  - " & sr_case_to_work &" SR cases with no CSR entered in MAXIS"
ElseIf renewal_option <> "Send NOMIs" Then
    end_msg = "Report not available yet."
End if

If renewal_option = "Send NOMIs" Then
	MAXIS_footer_month = CM_plus_1_mo							'Setting the footer month and year based on the review month. We do not run statistics in CM + 2
	MAXIS_footer_year = CM_plus_1_yr

	'creating a last day of recert variable - for NOMI this is the last day of the current month - which is determined here
	last_day_of_recert = CM_plus_1_mo & "/01/" & CM_plus_1_yr
	last_day_of_recert = dateadd("D", -1, last_day_of_recert)

	'creating the interview deadline date - this was the last day provided in the appointment notice
	interview_deadline_date = CM_mo & "/15/" & CM_yr

	'creating the response deadline date - this is the day client must respond by in order to prevent case closure
	response_deadline_date = CM_plus_1_mo & "/01/" & CM_plus_1_yr
	response_deadline_date = dateadd("D", -11, last_day_of_recert)

	'Activates worksheet based on user selection
	objExcel.worksheets(scenario_dropdown).Activate

	'Finding the last column that has something in it so we can add to the end.
	col_to_use = 0
	Do
		col_to_use = col_to_use + 1
		col_header = trim(ObjExcel.Cells(1, col_to_use).Value)
		If col_header = "NOMI Sent" Then notc_col = col_to_use
		If col_header = "NOMI Date" Then notc_date_col = col_to_use
		If col_header = "APPT NOTC Sent" Then appt_notc_col = col_to_use
		If col_header = "APPT NOTC Date" Then last_apt_notc_col = col_to_use + 1
	Loop until col_header = ""

	' MsgBox "NOTC Col - " & notc_col & vbCr & "NOTC Date Col - " & notc_date_col
	If notc_col = "" OR notc_date_col = "" Then
		last_apt_notc_col_letter = convert_digit_to_excel_column(last_apt_notc_col)

		'Insert columns in excel for additional information to be added
		column_place = last_apt_notc_col_letter & "1"
		Set objRange = objExcel.Range(column_place).EntireColumn

		If notc_date_col = "" Then objRange.Insert(xlShiftToRight)			'We neeed one more columns
		If notc_col = "" Then objRange.Insert(xlShiftToRight)			'We neeed one more columns

		If notc_date_col = "" Then
			' objRange.Insert(xlShiftToRight)			'We neeed one more columns
			notc_date_col = last_apt_notc_col + 1		'Setting the column to individual variables so we enter the found information in the right place

			ObjExcel.Cells(1, notc_date_col).Value = "NOMI Date"			'creating the column headers for the statistics information for the day of the run.
			objExcel.Cells(1, notc_date_col).Font.Bold = True		'bold font'
			ObjExcel.columns(notc_date_col).NumberFormat = "m/d/yy" 		'formatting as text
			ObjExcel.columns(notc_date_col).AutoFit() 						'fsizing the columns'
		End If
		If notc_col = "" Then
			' objRange.Insert(xlShiftToRight)			'We neeed one more columns
			notc_col = last_apt_notc_col 		'Setting the column to individual variables so we enter the found information in the right place
			' col_to_use = col_to_use + 1

			ObjExcel.Cells(1, notc_col).Value = "NOMI Sent"			'creating the column headers for the statistics information for the day of the run.
			objExcel.Cells(1, notc_col).Font.Bold = True		'bold font'
			ObjExcel.columns(notc_col).NumberFormat = "@" 		'formatting as text
			ObjExcel.columns(notc_col).AutoFit() 				'sizing the columns'
		End If

	End If
	' MsgBox "NOMI Sent Col - " & notc_col & vbCr &_
	' 	   "NOMI Date Col - " & notc_date_col & vbCr &_
	' 	   "APPT Sent Col - " & appt_notc_col & vbCr &_
	' 	   "APPT Date Col - " & last_apt_notc_col
	col_to_use = 1
	Do
		col_header = trim(ObjExcel.Cells(1, col_to_use).Value)
		If col_header = "CASH (" & date_header & ")" Then cash_stat_excel_col = col_to_use
		If col_header = "SNAP (" & date_header & ")" Then snap_stat_excel_col = col_to_use
		If col_header = "HC (" & date_header & ")" Then hc_stat_excel_col = col_to_use
		If col_header = "MAGI (" & date_header & ")" Then magi_stat_excel_col = col_to_use
		If col_header = "CAF Date (" & date_header & ")" Then recvd_date_excel_col = col_to_use
		If col_header = "Intvw Date (" & date_header & ")" Then intvw_date_excel_col = col_to_use

		col_to_use = col_to_use + 1
	Loop until col_header = ""

	today_mo = DatePart("m", date)
	today_mo = right("00" & today_mo, 2)

	today_day = DatePart("d", date)
	today_day = right("00" & today_day, 2)

	today_yr = DatePart("yyyy", date)
	today_yr = right(today_yr, 2)
	today_date = today_mo & "/" & today_day & "/" & today_yr

	'Now we loop through the whole Excel List and sending notices on the right cases
	excel_row = "2"		'starts at row 2'
	Do
		notc_col_info = trim(ObjExcel.Cells(excel_row, notc_col).Value)
		MAXIS_case_number 	= trim(ObjExcel.Cells(excel_row,  2).Value)			'getting the case number from the spreadsheet
		' MsgBox "row - " & excel_row & vbCr & "col - " & notc_col & vbCr & "val - *" & notc_col_info & "*"
		If notc_col_info = "" AND MAXIS_case_number <> "" Then
			' MsgBox excel_row
			forms_to_arep = ""
			forms_to_swkr = ""
			er_with_intherview = False
			MFIP_status = False
			SNAP_status = False
			appt_notc_sent = ""
			interview_date_as_of_today = ""
			caf_date_as_of_today = ""
			' programs = ""

			Call read_boolean_from_excel(ObjExcel.Cells(excel_row,  3).Value, er_with_intherview)
			Call read_boolean_from_excel(ObjExcel.cells(excel_row,  6).value, MFIP_status)
			Call read_boolean_from_excel(ObjExcel.cells(excel_row, 13).value, SNAP_status)

			appt_notc_sent = trim(ObjExcel.Cells(excel_row, appt_notc_col).Value)
			interview_date_as_of_today = trim(ObjExcel.Cells(excel_row, intvw_date_excel_col).Value)
			caf_date_as_of_today = trim(ObjExcel.Cells(excel_row, recvd_date_excel_col).Value)
			notes_info = Trim(ObjExcel.cells(excel_row, 25).value)

			send_nomi_now = False
			in_county = True

			' If er_with_intherview = True AND interview_date_as_of_today = "" AND appt_notc_sent = "Y" Then						'USING DIFFERENT CRITERA DUE TO WAIVED INTVW FOR SNAP DURING COVID-19'
			If er_with_intherview = True AND interview_date_as_of_today = "" AND appt_notc_sent = "Y" AND MFIP_status = True Then send_nomi_now = True

			If send_nomi_now = True Then
				If notes_info = "PRIV Case." Then send_nomi_now = False
				If notes_info <> "PRIV Case." then
					Call back_to_SELF
					Call navigate_to_MAXIS_screen("CASE", "CURR")
					EMReadscreen curr_primary_worker, 4, 21, 14
					If curr_primary_worker <> "X127" Then
						send_nomi_now = False
						in_county = False
					End If
				End If
			End If


			' If MFIP_status = True and SNAP_status = True Then
			' 	programs = "MFIP/SNAP"
			' ElseIf MFIP_status = True Then
			' 	programs = "MFIP"
			' ElseIf SNAP_status = True Then
			' 	programs = "SNAP"
			' End If

			If send_nomi_now = True Then
				Call start_a_new_spec_memo_and_continue(memo_started)

				IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY

					if caf_date_as_of_today <> "" then CALL write_variable_in_SPEC_MEMO("We received your Recertification Paperwork on " & caf_date_as_of_today & ".")
					if caf_date_as_of_today = "" then CALL write_variable_in_SPEC_MEMO("Your Recertification Paperwork has not yet been received.")
					CALL write_variable_in_SPEC_MEMO("")
					' CALL write_variable_in_SPEC_MEMO("You must have an interview by " & last_day_of_recert & " or your benefits will end. ")
					CALL write_variable_in_SPEC_MEMO("You may still need an interview by " & last_day_of_recert & " or your benefits will end. ")
					CALL write_variable_in_SPEC_MEMO("")
					Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
					Call write_variable_in_SPEC_MEMO("612-596-1300 between 8:00am and 4:30pm Monday thru Friday.")
					CALL write_variable_in_SPEC_MEMO("")
					'removal of in person verbiage during the COVID-19 PEACETIME STATE OF EMERGENCY
					' Call write_variable_in_SPEC_MEMO("If you wish to schedule an interview, call 612-596-1300. You may also come to any of the six offices below for an in-person interview between 8 and 4:30, Monday thru Friday.")
					' Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
					' Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
					' Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30 ")
					' Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
					' Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
					' Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
					' Call write_variable_in_SPEC_MEMO("(Hours are M - F 8-4:30 unless otherwise noted)")
					CALL write_variable_in_SPEC_MEMO("** You can submit documents Online at www.MNBenefits.org **")
					CALL write_variable_in_SPEC_MEMO("Other options for submitting documents to Hennepin County:")
					CALL write_variable_in_SPEC_MEMO(" - Mail, Fax, or Drop Boxes at service centers")
					CALL write_variable_in_SPEC_MEMO(" - Email with document attachment.EMAIL: hhsews@hennepin.us")
					CALL write_variable_in_SPEC_MEMO("   (Only attach PNG, JPG, TIF, DOC, PDF, or HTM file types)")
					CALL write_variable_in_SPEC_MEMO("")
					CALL write_variable_in_SPEC_MEMO("  ** If we do not hear from you by " & last_day_of_recert & "  **")
					CALL write_variable_in_SPEC_MEMO("  **   your benefits will end on " & last_day_of_recert & ".   **")


					PF4         'Submit the MEMO

					memo_row = 7                                            'Setting the row for the loop to read MEMOs
					ObjExcel.Cells(excel_row, notc_col).Value = "N"         'Defaulting this to 'N'
					Do
						EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
						EMReadScreen print_status, 7, memo_row, 67
						If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
							ObjExcel.Cells(excel_row, notc_col).Value = "Y"             'If we've found this then no reason to keep looking.
							ObjExcel.Cells(excel_row, notc_date_col).Value = today_date        'If we've found this then no reason to keep looking.
							successful_notices = successful_notices + 1                 'For statistical purposes
							Exit Do
						End If

						memo_row = memo_row + 1           'Looking at next row'
					Loop Until create_date = "        "
				ELSE
					ObjExcel.Cells(excel_row, notc_col).Value = "N"         'Setting this as N if the MEMO failed
					call back_to_SELF
				END IF


				If ObjExcel.Cells(excel_row, notc_col).Value = "Y" Then

					start_a_blank_case_note
					CALL write_variable_in_CASE_NOTE("*** NOMI Sent for " & REPT_month & "/" & REPT_year & " Recertification***")
					if caf_date_as_of_today <> "" then CALL write_variable_in_CASE_NOTE("* Recertification app received on " & caf_date_as_of_today)
					if caf_date_as_of_today = "" then CALL write_variable_in_CASE_NOTE("* Recertification app has NOT been received. Client must submit paperwork.")
					CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about how to call in for an interview.")
					CALL write_variable_in_CASE_NOTE("* Client must call 612-596-1300 to complete interview.")
					If forms_to_arep = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")
					If forms_to_swkr = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")
					call write_variable_in_case_note("---")
					call write_variable_in_case_note(worker_signature)

					PF3
				End If

			ElseIf er_with_intherview = True  AND MFIP_status = False Then
				ObjExcel.Cells(excel_row, notc_col).Value = "Not MFIP"
			ElseIf er_with_intherview = True AND interview_date_as_of_today <> "" Then
				ObjExcel.Cells(excel_row, notc_col).Value = "INTV Done"
			ElseIf er_with_intherview = True AND MFIP_status = TRUE AND appt_notc_sent <> "Y" Then
				ObjExcel.Cells(excel_row, notc_col).Value = "Check APPT NOTC"
			ElseIf er_with_intherview = False Then
				ObjExcel.Cells(excel_row, notc_col).Value = "N/A"
			ElseIf in_county = False Then
				ObjExcel.Cells(excel_row, notc_col).Value = "Out of County - " & right(curr_primary_worker, 2)
			End If
		End If
		excel_row = excel_row + 1
	Loop until MAXIS_case_number = ""

	'Finding all of the worksheets available in the file. We will likely open up the main 'Review Report' so the script will default to that one.
	For Each objWorkSheet In objWorkbook.Worksheets
		If objWorkSheet.Name = "NOTICES" Then
			objWorkSheet.Activate
			Exit For
		End If
	Next

	entry_row = 1

	objExcel.Cells(entry_row, 4).Value      = "NOMIs run on:"     'Date and time the script was completed
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value      = now
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 4).Value      = "Runtime (in seconds)"            'Enters the amount of time it took the script to run
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value      = timer - query_start_time
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 4).Value      = "Total Cases assesed"             'All cases from the spreadsheet
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value    	= excel_row - 2
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 4).Value      = "Total Cases with ER Interview"             'All cases from the spreadsheet
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value      = "=COUNTIFS(Table1[Interview ER],"&is_true&")"
	' total_row = entry_row
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 4).Value      = "Total MFIP Cases with ER Interview"             'All cases from the spreadsheet
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value      = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[MFIP Status],"&is_true&")"
	total_row = entry_row
	entry_row = entry_row + 1

	if successful_notices = "" then successful_notices = 0
	objExcel.Cells(entry_row, 4).Value      = "NOMIs Sent"        'number of notices that were successful
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value      = "=COUNTIFS(Table1[NOMI Sent]," & Chr(34) & "Y" & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
	appt_row = entry_row
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 4).Value      = "Percentage INTV Cases with NOMIs"           'calculation of the percent of successful notices
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value      = "=E" & appt_row & "/E" & total_row
	objExcel.Cells(entry_row, 5).NumberFormat = "0.00%"		'Formula should be percent
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 4).Value      = "Interviews Completed"        'number of notices that were successful
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value      = "=COUNTIFS(Table1[NOMI Sent]," & Chr(34) & "INTV Done" & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
	intv_row = entry_row
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 4).Value      = "Percentage INTV Done"           'calculation of the percent of successful notices
	objExcel.Cells(entry_row, 4).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 5).Value      = "=E" & intv_row & "/E" & total_row
	objExcel.Cells(entry_row, 5).NumberFormat = "0.00%"		'Formula should be percent
	entry_row = entry_row + 1

	ObjExcel.columns(4).AutoFit()
	ObjExcel.columns(5).AutoFit()
	end_msg = end_msg & vbCr & vbCr & "NOMIs have been sent on " & successful_notices & " cases today. Information added to the Review Report Excel Document"
End If

STATS_counter = STATS_counter - 1
script_end_procedure(end_msg)
