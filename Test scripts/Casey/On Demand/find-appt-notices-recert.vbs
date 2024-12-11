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

excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\11-24 Review Report.xlsx"

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(excel_file_path, True, True, ObjExcel, objWorkbook)

'Activates worksheet based on user selection
objExcel.worksheets("11-24 Review Report").Activate

MsgBox "WAIT and See"

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
today_date = #9/25/2024#
too_old_date = #9/24/2024#

'Now we loop through the whole Excel List and sending notices on the right cases
excel_row = "2"		'starts at row 2'
Do

	notc_col_info = trim(ObjExcel.Cells(excel_row, notc_col).Value)
	MAXIS_case_number 	= trim(ObjExcel.Cells(excel_row,  2).Value)			'getting the case number from the spreadsheet
	notice_found = "N"
	notice_date = ""
	' MsgBox "row - " & excel_row & vbCr & "col - " & notc_col & vbCr & "val - *" & notc_col_info & "*"
	If notc_col_info = "" AND MAXIS_case_number <> "" Then

		Call navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)
		If is_this_priv = True Then notice_found = "N/A"
		If is_this_priv = False Then

			note_row = 5
			Do
				EMReadScreen note_date, 8, note_row, 6                  'reading the note date

				EMReadScreen note_title, 55, note_row, 25               'reading the note header
				note_title = trim(note_title)

				If InStr(note_title, "*** Notice of") <> 0 and InStr(note_title, "Recertification Interview") <> 0 Then
					notice_found = "Y"
					notice_date = note_date
				End If
				If InStr(note_title, "Renewal Guidance") <> 0 Then
					notice_found = "RG"
					' notice_date = note_date
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
		ObjExcel.Cells(excel_row, notc_col).Value = notice_found
		ObjExcel.Cells(excel_row, notc_date_col).Value = notice_date
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
	total_row = entry_row
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 1).Value      = "Total MFIP Cases with ER Interview"             'All cases from the spreadsheet
	objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 2).Value      = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[MFIP Status],"&is_true&")"
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

MsgBox "DONE?"
MsgBox "No really, stop"







' 	notc_col_info = trim(ObjExcel.Cells(excel_row, notc_col).Value)
' 	MAXIS_case_number 	= trim(ObjExcel.Cells(excel_row,  2).Value)			'getting the case number from the spreadsheet
' 	' MsgBox "row - " & excel_row & vbCr & "col - " & notc_col & vbCr & "val - *" & notc_col_info & "*"
' 	If notc_col_info = "" AND MAXIS_case_number <> "" Then
' 		send_appt_notc = True
' 		If check_REVW = True Then
' 			send_appt_notc = False
' 			Call check_for_MAXIS(FALSE)		'making sure we haven't passworded out

' 			' MAXIS_case_number = case_number_to_check		'setting the case number for NAV functions
' 			call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)		'Go to STAT REVW and be sure the case is not privleged.
' 			If is_this_priv = FALSE Then
' 				EMReadScreen recvd_date, 8, 13, 37										'Reading the CAF Received Date and format
' 				recvd_date = replace(recvd_date, " ", "/")
' 				if recvd_date = "__/__/__" then recvd_date = ""

' 				EMReadScreen interview_date, 8, 15, 37									'Reading the interview date and format
' 				interview_date = replace(interview_date, " ", "/")
' 				if interview_date = "__/__/__" then interview_date = ""

' 				EMReadScreen cash_review_status, 1, 7, 40								'Reading the review status and format
' 				EMReadScreen snap_review_status, 1, 7, 60
' 				EMReadScreen hc_review_status, 1, 7, 73
' 				If cash_review_status = "_" Then cash_review_status = ""
' 				If snap_review_status = "_" Then snap_review_status = ""
' 				If hc_review_status = "_" Then hc_review_status = ""

' 				revw_status_all_n = True
' 				If cash_review_status = "I" OR cash_review_status = "U" OR cash_review_status = "A" Then revw_status_all_n = False
' 				If snap_review_status = "I" OR snap_review_status = "U" OR snap_review_status = "A" Then revw_status_all_n = False
' 				If hc_review_status = "I" OR hc_review_status = "U" OR hc_review_status = "A" Then revw_status_all_n = False

' 				If recvd_date = "" AND revw_status_all_n = True Then send_appt_notc = True
' 			End If
' 		End If
' 		If send_appt_notc = False Then ObjExcel.Cells(excel_row, notc_col).Value = "N/A"

' 		If send_appt_notc = True or check_REVW = False Then
' 			' MsgBox excel_row

' 			forms_to_arep = ""
' 			forms_to_swkr = ""
' 			programs = ""
' 			intvw_programs = ""
' 			renewal_guidance_needed = False
' 			renewal_guidance_confirmed = False

' 			Call read_boolean_from_excel(ObjExcel.Cells(excel_row,  3).Value, er_with_intherview)
' 			Call read_boolean_from_excel(objExcel.cells(excel_row,  6).value, MFIP_status)
' 			Call read_boolean_from_excel(objExcel.cells(excel_row,  7).value, DWP_status)
' 			Call read_boolean_from_excel(objExcel.cells(excel_row,  8).value, GA_status)
' 			Call read_boolean_from_excel(objExcel.cells(excel_row,  9).value, MSA_status)
' 			Call read_boolean_from_excel(objExcel.cells(excel_row, 10).value, GRH_status)
' 			Call read_boolean_from_excel(objExcel.cells(excel_row, 13).value, SNAP_status)

' 			REPT_full = REPT_month & "/" & REPT_year
' 			CASH_SR_Info = trim(objExcel.cells(excel_row, 11).value)
' 			CASH_ER_Info = trim(objExcel.cells(excel_row, 12).value)
' 			SNAP_SR_Info = trim(objExcel.cells(excel_row, 14).value)
' 			SNAP_ER_Info = trim(objExcel.cells(excel_row, 15).value)

' 			If MFIP_status = True and SNAP_status = True Then
' 				intvw_programs = "MFIP/SNAP"
' 			ElseIf MFIP_status = True Then
' 				intvw_programs = "MFIP"
' 			ElseIf SNAP_status = True Then
' 				intvw_programs = "SNAP"
' 			End If
' 			If CASH_ER_Info = REPT_full then
' 				If MFIP_status = True Then programs = programs & "/MFIP"
' 				If DWP_status = True Then programs = programs & "/DWP"
' 				If GA_status = True Then programs = programs & "/GA"
' 				If MSA_status = True Then programs = programs & "/MSA"
' 			End If
' 			If CASH_SR_Info = REPT_full OR CASH_ER_Info = REPT_full then
' 				If GRH_status = True Then programs = programs & "/GRH"
' 			End If
' 			If SNAP_SR_Info = REPT_full OR SNAP_ER_Info = REPT_full then
' 				If SNAP_status = True Then programs = programs & "/SNAP"
' 			End If
' 			If left(programs, 1) = "/" Then programs = right(programs, len(programs)-1)
' 			' MsgBox "Sending NOTC - " & MAXIS_case_number & " - excel row - " & excel_row & " Programs - " & programs

' 			interview_end_date = CM_plus_1_mo & "/15/" & CM_plus_1_yr
' 			last_day_of_recert = CM_plus_2_mo & "/01/" & CM_plus_2_yr
' 			last_day_of_recert = dateadd("D", -1, last_day_of_recert)

' 			notes_info = Trim(ObjExcel.cells(excel_row, 25).value)

' 			' If er_with_intherview = True Then
' 			If er_with_intherview = True AND MFIP_status = True Then
' 				'Writing the SPEC MEMO - dates will be input from the determination made earlier.
' 				' MsgBox "We're writing a MEMO here"
' 				Call start_a_new_spec_memo_and_continue(memo_started)

' 				IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY
' 					CALL create_appointment_letter_notice_recertification(programs, intvw_programs, interview_end_date, last_day_of_recert)

' 					memo_row = 7                                            'Setting the row for the loop to read MEMOs
' 					ObjExcel.Cells(excel_row, notc_col).Value = "N"         'Defaulting this to 'N'
' 					Do
' 						EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
' 						EMReadScreen print_status, 7, memo_row, 67
' 						If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
' 							ObjExcel.Cells(excel_row, notc_col).Value = "Y"             'If we've found this then no reason to keep looking.
' 							ObjExcel.Cells(excel_row, notc_date_col).Value = today_date        'If we've found this then no reason to keep looking.
' 							successful_notices = successful_notices + 1                 'For statistical purposes
' 							Exit Do
' 						End If

' 						memo_row = memo_row + 1           'Looking at next row'
' 					Loop Until create_date = "        "

' 				ELSE
' 					ObjExcel.Cells(excel_row, notc_col).Value = "N"         'Setting this as N if the MEMO failed
' 					call back_to_SELF
' 				END IF
' 				' ObjExcel.Cells(excel_row, 25).Value = "All progs - " & programs & " : INTVW Progs - " & intvw_programs
' 			Else
' 				ObjExcel.Cells(excel_row, notc_col).Value = "N/A"
' 				renewal_guidance_needed = True
' 				If notes_info = "PRIV Case." then renewal_guidance_needed = False
' 			End If

' 			If ObjExcel.Cells(excel_row, notc_col).Value = "Y" OR renewal_guidance_needed = True Then

' 				Call start_a_new_spec_memo_and_continue(memo_started)   'Starting a MEMO to send information about verifications

' 				IF memo_started = True THEN
' 					If renewal_guidance_needed = True Then
' 						CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case and is due by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ".")
' 						' CALL write_variable_in_SPEC_MEMO("")
' 					End If

' 					CALL write_variable_in_SPEC_MEMO("As a part of the Renewal Process we must receive recent verification of your information. To speed the renewal process, please send proofs with your renewal paperwork.")
' 					CALL write_variable_in_SPEC_MEMO("")
' 					CALL write_variable_in_SPEC_MEMO(" * Examples of income proofs: paystubs, employer statement,")
' 					CALL write_variable_in_SPEC_MEMO("   income reports, business ledgers, income tax forms, etc.")
' 					CALL write_variable_in_SPEC_MEMO("   *If a job has ended, send proof of the end of employment")
' 					CALL write_variable_in_SPEC_MEMO("   and last pay.")
' 					CALL write_variable_in_SPEC_MEMO("")
' 					CALL write_variable_in_SPEC_MEMO(" * Examples of housing cost proofs(if changed): rent/house")
' 					CALL write_variable_in_SPEC_MEMO("   payment receipt, mortgage, lease, subsidy, etc.")
' 					CALL write_variable_in_SPEC_MEMO("")
' 					CALL write_variable_in_SPEC_MEMO(" * Examples of medical cost proofs(if changed):")
' 					CALL write_variable_in_SPEC_MEMO("   prescription and medical bills, etc.")
' 					CALL write_variable_in_SPEC_MEMO("")
' 					If renewal_guidance_needed = False Then CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")
' 					If renewal_guidance_needed = True Then
' 						CALL digital_experience
' 						CALL write_variable_in_SPEC_MEMO("Once we receive and process your renewal paperwork, you will receive information BY MAIL with possible follow up or actions taken on your case. Call 612-596-1300 if you have additional questions.")
' 					End If

' 					PF4 'Submit the MEMO'

' 					If renewal_guidance_needed = True Then
' 						memo_row = 7                                            'Setting the row for the loop to read MEMOs
' 						Do
' 							EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
' 							EMReadScreen print_status, 7, memo_row, 67
' 							If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
' 								renewal_guidance_confirmed = True
' 								ObjExcel.Cells(excel_row, notc_col).Value = "RG"
' 								' ObjExcel.Cells(excel_row, 25).Value = "All progs - " & programs & " : INTVW Progs - " & intvw_programs
' 								Exit Do
' 							End If

' 							memo_row = memo_row + 1           'Looking at next row'
' 						Loop Until create_date = "        "
' 					End If
' 				End If

' 				If ObjExcel.Cells(excel_row, notc_col).Value = "Y" Then
' 					start_a_blank_case_note
' 					CALL write_variable_in_CASE_NOTE("*** Notice of " & intvw_programs & " Recertification Interview Sent ***")
' 					CALL write_variable_in_case_note("* A notice has been sent to client with detail about how to call in for an interview.")
' 					CALL write_variable_in_case_note("* Client must submit paperwork and call 612-596-1300 to complete interview.")
' 					If forms_to_arep = "Y" then call write_variable_in_case_note("* Copy of notice sent to AREP.")
' 					If forms_to_swkr = "Y" then call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
' 					call write_variable_in_case_note("---")
' 					CALL write_variable_in_case_note("Link to Domestic Violence Brochure sent to client in SPEC/MEMO as a part of interview notice.")
' 					call write_variable_in_case_note("---")
' 					call write_variable_in_case_note(worker_signature)
' 				ElseIf renewal_guidance_confirmed = True Then
' 					start_a_blank_case_note
' 					CALL write_variable_in_CASE_NOTE("Notice sent for " & programs & " Renewal Guidance")
' 					Call write_variable_in_case_note("* A renewal is due for this case for " & REPT_month & "/" & REPT_year)
' 					Call write_variable_in_case_note("* Reminder notice sent with forms due date and verification options.")
' 					Call write_variable_in_case_note("  -This is NOT an official verification request.-")
' 					Call write_variable_in_case_note("---")
' 					Call write_variable_in_case_note(worker_signature)
' 				End If
' 				PF3
' 			End If
' 		End If
' 	End If
' 	excel_row = excel_row + 1
' Loop until MAXIS_case_number = ""

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
	total_row = entry_row
	entry_row = entry_row + 1

	objExcel.Cells(entry_row, 1).Value      = "Total MFIP Cases with ER Interview"             'All cases from the spreadsheet
	objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 2).Value      = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[MFIP Status],"&is_true&")"
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
