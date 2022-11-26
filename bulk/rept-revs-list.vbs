'Required for statistical purposes===============================================================================
name_of_script = "BULK - REPT-REVS LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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
call changelog_update("11/26/2022", "Updated date handling and selection.", "Ilse Ferris, Hennepin County") ''#1060
call changelog_update("06/27/2018", "Added/updated closing message.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/12/2018", "Entering a supervisor X-Number in the Workers to Check will pull all X-Numbers listed under that supervisor in MAXIS. Addiional bug fix where script was missing cases.", "Casey Love, Hennepin County")
call changelog_update("09/25/2017", "Added handling for all months. Previously script only allowed user to select from current month or current month plus 2.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'This function is used to grab all active X numbers according to the supervisor X number(s) inputted
FUNCTION create_array_of_all_active_x_numbers_by_supervisor(array_name, supervisor_array)
	'Getting to REPT/USER
	CALL navigate_to_MAXIS_screen("REPT", "USER")
	'Sorting by supervisor
	PF5
	PF5
	'Reseting array_name
	array_name = ""
	'Splitting the list of inputted supervisors...
	supervisor_array = replace(supervisor_array, " ", "")
	supervisor_array = split(supervisor_array, ",")
	FOR EACH unit_supervisor IN supervisor_array
		IF unit_supervisor <> "" THEN
			'Entering the supervisor number and sending a transmit
			CALL write_value_and_transmit(unit_supervisor, 21, 12)
			MAXIS_row = 7
			DO
				EMReadScreen worker_ID, 8, MAXIS_row, 5
				worker_ID = trim(worker_ID)
				IF worker_ID = "" THEN EXIT DO
				array_name = trim(array_name & " " & worker_ID)
				MAXIS_row = MAXIS_row + 1
				IF MAXIS_row = 19 THEN
					PF8
					EMReadScreen end_check, 9, 24,14
					If end_check = "LAST PAGE" Then Exit Do
					MAXIS_row = 7
				END IF
			LOOP
		END IF
	NEXT
	'Preparing array_name for use...
	array_name = split(array_name)
END FUNCTION

'THE SCRIPT---------------------------------------------------

EMConnect "" 'Connects to BlueZone
Call check_for_MAXIS(False)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)		'inputs the MAXIS footer/month and year that is on the current MAXIS screen
day_of_month = DatePart("D", date)

review_month = MAXIS_footer_month
review_year = MAXIS_footer_year

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 291, 150, "REPT/REVS"
  EditBox 140, 10, 140, 15, worker_number
  CheckBox 135, 40, 160, 10, "OR check here to run this query county-wide.", all_workers_check
  EditBox 140, 55, 20, 15, review_month
  EditBox 165, 55, 20, 15, review_year
  CheckBox 10, 25, 40, 10, "SNAP?", SNAP_check
  CheckBox 10, 40, 50, 10, "Cash/GRH?", cash_check
  CheckBox 10, 55, 40, 10, "HC?", HC_check
  CheckBox 10, 95, 215, 10, "Read info from CASE/NOTE to identify potential renewal errors.", error_check
  CheckBox 10, 110, 210, 10, "Read SNAP/MFIP info from Auto-close notices.", notice_audit_check
  ButtonGroup ButtonPressed
    OkButton 195, 130, 40, 15
    CancelButton 240, 130, 40, 15
  Text 75, 30, 220, 10, "Enter the 7-digit worker numbers (ex: X127###), comma separated."
  Text 75, 60, 65, 10, "Report Month/Year:"
  GroupBox 5, 10, 60, 60, "Program(s):"
  Text 75, 15, 65, 10, "Worker(s) to check:"
  GroupBox 5, 80, 275, 45, "Additional Reporting Options:"
EndDialog

Do
	Do
		err_msg = ""
		Dialog Dialog1
		Cancel_without_confirmation
        Call validate_footer_month_entry(review_month, review_year, err_msg, "*")
        If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
        If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
        If review_month = CM_plus_2_mo and review_year = CM_plus_2_yr then
            If day_of_month < 16 then
                err_msg = err_msg & vbcr & "* You cannot select current month plus 2 until the 16th of the month."
            Else
                current_month_plus_2 = True
            End if
        End if
		If SNAP_check = 0 and cash_check = 0 and HC_check = 0 then err_msg = err_msg & vbNewLine & "* Select at least one program."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

review_date = review_month & "/01/" & review_year

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 3).Font.Bold = TRUE

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 4 'Starting with 4 because cols 1-3 are already used

If SNAP_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "SNAP?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	snap_actv_col = col_to_use
	col_to_use = col_to_use + 1
	SNAP_letter_col = convert_digit_to_excel_column(snap_actv_col)
End if
If cash_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "CASH?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	cash_actv_col = col_to_use
	col_to_use = col_to_use + 1
	cash_letter_col = convert_digit_to_excel_column(cash_actv_col)
End if
If HC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "HC?"		'First does HC col, then does exempt IR col, then MAGI col
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	HC_actv_col = col_to_use
	col_to_use = col_to_use + 1
	HC_letter_col = convert_digit_to_excel_column(HC_actv_col)
	ObjExcel.Cells(1, col_to_use).Value = "EXEMPT HC IR?"	'Exempt IR col
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	exempt_IR_col = col_to_use
	col_to_use = col_to_use + 1
	exempt_IR_letter_col = convert_digit_to_excel_column(exempt_IR_col)
	ObjExcel.Cells(1, col_to_use).Value = "MAGI?"		'Here's that MAGI col
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	MAGI_col = col_to_use
	col_to_use = col_to_use + 1
	MAGI_letter_col = convert_digit_to_excel_column(MAGI_col)
End if

'Does these two columns after the others, because they appear that way in the screen, but are always used.
'Only does these if current_month_plus_2 = False, as you cannot have a revw rec'd or an interview date before this point.
If current_month_plus_2 = False then
	ObjExcel.Cells(1, col_to_use).Value = "DATE REVW REC'D"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	revw_recd_date_col = col_to_use
	col_to_use = col_to_use + 1
	revw_recd_date_letter_col = convert_digit_to_excel_column(revw_recd_date_col)
	ObjExcel.Cells(1, col_to_use).Value = "INTERVIEW DATE"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	interview_date_col = col_to_use
	col_to_use = col_to_use + 1
	interview_date_letter_col = convert_digit_to_excel_column(interview_date_col)
End if

IF notice_audit_check = checked Then
	objExcel.Cells(1, col_to_use).value = "AUTOCLOSE NOTICE"
	objExcel.Cells(1, col_to_use).Font.Bold = True
	notice_column = col_to_use
	col_to_use = col_to_use + 1
END IF

IF error_check = checked Then
	objExcel.Cells(1, col_to_use).value = "POTENTIAL ERROR"
	objExcel.Cells(1, col_to_use).Font.Bold = True
	error_column = col_to_use
	col_to_use = col_to_use + 1
END IF

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else		'If worker numbers are litsted - this will create an array of workers to check
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'formatting array
	For each x1_number in x1s_from_dialog
		x1_number = trim(ucase(x1_number))					'Formatting the x numbers so there are no errors
		Call navigate_to_MAXIS_screen ("REPT", "USER")		'This part will check to see if the x number entered is a supervisor of anyone
		PF5
		PF5
		EMWriteScreen x1_number, 21, 12
		transmit
		EMReadScreen sup_id_check, 7, 7, 5					'This is the spot where the first person is listed under this supervisor
		IF sup_id_check <> "       " Then 					'If this frist one is not blank then this person is a supervisor
			supervisor_array = trim(supervisor_array & " " & x1_number)		'The script will add this x number to a list of supervisors
		Else
			If worker_array = "" then						'Otherwise this x number is added to a list of workers to run the script on
				worker_array = trim(x1_number)
			Else
				worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
			End if
		End If
		PF3
	Next

	If supervisor_array <> "" Then 				'If there are any x numbers identified as a supervisor, the script will run the function above
		Call create_array_of_all_active_x_numbers_by_supervisor (more_workers_array, supervisor_array)
		workers_to_add = join(more_workers_array, ", ")
		If worker_array = "" then				'Adding all x numbers listed under the supervisor to the worker array
			worker_array = workers_to_add
		Else
			worker_array = worker_array & ", " & trim(ucase(workers_to_add))
		End if
	End If

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Setting the variable for what's to come
excel_row = 2
all_case_numbers_array = "*"

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
    MAXIS_footer_month = "" 'clearing variable to prevent breaking when in Cm+2
	MAXIS_footer_year = ""
	EMWriteScreen MAXIS_footer_month, 20, 43 'needs to add date that isn't CM+2 other wise script cannot navigate back to REVS when running on multiple cases.
	Call write_value_and_transmit(MAXIS_footer_year, 20, 46)

	Call navigate_to_MAXIS_screen("REPT", "REVS")
    EmWriteScreen review_month, 20, 55        'Entering user selected review month/year from dialog
    Call write_value_and_transmit(review_year, 20, 58)
    Call write_value_and_transmit(worker, 21, 6)

	EMReadScreen has_content_check, 8, 7, 6    'Skips workers with no info
	If trim(has_content_check) <> "" then
		Do  'Grabbing each case number on screen
			MAXIS_row = 7    'Set variable for next do...loop
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6			'Reading case number
				EMReadScreen client_name, 15, MAXIS_row, 16		'Reading client name
				EMReadScreen cash_status, 2, MAXIS_row, 39		'Reading cash status
				EMReadScreen SNAP_status, 1, MAXIS_row, 45		'Reading SNAP status
				EMReadScreen HC_status, 1, MAXIS_row, 49			'Reading HC status
				EMReadScreen exempt_IR_status, 1, MAXIS_row, 51		'Reading exempt IR status
				EMReadScreen MAGI_status, 8, MAXIS_row, 54		'Reading MAGI status
				EMReadScreen revw_recd_date, 8, MAXIS_row, 62		'Reading review received date
				EMReadScreen interview_date, 8, MAXIS_row, 72		'Reading interview date

				'Certain MAXIS users have inadvertently bent the laws of MAXIS physics and have cash_status in column 34 instead of 35. Here-in lies the fix.
				cash_status = trim(replace(cash_status, " ", ""))

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				MAXIS_case_number = trim(MAXIS_case_number)
				If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")
				If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end

				'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
				If cash_status = "-" then cash_status = ""
				If SNAP_status = "-" then SNAP_status = ""
				If HC_status = "-" then HC_status = ""

				'The asterisk in the exempt IR column messes up the formula for Excel. Replacing with the word "exempt"
				If exempt_IR_status = "*" then exempt_IR_status = "Exempt"

				'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
				If T_check = checked THEN
					IF trim(case_status) = "T" and cash_check = checked THEN add_case_info_to_Excel = True
					If trim(SNAP_status) = "T" and SNAP_check = checked then add_case_info_to_Excel = True
					If trim(HC_status) = "T" and HC_check = checked then add_case_info_to_Excel = True
				ELSE
					If trim(cash_status) <> "" and cash_check = checked then add_case_info_to_Excel = True
					If trim(SNAP_status) <> "" and SNAP_check = checked then add_case_info_to_Excel = True
					If trim(HC_status) <> "" and HC_check = checked then add_case_info_to_Excel = True
				END IF

				'Cleaning up the blank revw_recd_date and interview_date variables
				revw_recd_date = trim(replace(revw_recd_date, "__ __ __", ""))
				interview_date = trim(replace(interview_date, "__ __ __", ""))

				'Adding the case to Excel
				If add_case_info_to_Excel = True then
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					If current_month_plus_2 = False then
						ObjExcel.Cells(excel_row, revw_recd_date_col).Value = replace(revw_recd_date, " ", "/")
						ObjExcel.Cells(excel_row, interview_date_col).Value = replace(interview_date, " ", "/")
					End if
					If SNAP_check = checked then ObjExcel.Cells(excel_row, snap_actv_col).Value = trim(SNAP_status)
					If cash_check = checked then ObjExcel.Cells(excel_row, cash_actv_col).Value = trim(cash_status)
					If HC_check = checked then
						ObjExcel.Cells(excel_row, HC_actv_col).Value = trim(HC_status)
						ObjExcel.Cells(excel_row, exempt_IR_col).Value = trim(exempt_IR_status)
						ObjExcel.Cells(excel_row, MAGI_col).Value = trim(MAGI_status)
					End if
					excel_row = excel_row + 1
				End if
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
next

'Going to the top of the list and checking each case for additional information
row_to_use = 2
'if notice_audit_check = checked THEN col_to_use = col_to_use + 1
IF notice_audit_check = checked OR error_check = checked THEN
	DO
		MAXIS_case_number = objExcel.Cells(row_to_use, 2).Value
		IF MAXIS_case_number <> "" THEN
			IF SNAP_check = checked THEN 'Checking SNAP notices only
				IF notice_audit_check = checked THEN 'THE FOLLOWING CHECKS THE CONTENTS OF AUTOCLOSE NOTICES
					IF objExcel.Cells(row_to_use, snap_actv_col).Value = "T"  OR objExcel.Cells(row_to_use, snap_actv_col).Value = "N" THEN 'Only concerned with notices for incomplete or terminated cases.
					'First it will read the SNAP autoclose notice and check the closure reasons listed.
						call navigate_to_MAXIS_screen("SPEC", "WCOM")
						row = 7
						col = 1
						DO
						EMSearch "Autoclose Notice", row, col
						IF row <> 0 THEN
							EMReadScreen prg_typ, 2, row, 26
							IF prg_typ = "FS" or prg_typ = "MF" THEN 'found a snap autoclose, need to read it
								EMWriteScreen "X", row, 13
								Transmit
								PF8 'skipping the first page, the important stuff is on page 2
								row = 10
								col = 1
								EMSearch "Report Form", row, col 'checking for CSRs 1st as they all use same notice.
								IF row <> 0 THEN
									notice_reason = "CSR not complete."
									PF3
									EXIT DO
								ELSE 'not a csr notice, read the closure reasons
									row = 10
									col = 1
									EMSearch "interview", row, col
									IF row <> 0 THEN notice_reason = "No interview."
									row = 10
									col = 1
									EMSearch "asked for", row, col
									IF row <> 0 THEN notice_reason = notice_reason & " Proofs."
									row = 10
									col = 1-3
									EMSearch "redetermination form", row, col
									IF row <> 0 THEN notice_reason = notice_reason & " No CAF."
									IF notice_reason = "" THEN notice_reason = "Check manually."
									PF3
									EXIT DO
								END IF
							END IF
							row = row + 1
						END IF
						If row = 0 THEN notice_reason = "No notice found, check case manually." 'There is likely an error on this case.
						LOOP UNTIL row = 21 or row = 0
						objExcel.Cells(row_to_use, notice_column).Value = notice_reason
					END IF
				END IF
				'This section checks for potential errors before notices are sent
				IF error_check = checked THEN
					IF ObjExcel.Cells(row_to_use, snap_actv_col).Value = "N" OR objExcel.cells(row_to_use, snap_actv_col).Value = "I" THEN
						call navigate_to_MAXIS_screen("CASE", "NOTE")
						row = 1
						col = 1
						IF ObjExcel.Cells(row_to_use, snap_actv_col).Value = "N" THEN
							DO
								IF ObjExcel.Cells(row_to_use, snap_actv_col).Value = "N" THEN EMSearch "received", row, col 'Looking for a case note for CSR Received or RECERT CAF received.
								IF row = 0 THEN EXIT DO
								IF row <> 0 THEN
									EMReadScreen case_note_date, 10, row, 6
									IF datediff("d", case_note_date, review_date) > 45 THEN EXIT DO 'We are only concerned with stuff received in the 45 days before recert.
									EMReadScreen case_note_type, 4, row, col - 4
									IF case_note_type = "CAF " THEN
										error_content = "Check for possible RECERT received."
										EXIT DO
									ELSEIF case_note_type = "CSR " THEN
										error_content = "Check for possible CSR received."
										EXIT DO
									ELSE
										row = row + 1
									END IF
								END IF
							LOOP UNTIL row = 19
						END IF
						IF objExcel.cells(row_to_use, snap_actv_col).Value = "I" THEN
							EMSearch "approved", row, col 'Looking for a case note saying "approved" when review is coded I
							IF row <> 0 THEN 'found the word approved, next it finds out when
								EMReadScreen case_note_date, 10, row, 6
								IF datediff("d", case_note_date, review_date) < 32 THEN error_content = "Check for potential approved review." 'review can only be approved 1 month prior
							END IF
						END IF
						objExcel.Cells(row_to_use, error_column).Value = error_content
						error_content = "" 'Reset data
					END IF
				END IF
			END IF
		row_to_use = row_to_use + 1
		END IF
	Loop Until MAXIS_case_number = ""
END IF

'IF error_check = checked THEN col_to_use = col_to_use + 1 'need to add the extra column for this outside the DO...LOOP
col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns
row_to_use = 3			'Declaring here before the following if...then statements

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'SNAP info
If SNAP_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "SNAP cases with unapproved review:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of SNAP cases coded " & chr(34) & "N" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=COUNTIFS(" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & chr(34) & "N" & chr(34) & ")/(COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 2, col_to_use - 1).Value = "Percentage of SNAP cases coded " & chr(34) & "I" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 2, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 2, col_to_use).Value = "=COUNTIFS(" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & chr(34) & "I" & chr(34) & ")/(COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 2, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 3, col_to_use - 1).Value = "Percentage of SNAP cases coded " & chr(34) & "U" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 3, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 3, col_to_use).Value = "=COUNTIFS(" & SNAP_letter_col & ":" & SNAP_letter_col & ", " & chr(34) & "U" & chr(34) & ")/(COUNTA(" & SNAP_letter_col & ":" & SNAP_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 3, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 4	'It's four rows we jump, because the SNAP stat takes up four rows
End if

'HC info
If HC_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "HC cases with unapproved review:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of HC cases coded " & chr(34) & "N" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=COUNTIFS(" & HC_letter_col & ":" & HC_letter_col & ", " & chr(34) & "N" & chr(34) & ")/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 2, col_to_use - 1).Value = "Percentage of HC cases coded " & chr(34) & "I" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 2, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 2, col_to_use).Value = "=COUNTIFS(" & HC_letter_col & ":" & HC_letter_col & ", " & chr(34) & "I" & chr(34) & ")/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 2, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 3, col_to_use - 1).Value = "Percentage of HC cases coded " & chr(34) & "U" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 3, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 3, col_to_use).Value = "=COUNTIFS(" & HC_letter_col & ":" & HC_letter_col & ", " & chr(34) & "U" & chr(34) & ")/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 3, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 4, col_to_use - 1).Value = "Percentage of HC cases with exempt IRs:"		'Row header
	objExcel.Cells(row_to_use + 4, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 4, col_to_use).Value = "=COUNTIFS(" & exempt_IR_letter_col & ":" & exempt_IR_letter_col & ", " & chr(34) & "exempt" & chr(34) & ")/(COUNTA(" & HC_letter_col & ":" & HC_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 4, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 5	'It's four five we jump, because the HC stat takes up five rows
End if

'cash info
If cash_check = checked then
	ObjExcel.Cells(row_to_use, col_to_use - 1).Value = "Cash cases with unapproved review:"	'Row header
	objExcel.Cells(row_to_use, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold
	ObjExcel.Cells(row_to_use, col_to_use).Value = "=COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1"	'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use - 1).Value = "Percentage of cash cases coded " & chr(34) & "N" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 1, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 1, col_to_use).Value = "=COUNTIFS(" & cash_letter_col & ":" & cash_letter_col & ", " & chr(34) & "N" & chr(34) & ")/(COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 1, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 2, col_to_use - 1).Value = "Percentage of cash cases coded " & chr(34) & "I" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 2, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 2, col_to_use).Value = "=COUNTIFS(" & cash_letter_col & ":" & cash_letter_col & ", " & chr(34) & "I" & chr(34) & ")/(COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 2, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	ObjExcel.Cells(row_to_use + 3, col_to_use - 1).Value = "Percentage of cash cases coded " & chr(34) & "U" & chr(34) & ":"	'Row header
	objExcel.Cells(row_to_use + 3, col_to_use - 1).Font.Bold = TRUE								'Row header should be bold
	ObjExcel.Cells(row_to_use + 3, col_to_use).Value = "=COUNTIFS(" & cash_letter_col & ":" & cash_letter_col & ", " & chr(34) & "U" & chr(34) & ")/(COUNTA(" & cash_letter_col & ":" & cash_letter_col & ") - 1)" 'Excel formula
	ObjExcel.Cells(row_to_use + 3, col_to_use).NumberFormat = "0.00%"		'Formula should be percent
	row_to_use = row_to_use + 4	'It's four rows we jump, because the cash stat takes up four rows
End if

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! Your REPT/REVS list has been created.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------11/26/2022
'--Tab orders reviewed & confirmed----------------------------------------------11/26/2022
'--Mandatory fields all present & Reviewed--------------------------------------11/26/2022
'--All variables in dialog match mandatory fields-------------------------------11/26/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------11/26/2022------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------11/26/2022------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------11/26/2022------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-11/26/2022------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------11/26/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------11/26/2022------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------11/26/2022------------------N/A
'--Out-of-County handling reviewed----------------------------------------------11/26/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------11/26/2022
'--BULK - review output of statistics and run time/count (if applicable)--------11/26/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------11/26/2022
'--Incrementors reviewed (if necessary)-----------------------------------------11/26/2022
'--Denomination reviewed -------------------------------------------------------11/26/2022
'--Script name reviewed---------------------------------------------------------11/26/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------11/26/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------11/26/2022
'--comment Code-----------------------------------------------------------------11/26/2022
'--Update Changelog for release/update------------------------------------------11/26/2022
'--Remove testing message boxes-------------------------------------------------11/26/2022
'--Remove testing code/unnecessary code-----------------------------------------11/26/2022
'--Review/update SharePoint instructions----------------------------------------11/26/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------11/26/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------11/26/2022
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------11/26/2022 - Will pull in separtely; pending pull request has updates to CLoS.
'--Complete misc. documentation (if applicable)---------------------------------11/26/2022
'--Update project team/issue contact (if applicable)----------------------------11/26/2022
