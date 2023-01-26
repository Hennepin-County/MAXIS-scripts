'Required for statistical purposes==========================================================================================
name_of_script = "BULK - REPT-ACTV LIST.vbs"
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
CALL changelog_update("01/27/2020", "Updated the handling for CCL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("05/07/2018", "Updated the characters to pull for the client's name.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/12/2018", "Entering a supervisor X-Number in the Workers to Check will pull all X-Numbers listed under that supervisor in MAXIS. Addiional bug fix where script was missing cases.", "Casey Love, Hennepin County")
Call changelog_update("12/10/2016", "Added IV-E, Child Care and FIATed case statuses to script. Also added closing message informing user that script has ended sucessfully.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
'get_county_code
'Connects to BlueZone
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 286, 130, "Pull REPT data into Excel dialog"
  EditBox 75, 20, 145, 15, worker_number
  EditBox 90, 110, 15, 15, MAXIS_footer_month
  EditBox 110, 110, 15, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 185, 110, 45, 15
    CancelButton 235, 110, 45, 15
  Text 25, 5, 240, 10, "Pull ACTIVE Helth Care Cases into Excel with Approved Budget Indicator "
  Text 10, 25, 65, 10, "Worker(s) to check:"
  Text 10, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  CheckBox 10, 60, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 10, 70, 150, 10, "Identity FIATed cases on the spreadsheet", FIAT_check
  Text 10, 85, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 10, 115, 80, 10, "Footer Month to Check:"
EndDialog

Do
	Do
		err_msg = ""
		Dialog dialog1
		cancel_without_confirmation
		If (all_workers_check = 0 AND worker_number = "") then err_msg = err_msg & vbCr & "Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""
		Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
	LOOP until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)




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

const worker_number_const 				= 0
const case_number_const 				= 1
const client_name_const 				= 2
const next_revw_date_const				= 3
const hc_status_code_const				= 4
const fiat_status_const					= 5
const number_approved_hc_progs_const	= 6
const MA_progs_with_budget				= 7
const last_const 						= 8

Dim ALL_ACTIVE_HC_CASES_ARRAY()
ReDim ALL_ACTIVE_HC_CASES_ARRAY(last_const, 0)

'Setting the variable for what's to come
all_case_numbers_array = "*"
hc_cases = 0

For each worker in worker_array
	worker = trim(ucase(worker))					'Formatting the worker so there are no errors
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "ACTV")
	EMWriteScreen worker, 21, 13
	TRANSMIT
	EMReadScreen user_worker, 7, 21, 71
	EMReadScreen p_worker, 7, 21, 13
	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

	'msgbox "worker " & worker

	IF worker_number = "X127CCL" or worker = "127CCL" THEN
		DO
			EmReadScreen worker_confirmation, 20, 3, 11 'looking for CENTURY PLAZA CLOSED
			EMWaitReady 0, 0
			'MsgBox "Are we waiting?"
		LOOP UNTIL worker_confirmation = "CENTURY PLAZA CLOSED"
	END IF

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then
		'Grabbing each case number on screen
		Do
		    'Set variable for next do...loop
			MAXIS_row = 7
			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			EMReadscreen number_of_pages, 4, 3, 76 'getting page number because to ensure it doesnt fail'
			number_of_pages = trim(number_of_pages)
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number
				EMReadScreen client_name, 21, MAXIS_row, 21			'Reading client name
				EMReadScreen next_revw_date, 8, MAXIS_row, 42		'Reading application date
				EMReadScreen HC_status, 1, MAXIS_row, 64			'Reading HC status
				EMReadScreen FIAT_status, 1, MAXIS_row, 77			'Reading the FIAT status of a case

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				MAXIS_case_number = trim(MAXIS_case_number)
				If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end
				If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") = 0 and HC_status <> " " and HC_status <> "I" then
					all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")

					ReDim Preserve ALL_ACTIVE_HC_CASES_ARRAY(last_const, hc_cases)

					ALL_ACTIVE_HC_CASES_ARRAY(worker_number_const, hc_cases) = worker
					ALL_ACTIVE_HC_CASES_ARRAY(case_number_const, hc_cases) = MAXIS_case_number
					ALL_ACTIVE_HC_CASES_ARRAY(client_name_const, hc_cases) = client_name
					next_revw_date = trim(next_revw_date)
					ALL_ACTIVE_HC_CASES_ARRAY(next_revw_date_const, hc_cases) = replace(next_revw_date, " ", "/")
					ALL_ACTIVE_HC_CASES_ARRAY(hc_status_code_const, hc_cases) = HC_status
					ALL_ACTIVE_HC_CASES_ARRAY(fiat_status_const, hc_cases) = FIAT_status
					hc_cases = hc_cases + 1

				End if
				MAXIS_row = MAXIS_row + 1
				MAXIS_case_number = ""			'Blanking out variable
				STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	END IF
next

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"
ObjExcel.Cells(1, 4).Value = "NEXT REVW DATE"
ObjExcel.Cells(1, 5).Value = "HC"
ObjExcel.Cells(1, 6).Value = "Approved HC Spans"
ObjExcel.Cells(1, 7).Value = "MA Budgets cover " & MAXIS_footer_month & "/" & MAXIS_footer_year
ObjExcel.Cells(1, 8).Value = "Needs Review and Approve"

FOR i = 1 to 8		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
NEXT
last_letter_col = "H"

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 9 'Starting with 5 because cols 1-4 are already used
If FIAT_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "FIAT"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	FIAT_actv_col = col_to_use
	col_to_use = col_to_use + 1
	FIAT_letter_col = convert_digit_to_excel_column(FIAT_actv_col)
	last_letter_col = FIAT_letter_col
End if

excel_row = 2

For each_hc_case = 0 to UBound(ALL_ACTIVE_HC_CASES_ARRAY, 2)
	Call Back_to_SELF
	MAXIS_case_number = ALL_ACTIVE_HC_CASES_ARRAY(case_number_const, each_hc_case)


	call navigate_to_MAXIS_screen("ELIG", "HC  ")
	EMWriteScreen MAXIS_footer_month, 19, 54
	EMWriteScreen MAXIS_footer_year, 19, 57
	transmit

	hc_row = 8
	approved_hc_programs = 0
	all_MA_budgets_approved = True
	approved_MA_exists = False
	Do
		EMReadScreen new_hc_elig_ref_numbs, 2, hc_row, 3
		EMReadScreen new_hc_elig_full_name, 17, hc_row, 7

		If new_hc_elig_ref_numbs = "  " Then
			new_hc_elig_ref_numbs = hc_elig_ref_numbs
			new_hc_elig_full_name = hc_elig_full_name
		End If
		hc_elig_ref_numbs = new_hc_elig_ref_numbs
		hc_elig_full_name = new_hc_elig_full_name

		hc_elig_full_name = trim(hc_elig_full_name)

		EMReadScreen clt_hc_prog, 4, hc_row, 28
		If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "    " Then

			EMReadScreen prog_status, 3, hc_row, 68
			If prog_status <> "APP" Then                        'Finding the approved version
				EMReadScreen total_versions, 2, hc_row, 64
				If total_versions = "01" Then
					hc_prog_elig_appd = False
				Else
					EMReadScreen current_version, 2, hc_row, 58
					' MsgBox "hc_row - " & hc_row & vbCr & "current_version - " & current_version
					If current_version = "01" Then
						hc_prog_elig_appd = False
					Else
						prev_version = right ("00" & abs(current_version) - 1, 2)
						EMWriteScreen prev_version, hc_row, 58
						transmit
						hc_prog_elig_appd = True
					End If

				End If
			Else
				hc_prog_elig_appd = True
			End If
		Else
			hc_prog_elig_appd = False
		End If

		If hc_prog_elig_appd = True Then
			approved_hc_programs = approved_hc_programs + 1

			EMReadScreen hc_prog_elig_major_program, 		4, hc_row, 28
			EMReadScreen hc_prog_elig_eligibility_result, 	8, hc_row, 41
			EMReadScreen hc_prog_elig_status, 				8, hc_row, 50
			EMReadScreen hc_prog_elig_app_indc, 				6, hc_row, 68
			EMReadScreen hc_prog_elig_magi_excempt, 			6, hc_row, 74


			hc_prog_elig_major_program = trim(hc_prog_elig_major_program)

			Call write_value_and_transmit("X", hc_row, 26)
			' MsgBox "MOVING - 1" & vbCr & hc_prog_elig_major_program(hc_prog_count) & vbCr & "MEMB " & hc_elig_ref_numbs(hc_prog_count)
			EMReadScreen hc_prog_elig_process_date, 8, 2, 73
			hc_prog_elig_process_date = DateAdd("d", 0, hc_prog_elig_process_date)

			' If DateDiff("'d", hc_prog_elig_process_date, date) = 0 Then
			If hc_prog_elig_major_program = "HC D" Then
				EMReadScreen hc_prog_elig_app_date, 8, 3, 73

				EMReadScreen hc_prog_elig_source_of_info, 		4, 9, 33
				EMReadScreen hc_prog_elig_responsible_county, 	2, 8, 78
				EMReadScreen hc_prog_elig_servicing_county, 	2, 9, 78

				EMReadScreen hc_prog_elig_test_application_withdrawn, 			6, 13, 22
				EMReadScreen hc_prog_elig_test_application_process_incomplete, 6, 14, 22
				EMReadScreen hc_prog_elig_test_no_new_prog_eligibility, 		6, 15, 22
				EMReadScreen hc_prog_elig_test_assistance_unit, 				6, 16, 22

				EMReadScreen hc_prog_elig_worker_msg_one, 78, 19, 3
			End If

			If hc_prog_elig_major_program = "MA" or hc_prog_elig_major_program = "EMA" Then
				transmit
				approved_MA_exists = True
				EMReadScreen hc_prog_elig_app_date, 8, 4, 73
				PF3
				hc_col = 17
				Do
					EMReadScreen budg_mo, 2, 6, hc_col + 2
					EMReadScreen budg_yr, 2, 6, hc_col + 5
					' MsgBox "BUDG MO/YR:" & vbCr & budg_mo & "/" & budg_yr & vbCr & "Col: " & hc_col
					If budg_mo = MAXIS_footer_month AND budg_yr = MAXIS_footer_year Then
						EMReadScreen hc_prog_elig_elig_type, 		2, 12, hc_col
						EMReadScreen hc_prog_elig_elig_standard, 	1, 12, hc_col + 5
						EMReadScreen hc_prog_elig_method, 			1, 13, hc_col + 4
						EMReadScreen hc_prog_elig_waiver, 			1, 14, hc_col + 4

						EMReadScreen hc_prog_elig_total_net_income, 9, 15, hc_col
						EMReadScreen hc_prog_elig_standard, 		9, 16, hc_col
						EMReadScreen hc_prog_elig_excess_income, 	9, 17, hc_col
						If trim(hc_prog_elig_total_net_income) = "" Then hc_prog_elig_total_net_income = "0.00"
						Exit Do

					End If
					hc_col = hc_col + 11

					If hc_col = 83 Then all_MA_budgets_approved = False
				Loop until hc_col = 83
			End If

			If hc_prog_elig_major_program = "QMB" or hc_prog_elig_major_program = "SLMB" or hc_prog_elig_major_program = "QI1" Then
				transmit
				EMReadScreen hc_prog_elig_app_date, 8, 4, 73
			End If
		End If

		clt_hc_prog = ""
		hc_prog_elig_appd = ""
		hc_prog_elig_major_program = ""
		hc_prog_elig_eligibility_result = ""
		hc_prog_elig_status = ""
		hc_prog_elig_app_indc = ""
		hc_prog_elig_elig_type = ""
		hc_prog_elig_elig_standard = ""
		hc_prog_elig_method = ""
		hc_prog_elig_total_net_income = ""
		hc_prog_elig_standard = ""

		Do
			EMReadScreen hhmm_check, 4, 3, 51
			If hhmm_check <> "HHMM" Then PF3
		Loop Until hhmm_check = "HHMM"

		hc_row = hc_row + 1
		EMReadScreen next_ref_numb, 2, hc_row, 3
		EMReadScreen next_maj_prog, 4, hc_row, 28
		' MsgBox "Row: " & hc_row & vbCr & "Next Ref Numb: " & next_ref_numb & vbCr & "Next Major Prog: " & next_maj_prog
	Loop until next_ref_numb = "  " and next_maj_prog = "    "
	If approved_MA_exists = False Then 	all_MA_budgets_approved = "No MA"
	ALL_ACTIVE_HC_CASES_ARRAY(number_approved_hc_progs_const, each_hc_case) = approved_hc_programs
	ALL_ACTIVE_HC_CASES_ARRAY(MA_progs_with_budget, each_hc_case) = all_MA_budgets_approved

	add_to_Excel = True
	If ALL_ACTIVE_HC_CASES_ARRAY(hc_status_code_const, each_hc_case) = "P" AND ALL_ACTIVE_HC_CASES_ARRAY(number_approved_hc_progs_const, each_hc_case) = 0 then add_to_Excel = False
	If add_to_Excel = True Then
		ObjExcel.Cells(excel_row, 1).Value = ALL_ACTIVE_HC_CASES_ARRAY(worker_number_const, each_hc_case)
		ObjExcel.Cells(excel_row, 2).Value = ALL_ACTIVE_HC_CASES_ARRAY(case_number_const, each_hc_case)
		ObjExcel.Cells(excel_row, 3).Value = ALL_ACTIVE_HC_CASES_ARRAY(client_name_const, each_hc_case)
		ObjExcel.Cells(excel_row, 4).Value = ALL_ACTIVE_HC_CASES_ARRAY(next_revw_date_const, each_hc_case)
		ObjExcel.Cells(excel_row, 5).Value = ALL_ACTIVE_HC_CASES_ARRAY(hc_status_code_const, each_hc_case)
		ObjExcel.Cells(excel_row, 6).Value = ALL_ACTIVE_HC_CASES_ARRAY(number_approved_hc_progs_const, each_hc_case)
		ObjExcel.Cells(excel_row, 7).Value = ALL_ACTIVE_HC_CASES_ARRAY(MA_progs_with_budget, each_hc_case)
		review_and_approve = False
		If ALL_ACTIVE_HC_CASES_ARRAY(number_approved_hc_progs_const, each_hc_case) = 0 or ALL_ACTIVE_HC_CASES_ARRAY(MA_progs_with_budget, each_hc_case) = False then review_and_approve = True
		ObjExcel.Cells(excel_row, 8).Value = review_and_approve
		If FIAT_check = checked then ObjExcel.Cells(excel_row, FIAT_actv_col).Value = ALL_ACTIVE_HC_CASES_ARRAY(fiat_status_const, each_hc_case)
		excel_row = excel_row + 1
	End If
Next
excel_row = excel_row - 1
Call Back_to_SELF

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

Const xlSrcRange = 1
Const xlYes = 1
table1Range = "A1:" & last_letter_col & excel_row
ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table1Range, xlYes).Name = "HCApprovalInfo"

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(1, col_to_use + 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(1, col_to_use + 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use + 2).Value = timer - query_start_time
' ObjExcel.Cells(2, col_to_use - 1).Value = "Number of pages found"	'Goes back one, as this is on the next row
' ObjExcel.Cells(2, col_to_use).Value = number_of_pages
ObjExcel.Cells(2, col_to_use - 1).Value = "Number of cases that need Review"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = "=COUNTIFS(HCApprovalInfo[Needs Review and Approve], "&chr(34)&"TRUE"&chr(34)&")"

'Autofitting columns
For col_to_autofit = col_to_use - 1 to col_to_use + 2
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

script_end_procedure("Success! Your REPT/ACTV list has been created.")