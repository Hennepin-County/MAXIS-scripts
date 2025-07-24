'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - CSR Prep List.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 30                      'manual run time in seconds
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
call changelog_update("07/01/2025", "Initial version.", "Casey Love, Hennepin Count")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

all_workers_check = checked
' worker_number = "X127EX3, X127EX4, X127EX5, X127EX7, X127ET5, X127ET6, X127ET7, X127ET8, X127EM2, X127EH9"

'THE SCRIPT-------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 226, 130, "Pull REPT data into Excel dialog"
  EditBox 75, 20, 145, 15, worker_number
  CheckBox 10, 60, 150, 10, "Check here to run this query county-wide.", all_workers_check
  DropListBox 10, 90, 185, 45, "Select One..."+chr(9)+"DAILS"+chr(9)+"MFIP"+chr(9)+"GA"+chr(9)+"GRH"+chr(9)+"UHFS", report_selection
  ButtonGroup ButtonPressed
    OkButton 125, 110, 45, 15
    CancelButton 175, 110, 45, 15
  Text 30, 5, 145, 10, "Create Report of Cases to Resolve for CSR"
  Text 10, 25, 65, 10, "Worker(s) to check:"
  Text 10, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 10, 80, 50, 10, "Report to Run:"
EndDialog


Do
	Do
		Dialog dialog1
		cancel_without_confirmation
		If (all_workers_check = 0 AND worker_number = "") then MsgBox "Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""
	LOOP until all_workers_check = 1 or worker_number <> ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

CASH_check = unchecked
GRH_check = unchecked
SNAP_check = unchecked
If report_selection = "MFIP" or report_selection = "GA" Then CASH_check = checked
If report_selection = "MFIP" Then CASH_prog_select = "MF"
If report_selection = "GA" Then CASH_prog_select = "GA"
If report_selection = "GRH" Then GRH_check = checked
If report_selection = "UHFS" Then
	SNAP_check = checked
	CASH_check = checked
	CASH_prog_select = "MF"
End If

If all_workers_check = checked Then file_name = "FULL - GA ONLY - " & report_selection & " - " & CM_plus_2_mo & "-" & CM_plus_2_yr & " SR Preparation Report.xlsx"
If all_workers_check = unchecked Then file_name = "Partial - " & report_selection & " - " & CM_plus_2_mo & "-" & CM_plus_2_yr & " SR Preparation Report.xlsx"
file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Cash Six-Month Reporting Transition\" & file_name

'Checking for MAXIS
Call check_for_MAXIS(True)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

const worker_col 	= 1
const case_numb_col	= 2
const name_col		= 3
const next_er_col 	= 4
const next_sr_col 	= 5
const appl_col		= 6
const CM_1_app_col 	= 7
const CM_1_rept_stat= 8
const dail_col 		= 9
const next_revw_col	= 10
const revw_type_col	= 11
const proc_date_col = 12

const cash_col 		= 13
const grh_col 		= 13
const snap_col 		= 14
const ssi_col		= 15

If report_selection = "UHFS" Then last_col = 15
If report_selection <> "UHFS" Then last_col = 13
If report_selection = "DAILS" Then last_col = 15

'Find all cases with FS & MF and get REVW dates from STAT/REVW - these all need manual update and approval
'QUESTION -- DO WE NEED TO UPDATE REVW FOR ALL CASES???
'MAYBE -- Send all MF and GA cases through background to see which have STAT Edits to resolve.

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, worker_col 	).Value = "WORKER"
ObjExcel.Cells(1, case_numb_col	).Value = "CASE NUMBER"
ObjExcel.Cells(1, name_col		).Value = "NAME"
ObjExcel.Cells(1, next_er_col 	).Value = "NEXT ER DATE"
ObjExcel.Cells(1, next_sr_col 	).Value = "NEXT SR DATE"
ObjExcel.Cells(1, appl_col 		).Value = "APPL DATE"
ObjExcel.Cells(1, CM_1_app_col 	).Value = "CM+1 APPROVAL DATE"
ObjExcel.Cells(1, CM_1_rept_stat).Value = "CM+1 REPORTING"
ObjExcel.Cells(1, dail_col 		).Value = "DAIL"
ObjExcel.Cells(1, next_revw_col ).Value = "NEXT REVW"
ObjExcel.Cells(1, revw_type_col ).Value = "REVW TYPE"
ObjExcel.Cells(1, proc_date_col ).Value = "PROCESSING DEADLINE"
If report_selection = "DAILS" Then
	ObjExcel.Cells(1, cash_col).Value = "CASH"
	ObjExcel.Cells(1, snap_col).Value = "SNAP"
	ObjExcel.Cells(1, ssi_col).Value = "SSI Exists"
End If
If CASH_check = checked Then ObjExcel.Cells(1, cash_col).Value = "CASH - " & CASH_prog_select
If GRH_check = checked 	Then ObjExcel.Cells(1, grh_col ).Value = "GRH"
If SNAP_check = checked Then ObjExcel.Cells(1, snap_col).Value = "SNAP"
If report_selection = "UHFS" Then ObjExcel.Cells(1, ssi_col).Value = "SSI Exists"

FOR i = 1 to last_col		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
NEXT
ObjExcel.columns(next_er_col).NumberFormat = "@" 		'formatting as text
ObjExcel.columns(next_sr_col).NumberFormat = "@" 		'formatting as text

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

If report_selection = "DAILS" Then

	For each worker in worker_array
		worker = trim(ucase(worker))					'Formatting the worker so there are no errors
		MAXIS_case_number = ""
		back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
		Call navigate_to_MAXIS_screen("DAIL", "DAIL")
		Call write_value_and_transmit(worker, 21, 6)
		' Call write_value_and_transmit("X", 4, 12)		'these messages are not all INFO messages so this is removed
		' EMWriteScreen " ", 7, 39
		' Call write_value_and_transmit("X", 13, 39)

        'Reads where the count of DAILs is listed. Used to verify DAIL is not empty.
        EMReadScreen number_of_dails, 1, 3, 67
        Do
			all_done = False
	        'If this space is blank the rest of the DAIL reading is skipped
            If number_of_dails = " " Then
                exit do
            End if
            'Because the script brings each new case to the top of the page, dail_row starts at 6.
            dail_row = 6
			case_recorded = ""

			Do
				dail_type = ""
				dail_msg = ""
				dail_month = ""
				MAXIS_case_number = ""
				actionable_dail = ""
				renewal_6_month_check = ""
                ' MsgBox "dail_row - " & dail_row & vbCr & "4"

				'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
				EMReadScreen new_case, 8, dail_row, 63
				new_case = trim(new_case)
				' MsgBox "new_case - " & new_case & vbCr & "dail_row - " & dail_row & vbCr & "case_recorded - " & case_recorded
				If case_recorded <> "" Then
					Do While new_case <> "CASE NBR"
						dail_row = dail_row + 1
						EMReadScreen full_case, 17, dail_row, 63
						If full_case = case_recorded Then dail_row = dail_row + 1
						EMReadScreen new_case, 8, dail_row, 63
						new_case = trim(new_case)
						EMReadScreen next_dail_check, 7, dail_row, 3
						If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
							'Attempt to navigate to the next page
							PF8
							EMReadScreen last_page_check, 21, 24, 2
							'Check if the last page of the DAIL has been reached, also handles for situations where the last DAIL has been deleted and it displays a 'NO MESSAGES' warning
							If last_page_check = "THIS IS THE LAST PAGE" or Instr(last_page_check, "NO MESSAGES") then
								all_done = true
								exit do
							Else
								dail_row = 6
							End if
						End if
					Loop
					case_recorded = ""
				End If
				IF all_done = true THEN exit do

				IF new_case <> "CASE NBR" THEN
					'If there is NOT a new case number, the script will top the message
					Call write_value_and_transmit("T", dail_row, 3)
				ELSEIF new_case = "CASE NBR" THEN
					'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
					Call write_value_and_transmit("T", dail_row + 1, 3)
				End if

				dail_row = 6

				'Determines the DAIL Type
				EMReadScreen dail_type, 4, dail_row, 6
				dail_type = trim(dail_type)

				'Determines the DAIL date
				EMReadScreen dail_month, 8, dail_row, 11
				dail_month = trim(dail_month)

				'Reads the DAIL msg to determine if it is an out-of-scope message
				EMReadScreen dail_msg, 60, dail_row, 20
				dail_msg = trim(dail_msg)
				' MsgBox "dail_msg: " & vbCr & dail_msg

				' If InStr(dail_msg, "MASS CHANGE NOT AUTO-APPROVED") <> 0 Then
				' If InStr(dail_msg, "GA: NOT AUTO-APPROVED") <> 0 Then
				If InStr(dail_msg, "NOT AUTO-APPROVED") <> 0 Then
					EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
					MAXIS_case_number = trim(MAXIS_case_number)

					ObjExcel.Cells(excel_row, worker_col).Value 	= worker					'ERROR FOR GA DAIL LIST HAPPENDED HERE 303:6
					ObjExcel.Cells(excel_row, case_numb_col).Value 	= MAXIS_case_number
					ObjExcel.Cells(excel_row, dail_col).Value 		= dail_type & " " & dail_month & "  -  " & dail_msg

					excel_row = excel_row + 1
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					case_recorded = left("CASE NBR: " & MAXIS_case_number & "       ", 17)
				End If
                ' MsgBox "dail_row - " & dail_row & vbCr & "1"
				dail_row = dail_row + 1
                ' MsgBox "dail_row - " & dail_row & vbCr & "2"

                'Checking for the last DAIL message. If it just processed the final message, the DAIL will appear blank but there is actually an invisible '_' at 6, 3. Handling to check for this and then navigate to the next page if needed. If it is on the last page, then it will exit the do loop
                EMReadScreen next_dail_check, 7, dail_row, 3
                If trim(next_dail_check) = "" or trim(next_dail_check) = "_" then
                    'Attempt to navigate to the next page
                    PF8
                    EMReadScreen last_page_check, 21, 24, 2
                    'Check if the last page of the DAIL has been reached, also handles for situations where the last DAIL has been deleted and it displays a 'NO MESSAGES' warning
                    If last_page_check = "THIS IS THE LAST PAGE" or Instr(last_page_check, "NO MESSAGES") then
                        all_done = true
                        exit do
                    Else
                        dail_row = 6
                    End if
                End if
                ' MsgBox "dail_row - " & dail_row & vbCr & "3"

			Loop
            IF all_done = true THEN exit do
        LOOP

	Next

Else

	For each worker in worker_array
		worker = trim(ucase(worker))					'Formatting the worker so there are no errors
		back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
		Call navigate_to_MAXIS_screen("REPT", "ACTV")
		EMWriteScreen worker, 21, 13
		TRANSMIT
		EMReadScreen user_worker, 7, 21, 71		'
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
					EMReadScreen client_name, 	21, MAXIS_row, 21		'Reading client name
					EMReadScreen CASH_1_status, 1, MAXIS_row, 54		'Reading cash status
					EMReadScreen CASH_1_prog, 	2, MAXIS_row, 51		'Reading cash status
					EMReadScreen CASH_2_status, 1, MAXIS_row, 58		'Reading cash status
					EMReadScreen CASH_2_prog, 	2, MAXIS_row, 56		'Reading cash status
					EMReadScreen SNAP_status, 	1, MAXIS_row, 61		'Reading SNAP status
					EMReadScreen GRH_status, 	1, MAXIS_row, 70		'Reading GRH status

					' MsgBox 	"CASH_1_status - " & CASH_1_status & vbCr &_
					' 		"CASH_1_prog - " & CASH_1_prog & vbCr &_
					' 		"CASH_2_status - " & CASH_2_status & vbCr &_
					' 		"CASH_2_prog - " & CASH_2_prog & vbCr &_
					' 		"GRH_status - " & GRH_status


					'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
					MAXIS_case_number = trim(MAXIS_case_number)
					If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
					all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")

					If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end

					'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)
					add_case_info_to_Excel = False
					If CASH_1_status <> " " and CASH_1_status <> "I" and CASH_check = checked and CASH_1_prog = CASH_prog_select then add_case_info_to_Excel = True
					If CASH_2_status <> " " and CASH_2_status <> "I" and CASH_check = checked and CASH_2_prog = CASH_prog_select then add_case_info_to_Excel = True
					If SNAP_status <> " " 	and SNAP_status <> "I" 	 and SNAP_check = checked 	then add_case_info_to_Excel = True
					If GRH_status <> " " 	and GRH_status <> "I" 	 and GRH_check = checked 	then add_case_info_to_Excel = True
					If report_selection = "GA" and GRH_status <> " " and GRH_status <> "I" Then add_case_info_to_Excel = False
					If report_selection = "UHFS" Then
						If SNAP_status = " " 	or SNAP_status = "I" Then add_case_info_to_Excel = False
						MFIP_Active = False
						If CASH_1_status <> " " and CASH_1_status <> "I" AND CASH_1_prog = CASH_prog_select Then MFIP_Active = True
						If CASH_2_status <> " " and CASH_2_status <> "I" AND CASH_2_prog = CASH_prog_select Then MFIP_Active = True
						If MFIP_Active = False Then add_case_info_to_Excel = False
					End If
					' MsgBox "add_case_info_to_Excel - " & add_case_info_to_Excel

					If add_case_info_to_Excel = True then
						ObjExcel.Cells(excel_row, worker_col).Value 	= worker
						ObjExcel.Cells(excel_row, case_numb_col).Value 	= MAXIS_case_number
						ObjExcel.Cells(excel_row, name_col).Value 		= client_name

						If CASH_check = checked and CASH_1_status <> " " and CASH_1_status <> "I" and (CASH_1_prog = CASH_prog_select) Then ObjExcel.Cells(excel_row, cash_col).Value 	= CASH_1_status
						If CASH_check = checked and CASH_2_status <> " " and CASH_2_status <> "I" and (CASH_2_prog = CASH_prog_select) Then ObjExcel.Cells(excel_row, cash_col).Value 	= CASH_2_status
						If SNAP_check = checked Then ObjExcel.Cells(excel_row, snap_col).Value = SNAP_status
						If GRH_check = checked Then ObjExcel.Cells(excel_row, grh_col).Value = GRH_status

						excel_row = excel_row + 1
						STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					End if
					MAXIS_row = MAXIS_row + 1
					add_case_info_to_Excel = ""	'Blanking out variable
					MAXIS_case_number = ""			'Blanking out variable
				Loop until MAXIS_row = 19
				PF8
			Loop until last_page_check = "THIS IS THE LAST PAGE"
		END IF
	Next

End If

Call back_to_SELF

'These are the constants that we need to create tables in Excel
Const xlSrcRange = 1
Const xlYes = 1
last_col_letter = convert_digit_to_excel_column(last_col)
table1Range = "A1:" & last_col_letter & excel_row-1
ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table1Range, xlYes).Name = "Table1"
ObjExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium9"

objExcel.ActiveWorkbook.SaveAs file_path



MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr
excel_row = 2
Do
	SR_mo = ""
	SR_yr = ""
	ER_mo = ""
	ER_yr = ""

	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, case_numb_col).Value)

	If report_selection = "DAILS" Then
		Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
		If is_this_priv = False Then
			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
			EMReadScreen case_name, 25, 21, 40
			ObjExcel.Cells(excel_row, name_col).Value = trim(case_name)
			If grh_case = True Then ObjExcel.Cells(excel_row, cash_col).Value = "GRH"
			If ga_case 	= True Then ObjExcel.Cells(excel_row, cash_col).Value = "GA"
			If mfip_case = True Then ObjExcel.Cells(excel_row, cash_col).Value = "MFIP"
			ObjExcel.Cells(excel_row, snap_col).Value = snap_status
		End If
	End If

	Call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)

	If is_this_priv = False Then
		If report_selection = "UHFS" or (report_selection = "DAILS" and ObjExcel.Cells(excel_row, cash_col).Value = "") Then
			Call write_value_and_transmit("X", 5, 58)
			EMReadScreen SR_mo, 2, 9, 26
			EMReadScreen SR_yr, 2, 9, 32
			EMReadScreen ER_mo, 2, 9, 64
			EMReadScreen ER_yr, 2, 9, 70

			transmit
			EMReadScreen next_revw, 8, 9, 57
			next_revw = replace(next_revw, " ", "/")
		Else
			Call write_value_and_transmit("X", 5, 35)
			EMReadScreen SR_mo, 2, 9, 26
			EMReadScreen SR_yr, 2, 9, 32
			EMReadScreen ER_mo, 2, 9, 64
			EMReadScreen ER_yr, 2, 9, 70

			transmit
			EMReadScreen next_revw, 8, 9, 37
			next_revw = replace(next_revw, " ", "/")

		End If

		ObjExcel.Cells(excel_row, next_er_col).Value = ER_mo & "/" & ER_yr
		ObjExcel.Cells(excel_row, next_sr_col).Value = SR_mo & "/" & SR_yr
		ObjExcel.Cells(excel_row, revw_type_col).Value = "Unknown"

		If SR_mo <> "__" Then sr_date = SR_mo & "/1/" & SR_yr
		If ER_mo <> "__" Then er_date = ER_mo & "/1/" & ER_yr
		If IsDate(next_revw) Then
			deadline_mo = DatePart("m", DateAdd("m", -2, next_revw))
			deadline_yr = DatePart("yyyy", DateAdd("m", -2, next_revw))
			ObjExcel.Cells(excel_row, next_revw_col).Value = next_revw
			If IsDate(er_date) and IsDate(sr_date) Then
				If DateDiff("d", sr_date, next_revw) = 0 Then ObjExcel.Cells(excel_row, revw_type_col).Value = "SR"
				If DateDiff("d", er_date, next_revw) = 0 Then ObjExcel.Cells(excel_row, revw_type_col).Value = "ER"
			ElseIf IsDate(er_date) Then
				If DateDiff("d", er_date, next_revw) = 0 Then ObjExcel.Cells(excel_row, revw_type_col).Value = "ER"
			ElseIf IsDate(sr_date) Then
				If DateDiff("d", sr_date, next_revw) = 0 Then ObjExcel.Cells(excel_row, revw_type_col).Value = "SR"
			End If
			ObjExcel.Cells(excel_row, proc_date_col).Value = deadline_mo & "/14/" & deadline_yr
		End If



		Call navigate_to_MAXIS_screen("STAT", "PROG")
		If report_selection = "UHFS" or (report_selection = "DAILS" and ObjExcel.Cells(excel_row, cash_col).Value = "") Then
			EMReadScreen appl_date, 8, 10, 33
			ObjExcel.Cells(excel_row, appl_col).Value = replace(appl_date, " ", "/")
		ElseIf report_selection = "GRH" or (report_selection = "DAILS" and ObjExcel.Cells(excel_row, cash_col).Value = "GRH") Then
			EMReadScreen appl_date, 8, 9, 33
			ObjExcel.Cells(excel_row, appl_col).Value = replace(appl_date, " ", "/")
		Else
			EMReadScreen appl_date, 8, 6, 33
			ObjExcel.Cells(excel_row, appl_col).Value = replace(appl_date, " ", "/")
		End If

		approval_date = ""
		reporting_status = ""
		If report_selection = "MFIP" or (report_selection = "DAILS" and ObjExcel.Cells(excel_row, cash_col).Value = "MFIP") Then
			Call back_to_SELF
			' call navigate_to_MAXIS_screen("ELIG", "    ")		'for MFIP, we need to navigate to the correct month FIRST from the main ELIG menu beecause there is sometimes a sig change panel
			' EMWriteScreen CM_mo, 20, 55
			' EMWriteScreen CM_yr, 20, 58
			' transmit
			' call navigate_to_MAXIS_screen("ELIG", "MFIP")

			' EMReadScreen sig_change_check, 4, 3, 38				'looking to see if the significant change panel is on this case
			' If sig_change_check = "MFSC" Then
			' 	'this is important because the command line is in a different place on the sig change panel so this call is slightly different
			' 	Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
			' Else
			' 	Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
			' End If
			' 'When the correct version is selected, MAXIS navigates to MFPR, even if MFSC exists, so once the correct version is selected, we can assume we are MFPR

			' If approved_version_found = True Then
			' 	EMReadScreen approval_date, 8, 3, 14		'this is the actual approval date - not the process date'
			' 	approval_date = DateAdd("d", 0, approval_date)
			' End If
			' ObjExcel.Cells(excel_row, CM_app_col).Value = approval_date

			Call back_to_SELF
			call navigate_to_MAXIS_screen("ELIG", "    ")		'for MFIP, we need to navigate to the correct month FIRST from the main ELIG menu beecause there is sometimes a sig change panel
			EMWriteScreen CM_plus_1_mo, 20, 55
			EMWriteScreen CM_plus_1_yr, 20, 58
			transmit
			call navigate_to_MAXIS_screen("ELIG", "MFIP")

			EMReadScreen sig_change_check, 4, 3, 38				'looking to see if the significant change panel is on this case
			If sig_change_check = "MFSC" Then
				'this is important because the command line is in a different place on the sig change panel so this call is slightly different
				Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
			Else
				Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
			End If
			'When the correct version is selected, MAXIS navigates to MFPR, even if MFSC exists, so once the correct version is selected, we can assume we are MFPR

			If approved_version_found = True Then
				EMReadScreen approval_date, 8, 3, 14		'this is the actual approval date - not the process date'
				approval_date = DateAdd("d", 0, approval_date)
				Call write_value_and_transmit("MFSM", 20, 71)
				EMReadScreen reporting_status, 11, 8, 31
				ObjExcel.Cells(excel_row, CM_1_rept_stat).Value = trim(reporting_status)
			End If
			ObjExcel.Cells(excel_row, CM_1_app_col).Value = approval_date

		ElseIf report_selection = "UHFS" or (report_selection = "DAILS" and ObjExcel.Cells(excel_row, cash_col).Value = "") Then
			call navigate_to_MAXIS_screen("ELIG", "FS  ")
			EMWriteScreen CM_plus_1_mo, 19, 54
			EMWriteScreen CM_plus_1_yr, 19, 57
			Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
			If approved_version_found = True Then
				EMReadScreen approval_date, 8, 3, 14		'this is the actual approval date - not the process date'
				approval_date = DateAdd("d", 0, approval_date)
				Call write_value_and_transmit("FSB1", 19, 70)
				EMReadScreen SSI_amount, 8, 12, 33
				SSI_amount = trim(SSI_amount)
				If SSI_amount = "" Then ObjExcel.Cells(excel_row, ssi_col).Value = False
				If SSI_amount <> "" Then ObjExcel.Cells(excel_row, ssi_col).Value = True
				Call write_value_and_transmit("FSSM", 19, 70)
				EMReadScreen reporting_status, 11, 8, 31
				ObjExcel.Cells(excel_row, CM_1_rept_stat).Value = trim(reporting_status)
			End If
			ObjExcel.Cells(excel_row, CM_1_app_col).Value = approval_date

		ElseIf report_selection = "GA" or (report_selection = "DAILS" and ObjExcel.Cells(excel_row, cash_col).Value = "GA") Then
			call navigate_to_MAXIS_screen("ELIG", "GA  ")
			EMWriteScreen elig_footer_month, 20, 54
			EMWriteScreen elig_footer_year, 20, 57
			Call find_last_approved_ELIG_version(20, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
			If approved_version_found = True Then
				EMReadScreen approval_date, 8, 3, 15		'this is the actual approval date - not the process date'
				approval_date = DateAdd("d", 0, approval_date)
				Call write_value_and_transmit("GASM", 20, 70)
				EMReadScreen reporting_status, 11, 8, 32
				ObjExcel.Cells(excel_row, CM_1_rept_stat).Value = trim(reporting_status)
			End If
			ObjExcel.Cells(excel_row, CM_1_app_col).Value = approval_date

		ElseIf report_selection = "GRH" or (report_selection = "DAILS" and ObjExcel.Cells(excel_row, cash_col).Value = "GRH") Then
			call navigate_to_MAXIS_screen("ELIG", "GRH ")
			EMWriteScreen elig_footer_month, 20, 55
			EMWriteScreen elig_footer_year, 20, 58
			Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
			If approved_version_found = True Then
				EMReadScreen approval_date, 8, 3, 14		'this is the actual approval date - not the process date'
				approval_date = DateAdd("d", 0, approval_date)
				Call write_value_and_transmit("GRSM", 20, 71)
				EMReadScreen reporting_status, 11, 7, 69
				ObjExcel.Cells(excel_row, CM_1_rept_stat).Value = trim(reporting_status)
			End If
			ObjExcel.Cells(excel_row, CM_1_app_col).Value = approval_date
		End If

	Else
		ObjExcel.Cells(excel_row, next_er_col).Value = "PRIV"
	End If

	Call back_to_SELF

	excel_row = excel_row + 1
	next_MAXIS_case_number = trim(ObjExcel.Cells(excel_row, case_numb_col).Value)
Loop until next_MAXIS_case_number = ""

objWorkbook.Save()		'saving the excel

col_to_use = last_col + 3	'Doing two because the wrap-up is two columns

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time
ObjExcel.Cells(3, col_to_use - 1).Value = "Number of pages found"	'Goes back one, as this is on the next row
ObjExcel.Cells(3, col_to_use).Value = number_of_pages
ObjExcel.Cells(4, col_to_use - 1).Value = "Total Number of Cases"	'Goes back one, as this is on the next row
ObjExcel.Cells(4, col_to_use).Value = "=COUNTA(Table1[WORKER])"

ObjExcel.Cells(5, col_to_use - 1).Value = "Cases to Process on:"	'Goes back one, as this is on the next row

BoldColLetter = convert_digit_to_excel_column(col_to_use-1)
next_deadline = DateAdd("d", 0, CM_mo&"/14/"&CM_yr)
If DateDiff("d", date, next_deadline) < 0 Then next_deadline = DateAdd("m", 1, next_deadline)
ObjExcel.Cells(6, col_to_use - 1).Value = next_deadline
ObjExcel.Cells(6, col_to_use).Value = "=COUNTIF(Table1[PROCESSING DEADLINE],"&chr(34)&"<"&chr(34)&"&"& BoldColLetter & "6)"
ObjExcel.Cells(6, col_to_use+1).Value = "Before"
row = 7
Do
	ObjExcel.Cells(row, col_to_use - 1).Value = next_deadline
	ObjExcel.Cells(row, col_to_use).Value = "=COUNTIF(Table1[PROCESSING DEADLINE],"& BoldColLetter & row &")"
	next_deadline = DateAdd("m", 1, next_deadline)
	row = row + 1
Loop until DateDiff("m", date, next_deadline) > 12
ObjExcel.Cells(row, col_to_use - 1).Value = next_deadline
ObjExcel.Cells(row, col_to_use).Value = "=COUNTIF(Table1[PROCESSING DEADLINE],"&chr(34)&">"&chr(34)&"&"& BoldColLetter & row &")"
ObjExcel.Cells(row, col_to_use+1).Value = "After"

For BoldRow = 1 to row
	objExcel.Cells(BoldRow, col_to_use - 1).Font.Bold = TRUE
Next
'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next
objWorkbook.Save()		'saving the excel


'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

script_end_procedure("Success! Your REPT/ACTV list has been created.")
