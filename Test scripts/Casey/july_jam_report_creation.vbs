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
Const xlSrcRange = 1
Const xlYes = 1

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



' review_report_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\" & report_date & " Review Report.xlsx"


july_jam_file_path = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Ex-Parte\July Jam Work Tracking.xlsx"


' 'Opening the Excel file, (now that the dialog is done)
' Set objExcel = CreateObject("Excel.Application")
' objExcel.Visible = True
' Set objWorkbook = objExcel.Workbooks.Add()
' objExcel.DisplayAlerts = True

' 'Changes name of Excel sheet to "Case information"
' ObjExcel.ActiveSheet.Name = report_date & " Review Report"

' 'formatting excel file with columns for case number and interview date/time
' objExcel.cells(1,  1).value = "X number"
' objExcel.cells(1,  2).value = "Case number"
' objExcel.cells(1,  3).value = "MA Status"
' objExcel.cells(1,  4).value = "MSP Status"
' objExcel.cells(1,  5).value = "Next HC SR"
' objExcel.cells(1,  6).value = "Next HC ER"
' objExcel.Cells(1,  7).value = "Notes"

' FOR i = 1 to 7									'formatting the cells'
' 	objExcel.Cells(1, i).Font.Bold = True		'bold font'
' 	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
' 	objExcel.Columns(i).AutoFit()				'sizing the columns'
' NEXT

' objExcel.ActiveSheet.ListObjects.add xlSrcRange,objExcel.Range("A1:G2"),,XlYes 'creating table for all 25 columns and 2 rows. Will increment as more cases/data columns are added.

' excel_row = 2
' REPT_month = "07"
' REPT_year = "23"

' back_to_self    'We need to get back to SELF and manually update the footer month
' Call navigate_to_MAXIS_screen("REPT", "REVS")
' EMWriteScreen REPT_month, 20, 55
' EMWriteScreen REPT_year, 20, 58
' transmit

' 'start of the FOR...next loop
' For each worker in worker_array
' 	worker = trim(worker)
' 	If worker = "" then exit for
' 	Call write_value_and_transmit(worker, 21, 6)   'writing in the worker number in the correct col

' 	'Grabbing case numbers from REVS for requested worker
' 	DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
' 		row = 7	'Setting or resetting this to look at the top of the list
' 		DO		'All of this loops until row = 19
' 			'Reading case information (case number, SNAP status, and cash status)
' 			EMReadScreen MAXIS_case_number, 8, row, 6
' 			MAXIS_case_number = trim(MAXIS_case_number)
' 			EMReadScreen SNAP_status, 1, row, 45
' 			EMReadScreen cash_status, 1, row, 39
' 			EmReadscreen HC_status, 1, row, 49

' 			'Navigates though until it runs out of case numbers to read
' 			IF MAXIS_case_number = "" then exit do

' 			'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
' 			If cash_status = "-" 	then cash_status = ""
' 			If SNAP_status = "-" 	then SNAP_status = ""
' 			If HC_status = "-" 		then HC_status = ""

' 			'Using if...thens to decide if a case should be added (status isn't blank)
' 			If trim(HC_status) = "N" or trim(HC_status) = "I" or trim(HC_status) = "U"  or trim(HC_status) = "A" or trim(HC_status) = "O" or trim(HC_status) = "D" or trim(HC_status) = "T" then
' 				'Adding the case information to Excel
' 				ObjExcel.Cells(excel_row, 1).value  = worker
' 				ObjExcel.Cells(excel_row, 2).value  = trim(MAXIS_case_number)
' 				excel_row = excel_row + 1
' 			End if

' 			row = row + 1    'On the next loop it must look to the next row
' 			MAXIS_case_number = "" 'Clearing variables before next loop
' 		Loop until row = 19		'Last row in REPT/REVS
' 		'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
' 		PF8
' 		EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
' 		'if max reviews are reached, the goes to next worker is applicable
' 	Loop until last_page_check = "THIS IS THE LAST PAGE"
' next

' 'Saves and closes the most the main spreadsheet before continuing
' objExcel.ActiveWorkbook.SaveAs "C:\Users\calo001\OneDrive - Hennepin County\Projects\Ex-Parte\July Jam Work Tracking.xlsx"

call excel_open(july_jam_file_path, True, True, ObjExcel, objWorkbook)

objExcel.cells(1,  1).value = "X number"
objExcel.cells(1,  2).value = "Case number"
objExcel.cells(1,  3).value = "MA Status"
objExcel.cells(1,  4).value = "MSP Status"
objExcel.cells(1,  5).value = "Next HC SR"
objExcel.cells(1,  6).value = "Next HC ER"
objExcel.Cells(1,  7).value = "Notes"

const worker_const		= 0
const case_number_const	= 1
const MA_status_const	= 2
const MSP_status_const	= 3
const HC_SR_status_const= 4
const HC_ER_status_const= 5
const notes_const 		= 6

Dim review_array()
ReDim review_array(notes_const, 0)


'Establish the reviews array
recert_cases = 0	            'incrementor for the array

' objExcel.worksheets(report_date & " Review Report").Activate  'Activates the review worksheet
excel_row = 2   'Excel start row reading the case information for the array

Do
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value 'reading case number
	MAXIS_case_number = trim(MAXIS_case_number)
	If MAXIS_case_number = "" then exit do

	worker = ObjExcel.Cells(excel_row, 1).Value

	ReDim Preserve review_array(notes_const, recert_cases)	'This resizes the array based on if master notes were found or not
	review_array(worker_const,          recert_cases) = trim(worker)
	review_array(case_number_const,     recert_cases) = MAXIS_case_number
	review_array(MA_status_const,       recert_cases) = ""
	review_array(MSP_status_const,      recert_cases) = ""
	review_array(HC_SR_status_const,    recert_cases) = ""
	review_array(HC_ER_status_const,    recert_cases) = ""
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

	Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv) 'function to check PRIV status
	If is_this_priv = True then
		review_array(notes_const, item) = "PRIV Case."
	Else
		EmReadscreen worker_prefix, 4, 21, 14
		If worker_prefix <> "X127" then
			review_array(notes_const, item) = "Out-of-County: " & right(worker_prefix, 2)
		Else
			'function to determine programs and the program's status---Yay Casey!
			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

			If case_active = False then
				review_array(notes_const, item) = "Case Not Active."
			Else
				'valuing the array variables from the inforamtion gathered in from CASE/CURR
				review_array(MA_status_const,   item) = ma_case
				review_array(MSP_status_const,  item) = msp_case

				CALL navigate_to_MAXIS_screen("STAT", "REVW")

				Call write_value_and_transmit("X", 5, 71) 'HC Review Information
				EmReadscreen HC_review_popup, 20, 4, 32
				If HC_review_popup = "HEALTH CARE RENEWALS" then
				'The script will now read the CSR MO/YR and the Recert MO/YR
					EMReadScreen CSR_mo, 2, 7, 27   'IR dates
					EMReadScreen CSR_yr, 2, 7, 33
					If CSR_mo = "__" or CSR_yr = "__" then
						EMReadScreen CSR_mo, 2, 7, 71   'IR/AR dates
						EMReadScreen CSR_yr, 2, 7, 77
					End if
					EMReadScreen recert_mo, 2, 8, 27
					EMReadScreen recert_yr, 2, 8, 33

					HC_CSR_date = CSR_mo & "/" & CSR_yr
					If HC_CSR_date = "__/__" then HC_CSR_date = ""

					HC_ER_date = recert_mo & "/" & recert_yr
					If HC_ER_date = "__/__" then HC_ER_date = ""

					EMReadScreen Ex_Parte_indicator, 1, 9, 27 'Y/N
					EMReadScreen Ex_Parte_mo, 2, 9, 71
					EMReadScreen Ex_Parte_yr, 4, 9, 74


					'Next HC ER and SR dates
					review_array(HC_SR_status_const, item) = HC_CSR_date
					review_array(HC_ER_status_const, item) = HC_ER_date

					Transmit 'to exit out of the pop-up screen
				Else
					Transmit 'to exit out of the pop-up screen
					review_array(notes_const, item) = "Unable to Access HC Review Information."
				End if
			End If
		End If
	End If

	objExcel.cells(excel_row,  3).value = review_array(MA_status_const,   item)
	objExcel.cells(excel_row,  4).value = review_array(MSP_status_const,  item)
	objExcel.cells(excel_row,  5).value = review_array(HC_SR_status_const, item)
	objExcel.cells(excel_row,  6).value = review_array(HC_ER_status_const, item)
	objExcel.Cells(excel_row,  7).value = review_array(notes_const, item)

	excel_row = excel_row + 1
	total_cases_review = total_cases_review + 1
	STATS_counter = STATS_counter + 1						'adds one instance to the stats counter
	MAXIS_case_number = ""
Next

'Formatting the columns to autofit after they are all finished being created.
FOR i = 1 to 7
	objExcel.Columns(i).autofit()
Next

'Saves and closes the main reivew report
objWorkbook.Save()
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

Call script_end_procedure("Success! The review report is ready.")

