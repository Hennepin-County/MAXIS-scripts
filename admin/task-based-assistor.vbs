'STATS GATHERING--------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - TASK BASED ASSISTOR.vbs"
start_time = timer
STATS_counter = 1  'sets the stats counter at one
STATS_manualtime = 100 'manual run time in seconds
STATS_denomination = "C" 			   'M is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY================================================================
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
CALL changelog_update("01/15/2021", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK

 'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
''---------------------------------------------------------------------------------------------------- previous day's assignment
assignment_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(assignment_date)       'finds the most recent previous working day for the file names
assignment_date = assignment_date & "" 'have to make it a string for the script to realize that it is a date '

BeginDialog Dialog1, 0, 0, 236, 115, "TASK BASED REVIEW"
  ButtonGroup ButtonPressed
    PushButton 175, 70, 50, 15, "Browse...", select_a_file_button
    OkButton 120, 95, 50, 15
    CancelButton 175, 95, 50, 15
  EditBox 10, 70, 150, 15, file_selection_path
  EditBox 175, 5, 50, 15, assignment_date
  Text 10, 35, 210, 10, "This script should be used to assist with the task based review."
  Text 10, 5, 155, 20, "Please enter the date the HSR was given the assignment:"
  GroupBox 5, 25, 225, 65, "Using this script:"
  Text 10, 50, 205, 20, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
EndDialog

'dialog and dialog DO...Loop
Do
 	Do
  	    err_msg = ""
  	    dialog Dialog1
  	    cancel_without_confirmation
  	    If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
  	    If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "Please select a file to continue."
		IF isdate(assignment_date) = False then err_msg = err_msg & vbnewline & "Please enter an assignment date."
  	    If err_msg <> "" Then MsgBox err_msg
 	Loop until err_msg = ""
 	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'setting the footer month to make the updates in'
CALL convert_date_into_MAXIS_footer_month(assignment_date, MAXIS_footer_month, MAXIS_footer_year)
CALL MAXIS_footer_month_confirmation
CALL ONLY_create_MAXIS_friendly_date(assignment_date)


'Opening today's list
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
'objExcel.worksheets("Report 1").Activate   'Activates the initial BOBI report

'Establishing array
DIM task_based_array()  'Declaring the array this is what this list is
ReDim task_based_array(interview_const, 0)  'Resizing the array 'that ,list is goign to have 20 parameter but to start with there is only one paparmeter it gets complicated - grid'
'for each row the column is going to be the same information type
'Creating constants to value the array elements this is why we create constants
const assignment_date_const 	= 0 '= "Date Assigned"
const excel_row_const			= 1 '=
const maxis_case_number_const 	= 2 '= "Case Number" - pretend this means 2
const case_name_const 			= 3 '= "Case Name"
const assigned_to_const 		= 4 '= "Assigned to"
const worker_number_const		= 5 '= "Assigned Worker X127#"
const case_note_const 			= 6 '=
const DAIL_count_const   		= 7 '= "DAIL Count
const DAIL_type_const   		= 9 '= "DAIL Count"
const case_status_const 		= 10 '=
const case_note_date_const 		= 11 '
const case_note_key_word_const 	= 12
const interview_const 			= 13 'Interview Completed

'setting the columns - using constant so that we know what is going on'
const excel_col_assignment_date = 1 'A'
const excel_col_case_number 	= 3 'C'
const excel_col_case_name 		= 4 'D'
const excel_col_assigned_to 	= 6 'F'
const excel_col_worker_number 	= 7 'G'
const excel_col_case_note   	= 8 'H' ended up being a true false due to macros was orginally a count
const excel_col_key_word		= 9 'I'
const excel_col_DAIL_count		= 10 'J'
const excel_col_DAIL_type   	= 11 'K' RECOMMEND REMOVAL
const excel_col_case_status 	= 15 'O'
const excel_col_interview		= 19 'S'

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start based on when picking up the information
entry_record = 0 'incrementor for the array and count

Do 'purpose is to read each excel row and to add into each excel array '
 	'Reading information from the Excel
 	MAXIS_case_number = objExcel.cells(excel_row, excel_col_case_number).Value
 	MAXIS_case_number = trim(MAXIS_case_number)
 	IF MAXIS_case_number = "" then exit do

   'Adding client information to the array - this is for READING FROM the excel
 	ReDim Preserve task_based_array(interview_const, entry_record)	'This resizes the array based on the number of cases
	task_based_array(maxis_case_number_const,  entry_record) = MAXIS_case_number
	task_based_array(assigned_to_const,  entry_record) = trim(objExcel.cells(excel_row, excel_col_assigned_to).Value)
	task_based_array(worker_number_const,   entry_record) =  trim(objExcel.cells(excel_row, excel_col_worker_number).Value)
	task_based_array(excel_row_const, entry_record) = excel_row
	'making space in the array for these variables, but valuing them as "" for now
  	entry_record = entry_record + 1			'This increments to the next entry in the array
  	stats_counter = stats_counter + 1 'Increment for stats counter
 	excel_row = excel_row + 1
Loop

back_to_self 'resetting MAXIS back to self before getting started
Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

'Loading of cases is complete. Reviewing the cases in the array.
For item = 0 to UBound(task_based_array, 2)
 	MAXIS_case_number   = task_based_array(maxis_case_number_const, item)
	assigned_to  		= task_based_array(assigned_to_const,   item)
	worker_number   	= task_based_array(worker_number_const,   item)

	CALL navigate_to_MAXIS_screen("CASE", "NOTE")
 	MAXIS_row = 5 'Defining row for the search feature.

 	EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip - checking in PROG and INQUIRY
	EMReadScreen county_code, 4, 21, 14  'Out of county cases from STAT
	EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to  navigate'
 	IF priv_check = "PRIV" THEN  'PRIV cases
  		EMReadscreen priv_worker, 26, 24, 46
  		task_based_array(case_status_const, item) = trim(priv_worker)
 	ELSEIf county_code <> "X127" THEN
	  task_based_array(case_status_const, item) = "OUT OF COUNTY CASE"
 	ELSEIF instr(case_invalid_error, "IS INVALID") THEN  'CASE xxxxxxxx IS INVALID FOR PERIOD 12/99
		task_based_array(case_status_const, item) = trim(case_invalid_error)
	ELSE
		EMReadScreen MAXIS_case_name, 27, 21, 40
		task_based_array(case_name_const, item) = trim(MAXIS_case_name)
		task_based_array(case_note_const, item) = "NO" 'defaulting to no to ensure we increment '
 	    DO
 	    	EMReadscreen case_note_date, 8, MAXIS_row, 6
			'MsgBox assignment_date & " ~ " & case_note_date & "~"
 	    	If trim(case_note_date) = "" THEN
				task_based_array(case_status_const, item) = "NO CASE NOTE"
				exit do
 	    	Else
 	    		IF case_note_date = assignment_date THEN 'weekends and the day prior has the date assigned confirmed by the SSR '
					task_based_array(case_note_date_const, item) = case_note_date
 	    			EMReadScreen case_note_worker_number, 7, MAXIS_row, 16
					'MsgBox worker_number & "~" & case_note_worker_number
 	    			IF worker_number = case_note_worker_number THEN
						task_based_array(case_note_const, item) = "YES"
 	    				case_note_count = case_note_count + 1
 	    				EMReadScreen case_note_header, 55, MAXIS_row, 25
  	    				case_note_header = lcase(trim(case_note_header))
 	    				If instr(case_note_header, "interview completed") then task_based_array(interview_const, item) = "TRUE"
	    			END IF
	    		END IF
 	    	END IF
		MAXIS_row = MAXIS_row + 1
 		IF MAXIS_row = 19 THEN
 			PF8 'moving to next case note page if at the end of the page
 			MAXIS_row = 5
 		END IF
  		LOOP UNTIL cdate(case_note_date) < cdate(assignment_date)   'repeats until the case note date is less than the assignment date
 		task_based_array(case_note_count_const, item) = case_note_count
	'END IF
 	CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
 	DO
 		EMReadScreen dail_check, 4, 2, 48
 		If next_dail_check <> "DAIL" then CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
 	LOOP UNTIL dail_check = "DAIL"

 	DAIL_count = 0	'these are the actionable DAIL counts only
 	dail_row = 5			'Because the script brings each new case to the top of the page, dail_row starts at 6.
 	DO
 		EmReadscreen number_of_dails, 1, 3, 67	'Reads where there count of dAILS is listed
 		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped

 		EMReadScreen DAIL_case_number, 8, dail_row, 73
 		DAIL_case_number = trim(DAIL_case_number)
 		If DAIL_case_number <> MAXIS_case_number then exit do
 	    'Determining if there is a new case number...
 	    EMReadScreen new_case, 8, dail_row, 63
 	    new_case = trim(new_case)
 	    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
 	     	Call write_value_and_transmit("T", dail_row, 3)
 	     	dail_row = 6
 	    ELSEIF new_case = "CASE NBR" THEN
 	     '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
 	     	Call write_value_and_transmit("T", dail_row + 1, 3)
 	     	dail_row = 6
 	    End if
 	    EMReadScreen DAIL_type, 4, dail_row, 6
 	    EMReadScreen dail_msg, 61, dail_row, 20
 	    dail_msg = trim(dail_msg)
		DAIL_type = trim(DAIL_type)
 	    EMReadScreen dail_month, 8, dail_row, 11
 	    dail_month = trim(dail_month)
 	    Call non_actionable_dails(actionable_dail)   'Function to evaluate the DAIL messages
 	    IF actionable_dail = True then dail_count = dail_count + 1
 	    dail_row = dail_row + 1

		'IF DAIL_type = "COLA" THEN DAIL_type = DAIL_all & "," ' this is to discuss the follow up on the case '
		'IF DAIL_type = "CLMS" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "CSES" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "ELIG" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "IEVS" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "INFC" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "IV-E" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "MEC2" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "PEPR" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "REIN" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "STAT" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "TIKL" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "ENDI" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "HIRE" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "ISPI" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "SSN " THEN DAIL_type = DAIL_all & "," 'ask about the trim '
		'IF DAIL_type = "SVES" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "BNDX" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "PARI" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "SDXS" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "SDXI" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "BEER" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "UNVI" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "UBEN" THEN DAIL_type = DAIL_all & ","
		'IF DAIL_type = "WAGE" THEN DAIL_type = DAIL_all & ","
'
		'DAIL_all = trim(DAIL_all)
		'If right(DAIL_all, 1) = "," THEN DAIL_all = left(DAIL_all, len(DAIL_all) - 1)

 	Loop
		task_based_array(DAIL_count_const, item)  = DAIL_count
		'task_based_array(DAIL_type_const, item)  = DAIL_type
	END IF
	CALL back_to_self
	'worker_number = "" 'clearing for the function' need to reset the worker number each time we go into the "next"
	DAIL_count = 0
	'DAIL_type = ""
Next

objExcel.Columns(1).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY

For item = 0 to UBound(task_based_array, 2)
 	excel_row = task_based_array(excel_row_const, item)
 	objExcel.Cells(excel_row, excel_col_case_name).Value  = task_based_array(case_name_const,   item)
 	objExcel.Cells(excel_row, excel_col_case_note).Value = task_based_array(case_note_const,   item)
 	objExcel.Cells(excel_row, excel_col_DAIL_count).Value = task_based_array(DAIL_count_const,  item)
 	'objExcel.Cells(excel_row, excel_col_DAIL_type).Value = task_based_array(DAIL_type_const,  item)
 	objExcel.Cells(excel_row, excel_col_case_status).Value = task_based_array(case_status_const, item)
 	objExcel.Cells(excel_row, excel_col_interview).Value = task_based_array(interview_const,   item)
Next

FOR i = 1 to 20							'formatting the cells'
	objExcel.Columns(i).AutoFit()		'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1   'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

script_end_procedure_with_error_report("Success your list has been updated, please review to ensure accuracy.")
