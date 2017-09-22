'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - INTERVIEW REQUIRED.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	70			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

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

'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog appointment_required_dialog, 0, 0, 286, 60, "Appointment required dialog"
  EditBox 70, 5, 210, 15, worker_number
  CheckBox 5, 45, 155, 10, "Select all active workers in the agency", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 175, 40, 50, 15
    CancelButton 230, 40, 50, 15
  Text 5, 25, 275, 10, "Enter the fulll 7-digit worker number, separate each with a comma if more than one."
  Text 5, 10, 60, 10, "Worker number(s):"
EndDialog

Function HCRE_panel_bypass() 
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function
		
'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
all_workers_check = 1		'defaulting the check box to checked

'DISPLAYS DIALOG
DO
	DO
		err_msg = ""
		Dialog appointment_required_dialog
		If ButtonPressed = 0 then StopScript
		If worker_number = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Enter a valid worker number."
		if worker_number <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Enter a worker number OR select the entire agency, not both." 
		If datePart("d", date) < 16 then err_msg = err_msg & VbNewLine & "* This is not a valid time period for REPT/REVS until the 16th of the month. Please select a new time period."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in		

REPT_month = CM_plus_2_mo
REPT_year  = CM_plus_2_yr

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
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

'We need to get back to SELF and manually update the footer month
back_to_self
'clears all data from the SELF screen
EMWriteScreen "____", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen"____", 21, 70
transmit
transmit

Call navigate_to_MAXIS_screen("REPT", "REVS")
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit
EMWriteScreen REPT_month, 20, 55
EMWriteScreen REPT_year, 20, 58
transmit

'establishes values for variables and declaring the arrays
reviews_total = 0
total_cases_review = 0
DIM REVS_array()
REDim REVS_array(0)

'start of the FOR...next loop
For each worker in worker_array
	If trim(worker) = "" then exit for
	worker_number = trim(worker)
	'writing in the worker number in the correct col
	EMWriteScreen worker, 21, 6
	transmit

    'Grabbing case numbers from REVS for requested worker
	DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
		MAXIS_row = 7	'Setting or resetting this to look at the top of the list
		DO		'All of this loops until MAXIS_row = 19
			'Reading case information (case number, SNAP status, and cash status)
			EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6
			MAXIS_case_number = trim(MAXIS_case_number)
			EMReadScreen SNAP_status, 1, MAXIS_row, 45
			EMReadScreen cash_status, 1, MAXIS_row, 39
            
			'Navigates though until it runs out of case numbers to read
			IF MAXIS_case_number = "" then exit do

			'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
			If cash_status = "-" 	then cash_status = ""
			If SNAP_status = "-" 	then SNAP_status = ""
			If HC_status = "-" 		then HC_status = ""

			'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
			If ( ( trim(SNAP_status) = "N" or trim(SNAP_status) = "I" or trim(SNAP_status) = "U" ) or ( trim(cash_status) = "N" or trim(cash_status) = "I" or trim(cash_status) = "U" ) ) then
				ReDim Preserve REVS_array(reviews_total)				        'This resizes the array based on the number of members being added to the array
				REVS_array(reviews_total) = MAXIS_case_number
				reviews_total = reviews_total + 1
				total_cases_review = total_cases_review + 1
			End if
			'On the next loop it must look to the next row
			MAXIS_row = MAXIS_row + 1

			'Clearing variables before next loop
			add_to_array = ""
			MAXIS_case_number = ""
		Loop until MAXIS_row = 19		'Last row in REPT/REVS
		'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
		PF8
		EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
        'if max reviews are reached, the goes to next worker is applicable
	Loop until last_page_check = "THIS IS THE LAST PAGE"
next

DIM Required_appt_array()
ReDim Required_appt_array(8, 0)

'constants for array
const basket_number = 0
const case_number	= 1
const active_progs  = 2
const case_interp	= 3
const case_lang		= 4
const phone_one		= 5	
const phone_two		= 6	
const phone_three	= 7	

worker_number = ""
back_to_SELF

recert_cases = 0	'value for the array

'DO 'Loops until there are no more cases in the Excel list
For each reviews_total in REVS_array
	MAXIS_case_number = reviews_total 
	recert_status = "NO"	'Defaulting this to no because if SNAP or MFIP are not active - no recert will be scheduled
	CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
	EMReadScreen wrkr_numb, 7, 21, 21
	
	'Checking for PRIV cases.
	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" THEN 
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
	ELSE						'For all of the cases that aren't privileged...
		MFIP_ACTIVE = FALSE		'Setting some variables for the loop
		SNAP_ACTIVE = False		

		SNAP_status_check = ""
		MFIP_prog_1_check = ""
		MFIP_status_1_check = ""
		MFIP_prog_2_check = ""
		MFIP_status_2_check = ""

		'Reading the status and program
		EMReadScreen SNAP_status_check, 4, 10, 74		'checking the SNAP status
		EMReadScreen MFIP_prog_1_check, 2, 6, 67		'checking for an active MFIP case
		EMReadScreen MFIP_status_1_check, 4, 6, 74
		EMReadScreen MFIP_prog_2_check, 2, 6, 67		'checking for an active MFIP case
		EMReadScreen MFIP_status_2_check, 4, 6, 74

		IF SNAP_status_check = "ACTV" Then SNAP_ACTIVE = TRUE
		
		'Logic to determine if MFIP is active
		If MFIP_prog_1_check = "MF" Then
			If MFIP_status_1_check = "ACTV" Then MFIP_ACTIVE = TRUE
		ElseIf MFIP_prog_2_check = "MF" Then
			If MFIP_status_2_check = "ACTV" Then MFIP_ACTIVE = TRUE
		End If
		
		HCRE_panel_bypass	'function I created to ensure that we don't get trapped in the HCRE panel

		'Going to STAT/REVW to to check for ER vs CSR for SNAP cases
		CALL navigate_to_MAXIS_screen("STAT", "REVW")
		If MFIP_ACTIVE = TRUE Then recert_status = "YES"	'MFIP will only have an ER - so if listed on REVS - will be an ER - don't need to check dates
		If SNAP_ACTIVE = TRUE Then
			EMReadScreen SNAP_review_check, 8, 9, 57
			If SNAP_review_check = "__ 01 __" then 		'If this is blank there are big issues
				recert_status = "NO"
			Else
				EMwritescreen "x", 5, 58		'Opening the SNAP popup
				Transmit
				DO
				    EMReadScreen SNAP_popup_check, 7, 5, 43
				LOOP until SNAP_popup_check = "Reports"

				'The script will now read the CSR MO/YR and the Recert MO/YR
				EMReadScreen CSR_mo, 2, 9, 26
				EMReadScreen CSR_yr, 2, 9, 32
				EMReadScreen recert_mo, 2, 9, 64
				EMReadScreen recert_yr, 2, 9, 70

				'Comparing CSR and ER daates to the month of REVS review
				IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN recert_status = "NO"
				If recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) THEN recert_status = "NO"
				IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) THEN recert_status = "YES"
			End If
		End If 
		
		If recert_status = "YES" then 
			Redim Preserve Required_appt_array(8, 	recert_cases)
			Required_appt_array (case_number, 		recert_cases) = MAXIS_case_number
			Required_appt_array (x1number, 			recert_cases) = wrkr_numb
			IF MFIP_ACTIVE = TRUE AND SNAP_ACTIVE = FALSE Then Required_appt_array (active_progs, recert_cases) = "MFIP"
			If MFIP_ACTIVE = TRUE AND SNAP_ACTIVE = TRUE  Then Required_appt_array (active_progs, recert_cases) = "MFIP & SNAP"
			If MFIP_ACTIVE = FALSE AND SNAP_ACTIVE = TRUE Then Required_appt_array (active_progs, recert_cases) = "SNAP"
			
			'Gathering the phone numbers
			call navigate_to_MAXIS_screen("STAT", "ADDR")
			EMReadScreen phone_number_one, 16, 17, 43	' if phone numbers are blank it doesn't add them to EXCEL
			If phone_number_one <> "( ___ ) ___ ____" then Required_appt_array (phone_one, recert_cases) = phone_number_one
			EMReadScreen phone_number_two, 16, 18, 43
			If phone_number_two <> "( ___ ) ___ ____" then Required_appt_array (phone_two, recert_cases) = phone_number_two
			EMReadScreen phone_number_three, 16, 19, 43
			If phone_number_three <> "( ___ ) ___ ____" then Required_appt_array (phone_three, recert_cases) = phone_number_three	
			
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
			
			Required_appt_array (case_interp,  recert_cases) = interpreter_code
			Required_appt_array (case_lang,    recert_cases) = language_coded
			recert_cases = recert_cases + 1
			STATS_counter = STATS_counter + 1						'adds one instance to the stats counter
		End if 
	End if 	
Next

'----------------------------------------------------------------------------------------------------EXCEL INPUT
'Opening the Excel file, (now that the dialog is done)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'formatting excel file with columns for case number and interview date/time
objExcel.cells(1, 1).value 	= "X number"
objExcel.cells(1, 2).value 	= "Case number"
objExcel.cells(1, 3).value 	= "Programs"
objExcel.cells(1, 4).value 	= "Case language"
objExcel.Cells(1, 5).value 	= "Interpreter"
objExcel.cells(1, 6).value 	= "Phone # One"
objExcel.cells(1, 7).value 	= "Phone # Two"
objExcel.Cells(1, 8).value 	= "Phone # Three"
'objExcel.cells(1, 9).value 	= "Privileged Cases"
	
FOR i = 1 to 9									'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Adding the case information to Excel
excel_row = 2
For item = 0 to UBound(Required_appt_array, 2)
	ObjExcel.Cells(excel_row, 1).value  = Required_appt_array (x1number,     item)
	ObjExcel.Cells(excel_row, 2).value  = Required_appt_array (case_number,  item)
	ObjExcel.Cells(excel_row, 3).value  = Required_appt_array (active_progs, item)
	ObjExcel.Cells(excel_row, 4).value  = Required_appt_array (case_lang,    item)
	ObjExcel.Cells(excel_row, 5).value  = Required_appt_array (case_interp,  item)
	ObjExcel.Cells(excel_row, 6).value = Required_appt_array (phone_one,     item)
	ObjExcel.Cells(excel_row, 7).value = Required_appt_array (phone_two,     item)
	ObjExcel.Cells(excel_row, 8).value = Required_appt_array (phone_three,   item)
	excel_row = excel_row + 1 
Next

''Creating the list of privileged cases and adding to the spreadsheet
'If priv_case_list <> "" Then
'	priv_case_list = right(priv_case_list, (len(priv_case_list)-1))
'	prived_case_array = split(priv_case_list, "|")
'	
'	excel_row = 2
'
'	FOR EACH MAXIS_case_number in prived_case_array
'		objExcel.cells(excel_row, 9).value = MAXIS_case_number
'		excel_row = excel_row + 1
'	NEXT
'End If

'Query date/time/runtime info
objExcel.Cells(1, 10).Font.Bold = TRUE
objExcel.Cells(2, 10).Font.Bold = TRUE
objExcel.Cells(3, 10).Font.Bold = TRUE
objExcel.Cells(4, 10).Font.Bold = TRUE
ObjExcel.Cells(1, 10).Value = "Query date and time:"	
ObjExcel.Cells(2, 10).Value = "Query runtime (in seconds):"	
ObjExcel.Cells(3, 10).Value = "Total reviews:"
ObjExcel.Cells(4, 10).Value = "Interview required:"
ObjExcel.Cells(1, 11).Value = now
ObjExcel.Cells(2, 11).Value = timer - query_start_time
ObjExcel.Cells(3, 11).Value = total_cases_review
ObjExcel.Cells(4, 11).Value = recert_cases

'Formatting the columns to autofit after they are all finished being created.
FOR i = 1 to 11
	objExcel.Columns(i).autofit()
Next

STATS_counter = STATS_counter - 1
script_end_procedure("Success! The file is ready to clean up and submit.")