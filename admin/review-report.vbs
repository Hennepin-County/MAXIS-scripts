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
		MsgBox "SOCWKR found!" & vbNewLine & "ROW - " & row & vbNewLine & "COL - " & col & vbNewLine & "~" & this_is_it & "~"
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
			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)

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
			call navigate_to_MAXIS_screen("STAT", "ADDR")
			EMReadScreen phone_number_one, 16, 17, 43	' if phone numbers are blank it doesn't add them to EXCEL
			If phone_number_one <> "( ___ ) ___ ____" then review_array(phone_1_const, incrementor_var) = phone_number_one
			EMReadScreen phone_number_two, 16, 18, 43
			If phone_number_two <> "( ___ ) ___ ____" then review_array(phone_2_const, incrementor_var) = phone_number_two
			EMReadScreen phone_number_three, 16, 19, 43
			If phone_number_three <> "( ___ ) ___ ____" then review_array(phone_3_const, incrementor_var) = phone_number_three

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

'DISPLAYS DIALOG
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 85, "Review Report"
  ' DropListBox 90, 35, 90, 15, "Select one..."+chr(9)+"Create Renewal Report"+chr(9)+"Discrepancy Run", renewal_option
  DropListBox 90, 35, 90, 15, "Select one..."+chr(9)+"Create Renewal Report"+chr(9)+"Discrepancy Run"+chr(9)+"Collect Statistics"+chr(9)+"Send Appointment Letters"+chr(9)+"Create Worklist", renewal_option
  ButtonGroup ButtonPressed
    OkButton 95, 65, 40, 15
    CancelButton 140, 65, 40, 15
  EditBox 70, 5, 110, 15, worker_number
  CheckBox 5, 55, 70, 10, "Select all agency.", all_workers_check
  CheckBox 5, 70, 70, 10, "Select for CM + 2.", CM_plus_two_checkbox
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

If renewal_option = "Collect Statistics" OR renewal_option = "Create Worklist" Then

	'If we are collecting statistics, we may be running on a current or past month, we need to clarify which month we are looking at.'
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 115, 55, "Select REVW Month for Statistics"
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
End If

report_date = REPT_month & "-" & REPT_year  'establishing review date

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
	MAXIS_footer_month = REPT_month							'Setting the footer month and year based on the review month. We do not run statistics in CM + 2
	MAXIS_footer_year = REPT_year

	'This is where the review report is currently saved.
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\" & report_date & " Review Report.xlsx"

	'Initial Dialog which requests a file path for the excel file
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 361, 70, "On Demand Recertifications"
	  EditBox 130, 20, 175, 15, excel_file_path
	  ButtonGroup ButtonPressed
	    PushButton 310, 20, 45, 15, "Browse...", select_a_file_button
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

	'Stats option ignores the 'list of workers' since it works off of an existing Excel, it needs to pull all of the workers
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)

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
		'Here we look through the entire array until we find a match
		For revs_item = 0 to UBound(review_array, 2)
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
				Exit For						'if we found a match, we should stop looking
			End If
		Next
		'if the case was not found in the array, we need to look in STAT for the information
		If found_in_array = FALSE AND case_number_to_check <> "" Then
			Call check_for_MAXIS(FALSE)		'making sure we haven't passworded out

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

				If cash_review_status <> "" Then ObjExcel.Cells(excel_row, cash_stat_excel_col).Value = cash_review_status		'Enter all the information into Excel
				If snap_review_status <> "" Then ObjExcel.Cells(excel_row, snap_stat_excel_col).Value = snap_review_status
				If hc_review_status <> "" Then ObjExcel.Cells(excel_row, hc_stat_excel_col).Value = hc_review_status
				If recvd_date <> "" Then ObjExcel.Cells(excel_row, recvd_date_excel_col).Value = recvd_date
				If interview_date <> "" Then ObjExcel.Cells(excel_row, intvw_date_excel_col).Value = interview_date
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

			excel_row = excel_row + 1		'going to the next excel
		End If
	Next

	'Going to another sheet, to enter worker-specific statistics and naming it
	sheet_name = "Statistics from " & date_month & "-" & date_day
	ObjExcel.Worksheets.Add().Name = sheet_name

	'Now we add all the information into Excel to calculate stats information and format
	ObjExcel.Cells(1, 2).Value = "ER with Interview"
	ObjExcel.Cells(1, 4).Value = "ER - No Interview"
	ObjExcel.Cells(1, 5).Value = "CSR"
	ObjExcel.Cells(1, 6).Value = "PRIV"
	ObjExcel.Cells(1, 8).Value = "Total"
	ObjExcel.Range("B1:C1").Merge
	ObjExcel.Range("F1:G1").Merge
	ObjExcel.Range("H1:I1").Merge
	ObjExcel.Cells(1, 2).HorizontalAlignment = -4108		'Center'
	ObjExcel.Cells(1, 4).HorizontalAlignment = -4108		'Center'
	ObjExcel.Cells(1, 5).HorizontalAlignment = -4108		'Center'
	ObjExcel.Cells(1, 6).HorizontalAlignment = -4108		'Center'
	ObjExcel.Cells(1, 8).HorizontalAlignment = -4108		'Center'


	ObjExcel.Cells(2, 1).Value = "All"
	ObjExcel.Cells(3, 1).Value = "Apps Received"
	ObjExcel.Cells(4, 1).Value = "No App in MX"
	ObjExcel.Cells(5, 1).Value = "Percent Received"
	ObjExcel.Cells(6, 1).Value = "Interview Completed"
	ObjExcel.Cells(7, 1).Value = "No Interview"
	ObjExcel.Cells(8, 1).Value = "Percent of Interviews Done"
	For i = 2 to 8
		ObjExcel.Cells(i, 1).Font.Bold = TRUE
	Next

	ObjExcel.Cells(10, 2).Value = "Interview Count"
	ObjExcel.Cells(10, 3).Value = "App Recvd Count"
	ObjExcel.Cells(10, 4).Value = "App Recvd Count"
	ObjExcel.Cells(10, 5).Value = "App Recvd Count"
	ObjExcel.Cells(10, 6).Value = "Interview Count"
	ObjExcel.Cells(10, 7).Value = "App Recvd Count"
	ObjExcel.Cells(10, 8).Value = "Interview Count"
	ObjExcel.Cells(10, 9).Value = "App Recvd Count"
	For i = 2 to 9
		ObjExcel.Cells(10, i).Font.Bold = TRUE
	Next

	ObjExcel.Cells(1, 13).Value = "Apps Received"
	ObjExcel.Range("M1:N1").Merge
	ObjExcel.Cells(2, 13).Value = "Count"
	ObjExcel.Cells(2, 14).Value = "%"
	ObjExcel.Cells(1, 15).Value = "Interviews Completed"
	ObjExcel.Range("O1:P1").Merge
	ObjExcel.Cells(2, 15).Value = "Count"
	ObjExcel.Cells(2, 16).Value = "%"
	ObjExcel.Cells(1, 17).Value = "REVW - I"
	ObjExcel.Range("Q1:R1").Merge
	ObjExcel.Cells(2, 17).Value = "Count"
	ObjExcel.Cells(2, 18).Value = "%"
	ObjExcel.Cells(1, 19).Value = "REVW - U"
	ObjExcel.Range("S1:T1").Merge
	ObjExcel.Cells(2, 19).Value = "Count"
	ObjExcel.Cells(2, 20).Value = "%"
	ObjExcel.Cells(1, 21).Value = "REVW - N"
	ObjExcel.Range("U1:V1").Merge
	ObjExcel.Cells(2, 21).Value = "Count"
	ObjExcel.Cells(2, 22).Value = "%"
	ObjExcel.Cells(1, 23).Value = "REVW - A"
	ObjExcel.Range("W1:X1").Merge
	ObjExcel.Cells(2, 23).Value = "Count"
	ObjExcel.Cells(2, 24).Value = "%"
	ObjExcel.Cells(1, 25).Value = "REVW - O"
	ObjExcel.Range("Y1:Z1").Merge
	ObjExcel.Cells(2, 25).Value = "Count"
	ObjExcel.Cells(2, 26).Value = "%"
	ObjExcel.Cells(1, 27).Value = "REVW - T"
	ObjExcel.Range("AA1:AB1").Merge
	ObjExcel.Cells(2, 27).Value = "Count"
	ObjExcel.Cells(2, 28).Value = "%"
	ObjExcel.Cells(1, 29).Value = "REVW - D"
	ObjExcel.Range("AC1:AD1").Merge
	ObjExcel.Cells(2, 29).Value = "Count"
	ObjExcel.Cells(2, 30).Value = "%"
	ObjExcel.Cells(1, 31).Value = "Totals"
	ObjExcel.Range("A1").EntireRow.Font.Size = "14"
	for i = 13 to 31
		ObjExcel.Cells(2, i).Font.Bold = True
	next

	ObjExcel.Cells(3, 11).Value = "ER with Interview"
	ObjExcel.Cells(3, 12).Value = "All"
	ObjExcel.Cells(4, 12).Value = "Cash"
	ObjExcel.Cells(5, 12).Value = "SNAP"
	ObjExcel.Cells(6, 11).Value = "ER - No Interview"
	ObjExcel.Cells(6, 12).Value = "All"
	ObjExcel.Cells(7, 12).Value = "Cash"
	ObjExcel.Cells(8, 12).Value = "SNAP"
	ObjExcel.Cells(9, 11).Value = "CSR"
	ObjExcel.Cells(9, 12).Value = "All"
	ObjExcel.Cells(10, 12).Value = "GRH"
	ObjExcel.Cells(11, 12).Value = "SNAP"
	ObjExcel.Cells(12, 11).Value = "PRIV"
	ObjExcel.Cells(12, 12).Value = "All"
	ObjExcel.Cells(13, 12).Value = "Cash"
	ObjExcel.Cells(14, 12).Value = "SNAP"
	ObjExcel.Cells(15, 11).Value = "Total"
	ObjExcel.Cells(15, 12).Value = "All"
	ObjExcel.Cells(16, 12).Value = "Cash"
	ObjExcel.Cells(17, 12).Value = "SNAP"
	for i = 3 to 17
		ObjExcel.Cells(i, 11).Font.Bold = True
		ObjExcel.Cells(i, 12).Font.Bold = True
	next

	first_of_rept_month = REPT_month & "/1/" & REPT_year
	first_of_rept_month = DateAdd("d", 0, first_of_rept_month)
	search_start = DateAdd("m", -2, first_of_rept_month)
	search_start = DateAdd("d", 15, search_start)

	date_row = 11
	the_date = search_start
	Do
		ObjExcel.Cells(date_row, 1).Value = the_date
		the_date = DateAdd("d", 1, the_date)
		date_row = date_row + 1
	Loop until DateDiff("d", the_date, date) < 0
	' chr(34) - QUOTATION MARKS
	is_not_blank = chr(34) & "<>" & chr(34)
	is_blank = chr(34) & chr(34)
	is_true = chr(34)&"TRUE"&chr(34)
	is_false = chr(34)&"FALSE"&chr(34)

	ObjExcel.Cells(2, 2).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&")"
	ObjExcel.Cells(3, 2).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(4, 2).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(5, 2).Value = "=B3/B2"
	ObjExcel.Cells(5, 2).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 2).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[Intvw Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(7, 2).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[Intvw Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(8, 2).Value = "=B6/B2"
	ObjExcel.Cells(8, 2).NumberFormat = "0.00%"
	ObjExcel.Range("B2:C2").Merge
	ObjExcel.Range("B3:C3").Merge
	ObjExcel.Range("B4:C4").Merge
	ObjExcel.Range("B5:C5").Merge
	ObjExcel.Range("B6:C6").Merge
	ObjExcel.Range("B7:C7").Merge
	ObjExcel.Range("B8:C8").Merge

	ObjExcel.Cells(2, 4).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&")"
	ObjExcel.Cells(3, 4).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(4, 4).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(5, 4).Value = "=D3/D2"
	ObjExcel.Cells(5, 4).NumberFormat = "0.00%"

	ObjExcel.Cells(2, 5).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&")"
	ObjExcel.Cells(3, 5).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(4, 5).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CAF Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(5, 5).Value = "=E3/E2"
	ObjExcel.Cells(5, 5).NumberFormat = "0.00%"

	ObjExcel.Cells(2, 6).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&")"
	ObjExcel.Cells(3, 6).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CAF Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(4, 6).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CAF Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(5, 6).Value = "=F3/F2"
	ObjExcel.Cells(5, 6).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 6).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[Intvw Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(7, 6).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[Intvw Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(8, 6).Value = "=F3/F2"
	ObjExcel.Cells(8, 6).NumberFormat = "0.00%"
	ObjExcel.Range("F2:G2").Merge
	ObjExcel.Range("F3:G3").Merge
	ObjExcel.Range("F4:G4").Merge
	ObjExcel.Range("F5:G5").Merge
	ObjExcel.Range("F6:G6").Merge
	ObjExcel.Range("F7:G7").Merge
	ObjExcel.Range("F8:G8").Merge

	ObjExcel.Cells(2, 8).Value = "=COUNTA(Table1[Case number])"
	ObjExcel.Cells(3, 8).Value = "=COUNTIFS(Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(4, 8).Value = "=COUNTIFS(Table1[CAF Date ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(5, 8).Value = "=H3/H2"
	ObjExcel.Cells(5, 8).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 8).Value = "=COUNTIFS(Table1[Intvw Date ("&date_header&")],"&is_not_blank&")"
	ObjExcel.Cells(7, 8).Value = "=COUNTIFS(Table1[Intvw Date ("&date_header&")],"&is_blank&")"
	ObjExcel.Cells(8, 8).Value = "=H3/H2"
	ObjExcel.Cells(8, 8).NumberFormat = "0.00%"
	ObjExcel.Range("H2:I2").Merge
	ObjExcel.Range("H3:I3").Merge
	ObjExcel.Range("H4:I4").Merge
	ObjExcel.Range("H5:I5").Merge
	ObjExcel.Range("H6:I6").Merge
	ObjExcel.Range("H7:I7").Merge
	ObjExcel.Range("H8:I8").Merge

	stats_row = 11
	Do
		ObjExcel.Cells(stats_row, 2).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[Intvw Date ("&date_header&")], A"&stats_row&")"
		ObjExcel.Cells(stats_row, 3).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CAF Date ("&date_header&")], A"&stats_row&")"
		ObjExcel.Cells(stats_row, 4).Value = "=COUNTIFS(Table1[Interview ER], "&is_false&",Table1[No Interview ER],"&is_true&",Table1[CAF Date ("&date_header&")], A"&stats_row&")"
		ObjExcel.Cells(stats_row, 5).Value = "=COUNTIFS(Table1[Interview ER], "&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CAF Date ("&date_header&")], A"&stats_row&")"
		ObjExcel.Cells(stats_row, 6).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[Intvw Date ("&date_header&")], A"&stats_row&")"
		ObjExcel.Cells(stats_row, 7).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CAF Date ("&date_header&")], A"&stats_row&")"
		ObjExcel.Cells(stats_row, 8).Value = "=COUNTIFS(Table1[Intvw Date ("&date_header&")], A"&stats_row&")"
		ObjExcel.Cells(stats_row, 9).Value = "=COUNTIFS(Table1[CAF Date ("&date_header&")], A"&stats_row&")"
		stats_row = stats_row + 1
		next_row_date = ObjExcel.Cells(stats_row, 1).Value
	Loop until next_row_date = ""
	last_row = stats_row - 1


	ObjExcel.Cells(3, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&")"
	ObjExcel.Cells(4, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(5, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(3, 13).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(4, 13).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(5, 13).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(3, 14).Value = "=M3/AE3"
	ObjExcel.Cells(4, 14).Value = "=M4/AE4"
	ObjExcel.Cells(5, 14).Value = "=M5/AE5"
	ObjExcel.Cells(3, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 14).NumberFormat = "0.00%"
		ObjExcel.Cells(3, 15).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(4, 15).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(5, 15).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(3, 16).Value = "=O3/AE3"
	ObjExcel.Cells(4, 16).Value = "=O4/AE4"
	ObjExcel.Cells(5, 16).Value = "=O5/AE5"
	ObjExcel.Cells(3, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(3, 17).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(4, 17).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(5, 17).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(3, 18).Value = "=Q3/AE3"
	ObjExcel.Cells(4, 18).Value = "=Q4/AE4"
	ObjExcel.Cells(5, 18).Value = "=Q5/AE5"
	ObjExcel.Cells(3, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(3, 19).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(4, 19).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(5, 19).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(3, 20).Value = "=Q3/AE3"
	ObjExcel.Cells(4, 20).Value = "=Q4/AE4"
	ObjExcel.Cells(5, 20).Value = "=Q5/AE5"
	ObjExcel.Cells(3, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(3, 21).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(4, 21).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(5, 21).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(3, 22).Value = "=U3/AE3"
	ObjExcel.Cells(4, 22).Value = "=U4/AE4"
	ObjExcel.Cells(5, 22).Value = "=U5/AE5"
	ObjExcel.Cells(3, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(3, 23).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(4, 23).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(5, 23).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(3, 24).Value = "=W3/AE3"
	ObjExcel.Cells(4, 24).Value = "=W4/AE4"
	ObjExcel.Cells(5, 24).Value = "=W5/AE5"
	ObjExcel.Cells(3, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(3, 25).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(4, 25).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(5, 25).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(3, 26).Value = "=Y3/AE3"
	ObjExcel.Cells(4, 26).Value = "=Y4/AE4"
	ObjExcel.Cells(5, 26).Value = "=Y5/AE5"
	ObjExcel.Cells(3, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(3, 27).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(4, 27).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(5, 27).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(3, 28).Value = "=AA3/AE3"
	ObjExcel.Cells(4, 28).Value = "=AA4/AE4"
	ObjExcel.Cells(5, 28).Value = "=AA5/AE5"
	ObjExcel.Cells(3, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(3, 29).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(4, 29).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(5, 29).Value = "=COUNTIFS(Table1[Interview ER], "&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(3, 30).Value = "=AC3/AE3"
	ObjExcel.Cells(4, 30).Value = "=AC4/AE4"
	ObjExcel.Cells(5, 30).Value = "=AC5/AE5"
	ObjExcel.Cells(3, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(4, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(5, 30).NumberFormat = "0.00%"


	ObjExcel.Cells(6, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&")"
	ObjExcel.Cells(7, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(8, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(6, 13).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(7, 13).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(8, 13).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(6, 14).Value = "=M6/AE6"
	ObjExcel.Cells(7, 14).Value = "=M7/AE7"
	ObjExcel.Cells(8, 14).Value = "=M8/AE8"
	ObjExcel.Cells(6, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(7, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(8, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 17).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(7, 17).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(8, 17).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(6, 18).Value = "=Q6/AE6"
	ObjExcel.Cells(7, 18).Value = "=Q7/AE7"
	ObjExcel.Cells(8, 18).Value = "=Q8/AE8"
	ObjExcel.Cells(6, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(7, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(8, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 19).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(7, 19).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(8, 19).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(6, 20).Value = "=Q6/AE6"
	ObjExcel.Cells(7, 20).Value = "=Q7/AE7"
	ObjExcel.Cells(8, 20).Value = "=Q8/AE8"
	ObjExcel.Cells(6, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(7, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(8, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 21).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(7, 21).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(8, 21).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(6, 22).Value = "=U6/AE6"
	ObjExcel.Cells(7, 22).Value = "=U7/AE7"
	ObjExcel.Cells(8, 22).Value = "=U8/AE8"
	ObjExcel.Cells(6, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(7, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(8, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 23).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(7, 23).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(8, 23).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(6, 24).Value = "=W6/AE6"
	ObjExcel.Cells(7, 24).Value = "=W7/AE7"
	ObjExcel.Cells(8, 24).Value = "=W8/AE8"
	ObjExcel.Cells(6, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(7, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(8, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 25).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(7, 25).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(8, 25).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(6, 26).Value = "=Y6/AE6"
	ObjExcel.Cells(7, 26).Value = "=Y7/AE7"
	ObjExcel.Cells(8, 26).Value = "=Y8/AE8"
	ObjExcel.Cells(6, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(7, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(8, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 27).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(7, 27).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(8, 27).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(6, 28).Value = "=AA6/AE6"
	ObjExcel.Cells(7, 28).Value = "=AA7/AE7"
	ObjExcel.Cells(8, 28).Value = "=AA8/AE8"
	ObjExcel.Cells(6, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(7, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(8, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(6, 29).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(7, 29).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(8, 29).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(6, 30).Value = "=AC6/AE6"
	ObjExcel.Cells(7, 30).Value = "=AC7/AE7"
	ObjExcel.Cells(8, 30).Value = "=AC8/AE8"
	ObjExcel.Cells(6, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(7, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(8, 30).NumberFormat = "0.00%"


	ObjExcel.Cells(9, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&")"
	ObjExcel.Cells(10, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(11, 31).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(9, 13).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(10, 13).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(11, 13).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(9, 14).Value = "=M9/AE9"
	ObjExcel.Cells(10, 14).Value = "=M10/AE10"
	ObjExcel.Cells(11, 14).Value = "=M11/AE11"
	ObjExcel.Cells(9, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(10, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(11, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(9, 17).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&_
	",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(10, 17).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(11, 17).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(9, 18).Value = "=Q9/AE9"
	ObjExcel.Cells(10, 18).Value = "=Q10/AE10"
	ObjExcel.Cells(11, 18).Value = "=Q11/AE11"
	ObjExcel.Cells(9, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(10, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(11, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(9, 19).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(10, 19).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(11, 19).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(9, 20).Value = "=Q9/AE9"
	ObjExcel.Cells(10, 20).Value = "=Q10/AE10"
	ObjExcel.Cells(11, 20).Value = "=Q11/AE11"
	ObjExcel.Cells(9, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(10, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(11, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(9, 21).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(10, 21).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(11, 21).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(9, 22).Value = "=U9/AE9"
	ObjExcel.Cells(10, 22).Value = "=U10/AE10"
	ObjExcel.Cells(11, 22).Value = "=U11/AE11"
	ObjExcel.Cells(9, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(10, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(11, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(9, 23).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(10, 23).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(11, 23).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(9, 24).Value = "=W9/AE9"
	ObjExcel.Cells(10, 24).Value = "=W10/AE10"
	ObjExcel.Cells(11, 24).Value = "=W11/AE11"
	ObjExcel.Cells(9, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(10, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(11, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(9, 25).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(10, 25).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(11, 25).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(9, 26).Value = "=Y9/AE9"
	ObjExcel.Cells(10, 26).Value = "=Y10/AE10"
	ObjExcel.Cells(11, 26).Value = "=Y11/AE11"
	ObjExcel.Cells(9, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(10, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(11, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(9, 27).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(10, 27).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(11, 27).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(9, 28).Value = "=AA9/AE9"
	ObjExcel.Cells(10, 28).Value = "=AA10/AE10"
	ObjExcel.Cells(11, 28).Value = "=AA11/AE11"
	ObjExcel.Cells(9, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(10, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(11, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(9, 29).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(10, 29).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(11, 29).Value = "=COUNTIFS(Table1[Interview ER],"&is_false&",Table1[No Interview ER],"&is_false&", Table1[Current SR],"&is_true&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(9, 30).Value = "=AC9/AE9"
	ObjExcel.Cells(10, 30).Value = "=AC10/AE10"
	ObjExcel.Cells(11, 30).Value = "=AC11/AE11"
	ObjExcel.Cells(9, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(10, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(11, 30).NumberFormat = "0.00%"


	ObjExcel.Cells(12, 31).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&")"
	ObjExcel.Cells(13, 31).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(14, 31).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(12, 13).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(13, 13).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(14, 13).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(12, 14).Value = "=M12/AE12"
	ObjExcel.Cells(13, 14).Value = "=M13/AE13"
	ObjExcel.Cells(14, 14).Value = "=M14/AE14"
	ObjExcel.Cells(12, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(12, 15).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(13, 15).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(14, 15).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(12, 16).Value = "=O12/AE12"
	ObjExcel.Cells(13, 16).Value = "=O13/AE13"
	ObjExcel.Cells(14, 16).Value = "=O14/AE14"
	ObjExcel.Cells(12, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(12, 17).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(13, 17).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(14, 17).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(12, 18).Value = "=Q12/AE12"
	ObjExcel.Cells(13, 18).Value = "=Q13/AE13"
	ObjExcel.Cells(14, 18).Value = "=Q14/AE14"
	ObjExcel.Cells(12, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(12, 19).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(13, 19).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(14, 19).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(12, 20).Value = "=Q12/AE12"
	ObjExcel.Cells(13, 20).Value = "=Q13/AE13"
	ObjExcel.Cells(14, 20).Value = "=Q14/AE14"
	ObjExcel.Cells(12, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(12, 21).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(13, 21).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(14, 21).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(12, 22).Value = "=U12/AE12"
	ObjExcel.Cells(13, 22).Value = "=U13/AE13"
	ObjExcel.Cells(14, 22).Value = "=U14/AE14"
	ObjExcel.Cells(12, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(12, 23).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(13, 23).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(14, 23).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(12, 24).Value = "=W12/AE12"
	ObjExcel.Cells(13, 24).Value = "=W13/AE13"
	ObjExcel.Cells(14, 24).Value = "=W14/AE14"
	ObjExcel.Cells(12, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(12, 25).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(13, 25).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(14, 25).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(12, 26).Value = "=Y12/AE12"
	ObjExcel.Cells(13, 26).Value = "=Y13/AE13"
	ObjExcel.Cells(14, 26).Value = "=Y14/AE14"
	ObjExcel.Cells(12, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(12, 27).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(13, 27).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(14, 27).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(12, 28).Value = "=AA12/AE12"
	ObjExcel.Cells(13, 28).Value = "=AA13/AE13"
	ObjExcel.Cells(14, 28).Value = "=AA14/AE14"
	ObjExcel.Cells(12, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(12, 29).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(13, 29).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(14, 29).Value = "=COUNTIFS(Table1[Notes], "&chr(34)&"PRIV Case."&chr(34)&",Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(12, 30).Value = "=AC12/AE12"
	ObjExcel.Cells(13, 30).Value = "=AC13/AE13"
	ObjExcel.Cells(14, 30).Value = "=AC14/AE14"
	ObjExcel.Cells(12, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(13, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(14, 30).NumberFormat = "0.00%"


	ObjExcel.Cells(15, 31).Value = "=COUNTA(Table1[Case number])"
	ObjExcel.Cells(16, 31).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(17, 31).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&")"

	ObjExcel.Cells(15, 13).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(16, 13).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(17, 13).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[CAF Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(15, 14).Value = "=M15/AE15"
	ObjExcel.Cells(16, 14).Value = "=M16/AE16"
	ObjExcel.Cells(17, 14).Value = "=M17/AE17"
	ObjExcel.Cells(15, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 14).NumberFormat = "0.00%"
	ObjExcel.Cells(15, 15).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")+COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[SNAP ("&date_header&")], "&is_blank&", Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(16, 15).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(17, 15).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&is_not_blank&",Table1[Intvw Date ("&date_header&")], "&is_not_blank&")"
	ObjExcel.Cells(15, 16).Value = "=O15/AE15"
	ObjExcel.Cells(16, 16).Value = "=O16/AE16"
	ObjExcel.Cells(17, 16).Value = "=O17/AE17"
	ObjExcel.Cells(15, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 16).NumberFormat = "0.00%"
	ObjExcel.Cells(15, 17).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(16, 17).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(17, 17).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"I"&chr(34)&")"
	ObjExcel.Cells(15, 18).Value = "=Q15/AE15"
	ObjExcel.Cells(16, 18).Value = "=Q16/AE16"
	ObjExcel.Cells(17, 18).Value = "=Q17/AE17"
	ObjExcel.Cells(15, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 18).NumberFormat = "0.00%"
	ObjExcel.Cells(15, 19).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(16, 19).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(17, 19).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"U"&chr(34)&")"
	ObjExcel.Cells(15, 20).Value = "=Q15/AE15"
	ObjExcel.Cells(16, 20).Value = "=Q16/AE16"
	ObjExcel.Cells(17, 20).Value = "=Q17/AE17"
	ObjExcel.Cells(15, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 20).NumberFormat = "0.00%"
	ObjExcel.Cells(15, 21).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(16, 21).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(17, 21).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"N"&chr(34)&")"
	ObjExcel.Cells(15, 22).Value = "=U15/AE15"
	ObjExcel.Cells(16, 22).Value = "=U16/AE16"
	ObjExcel.Cells(17, 22).Value = "=U17/AE17"
	ObjExcel.Cells(15, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 22).NumberFormat = "0.00%"
	ObjExcel.Cells(15, 23).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(16, 23).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(17, 23).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"A"&chr(34)&")"
	ObjExcel.Cells(15, 24).Value = "=W15/AE15"
	ObjExcel.Cells(16, 24).Value = "=W16/AE16"
	ObjExcel.Cells(17, 24).Value = "=W17/AE17"
	ObjExcel.Cells(15, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 24).NumberFormat = "0.00%"
	ObjExcel.Cells(15, 25).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(16, 25).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(17, 25).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"O"&chr(34)&")"
	ObjExcel.Cells(15, 26).Value = "=Y15/AE15"
	ObjExcel.Cells(16, 26).Value = "=Y16/AE16"
	ObjExcel.Cells(17, 26).Value = "=Y17/AE17"
	ObjExcel.Cells(15, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 26).NumberFormat = "0.00%"
	ObjExcel.Cells(15, 27).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(16, 27).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(17, 27).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"T"&chr(34)&")"
	ObjExcel.Cells(15, 28).Value = "=AA15/AE15"
	ObjExcel.Cells(16, 28).Value = "=AA16/AE16"
	ObjExcel.Cells(17, 28).Value = "=AA17/AE17"
	ObjExcel.Cells(15, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 28).NumberFormat = "0.00%"
	ObjExcel.Cells(15, 29).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")+COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&",Table1[SNAP ("&date_header&")], "&is_blank&")"
	ObjExcel.Cells(16, 29).Value = "=COUNTIFS(Table1[CASH ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(17, 29).Value = "=COUNTIFS(Table1[SNAP ("&date_header&")], "&chr(34)&"D"&chr(34)&")"
	ObjExcel.Cells(15, 30).Value = "=AC15/AE15"
	ObjExcel.Cells(16, 30).Value = "=AC16/AE16"
	ObjExcel.Cells(17, 30).Value = "=AC17/AE17"
	ObjExcel.Cells(15, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(16, 30).NumberFormat = "0.00%"
	ObjExcel.Cells(17, 30).NumberFormat = "0.00%"

	ObjExcel.Range("M3:P3").Interior.ColorIndex = 6
	ObjExcel.Range("W3:X3").Interior.ColorIndex = 6
	ObjExcel.Range("AE3:AE3").Interior.ColorIndex = 6
	ObjExcel.Range("M6:N6").Interior.ColorIndex = 6
	ObjExcel.Range("Q6:R6").Interior.ColorIndex = 6
	ObjExcel.Range("W6:X6").Interior.ColorIndex = 6
	ObjExcel.Range("AE6:AE6").Interior.ColorIndex = 6
	ObjExcel.Range("M9:N9").Interior.ColorIndex = 6
	ObjExcel.Range("Q9:R9").Interior.ColorIndex = 6
	ObjExcel.Range("W9:X9").Interior.ColorIndex = 6
	ObjExcel.Range("AE9:AE9").Interior.ColorIndex = 6
	ObjExcel.Range("W15:X15").Interior.ColorIndex = 6

	'Query date/time/runtime info
	objExcel.Cells(1, 33).Font.Bold = TRUE
	objExcel.Cells(2, 33).Font.Bold = TRUE
	ObjExcel.Cells(1, 33).Value = "Query date and time:"
	ObjExcel.Cells(2, 33).Value = "Query runtime (in seconds):"
	ObjExcel.Cells(1, 34).Value = now
	ObjExcel.Cells(2, 34).Value = timer - query_start_time

	'https://docs.microsoft.com/en-us/office/vba/api/excel.xlcolorindex - This is where you find all the excel numbers to use are.
	border_array = array("B1:C"&last_row, "D1:D"&last_row, "E1:E"&last_row, "F1:G"&last_row, "H1:I"&last_row, "A2:I2", "A3:I5", "A6:I8", "B10:I10", "K3:AE5", "K6:AE8", "K9:AE11", "K12:AE14", "K15:AE17", "M1:N17", "O1:P17",_
	 					 "Q1:R17", "S1:T17", "U1:V17", "W1:X17", "Y1:Z17", "AA1:AB17", "AC1:AD17", "AE1:AE17")

	For each group in border_array
		With ObjExcel.ActiveSheet.Range(group)
			With .Borders(7)	'left'
				.LineStyle = 1
				.Weight = 2
				.ColorIndex = -4105
			End With
			With .Borders(8)	'Top'
				.LineStyle = 1
				.Weight = 2
				.ColorIndex = -4105
			End With
			With .Borders(9)	'Bottom'
				.LineStyle = 1
				.Weight = 2
				.ColorIndex = -4105
			End With
			With .Borders(10)	'Right'
				.LineStyle = 1
				.Weight = 2
				.ColorIndex = -4105
			End With
		End With
	Next

	For xl_col = 1 to 34
		ObjExcel.columns(xl_col).AutoFit()
	Next

	run_time = timer - query_start_time
	end_msg = "Case details have been added to the Review Report" & vbCr & vbCr & "Run time: " & run_time & " seconds."
ElseIf renewal_option = "Send Appointment Letters" Then
	MAXIS_footer_month = CM_mo							'Setting the footer month and year based on the review month. We do not run statistics in CM + 2
	MAXIS_footer_year = CM_yr

	'This is where the review report is currently saved.
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\" & report_date & " Review Report.xlsx"

	'Initial Dialog which requests a file path for the excel file
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 361, 65, "On Demand Recertifications - Send Appointment Notices"
	  EditBox 130, 20, 175, 15, excel_file_path
	  ButtonGroup ButtonPressed
		PushButton 310, 20, 45, 15, "Browse...", select_a_file_button
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

	objRange.Insert(xlShiftToRight)			'We neeed one more columns

	notc_col = col_to_use		'Setting the column to individual variables so we enter the found information in the right place

	date_month = DatePart("m", date)		'Creating a variable to enter in the column headers
	date_day = DatePart("d", date)
	date_header = date_month & "-" & date_day

	ObjExcel.Cells(1, notc_col).Value = "APPT NOTC on " & date_header & ""			'creating the column headers for the statistics information for the day of the run.

	FOR i = col_to_use to col_to_use + 5									'formatting the cells'
		objExcel.Cells(1, i).Font.Bold = True		'bold font'
		ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
		objExcel.Columns(i).AutoFit()				'sizing the columns'
	NEXT

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
		MAXIS_case_number 	= trim(ObjExcel.Cells(excel_row,  2).Value)			'getting the case number from the spreadsheet
		forms_to_arep = ""
		forms_to_swkr = ""

		Call read_boolean_from_excel(ObjExcel.Cells(excel_row,  3).Value, er_with_intherview)
		Call read_boolean_from_excel(objExcel.cells(excel_row,  6).value, MFIP_status)
		Call read_boolean_from_excel(objExcel.cells(excel_row, 13).value, SNAP_status)

		' If er_with_intherview = True Then
		' 	MsgBox er_with_intherview & vbNewLine & "READING AS A BOOLEAN and TRUE"
		' ElseIf er_with_intherview = False Then
		' 	MsgBox er_with_intherview & vbNewLine & "READING AS A BOOLEAN and FALSE"
		' Else
		' 	MsgBox er_with_intherview & vbNewLine & "Sad"
		' End If
		If MFIP_status = True and SNAP_status = True Then programs = "MFIP/SNAP"
		If MFIP_status = True Then programs = "MFIP"
		If SNAP_status = True Then programs = "SNAP"
		interview_end_date = CM_plus_1_mo & "/15/" & CM_plus_1_yr
		last_day_of_recert = CM_plus_2_mo & "/01/" & CM_plus_2_yr
	    last_day_of_recert = dateadd("D", -1, last_day_of_recert)

		If er_with_intherview = True Then
			'Writing the SPEC MEMO - dates will be input from the determination made earlier.
			' MsgBox "We're writing a MEMO here"
			Call start_a_new_spec_memo_and_continue(memo_started)

			IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY

				CALL write_variable_in_SPEC_MEMO("The Department of Human Services sent you a packet of paperwork. This paperwork is to renew your " & programs & " case.")
				CALL write_variable_in_SPEC_MEMO("")
				' CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You must also complete an interview for your " & programs & " case to continue.")
				CALL write_variable_in_SPEC_MEMO("Please sign, date and return the renewal paperwork by " & CM_plus_1_mo & "/08/" & CM_plus_1_yr & ". You may need to complete an interview for your " & programs & " case to continue.")
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
				CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
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
						successful_notices = successful_notices + 1                 'For statistical purposes
						Exit Do
					End If

					memo_row = memo_row + 1           'Looking at next row'
				Loop Until create_date = "        "

			ELSE
				ObjExcel.Cells(excel_row, notc_col).Value = "N"         'Setting this as N if the MEMO failed
				call back_to_SELF
			END IF
		Else
			ObjExcel.Cells(excel_row, notc_col).Value = "N/A"
		End If

		If ObjExcel.Cells(excel_row, notc_col).Value = "Y" Then

			Call start_a_new_spec_memo_and_continue(memo_started)   'Starting a MEMO to send information about verifications

			IF memo_started = True THEN

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
				CALL write_variable_in_SPEC_MEMO("You now have an option to use an email to return documents to Hennepin County. Write the case number and full name associated with the case in the body of the email. Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure. To obtain information about your case please contact your worker. EMAIL: hhsews@hennepin.us ")
				CALL write_variable_in_SPEC_MEMO("If you have questions about the type of verifications needed, call 612-596-1300 and someone will assist you.")

				PF4 'Submit the MEMO'


			End If

			start_a_blank_case_note
			EMSendKey("*** Notice of " & programs & " Recertification Interview Sent ***")
			CALL write_variable_in_case_note("* A notice has been sent to client with detail about how to call in for an interview.")
			CALL write_variable_in_case_note("* Client must submit paperwork and call 612-596-1300 to complete interview.")
			If forms_to_arep = "Y" then call write_variable_in_case_note("* Copy of notice sent to AREP.")
			If forms_to_swkr = "Y" then call write_variable_in_case_note("* Copy of notice sent to Social Worker.")
			call write_variable_in_case_note("---")
			CALL write_variable_in_case_note("Link to Domestic Violence Brochure sent to client in SPEC/MEMO as a part of interview notice.")
			call write_variable_in_case_note("---")
			call write_variable_in_case_note(worker_signature)

			PF3
		End If

		excel_row = excel_row + 1
	Loop until MAXIS_case_number = ""

	is_true = chr(34)&"TRUE"&chr(34)

	'Going to another sheet, to enter worker-specific statistics and naming it
	sheet_name = "APPT NOTC on " & date_month & "-" & date_day
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

    if successful_notices = "" then successful_notices = 0
    objExcel.Cells(entry_row, 1).Value      = "Appointment Notices Sent"        'number of notices that were successful
    objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
	objExcel.Cells(entry_row, 2).Value      = "=COUNTIFS(Table1[APPT NOTC on " & date_header & "]," & Chr(34) & "Y" & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
    appt_row = entry_row
    entry_row = entry_row + 1

    objExcel.Cells(entry_row, 1).Value      = "Percentage successful"           'calculation of the percent of successful notices
    objExcel.Cells(entry_row, 1).Font.Bold 	= TRUE
    objExcel.Cells(entry_row, 2).Value      = "=B" & appt_row & "/B" & total_row
    objExcel.Cells(entry_row, 2).NumberFormat = "0.00%"		'Formula should be percent
    entry_row = entry_row + 1



ElseIf renewal_option = "Create Worklist" Then

	MAXIS_footer_month = REPT_month							'Setting the footer month and year based on the review month. We do not run statistics in CM + 2
	MAXIS_footer_year = REPT_year

	'This is where the review report is currently saved.
	excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\" & report_date & " Review Report.xlsx"

	'Initial Dialog which requests a file path for the excel file
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 361, 70, "On Demand Recertifications"
	  EditBox 130, 20, 175, 15, excel_file_path
	  ButtonGroup ButtonPressed
		PushButton 310, 20, 45, 15, "Browse...", select_a_file_button
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

	'Stats option ignores the 'list of workers' since it works off of an existing Excel, it needs to pull all of the workers
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)

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
Else
    end_msg = "No discrepancy report available yet."
End if

STATS_counter = STATS_counter - 1
script_end_procedure(end_msg)
