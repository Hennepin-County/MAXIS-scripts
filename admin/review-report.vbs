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
const SNAP_next_ER_const    = 14
const MA_status_const       = 15
const MSP_status_const      = 16
const HC_next_SR_const      = 17
const HC_next_ER_const      = 18
const Language_const        = 19
const Interpreter_const     = 20
const phone_1_const         = 21
const phone_2_const         = 22
const phone_3_const         = 23
const CASH_revw_status_const= 24
const SNAP_revw_status_const= 25
const HC_revw_status_const	= 26
const HC_MAGI_code_const	= 27
const review_recvd_const	= 28
const interview_date_const	= 29
const notes_const           = 30

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
  DropListBox 90, 35, 90, 15, "Select one..."+chr(9)+"Create Renewal Report"+chr(9)+"Discrepancy Run"+chr(9)+"Collect Statistics"+chr(9)+"Create Worklist", renewal_option
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

If renewal_option = "Collect Statistics" Then

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
    			If ( ( trim(SNAP_status) = "N" or trim(SNAP_status) = "I" or trim(SNAP_status) = "U" ) or ( trim(cash_status) = "N" or trim(cash_status) = "I" or trim(cash_status) = "U" ) _
                or ( trim(HC_status) = "N" or trim(HC_status) = "I" or trim(HC_status) = "U" )) then
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
        'Incremented variables
        recert_cases = recert_cases + 1                 'array incrementor
        STATS_counter = STATS_counter + 1               'stats incrementor
        excel_row = excel_row + 1                       'Excel row incrementor
    LOOP

    '----------------------------------------------------------------------------------------------------MAXIS TIME
    back_to_SELF
    MAXIS_footer_month = CM_mo_plus_one
    MAXIS_footer_year = CM_yr_plus_one
    Call MAXIS_footer_month_confirmation

    total_cases_review = 0  'for total recert counts for stats
    excel_row = 2          'resetting excel_row to output the array information

    'DO 'Loops until there are no more cases in the Excel list
    For item = 0 to Ubound(review_array, 2)
    	MAXIS_case_number = review_array(case_number_const, item)

        Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv) 'function to check PRIV status
        If is_this_priv = True then
            review_array(notes_const, item) = "PRIV Case."
            review_array(interview_const, item) = ""
            review_array(no_interview_const, item) = ""
            review_array(current_SR_const, item) = ""
        Else
            EmReadscreen worker_prefix, 4, 21, 14
            If worker_prefix <> "X127" then
                review_array(notes_const, i) = "Out-of-County: " & right(worker_prefix, 2)
                review_array(notes_const, item) = "PRIV Case."
                review_array(interview_const, item) = ""
                review_array(no_interview_const, item) = ""
                review_array(current_SR_const, item) = ""
            Else
                'function to determine programs and the program's status---Yay Casey!
                Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)

                If case_active = False then
                    review_array(notes_const, item) = "Case Not Active."
                Else
                    'valuing the array variables from the inforamtion gathered in from CASE/CURR
                    review_array(MFIP_status_const, item) = mfip_case
                    review_array(DWP_status_const,  item) = dwp_case
                    review_array(GA_status_const,   item) = ga_case
                    review_array(MSA_status_const,  item) = msa_case
                    review_array(GRH_status_const,  item) = grh_case
                    review_array(SNAP_status_const, item) = snap_case
                    review_array(MA_status_const,   item) = ma_case
                    review_array(MSP_status_const,  item) = msp_case
                    '----------------------------------------------------------------------------------------------------STAT/REVW
                    CALL navigate_to_MAXIS_screen("STAT", "REVW")

                    If family_cash_case = True or adult_cash_case = True then
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
                            IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN
                                review_array(current_SR_const, item) = True
                            Else
                                review_array(current_SR_const, item) = False
                            End if

                            'Determining if a case is ER, and if it meets interview requirement
                            IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then
                                If mfip_case = True then review_array(interview_const, item) = True             'MFIP interview requirement
                                IF adult_cash_case = True then review_array(no_interview_const, item) = True    'Adult CASH programs do not meet interview requirement
                            Elseif recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) then
                                review_array(interview_const, item) = False
                                review_array(no_interview_const, item) = False
                            End if

                            'Next CASH ER and SR dates
                            review_array(CASH_next_SR_const, item) = CASH_CSR_date
                            review_array(CASH_next_ER_const, item) = CASH_ER_date
                        Else
                            review_array(notes_const, i) = "Unable to Access CASH Review Information."
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
                            IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN
                                review_array(current_SR_const, item) = True
                            Else
                                review_array(current_SR_const, item) = False
                            End if

                            If recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) then review_array(interview_const, item) = False
                            IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then review_array(interview_const, item) = True

                            'Next SNAP ER and SR dates
                            review_array(SNAP_next_SR_const, item) = SNAP_CSR_date
                            review_array(SNAP_next_ER_const, item) = SNAP_ER_date
                        Else
                            review_array(notes_const, i) = "Unable to Access FS Review Information."
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
                            IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN
                                review_array(current_SR_const, item) = True
                            Else
                                review_array(current_SR_const, item) = False
                            End if

                            IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) then review_array(no_interview_const, item) = True
                            If recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) then review_array(no_interview_const, item) = False

                            'Next HC ER and SR dates
                            review_array(HC_next_SR_const, item) = HC_CSR_date
                            review_array(HC_next_ER_const, item) = HC_ER_date

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
        	    If phone_number_one <> "( ___ ) ___ ____" then review_array(phone_1_const, item) = phone_number_one
        	    EMReadScreen phone_number_two, 16, 18, 43
        	    If phone_number_two <> "( ___ ) ___ ____" then review_array(phone_2_const, item) = phone_number_two
        	    EMReadScreen phone_number_three, 16, 19, 43
        	    If phone_number_three <> "( ___ ) ___ ____" then review_array(phone_3_const, item) = phone_number_three

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

        	    review_array(Interpreter_const, item) = interpreter_code
        	    review_array(Language_const, item) = language_coded
            End if
        End if
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
ElseIf renewal_option = "Collect Statistics" Then


	t_drive = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team"
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
	  ' Text 10, 25, 340, 30, "This script will send an Appointment Notice or NOMI for recertification for a list of cases in a county that currently has an On Demand Waiver in effect for interviews. If your county does not have this waiver, this script should not be used."
	  Text 10, 25, 120, 10, "Select an Excel file for recert cases:"
	EndDialog

	'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
	'Show initial dialog
	Do
		Dialog Dialog1
		cancel_confirmation
		If ButtonPressed = select_a_file_button then call file_selection_system_dialog(excel_file_path, ".xlsx")
	Loop until ButtonPressed = OK and excel_file_path <> ""

	'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
	call excel_open(excel_file_path, True, True, ObjExcel, objWorkbook)

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

	col_to_use = 0
	Do
		col_to_use = col_to_use + 1
		col_header = trim(ObjExcel.Cells(1, col_to_use).Value)
	Loop until col_header = ""

	last_col_letter = convert_digit_to_excel_column(col_to_use)
	'Insert columns in excel for additional information to be added
	column_end = last_col_letter & "1"
	Set objRange = objExcel.Range(column_end).EntireColumn

	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)
	objRange.Insert(xlShiftToRight)

	cash_stat_excel_col = col_to_use
	snap_stat_excel_col = col_to_use + 1
	hc_stat_excel_col = col_to_use + 2
	magi_stat_excel_col = col_to_use + 3
	recvd_date_excel_col = col_to_use + 4
	intvw_date_excel_col = col_to_use + 5

	date_month = DatePart("m", date)
	date_day = DatePart("d", date)
	date_header = date_month & "/" & date_day

	ObjExcel.Cells(1, cash_stat_excel_col).Value = "CASH (" & date_header & ")"
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
	' worker_array = ARRAY("X127ET6", "X127EZ2", "X127ET4", "X127ET5")

	'Establish the reviews array
	recert_cases = 0	            'incrementor for the array


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
				EMReadScreen MAGI_status, 4, row, 55
				EMReadScreen recvd_date, 8, row, 62
				EMReadScreen intvw_date, 8, row, 72

    			'Navigates though until it runs out of case numbers to read
    			IF MAXIS_case_number = "" then exit do

    			'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
    			If cash_status = "-" 	then cash_status = ""
    			If SNAP_status = "-" 	then SNAP_status = ""
    			If HC_status = "-" 		then HC_status = ""

    			'Using if...thens to decide if a case should be added (status isn't blank)
    			If ( ( trim(SNAP_status) = "N" or trim(SNAP_status) = "I" or trim(SNAP_status) = "U" ) or ( trim(cash_status) = "N" or trim(cash_status) = "I" or trim(cash_status) = "U" ) _
                or ( trim(HC_status) = "N" or trim(HC_status) = "I" or trim(HC_status) = "U" )) then
                    'Adding the case information to the array
					ReDim Preserve review_array(notes_const, recert_cases)

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

                    recert_cases = recert_cases + 1
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

	excel_row = "2"
	Do
		case_number_to_check = trim(ObjExcel.Cells(excel_row, 2).Value)
		' MsgBox "Case Number to Check - " & case_number_to_check
		For revs_item = 0 to UBound(review_array, 2)
			' MsgBox "Case Number to Check - " & case_number_to_check & vbCr & "ARRAY Case Number - " & review_array(case_number_const, revs_item)

			If case_number_to_check = review_array(case_number_const, revs_item) Then
				ObjExcel.Cells(excel_row, cash_stat_excel_col).Value = review_array(CASH_revw_status_const, revs_item)
				ObjExcel.Cells(excel_row, snap_stat_excel_col).Value = review_array(SNAP_revw_status_const, revs_item)
				ObjExcel.Cells(excel_row, hc_stat_excel_col).Value = review_array(HC_revw_status_const, revs_item)
				ObjExcel.Cells(excel_row, magi_stat_excel_col).Value = review_array(HC_MAGI_code_const, revs_item)
				ObjExcel.Cells(excel_row, recvd_date_excel_col).Value = review_array(review_recvd_const, revs_item)
				ObjExcel.Cells(excel_row, intvw_date_excel_col).Value = review_array(interview_date_const, revs_item)
				Exit For
			End If
		Next
		excel_row = excel_row + 1
	Loop until case_number_to_check = ""

Else
    msgbox "No discrepancy report available yet."
End if

STATS_counter = STATS_counter - 1
script_end_procedure("Success! The review report is ready.")
