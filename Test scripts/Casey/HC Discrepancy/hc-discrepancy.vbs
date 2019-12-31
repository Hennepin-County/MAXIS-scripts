'Required for statistical purposes==========================================================================================
name_of_script = "BULK - SPENDDOWN REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 255                      'manual run time in seconds
STATS_denomination = "M"       			   'M is for each Member
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
CALL changelog_update("01/25/2018", "Entering a supervisor X-Number in the Workers to Check will pull all X-Numbers listed under that supervisor in MAXIS. Addiional bug fix where script was missing cases.", "Casey Love, Hennepin County")
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

function navigate_to_spec_MMIS_region(group_security_selection)
'--- This function is to be used when navigating to MMIS from another function in BlueZone (MAXIS, PRISM, INFOPAC, etc.)
'~~~~~ group_security_selection: region of MMIS to access - programed options are "CTY ELIG STAFF/UPDATE", "GRH UPDATE", "GRH INQUIRY", "MMIS MCRE"
'===== Keywords: MMIS, navigate
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then
			Do
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 81, 25, "Dialog"
                  Text 10, 10, 65, 10, "Need Password"
                EndDialog

                dialog Dialog1

				CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	 		Loop until are_we_passworded_out = false					'loops until user passwords back in
		End if
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	group_security_selection = UCASE(group_security_selection)

	EMReadScreen MMIS_menu, 24, 3, 30
	If MMIS_menu = "GROUP SECURITY SELECTION" Then
		EMReadScreen mmis_group_selection, 4, 1, 65
		EMReadScreen mmis_group_type, 4, 1, 57

		correct_group = FALSE

		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			mmis_group_selection_part = left(mmis_group_selection, 2)

			If mmis_group_selection_part = "C3" Then correct_group = TRUE
			If mmis_group_selection_part = "C4" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the County Eligibility region. The script will now stop.")

		Case "GRH UPDATE"
			If mmis_group_selection  = "GRHU" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Update region. The script will now stop.")

		Case "GRH INQUIRY"
			If mmis_group_selection  = "GRHI" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Inquiry region. The script will now stop.")

		Case "MMIS MCRE"
			If mmis_group_selection  = "EK01" Then correct_group = TRUE
			If mmis_group_selection  = "EKIQ" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the MCRE region. The script will now stop.")

		End Select

	Else
		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			row = 1
			col = 1
			EMSearch " C3", row, col
			If row <> 0 Then
				EMWriteScreen "X", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "X", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "X", row, col - 3
			transmit

		Case "GRH UPDATE"
			row = 1
			col = 1
			EMSearch "GRHU", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH INQUIRY"
			row = 1
			col = 1
			EMSearch "GRHI", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH Inquiry area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "MMIS MCRE"
			row = 1
			col = 1
			EMSearch "EK01", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the MCRE area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		End Select
	End If
end function

'function specific to this script - running_stopwatch and MX_environment are defined outside of this function
'meant to keep MMIS from passwording out while this long bulk script is running
function keep_MMIS_passworded_in()
    If timer - running_stopwatch > 720 Then         'this means the script has been running for more than 12 minutes since we last popped in to MMIS
        Call navigate_to_spec_MMIS_region("CTY ELIG STAFF/UPDATE")      'Going to MMIS'
        Call navigate_to_MAXIS(MX_environment)                          'going back to MAXIS'

        running_stopwatch = timer                                       'resetting the stopwatch'
    End If
end function


'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
get_county_code

'Connects to BlueZone
EMConnect ""

'Checking for MAXIS
Call check_for_MAXIS(True)
Call back_to_SELF
EmReadscreen MX_environment, 13, 22, 48
MX_environment = trim(MX_environment)
Call navigate_to_spec_MMIS_region("CTY ELIG STAFF/UPDATE")      'Going to MMIS'
Call navigate_to_MAXIS(MX_environment)

running_stopwatch = timer

'Setting up constants for ease of reading the array
Const wrk_num       = 0
Const case_num      = 1
Const next_revw     = 2
Const clt_name      = 3
Const ref_numb      = 4
Const clt_pmi       = 5
Const hc_prog_one   = 6
Const mmis_end_one  = 7
Const disc_one      = 8

Const elig_type_one = 9
Const elig_std_one  = 10
Const elig_mthd_one = 11
Const elig_waiv     = 12
Const mobl_spdn     = 13
Const spd_pd        = 14
Const hc_excess     = 15
Const hc_prog_two   = 16
Const mmis_end_two  = 17
Const disc_two      = 18

Const elig_type_two = 19
Const elig_std_two  = 20
Const elig_mthd_two = 21
Const mmis_spdn     = 22
Const error_notes   = 23
Const add_xcl       = 24

'Setting up the arrays to be dynamic
Dim HC_CASES_ARRAY()
ReDim HC_CASES_ARRAY (3, 0)

Dim HC_CLIENTS_DETAIL_ARRAY ()
ReDim HC_CLIENTS_DETAIL_ARRAY (add_xcl, 0)

'Setting this variable to determine a filter later
one_month_only = FALSE
'defining current footer month so the script doesn't go to old things
MAXIS_footer_month = right("00" & datepart("m", date), 2)
MAXIS_footer_year = right("00" & datepart("yyyy", date), 2)

'Setting the initial path for the excel file to be found at - so we don't have to clickity click a bunch to get to the right file.
hc_cases_excel_file_path = ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 155, "Pull REPT data into Excel dialog"
  EditBox 85, 20, 130, 15, worker_number
  EditBox 5, 110, 210, 15, hc_cases_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 165, 90, 50, 15, "Browse...", select_a_file_button
    OkButton 110, 135, 50, 15
    CancelButton 165, 135, 50, 15
  Text 50, 5, 125, 10, "*** REPT ON MAXIS SPENDDOW ***"
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 5, 60, 210, 25, "** If a supervisor 'x1 number' is entered, the script will add the 'x1 numbers' of all workers listed in MAXIS under that supervisor number."
  Text 100, 85, 15, 10, "OR"
  Text 5, 95, 135, 10, "Select an Excel file of MAXIS MA cases:"
EndDialog

'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
'Show initial dialog
Do
    Do
        err_msg = ""

    	Dialog Dialog1
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then
            call file_selection_system_dialog(hc_cases_excel_file_path, ".xlsx")
            err_msg = "LOOP" & err_msg
        End If
        If trim(worker_number) = "" AND trim(hc_cases_excel_file_path) = "" Then err_msg = err_msg & vbNewLine & "* Choose a source of cases, either a BOBI report or Basket Numbers."
        If trim(worker_number) <> "" AND trim(hc_cases_excel_file_path) <> "" Then err_msg = err_msg & vbNewLine & "* Choose one source of cases. Both a BOBI report and Basket Numbers cannot be entered at the same time."

        If err_msg <> "" Then
            If left(err_msg, 4) <> "LOOP" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
        End If
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

revw_month_list = "ALL"
'Sets the script up to only pull cases for certain months if selected from the dialog
If revw_month_list <> "ALL" AND revw_month_list <> "" Then
	one_month_only = TRUE 			'If any month is selected the script needs to filter
	Select Case revw_month_list
		Case "January"
			month_selected = 1
		Case "February"
			month_selected = 2
		Case "March"
			month_selected = 3
		Case "April"
			month_selected = 4
		Case "May"
			month_selected = 5
		Case "June"
			month_selected = 6
		Case "July"
			month_selected = 7
		Case "August"
			month_selected = 8
		Case "September"
			month_selected = 9
		Case "October"
			month_selected = 10
		Case "November"
			month_selected = 11
		Case "December"
			month_selected = 12
	End Select
End If

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

worker_number = trim(worker_number)

'Checking for MAXIS
Call check_for_MAXIS(True)

'If worker number information is selected, we need to gather HC client information from REPT ACTV and CASE PERS
If worker_number <> "" then

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

    hc_clt = 0      'setting this for the beginning of the array creation

    'Getting all the cases with HC active for each worker
    For each worker in worker_array
    	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
    	Call navigate_to_MAXIS_screen("rept", "actv")	'going to rept actv for each worker
    	EMWriteScreen worker, 21, 13
    	transmit
    	EMReadScreen user_worker, 7, 21, 71
    	EMReadScreen p_worker, 7, 21, 13
    	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

    	'Skips workers with no info
    	EMReadScreen has_content_check, 1, 7, 8
    	If has_content_check <> " " then

    		'Grabbing each case number on screen
    		Do
    			'Set variable for next do...loop
    			MAXIS_row = 7

    			'Checking for the last page of cases.
    			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
    			Do
    				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number
    				EMReadScreen client_name, 21, MAXIS_row, 21			'Reading client name
    				EMReadScreen next_revw_date, 8, MAXIS_row, 42		'Reading application date
    				EMReadScreen HC_status, 1, MAXIS_row, 64			'Reading HC status

    				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
    				MAXIS_case_number = trim(MAXIS_case_number)
    				If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
    				all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")

    				If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end

    				'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)
    				If HC_status = "A" then
    					If one_month_only = TRUE Then 						'If user has selected to only get cases with a certain reveiw month
    						If trim(next_revw_date) = "" Then
    							case_error = MsgBox ("Case " & MAXIS_case_number & " does not have a review listed, please check that STAT is coded correctly for this case." & vbNewLine & vbNewLine & "This case will not be added to the report, you should check for a spenddown manually.", vbAlert, "No Review Date")
    						Else
    							revw_month = abs(left(next_revw_date, 2))
    							If revw_month = month_selected Then 			'Compares the review month to the variable defined above in the Select Case
    								ReDim Preserve HC_CASES_ARRAY (3, hc_clt)		'Adds information about case with active HC to an array
    								HC_CASES_ARRAY(wrk_num, hc_clt)   = worker
    								HC_CASES_ARRAY(case_num, hc_clt)  = MAXIS_case_number
    								HC_CASES_ARRAY(next_revw, hc_clt) = next_revw_date
    								hc_clt = hc_clt + 1
    							End If
    						End If
    					Else
    						ReDim Preserve HC_CASES_ARRAY (3, hc_clt)			'Adds information about case with active HC to an array
    						HC_CASES_ARRAY(wrk_num, hc_clt)   = worker
    						HC_CASES_ARRAY(case_num, hc_clt)  = MAXIS_case_number
    						HC_CASES_ARRAY(next_revw, hc_clt) = next_revw_date
    						hc_clt = hc_clt + 1
    					End If
    				End If

    				MAXIS_row = MAXIS_row + 1       'going to the next case on the REPT/ACTV
    			Loop until MAXIS_row = 19           'This is the end of the page on the REPT
    			PF8                                 'Go to the next page of the REPT
    		Loop until last_page_check = "THIS IS THE LAST PAGE"      'This shows the end of the REPT
    	End if
        Call keep_MMIS_passworded_in
    next

    hc_clt = 0          'Resetting this as we are using this for the NEW array we are creating

    'The script will now look in each case at MOBL to identify clients that have spenddown listed on MOBL'\
    For hc_case = 0 to UBound(HC_CASES_ARRAY, 2)
        back_to_SELF                                                'Back to SELF at the beginning of each run so that we don't end up in the wrong case
    	MAXIS_case_number = HC_CASES_ARRAY(case_num, hc_case)		'defining case number for functions to use

        Call navigate_to_MAXIS_screen("CASE", "PERS")               'Getting client eligibility of HC from CASE PERS
        pers_row = 10                                               'This is where client information starts on CASE PERS
        Do
            EmReadscreen clt_hc_status, 1, pers_row, 61             'reading the HC status of each client
            If clt_hc_status = "A" Then                             'if HC is active then we will add this client to the array to find additional information
                EmReadscreen pers_ref_numb,  2, pers_row, 3         'reading the client information
                EmReadscreen pers_pmi_numb,  8, pers_row, 34
                EmReadscreen pers_last_name, 15, pers_row, 6
                EmReadscreen pers_frst_name, 11, pers_row, 22

                pers_pmi_numb = trim(pers_pmi_numb)
                pers_last_name = trim(pers_last_name)
                pers_frst_name = trim(pers_frst_name)

                ReDim Preserve HC_CLIENTS_DETAIL_ARRAY (add_xcl, hc_clt)        'adding client information to the client array

                HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) = HC_CASES_ARRAY(wrk_num,  hc_case)
                HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt) = HC_CASES_ARRAY(case_num, hc_case)
                'HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt) = ObjExcel.Cells(excel_row, ). Value
                HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt) = pers_last_name & ", " & pers_frst_name
                HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) = pers_ref_numb
                HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt) = pers_pmi_numb

                hc_clt = hc_clt + 1     'incrementing the array
            End If

            pers_row = pers_row + 3         'next client information is 3 rows down
            If pers_row = 19 Then           'this is the end of the list of client on each list
                PF8                         'going to the next page of client information
                pers_row = 10
                EmReadscreen end_of_list, 9, 24, 14
                If end_of_list = "LAST PAGE" Then Exit Do
            End If
            EmReadscreen next_pers_ref_numb, 2, pers_row, 3     'this reads for the end of the list

        Loop until next_pers_ref_numb = "  "

        'creating an end of script message
        all_the_workers = join(worker_array, ", ")
        end_msg = "Success! Client HC Eligibility and MMIS coding for workers: " & all_the_workers & " have been added to the spreadsheet."

        Call keep_MMIS_passworded_in
    Next

'If there are no worker numbers entered, then we are going to use a BOBI list of active HC clients
Else

    'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
    call excel_open(hc_cases_excel_file_path, True, True, ObjExcel, objWorkbook)

    excel_row_to_start = "5"    'presetting this before the dialog since BOBI case information starts on row 5

    'This is the dialog to limit the script run as the BOBI is in the tens of thousands
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 171, 115, "How long to run?"
      EditBox 25, 30, 30, 15, stop_time
      EditBox 65, 75, 30, 15, excel_row_to_start
      EditBox 65, 95, 30, 15, excel_row_to_end
      ButtonGroup ButtonPressed
        OkButton 115, 95, 50, 15
      Text 10, 10, 140, 20, "To time limit the run of the script enter the numeber of hours to run the script:"
      Text 65, 35, 50, 10, "Hours"
      Text 10, 55, 145, 20, "The run can be limited by indicating which rows of the Excel file to review/process:"
      Text 15, 80, 50, 10, "Excel to start"
      Text 15, 100, 45, 10, "Excel to end"
    EndDialog

    'showing the dialog
    Do
        Do
            err_msg = ""
            dialog Dialog1

            If trim(stop_time) <> "" AND IsNumeric(stop_time) = FALSE Then err_msg = err_msg & vbNewLine & "- Number of hours should be a number."
            If trim(excel_row_to_start) = "" Then err_msg = err_msg & vbNewLine & "- Indicate the excel row to start the run at."
            If trim(excel_row_to_start) <> "" AND IsNumeric(excel_row_to_start) = FALSE Then err_msg = err_msg & vbNewLine & "- Start row of Excel should be a number."
            If trim(excel_row_to_end) <> "" AND IsNumeric(excel_row_to_end) = FALSE Then err_msg = err_msg & vbNewLine & "- End row of Excel should be a number."

            If err_msg <> "" Then MsgBox "** Please Resolve the Following to Continue:" & vbNew & err_msg

        Loop until err_msg = ""
        call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

    'setting these to numbers
    excel_row = excel_row_to_start * 1
    If trim(excel_row_to_end) <> "" Then excel_row_to_end = excel_row_to_end * 1
    hc_clt = 0      'setting the beginning of the array

    'making stop time a number
    If trim(stop_time) <> "" Then
        stop_time = FormatNumber(stop_time, 2,          0,                 0,                      0)
                                'number     dec places  leading 0 - FALSE    neg nbr in () - FALSE   use deliminator(comma) - FALSE
        stop_time = stop_time * 60 * 60     'tunring hours to seconds

        'Since this happens in phases we need to limit the number of cases at the beginning to fit within the time frame.
        If trim(excel_row_to_end) = "" Then                 'if the excel row to end at was not predefined, use the average time per line to determine how much of the list ot read
            number_of_rows_to_review = stop_time/2.5        'TODO track if 2.5 seconds per case is a good time reference
            excel_row_to_end = excel_row + number_of_rows_to_review
        End If

        end_time = timer + stop_time        'timer is the number of seconds from 12:00 AM so we need to add the hours to run to the time to determine at what point the script should exit the loop
    Else
        end_time = 84600    'sets the end time for 11:30 PM so that is doesn't end out
    End If

    Do
        ReDim Preserve HC_CLIENTS_DETAIL_ARRAY (add_xcl, hc_clt)        'redim the array to add another case

        HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) = ObjExcel.Cells(excel_row, 3). Value       'adding information from the spreadsheet to the array
        HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt) = ObjExcel.Cells(excel_row, 2). Value
        'HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt) = ObjExcel.Cells(excel_row, ). Value
        HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt) = ObjExcel.Cells(excel_row, 8). Value
        HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) = right(ObjExcel.Cells(excel_row, 7). Value, 2)
        HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt) = ObjExcel.Cells(excel_row, 6). Value
        HC_CLIENTS_DETAIL_ARRAY (add_xcl,   hc_clt) = excel_row

        ' MsgBox "Worker: " & HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) & vbNewLine &_
        '        "Case: " & HC_CLIENTS_DETAIL_ARRAY (case_num,   hc_clt) & vbNewLine &_
        '        "Client: " & HC_CLIENTS_DETAIL_ARRAY (clt_name,   hc_clt) & vbNewLine &_
        '        "Ref Number: " & HC_CLIENTS_DETAIL_ARRAY (ref_numb,   hc_clt) & vbNewLine &_
        '        "PMI: " & HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt)
        excel_row = excel_row + 1           'incrementing to the next row and next place in the array
        hc_clt = hc_clt + 1

        next_case_number = ObjExcel.Cells(excel_row, 2). Value      'looking to see if we found the end of the list
        next_case_number = trim(next_case_number)
        Call keep_MMIS_passworded_in
        If excel_row = excel_row_to_end Then Exit Do
    Loop until next_case_number = ""

    'Setting the end message
    end_msg = "Success! Client HC Eligibility and MMIS coding for row " & excel_row_to_start & " to " & excel_row_to_end & " have been added to the spreadsheet."

    'closing excel
    ObjExcel.Quit
    Set ObjExcel = Nothing
End If

'Now the array is created - it is the same if we got it from REPT/PND2 or the BOBI
'Information gathering in MAXIS now for every client on HC on the list
For hc_clt = 0 to UBOUND(HC_CLIENTS_DETAIL_ARRAY, 2)
    back_to_SELF                                                        'resetting at each loop
    MAXIS_case_number = HC_CLIENTS_DETAIL_ARRAY(case_num, hc_clt)		'defining case number for functions to use
    CLIENT_reference_number = HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt)
    Call navigate_to_MAXIS_screen ("ELIG", "HC")						'Goes to ELIG HC
    APPROVAL_NEEDED = FALSE                                             'setting some booleans
    found_elig = FALSE
    client_found = FALSE
    row = 8                                                             'begining of the list of HH Membs in ELIG/HC
    Do
        EMReadScreen check_for_priv, 10, 24, 14                         'Some cases from the BOBI are high level priv and we cannot look at details
        If check_for_priv = "PRIVILEGED" Then
            HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = "PRIV"
            Exit Do
        End If

        EMReadscreen elig_clt, 2, row, 3                                'reading the information on the row to see if it is for the client
        EmReadscreen prog_exists, 1, row, 10
        If elig_clt = CLIENT_reference_number Then                      'If these match, we have found the client to find additional HC details
            'MsgBox "Elig Clt: " & elig_clt & vbNewLine & "Ref Number: " &CLIENT_reference_number
            client_found = TRUE                                         'setting the boolean for the rest to search for more
            EMReadScreen prog, 10, row, 28                              'reading all the program details on ELIG Memb List
            EMReadScreen version, 2, row, 58
            EMReadScreen app_indc, 6, row, 68

            prog = trim(prog)                                           'formatting the information that was read.
            app_indc = trim(app_indc)

            If prog = "NO REQUEST" OR prog = "NO VERSION" OR prog = "" Then                          'If there is no span known for the member, it will be indicated in this way and we can't see any more
                HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ HC information does not appear to be in MAXIS. ELIG/HC for this member - " & prog
                Exit Do
            End If

            If app_indc <> "APP" Then                                   'If the version is not Approved then we should try to find the approved version
                if version = "01" Then                                  'If this is the only version, it will be 01 and we will have to take an unapproved version
                    found_elig = TRUE
                    APPROVAL_NEEDED = TRUE                              'This indicates that the ELIG information in MAXIS may be out of date
                Else                                                    'if this isn't version 01 then we are going to try to find the approved version
                    Do                                                  'this is on a loop because we may need to look at multiple previous versions
                        EMReadScreen version, 2, row, 58                'reading the version (it is here as well because we need to reread it at every loop)
                        'MsgBox "1 - Version: " & version
                        version = version * 1                           'making this a number and not a string
                        prev_verision = version - 1                     'going to the number before the previous version
                        prev_verision = right("00" & prev_verision, 2)  'making it a string

                        EMWriteScreen prev_verision, row, 58            'writing the previous version on to the current row
                        transmit                                        'transmit to pull the detail about that version
                        EMReadScreen app_indc, 6, row, 68               'reading if this version has been approved or not
                        app_indc = trim(app_indc)
                        If app_indc = "APP" Then                        'If this version is approved - we do not need to loop any more
                            found_elig = TRUE
                            Exit Do
                        End If
                        'MsgBox "Loop 2 - prev_verision: " & prev_verision
                    Loop until prev_verision = "01"                     'We can only go to 01
                    EMReadScreen version, 2, row, 58
                    EMReadScreen app_indc, 6, row, 68
                    app_indc = trim(app_indc)
                    If version = "01" AND app_indc <> "APP" Then APPROVAL_NEEDED = TRUE     'if we finally get to version 01 and still have not found an approve version, this is set here
                End If
            Else
                found_elig = TRUE                                       'if the first version found is approved then we have already found the elig version
            End If

            If found_elig = TRUE Then                                   'if we found the elig information
                EMReadScreen prog, 10, row, 28                          'the script now reads the actual HC detail
                EMReadScreen result, 7, row, 41
                EMReadScreen hc_status, 7, row, 50

                prog = trim(prog)
                result = trim(result)
                hc_status = trim(hc_status)

                If result = "ELIG" AND hc_status = "ACTIVE" Then        'the clients that are eligible and active on ELIG HC - we will look in the HC Summ for more information
                    HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = prog      'setting this to the array

                    EmWriteScreen "X", row, 26                          'opening the HC BSUM
                    transmit

                    If prog = "MA" or prog = "IMD" Then                 'for the programs MA or IMD the information is in a certain place
                        If left(HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then   'If the name was not on the BOBI and is just listed on X's then we read the actual name here
                            EmReadscreen the_name, 30, 5, 20
                            the_name = trim(the_name)
                            HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt) = the_name
                        End If
                        mo_col = 19                                     'setting the column for reading the month and year of the HC information for the client
                        yr_col = 22
                        Do                                              'we will look through each of the 6 months in the budget to find the current month and year
                            EMReadScreen bsum_mo, 2, 6, mo_col          'reading the month and year
                            EMReadScreen bsum_yr, 2, 6, yr_col

                            If bsum_mo = MAXIS_footer_month and bsum_yr = MAXIS_footer_year Then Exit Do        'if it is this month and year, we found the right month and year
                            mo_col = mo_col + 11                        'if it doesn't match, then we go to the next - which is 11 over
                            yr_col = yr_col + 11
                            'MsgBox "Loop 3 - month col: " & mo_col
                        Loop until mo_col = 74                          'this is the last month

                        EMReadScreen reference, 2, 5, 16                'this is the reference number

                        EMReadScreen prog, 4, 11, mo_col                'reading all of the detail in this month of BSUM
                        EMReadScreen pers_type, 2, 12, mo_col-2
                        EMReadScreen pers_std, 1, 12, yr_col
                        EMReadScreen pers_mthd, 1, 13, yr_col-1
                        EMReadScreen pers_waiv, 1, 14, yr_col-1

                        'sometimes the month is not correctly found because of old budgets, this sets error information here because the case needs to be looked at manually
                        If prog = "    " Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ HC ELIG Budget may need approval or budget needs to be aligned."

                        'Took this out because I think it is for a different script issue with a dirrect run and no longer needed. Will test with it gone for a while'
                        ' If pers_type = "__" Then                        'TODO - REMOVE ON 9/28/18 if no longer needed - determine why I have this here - look at case # 132245
                        '     EMReadScreen cur_mo_test, 6, 7, mo_col
                        '     cur_mo_test = trim(cur_mo_test)
                        '     'MsgBox "This is come up when person test is __" & vbNewLine & "cur_mo_test is " & cur_mo_test
                        '     pers_type = cur_mo_test
                        '     pers_std = ""
                        '     pers_mthd = ""
                        ' End If

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_one, hc_clt) = pers_type     'setting all of the read information is added to the array
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_one,  hc_clt) = pers_std
                        HC_CLIENTS_DETAIL_ARRAY (elig_mthd_one, hc_clt) = pers_mthd
                        HC_CLIENTS_DETAIL_ARRAY (elig_waiv, hc_clt) = pers_waiv

                        'if this was found to be true in this loop, will add error note that the case needs review and approval
                        If APPROVAL_NEEDED = TRUE THen HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) & " ~ Budget Needs Approval"

                        EMWriteScreen "X", 18, 3        'Going in to MOBL
                        transmit

                        mobl_row = 6                    'setting the top of the list in MOBL (this lists the whole HH)'
                        Do
                            EMReadScreen ref_nbr, 2, mobl_row, 6    'reading the reference number
                            if ref_nbr = reference Then             'if this is the client we are looking at
                                EMReadScreen type_of_spenddown, 20, mobl_row, 39        'reading the type of spenddown indicated for this client
                                HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt) = trim(type_of_spenddown)   'adding this to the array
                                If type_of_spenddown <> "NO SPENDDOWN" Then             'if there is a spenddown, we will determine the period it applies for
                                    EMReadScreen period, 13, mobl_row, 61
                                    HC_CLIENTS_DETAIL_ARRAY (spd_pd, hc_clt) = period

                                    If HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt) = "WAIVER OBLIGATION" AND HC_CLIENTS_DETAIL_ARRAY (elig_waiv, hc_clt) = "_" Then HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) & " ~ Spenddown type is 'Waiver Obligation' but no waiver is indicated in ELIG."
                                End If
                                Exit Do         'if we found the right member, then we don't need to look any more
                            End if
                            mobl_row = mobl_row + 1     'looking at the next row for the right person
                            'MsgBox "Loop 4 - Reference number: " & ref_nbr
                        Loop Until ref_nbr = "  "       'this is the end of the list
                        PF3
                    Else                                                            'this is for programs other than MA or IMD - typically QMB, SLMB, or QI
                        If left(HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then       'for some clients that don't have an actual name
                            EmReadscreen the_name, 30, 5, 15
                            the_name = trim(the_name)
                            HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt) = the_name
                        End If
                        EMReadScreen pers_type, 2, 6, 56                                'reading the type and standard
                        EMReadScreen pers_std, 1, 6, 64

                        HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = prog          'adding this to the array

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_one, hc_clt) = pers_type
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_one,  hc_clt) = pers_std
                    End If
                    PF3
                End If
            End If

            Do                                              'this is after the first program is listed, there may be a second program
                row = row + 1                               'looking at the next row

                EmReadscreen next_client_ref, 2, row, 3     'reading the reference number and program'
                EmReadscreen next_prog, 4, row, 28

                next_prog = trim(next_prog)
                If next_client_ref <> "  " Then Exit Do     'if the next line has a different reference number listed then no more to read
                If next_prog = "" Then Exit Do              'if there is no program listed on the next line, there is no more to read

                found_elig = FALSE                          'setting this at the beginning of each loop
                EMReadScreen prog, 10, row, 28              'reading the program information from the current row
                EMReadScreen version, 2, row, 58
                EMReadScreen app_indc, 6, row, 68

                prog = trim(prog)
                app_indc = trim(app_indc)

                If prog = "NO REQUEST" OR prog = "NO VERSION" OR prog = "" Then                          'If there is no span known for the member, it will be indicated in this way and we can't see any more
                    HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ HC information does not appear to be in MAXIS. ELIG/HC for this member - " & prog
                    Exit Do
                End If

                If app_indc <> "APP" Then                   'If this has not been approved then we will try to find an approved version
                    if version = "01" Then                  'if we are at version 01 - there are no other ones to check
                        found_elig = TRUE
                        APPROVAL_NEEDED = TRUE
                    Else                                    'if not at version 01 then we will loop through the versions to find the approved one
                        Do
                            EMReadScreen version, 2, row, 58                    'reading this at the beginning of each loop
                            'MsgBox "2 - Version: " & version
                            version = version * 1                               'make it a number
                            prev_verision = version - 1                         'go back one
                            prev_verision = right("00" & prev_verision, 2)      'make it a string

                            EMWriteScreen prev_verision, row, 58                'writing in the previous version and transmitting to pull the information up
                            transmit
                            EMReadScreen app_indc, 6, row, 68                   'determe if it has been approved
                            app_indc = trim(app_indc)
                            If app_indc = "APP" Then
                                found_elig = TRUE
                                Exit Do                                         'leave the loop at this version if it has been approved
                            End If
                            'MsgBox "Loop 6 - prev_verision: " & prev_verision
                        Loop until prev_verision = "01"                         'can't go back any further
                        EMReadScreen version, 2, row, 58
                        EMReadScreen app_indc, 6, row, 68
                        app_indc = trim(app_indc)
                        If version = "01" AND app_indc <> "APP" Then APPROVAL_NEEDED = TRUE
                    End If
                Else
                    found_elig = TRUE
                End If

                If found_elig = TRUE Then                                       'this was set in the code above.
                    EmWriteScreen "X", row, 26                                  'opening BSUM
                    transmit                                                    'we don't need to determine program because a second programs is always medicare savings progs

                    If left(HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then       'finding the correct name if the case is priv but I have access
                        EmReadscreen the_name, 30, 5, 15
                        the_name = trim(the_name)
                        HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt) = the_name
                    End If

                    EMReadScreen pers_type, 2, 6, 56                            'reading the type and standard
                    EMReadScreen pers_std, 1, 6, 64

                    If HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt) <> "" Then      'this adds it to the array after determining WHICH part it belongs in
                        HC_CLIENTS_DETAIL_ARRAY (hc_prog_two,   hc_clt) = prog

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_two, hc_clt) = pers_type
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_two,  hc_clt) = pers_std
                    Else
                        HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = prog

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_one, hc_clt) = pers_type
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_one,  hc_clt) = pers_std
                    End If
                    PF3
                End If
                'MsgBox "Loop 5 - the row: " & row
            Loop until row = 20
        End If
        row = row + 1       'incrementing

        If row = 18 Then    'this is the last line of the HH Members on ELIG
            PF8             'goes to the next page and resets the row
            row = 8

            EMReadScreen is_there_more, 9, 24, 14       'reading for the last page of the list
            If is_there_more = "LAST PAGE" Then
                If client_found = FALSE Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ Member number not found on ELIG/HC."
                Exit Do
            End If
        End If
    Loop until client_found = TRUE
    Call keep_MMIS_passworded_in
Next

'need to get to ground zero
Call back_to_SELF
Call navigate_to_spec_MMIS_region("CTY ELIG STAFF/UPDATE")      'Going to MMIS'
'TODO add functionality to check for MMIS before the full run and use of the timer to check in to MMIS periodically through the run'

'Looping through each of the HC clients while in MMIS
For hc_clt = 0 to UBOUND(HC_CLIENTS_DETAIL_ARRAY, 2)
    STATS_counter = STATS_counter + 1       'incrementing for each client HC reviewed - it is here because this is the part that the timer will cut out on
    PMI_Number = right("00000000" & HC_CLIENTS_DETAIL_ARRAY(clt_pmi, hc_clt), 8)    'making this 8 charactes because MMIS

    If HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) <> "PRIV" Then                  'Can't look at priv case information so we will ignore them
        EmWriteScreen "I", 2, 19                                                    'read only
        EmWriteScreen PMI_Number, 4, 19                                             'enter through the PMI so it isn't case specific
        transmit

        EmWriteScreen "RELG", 1, 8                  'go to RELG where all the elig detail is
        transmit

        relg_row = 6                                'beginning of the list.
        span_found = FALSE                          'setting this for each client loop
        Do
            EmReadscreen relg_prog, 2, relg_row, 10 'reading the prog and elig type information
            EmReadscreen relg_elig, 2, relg_row, 33
            'MsgBox relg_prog & " - " & relg_elig

            'If the program matches and the elig type matches we will read for an end date
            If relg_prog = left(HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt), 2) AND relg_elig = HC_CLIENTS_DETAIL_ARRAY(elig_type_one, hc_clt) Then
                span_found = TRUE           'setting this for later/next loop
                EmReadscreen relg_end_dt, 8, relg_row+1, 36     'this is where the end date is
                'MsgBox "End Date - " & relg_end_dt
                HC_CLIENTS_DETAIL_ARRAY(mmis_end_one, hc_clt) = relg_end_dt     'setting the end date in to the array
                If relg_end_dt <> "99/99/99" Then           '99/99/99 is a no end date - if there is an actual end date, we are going to compare it with today
                    'if the difference is over 0, the end date is before today and it is not active.
                    If DateDiff("d", relg_end_dt, date) > 0 Then HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) = "MMIS SPAN ENDED for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt)
                End If
            ElseIf relg_prog = left(HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt), 2) Then       'sometimes the program matches but the elig type does not - HC is still active in MMIS but wrong
                EmReadscreen relg_end_dt, 8, relg_row+1, 36         'reading the end date
                'if there is no end date or the end date is after today then this span is active
                if relg_end_dt <> "99/99/99" Then
                    difference_between = DateDiff("d", relg_end_dt, date)
                Else
                    difference_between = 1
                End If
                If relg_end_dt = "99/99/99" OR difference_between < 0 Then
                    HC_CLIENTS_DETAIL_ARRAY(mmis_end_one, hc_clt) = relg_end_dt         'adding it to the array and adding a message about the wrong elig type
                    HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) = "MMIS SPAN for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt) & " has the wrong ELIG TYPE"
                    span_found = TRUE
                End If
            End If

            'For MA cases, look for a spenddown inicator in the span to compare it to the spandown indicated in MAXIS Elig
            If relg_prog = "MA" and span_found = TRUE Then
                EmReadscreen spd_indct, 1, relg_row+2, 62
                If HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt) = "NO SPENDDOWN" and spd_indct = "Y" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ No spenddown indicated in MAXIS but MMIS spenddown indicator is Y."
                If HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt) <> "NO SPENDDOWN" and left(HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt), 15) <> "MONTHLY PREMIUM" and spd_indct <> "Y" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ MAXIS ELIG indicates and Spenddown but MMIS span does not."
            End If

            'Once PROG is blank - there are no more spans to review
            If relg_prog = "  " Then Exit Do
            relg_row = relg_row + 4         'next span on RELG'
            If relg_row = 22 Then           'this is the end of RELG and we need to go to a new page
                PF8
                relg_row = 6
            End If
            EmReadscreen end_of_list, 7, 24, 26     'This is the end of the list
            If end_of_list = "NO MORE" Then Exit Do
        Loop until span_found = TRUE
        'If we exited before finding the right Span then an error is added that a span does not exist.
        If span_found = FALSE Then HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) = "No MMIS SPAN for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt)

        EmWriteScreen "RELG", 1, 8      'This takes us back to the top in case we had to PF8 down'
        transmit

        'If there is a second program for this client, we are goind to do it all over again.
        If HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt) <> "" Then
            relg_row = 6                'top of the list of Spans
            span_found = FALSE          'reset this for the next program
            Do
                EmReadscreen relg_prog, 2, relg_row, 10     'reading program and elig type'
                EmReadscreen relg_elig, 2, relg_row, 33
                'MsgBox "2 - " & relg_prog & " - " & relg_elig

                'if both match, getting th end date
                If relg_prog = left(HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt), 2) AND relg_elig = HC_CLIENTS_DETAIL_ARRAY(elig_type_two, hc_clt) Then
                    span_found = TRUE
                    EmReadscreen relg_end_dt, 8, relg_row+1, 36                 'reading the end date
                    'MsgBox "2 - End Date - " & relg_end_dt
                    HC_CLIENTS_DETAIL_ARRAY(mmis_end_two, hc_clt) = relg_end_dt 'setting it to the array
                    If relg_end_dt <> "99/99/99" Then                           'looking to see is the span has ended.
                        If DateDiff("d", relg_end_dt, date) > 0 Then HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) = "MMIS SPAN ENDED for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt)
                    End If
                ElseIf relg_prog = left(HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt), 2) Then       'if only the program matches
                    EmReadscreen relg_end_dt, 8, relg_row+1, 36                 'reading the end date
                    if relg_end_dt <> "99/99/99" Then
                        difference_between = DateDiff("d", relg_end_dt, date)
                    Else
                        difference_between = 1
                    End If
                    If relg_end_dt = "99/99/99" OR difference_between < 0 Then
                        HC_CLIENTS_DETAIL_ARRAY(mmis_end_two, hc_clt) = relg_end_dt         'adding it to the array and adding a message about the wrong elig type
                        HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) = "MMIS SPAN for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt) & " has the wrong ELIG TYPE"
                        span_found = TRUE
                    End If
                End If

                If relg_prog = "  " Then Exit Do            'leaving the loop if we are at the end of the RELG list'
                relg_row = relg_row + 4                     'next span
                If relg_row = 22 Then                       'next page
                    PF8
                    relg_row = 6
                End If
                EmReadscreen end_of_list, 7, 24, 26
                If end_of_list = "NO MORE" Then Exit Do     'last page
            Loop until span_found = TRUE
            'adding a message if no span was found for this program
            If span_found = FALSE Then HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) = "No MMIS SPAN for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt)
        End If

        EmWriteScreen "RKEY", 1, 8  'back to the beginning for the next client/loop'
        transmit
    End If
    'if for some reason no HC programs were in MAXIS to begin with - adding this detail to the message
    If HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt) = "" AND HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt) = "" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ No HC Programs found in MAXIS ELIG."

    'this ends the script run if we have hid the time that was indicated to run for
    If timer > end_time Then
        end_msg = "Success! Script has run for " & stop_time/60/60 & " hours and has finished." & vbNewLine & "Last row from the BOBI reviewd and added: " & HC_CLIENTS_DETAIL_ARRAY (add_xcl, hc_clt) & vbNewLine & end_msg
        Exit For
    End If
Next

'Opening a new Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

col_to_use = 1
'Setting the column headers and defining the column numbers for entry of client information
ObjExcel.Cells(1, col_to_use).Value = "WORKER"
worker_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "CASE NUMBER"
case_numb_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "REF NO"
ref_numb_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "NAME"
name_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "PMI"
pmi_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "1st Prog"
prog_one_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "ELIG TYPE"
elig_one_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "MMIS End Date"
mmis_one_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "2nd PROG"
prog_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "ELIG TYPE"
elig_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "MMIS End Date"
mmis_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "SPDWN ON MOBL"
spdn_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "WAIVER"
waiver_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "ERRORS"
errors_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Rows(1).Font.Bold = TRUE
excel_row = 2

'Adding all client information to a spreadsheet for your viewing pleasure
For hc_clt = 0 to UBound(HC_CLIENTS_DETAIL_ARRAY, 2)
	ObjExcel.Cells(excel_row, worker_col).Value       = HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt)
	ObjExcel.Cells(excel_row, case_numb_col).Value    = HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt)
	ObjExcel.Cells(excel_row, ref_numb_col).Value     = "Memb " & HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt)
	ObjExcel.Cells(excel_row, name_col).Value         = HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt)
	ObjExcel.Cells(excel_row, pmi_col).Value          = HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt)
    ObjExcel.Cells(excel_row, prog_one_col).Value     = HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt)
    If HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = "MA" Then
	    ObjExcel.Cells(excel_row, elig_one_col).Value  = HC_CLIENTS_DETAIL_ARRAY (elig_type_one,   hc_clt) & "-" & HC_CLIENTS_DETAIL_ARRAY(elig_std_one, hc_clt) & " - Method: " & HC_CLIENTS_DETAIL_ARRAY(elig_mthd_one, hc_clt)
	    If HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt) <> "NO SPENDDOWN" Then
            ObjExcel.Cells(excel_row, spdn_col).Value  = HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt) & " for " & HC_CLIENTS_DETAIL_ARRAY(spd_pd, hc_clt)
        Else
            ObjExcel.Cells(excel_row, spdn_col).Value  = HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt)
        End If
    Else
        ObjExcel.Cells(excel_row, elig_one_col).Value  = HC_CLIENTS_DETAIL_ARRAY (elig_type_one,   hc_clt) & "-" & HC_CLIENTS_DETAIL_ARRAY(elig_std_one, hc_clt)
    End If
    ObjExcel.Cells(excel_row, mmis_one_col).Value       = HC_CLIENTS_DETAIL_ARRAY(mmis_end_one, hc_clt)

    ObjExcel.Cells(excel_row, prog_two_col).Value     = HC_CLIENTS_DETAIL_ARRAY (hc_prog_two,   hc_clt)
    If HC_CLIENTS_DETAIL_ARRAY (hc_prog_two,   hc_clt) = "MA" Then
        ObjExcel.Cells(excel_row, elig_two_col).Value  = HC_CLIENTS_DETAIL_ARRAY (elig_type_two,   hc_clt) & "-" & HC_CLIENTS_DETAIL_ARRAY(elig_std_two, hc_clt) & " - Method: " & HC_CLIENTS_DETAIL_ARRAY(elig_mthd_two, hc_clt)
        If HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt) <> "NO SPENDDOWN" Then
            ObjExcel.Cells(excel_row, spdn_col).Value  = HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt) & " for " & HC_CLIENTS_DETAIL_ARRAY(spd_pd, hc_clt)
        Else
            ObjExcel.Cells(excel_row, spdn_col).Value  = HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt)
        End If
    Else
        If HC_CLIENTS_DETAIL_ARRAY(elig_type_two, hc_clt) <> "" Then ObjExcel.Cells(excel_row, elig_two_col).Value  = HC_CLIENTS_DETAIL_ARRAY (elig_type_two,   hc_clt) & "-" & HC_CLIENTS_DETAIL_ARRAY(elig_std_two, hc_clt)
    End If
    ObjExcel.Cells(excel_row, mmis_two_col).Value       = HC_CLIENTS_DETAIL_ARRAY(mmis_end_two, hc_clt)

    If HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) <> "" OR HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) <> "" Then ObjExcel.Rows(excel_row).Interior.ColorIndex = 6

    ObjExcel.Cells(excel_row, waiver_col).Value     = HC_CLIENTS_DETAIL_ARRAY(elig_waiv, hc_clt)
    If HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) <>"" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ " & HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt)
    If HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) <>"" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ " & HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt)

    ObjExcel.Cells(excel_row, errors_col).Value     = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt)
	excel_row = excel_row + 1      'next row
Next

col_to_use = col_to_use + 1     'moving over one extra for script run details.

'Query date/time/runtime info
objExcel.Cells(2, col_to_use).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use+1).Value = now
ObjExcel.Cells(1, col_to_use+1).Font.Bold = FALSE
ObjExcel.Cells(2, col_to_use).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use+1).Value = timer - query_start_time

'Autofitting columns
For col_to_autofit = 1 to col_to_use+1
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure(end_msg)
