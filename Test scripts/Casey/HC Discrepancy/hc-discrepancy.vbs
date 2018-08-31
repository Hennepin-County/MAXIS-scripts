'Required for statistical purposes==========================================================================================
name_of_script = "BULK - SPENDDOWN REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 72                      'manual run time in seconds
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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

'DIALOGS----------------------------------------------------------------------
' BeginDialog find_spenddowns_month_spec_dialog, 0, 0, 221, 180, "Pull REPT data into Excel dialog"
'   EditBox 85, 20, 130, 15, worker_number
'   DropListBox 125, 85, 80, 45, "ALL"+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", revw_month_list
'   CheckBox 5, 105, 150, 10, "Check here to have the script check MMIS", MMIS_checkbox
'   CheckBox 5, 120, 150, 10, "Check here to run this query county-wide.", all_workers_check
'   ButtonGroup ButtonPressed
'     OkButton 110, 160, 50, 15
'     CancelButton 165, 160, 50, 15
'   Text 50, 5, 125, 10, "*** REPT ON MAXIS SPENDDOW ***"
'   Text 5, 25, 65, 10, "Worker(s) to check:"
'   Text 5, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
'   Text 5, 60, 210, 25, "** If a supervisor 'x1 number' is entered, the script will add the 'x1 numbers' of all workers listed in MAXIS under that supervisor number."
'   Text 5, 90, 120, 10, "Only pull cases with next review in:"
'   Text 5, 135, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
' EndDialog

BeginDialog find_spenddowns_month_spec_dialog, 0, 0, 221, 200, "Pull REPT data into Excel dialog"
  EditBox 85, 20, 130, 15, worker_number
  EditBox 5, 120, 210, 15, hc_cases_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 165, 140, 50, 15, "Browse...", select_a_file_button
  'DropListBox 125, 160, 90, 45, "ALL"+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", revw_month_list
  ButtonGroup ButtonPressed
    OkButton 110, 180, 50, 15
    CancelButton 165, 180, 50, 15
  Text 50, 5, 125, 10, "*** REPT ON MAXIS SPENDDOW ***"
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 5, 60, 210, 25, "** If a supervisor 'x1 number' is entered, the script will add the 'x1 numbers' of all workers listed in MAXIS under that supervisor number."
  Text 100, 90, 15, 10, "OR"
  Text 5, 105, 135, 10, "Select an Excel file of MAXIS MA cases:"
  'Text 5, 165, 120, 10, "Only pull cases with next review in:"
EndDialog

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
get_county_code

'Connects to BlueZone
EMConnect ""

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

'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
'Show initial dialog
Do
    Do
        err_msg = ""

    	Dialog find_spenddowns_month_spec_dialog
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
    Next

'If there are no worker numbers entered, then we are going to use a BOBI list of active HC clients
Else

    'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
    call excel_open(hc_cases_excel_file_path, True, True, ObjExcel, objWorkbook)

    excel_row_to_start = "5"    'presetting this before the dialog since BOBI case information starts on row 5

    'This is the dialog to limit the script run as the BOBI is in the tens of thousands
    BeginDialog Dialog1, 0, 0, 176, 140, "Dialog"
      EditBox 25, 55, 30, 15, stop_time
      EditBox 65, 100, 30, 15, excel_row_to_start
      EditBox 65, 120, 30, 15, excel_row_to_end
      ButtonGroup ButtonPressed
        OkButton 115, 120, 50, 15
      Text 5, 10, 165, 10, "This run of the script will review and help process: "
      Text 5, 20, 165, 10, process_option
      Text 10, 35, 140, 20, "To time limit the run of the script enter the numeber of hours to run the script:"
      Text 65, 60, 50, 10, "Hours"
      Text 10, 80, 145, 20, "The run can be limited by indicating which rows of the Excel file to review/process:"
      Text 15, 105, 50, 10, "Excel to start"
      Text 15, 125, 45, 10, "Excel to end"
    EndDialog

    'showing the dialog
    Do
        Do
            err_msg = ""
            dialog Dialog1

            If trim(stop_time) <> "" AND IsNumeric(stop_time) = FALSE Then err_msg = err_msg & vbNewLine & "- Number of hours should be a number."
            If trim(excel_row_to_start) <> "" AND IsNumeric(excel_row_to_start) = FALSE Then err_msg = err_msg & vbNewLine & "- Start row of Excel should be a number."
            If trim(excel_row_to_end) <> "" AND IsNumeric(excel_row_to_end) = FALSE Then err_msg = err_msg & vbNewLine & "- End row of Excel should be a number."

            If err_msg <> "" Then MsgBox "** Please Resolve the Following to Continue:" & vbNew & err_msg

        Loop until err_msg = ""
        call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

    'setting these to numbers
    excel_row = excel_row_to_start * 1
    excel_row_to_end = excel_row_to_end * 1
    hc_clt = 0      'setting the beginning of the array

    'TODO add time handling base it on average time per client

    Do
        'ObjExcel.Cells(). Value
        ReDim Preserve HC_CLIENTS_DETAIL_ARRAY (add_xcl, hc_clt)

        HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) = ObjExcel.Cells(excel_row, 3). Value
        HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt) = ObjExcel.Cells(excel_row, 2). Value
        'HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt) = ObjExcel.Cells(excel_row, ). Value
        HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt) = ObjExcel.Cells(excel_row, 8). Value
        HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) = right(ObjExcel.Cells(excel_row, 7). Value, 2)
        HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt) = ObjExcel.Cells(excel_row, 6). Value

        ' MsgBox "Worker: " & HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) & vbNewLine &_
        '        "Case: " & HC_CLIENTS_DETAIL_ARRAY (case_num,   hc_clt) & vbNewLine &_
        '        "Client: " & HC_CLIENTS_DETAIL_ARRAY (clt_name,   hc_clt) & vbNewLine &_
        '        "Ref Number: " & HC_CLIENTS_DETAIL_ARRAY (ref_numb,   hc_clt) & vbNewLine &_
        '        "PMI: " & HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt)
        excel_row = excel_row + 1
        hc_clt = hc_clt + 1
        next_case_number = ObjExcel.Cells(excel_row, 2). Value
        next_case_number = trim(next_case_number)
        If excel_row = excel_row_to_end Then Exit Do
    Loop until next_case_number = ""

    end_msg = "Success! Client HC Eligibility and MMIS coding for row " & excel_row_to_start & " to " & excel_row_to_end & " have been added to the spreadsheet."

    ObjExcel.Quit
    Set ObjExcel = Nothing
End If

For hc_clt = 0 to UBOUND(HC_CLIENTS_DETAIL_ARRAY, 2)
    back_to_SELF
    MAXIS_case_number = HC_CLIENTS_DETAIL_ARRAY(case_num, hc_clt)		'defining case number for functions to use
    CLIENT_reference_number = HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt)
    Call navigate_to_MAXIS_screen ("ELIG", "HC")						'Goes to ELIG HC
    APPROVAL_NEEDED = FALSE
    found_elig = FALSE
    client_found = FALSE
    row = 8
    Do
        EMReadScreen check_for_priv, 10, 24, 14
        If check_for_priv = "PRIVILEGED" Then
            HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = "PRIV"
            Exit Do
        End If
        'TODO add handling for more than 1 page of elig results (BOBI has one that is 168XXX or 169XXX)'
        EMReadscreen elig_clt, 2, row, 3
        EmReadscreen prog_exists, 1, row, 10
        If elig_clt = CLIENT_reference_number Then
            'MsgBox "Elig Clt: " & elig_clt & vbNewLine & "Ref Number: " &CLIENT_reference_number
            client_found = TRUE
            EMReadScreen prog, 10, row, 28
            EMReadScreen version, 2, row, 58
            EMReadScreen app_indc, 6, row, 68

            prog = trim(prog)
            app_indc = trim(app_indc)

            If prog = "NO REQUEST" Then Exit DO
            If prog = "NO VERSION" Then Exit DO
            If prog = "" Then Exit Do

            If app_indc <> "APP" Then
                if version = "01" Then
                    found_elig = TRUE
                    APPROVAL_NEEDED = TRUE
                Else
                    Do
                        EMReadScreen version, 2, row, 58
                        'MsgBox "1 - Version: " & version
                        version = version * 1
                        prev_verision = version - 1
                        prev_verision = right("00" & prev_verision, 2)

                        EMWriteScreen prev_verision, row, 58
                        transmit
                        EMReadScreen app_indc, 6, row, 68
                        app_indc = trim(app_indc)
                        If app_indc = "APP" Then
                            found_elig = TRUE
                            Exit Do
                        End If
                        'MsgBox "Loop 2 - prev_verision: " & prev_verision
                    Loop until prev_verision = "01"
                    EMReadScreen version, 2, row, 58
                    EMReadScreen app_indc, 6, row, 68
                    app_indc = trim(app_indc)
                    If version = "01" AND app_indc <> "APP" Then APPROVAL_NEEDED = TRUE
                End If
            Else
                found_elig = TRUE
            End If

            If found_elig = TRUE Then
                EMReadScreen prog, 10, row, 28
                EMReadScreen result, 7, row, 41
                EMReadScreen hc_status, 7, row, 50

                prog = trim(prog)
                result = trim(result)
                hc_status = trim(hc_status)

                If result = "ELIG" AND hc_status = "ACTIVE" Then
                    HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = prog

                    EmWriteScreen "X", row, 26
                    transmit

                    If prog = "MA" or prog = "IMD" Then
                        If left(HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then
                            EmReadscreen the_name, 30, 5, 20
                            the_name = trim(the_name)
                            HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt) = the_name
                        End If
                        mo_col = 19
                        yr_col = 22
                        Do
                            EMReadScreen bsum_mo, 2, 6, mo_col
                            EMReadScreen bsum_yr, 2, 6, yr_col

                            If bsum_mo = MAXIS_footer_month and bsum_yr = MAXIS_footer_year Then Exit Do
                            mo_col = mo_col + 11
                            yr_col = yr_col + 11
                            'MsgBox "Loop 3 - month col: " & mo_col
                        Loop until mo_col = 74

                        EMReadScreen cname, 35, 5, 20
                        EMReadScreen reference, 2, 5, 16

                        EMReadScreen prog, 4, 11, mo_col
                        EMReadScreen pers_type, 2, 12, mo_col-2
                        EMReadScreen pers_std, 1, 12, yr_col
                        EMReadScreen pers_mthd, 1, 13, yr_col-1
                        EMReadScreen pers_waiv, 1, 14, yr_col-1

                        If prog = "    " Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ HC ELIG Budget may need approval or budget needs to be aligned."

                        If pers_type = "__" Then
                            EMReadScreen cur_mo_test, 6, 7, mo_col
                            cur_mo_test = trim(cur_mo_test)
                            pers_type = cur_mo_test
                            pers_std = ""
                            pers_mthd = ""
                        End If

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_one, hc_clt) = pers_type
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_one,  hc_clt) = pers_std
                        HC_CLIENTS_DETAIL_ARRAY (elig_mthd_one, hc_clt) = pers_mthd
                        HC_CLIENTS_DETAIL_ARRAY (elig_waiv, hc_clt) = pers_waiv

                        If APPROVAL_NEEDED = TRUE THen HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) & " ~ SPAN Needs Approval"

                        EMWriteScreen "X", 18, 3        'Going in to MOBL
                        transmit

                        mobl_row = 6
                        Do
                            EMReadScreen ref_nbr, 2, mobl_row, 6
                            if ref_nbr = reference Then
                                EMReadScreen type_of_spenddown, 20, mobl_row, 39
                                HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt) = trim(type_of_spenddown)
                                If type_of_spenddown <> "NO SPENDDOWN" Then
                                    EMReadScreen period, 13, mobl_row, 61
                                    HC_CLIENTS_DETAIL_ARRAY (spd_pd, hc_clt) = period

                                    If HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt) = "WAIVER OBLIGATION" AND HC_CLIENTS_DETAIL_ARRAY (elig_waiv, hc_clt) = "_" Then HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) & " ~ Spenddown type is 'Waiver Obligation' but no waiver is indicated in ELIG."
                                End If
                                Exit Do
                            End if
                            mobl_row = mobl_row + 1
                            'MsgBox "Loop 4 - Reference number: " & ref_nbr
                        Loop Until ref_nbr = "  "
                        PF3
                    Else
                        If left(HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then
                            EmReadscreen the_name, 30, 5, 15
                            the_name = trim(the_name)
                            HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt) = the_name
                        End If
                        EMReadScreen pers_type, 2, 6, 56
                        EMReadScreen pers_std, 1, 6, 64

                        HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = prog

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_one, hc_clt) = pers_type
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_one,  hc_clt) = pers_std
                    End If
                    PF3
                End If
            End If

            Do
                row = row + 1

                EmReadscreen next_client_ref, 2, row, 3
                EmReadscreen next_prog, 4, row, 28

                next_prog = trim(next_prog)
                If next_client_ref <> "  " Then Exit Do
                If next_prog = "" Then Exit Do

                found_elig = FALSE
                EMReadScreen prog, 10, row, 28
                EMReadScreen version, 2, row, 58
                EMReadScreen app_indc, 6, row, 68

                prog = trim(prog)
                app_indc = trim(app_indc)

                If prog = "NO REQUEST" Then Exit DO
                If prog = "NO VERSION" Then Exit DO
                If prog = "" Then Exit Do

                If app_indc <> "APP" Then
                    if version = "01" Then
                        found_elig = TRUE
                        APPROVAL_NEEDED = TRUE
                    Else
                        Do
                            EMReadScreen version, 2, row, 58
                            'MsgBox "2 - Version: " & version
                            version = version * 1
                            prev_verision = version - 1
                            prev_verision = right("00" & prev_verision, 2)

                            EMWriteScreen prev_verision, row, 58
                            transmit
                            EMReadScreen app_indc, 6, row, 68
                            app_indc = trim(app_indc)
                            If app_indc = "APP" Then
                                found_elig = TRUE
                                Exit Do
                            End If
                            'MsgBox "Loop 6 - prev_verision: " & prev_verision
                        Loop until prev_verision = "01"
                        EMReadScreen version, 2, row, 58
                        EMReadScreen app_indc, 6, row, 68
                        app_indc = trim(app_indc)
                        If version = "01" AND app_indc <> "APP" Then APPROVAL_NEEDED = TRUE
                    End If
                Else
                    found_elig = TRUE
                End If

                If found_elig = TRUE Then
                    EmWriteScreen "X", row, 26
                    transmit

                    If left(HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then
                        EmReadscreen the_name, 30, 5, 15
                        the_name = trim(the_name)
                        HC_CLIENTS_DETAIL_ARRAY (clt_name, hc_clt) = the_name
                    End If

                    EMReadScreen pers_type, 2, 6, 56
                    EMReadScreen pers_std, 1, 6, 64

                    If HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt) <> "" Then
                        HC_CLIENTS_DETAIL_ARRAY (hc_prog_two,   hc_clt) = prog

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_two, hc_clt) = pers_type
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_two,  hc_clt) = pers_std
                    End If
                    PF3
                End If
                'MsgBox "Loop 5 - the row: " & row
            Loop until row = 20
        End If
        row = row + 1
        'MsgBox "Loop 1 - client found: " & client_found
    Loop until client_found = TRUE

Next

Call back_to_SELF
Call navigate_to_spec_MMIS_region("CTY ELIG STAFF/UPDATE")

For hc_clt = 0 to UBOUND(HC_CLIENTS_DETAIL_ARRAY, 2)
    PMI_Number = right("00000000" & HC_CLIENTS_DETAIL_ARRAY(clt_pmi, hc_clt), 8)

    If HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) <> "PRIV" Then
        EmWriteScreen "I", 2, 19
        EmWriteScreen PMI_Number, 4, 19
        transmit

        EmWriteScreen "RELG", 1, 8
        transmit

        relg_row = 6
        span_found = FALSE
        Do
            EmReadscreen relg_prog, 2, relg_row, 10
            EmReadscreen relg_elig, 2, relg_row, 33
            'MsgBox relg_prog & " - " & relg_elig

            If relg_prog = left(HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt), 2) AND relg_elig = HC_CLIENTS_DETAIL_ARRAY(elig_type_one, hc_clt) Then
                span_found = TRUE
                EmReadscreen relg_end_dt, 8, relg_row+1, 36
                'MsgBox "End Date - " & relg_end_dt
                HC_CLIENTS_DETAIL_ARRAY(mmis_end_one, hc_clt) = relg_end_dt
                If relg_end_dt <> "99/99/99" Then
                    If DateDiff("d", relg_end_dt, date) > 0 Then HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) = "MMIS SPAN ENDED for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt)
                End If
            ElseIf relg_prog = left(HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt), 2) Then
                EmReadscreen relg_end_dt, 8, relg_row+1, 36
                If relg_end_dt = "99/99/99" Then
                    HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) = "MMIS SPAN for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt) & " has the wrong ELIG TYPE"
                    span_found = TRUE
                End If
            End If

            If relg_prog = "MA" and span_found = TRUE Then
                EmReadscreen spd_indct, 1, relg_row+2, 62
                If HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt) = "NO SPENDDOWN" and spd_indct = "Y" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ No spenddown indicated in MAXIS but MMIS spenddown indicator is Y."
                If HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt) <> "NO SPENDDOWN" and left(HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt), 15) <> "MONTHLY PREMIUMN" and spd_indct <> "Y" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ MAXIS ELIG indicates and Spenddown but MMIS span does not."
            End If

            If relg_prog = "  " Then Exit Do
            relg_row = relg_row + 4
            If relg_row = 22 Then
                PF8
                relg_row = 6
            End If
            EmReadscreen end_of_list, 7, 24, 26
            If end_of_list = "NO MORE" Then Exit Do
        Loop until span_found = TRUE
        If span_found = FALSE Then HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) = "No MMIS SPAN for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt)

        EmWriteScreen "RELG", 1, 8
        transmit

        If HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt) <> "" Then
            relg_row = 6
            span_found = FALSE
            Do
                EmReadscreen relg_prog, 2, relg_row, 10
                EmReadscreen relg_elig, 2, relg_row, 33
                'MsgBox "2 - " & relg_prog & " - " & relg_elig

                If relg_prog = left(HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt), 2) AND relg_elig = HC_CLIENTS_DETAIL_ARRAY(elig_type_two, hc_clt) Then
                    span_found = TRUE
                    EmReadscreen relg_end_dt, 8, relg_row+1, 36
                    'MsgBox "2 - End Date - " & relg_end_dt
                    HC_CLIENTS_DETAIL_ARRAY(mmis_end_two, hc_clt) = relg_end_dt
                    If relg_end_dt <> "99/99/99" Then
                        If DateDiff("d", relg_end_dt, date) > 0 Then HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) = "MMIS SPAN ENDED for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt)
                    End If
                ElseIf relg_prog = left(HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt), 2) Then
                    EmReadscreen relg_end_dt, 8, relg_row+1, 36
                    If relg_end_dt = "99/99/99" Then
                        HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) = "MMIS SPAN for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt) & " has the wrong ELIG TYPE"
                        span_found = TRUE
                    End If
                End If

                If relg_prog = "  " Then Exit Do
                relg_row = relg_row + 4
                If relg_row = 22 Then
                    PF8
                    relg_row = 6
                End If
                EmReadscreen end_of_list, 7, 24, 26
                If end_of_list = "NO MORE" Then Exit Do
            Loop until span_found = TRUE
            If span_found = FALSE Then HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) = "No MMIS SPAN for " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt)

        End If

        EmWriteScreen "RKEY", 1, 8
        transmit
    End If

    If HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt) = "" AND HC_CLIENTS_DETAIL_ARRAY(hc_prog_two, hc_clt) = "" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ No HC Programs found in MAXIS ELIG."

Next


'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

col_to_use = 1
'Setting the first 4 col as worker, case number, name, and APPL date
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

' ObjExcel.Cells(1, col_to_use).Value = "NEXT REVW DATE"
' revw_date_col = col_to_use
' col_to_use = col_to_use + 1

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

If MMIS_checkbox = checked Then
	ObjExcel.Cells(1, col_to_use).Value = "MMIS SPDWN"
    mmis_spdn_col = col_to_use
    col_to_use = col_to_use + 1
End If

ObjExcel.Cells(1, col_to_use).Value = "ERRORS"
errors_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Rows(1).Font.Bold = TRUE
excel_row = 2

'Adding all client information to a spreadsheet for your viewing pleasure
For hc_clt = 0 to UBound(HC_CLIENTS_DETAIL_ARRAY, 2)
	'If HC_CLIENTS_DETAIL_ARRAY(add_xcl, hc_clt) = TRUE Then
		ObjExcel.Cells(excel_row, worker_col).Value       = HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt)
		ObjExcel.Cells(excel_row, case_numb_col).Value    = HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt)
		ObjExcel.Cells(excel_row, ref_numb_col).Value     = "Memb " & HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt)
		ObjExcel.Cells(excel_row, name_col).Value         = HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt)
		ObjExcel.Cells(excel_row, pmi_col).Value          = HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt)
		'ObjExcel.Cells(excel_row, revw_date_col).Value    = HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt)
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

        'ObjExcel.Cells(excel_row, 9).Value  = HC_CLIENTS_DETAIL_ARRAY (hc_excess, hc_clt)
		'ObjExcel.Cells(excel_row, 10).Value = HC_CLIENTS_DETAIL_ARRAY (mmis_spdn, hc_clt)
        ObjExcel.Cells(excel_row, waiver_col).Value     = HC_CLIENTS_DETAIL_ARRAY(elig_waiv, hc_clt)
        'If left(HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt), 3) = " ~ " Then HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) = right(HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt), len(HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_cl)-3))
        If HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt) <>"" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ " & HC_CLIENTS_DETAIL_ARRAY(disc_one, hc_clt)
        If HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt) <>"" Then HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt) & " ~ " & HC_CLIENTS_DETAIL_ARRAY(disc_two, hc_clt)

        ObjExcel.Cells(excel_row, errors_col).Value     = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt)
		excel_row = excel_row + 1
	'End If
Next

col_to_use = col_to_use + 1

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
