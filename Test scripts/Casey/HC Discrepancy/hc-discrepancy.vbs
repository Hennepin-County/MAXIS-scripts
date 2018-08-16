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

function navigate_to_MMIS_region(group_security_selection)
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
	If MMIS_menu <> "GROUP SECURITY SELECTION" Then
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
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
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
  DropListBox 125, 160, 90, 45, "ALL"+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", revw_month_list
  ButtonGroup ButtonPressed
    OkButton 110, 180, 50, 15
    CancelButton 165, 180, 50, 15
  Text 50, 5, 125, 10, "*** REPT ON MAXIS SPENDDOW ***"
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 5, 60, 210, 25, "** If a supervisor 'x1 number' is entered, the script will add the 'x1 numbers' of all workers listed in MAXIS under that supervisor number."
  Text 100, 90, 15, 10, "OR"
  Text 5, 105, 135, 10, "Select an Excel file of MAXIS MA cases:"
  Text 5, 165, 120, 10, "Only pull cases with next review in:"
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
Const elig_type_one = 7
Const elig_std_one  = 8
Const elig_mthd_one = 9
Const elig_waiv     = 10
Const mobl_spdn     = 11
Const spd_pd        = 12
Const hc_excess     = 13
Const hc_prog_two   = 14
Const elig_type_two = 15
Const elig_std_two  = 16
Const elig_mthd_two = 17
Const mmis_spdn     = 18
Const error_notes   = 19
Const add_xcl       = 20

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
            If left(err_msg, 4) = "LOOP" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
        End If
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false


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

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
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

    hc_clt = 0

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


    				MAXIS_row = MAXIS_row + 1
    			Loop until MAXIS_row = 19
    			PF8
    		Loop until last_page_check = "THIS IS THE LAST PAGE"
    	End if
    next

    hc_clt = 0

    'The script will now look in each case at MOBL to identify clients that have spenddown listed on MOBL'\
    For hc_case = 0 to UBound(HC_CASES_ARRAY, 2)
        back_to_SELF
    	MAXIS_case_number = HC_CASES_ARRAY(case_num, hc_case)		'defining case number for functions to use
    	Call navigate_to_MAXIS_screen ("ELIG", "HC")						'Goes to ELIG HC
        row = 8
        Do
        'For row = 8 to 20
            found_elig = FALSE
            clt_hc_active = FALSE
            APPROVAL_NEEDED = FALSE
            Do
                EMReadScreen prog, 10, row, 28
                EMReadScreen result, 7, row, 41
                EMReadScreen hc_status, 7, row, 50
                EMReadScreen version, 2, row, 58
                EMReadScreen app_indc, 6, row, 68

                'MsgBox "Row - " & row & vbNewLine & "Version - " & version
                prog = trim(prog)
                result = trim(result)
                hc_status = trim(hc_status)
                app_indc = trim(app_indc)

                If prog = "NO REQUEST" Then Exit DO
                If prog = "NO VERSION" Then Exit DO
                If prog = "" Then Exit Do

                If app_indc <> "APP" Then
                    if version = "01" Then
                        found_elig = TRUE
                        APPROVAL_NEEDED = TRUE
                    Else
                        version = version * 1
                        prev_verision = version - 1
                        prev_verision = right("00" & prev_verision, 2)

                        EMWriteScreen prev_verision, row, 58
                        transmit
                    End If
                Else
                    found_elig = TRUE
                    If result = "ELIG" Then clt_hc_active = TRUE
                End If

            Loop until found_elig = TRUE



            If clt_hc_active = TRUE Then

                second_elig = FALSE                         'finding if there is another HC program open.
                EMReadScreen next_member, 16, row+1, 7
                EMReadScreen next_prog, 10, row+1, 28
                next_member = trim(next_member)
                next_prog = trim(next_prog)

                If next_member = "" AND next_prog <> "" Then
                    EMReadScreen app_indc, 6, row+1, 68
                    EMReadScreen result, 7, row+1, 41
                    EMReadScreen version, 2, row+1, 58
                    app_indc = trim(app_indc)
                    result = trim(result)

                    If app_indc = "APP" Then
                        If result = "ELIG" Then
                            add_row = 1
                            second_elig = TRUE
                        End If
                    Else
                        If version = "01"Then
                            If result = "ELIG" Then
                                APPROVAL_NEEDED = TRUE
                                add_row = 1
                            End If
                        Else
                            version = version * 1
                            For hc_vers = 1 to version-1
                                prev_verision = version - hc_vers
                                prev_verision = right("00" & prev_verision, 2)

                                EMWriteScreen prev_verision, row+1, 58
                                transmit

                                EMReadScreen this_app_indc, 6, row+1, 68
                                EMReadScreen this_result, 7, row+1, 41

                                If this_app_indc = "APP" Then
                                    If this_result = "ELIG" Then add_row = 1
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If

                EMReadScreen next_member, 16, row+2, 7          'Looking if there is a thrid row for a single member'
                EMReadScreen next_prog, 10, row+2, 28
                next_member = trim(next_member)
                next_prog = trim(next_prog)

                If next_member = "" AND next_prog <> "" Then
                    EMReadScreen app_indc, 6, row+2, 68
                    EMReadScreen result, 7, row+2, 41
                    EMReadScreen version, 2, row+2, 58
                    app_indc = trim(app_indc)
                    result = trim(result)

                    If app_indc = "APP" Then
                        If result = "ELIG" Then
                            add_row = 2
                            second_elig = TRUE
                        End If
                    Else
                        If version = "01"Then
                            If result = "ELIG" Then
                                APPROVAL_NEEDED = TRUE
                                add_row = 2
                            End If
                        Else
                            version = version * 1
                            For hc_vers = 1 to version-1
                                prev_verision = version - hc_vers
                                prev_verision = right("00" & prev_verision, 2)

                                EMWriteScreen prev_verision, row+2, 58
                                transmit

                                EMReadScreen this_app_indc, 6, row+2, 68
                                EMReadScreen this_result, 7, row+2, 41

                                If this_app_indc = "APP" Then
                                    If this_result = "ELIG" Then add_row = 2
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If

                If add_row <> "" Then second_elig = TRUE

                EMWriteScreen "X", row, 26		'Goes into the HC ELIG - BSUM
                transmit
                If prog = "MA" Then
                    mo_col = 19
                    yr_col = 22
                    Do
                        EMReadScreen bsum_mo, 2, 6, mo_col
                        EMReadScreen bsum_yr, 2, 6, yr_col

                        If bsum_mo = MAXIS_footer_month and bsum_yr = MAXIS_footer_year Then Exit Do
                        mo_col = mo_col + 11
                        yr_col = yr_col + 11
                    Loop until mo_col = 74

                    EMReadScreen cname, 35, 5, 20
                    EMReadScreen reference, 2, 5, 16

                    EMReadScreen prog, 4, 11, mo_col
                    EMReadScreen pers_type, 2, 12, mo_col-2
                    EMReadScreen pers_std, 1, 12, yr_col
                    EMReadScreen pers_mthd, 1, 13, yr_col-1
                    EMReadScreen pers_waiv, 1, 14, yr_col-1

                    If pers_type = "__" Then
                        EMReadScreen cur_mo_test, 6, 7, mo_col
                        cur_mo_test = trim(cur_mo_test)
                        pers_type = cur_mo_test
                        pers_std = ""
                        pers_mthd = ""
                    End If

                    ReDim Preserve HC_CLIENTS_DETAIL_ARRAY (add_xcl, hc_clt)			'Adding any client with a spenddown to a new array

                    HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) = HC_CASES_ARRAY(wrk_num, hc_case)
                    HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt) = MAXIS_case_number
                    HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt) = replace(HC_CASES_ARRAY(next_revw, hc_case), " ", "/")
                    HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt) = cname
                    HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) = reference
                    HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = trim(prog)

                    HC_CLIENTS_DETAIL_ARRAY (elig_type_one, hc_clt) = pers_type
                    HC_CLIENTS_DETAIL_ARRAY (elig_std_one,  hc_clt) = pers_std
                    HC_CLIENTS_DETAIL_ARRAY (elig_mthd_one, hc_clt) = pers_mthd
                    HC_CLIENTS_DETAIL_ARRAY (elig_waiv, hc_clt) = pers_waiv

                    If APPROVAL_NEEDED = TRUE THen HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) = "SPAN Needs Approval"

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
                            End If
                            Exit Do
                        End if
                        mobl_row = mobl_row + 1
                    Loop Until ref_nbr = "  "
                    PF3
                Else
                    EMReadScreen cname, 35, 5, 15
                    EMReadScreen reference, 2, 5, 11

                    EMReadScreen pers_type, 2, 6, 56
                    EMReadScreen pers_std, 1, 6, 64

                    ReDim Preserve HC_CLIENTS_DETAIL_ARRAY (add_xcl, hc_clt)			'Adding any client with a spenddown to a new array

                    HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) = HC_CASES_ARRAY(wrk_num, hc_case)
                    HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt) = MAXIS_case_number
                    HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt) = replace(HC_CASES_ARRAY(next_revw, hc_case), " ", "/")
                    HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt) = trim(cname)
                    HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) = reference
                    HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = prog

                    HC_CLIENTS_DETAIL_ARRAY (elig_type_one, hc_clt) = pers_type
                    HC_CLIENTS_DETAIL_ARRAY (elig_std_one,  hc_clt) = pers_std
                End If
                PF3

                If second_elig = TRUE Then
                    row = row + add_row
                    EMReadScreen prog, 10, row, 28
                    prog = trim(prog)
                    EMWriteScreen "X", row, 26		'Goes into the HC ELIG - BSUM
                    transmit
                    If prog = "MA" Then
                        mo_col = 19
                        yr_col = 22
                        Do
                            EMReadScreen bsum_mo, 2, 6, mo_col
                            EMReadScreen bsum_yr, 2, 6, yr_col

                            If bsum_mo = MAXIS_footer_month and bsum_yr = MAXIS_footer_year Then Exit Do
                            mo_col = mo_col + 11
                            yr_col = yr_col + 11
                        Loop until mo_col = 74

                        EMReadScreen prog, 4, 11, mo_col
                        EMReadScreen pers_type, 2, 12, mo_col-2
                        EMReadScreen pers_std, 1, 12, yr_col
                        EMReadScreen pers_mthd, 1, 13, yr_col-1
                        EMReadScreen pers_waiv, 1, 14, yr_col-1

                        If pers_type = "__" Then
                            EMReadScreen cur_mo_test, 6, 7, mo_col
                            cur_mo_test = trim(cur_mo_test)
                            pers_type = cur_mo_test
                            pers_std = ""
                            pers_mthd = ""
                        End If

                        HC_CLIENTS_DETAIL_ARRAY (hc_prog_two,   hc_clt) = trim(prog)

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_two, hc_clt) = pers_type
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_two,  hc_clt) = pers_std
                        HC_CLIENTS_DETAIL_ARRAY (elig_mthd_two, hc_clt) = pers_mthd

                        If APPROVAL_NEEDED = TRUE THen HC_CLIENTS_DETAIL_ARRAY (error_notes, hc_clt) = "SPAN Needs Approval"

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
                                End If
                                Exit Do
                            End if
                            mobl_row = mobl_row + 1
                        Loop Until ref_nbr = "  "
                        PF3
                    Else
                        EMReadScreen cname, 35, 5, 15
                        EMReadScreen reference, 2, 5, 11

                        EMReadScreen pers_type, 2, 6, 56
                        EMReadScreen pers_std, 1, 6, 64

                        HC_CLIENTS_DETAIL_ARRAY (hc_prog_two,   hc_clt) = prog

                        HC_CLIENTS_DETAIL_ARRAY (elig_type_two, hc_clt) = pers_type
                        HC_CLIENTS_DETAIL_ARRAY (elig_std_two,  hc_clt) = pers_std
                    End If
                    PF3
                End If
                HC_CLIENTS_DETAIL_ARRAY(add_xcl, hc_clt) = TRUE
                'MsgBox "Case Number: " & HC_CLIENTS_DETAIL_ARRAY(case_num, hc_clt) & vbNewLine & "HC - " & HC_CLIENTS_DETAIL_ARRAY(hc_prog_one, hc_clt)
                hc_clt = hc_clt + 1
            End If
            row = row + 1
            Call navigate_to_MAXIS_screen("ELIG", "HC")
        'Next
        Loop until row = 20
    Next
Else

    'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
    call excel_open(hc_cases_excel_file_path, True, True, ObjExcel, objWorkbook)

    excel_row = 5
    hc_clt = 0

    Do
        ObjExcel.Cells(). Value
        ReDim Preserve HC_CLIENTS_DETAIL_ARRAY (add_xcl, hc_clt)

        HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) = ObjExcel.Cells(excel_row, 3). Value
        HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt) = ObjExcel.Cells(excel_row, 2). Value
        'HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt) = ObjExcel.Cells(excel_row, ). Value
        HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt) = ObjExcel.Cells(excel_row, 8). Value
        HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) = ObjExcel.Cells(excel_row, 7). Value
        HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt) = ObjExcel.Cells(excel_row, 6). Value

        'QUESTION - do we need review date???'

        back_to_SELF
        MAXIS_case_number = HC_CLIENTS_DETAIL_ARRAY(case_num, hc_case)		'defining case number for functions to use
        CLIENT_reference_number = HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt)
        Call navigate_to_MAXIS_screen ("ELIG", "HC")						'Goes to ELIG HC
        row = 8
        Do
            EMReadScreen prog, 10, row, 28
            EMReadScreen result, 7, row, 41
            EMReadScreen hc_status, 7, row, 50
            EMReadScreen version, 2, row, 58
            EMReadScreen app_indc, 6, row, 68

            'MsgBox "Row - " & row & vbNewLine & "Version - " & version
            prog = trim(prog)
            result = trim(result)
            hc_status = trim(hc_status)
            app_indc = trim(app_indc)

            If prog = "NO REQUEST" Then Exit DO
            If prog = "NO VERSION" Then Exit DO
            If prog = "" Then Exit Do

            If app_indc <> "APP" Then
                if version = "01" Then
                    found_elig = TRUE
                    APPROVAL_NEEDED = TRUE
                Else
                    version = version * 1
                    prev_verision = version - 1
                    prev_verision = right("00" & prev_verision, 2)

                    EMWriteScreen prev_verision, row, 58
                    transmit
                End If
            Else
                found_elig = TRUE
                If result = "ELIG" Then clt_hc_active = TRUE
            End If

        Loop until found_elig = TRUE




        HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) =
        HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) =
        HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) =

        HC_CLIENTS_DETAIL_ARRAY (hc_prog_one,   hc_clt) = trim(prog)

        HC_CLIENTS_DETAIL_ARRAY (elig_type_one, hc_clt) = pers_type
        HC_CLIENTS_DETAIL_ARRAY (elig_std_one,  hc_clt) = pers_std
        HC_CLIENTS_DETAIL_ARRAY (elig_mthd_one, hc_clt) = pers_mthd
        HC_CLIENTS_DETAIL_ARRAY (elig_waiv, hc_clt) = pers_waiv


        Const next_revw     = 2

        Const hc_prog_one   = 6
        Const elig_type_one = 7
        Const elig_std_one  = 8
        Const elig_mthd_one = 9
        Const elig_waiv     = 10
        Const mobl_spdn     = 11
        Const spd_pd        = 12
        Const hc_excess     = 13
        Const hc_prog_two   = 14
        Const elig_type_two = 15
        Const elig_std_two  = 16
        Const elig_mthd_two = 17
        Const mmis_spdn     = 18
        Const error_notes   = 19
        Const add_xcl       = 20

        excel_row = excel_row + 1
        next_case_number = ObjExcel.Cells(excel_row, 2). Value
        next_case_number = trim(next_case_number)
    Loop until next_case_number = ""
End if

'Setting the variable for what's to come
excel_row = 2
all_case_numbers_array = "*"



    ' 'THIS IS TE SPENDDOWN REPORT CODE'
    ' row = 8
	' Do										'Looks at each row in HC Elig to find the first MA span
	' 	EMReadScreen prog, 2, row, 28
	' 	If prog = "MA" Then
	' 		EMWriteScreen "X", row, 26		'Goes into it
	' 		transmit
	' 		EMReadScreen panel_check, 4, 3, 57
	' 		If panel_check = "BSUM" Then
	' 			Exit Do
	' 		Else
	' 			Transmit
	' 		End If
	' 	End if
	' 	row = row + 1
	' Loop until row = 20
	' If row <> 20 Then 						'Once in the span, opens MOBL
	' 	EMWriteScreen "X", 18, 3
	' 	transmit
	' 	Do
	' 		EMReadScreen MOBL_check, 4, 3, 49
	' 		If MOBL_check <> "MOBL" Then
	' 			row = row + 1
	' 			PF3
	' 			EMReadScreen prog, 2, row, 28
	' 			If prog = "MA" Then
	' 				EMWriteScreen "X", row, 26		'Goes into it
	' 				transmit
	' 				EMWriteScreen "X", 18, 3
	' 				transmit
	' 			End if
	' 		End If
	' 	Loop until row = 20 OR MOBL_check = "MOBL"
	' 	row = 6
	' 	Do									'reads each line on MOBL and saves the clt information for any client that has a spenddown indicated on MOBL
	' 		EMReadScreen spd_type, 20, row, 39
	' 		spd_type = trim(spd_type)
	' 		If spd_type = "" Then Exit Do			'Leaves the do loop once a blank line is found
	' 		If spd_type <> "NO SPENDDOWN" Then 		'Anything other than this indicates MAXIS thinks there is a spenddown
	' 			EMReadScreen reference, 2, row, 6
	' 			EMReadScreen period, 13, row, 61
	' 			EMReadScreen cname, 21, row, 10
	' 			cname = trim(cname)
	' 			If cname = "" Then EMReadScreen cname, 21, row - 1, 10
	' 			cname = trim(cname)
    '
	' 			ReDim Preserve HC_CLIENTS_DETAIL_ARRAY (12, hc_clt)			'Adding any client with a spenddown to a new array
    '
	' 			HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt) = HC_CASES_ARRAY(wrk_num, hc_case)
	' 			HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt) = MAXIS_case_number
	' 			HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt) = replace(HC_CASES_ARRAY(next_revw, hc_case), " ", "/")
	' 			HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt) = cname
	' 			HC_CLIENTS_DETAIL_ARRAY (ref_numb,  hc_clt) = reference
	' 			HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt) = spd_type
	' 			HC_CLIENTS_DETAIL_ARRAY (spd_pd,    hc_clt) = period
    '
	' 			hc_clt = hc_clt + 1
    '
	' 		End If
	' 		row = row + 1
	' 	Loop until row = 19
	' End If

''MY CODE===============================================
'This bit will look to see if there are any cases that have a possible spenddown.
'Occasionally the criteria selected produce no cases and this explains this to the user.
If UBound(HC_CLIENTS_DETAIL_ARRAY, 2) = 0 AND HC_CLIENTS_DETAIL_ARRAY(case_num, 0) = "" Then
	all_workers = Join(worker_array, ", ")
	If one_month_only = True Then
		selected_time = " for the month of " & revw_month_list & "."
	Else
		selected_time = "."
	End If
	end_msg = "Success! The script has completed!" & vbNewLine & "NO HC Cases FOUND!" & vbNewLine & vbNewLine &_
	          "The script has checked REPT/ACTV for the case loads under worker number(s) " & all_workers & selected_time & vbNewLine &_
			  "None of the active HC cases have a spenddown indicated on MOBL." & vbNewLine & vbNewLine &_
			  "No report will be generated, the script has completed."
	script_end_procedure(end_msg)
End If
'
' For hc_clt = 0 to UBound(HC_CLIENTS_DETAIL_ARRAY, 2)
'     MAXIS_case_number = HC_CLIENTS_DETAIL_ARRAY(case_num, hc_clt)
'     Call navigate_to_MAXIS_screen("STAT", "MEMB")
'     EMWriteScreen HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt), 20, 76
'     transmit
'
'     EMReadScreen memb_pmi, 8, 4, 46
'     memb_pmi = right("00000000" & memb_pmi, 8)
'     HC_CLIENTS_DETAIL_ARRAY(clt_pmi, hc_clt) = memb_pmi
'
'     EMReadScreen memb_age, 3, 8, 76
'     memb_age = trim(memb_age)
'
'     Call navigate_to_MAXIS_screen ("STAT", "MEDI")
'     EMWriteScreen HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt), 20, 76
'     transmit
'
'     EMReadScreen part_a_end, 8, 15, 35
'     EMReadScreen part_b_end, 8, 15, 65
'     EMReadScreen medi_part_a_prem, 8, 7, 46
'     EMReadScreen medi_part_b_prem, 8, 7, 73
'     EMReadScreen apply_prem_to_spdn, 1, 11, 71
'
'     Call navigate_to_MAXIS_screen ("STAT", "PREG")
'     EMWriteScreen HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt), 20, 76
'     transmit
'
'     EMReadScreen preg_end_date, 8, 12, 53
'     preg_end_date = replace(preg_end_date, " ", "/")
'
'     Call navigate_to_MAXIS_screen ("STAT", "DISA")
'     EMWriteScreen HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt), 20, 76
'     transmit
'
'     EMReadScreen disa_waiver, 1, 14, 59
'     EMReadScreen disa_hc_status, 2, 13, 59
'     EMReadScreen disa_verif, 1, 13, 69
'     EMReadScreen disa_1619_status, 1, 16, 59
'
'     EMReadScreen disa_end_date, 10, 6, 69
'     disa_end_date = replace(disa_end_date, " ", "/")
'
'     EMReadScreen disa_cert_end_date, 10, 7, 69
'     disa_cert_end_date = replace(disa_cert_end_date, " ", "/")
'
'     client_earned_income = 0
'
'     Call navigate_to_MAXIS_screen ("STAT", "JOBS")
'     EMWriteScreen HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt), 20, 76
'     transmit
'
'     EMReadScreen jobs_versions, 1, 2, 78            'Reading the total number of jobs for this member
'     If jobs_versions <> "0" Then                    'If there are no jobs - nothing to look at
'         jobs_versions = jobs_versions * 1           'making the number of jobs listed an actual number
'
'         For the_job = 1 to jobs_versions            'go through each job
'             EMWriteScreen "0", 20, 79               'enter a leading 0 in the panel number indicated
'             EMWriteScreen the_job, 20, 80           'enter the job number just after the 0
'             transmit
'
'             EMWriteScreen "X", 19, 48               'Opening the HC In Est pop-up
'             transmit
'
'             EMReadScreen income_per_pay, 8, 11, 63  'Reading the income listed in the pop-up
'             PF3                                     'backing out to the main panel
'
'             income_per_pay = trim(income_per_pay)       'formating the amount of income read
'             if income_per_pay = "________" Then income_per_pay = 0
'             income_per_pay = income_per_pay * 1
'
'             EMReadScreen pay_freq, 1, 18, 35            'reading the pay frequency
'             if pay_freq = "_" Then pay_freq = 1
'
'             monthly_income = pay_freq * income_per_pay      'calculating the monthly income
'
'             client_earned_income = client_earned_income + monthly_income    'adding the monthly income from each job to the running total
'         Next
'     End If
'
'     Call navigate_to_MAXIS_screen ("STAT", "BUSI")                          'Going to BUSI for this member
'     EMWriteScreen HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt), 20, 76
'     transmit
'
'     EMReadScreen busi_versions, 1, 2, 78            'Reading the total number of BUSI panels for this member
'     If busi_versions <> "0" Then                    'If there are no BUSI panels - nothing to look at
'         busi_versions = busi_versions * 1           'making the number of BUSI panels listed an actual number
'
'         For the_busi = 1 to busi_versions            'go through each BUSI
'             EMWriteScreen "0", 20, 79               'enter a leading 0 in the panel number indicated
'             EMWriteScreen the_busi, 20, 80           'enter the busi number just after the 0
'             transmit
'
'             EMWriteScreen "X", 17, 27               'Opening the HC In Est pop-up
'             transmit
'
'             EMReadScreen method_a_income, 8, 15, 54
'             EMReadScreen method_b_income, 8, 16, 54
'
'             method_a_income = trim(method_a_income)
'             method_b_income = trim(method_b_income)
'
'             method_a_income = method_a_income * 1
'             method_b_income = method_b_income * 1
'
'             PF3                                     'Going back to the main panel
'
'             If HC_CLIENTS_DETAIL_ARRAY(elig_mthd, hc_clt) = "A" Then
'                 client_earned_income = client_earned_income + method_a_income
'             ElseIf HC_CLIENTS_DETAIL_ARRAY(elig_mthd, hc_clt) = "B" Then
'                 client_earned_income = client_earned_income + method_b_income
'             Else
'                 If method_a_income >= method_b_income Then client_earned_income = client_earned_income + method_a_income
'                 If method_a_income < method_b_income Then client_earned_income = client_earned_income + method_b_income
'             End If
'         Next
'     End If
'
'
'
' Next

' navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")
' EmWriteScreen "X", 8, 3                             'Entering Recipient file application'
' transmit
'
' For hc_clt = 0 to UBound(HC_CLIENTS_DETAIL_ARRAY, 2)
'     PMI_Number = HC_CLIENTS_DETAIL_ARRAY(clt_pmi, hc_clt)
'
'     EmWriteScreen "I", 2, 19                'Going to RSUM for the individual'
'     EmWriteScreen PMI_Number, 4, 19
'     transmit
'
'     EMReadScreen MMIS_waiver, 1, 15, 15     'Read waiver type and dates
'     EMReadScreen MMIS_waiver_begin, 8, 15, 25
'     EMReadScreen MMIS_waiver_through, 8, 15, 46
'
'     EMReadScreen MMIS_PPHP_begin, 8, 16, 20     'Read PPHP dates and plan
'     EMReadScreen MMIS_PPHP_end, 8, 16, 37
'     EMReadScreen MMIS_PPHP_plan, 10, 16, 52
'     EMReadScreen MMIS_PPHP_prod_ID, 5, 16, 72
'
'     EmWriteScreen "RELG", 1, 8                  'Go to RELG
'     transmit
'     'Find open Spans
'     relg_row = 6
'     Do
'         EMReadScreen MMIS_span_elig_end, 8, relg_row + 1, 36
'
'
'     Loop
'
'     'Go to RSPL
'     'Review for spenddowns
'
' Next
'MY CODE END=========================================='

' 'Gathering additional information about each client with a spenddown indicated - SPENDDOWN REPORT CODE
' For hc_clt = 0 to UBound(HC_CLIENTS_DETAIL_ARRAY, 2)
' 	spd_amt = 0			'Reset the variable for each run
' 	MAXIS_case_number = HC_CLIENTS_DETAIL_ARRAY(case_num, hc_clt)				'Setting the case number for global functions
' 	HC_CLIENTS_DETAIL_ARRAY(add_xcl, hc_clt) = TRUE
' 	Call navigate_to_MAXIS_screen ("CASE", "PERS")								'Confirming clt is active HC this month
' 	row = 9
' 	Do
' 		EMReadScreen person, 2, row, 3
' 		If person = HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt) Then
' 			EMReadScreen hc_stat, 1, row, 61
' 			If hc_stat = "I" OR hc_stat = "D" Then HC_CLIENTS_DETAIL_ARRAY(add_xcl, hc_clt) = FALSE 	'If not, case will not be added to report
' 			Exit Do
' 		Else
' 			row = row + 1
' 			If row = 18 Then
' 				EMReadScreen next_page, 7, row, 3
' 				If next_page = "More: +" Then
' 					PF8
' 					row = 9
' 				End If
' 			End If
' 		End If
' 	Loop until row = 18
' 	IF HC_CLIENTS_DETAIL_ARRAY(add_xcl, hc_clt) = TRUE Then 						'If clt is actve HC
' 		STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
' 		Call navigate_to_MAXIS_screen ("ELIG", "HC")								'Need a closer look at HC
' 		row = 8
' 		Do
' 			EMReadScreen person, 2, row, 3									'Finding the correct person on HC ELIG
' 			If person = HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt) Then
' 				Do
' 					EMReadScreen prog, 2, row, 28							'Find the line that has this persons MA listed on it /NOT QMB etc
' 					If prog = "MA" Then
' 						counter = 1
' 						Do 													'Here this will find if this version is approved
' 							EMReadScreen app_indc, 5, row, 68
' 							app_indc = trim(app_indc)
' 							If app_indc = "APP" Then						'If approved, open this version
' 								Call write_value_and_transmit("x", row, 26)
' 								Exit Do
' 							End If
' 							EMReadScreen this_version, 2, row, 58
' 							If this_version = "00" Then
' 								EMReadScreen elig_month, 2, 20, 56			'If the earliest version in this month was not approved then it goes to the previous month
' 								If elig_month <> "01" Then
' 									last_month = right("00" & (abs(elig_month)-1), 2)
' 									EMWriteScreen last_month, 20, 56
' 									transmit
' 								Else
' 									last_month = "12"
' 									EMReadScreen elig_year, 2, 20, 59
' 									last_year = right("00" & (abs(elig_year)-1), 2)
' 									EMWriteScreen last_month, 20, 56
' 									EMWriteScreen last_year, 20, 59
' 								End If
' 								counter = counter + 1
' 							ElseIf this_version <> "01" Then 					'Checking to see if there is a pervious version listed in this month
' 								prev_verision = right("00" & (abs(this_version)-1), 2)	'If so, it will go to the previous version
' 								EMWriteScreen prev_verision, row, 58
' 								transmit
' 							Else
' 								EMReadScreen elig_month, 2, 20, 56			'If the earliest version in this month was not approved then it goes to the previous month
' 								If elig_month <> "01" Then
' 									last_month = right("00" & (abs(elig_month)-1), 2)
' 									EMWriteScreen last_month, 20, 56
' 									transmit
' 								Else
' 									last_month = "12"
' 									EMReadScreen elig_year, 2, 20, 59
' 									last_year = right("00" & (abs(elig_year)-1), 2)
' 									EMWriteScreen last_month, 20, 56
' 									EMWriteScreen last_year, 20, 59
' 								End If
' 								counter = counter + 1
' 							End If
' 						Loop until counter = 6				'Only looks at 6 months
' 					ELSE					'If this person could not be found then the report will list no version was found
' 						row = row + 1
' 						EMReadScreen person, 2, row, 3
' 						If person <> "  "  Then
' 							HC_CLIENTS_DETAIL_ARRAY(hc_type, hc_clt) = "NO HC VERSION"
' 							Exit Do
' 						End If
' 					End If
' 				Loop Until row = 20
' 				Exit Do
' 			Else
' 				row = row + 1
' 			End If
' 		Loop until row = 20
' 		EMReadScreen bsum_check, 4, 3, 57		'Confirming that HC Elig has been opened for this person
' 		If bsum_check = "BSUM" Then
' 			col = 19
' 			Do									'Finding the current month in elig to get the current elig type
' 				EMReadScreen span_month, 2, 6, col
' 				If span_month = MAXIS_footer_month Then		'reading the ELIG TYPE
' 					EMReadScreen pers_type, 2, 12, col - 2
' 					EMReadScreen std, 1, 12, col + 3
' 					EMReadScreen meth, 1, 13, col + 2
' 					Exit Do
' 				End If
' 				col = col + 11
' 				If col = 85 Then 		'If this month was not found then it reads the LAST elig type in elig
' 					EMReadScreen pers_type, 2, 12, 72		'ONLY saves this information if an actual elig type was found
' 					If pers_type <> "11" AND pers_type <> "09" AND pers_type <> "PX" AND pers_type <> "PC" AND pers_type <> "CB" AND pers_type <> "CK" AND pers_type <> "CX" AND pers_type <> "CM" AND pers_type <> "AA" AND pers_type <> "AX" AND pers_type <> "BT" AND pers_type <> "DT" AND pers_type <> "15" AND pers_type <> "16" AND pers_type <> "DC" AND pers_type <> "EX" AND pers_type <> "DX" AND pers_type <> "DP" AND pers_type <> "BC" AND pers_type <> "RM" AND pers_type <> "10" AND pers_type <> "25" Then
' 						pers_type = ""
' 					Else
' 						EMReadScreen std, 1, 12, 77
' 						EMReadScreen meth, 1, 13, 76
' 					End If
' 				End If
' 			Loop until col = 85
' 			If pers_type = "" Then 				'Setting the elig type to readable format
' 				HC_CLIENTS_DETAIL_ARRAY(hc_type, hc_clt) = "ELIG Type Not Found"
' 			Else
' 				HC_CLIENTS_DETAIL_ARRAY(hc_type, hc_clt) = pers_type & "-" & std & " Method: " & meth
' 				pers_type = ""
' 				std = ""
' 				meth = ""
' 			End If
' 			spd_amt = 0
' 			col = 18
' 			Do 				'This will gather the 6 month standard AND the budgeted income to calculate the HC overage
' 				EMReadScreen month_net_inc, 8, 15, col
' 				EMReadScreen month_std_inc, 8, 16, col
' 				month_net_inc = trim(month_net_inc)
' 				If month_net_inc = "" Then month_net_inc = 0
' 				month_std_inc = trim(month_std_inc)
' 				If month_std_inc = "" Then month_std_inc = 0
' 				tot_net_inc = tot_net_inc + abs(month_net_inc)
' 				tot_std_inc = tot_std_inc + abs(trim(month_std_inc))
' 				col = col + 11
' 			Loop until col = 84
' 			spd_amt =  tot_net_inc - tot_std_inc
' 			If spd_amt < 0 Then spd_amt = 0
' 			HC_CLIENTS_DETAIL_ARRAY(hc_excess, hc_clt) = spd_amt
' 			'NOTE that Cert Period Amount popup was NOT used as it appears to change what is listed on MOBL if the spenddown was in error
' 			'We do not want bulk reports to make alterations to cases without worker review and approval
' 		End If
'
' 		'Goes to get PMI
' 		Call navigate_to_MAXIS_screen ("STAT", "MEMB")
' 		EMWriteScreen HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt), 20, 76
' 		transmit
' 		EMReadScreen pmi, 8, 4, 46
' 		HC_CLIENTS_DETAIL_ARRAY(clt_pmi, hc_clt) = right("00000000" & replace(pmi, "_", ""), 8)
' 	End If
' 	back_to_self
' Next
'
' If MMIS_checkbox = checked Then
' 	'Now it will look for MMIS on both screens, and enter into it..
' 	attn
' 	EMReadScreen MMIS_A_check, 7, 15, 15
' 	If MMIS_A_check = "RUNNING" then
' 		EMSendKey "10"
' 		transmit
' 	Else
' 		attn
' 		EMConnect "B"
' 		attn
' 		EMReadScreen MMIS_B_check, 7, 15, 15
' 		If MMIS_B_check <> "RUNNING" then
' 			MMIS_checkbox = unchecked
' 			script_continue = MsgBox ("MMIS does not appear to be running." & vbNewLine & "Do you wish to have the report without the MMIS Spenddown Indicator checked?", vbYesNo + vbQuestion, "MMIS not running")
' 			IF script_continue = vbNo Then script_end_procedure ("Script has ended with no report generated. To have MMIS information gathered, be sure to have MMIS running and not be passworded out.")
' 		Else
' 			EMSendKey "10"
' 			transmit
' 		End if
' 	End if
' End If
'
' If MMIS_checkbox = checked Then
' 	EMFocus 'Bringing window focus to the second screen if needed.
'
' 	'Sending MMIS back to the beginning screen and checking for a password prompt
' 	Do
' 		PF6
' 		EMReadScreen password_prompt, 38, 2, 23
' 	  	IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then
' 		  	MMIS_checkbox = unchecked
' 		  	script_continue = MsgBox ("MMIS does not appear to be running." & vbNewLine & "Do you wish to have the report without the MMIS Spenddown Indicator checked?", vbYesNo + vbQuestion, "MMIS not running")
' 		  	IF script_continue = vbNo Then script_end_procedure ("Script has ended with no report generated. To have MMIS information gathered, be sure to have MMIS running and not be passworded out.")
' 			Exit Do
' 		End If
' 	  	EMReadScreen session_start, 18, 1, 7
' 	Loop until session_start = "SESSION TERMINATED"
' End If
'
' If MMIS_checkbox = checked Then
' 	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themself into MMIS the first time!)
' 	EMWriteScreen "mw00", 1, 2
' 	transmit
' 	transmit
'
' 	'Finding the right MMIS, if needed, by checking the header of the screen to see if it matches the security group selector
' 	EMReadScreen MMIS_security_group_check, 21, 1, 35
' 	If MMIS_security_group_check = "MMIS MAIN MENU - MAIN" then
' 		EMSendKey "x"
' 		transmit
' 	End if
'
' 	'Now it finds the recipient file application feature and selects it.
' 	row = 1
' 	col = 1
' 	EMSearch "RECIPIENT FILE APPLICATION", row, col
' 	EMWriteScreen "x", row, col - 3
' 	transmit
'
' 	For hc_clt = 0 to UBound(HC_CLIENTS_DETAIL_ARRAY, 2)			'Opens RELG for each client to get spenddown indicator
' 		indicator = ""
' 		'Now we are in RKEY, and it navigates into the case, transmits, and makes sure we've moved to the next screen.
' 		EMWriteScreen "i", 2, 19
' 		EMWriteScreen HC_CLIENTS_DETAIL_ARRAY(clt_pmi, hc_clt), 4, 19	'Enters PMI
' 		transmit		'Goes to RSUM
' 		EMReadscreen RKEY_check, 4, 1, 52
' 		If RKEY_check = "RKEY" then 		'Confirms that we have moved past RKEY
' 			HC_CLIENTS_DETAIL_ARRAY (mmis_spdn, hc_clt) = "Not Found"
' 		Else
' 			EMWriteScreen "RELG", 1, 8		'Goes to RELG
' 			transmit
' 			row = 7
' 			Do 				'Finding the openended OR future close MA span
' 				EMReadscreen elig_end, 8, row, 36
' 				IF elig_end <> "99/99/99" Then after_now = DateDiff("d", date, elig_end)
' 				If elig_end = "99/99/99" OR after_now < 0 Then
' 					EMReadscreen prg, 2, row-1, 10
' 					IF prg = "MA" Then 			'Reads the spenddown indicator
' 						EMReadscreen indicator, 1, row + 1, 62
' 						Exit Do
' 					End If
' 				End If
' 				row = row + 4
' 			Loop until row = 23
'
' 			PF6
' 			EMWriteScreen "        ", 4, 19		'Blanking out the PMI for safety
'
' 			If indicator = "" Then 				'Setting the indicator to the array
' 				HC_CLIENTS_DETAIL_ARRAY (mmis_spdn, hc_clt) = "Not Found"
' 			Else
' 				HC_CLIENTS_DETAIL_ARRAY (mmis_spdn, hc_clt) = indicator
' 			End If
' 		End If
'
' 	Next
' End If

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

ObjExcel.Cells(1, col_to_use).Value = "NEXT REVW DATE"
revw_date_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "1st Prog"
prog_one_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "ELIG TYPE"
elig_one_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "2nd PROG"
prog_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "ELIG TYPE"
elig_two_col = col_to_use
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

'Adding all client information to a spreadsheet for your viewing pleasure
For hc_clt = 0 to UBound(HC_CLIENTS_DETAIL_ARRAY, 2)
	If HC_CLIENTS_DETAIL_ARRAY(add_xcl, hc_clt) = TRUE Then
		ObjExcel.Cells(excel_row, worker_col).Value       = HC_CLIENTS_DETAIL_ARRAY (wrk_num,   hc_clt)
		ObjExcel.Cells(excel_row, case_numb_col).Value    = HC_CLIENTS_DETAIL_ARRAY (case_num,  hc_clt)
		ObjExcel.Cells(excel_row, ref_numb_col).Value     = "Memb " & HC_CLIENTS_DETAIL_ARRAY(ref_numb, hc_clt)
		ObjExcel.Cells(excel_row, name_col).Value         = HC_CLIENTS_DETAIL_ARRAY (clt_name,  hc_clt)
		ObjExcel.Cells(excel_row, pmi_col).Value          = HC_CLIENTS_DETAIL_ARRAY (clt_pmi,   hc_clt)
		ObjExcel.Cells(excel_row, revw_date_col).Value    = HC_CLIENTS_DETAIL_ARRAY (next_revw, hc_clt)
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

        ObjExcel.Cells(excel_row, prog_two_col).Value     = HC_CLIENTS_DETAIL_ARRAY (hc_prog_two,   hc_clt)
        If HC_CLIENTS_DETAIL_ARRAY (hc_prog_two,   hc_clt) = "MA" Then
            ObjExcel.Cells(excel_row, elig_two_col).Value  = HC_CLIENTS_DETAIL_ARRAY (elig_type_two,   hc_clt) & "-" & HC_CLIENTS_DETAIL_ARRAY(elig_std_two, hc_clt) & " - Method: " & HC_CLIENTS_DETAIL_ARRAY(elig_mthd_two, hc_clt)
            If HC_CLIENTS_DETAIL_ARRAY(mobl_spdn, hc_clt) <> "NO SPENDDOWN" Then
                ObjExcel.Cells(excel_row, spdn_col).Value  = HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt) & " for " & HC_CLIENTS_DETAIL_ARRAY(spd_pd, hc_clt)
            Else
                ObjExcel.Cells(excel_row, spdn_col).Value  = HC_CLIENTS_DETAIL_ARRAY (mobl_spdn, hc_clt)
            End If
        Else
            ObjExcel.Cells(excel_row, elig_two_col).Value  = HC_CLIENTS_DETAIL_ARRAY (elig_type_two,   hc_clt) & "-" & HC_CLIENTS_DETAIL_ARRAY(elig_std_two, hc_clt)
        End If


        'ObjExcel.Cells(excel_row, 9).Value  = HC_CLIENTS_DETAIL_ARRAY (hc_excess, hc_clt)
		'ObjExcel.Cells(excel_row, 10).Value = HC_CLIENTS_DETAIL_ARRAY (mmis_spdn, hc_clt)
        ObjExcel.Cells(excel_row, waiver_col).Value     = HC_CLIENTS_DETAIL_ARRAY(elig_waiv, hc_clt)
        ObjExcel.Cells(excel_row, errors_col).Value     = HC_CLIENTS_DETAIL_ARRAY(error_notes, hc_clt)
		excel_row = excel_row + 1
	End If
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
script_end_procedure("Success! All cases for selected workers that appear to have a Spenddown indicated in MAXIS have been added to the Excel Spreadsheet.")
