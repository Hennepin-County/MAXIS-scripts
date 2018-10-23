'Required for statistical purposes==========================================================================================
name_of_script = "BULK - Resolve HC EOMC in MMIS.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 1                       'manual run time in seconds
STATS_denomination = "M"       							'C is for each CASE
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
call changelog_update("09/21/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS==================================================================================================================
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

'function specific to this script - running_stopwatch and MX_environment are defined outside of this function
'meant to keep MMIS from passwording out while this long bulk script is running
function keep_MMIS_passworded_in()
    If timer - running_stopwatch > 720 Then         'this means the script has been running for more than 12 minutes since we last popped in to MMIS
        Call navigate_to_spec_MMIS_region("CTY ELIG STAFF/UPDATE")      'Going to MMIS'
        Call navigate_to_MAXIS(MX_environment)                          'going back to MAXIS'

        running_stopwatch = timer                                       'resetting the stopwatch'
    End If
end function
'===========================================================================================================================

'ELEMENTS TO DECLARE------------------------------------------------------------------------------------------------------------------
'Variables that need to be defined early in the script
all_case_numbers_array = " "					'Creating blank variable for the future array
two_digit_county_code = "27"                    'hard coding defining this because it is county specific'
is_not_blank_excel_string = Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34)	'This is the string required to tell excel to ignore blank cells in a COUNTIFS function
MAXIS_footer_month = CM_mo                      'defining footer month and year - we will always start in current month - EOMC doesn't function to make changes in other months
MAXIS_footer_year = CM_yr
first_of_next_month = CM_plus_1_mo & "/1/" & CM_plus_1_yr           'creating a date varibale for the first of the next month
last_day_of_this_month = DateAdd("d", -1, first_of_next_month)      'creating a date variable with the LAST day of the current month

'formatting the last day of the month in to MM/DD/YY for entry in to MMIS
last_day_mo = DatePart("m", last_day_of_this_month)
last_day_mo = right("00" & last_day_mo, 2)
last_day_day = DatePart("d", last_day_of_this_month)
last_day_day = right("00" & last_day_day, 2)
last_day_yr = DatePart("yyyy", last_day_of_this_month)
last_day_yr = right("00" & last_day_yr, 2)
mmis_last_day_date = last_day_mo & "/" & last_day_day & "/" & last_day_yr

'Setting amounts
total_savings = 0                   'setting this at zero so that we can add up what we save
capitation_11x      = 864.45        'capitation amounts set annually by DHS - eventually we need to move this to FuncLib
capitation_PW       = 1174.15
capitation_1        = 243.67
capitation_2_15     = 244.00
capitation_16_20    = 267.10
capitation_21_49    = 794.03
capitation_50_64    = 1058.51
capitation_65       = 2354.34

capitation_QMB      = 104.90
capitation_SLMB     = 104.90
capitation_QI1      = 104.90

'Constants
Const basket_nbr            = 0
Const case_nbr              = 1
Const clt_name              = 2
Const autoclose             = 3
Const hc_close_stat         = 4
Const clt_pmi               = 5
Const clt_ref_nbr           = 6
Const clt_age               = 7
Const hc_prog_one           = 8
Const elig_type_one         = 9
Const prog_one_end          = 10
Const hc_prog_two           = 11
Const elig_type_two         = 12
Const prog_two_end          = 13
Const MMIS_span_one         = 14
Const MMIS_curr_end_one     = 15
Const MMIS_new_end_one      = 16
Const RELG_page_one         = 17
Const RELG_row_one          = 18
Const MMIS_span_two         = 19
Const MMIS_curr_end_two     = 20
Const MMIS_new_end_two      = 21
Const RELG_page_two         = 22
Const RELG_row_two          = 23
Const clt_savings           = 24
Const capitation_ended      = 25
Const err_notes             = 26

'Arrays
Dim EOMC_CASES_ARRAY()
ReDim EOMC_CASES_ARRAY(hc_close_stat, 0)

Dim EOMC_CLIENT_ARRAY()
ReDim EOMC_CLIENT_ARRAY(err_notes, 0)
'THE SCRIPT-----------------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'dialog to restrict how many baskets the script is run on AND decide if the script will be run to change or just look up information
BeginDialog EOMC_dialog, 0, 0, 351, 65, "Workers to check EOMC"
  EditBox 90, 10, 255, 15, list_of_workers
  CheckBox 10, 45, 140, 10, "Check here to have script update MMIS", change_checkbox
  ButtonGroup ButtonPressed
    OkButton 240, 45, 50, 15
    CancelButton 295, 45, 50, 15
  Text 5, 15, 85, 10, "List of Workers to check:"
  Text 95, 30, 125, 10, "(Leave blank to run on entrire county)"
EndDialog

'Showing the dialog
Do
    Dialog EOMC_dialog

    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checking for MAXIS
Call check_for_MAXIS(True)
Call back_to_SELF                                               'starting at the SELF panel
EMReadScreen MX_environment, 13, 22, 48                         'seeing which MX environment we are in
MX_environment = trim(MX_environment)
Call navigate_to_spec_MMIS_region("CTY ELIG STAFF/UPDATE")      'Going to MMIS'
Call navigate_to_MAXIS(MX_environment)                          'going back to MAXIS

running_stopwatch = timer               'setting the running timer so we log in to MMIS within every 15 mintues so we don't password out

make_changes = FALSE                    'setting this at the start
If change_checkbox = checked Then make_changes = TRUE   'if the dialog has indicated that changes should be changed reset this to true

list_of_workers = trim(list_of_workers)
If list_of_workers = "" Then            'if this is blank then we are going to search the entire county
    call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else                                    'if x numbers are entered, then we are going to look at JUST those baskets
    worker_array = split(list_of_workers, ",")
End If

list_of_cases = 0                       'array incrementer
For each worker in worker_array         'going through each worker in the list of workers
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
    worker = trim(worker)               'making sure there aren't any spaces around the basket number
	Call navigate_to_MAXIS_screen("rept", "eomc")      'go to EOMC for the correct worker number
	EMWriteScreen worker, 21, 16
	transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 5
	If has_content_check <> " " then
        Do
            EMReadScreen last_page_check, 21, 24, 2	'because on REPT/EOMC it displays right away, instead of when the second F8 is sent

            'Set variable for next do...loop
            MAXIS_row = 7
            Do
                EMReadScreen MAXIS_case_number, 8, MAXIS_row, 7		   'Reading case number
                EMReadScreen client_name, 25, MAXIS_row, 16		       'Reading client name
                EMReadScreen cash_status, 4, MAXIS_row, 43		       'Reading cash status
                EMReadScreen SNAP_status, 4, MAXIS_row, 53		       'Reading SNAP status
                EMReadScreen HC_status, 4, MAXIS_row, 58			   'Reading HC status
                EMReadScreen GRH_status, 4, MAXIS_row, 68			   'Reading GRH status

                'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
                MAXIS_case_number = trim(MAXIS_case_number)
                If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
                all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")

                If MAXIS_case_number = "" then exit do			'Exits do if we reach the end

                If HC_status <> "    " then                     'we are only going to add cases closing HC
                    ReDim Preserve EOMC_CASES_ARRAY(hc_close_stat, list_of_cases)       'resizing the array
                    EOMC_CASES_ARRAY(autoclose, list_of_cases) = FALSE                  'default to false

                    EOMC_CASES_ARRAY(basket_nbr, list_of_cases) = worker                'adding the cse information
                    EOMC_CASES_ARRAY(case_nbr, list_of_cases) = MAXIS_case_number
                    EOMC_CASES_ARRAY(clt_name, list_of_cases) = client_name
                    If right(HC_status, 1) = "A" then EOMC_CASES_ARRAY(autoclose, list_of_cases) = TRUE     'if these are autoclose- redefining this
                    EOMC_CASES_ARRAY(hc_close_stat, list_of_cases) = trim(HC_status)

                    list_of_cases = list_of_cases + 1       'incrementing
                End If

                MAXIS_row = MAXIS_row + 1       'look at the next case
                MAXIS_case_number = ""			'Blanking out variable
            Loop until MAXIS_row = 19
            PF8     'go to the next page
        Loop until last_page_check = "THIS IS THE LAST PAGE"
    End if
    Call keep_MMIS_passworded_in                'every 12 mintues or so, the script will pop in to MMIS to make sure we are passworded in
Next

hc_clt = 0
'The script will now look in each case at MOBL to identify clients on HC for each case
For hc_case = 0 to UBound(EOMC_CASES_ARRAY, 2)
    back_to_SELF                                                'Back to SELF at the beginning of each run so that we don't end up in the wrong case
    MAXIS_case_number = EOMC_CASES_ARRAY(case_nbr, hc_case)		'defining case number for functions to use
    on_page = 1                                                 'saving which page we are on because '

    ' If MAXIS_case_number = "" Then MsgBox "Case number: " &  EOMC_CASES_ARRAY(case_nbr, hc_case)
    Call navigate_to_MAXIS_screen("CASE", "PERS")               'Getting client eligibility of HC from CASE PERS
    pers_row = 10                                               'This is where client information starts on CASE PERS
    Do
        clt_hc_ending = FALSE                                   'defining this at the beginning of each row of CASE PERS
        EMReadScreen clt_hc_status, 1, pers_row, 61             'reading the HC status of each client
        If clt_hc_status = "A" Then                             'if HC is active then we will add this client to the array to find additional information
            If EOMC_CASES_ARRAY(autoclose, hc_case) = TRUE Then clt_hc_ending = TRUE        'if client is active on a case that will AutoClose, then they will have HC ended
            If EOMC_CASES_ARRAY(autoclose, hc_case) = FALSE Then                'If the case is not autoclosing we need to look and see if this client's HC will close
                EMReadScreen pers_pmi_numb,  8, pers_row, 34                    'reading the PMI

                EMWriteScreen CM_plus_1_mo, 19, 54                              'go to the next month
                EMWriteScreen CM_plus_1_yr, 19, 57
                transmit

                If on_page > 1 Then                                             'if we had to PF8 on current month, doing the same in CM+1
                    For the_page = 2 to on_page
                        PF8
                    Next
                End If

                EMReadScreen the_pmi, 8, pers_row, 34                           'reading the pmi from the same place where we read it in CM
                If the_pmi <> pers_pmi_numb Then                                'if they don't match, we are going to look for the person
                    save_pers_row = pers_row                                    'saving the row because we need it when we go back to CM
                    Do                                                          'go to the first page
                        PF7
                        EMReadScreen at_beginning, 10, 24, 14
                    Loop until at_beginning = "FIRST PAGE"
                    pers_row = 10                                               'This is where client information starts on CASE PERS
                    Do
                        EMReadScreen the_pmi, 8, pers_row, 34                   'reading the pmi on each road

                        If the_pmi = pers_pmi_numb Then Exit Do                 'if we find a match, we are at the right person

                        pers_row = pers_row + 3         'next client information is 3 rows down
                        If pers_row = 19 Then           'this is the end of the list of client on each list
                            PF8                         'going to the next page of client information
                            pers_row = 10
                            EMReadScreen end_of_list, 9, 24, 14
                            If end_of_list = "LAST PAGE" Then Exit Do
                        End If
                        EMReadScreen next_pers_ref_numb, 2, pers_row, 3     'this reads for the end of the list

                    Loop until next_pers_ref_numb = "  "

                    If next_pers_ref_numb = "  " Then       'if we can't find the right person, then the person was removed from the case and the HC has ended
                        clt_hc_ending = TRUE
                    Else                                    'if we did find the right person, we check the status
                        EMReadScreen clt_hc_status, 1, pers_row, 61
                        If clt_hc_status = "I" Then clt_hc_ending = TRUE        'If the status is 'I - inactive' in CM+1 then HC ended for this client
                    End If

                    pers_row = save_pers_row                'resetting the row variable for the next client review
                Else        'if the pmi matched right away, we look to see the status in CM+1
                    EMReadScreen clt_hc_status, 1, pers_row, 61
                    If clt_hc_status = "I" Then clt_hc_ending = TRUE            'If the status is 'I - inactive' in CM+1 then HC ended for this client
                End If

                EMWriteScreen MAXIS_footer_month, 19, 54        'going back to CM
                EMWriteScreen MAXIS_footer_year, 19, 57

                transmit

                If on_page > 1 Then                 'going back to the right page of CASE/PERS
                    For the_page = 2 to on_page
                        PF8
                    Next
                End If

            End If

            If clt_hc_ending = TRUE Then            'If the HC is ending for this client, the client information is added to the client array

                EMReadScreen pers_ref_numb,  2, pers_row, 3         'reading the client information
                EMReadScreen pers_pmi_numb,  8, pers_row, 34
                EMReadScreen pers_last_name, 15, pers_row, 6
                EMReadScreen pers_frst_name, 11, pers_row, 22

                pers_pmi_numb = trim(pers_pmi_numb)                 'formatting the information read
                pers_last_name = trim(pers_last_name)
                pers_frst_name = trim(pers_frst_name)

                ReDim Preserve EOMC_CLIENT_ARRAY (err_notes, hc_clt)        'resizing the Client array

                EOMC_CLIENT_ARRAY (basket_nbr,   hc_clt) = EOMC_CASES_ARRAY(basket_nbr,  hc_case)       'some information is saved from the CASES array
                EOMC_CLIENT_ARRAY (case_nbr,  hc_clt) = EOMC_CASES_ARRAY(case_nbr, hc_case)
                EOMC_CLIENT_ARRAY (autoclose, hc_clt) = EOMC_CASES_ARRAY(autoclose, hc_case)
                EOMC_CLIENT_ARRAY (hc_close_stat, hc_clt) = EOMC_CASES_ARRAY(hc_close_stat, hc_case)
                'EOMC_CLIENT_ARRAY (next_revw, hc_clt) = ObjExcel.Cells(excel_row, ). Value
                EOMC_CLIENT_ARRAY (clt_name,  hc_clt) = pers_last_name & ", " & pers_frst_name          'some information is saved from the CASE/PERS information
                EOMC_CLIENT_ARRAY (clt_ref_nbr,  hc_clt) = pers_ref_numb
                EOMC_CLIENT_ARRAY (clt_pmi,   hc_clt) = pers_pmi_numb

                hc_clt = hc_clt + 1     'incrementing the array
            End If
        End If

        pers_row = pers_row + 3         'next client information is 3 rows down
        If pers_row = 19 Then           'this is the end of the list of client on each list
            PF8                         'going to the next page of client information
            on_page = on_page + 1       'saving that we have gone to a new page
            pers_row = 10               'resetting the row to read at the top of the next page
            EMReadScreen end_of_list, 9, 24, 14
            If end_of_list = "LAST PAGE" Then Exit Do
        End If
        EMReadScreen next_pers_ref_numb, 2, pers_row, 3     'this reads for the end of the list

    Loop until next_pers_ref_numb = "  "

    Call keep_MMIS_passworded_in        'making sure we are not passworded out in MMIS
Next

'Now the client array is created
'Information gathering in MAXIS now for every client on HC on the list
For hc_clt = 0 to UBOUND(EOMC_CLIENT_ARRAY, 2)
    back_to_SELF                                                       'resetting at each loop
    MAXIS_case_number = EOMC_CLIENT_ARRAY(case_nbr, hc_clt)		       'defining case number for functions to use
    CLIENT_reference_number = EOMC_CLIENT_ARRAY (clt_ref_nbr,  hc_clt) 'setting this to a more usable variable

    ' If MAXIS_case_number = "" Then MsgBox "Case number: " & EOMC_CLIENT_ARRAY(case_nbr, hc_case) & vbNewLine & "Client: " & EOMC_CLIENT_ARRAY (clt_ref_nbr,  hc_clt) & vbNewLines & "PMI: " & EOMC_CLIENT_ARRAY (clt_pmi,   hc_clt)

    Call navigate_to_MAXIS_screen("STAT", "MEMB")                       'going to MEMB to get age for identifying correct capitation
    EMWriteScreen CLIENT_reference_number, 20, 76
    EMReadScreen age_of_client, 3, 8, 76
    age_of_client = trim(age_of_client)
    If age_of_client = "" Then age_of_client = 0
    EOMC_CLIENT_ARRAY(clt_age, hc_clt) = age_of_client*1

    Call navigate_to_MAXIS_screen ("ELIG", "HC")						'Goes to ELIG HC
    APPROVAL_NEEDED = FALSE                                             'setting some booleans
    found_elig = FALSE
    client_found = FALSE
    row = 8                                                             'begining of the list of HH Membs in ELIG/HC
    Do
        EMReadScreen check_for_priv, 10, 24, 14                         'Some cases from the BOBI are high level priv and we cannot look at details
        If check_for_priv = "PRIVILEGED" Then
            EOMC_CLIENT_ARRAY(err_notes, hc_clt) = "PRIV"
            Exit Do
        End If

        EMReadScreen elig_clt, 2, row, 3                                'reading the information on the row to see if it is for the client
        EMReadScreen prog_exists, 1, row, 10
        If elig_clt = CLIENT_reference_number Then                      'If these match, we have found the client to find additional HC details
            'MsgBox "Elig Clt: " & elig_clt & vbNewLine & "Ref Number: " &CLIENT_reference_number
            client_found = TRUE                                         'setting the boolean for the rest to search for more
            EMReadScreen prog, 10, row, 28                              'reading all the program details on ELIG Memb List
            EMReadScreen version, 2, row, 58
            EMReadScreen app_indc, 6, row, 68

            prog = trim(prog)                                           'formatting the information that was read.
            app_indc = trim(app_indc)

            If prog = "NO REQUEST" OR prog = "NO VERSION" OR prog = "" Then                          'If there is no span known for the member, it will be indicated in this way and we can't see any more
                EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ HC information does not appear to be in MAXIS. ELIG/HC for this member - " & prog
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
                EMReadScreen hc_status, 7, row, 50

                prog = trim(prog)
                hc_status = trim(hc_status)

                If hc_status = "ACTIVE" Then
                'If result = "ELIG" AND hc_status = "ACTIVE" Then        'the clients that are eligible and active on ELIG HC - we will look in the HC Summ for more information
                    EOMC_CLIENT_ARRAY (hc_prog_one,   hc_clt) = prog      'setting this to the array

                    EMWriteScreen "X", row, 26                          'opening the HC BSUM
                    transmit

                    If prog = "MA" or prog = "IMD" or prog = "EMA" Then                 'for the programs MA or IMD the information is in a certain place
                        If left(EOMC_CLIENT_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then   'If the name was not on the BOBI and is just listed on X's then we read the actual name here
                            EMReadScreen the_name, 30, 5, 20
                            the_name = trim(the_name)
                            EOMC_CLIENT_ARRAY (clt_name, hc_clt) = the_name
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
                        If prog = "    " Then EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ HC ELIG Budget may need approval or budget needs to be aligned."

                        'Took this out because I think it is for a different script issue with a dirrect run and no longer needed. Will test with it gone for a while'
                        ' If pers_type = "__" Then                        'TODO - REMOVE ON 10/30/18 if no longer needed - determine why I have this here - look at case # 132245
                        '     EMReadScreen cur_mo_test, 6, 7, mo_col
                        '     cur_mo_test = trim(cur_mo_test)
                        '     'MsgBox "This is come up when person test is __" & vbNewLine & "cur_mo_test is " & cur_mo_test
                        '     pers_type = cur_mo_test
                        '     pers_std = ""
                        '     pers_mthd = ""
                        ' End If

                        EOMC_CLIENT_ARRAY (elig_type_one, hc_clt) = pers_type     'setting all of the read information is added to the array

                        'if this was found to be true in this loop, will add error note that the case needs review and approval
                        If APPROVAL_NEEDED = TRUE THen EOMC_CLIENT_ARRAY (err_notes, hc_clt) = EOMC_CLIENT_ARRAY (err_notes, hc_clt) & " ~ Budget Needs Approval"

                    Else                                                            'this is for programs other than MA or IMD - typically QMB, SLMB, or QI
                        If left(EOMC_CLIENT_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then       'for some clients that don't have an actual name
                            EMReadScreen the_name, 30, 5, 15
                            the_name = trim(the_name)
                            EOMC_CLIENT_ARRAY (clt_name, hc_clt) = the_name
                        End If
                        EMReadScreen pers_type, 2, 6, 56                                'reading the type and standard
                        EMReadScreen pers_std, 1, 6, 64

                        EOMC_CLIENT_ARRAY (hc_prog_one,   hc_clt) = prog          'adding this to the array

                        EOMC_CLIENT_ARRAY (elig_type_one, hc_clt) = pers_type
                    End If
                    PF3
                End If
            End If

            Do                                              'this is after the first program is listed, there may be a second program
                row = row + 1                               'looking at the next row

                EMReadScreen next_client_ref, 2, row, 3     'reading the reference number and program'
                EMReadScreen next_prog, 4, row, 28

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
                    EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ HC information does not appear to be in MAXIS. ELIG/HC for this member - " & prog
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

                EMReadScreen hc_status, 7, row, 50
                hc_status = trim(hc_status)

                If hc_status <> "ACTIVE" Then found_elig = FALSE

                If found_elig = TRUE Then                                       'this was set in the code above.
                    EMWriteScreen "X", row, 26                                  'opening BSUM
                    transmit                                                    'we don't need to determine program because a second programs is always medicare savings progs

                    If left(EOMC_CLIENT_ARRAY (clt_name, hc_clt), 5) = "XXXXX" Then       'finding the correct name if the case is priv but I have access
                        EMReadScreen the_name, 30, 5, 15
                        the_name = trim(the_name)
                        EOMC_CLIENT_ARRAY (clt_name, hc_clt) = the_name
                    End If

                    EMReadScreen pers_type, 2, 6, 56                            'reading the type and standard
                    EMReadScreen pers_std, 1, 6, 64

                    If EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) <> "" Then      'this adds it to the array after determining WHICH part it belongs in
                        EOMC_CLIENT_ARRAY (hc_prog_two,   hc_clt) = prog

                    Else
                        EOMC_CLIENT_ARRAY (hc_prog_one,   hc_clt) = prog

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
                If client_found = FALSE Then EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ Member number not found on ELIG/HC."
                Exit Do
            End If
        End If
    Loop until client_found = TRUE
    If EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) <> "" Then EOMC_CLIENT_ARRAY(prog_one_end, hc_clt) = last_day_of_this_month
    If EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) <> "" Then EOMC_CLIENT_ARRAY(prog_two_end, hc_clt) = last_day_of_this_month
    EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = 0
    Call keep_MMIS_passworded_in
Next

'need to get to ground zero
Call back_to_SELF
Call navigate_to_spec_MMIS_region("CTY ELIG STAFF/UPDATE")      'Going to MMIS'

'Looping through each of the HC clients while in MMIS
For hc_clt = 0 to UBOUND(EOMC_CLIENT_ARRAY, 2)
    STATS_counter = STATS_counter + 1       'incrementing for each client HC reviewed - it is here because this is the part that the timer will cut out on
    PMI_Number = right("00000000" & EOMC_CLIENT_ARRAY(clt_pmi, hc_clt), 8)    'making this 8 charactes because MMIS
    MAXIS_case_number = right("00000000" & EOMC_CLIENT_ARRAY(case_nbr, hc_clt), 8)

    If EOMC_CLIENT_ARRAY(err_notes, hc_clt) <> "PRIV" Then                  'Can't look at priv case information so we will ignore them
        EMWriteScreen "I", 2, 19                                                    'read only
        EMWriteScreen PMI_Number, 4, 19                                             'enter through the PMI so it isn't case specific
        transmit

        EMWriteScreen "RELG", 1, 8                  'go to RELG where all the elig detail is
        transmit
        'MsgBox "To RELG"

        relg_row = 6                                'beginning of the list.
        span_found = FALSE                          'setting this for each client loop
        EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) = 1
        Do
            EMReadScreen relg_prog, 2, relg_row, 10 'reading the prog and elig type information
            EMReadScreen relg_elig, 2, relg_row, 33
            EMReadScreen relg_case_num, 8, relg_row, 73 'reading the case number for this span

            If relg_case_num = MAXIS_case_number Then       'only look at a SPAN if it is for the right case number
                'If the program matches and the elig type matches we will read for an end date
                If relg_prog = left(EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt), 2) AND relg_elig = EOMC_CLIENT_ARRAY(elig_type_one, hc_clt) Then
                    span_found = TRUE           'setting this for later/next loop
                    EMReadScreen relg_end_dt, 8, relg_row+1, 36     'this is where the end date is
                    'MsgBox "End Date - " & relg_end_dt
                    EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = relg_end_dt     'setting the end date in to the array
                    EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = relg_row
                ElseIf relg_prog = "EH" AND EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "EMA" Then
                    span_found = TRUE           'setting this for later/next loop
                    EMReadScreen relg_end_dt, 8, relg_row+1, 36     'this is where the end date is
                    'MsgBox "End Date - " & relg_end_dt
                    EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = relg_end_dt     'setting the end date in to the array
                    EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = relg_row
                ElseIf relg_prog = "SL" AND EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "QI1" Then
                    If relg_elig = EOMC_CLIENT_ARRAY(elig_type_one, hc_clt) Then
                        span_found = TRUE           'setting this for later/next loop
                        EMReadScreen relg_end_dt, 8, relg_row+1, 36     'this is where the end date is
                        'MsgBox "End Date - " & relg_end_dt
                        EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = relg_end_dt     'setting the end date in to the array
                        EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = relg_row
                    Else
                        EMReadScreen relg_end_dt, 8, relg_row+1, 36     'this is where the end date is
                        EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = relg_end_dt         'adding it to the array and adding a message about the wrong elig type
                        EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ MMIS SPAN for " & EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) & " has the wrong ELIG TYPE"
                        span_found = TRUE
                        EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = relg_row
                    End If
                ElseIf relg_prog = left(EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt), 2) Then       'sometimes the program matches but the elig type does not - HC is still active in MMIS but wrong
                    EMReadScreen relg_end_dt, 8, relg_row+1, 36     'this is where the end date is
                    EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = relg_end_dt         'adding it to the array and adding a message about the wrong elig type
                    EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ MMIS SPAN for " & EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) & " has the wrong ELIG TYPE"
                    span_found = TRUE
                    EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = relg_row
                End If
            End If
            'Once PROG is blank - there are no more spans to review
            If relg_prog = "  " Then Exit Do
            relg_row = relg_row + 4         'next span on RELG'
            If relg_row = 22 Then           'this is the end of RELG and we need to go to a new page
                PF8
                relg_row = 6
                EMReadScreen end_of_list, 7, 24, 26     'This is the end of the list
                If end_of_list = "NO MORE" Then Exit Do
                EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) = EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) + 1
            End If

            ' If span_found = TRUE Then
            '     Msgbox "Found HC in MMIS" & vbNewLine & "On Row: " & EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) & vbNewLine & EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) & " is ended on: " & EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt)
            ' End If

        Loop until span_found = TRUE
        'If we exited before finding the right Span then an error is added that a span does not exist.

        EMWriteScreen "RELG", 1, 8      'This takes us back to the top in case we had to PF8 down'
        transmit
        'MsgBox "To RELG"

        'If there is a second program for this client, we are goind to do it all over again.
        If EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) <> "" Then
            relg_row = 6                'top of the list of Spans
            span_found = FALSE          'reset this for the next program
            Do
                EMReadScreen relg_prog, 2, relg_row, 10     'reading program and elig type'
                EMReadScreen relg_elig, 2, relg_row, 33
                EMReadScreen relg_case_num, 8, relg_row, 73
                'MsgBox "2 - " & relg_prog & " - " & relg_elig

                If relg_case_num = MAXIS_case_number Then
                    'if both match, getting th end date
                    If relg_prog = left(EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt), 2) AND relg_elig = EOMC_CLIENT_ARRAY(elig_type_two, hc_clt) Then
                        span_found = TRUE
                        EMReadScreen relg_end_dt, 8, relg_row+1, 36                 'reading the end date
                        'MsgBox "2 - End Date - " & relg_end_dt
                        EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = relg_end_dt 'setting it to the array
                        EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = relg_row
                    ElseIf relg_prog = left(EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt), 2) Then       'if only the program matches
                        EMReadScreen relg_end_dt, 8, relg_row+1, 36                 'reading the end date

                        EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = relg_end_dt         'adding it to the array and adding a message about the wrong elig type
                        EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ MMIS SPAN for " & EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) & " has the wrong ELIG TYPE"
                        span_found = TRUE
                        EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = relg_row
                    End If
                End If

                If relg_prog = "  " Then Exit Do            'leaving the loop if we are at the end of the RELG list'
                relg_row = relg_row + 4                     'next span
                If relg_row = 22 Then           'this is the end of RELG and we need to go to a new page
                    PF8
                    relg_row = 6
                    EMReadScreen end_of_list, 7, 24, 26     'This is the end of the list
                    If end_of_list = "NO MORE" Then Exit Do
                    EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) = EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) + 1
                End If

                ' 'for testing
                ' If span_found = TRUE Then
                '     Msgbox "Found HC in MMIS" & vbNewLine & "On Row: " & EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) & vbNewLine & EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) & " is ended on: " & EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt)
                ' End If

            Loop until span_found = TRUE
            'adding a message if no span was found for this program
            If span_found = FALSE Then EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ No MMIS SPAN for " & EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt)
        End If

        EMWriteScreen "RKEY", 1, 8  'back to the beginning for the next client/loop'
        transmit

    End If
    'if for some reason no HC programs were in MAXIS to begin with - adding this detail to the message
    If EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "" AND EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "" Then EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ No HC Programs found in MAXIS ELIG."

    'this block updated MMIS with the new end date if the change option was selected at the beginning
    If make_changes = TRUE Then
        If EOMC_CLIENT_ARRAY(autoclose, hc_clt) = FALSE Then                                'autoclose cases should close on their own.
            If EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = "99/99/99" Then               'if the span has an open end date
                EMWriteScreen "C", 2, 19                                                    'going in to change
                EMWriteScreen PMI_Number, 4, 19                                             'enter through the PMI because navigation is easier
                transmit

                EMWriteScreen "RELG", 1, 8                  'go to RELG where all the elig detail is
                transmit

                If EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) > 1 Then                        'we saved the RELG page so going back to it
                    for forward = 2 to EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt)
                        PF8
                    next
                End If

                'determine where the SPAN is
                If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 18 Then            'if the span is the last on the page
                    PF8                                                         'it is now on the next page
                    relg_row = 10                                               'the top span is empty in change so the known spans start at 10
                Else                                                            'if it wasn't the last on the page
                    relg_row = EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) + 4      'since the top row is empty, the row is 4 down from where it was in inquiry
                End If

                EMWriteScreen mmis_last_day_date, relg_row+1, 36                'entering the last day of the current month to the end date on the span
                EMWriteScreen "C", relg_row+1, 62                               'updating status to 'closed'

                PF3                                     'save and check for warning message
                EMReadScreen warn_msg, 7, 24, 2
                If warn_msg = "WARNING" Then PF3

                PF3                                     'save and go back to RKEY
                EMWriteScreen "X",8, 3
                transmit
                'NOW THE INFORMATION IS SAVED'

                'We are going to confirm the information
                EMWriteScreen "I", 2, 19                    'read only
                EMWriteScreen PMI_Number, 4, 19             'enter through the PMI so it isn't case specific
                transmit

                EMReadScreen pph_end_date, 8, 16, 37        'looking at the most recent capitation end date

                EMWriteScreen "RELG", 1, 8                  'go to RELG where all the elig detail is
                transmit

                If EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) > 1 Then            'now getting to the right page and row to read the span
                    for forward = 2 to EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt)
                        PF8
                    next
                End If

                relg_row = EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt)

                EMReadScreen confirm_mmis_end, 8, relg_row+1, 36        'reading the end date and status to make sure the change was successful
                EMReadScreen confirm_mmis_stat, 1, relg_row+1, 62

                If confirm_mmis_end = mmis_last_day_date AND confirm_mmis_stat = "C" Then   'if they match, we will update the array information

                    EOMC_CLIENT_ARRAY(MMIS_new_end_one, hc_clt) = mmis_last_day_date        'This adds the last day of the month to the new MMIS end date

                    'This part reviews for capitation savings
                    If EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "MA" OR EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "IMD" OR EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "EMA" Then

                        If pph_end_date = mmis_last_day_date Then       'if a PPH span is ending the same date, then the closure is caused by our update
                            'the savings are added to the client array here
                            If EOMC_CLIENT_ARRAY(elig_type_one, hc_clt) = "PX" OR EOMC_CLIENT_ARRAY(elig_type_one, hc_clt) = "PC" Then
                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_PW
                            Else
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 1 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_11x
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) = 1 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_1
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 1 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 16 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_2_15
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 15 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 21 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_16_20
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 20 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 50 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_21_49
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 49 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 65 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_50_64
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 65 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_65
                            ENd If

                            EOMC_CLIENT_ARRAY(capitation_ended, hc_clt) = TRUE      'setting the capitation ending as true
                        Else
                            EOMC_CLIENT_ARRAY(capitation_ended, hc_clt) = FALSE     'if these don't match, then a capitation was not ended.
                        End If

                    'for non-MA cases the savings is based on the medicare premium
                    ElseIf EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "QI1" Then
                        EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_QI1
                    ElseIf EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "QMB" Then
                        EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_QMB
                    ElseIf EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "SLMB" Then
                        EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_SLMB
                    End If
                End If

                PF6     'backing out of the MMIS client information

            End If

            If EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = "99/99/99" Then               'if program 2 is open ended
                EMWriteScreen "C", 2, 19                                                    'entering change option
                EMWriteScreen PMI_Number, 4, 19                                             'enter through the PMI because navigation is easier
                transmit

                EMWriteScreen "RELG", 1, 8                  'go to RELG where all the elig detail is
                transmit

                If EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) > 1 Then            'we saved the row the span was found at - going back there
                    for forward = 2 to EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt)
                        PF8
                    next
                End If

                'determine where the SPAN is
                If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 18 Then            'if the span is the last on the page
                    PF8                                                         'it is now on the next page
                    relg_row = 10                                               'the top span is empty in change so the known spans start at 10
                Else                                                            'if it wasn't the last on the page
                    relg_row = EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) + 4      'since the top row is empty, the row is 4 down from where it was in inquiry
                End If

                EMWriteScreen mmis_last_day_date, relg_row+1, 36                'entering the last day of the current month and changing status to closed
                EMWriteScreen "C", relg_row+1, 62

                PF3                                     'saving, checking for warning message and saving again
                EMReadScreen warn_msg, 7, 24, 2
                If warn_msg = "WARNING" Then PF3

                PF3                                     'saving all the way, then going back to RKEY
                EMWriteScreen "X",8, 3
                transmit
                'NOW THE INFORMATION IS SAVED'

                'We are going to confirm the information
                EMWriteScreen "I", 2, 19                                                    'read only
                EMWriteScreen PMI_Number, 4, 19                                             'enter through the PMI so it isn't case specific
                transmit

                EMReadScreen pph_end_date, 8, 16, 37

                EMWriteScreen "RELG", 1, 8                  'go to RELG where all the elig detail is
                transmit

                If EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) > 1 Then
                    for forward = 2 to EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt)
                        PF8
                    next
                End If

                relg_row = EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt)

                EMReadScreen confirm_mmis_end, 8, relg_row+1, 36
                EMReadScreen confirm_mmis_stat, 1, relg_row+1, 62

                If confirm_mmis_end = mmis_last_day_date AND confirm_mmis_stat = "C" Then

                    EOMC_CLIENT_ARRAY(MMIS_new_end_two, hc_clt) = mmis_last_day_date

                    If EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "MA" OR EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "IMD" OR EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "EMA" Then

                        If pph_end_date = mmis_last_day_date Then
                            If EOMC_CLIENT_ARRAY(elig_type_two, hc_clt) = "PX" OR EOMC_CLIENT_ARRAY(elig_type_two, hc_clt) = "PC" Then
                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_PW
                            Else
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 1 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_11x
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) = 1 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_1
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 1 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 16 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_2_15
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 15 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 21 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_16_20
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 20 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 50 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_21_49
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 49 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 65 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_50_64
                                If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 65 Then EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_65
                            ENd If

                            EOMC_CLIENT_ARRAY(capitation_ended, hc_clt) = TRUE
                        Else
                            EOMC_CLIENT_ARRAY(capitation_ended, hc_clt) = FALSE
                        End If

                    ElseIf EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "QI1" Then
                        EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_QI1
                    ElseIf EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "QMB" Then
                        EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_QMB
                    ElseIf EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "SLMB" Then
                        EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_SLMB
                    End If
                End If

                PF6

                'MsgBox "To RELG"
            End If
        End If
    Else
        'If we are not changing - we are still going to look for capitation end information
        EMWriteScreen "I", 2, 19                    'read only
        EMWriteScreen PMI_Number, 4, 19             'enter through the PMI so it isn't case specific
        transmit

        EMReadScreen pph_end_date, 8, 16, 37        'looking at the most recent capitation end date

        If pph_end_date = mmis_last_day_date Then
            EOMC_CLIENT_ARRAY(capitation_ended, hc_clt) = TRUE
        Else
            EOMC_CLIENT_ARRAY(capitation_ended, hc_clt) = FALSE
        End If

        PF6
    End If

    total_savings = total_savings + EOMC_CLIENT_ARRAY(clt_savings, hc_clt)      'adding client savings to the total savings for the script run.
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

ObjExcel.Cells(1, col_to_use).Value = "EOMC Status"
eomc_stat_col = col_to_use
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

ObjExcel.Cells(1, col_to_use).Value = "MAXIS End Date"
MAXIS_end_one_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "CURR MMIS End Date"
mmis_one_col = col_to_use
col_to_use = col_to_use + 1

If make_changes = TRUE Then
    ObjExcel.Cells(1, col_to_use).Value = "NEW MMIS End Date"
    new_mmis_one_col = col_to_use
    col_to_use = col_to_use + 1

    ObjExcel.Cells(1, col_to_use).Value = "PPH Cap"
    cap_col = col_to_use
    col_to_use = col_to_use + 1
End If

ObjExcel.Cells(1, col_to_use).Value = "2nd PROG"
prog_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "ELIG TYPE"
elig_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "MAXIS End Date"
MAXIS_end_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "CURR MMIS End Date"
mmis_two_col = col_to_use
col_to_use = col_to_use + 1

If make_changes = TRUE Then
    ObjExcel.Cells(1, col_to_use).Value = "NEW MMIS End Date"
    new_mmis_two_col = col_to_use
    col_to_use = col_to_use + 1

    ObjExcel.Cells(1, col_to_use).Value = "SAVINGS"
    savings_col = col_to_use
    col_to_use = col_to_use + 1
End If

ObjExcel.Cells(1, col_to_use).Value = "ERRORS"
errors_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Rows(1).Font.Bold = TRUE
excel_row = 2

'Adding all client information to a spreadsheet for your viewing pleasure
For hc_clt = 0 to UBound(EOMC_CLIENT_ARRAY, 2)
	ObjExcel.Cells(excel_row, worker_col).Value        = EOMC_CLIENT_ARRAY (basket_nbr,   hc_clt)
	ObjExcel.Cells(excel_row, case_numb_col).Value     = EOMC_CLIENT_ARRAY (case_nbr,  hc_clt)
    ObjExcel.Cells(excel_row, eomc_stat_col).Value     = EOMC_CLIENT_ARRAY (hc_close_stat, hc_clt)
	ObjExcel.Cells(excel_row, ref_numb_col).Value      = "Memb " & EOMC_CLIENT_ARRAY(clt_ref_nbr, hc_clt)
	ObjExcel.Cells(excel_row, name_col).Value          = EOMC_CLIENT_ARRAY (clt_name,  hc_clt)
	ObjExcel.Cells(excel_row, pmi_col).Value           = EOMC_CLIENT_ARRAY (clt_pmi,   hc_clt)

    ObjExcel.Cells(excel_row, prog_one_col).Value       = EOMC_CLIENT_ARRAY (hc_prog_one,   hc_clt)
    ObjExcel.Cells(excel_row, elig_one_col).Value       = EOMC_CLIENT_ARRAY (elig_type_one,   hc_clt)
    ObjExcel.Cells(excel_row, MAXIS_end_one_col).Value  = EOMC_CLIENT_ARRAY(prog_one_end, hc_clt)
    ObjExcel.Cells(excel_row, mmis_one_col).Value       = EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt)

    ObjExcel.Cells(excel_row, prog_two_col).Value       = EOMC_CLIENT_ARRAY (hc_prog_two,   hc_clt)
    ObjExcel.Cells(excel_row, elig_two_col).Value       = EOMC_CLIENT_ARRAY (elig_type_two,   hc_clt)
    ObjExcel.Cells(excel_row, MAXIS_end_two_col).Value  = EOMC_CLIENT_ARRAY(prog_two_end, hc_clt)
    ObjExcel.Cells(excel_row, mmis_two_col).Value       = EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt)

    ObjExcel.Cells(excel_row, errors_col).Value         = EOMC_CLIENT_ARRAY(err_notes, hc_clt)

    If make_changes = TRUE Then
        ObjExcel.Cells(excel_row, new_mmis_one_col).Value   = EOMC_CLIENT_ARRAY(MMIS_new_end_one, hc_clt)
        ObjExcel.Cells(excel_row, cap_col).Value            = EOMC_CLIENT_ARRAY(capitation_ended, hc_clt)

        ObjExcel.Cells(excel_row, new_mmis_two_col).Value   = EOMC_CLIENT_ARRAY(MMIS_new_end_two, hc_clt)
        ObjExcel.Cells(excel_row, savings_col).Value        = EOMC_CLIENT_ARRAY(clt_savings, hc_clt)
        ObjExcel.Cells(excel_row, savings_col).NumberFormat = "$#,##0.00"
    End If
	excel_row = excel_row + 1      'next row
Next

col_to_use = col_to_use + 1     'moving over one extra for script run details.

'Query date/time/runtime info
objExcel.Cells(2, col_to_use).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use).Value = "Query date and time:"
ObjExcel.Cells(1, col_to_use+1).Value = now
ObjExcel.Cells(1, col_to_use+1).Font.Bold = FALSE
ObjExcel.Cells(2, col_to_use).Value = "Query runtime (in seconds):"
ObjExcel.Cells(2, col_to_use+1).Value = timer - query_start_time
ObjExcel.Cells(3, col_to_use).Value = "Total Savings:"
ObjExcel.Cells(3, col_to_use+1).Value = total_savings
ObjExcel.Cells(3, col_to_use+1).NumberFormat = "$#,##0.00"

'Autofitting columns
For col_to_autofit = 1 to col_to_use+1
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'setting a freeze row for easy scrolling
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

script_end_procedure("All Done")
