'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - RESOLVE HC EOMC IN MMIS.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 160                       'manual run time in seconds
STATS_denomination = "M"       							'C is for each CASE
'END OF stats block==============================================================================================
' run_locally = TRUE
'TODO Figure out what is going on with the funclib on this - I need to run locally ?!'
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
call changelog_update("06/24/2019", "Moved the Excel update to within the change loop in case of a script run failure.", "Casey Love, Hennepin County")
call changelog_update("05/24/2019", "Added statistics detail to final script run when MMIS is updated.", "Casey Love, Hennepin County")
call changelog_update("02/19/2019", "Changed Medicare savings amounts to $135.50.", "Casey Love, Hennepin County")
call changelog_update("11/21/2018", "Removed custom function navigate_to_spec_MMIS_region(group_security_selection). Added test function navigate_to_MAXIS_test. Updated function 'keep_MMIS_passworded_in()' that calls these functions.", "Ilse Ferris, Hennepin County")
call changelog_update("09/21/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'function specific to this script - running_stopwatch and MX_environment are defined outside of this function
'meant to keep MMIS from passwording out while this long bulk script is running
function keep_MMIS_passworded_in(mmis_area, maxis_area)
    ' MsgBox running_stopwatch & vbNewLine & timer
    If timer - running_stopwatch > 720 Then         'this means the script has been running for more than 12 minutes since we last popped in to MMIS
        Call navigate_to_MMIS_region(mmis_area)      'Going to MMIS'
        'MsgBox "In MMIS"
        Call navigate_to_MAXIS(maxis_area)                       'going back to MAXIS'
        'MsgBox "Back to MAXIS"
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
first_of_month_after = CM_plus_2_mo & "/1/" & CM_plus_2_yr
last_day_of_month_after = DateAdd("d", -1, first_of_month_after)

'formatting the last day of the month in to MM/DD/YY for entry in to MMIS
last_day_mo = DatePart("m", last_day_of_this_month)
last_day_mo = right("00" & last_day_mo, 2)
last_day_day = DatePart("d", last_day_of_this_month)
last_day_day = right("00" & last_day_day, 2)
last_day_yr = DatePart("yyyy", last_day_of_this_month)
last_day_yr = right("00" & last_day_yr, 2)
mmis_last_day_date = last_day_mo & "/" & last_day_day & "/" & last_day_yr

last_day_mo = DatePart("m", last_day_of_month_after)
last_day_mo = right("00" & last_day_mo, 2)
last_day_day = DatePart("d", last_day_of_month_after)
last_day_day = right("00" & last_day_day, 2)
last_day_yr = DatePart("yyyy", last_day_of_month_after)
last_day_yr = right("00" & last_day_yr, 2)
mmis_last_day_after_cap = last_day_mo & "/" & last_day_day & "/" & last_day_yr


'Setting amounts
total_savings = 0                   'setting this at zero so that we can add up what we save
capitation_11x      = 1051.66'938.14        'capitation amounts set annually by DHS - eventually we need to move this to FuncLib
capitation_PW       = 791.45'1241.53
capitation_1        = 256.59'235.68
capitation_2_15     = 256.57'236.03
capitation_16_20    = 282.16'261.67
capitation_21_49    = 932.53'808.96
capitation_21_49_ax = 996.84'800.89
capitation_50_64    = 1260.16'1118.34
capitation_50_64_ax = 988.53'801.27
capitation_65       = 3094.54'2681.89

capitation_QMB      = 164.90'135.50
capitation_SLMB     = 164.90'135.50
capitation_QI1      = 164.90'135.50

Const xlSrcRange = 1
Const xlYes = 1

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
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 351, 75, "Start Resolve HC EOMC in MMIS Run"
  EditBox 90, 10, 255, 15, list_of_workers
  CheckBox 10, 45, 140, 10, "Check here to have script update MMIS", change_checkbox
  CheckBox 10, 60, 130, 10, "Check here if it is after CAPITATION.", after_capitation_checkbox
  ButtonGroup ButtonPressed
    OkButton 240, 55, 50, 15
    CancelButton 295, 55, 50, 15
  Text 5, 15, 85, 10, "List of Workers to check:"
  Text 95, 30, 125, 10, "(Leave blank to run on entrire county)"
EndDialog

'Showing the dialog
Do
    Dialog Dialog1
    cancel_without_confirmation

    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checking for MAXIS
Call check_for_MAXIS(True)
Call back_to_SELF                                               'starting at the SELF panel
EMReadScreen MX_environment, 13, 22, 48                         'seeing which MX environment we are in
MX_environment = trim(MX_environment)
Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")        'Going to MMIS'
Call navigate_to_MAXIS(MX_environment)                          'going back to MAXIS
running_stopwatch = timer               'setting the running timer so we log in to MMIS within every 15 mintues so we don't password out

make_changes = FALSE                    'setting this at the start
If change_checkbox = checked Then make_changes = TRUE   'if the dialog has indicated that changes should be changed reset this to true
If after_capitation_checkbox = checked Then mmis_last_day_date = mmis_last_day_after_cap
' MsgBox "MMIS last day to update to - " & mmis_last_day_date

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
	Call navigate_to_MAXIS_screen("REPT", "EOMC")      'go to EOMC for the correct worker number
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
    Call keep_MMIS_passworded_in("CTY ELIG STAFF/UPDATE", MX_environment)                'every 12 mintues or so, the script will pop in to MMIS to make sure we are passworded in
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
        'MsgBox clt_hc_status
        If clt_hc_status = "A" Then                             'if HC is active then we will add this client to the array to find additional information
            'MsgBox EOMC_CASES_ARRAY(autoclose, hc_case)
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

                'mSGbOX MAXIS_case_number
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

    Call keep_MMIS_passworded_in("CTY ELIG STAFF/UPDATE", MX_environment)        'making sure we are not passworded out in MMIS
Next

'Now the client array is created
'Information gathering in MAXIS now for every client on HC on the list
For hc_clt = 0 to UBOUND(EOMC_CLIENT_ARRAY, 2)
    back_to_SELF                                                       'resetting at each loop
    MAXIS_case_number = EOMC_CLIENT_ARRAY(case_nbr, hc_clt)		       'defining case number for functions to use
    CLIENT_reference_number = EOMC_CLIENT_ARRAY (clt_ref_nbr,  hc_clt) 'setting this to a more usable variable

    ' If MAXIS_case_number = "" Then MsgBox "Case number: " & EOMC_CLIENT_ARRAY(case_nbr, hc_case) & vbNewLine & "Client: " & EOMC_CLIENT_ARRAY (clt_ref_nbr,  hc_clt) & vbNewLines & "PMI: " & EOMC_CLIENT_ARRAY (clt_pmi,   hc_clt)

    Call navigate_to_MAXIS_screen("STAT", "MEMB")                       'going to MEMB to get age for identifying correct capitation
    Call write_value_and_transmit(CLIENT_reference_number, 20, 76)
	EMWaitReady 0, 0
	EMReadScreen access_denied_check, 13, 24, 2
	If access_denied_check = "ACCESS DENIED" Then
		PF10
        EMWaitReady 0, 0
		EOMC_CLIENT_ARRAY(clt_age, hc_clt) = 2 'TODO - decide what 'age' we want to assign when a MEMB panel is access denied
	Else
		EMReadScreen age_of_client, 3, 8, 76
		age_of_client = trim(age_of_client)
		If age_of_client = "" Then age_of_client = 0
		EOMC_CLIENT_ARRAY(clt_age, hc_clt) = age_of_client*1
	End If

    Call navigate_to_MAXIS_screen ("ELIG", "HC  ")						'Goes to ELIG HC
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
    Call keep_MMIS_passworded_in("CTY ELIG STAFF/UPDATE", MX_environment)
Next

'need to get to ground zero
Call back_to_SELF
Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")      'Going to MMIS'
'MsgBox "Pause"
EMReadScreen check_in_MMIS, 18, 1, 7

If check_in_MMIS = "SESSION TERMINATED" Then
    EMWriteScreen "MW00",1, 2
    transmit
    transmit

    EMWriteScreen "X", 8, 3
    transmit
End If


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
    new_mmis_one_col_letter_col = convert_digit_to_excel_column(new_mmis_one_col)
    col_to_use = col_to_use + 1

    ObjExcel.Cells(1, col_to_use).Value = "PPH Cap"
    cap_col = col_to_use
    col_to_use = col_to_use + 1
End If

ObjExcel.Cells(1, col_to_use).Value = "2nd PROG"
prog_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "ELIG TYPE - 2"
elig_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "MAXIS End Date - 2"
MAXIS_end_two_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Cells(1, col_to_use).Value = "CURR MMIS End Date - 2"
mmis_two_col = col_to_use
mmis_two_letter_col = convert_digit_to_excel_column(mmis_two_col)
col_to_use = col_to_use + 1

If make_changes = TRUE Then
    ObjExcel.Cells(1, col_to_use).Value = "NEW MMIS End Date - 2"
    new_mmis_two_col = col_to_use
    new_mmis_two_col_letter_col = convert_digit_to_excel_column(new_mmis_two_col)
    col_to_use = col_to_use + 1

    ObjExcel.Cells(1, col_to_use).Value = "SAVINGS"
    savings_col = col_to_use
    savings_letter_col = convert_digit_to_excel_column(savings_col)
    col_to_use = col_to_use + 1
End If

ObjExcel.Cells(1, col_to_use).Value = "ERRORS"
errors_col = col_to_use
col_to_use = col_to_use + 1

ObjExcel.Rows(1).Font.Bold = TRUE
excel_row = 2

'Looping through each of the HC clients while in MMIS
For hc_clt = 0 to UBOUND(EOMC_CLIENT_ARRAY, 2)
    Call get_to_RKEY
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
            EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) = 1
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

        call get_to_RKEY
    End If
    'if for some reason no HC programs were in MAXIS to begin with - adding this detail to the message
    If EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "" AND EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "" Then EOMC_CLIENT_ARRAY(err_notes, hc_clt) = EOMC_CLIENT_ARRAY(err_notes, hc_clt) & " ~ No HC Programs found in MAXIS ELIG."

    'this block updated MMIS with the new end date if the change option was selected at the beginning
    If make_changes = TRUE Then
        If EOMC_CLIENT_ARRAY(autoclose, hc_clt) = FALSE Then                                'autoclose cases should close on their own.
            ' Update_two = FALSE
            ' have_to_go_to_next_page = FALSE
            ' If EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = "99/99/99" AND EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = "99/99/99" Then
            '     Update_two = TRUE
            '     If EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) > 1 OR
            '         have_to_go_to_next_page = TRUE
            '
            '     End If
            '     If EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) > 1 Then
            '         have_to_go_to_next_page = TRUE
            '
            '     End If
            ' End If
            If EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = "99/99/99" or EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = "99/99/99" Then               'if the span has an open end date
                EMWriteScreen "C", 2, 19                                                    'going in to change
                EMWriteScreen PMI_Number, 4, 19                                             'enter through the PMI because navigation is easier
                ' MsgBox "Going in to CHANGE" & vbNewLine & "Prog 1 - " & EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) & " END - " & EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) & vbNewLine & "PROG 2 - " &  EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) & " END - " & EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt)
                transmit

                EMWriteScreen "RELG", 1, 8                  'go to RELG where all the elig detail is
                transmit

                prog_one_order = 0
                prog_two_order = 0
                on_page = 1

                'THIS IS MY NEW CODE TO TRY THE THINGS
                If EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = "99/99/99" Then
                    ' MsgBox "PROG 1" & vbNewLine & "PAGE - " & EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) & vbNewLine & "ROW - " & EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt)
                    If EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) = 1 Then
                        If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 6 Then prog_one_order = 1
                        If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 10 Then prog_one_order = 2
                        If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 14 Then prog_one_order = 3
                        If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 18 Then prog_one_order = 4
                    ElseIf EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) = 2 Then
                        If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 6 Then prog_one_order = 5
                        If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 10 Then prog_one_order = 6
                        If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 14 Then prog_one_order = 7
                        If EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt) = 18 Then prog_one_order = 8
                    End If
                End If

                If EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = "99/99/99" Then
                    ' MsgBox "PROG 2" & vbNewLine & "PAGE - " & EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) & vbNewLine & "ROW - " & EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt)
                    If EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) = 1 Then
                        If EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = 6 Then prog_two_order = 1
                        If EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = 10 Then prog_two_order = 2
                        If EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = 14 Then prog_two_order = 3
                        If EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = 18 Then prog_two_order = 4
                    ElseIf EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) = 2 Then
                        If EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = 6 Then prog_two_order = 5
                        If EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = 10 Then prog_two_order = 6
                        If EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = 14 Then prog_two_order = 7
                        If EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt) = 18 Then prog_two_order = 8
                    End If
                End If

                ' MsgBox "PROG one order - " & prog_one_order & vbNewLine & "PROG two order - " & prog_two_order

                For each_prog = 8 to 1 Step -1

                    If each_prog = prog_one_order OR each_prog = prog_two_order Then
                        Select Case each_prog

                            Case 1
                                relg_row = 10
                            Case 2
                                relg_row = 14
                            Case 3
                                relg_row = 18
                            Case 4
                                relg_row = 10
                            Case 5
                                relg_row = 14
                            Case 6
                                relg_row = 18
                            Case 7
                                relg_row = 10
                            Case 8
                                relg_row = 14

                        End Select
                        If each_prog > 6 Then
                            Do While on_page <> 3
                                PF8
                                on_page = on_page + 1
                            Loop
                        ElseIf each_prog > 3 and each_prog < 7 Then
                            If on_page > 2 Then
                                PF7
                                on_page = on_page - 1
                            ElseIf on_page < 2 Then
                                PF8
                                on_page = on_page + 1
                            End If
                        Else
                            Do While on_page <> 1
                                PF7
                                on_page = on_page - 1
                            Loop

                        End If

                        EMWriteScreen mmis_last_day_date, relg_row+1, 36                'entering the last day of the current month and changing status to closed
                        EMWriteScreen "C", relg_row+1, 62

                        ' MsgBox "ORDER - " & each_prog & vbNewLine & "ROW - " & relg_row

                    End If
                Next

                PF3                                     'save and check for warning message
                EMReadScreen warn_msg, 7, 24, 2
                If warn_msg = "WARNING" Then PF3

                EmReadscreen check_for_RKEY, 4, 1, 52
                If check_for_RKEY <> "RKEY" Then PF6


                PF3                                     'save and go back to RKEY
                'NOW THE INFORMATION IS SAVED'
                Call get_to_RKEY

                'We are going to confirm the information
                EMWriteScreen "I", 2, 19                    'read only
                EMWriteScreen PMI_Number, 4, 19             'enter through the PMI so it isn't case specific
                transmit

                EMReadScreen pph_end_date, 8, 16, 37        'looking at the most recent capitation end date

                EMWriteScreen "RELG", 1, 8                  'go to RELG where all the elig detail is
                transmit

                If EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = "99/99/99" Then
                    If EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt) > 1 Then            'now getting to the right page and row to read the span
                        for forward = 2 to EOMC_CLIENT_ARRAY(RELG_page_one, hc_clt)
                            PF8
                        next
                    End If

                    relg_row = EOMC_CLIENT_ARRAY(RELG_row_one, hc_clt)

                    EMReadScreen relg_prog, 2, relg_row, 10 'reading the prog and elig type information
                    EMReadScreen relg_elig, 2, relg_row, 33
                    EMReadScreen relg_case_num, 8, relg_row, 73 'reading the case number for this span

                    If relg_prog = "EH" Then relg_prog = "EMA"
                    IF EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "QI1" AND relg_prog = "SL" Then relg_prog = "QI1"

                    IF left(EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt), 2) = relg_prog AND MAXIS_case_number = relg_case_num Then

                    Else
                        Do
                            pf7
                            EmReadScreen top_check, 13, 24, 2
                        Loop Until top_check = "CANNOT SCROLL"

                        relg_row = 6
                        Do
                            EMReadScreen relg_prog, 2, relg_row, 10 'reading the prog and elig type information
                            EMReadScreen relg_elig, 2, relg_row, 33
                            EMReadScreen relg_case_num, 8, relg_row, 73 'reading the case number for this span

                            If relg_prog = "EH" Then relg_prog = "EMA"
                            IF EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt) = "QI1" AND relg_prog = "SL" Then relg_prog = "QI1"

                            IF left(EOMC_CLIENT_ARRAY(hc_prog_one, hc_clt), 2) = relg_prog AND MAXIS_case_number = relg_case_num Then Exit Do

                            relg_row = relg_row + 4
                            If relg_row = 22 Then
                                PF8
                                relg_row = 6
                            End If
                            EmReadScreen bottom_check, 13, 24, 2
                        Loop Until bottom_check = "CANNOT SCROLL"

                    End If

                    EMReadScreen confirm_mmis_end, 8, relg_row+1, 36        'reading the end date and status to make sure the change was successful
                    EMReadScreen confirm_mmis_stat, 1, relg_row+1, 62

                    If confirm_mmis_end <> "99/99/99" Then

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
                                        If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 20 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 50 Then
                                            If EOMC_CLIENT_ARRAY(elig_type_one, hc_clt) = "AX" Then
                                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_21_49_ax
                                            Else
                                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_21_49
                                            End If
                                        End If
                                        If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 49 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 65 Then
                                            If EOMC_CLIENT_ARRAY(elig_type_one, hc_clt) = "AX" Then
                                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_50_64_ax
                                            Else
                                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_50_64
                                            End If
                                        End If
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
                    End If
                    Do
                        PF7

                        EmReadscreen look_for_top, 12, 24,27
                    Loop Until look_for_top = "NO MORE DATA"
                End If

                If EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = "99/99/99" Then
                    If EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt) > 1 Then
                        for forward = 2 to EOMC_CLIENT_ARRAY(RELG_page_two, hc_clt)
                            PF8
                        next
                    End If

                    relg_row = EOMC_CLIENT_ARRAY(RELG_row_two, hc_clt)

                    EMReadScreen relg_prog, 2, relg_row, 10 'reading the prog and elig type information
                    EMReadScreen relg_elig, 2, relg_row, 33
                    EMReadScreen relg_case_num, 8, relg_row, 73 'reading the case number for this span

                    If relg_prog = "EH" Then relg_prog = "EMA"
                    IF EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "QI1" AND relg_prog = "SL" Then relg_prog = "QI1"

                    IF left(EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt), 2) = relg_prog AND MAXIS_case_number = relg_case_num Then

                    Else
                        Do
                            pf7
                            EmReadScreen top_check, 13, 24, 2
                        Loop Until top_check = "CANNOT SCROLL"

                        relg_row = 6
                        Do
                            EMReadScreen relg_prog, 2, relg_row, 10 'reading the prog and elig type information
                            EMReadScreen relg_elig, 2, relg_row, 33
                            EMReadScreen relg_case_num, 8, relg_row, 73 'reading the case number for this span

                            If relg_prog = "EH" Then relg_prog = "EMA"
                            IF EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt) = "QI1" AND relg_prog = "SL" Then relg_prog = "QI1"

                            IF left(EOMC_CLIENT_ARRAY(hc_prog_two, hc_clt), 2) = relg_prog AND MAXIS_case_number = relg_case_num Then Exit Do

                            relg_row = relg_row + 4
                            If relg_row = 22 Then
                                PF8
                                relg_row = 6
                            End If
                            EmReadScreen bottom_check, 13, 24, 2
                        Loop Until bottom_check = "CANNOT SCROLL"

                    End If

                    EMReadScreen confirm_mmis_end, 8, relg_row+1, 36
                    EMReadScreen confirm_mmis_stat, 1, relg_row+1, 62

                    If confirm_mmis_end <> "99/99/99" Then

                        If  confirm_mmis_end = mmis_last_day_date AND confirm_mmis_stat = "C" Then

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
                                        If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 20 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 50 Then
                                            If EOMC_CLIENT_ARRAY(elig_type_one, hc_clt) = "AX" Then
                                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_21_49_ax
                                            Else
                                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_21_49
                                            End If
                                        End If
                                        If EOMC_CLIENT_ARRAY(clt_age, hc_clt) > 49 AND EOMC_CLIENT_ARRAY(clt_age, hc_clt) < 65 Then
                                            If EOMC_CLIENT_ARRAY(elig_type_one, hc_clt) = "AX" Then
                                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_50_64_ax
                                            Else
                                                EOMC_CLIENT_ARRAY(clt_savings, hc_clt) = EOMC_CLIENT_ARRAY(clt_savings, hc_clt) + capitation_50_64
                                            End If
                                        End If
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
                    End If

                End If
                PF6     'backing out of the MMIS client information

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

    total_savings = total_savings + EOMC_CLIENT_ARRAY(clt_savings, hc_clt)      'adding client savings to the total savings for the script run.
Next


excel_row = excel_row - 1
col_to_use = col_to_use + 1     'moving over one extra for script run details.

'Query date/time/runtime info
objExcel.Cells(2, col_to_use).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use).Value = "Query date and time:"
ObjExcel.Cells(1, col_to_use+1).Value = now
ObjExcel.Cells(1, col_to_use+1).Font.Bold = FALSE
ObjExcel.Cells(2, col_to_use).Value = "Query runtime (in seconds):"
ObjExcel.Cells(2, col_to_use+1).Value = timer - query_start_time

If make_changes = TRUE Then
    ObjExcel.Cells(3, col_to_use).Value = "Total Savings:"
    ObjExcel.Cells(4, col_to_use).Value = "Total PMIs:"
    ObjExcel.Cells(5, col_to_use).Value = "Manyally Closed PMIs:"
    ObjExcel.Cells(6, col_to_use).Value = "MMIS Span Not Updated:"
    ObjExcel.Cells(7, col_to_use).Value = "PMIs with savings:"
    ObjExcel.Cells(8, col_to_use).Value = "PMIs Updated MMIS by script:"

    is_not_blank_excel_string = chr(34) & "<>" & chr(34)

    ObjExcel.Cells(3, col_to_use+1).Value = "=SUM(" & savings_letter_col & "2:" & savings_letter_col & excel_row & ")"
    ObjExcel.Cells(3, col_to_use+1).NumberFormat = "$#,##0.00"
    ObjExcel.Cells(4, col_to_use+1).Value = "=COUNTIF(B2:B" & excel_row & ", " & is_not_blank_excel_string & ")"
    ObjExcel.Cells(5, col_to_use+1).Value = "=COUNTIF(C2:C" & excel_row & ", " & chr(34) & "HC" & chr(34) &")"
    ObjExcel.Cells(6, col_to_use+1).Value = "=COUNTIFS(C2:C" & excel_row & ", " & chr(34) & "HC" & chr(34) & ", J2:J" & excel_row & ", " & chr(34) & "99/99/99" & chr(34) &_
     ")+COUNTIFS(C2:C" & excel_row & ", " & chr(34) & "HC" & chr(34) & ", " & mmis_two_letter_col & "2:" & mmis_two_letter_col & excel_row & ", " & chr(34) & "99/99/99" & chr(34) & ", J2:J" & excel_row & ", " & is_not_blank_excel_string & chr(38) & chr(34) & "99/99/99" & chr(34) & ")"
    ObjExcel.Cells(7, col_to_use+1).Value = "=COUNTIF(" & savings_letter_col & "2:" & savings_letter_col & excel_row & ", " & is_not_blank_excel_string & chr(38) & "0)"
    ObjExcel.Cells(8, col_to_use+1).Value = "=COUNTIF(" & new_mmis_one_col_letter_col & "2:" & new_mmis_one_col_letter_col & excel_row & ", " & is_not_blank_excel_string & ")+COUNTIFS(" & new_mmis_two_col_letter_col & "2:" & new_mmis_two_col_letter_col & excel_row & ", " & is_not_blank_excel_string & ", " & new_mmis_one_col_letter_col & "2:" & new_mmis_one_col_letter_col & excel_row & ", " & chr(34) & chr(34) & ")"
End If

'Autofitting columns
For col_to_autofit = 1 to col_to_use+1
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

table_range = "A1:S" & excel_row
table_name = "Table1"

ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table_range, xlYes).Name = table_name
ObjExcel.ActiveSheet.ListObjects(table_name).TableStyle = "TableStyleMedium3"

'setting a freeze row for easy scrolling
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

curr_day = DatePart("d", date)
curr_mo = DatePart("m", date)
curr_yr = DatePart("yyyy", date)
file_friendly_date = curr_mo & "-" & curr_day & "-" & right(2, curr_yr)
EOMC_report_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\HC Discrepancy\EOMC\"
file_name = CM_plus_1_mo & "-20" & CM_plus_1_yr & " Change Run - " & file_friendly_date & ".xlsx"
objExcel.ActiveWorkbook.SaveAs EOMC_report_folder & file_name

If make_changes = TRUE Then
    ' MsgBox "Starting new functionality"
    'Create Worklist Excel'
    'Opening the Excel file
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add()
    objExcel.DisplayAlerts = True

    'Name for the current sheet'
    ObjExcel.ActiveSheet.Name = "MMIS Span is Open"
    on_loop = 1
    mmis_last_day_date = DateValue(mmis_last_day_date)

    Do
        ObjExcel.Cells(1, 1).Value = "Worklist"
        If on_loop = 1 Then ObjExcel.Cells(1, 3).Value = "MMIS still open"
        If on_loop = 2 Then ObjExcel.Cells(1, 3).Value = "MMIS Span End Date Error"
        If on_loop = 3 Then ObjExcel.Cells(1, 3).Value = "MMIS Span Future End Date"
        If on_loop = 4 Then ObjExcel.Cells(1, 3).Value = "MAXIS Budget Error"

        ObjExcel.Cells(2, 1).Value = "Cases In List"
        ObjExcel.Cells(2, 3).Value = "=COUNTIF(G:G, " & is_not_blank_excel_string & ") - 1"

        ObjExcel.Cells(3, 1).Value = "Date Assigned"
        ObjExcel.Cells(3, 3).Value = date & ""

        ObjExcel.Cells(4, 1).Value = "Date Completed"

        ObjExcel.Cells(5, 1).Value = "Instructions"
        If on_loop = 1 Then ObjExcel.Cells(5, 3).Value = "End MMIS Span"
        If on_loop = 2 Then ObjExcel.Cells(5, 3).Value = "Update MMIS date to match MAXIS closure."
        If on_loop = 3 Then ObjExcel.Cells(5, 3).Value = "Check case and align dates."
        If on_loop = 4 Then ObjExcel.Cells(5, 3).Value = "Determine why Budgets are not being approved."

        ObjExcel.Cells(6, 1).Value = "Goal"
        If on_loop = 1 OR on_loop = 2 OR on_loop = 3 Then ObjExcel.Cells(6, 3).Value = "To reduce discrepancies between MAXIS and MMIS closures."
        If on_loop = 4 Then ObjExcel.Cells(6, 3).Value = "Discover areas of increased need for direction/work."

        col_to_use = 1

        'Excel headers and formatting the columns
        If on_loop = 4 Then
            objExcel.Cells(8, 1).Value  = "WORKER"
            objExcel.Cells(8, 2).Value  = "CASE NUMBER"
            objExcel.Cells(8, 3).Value  = "EOMC Status"
            objExcel.Cells(8, 4).Value  = "REF NO"
            objExcel.Cells(8, 5).Value  = "NAME"
            objExcel.Cells(8, 6).Value  = "PMI"
            objExcel.Cells(8, 7).Value  = "1st Prog"
            objExcel.Cells(8, 8).Value  = "ELIG TYPE"
            objExcel.Cells(8, 9).Value  = "2nd PROG"
            objExcel.Cells(8, 10).Value = "ELIG TYPE - 2"
            objExcel.Cells(8, 11).Value = "SAVINGS"
            objExcel.Cells(8, 12).Value = "ERRORS"
            objExcel.Cells(8, 13).Value = "REASON"
            objExcel.Cells(8, 14).Value = "MISSING MONTHS"

        Else
            objExcel.Cells(8, 1).Value  = "WORKER"
            objExcel.Cells(8, 2).Value  = "CASE NUMBER"
            objExcel.Cells(8, 3).Value  = "EOMC Status"
            objExcel.Cells(8, 4).Value  = "REF NO"
            objExcel.Cells(8, 5).Value  = "NAME"
            objExcel.Cells(8, 6).Value  = "PMI"
            objExcel.Cells(8, 7).Value  = "1st Prog"
            objExcel.Cells(8, 8).Value  = "ELIG TYPE"
            objExcel.Cells(8, 9).Value  = "MAXIS End Date"
            objExcel.Cells(8, 10).Value = "CURR MMIS End Date"
            objExcel.Cells(8, 11).Value = "NEW MMIS End Date"
            objExcel.Cells(8, 12).Value = "PPH Cap"
            objExcel.Cells(8, 13).Value = "2nd PROG"
            objExcel.Cells(8, 14).Value = "ELIG TYPE - 2"
            objExcel.Cells(8, 15).Value = "MAXIS End Date - 2"
            objExcel.Cells(8, 16).Value = "CURR MMIS End Date - 2"
            objExcel.Cells(8, 17).Value = "NEW MMIS End Date - 2"
            objExcel.Cells(8, 18).Value = "SAVINGS"
            objExcel.Cells(8, 19).Value = "ERRORS"
            objExcel.Cells(8, 20).Value = "ACTIONS"
            objExcel.Cells(8, 21).Value = "NOTES"

        End If

        For i = 1 to col_to_use
            ObjExcel.Cells(8, i).Font.Bold = TRUE
        Next

        excel_row = 9
        'Looping through each of the HC clients while in MMIS
        For hc_clt = 0 to UBOUND(EOMC_CLIENT_ARRAY, 2)
            write_this_entry = FALSE
            If on_loop = 1 Then
                If EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = "99/99/99" AND trim(EOMC_CLIENT_ARRAY(MMIS_new_end_one, hc_clt)) = "" Then
                    write_this_entry = TRUE
                Else
                    If EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = "99/99/99" AND trim(EOMC_CLIENT_ARRAY(MMIS_new_end_two, hc_clt)) = "" Then write_this_entry = TRUE
                End If
            ElseIf on_loop = 2 Then
                If EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) <> "" AND EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) <> "99/99/99" Then
                    EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = DateValue(EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt))
                    If DateDiff("d", mmis_last_day_date, EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt)) < 0 OR EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = "" Then write_this_entry = TRUE
                End If
                If write_this_entry = FALSE Then
                    If EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) <> "" AND EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) <> "99/99/99" Then
                        EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = DateValue(EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt))
                        If DateDiff("d", mmis_last_day_date, EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt)) < 0 OR EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = "" Then write_this_entry = TRUE
                    End If
                End If
            ElseIf on_loop = 3 Then
                If EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) <> "" AND EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) <> "99/99/99" Then
                    EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt) = DateValue(EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt))
                    If DateDiff("d", mmis_last_day_date, EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt)) > 0 Then write_this_entry = TRUE
                End If
                If write_this_entry = FALSE Then
                    If EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) <> "" AND EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) <> "99/99/99" Then
                        EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt) = DateValue(EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt))
                        If DateDiff("d", mmis_last_day_date, EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt)) > 0 Then write_this_entry = TRUE
                    End If
                End If
            ElseIf on_loop = 4 Then
                If InStr(EOMC_CLIENT_ARRAY(err_notes, hc_clt), "Budget Needs Approval") <> 0 Then write_this_entry = TRUE
            End If

            If write_this_entry = TRUE Then
                ObjExcel.Cells(excel_row, 1).Value              = EOMC_CLIENT_ARRAY (basket_nbr,   hc_clt)
                ObjExcel.Cells(excel_row, 2).Value              = EOMC_CLIENT_ARRAY (case_nbr,  hc_clt)
                ObjExcel.Cells(excel_row, 3).Value              = EOMC_CLIENT_ARRAY (hc_close_stat, hc_clt)
                ObjExcel.Cells(excel_row, 4).Value              = "Memb " & EOMC_CLIENT_ARRAY(clt_ref_nbr, hc_clt)
                ObjExcel.Cells(excel_row, 5).Value              = EOMC_CLIENT_ARRAY (clt_name,  hc_clt)
                ObjExcel.Cells(excel_row, 6).Value              = EOMC_CLIENT_ARRAY (clt_pmi,   hc_clt)

                ObjExcel.Cells(excel_row, 7).Value              = EOMC_CLIENT_ARRAY (hc_prog_one,   hc_clt)
                ObjExcel.Cells(excel_row, 8).Value              = EOMC_CLIENT_ARRAY (elig_type_one,   hc_clt)

                If on_loop = 4 Then
                    ObjExcel.Cells(excel_row, 9).Value          = EOMC_CLIENT_ARRAY (hc_prog_two,   hc_clt)
                    ObjExcel.Cells(excel_row, 10).Value         = EOMC_CLIENT_ARRAY (elig_type_two,   hc_clt)
                    ObjExcel.Cells(excel_row, 11).Value         = EOMC_CLIENT_ARRAY(clt_savings, hc_clt)
                    ObjExcel.Cells(excel_row, 11).NumberFormat  = "$#,##0.00"

                    ObjExcel.Cells(excel_row, 12).Value         = EOMC_CLIENT_ARRAY(err_notes, hc_clt)
                Else
                    ObjExcel.Cells(excel_row, 9).Value          = EOMC_CLIENT_ARRAY(prog_one_end, hc_clt)
                    ObjExcel.Cells(excel_row, 10).Value         = EOMC_CLIENT_ARRAY(MMIS_curr_end_one, hc_clt)
                    ObjExcel.Cells(excel_row, 11).Value         = EOMC_CLIENT_ARRAY(MMIS_new_end_one, hc_clt)
                    ObjExcel.Cells(excel_row, 12).Value         = EOMC_CLIENT_ARRAY(capitation_ended, hc_clt)

                    ObjExcel.Cells(excel_row, 13).Value         = EOMC_CLIENT_ARRAY (hc_prog_two,   hc_clt)
                    ObjExcel.Cells(excel_row, 14).Value         = EOMC_CLIENT_ARRAY (elig_type_two,   hc_clt)
                    ObjExcel.Cells(excel_row, 15).Value         = EOMC_CLIENT_ARRAY(prog_two_end, hc_clt)
                    ObjExcel.Cells(excel_row, 16).Value         = EOMC_CLIENT_ARRAY(MMIS_curr_end_two, hc_clt)
                    ObjExcel.Cells(excel_row, 17).Value         = EOMC_CLIENT_ARRAY(MMIS_new_end_two, hc_clt)
                    ObjExcel.Cells(excel_row, 18).Value         = EOMC_CLIENT_ARRAY(clt_savings, hc_clt)
                    ObjExcel.Cells(excel_row, 18).NumberFormat  = "$#,##0.00"

                    ObjExcel.Cells(excel_row, 19).Value         = EOMC_CLIENT_ARRAY(err_notes, hc_clt)
                End If

            	excel_row = excel_row + 1      'next row
            End If
        Next

        For i = 1 to col_to_use
            ObjExcel.columns(col_to_autofit).AutoFit()
        Next

        For xl_row = 1 to 5
            ObjExcel.Range("A" & xl_row & ":B" & xl_row).Merge
            ObjExcel.Cells(xl_row, 1).HorizontalAlignment = -4152       'Aligns text in Excel Cell to the right
            ObjExcel.Range("C" & xl_row & ":E" & xl_row).Merge
            ObjExcel.Cells(xl_row, 3).HorizontalAlignment = -4108       'Aligns text in Excel Cell to the center
        Next
        ObjExcel.Range("A6:B6").Merge
        ObjExcel.Cells(6, 1).HorizontalAlignment = -4152       'Aligns text in Excel Cell to the right
        ObjExcel.Range("C6:G6").Merge
        ObjExcel.Cells(6, 3).HorizontalAlignment = -4108       'Aligns text in Excel Cell to the center

		table_range = "A8:U" & excel_row-1
		If on_loop = 4 Then table_range = "A8:N" & excel_row-1

		If on_loop = 1 Then table_name = "MMISOpenTable"
		If on_loop = 2 Then table_name = "EndDateErrorTable"
		If on_loop = 3 Then table_name = "FutureEndDateTable"
		If on_loop = 4 Then table_name = "BudgetErrorTable"

		ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, table_range, xlYes).Name = table_name
		ObjExcel.ActiveSheet.ListObjects(table_name).TableStyle = "TableStyleMedium3"

        on_loop = on_loop + 1
        If on_loop = 2 Then ObjExcel.Worksheets.Add().Name = "MMIS Span End Date Error"
        If on_loop = 3 Then ObjExcel.Worksheets.Add().Name = "MMIS Span Future End Date"
        If on_loop = 4 Then ObjExcel.Worksheets.Add().Name = "MAXIS Budget Error"

    Loop until on_loop = 5
End If

EOMC_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Discrepancy HC\End of Month Closures\"
file_name = CM_plus_1_mo & "-" & CM_plus_1_yr & " Closures - EOMC workslist.xlsx"
objExcel.ActiveWorkbook.SaveAs EOMC_folder & file_name

email_subject = "HC End of Month Closures list is Ready"

email_body = "Hello QI!" & "<br>" & "<br>"
email_body = email_body & vbCr & "The Excel file has been created to review HC cases that are set to close for " & CM_plus_1_mo & "/" & CM_plus_1_yr & "." & "<br>"
email_body = email_body & vbCr & "This list is here:" & "<br>"
email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Discrepancy HC\End of Month Closures\" & file_name & chr(34) & ">" & file_name & "</a><br>" & "<br>"

email_body = email_body & vbCr & vbCr & "This script has attempted to align MMIS to the MAXIS HC eligibility, but some cases need manual review/action." & "<br>"
email_body = email_body & vbCr & "There is an instruction document here:" & "<br>"
email_body = email_body & "&emsp;&ensp;" & "- " & "<a href=" & chr(34) & "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\Discrepancy HC\End of Month Closures\" & "EOMC Work List Instructions.docx" & chr(34) & ">" & "EOMC Work List Instructions.docx" & "</a><br>" & "<br>"

email_body = email_body & vbCr & "Please reach out to Tanya with questions about this assignment." & "<br>"
email_body = email_body & vbCr & vbCr & "Thank you!"


'function labels		  email_from, 							  email_recip, 				 email_recip_CC, 		    email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email
Call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "HSPH.EWS.QI@hennepin.us", "Tanya.Payne@hennepin.us", "", 			 email_subject, 1, 				  False, 		email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, False, 				   email_attachment_array, True)

script_end_procedure("EOMC Automation completed and worklists created. Email sent to QI. Script run is complete.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------7/28/2023
'--Tab orders reviewed & confirmed----------------------------------------------7/28/2023
'--Mandatory fields all present & Reviewed--------------------------------------7/28/2023
'--All variables in dialog match mandatory fields-------------------------------7/28/2023
'Review dialog names for content and content fit in dialog----------------------7/28/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------7/28/2023					This script has the cool - keep MMIS logged in functionality
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------7/28/2023
'--Out-of-County handling reviewed----------------------------------------------N/A							Out of county not needed because cases are pulled from REPT/EOMC which is county specific
'--script_end_procedures (w/ or w/o error messaging)----------------------------7/28/2023
'--BULK - review output of statistics and run time/count (if applicable)--------7/28/2023
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------7/28/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------7/28/2023
'--Incrementors reviewed (if necessary)-----------------------------------------7/28/2023
'--Denomination reviewed -------------------------------------------------------7/28/2023
'--Script name reviewed---------------------------------------------------------7/28/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------7/28/2023					starts with 0 incrementor

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------7/28/2023
'--comment Code-----------------------------------------------------------------7/28/2023
'--Update Changelog for release/update------------------------------------------7N/A
'--Remove testing message boxes-------------------------------------------------7/28/2023
'--Remove testing code/unnecessary code-----------------------------------------7/28/2023
'--Review/update SharePoint instructions----------------------------------------7/28/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------7/28/2023
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------7/28/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------7/28/2023
'--Complete misc. documentation (if applicable)---------------------------------7/28/2023
'--Update project team/issue contact (if applicable)----------------------------7/28/2023