'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - HC Eligibility.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 1                       'manual run time in seconds
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
call changelog_update("09/30/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS--------------------------------------------------------------------------------------------------
'Variables


'Constants
Const clt_ref_nbr       = 0
Const clt_first_name    = 1
Const clt_last_name     = 2
Const clt_full_name     = 3
Const clt_pmi           = 4
Const foot_mo           = 5
Const foot_yr           = 6
Const hc_test_one       = 7
Const hc_prog_one       = 8
Const hc_elig_one       = 9
Const hc_std_one        = 10
Const hc_meth_one       = 11
Const hc_waiver_one     = 12
Const hc_spdwn_one      = 13
Const mmis_span_one     = 14
Const mmis_end_one      = 15
Const mmis_stat_one     = 16
Const hc_test_two       = 17
Const hc_prog_two       = 18
Const hc_elig_two       = 19
Const hc_std_two        = 20
Const hc_meth_two       = 21
Const mmis_span_two     = 22
Const mmis_end_two      = 23
Const mmis_stat_two     = 24
Const approval_today    = 25
Const action_type       = 26
Const err_notes         = 27

'Arrays
dim CLIENT_HC_ELIG_ARRAY ()
redim CLIENT_HC_ELIG_ARRAY (err_notes, 0)

'--------------------------------------------------------------------------------------------------------------
'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS
EMConnect ""
'Finds the case number
call MAXIS_case_number_finder(MAXIS_case_number)

BeginDialog elig_dlg, 0, 0, 121, 70, "Case Information"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  EditBox 80, 30, 15, 15, start_mo
  EditBox 100, 30, 15, 15, start_yr
  ButtonGroup ButtonPressed
    OkButton 45, 50, 35, 15
    CancelButton 80, 50, 35, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 10, 35, 70, 10, "First month of action:"
EndDialog

Do
	Do
		'Adding err_msg handling
		err_msg = ""

        Dialog elig_dlg

        If len(MAXIS_case_number) > 7 Then err_msg = err_msg & vbNewLine & "* Review the case number, it appears to be too long."
        If trim(MAXIS_case_number) = "" Then err_msg = err_msg & vbNewLine & "* Enter a case number."
        If IsNumeric(MAXIS_case_number) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid case number"
        If len(start_mo) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
        If len(start_yr) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."

        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
Loop until are_we_passworded_out = false

MAXIS_case_number = trim(MAXIS_case_number)
MAXIS_footer_month = start_mo
MAXIS_footer_year = start_yr

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

Call navigate_to_MAXIS_screen("ELIG", "HC  ")

EMReadScreen hc_elig_check, 4, 3, 51
If hc_elig_check <> "HHMM" Then script_end_procedure("No HC ELIG results exist, resolve edits and approve new version and run the script again.")
EMWriteScreen approval_month, 20, 56            'Goes to the next month and checks that elig results exist
EMWriteScreen approval_year,  20, 59
transmit
If hc_elig_check <> "HHMM" Then script_end_procedure("No HC ELIG results exist, resolve edits and approve new version and run the script again.")

'Read for each person on HC in the start month and year - any approval done in the current day
hc_clt = 0
row = 8                                          'Reads each line of Elig HC to find all the approved programs in a case
Do
    EMReadScreen elig_ref_num, 2, row, 3
    EMReadScreen elig_hc_prog, 12, row, 28
    elig_hc_prog = trim(elig_hc_prog)

    'looking for clients with HC eligibility
    If elig_hc_prog <> "NO VERSION" AND elig_hc_prog <> "NO REQUEST" AND elig_hc_prog <> "" Then
        ReDim Preserve CLIENT_HC_ELIG_ARRAY(err_notes, hc_clt)
        Do
            EMReadScreen prog_status, 3, row, 68
            If prog_status <> "APP" Then                        'Finding the approved version
                EMReadScreen total_versions, 2, row, 64
                If total_versions = "01" Then
                    CLIENT_HC_ELIG_ARRAY(clt_ref_nbr, hc_clt)   = elig_ref_num
                    CLIENT_HC_ELIG_ARRAY(hc_prog_one, hc_clt)   = elig_hc_prog
                    CLIENT_HC_ELIG_ARRAY(foot_mo, hc_clt)       = MAXIS_footer_month
                    CLIENT_HC_ELIG_ARRAY(foot_yr, hc_clt)       = MAXIS_footer_year
                    CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt)= FALSE
                    CLIENT_HC_ELIG_ARRAY(err_notes, hc_clt)     = "HC eligiblity not approved in MAXIS"
                    Exit Do
                Else
                    EMReadScreen current_version, 2, row, 58
                    If current_version = "01" Then
                        CLIENT_HC_ELIG_ARRAY(clt_ref_nbr, hc_clt)   = elig_ref_num
                        CLIENT_HC_ELIG_ARRAY(hc_prog_one, hc_clt)   = elig_hc_prog
                        CLIENT_HC_ELIG_ARRAY(foot_mo, hc_clt)       = MAXIS_footer_month
                        CLIENT_HC_ELIG_ARRAY(foot_yr, hc_clt)       = MAXIS_footer_year
                        CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt)= FALSE
                        CLIENT_HC_ELIG_ARRAY(err_notes, hc_clt)     = "HC eligiblity not approved in MAXIS"
                        Exit Do
                    End If
                    prev_version = right ("00" & abs(current_version) - 1, 2)
                    EMWriteScreen prev_version, row, 58
                    transmit
                End If
            End If
        Loop until current_version = "01" OR prog_status = "APP"

        If CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt) <> FALSE Then
            EMWriteScreen "x", row, 26
            transmit
            'TODO see if the process date and application date are in the same place for all programs
            EMReadScreen process_date, 8, 2, 73
            If DateValue(process_date) <> date then
                CLIENT_HC_ELIG_ARRAY(clt_ref_nbr, hc_clt)   = elig_ref_num
                CLIENT_HC_ELIG_ARRAY(hc_prog_one, hc_clt)   = elig_hc_prog
                CLIENT_HC_ELIG_ARRAY(foot_mo, hc_clt)       = MAXIS_footer_month
                CLIENT_HC_ELIG_ARRAY(foot_yr, hc_clt)       = MAXIS_footer_year
                CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt)= FALSE
                CLIENT_HC_ELIG_ARRAY(err_notes, hc_clt)     = "HC was not approved today."
                Exit Do

            ElseIF elig_hc_prog = "MA" OR elig_hc_prog = "IMD" OR elig_hc_prog = "EMA" Then
                EMReadScreen appl_month, 2, 3, 73
                EMReadScreen appl_year, 2, 3, 79

                mo_col = 19                                     'setting the column for reading the month and year of the HC information for the client
                yr_col = 22
                Do                                              'we will look through each of the 6 months in the budget to find the current month and year
                    EMReadScreen bsum_mo, 2, 6, mo_col          'reading the month and year
                    EMReadScreen bsum_yr, 2, 6, yr_col

                    If bsum_mo = MAXIS_footer_month and bsum_yr = MAXIS_footer_year Then Exit Do        'if it is this month and year, we found the right month and year
                    mo_col = mo_col + 11                        'if it doesn't match, then we go to the next - which is 11 over
                    yr_col = yr_col + 11
                    'MsgBox "Loop 3 - month col: " & mo_col
                Loop until mo_col = 85                          'this is the last month

                If mo_col = 85 Then
                    CLIENT_HC_ELIG_ARRAY(clt_ref_nbr, hc_clt)   = elig_ref_num
                    CLIENT_HC_ELIG_ARRAY(hc_prog_one, hc_clt)   = elig_hc_prog
                    CLIENT_HC_ELIG_ARRAY(foot_mo, hc_clt)       = MAXIS_footer_month
                    CLIENT_HC_ELIG_ARRAY(foot_yr, hc_clt)       = MAXIS_footer_year
                    CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt)= FALSE
                    CLIENT_HC_ELIG_ARRAY(err_notes, hc_clt)     = "Month not covered in the approved MAXIS Elig."
                End If
            End If
        End If

        'TODO Add special functionality for LTC/Waiver cases
        If CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt) <> FALSE Then
            If elig_hc_prog = "MA" OR elig_hc_prog = "IMD" OR elig_hc_prog = "EMA" Then
                EMReadScreen pers_test, 6, 7, mo_col

                EMReadScreen prog, 4, 11, mo_col                'reading all of the detail in this month of BSUM
                EMReadScreen pers_type, 2, 12, mo_col-2
                EMReadScreen pers_std, 1, 12, yr_col
                EMReadScreen pers_mthd, 1, 13, yr_col-1
                EMReadScreen pers_waiv, 1, 14, yr_col-1

                CLIENT_HC_ELIG_ARRAY(clt_ref_nbr, hc_clt)   = elig_ref_num
                CLIENT_HC_ELIG_ARRAY(foot_mo, hc_clt)       = MAXIS_footer_month
                CLIENT_HC_ELIG_ARRAY(foot_yr, hc_clt)       = MAXIS_footer_year
                CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt)= TRUE

                CLIENT_HC_ELIG_ARRAY(hc_test_one, hc_clt)   = trim(pers_test)
                CLIENT_HC_ELIG_ARRAY(hc_prog_one, hc_clt)   = trim(prog)
                CLIENT_HC_ELIG_ARRAY(hc_elig_one, hc_clt)   = pers_type
                CLIENT_HC_ELIG_ARRAY(hc_std_one, hc_clt)    = pers_std
                CLIENT_HC_ELIG_ARRAY(hc_meth_one, hc_clt)   = pers_mthd
                CLIENT_HC_ELIG_ARRAY(hc_waiver_one, hc_clt) = pers_waiv

                'Looking in this span to see if there are any additional months. '
                Do
                    mo_col = mo_col + 11
                    yr_col = yr_col + 11
                    If mo_col = 85 Then Exit Do

                    EMReadScreen bsum_mo, 2, 6, mo_col          'reading the month and year
                    EMReadScreen bsum_yr, 2, 6, yr_col

                    If bsum_mo <> "  " AND bsum_yr <> "  " Then
                        hc_clt = hc_clt + 1
                        ReDim Preserve CLIENT_HC_ELIG_ARRAY(err_notes, hc_clt)

                        EMReadScreen pers_test, 6, 7, mo_col

                        EMReadScreen prog, 4, 11, mo_col                'reading all of the detail in this month of BSUM
                        EMReadScreen pers_type, 2, 12, mo_col-2
                        EMReadScreen pers_std, 1, 12, yr_col
                        EMReadScreen pers_mthd, 1, 13, yr_col-1
                        EMReadScreen pers_waiv, 1, 14, yr_col-1

                        CLIENT_HC_ELIG_ARRAY(clt_ref_nbr, hc_clt)   = elig_ref_num
                        CLIENT_HC_ELIG_ARRAY(foot_mo, hc_clt)       = bsum_mo
                        CLIENT_HC_ELIG_ARRAY(foot_yr, hc_clt)       = bsum_yr
                        CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt)= TRUE

                        CLIENT_HC_ELIG_ARRAY(hc_test_one, hc_clt)   = trim(pers_test)
                        CLIENT_HC_ELIG_ARRAY(hc_prog_one, hc_clt)   = trim(prog)
                        CLIENT_HC_ELIG_ARRAY(hc_elig_one, hc_clt)   = pers_type
                        CLIENT_HC_ELIG_ARRAY(hc_std_one, hc_clt)    = pers_std
                        CLIENT_HC_ELIG_ARRAY(hc_meth_one, hc_clt)   = pers_mthd
                        CLIENT_HC_ELIG_ARRAY(hc_waiver_one, hc_clt) = pers_waiv

                    End If
                Loop until bsum_mo = "  " AND bsum_yr = "  "
            Else


                EMReadScreen pers_type, 2, 6, 56                                'reading the type and standard
                EMReadScreen pers_std, 1, 6, 64

                transmit
                transmit

                EMReadScreen pers_test, 10, 9, 34
                EMReadScreen appl_month, 2, 3, 73
                EMReadScreen appl_year, 2, 3, 79

                CLIENT_HC_ELIG_ARRAY(clt_ref_nbr, hc_clt)   = elig_ref_num
                CLIENT_HC_ELIG_ARRAY(foot_mo, hc_clt)       = MAXIS_footer_month
                CLIENT_HC_ELIG_ARRAY(foot_yr, hc_clt)       = MAXIS_footer_year
                CLIENT_HC_ELIG_ARRAY(approval_today, hc_clt)= TRUE

                CLIENT_HC_ELIG_ARRAY(hc_test_one, hc_clt)   = trim(pers_test)
                CLIENT_HC_ELIG_ARRAY(hc_prog_one, hc_clt)   = elig_hc_prog
                CLIENT_HC_ELIG_ARRAY(hc_elig_one, hc_clt)   = pers_type
                CLIENT_HC_ELIG_ARRAY(hc_std_one, hc_clt)    = pers_std

            End If
        End If


        row = row + 1

        EMReadScreen next_elig_ref_num, 2, row, 3
        EMReadScreen next_elig_hc_prog, 12, row, 28
        next_elig_hc_prog = trim(next_elig_hc_prog)

        If next_elig_ref_num = "  " AND next_elig_hc_prog <> "" Then
            hc_clt = hc_clt + 1
            ReDim Preserve CLIENT_HC_ELIG_ARRAY(err_notes, hc_clt)


        End If
    Else
        row = row + 1
    End If
    EMReadScreen next_elig_hc_prog, 12, row, 28
    next_elig_hc_prog = trim(next_elig_hc_prog)

Loop until next_elig_hc_prog = ""
'Identify elig or inelig to determine approval vs closure vs denial
'TODO figure out how approval/denial/closure look different'
'create dynamic dialog for EACH client and have it specific to the elig information found
'Look at the the following months to ensure nothing has changed.

'case note detail of approval
