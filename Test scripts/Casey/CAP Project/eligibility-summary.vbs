'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ELIGIBILITY SUMMARY.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each Case
'END OF stats block=========================================================================================================

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
call changelog_update("06/06/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS==================================================================================================================

'-------------------------------FUNCTIONS WE INVENTED THAT WILL SOON BE ADDED TO FUNCLIB
FUNCTION date_array_generator(initial_month, initial_year, date_array)
	'defines an intial date from the initial_month and initial_year parameters
	initial_date = initial_month & "/1/" & initial_year
	'defines a date_list, which starts with just the initial date
	date_list = initial_date
	'This loop creates a list of dates
	Do
		If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
		working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
		date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
	Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

	'Splits this into an array
	date_array = split(date_list, "|")
End function

'===========================================================================================================================

'DECLARATIONS===============================================================================================================
Dim ALL_APPROVALS_ARRAY()
ReDim ALL_APPROVALS_ARRAY(app_notes, 0)
'Constants
Const app_prog                  = 0
Const app_mo                    = 1
Const app_yr                    = 2
Const app_nav                   = 3
Const app_done                  = 4
Const elig_memb                 = 5
Const app_version               = 6

Const hc_maj_prog               = 7
Const hc_elig_type              = 8
Const hc_elig_stnd              = 9
Const hc_elig_mthd              = 10
Const hc_waiv_type              = 11

Const app_type                  = 12
Const app_benefit               = 13
Const app_mf_mf                 = 14
Const app_mf_fs                 = 15
Const app_mf_hg                 = 16
Const app_recoup                = 17

Const app_vendor_amt_one        = 18
Const app_vendor_info_one       = 19
Const app_vendor_amt_two        = 20
Const app_vendor_info_two       = 21
Const app_vendor_amt_three      = 22
Const app_vendor_info_three     = 23
Const app_budg_type             = 24
Const app_rept_status           = 25

Const total_budgeted_income     = 26
Const earned_income             = 27
Const ei_disregard              = 28
Const unearned_income           = 29
Const total_deductions          = 30
Const shelter_expense           = 31
Const utility_expense           = 32

Const app_memb_detail           = 33

Const app_notes     = 34

'===========================================================================================================================

EMConnect ""

'Attempt to gather case number and footer month/year
'Finds the case number
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(initial_footer_month, initial_footer_year)

'Dialog to confirm case number AND the first footer month/year of approval.
'IF multiple programs with different approval months then enter the very first of all of them.
'Approval month means month in which an APP was done.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 131, 105, "Case Number Dialog"
  EditBox 75, 10, 50, 15, MAXIS_case_number
  EditBox 90, 30, 15, 15, initial_footer_month
  EditBox 110, 30, 15, 15, initial_footer_year
  ButtonGroup ButtonPressed
    OkButton 20, 85, 50, 15
    CancelButton 75, 85, 50, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 10, 35, 80, 10, "Footer Month and Year:"
  Text 15, 55, 110, 25, "List the first month that was acted on for ANY program as the footer month and year."
EndDialog

Do
    Do
        err_msg = ""

        dialog dialog1

        cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
        If IsNumeric(initial_footer_month) = FALSE OR len(initial_footer_month) > 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid footer month"
        If IsNumeric(initial_footer_year) = FALSE OR len(initial_footer_year) > 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid footer year"


    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

initial_footer_month = right("00" & initial_footer_month, 2)
initial_footer_year = right(initial_footer_year, 2)

Call back_to_SELF                       'need to gather some detail to have the correct script run

developer_mode = FALSE                  'allowing worker to exit if started in Inquiry on accident
EMReadScreen MX_region, 12, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are in INQUIRY. Income information cannot be saved to STAT and a CASE/NOTE cannot be created." & vbNewLine & vbNewLine & "Do you wish to continue?", vbQuestion + vbYesNo, "Continue in Inquiry?")
    If continue_in_inquiry = vbNo Then script_end_procedure("Script ended since it was started in Inquiry.")
    developer_mode = TRUE
End If
If developer_mode = TRUE then MsgBox "Developer Mode ACTIVATED!"        'developer mode difference is that the MAXIS update detail is shown in a messagebox instead of updating the panel

'creates an array of all months from initial to CM+1
CALL date_array_generator(initial_footer_month, initial_footer_year, APPROVAL_MONTHS_ARRAY)


'Go to ELIG summ for all months from initial for CM + 1 and gather an array of each program for each month.
'Make sure the ELIG is from the current date.
Call back_to_SELF
Call navigate_to_MAXIS_screen ("ELIG", "SUMM")

each_approval = 0

For each approval_month in APPROVAL_MONTHS_ARRAY
    MAXIS_footer_month = DatePart("m", approval_month)          'setting the footer month and year to the next month
    MAXIS_footer_year = DatePart("yyyy", approval_month)

    MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
    MAXIS_footer_year = right(MAXIS_footer_year, 2)

    EMWriteScreen MAXIS_footer_month, 19, 56            'Entering the footer month and going to the correct ELIG/SUMM screen
    EMWriteScreen MAXIS_footer_year, 19, 59
    transmit

    elig_row = 7
    Do
        EMReadScreen version_date, 8, elig_row, 48

        If version_date <> "        " Then
            If DateDiff("d", version_date, date) = 0 Then
                ReDim Preserve ALL_APPROVALS_ARRAY(app_notes, each_approval)

                EMReadScreen elig_prog, 12, elig_row, 22
                elig_prog = trim(elig_prog)
                ALL_APPROVALS_ARRAY(app_prog, each_approval) = elig_prog
                ALL_APPROVALS_ARRAY(app_nav, each_approval) = elig_prog
                If elig_prog = "Food Support" Then
                    ALL_APPROVALS_ARRAY(app_prog, each_approval) = "SNAP"
                    ALL_APPROVALS_ARRAY(app_nav, each_approval) ="FS"
                ElseIf elig_prog = "Cash Denial" Then
                    ALL_APPROVALS_ARRAY(app_prog, each_approval) = "CASH"
                    ALL_APPROVALS_ARRAY(app_nav, each_approval) ="DENY"
                End If
                ALL_APPROVALS_ARRAY(app_mo, each_approval) = MAXIS_footer_month
                ALL_APPROVALS_ARRAY(app_yr, each_approval) = MAXIS_footer_year
                ALL_APPROVALS_ARRAY(app_done, each_approval) = FALSE

                each_approval = each_approval + 1
            End If
        End If

        elig_row = elig_row + 1
    Loop until elig_row = 18

Next

For the_approval = 0 to UBOUND(ALL_APPROVALS_ARRAY, 2)
    Call back_to_SELF
    MAXIS_footer_month = ALL_APPROVALS_ARRAY(app_mo, the_approval)
    MAXIS_footer_year = ALL_APPROVALS_ARRAY(app_yr, the_approval)

    Call navigate_to_MAXIS_screen("ELIG", ALL_APPROVALS_ARRAY(app_nav, the_approval))

    row = 1
    col = 1
    EMSearch "Command:", row, col           'the command line is in different places for different ELIG programs - so we have to search for it
    EMWriteScreen "99", row, col + 17       'putting '99' into the 3rd field of command pulls up a list of all the approvals done
    transmit

    vers_row = 7
    Do
        version_number = ""
        EMReadScreen version_status, 8, vers_row, 50
        If version_status = "APPROVED" Then
            EMReadScreen process_date, 8, vers_row, 26
            If DateDiff("d", process_date, date) = 0 Then
                ALL_APPROVALS_ARRAY(app_done, the_approval) = TRUE
                EMReadScreen version_number, 2, vers_row, 22
                version_number = trim(version_number)
                ALL_APPROVALS_ARRAY(app_version, the_approval) = right("00" & version_number, 2)
            End If
        End If

        vers_row = vers_row + 1
    Loop until version_status = "        "
    transmit

Next

'NOW we go to ELIG/HC for all of the months as HC is not listed on ELIG/SUMM
For each approval_month in APPROVAL_MONTHS_ARRAY
    MAXIS_footer_month = DatePart("m", approval_month)          'setting the footer month and year to the next month
    MAXIS_footer_year = DatePart("yyyy", approval_month)

    MAXIS_footer_month = right("00" & MAXIS_footer_month, 2)
    MAXIS_footer_year = right(MAXIS_footer_year, 2)

    Call back_to_SELF
    Call navigate_to_MAXIS_screen("ELIG", "HC  ")
    EMReadScreen hc_elig_check, 4, 3, 51

    If hc_elig_check = "HHMM" Then

        row = 8                                          'Reads each line of Elig HC to find all the approved programs in a case
        Do
            EMReadScreen clt_ref_num, 2, row, 3
            EMReadScreen clt_hc_prog, 4, row, 28
            'MsgBox clt_hc_prog
            If clt_ref_num = "  " AND clt_hc_prog <> "    " then        'If a client has more than 1 program - the ref number is only listed at the top one
                prev = 1
                Do
                    EMReadScreen clt_ref_num, 2, row - prev, 3
                    prev = prev + 1
                Loop until clt_ref_num <> "  "
            End If
            If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "    " Then     'Gets additional information for all clts with HC programs on this case
                Do
                    EMReadScreen prog_status, 3, row, 68
                    'MsgBox prog_status
                    If prog_status <> "APP" Then                        'Finding the approved version
                        EMReadScreen total_versions, 2, row, 64
                        If total_versions = "01" Then
                            error_processing_msg = error_processing_msg & vbNewLine & "Appears HC eligibility was not approved in " & approval_month & "/" & approval_year & " for " & clt_ref_num & ", please approve HC and rerunscript."
                            Exit Do
                        Else
                            EMReadScreen current_version, 2, row, 58
                            If current_version = "01" Then
                                error_processing_msg = error_processing_msg & vbNewLine & "Appears HC eligibility was not approved in " & approval_month & "/" & approval_year & " for " & clt_ref_num & ", please approve HC and rerunscript."
                                Exit Do
                            End If
                            prev_version = right ("00" & abs(current_version) - 1, 2)
                            EMWriteScreen prev_version, row, 58
                            transmit
                        End If
                    Else
                        EMReadScreen elig_result, 8, row, 41        'Goes into the elig version to get the major program and elig type
                        EMWriteScreen "x", row, 26
                        transmit

                        elig_col = 19
                        Do
                            EMReadScreen elig_mo, 2, 6, elig_col
                            EMReadScreen elig_yr, 2, 6, elig_col + 3

                            If elig_mo = MAXIS_footer_month AND elig_yr = MAXIS_footer_year Then
                                'MsgBox elig_col
                                Exit Do
                            Else
                                elig_col = elig_col + 11
                            End If

                        Loop Until elig_col = 85

                        If elig_col < 85 Then
                            EMReadScreen major_prog, 4, 11, elig_col
                            EMReadScreen elig_type, 2, 12, elig_col - 2
                            EMReadScreen elig_stnd, 1, 12, elig_col + 3
                            EMReadScreen elig_mthd, 1, 13, elig_col + 2
                            EMReadScreen waiver_check, 1, 14, elig_col + 2        'Checking to see if case may be LTC or Waiver'

                            Do
                                transmit
                                EMReadScreen hc_screen_check, 8, 5, 3
                            Loop until hc_screen_check = "Program:"

                            EMReadScreen process_date, 8, 2, 73
                            EMReadScreen app_date, 8, 4, 73

                            If DateDiff("d", process_date, date) = 0 AND DateDiff("d", app_date, date) = 0 Then
                                ReDim Preserve ALL_APPROVALS_ARRAY(app_notes, each_approval)

                                ALL_APPROVALS_ARRAY(elig_memb, each_approval) = clt_ref_num
                                ALL_APPROVALS_ARRAY(app_prog, each_approval) = "HC"
                                ALL_APPROVALS_ARRAY(app_nav, each_approval) = "HC  "
                                ALL_APPROVALS_ARRAY(app_mo, each_approval) = MAXIS_footer_month
                                ALL_APPROVALS_ARRAY(app_yr, each_approval) = MAXIS_footer_year
                                ALL_APPROVALS_ARRAY(app_done, each_approval) = TRUE

                                ALL_APPROVALS_ARRAY(hc_maj_prog, each_approval) = trim(major_prog)
                                ALL_APPROVALS_ARRAY(hc_elig_type, each_approval) = elig_type
                                ALL_APPROVALS_ARRAY(hc_elig_stnd, each_approval) = elig_stnd
                                ALL_APPROVALS_ARRAY(hc_elig_mthd, each_approval) = elig_mthd
                                ALL_APPROVALS_ARRAY(hc_waiv_type, each_approval) = replace(waiver_check, "_", "")

                                each_approval = each_approval + 1
                            End If

                        End If

                        transmit
                    End If
                Loop until current_version = "01" OR prog_status = "APP"
            End If
            row = row + 1
        Loop until clt_hc_prog = "    "


    End If
Next

'TESTING'
For the_approval = 0 to UBOUND(ALL_APPROVALS_ARRAY, 2)
    Call back_to_SELF
    MAXIS_footer_month = ALL_APPROVALS_ARRAY(app_mo, the_approval)
    MAXIS_footer_year = ALL_APPROVALS_ARRAY(app_yr, the_approval)

    Call navigate_to_MAXIS_screen("ELIG", ALL_APPROVALS_ARRAY(app_nav, the_approval))

    If ALL_APPROVALS_ARRAY(app_prog, the_approval) = "HC" Then
        MsgBox "APPROVAL: " & ALL_APPROVALS_ARRAY(app_prog, the_approval) & " for: " & ALL_APPROVALS_ARRAY(app_mo, the_approval) & "/" & ALL_APPROVALS_ARRAY(app_yr, the_approval) & vbNewLine &_
               "Navigation: ELIG/" & ALL_APPROVALS_ARRAY(app_nav, the_approval) & vbNewLine &_
               "Approval DONE - " & ALL_APPROVALS_ARRAY(app_done, the_approval) & vbNewLine & vbNewLine &_
               "Memb " & ALL_APPROVALS_ARRAY(elig_memb, the_approval) & " - " & ALL_APPROVALS_ARRAY(hc_maj_prog, the_approval) & " : " & ALL_APPROVALS_ARRAY(hc_elig_type, the_approval) & "-" & ALL_APPROVALS_ARRAY(hc_elig_stnd, the_approval)
    Else
        MsgBox "APPROVAL: " & ALL_APPROVALS_ARRAY(app_prog, the_approval) & " for: " & ALL_APPROVALS_ARRAY(app_mo, the_approval) & "/" & ALL_APPROVALS_ARRAY(app_yr, the_approval) & vbNewLine &_
               "Navigation: ELIG/" & ALL_APPROVALS_ARRAY(app_nav, the_approval) & vbNewLine &_
               "Approval DONE - " & ALL_APPROVALS_ARRAY(app_done, the_approval)
    End If
Next


'gather additional information for each array item (program and month)
'Need to be able to indicate if there is a change from the previous month to this month
'Need to determine if this ia approval, denial or closure
For the_approval = 0 to UBOUND(ALL_APPROVALS_ARRAY, 2)
    Call back_to_SELF

    If ALL_APPROVALS_ARRAY(app_done, the_approval) = TRUE Then
        MAXIS_footer_month = ALL_APPROVALS_ARRAY(app_mo, the_approval)
        MAXIS_footer_year = ALL_APPROVALS_ARRAY(app_yr, the_approval)

        Call navigate_to_MAXIS_screen("ELIG", ALL_APPROVALS_ARRAY(app_nav, the_approval))

        row = 1
        col = 1
        EMSearch "Command:", row, col
        EMWriteScreen ALL_APPROVALS_ARRAY(app_version, the_approval), row, col + 17
        transmit

        Select Case ALL_APPROVALS_ARRAY(app_prog, the_approval)

            Case "DWP"

            Case "MFIP"

            Case "MSA"

            Case "GA"

            Case "CASH"

            Case "GRH"

            Case "IVE"

            Case "EMER"

            Case "SNAP"

                'Read for the Eligibile Members and if their income is counted.
                elig_row = 7
                Do
                    EMReadScreen elig_ref_number, 2, elig_row, 10
                    EMReadScreen elig_memb_code, 12, elig_row, 35
                    EMReadScreen elig_status, 10, elig_row, 57

                    elig_memb_code = trim(elig_memb_code)
                    elig_status = trim(elig_status)

                    If elig_status = "ELIGIBLE" Then
                        ALL_APPROVALS_ARRAY(app_memb_detail, the_approval) = ALL_APPROVALS_ARRAY(app_memb_detail, the_approval) & "~" & elig_ref_number & " - " & elig_memb_code & " - " & elig_status
                        If ALL_APPROVALS_ARRAY(elig_memb, the_approval) = "" Then
                            ALL_APPROVALS_ARRAY(elig_memb, the_approval) = elig_ref_nbr
                        Else
                            ALL_APPROVALS_ARRAY(elig_memb, the_approval) = ALL_APPROVALS_ARRAY(elig_memb, the_approval) & ", " & elig_ref_nbr
                        End If
                    ElseIf elig_status = "INELIGIBLE" Then

                        ALL_APPROVALS_ARRAY(app_memb_detail, the_approval) = ALL_APPROVALS_ARRAY(app_memb_detail, the_approval) & "~" & elig_ref_number & " - " & elig_memb_code & " - " & elig_status

                        EMWriteScreen "X", elig_row, 5              'opening the person tests for ineligible members
                        transmit

                        EMReadScreen abawd_test, 6, 6, 20
                        EMReadScreen absence_test, 6, 7, 20
                        EMReadScreen roomer_test, 6, 8, 20
                        EMReadScreen boarder_test, 6, 9, 20
                        EMReadScreen citizenship_test, 6, 10, 20
                        EMReadScreen ctzn_coop_test, 6, 11, 20
                        EMReadScreen cmdty_test, 6, 12, 20
                        EMReadScreen disq_test, 6, 13, 20
                        EMReadScreen dupl_asst_test, 6, 14, 20

                        EMReadScreen fraud_test, 6, 6, 54
                        EMReadScreen student_test, 6, 7, 54
                        EMReadScreen institution_test, 6, 8, 54
                        EMReadScreen mfip_elig_test, 6, 9, 54
                        EMReadScreen non_applcnt_test, 6, 10, 54
                        EMReadScreen resi_test, 6, 11, 54
                        EMReadScreen ssn_coop_test, 6, 12, 54
                        EMReadScreen unit_mbr_test, 6, 13, 54
                        EMReadScreen work_reg_test, 6, 14, 54

                        If abawd_test = "FAILED" Then
                        If absence_test = "FAILED" Then
                        If roomer_test = "FAILED" Then
                        If boarder_test = "FAILED" Then
                        If citizenship_test = "FAILED" Then
                        If ctzn_coop_test = "FAILED" Then
                        If cmdty_test = "FAILED" Then
                        If disq_test = "FAILED" Then
                        If dupl_asst_test = "FAILED" Then

                        If fraud_test = "FAILED" Then
                        If student_test = "FAILED" Then
                        If institution_test = "FAILED" Then
                        If mfip_elig_test = "FAILED" Then
                        If non_applcnt_test = "FAILED" Then
                        If resi_test = "FAILED" Then
                        If ssn_coop_test = "FAILED" Then
                        If unit_mbr_test = "FAILED" Then
                        If work_reg_test = "FAILED" Then

                        transmit
                    End If

                    elig_row = elig_row + 1
                    If elig_row = 18 Then
                        PF8
                        elig_row = 7
                    End If

                Loop until elig_ref_number = "  " AND elig_memb_code = "" AND elig_status = ""

                transmit        'Now we are at 'Case Results'

                EMReadScreen appl_wthdrwn_clsd_test, 6, 7, 9
                EMReadScreen applcnt_elig_test, 6, 8, 9
                EMReadScreen comdty_test, 6, 9, 9
                EMReadScreen disq_test, 6, 10, 9
                EMReadScreen dupl_asst_test, 6, 11, 9
                EMReadScreen elig_prsn_test, 6, 12, 9
                EMReadScreen fail_coop_test, 6, 13, 9
                EMReadScreen fail_file_test, 6, 14, 9
                EMReadScreen prosp_gross_test, 6, 15, 9
                EMReadScreen prosp_net_test, 6, 16, 9

                EMReadScreen recert_test, 6, 7, 49
                EMReadScreen resi_test, 6, 8, 49
                EMReadScreen resource_test, 6, 9, 49
                EMReadScreen retro_gross_test, 6, 10, 49
                EMReadScreen retro_net_test, 6, 11, 49
                EMReadScreen strike_test, 6, 12, 49
                EMReadScreen transfer_resource_test, 6, 13, 49
                EMReadScreen verif_test, 6, 14, 49
                EMReadScreen vol_quit_test, 6, 15, 49
                EMReadScreen work_reg_test, 6, 16, 49


            Case "HC"


        End Select
    End If
Next



'Create a seperate case note for each program action
script_end_procedure_with_error_report("")
