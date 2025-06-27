'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - Individual Banked Months Notes.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 0			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================
'run_locally = TRUE
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
call changelog_update("06/05/2019", "Added a 'NOTES' section to add additional detail about action taken to the CASE//NOTE.", "Casey Love, Hennepin County")
call changelog_update("12/14/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""

'Checks to make sure we're in MAXIS
call check_for_MAXIS(True)

Call MAXIS_case_number_finder(MAXIS_case_number)
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 100, "Dialog"
  EditBox 70, 10, 60, 15, MAXIS_case_number
  EditBox 240, 10, 15, 15, MAXIS_footer_month
  EditBox 260, 10, 15, 15, MAXIS_footer_year
  EditBox 70, 30, 35, 15, HH_member_number
  DropListBox 160, 30, 115, 45, "Select One ..."+chr(9)+"WREG Updated - NO APPROVAL"+chr(9)+"Approval Done", note_to_make
  EditBox 40, 55, 235, 15, other_details
  ButtonGroup ButtonPressed
    OkButton 170, 80, 50, 15
    CancelButton 225, 80, 50, 15
  Text 15, 15, 50, 10, "Case Number:"
  Text 165, 15, 75, 10, "Initial Month Updated"
  Text 30, 35, 30, 10, "Member"
  Text 120, 35, 40, 10, "Note Type:"
  Text 10, 60, 30, 10, "Details:"
EndDialog

Do
    err_msg = ""

    Dialog Dialog1

    cancel_without_confirmation

    If MAXIS_case_number = "" Then err_msg = err_msg & vbNewLine & "* Enter a case number."
    If HH_member_number = "" Then err_msg = err_msg & vbNewLine & "* Enter the HH Member Number"
    If IsNumeric(MAXIS_case_number) = FALSE Then err_msg = err_msg & vbNewLine & "* * Case number appears to be invalid, check."
    If note_to_make = "Select One ..." Then err_msg = err_msg & vbNewLine & "* Choose which type of case note is needed."

    If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

Loop until err_msg = ""

start_month = MAXIS_footer_month
start_year = MAXIS_footer_year
HH_member_number = right("00"&HH_member_number, 2)
other_notes = ""
' If BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case) = "BANKED MONTH" Then
'     other_notes = other_notes & MAXIS_footer_month & "/" & MAXIS_footer_year & " for " & HH_memb & " is " & BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case) & " - Banked Month: " & month_tracker_nbr & ".; "
'
' Else
'     other_notes = other_notes & MAXIS_footer_month & "/" & MAXIS_footer_year & " for " & HH_memb & " is " & BANKED_MONTHS_CASES_ARRAY(month_indicator +9, the_case) & "; "
' End If
Do
    Call navigate_to_MAXIS_screen("STAT", "WREG")
    panel_date = cdate(MAXIS_footer_month & "/01/" & MAXIS_footer_year)
    If panel_date > cdate("6/30/2025") Then
        ET_col = 78
    Else
        ET_col = 80
    End If
    EMWriteScreen HH_member_number, 20, 76
    transmit

    EMReadScreen new_fset_status, 2, 8, 50
    EMReadScreen new_abawd_status, 2, 13, 50
    EMReadScreen new_bm_tracker, 1, 14, 50
    EMReadScreen new_fset_funds, 1, 8, ET_col

    If new_fset_status = "30" AND new_abawd_status = "13" Then
        other_notes = other_notes & " * " & MAXIS_footer_month & "/" & MAXIS_footer_year & " for MEMB " & HH_member_number & " is BANKED - Banked Month: " & new_bm_tracker & ".###"
    ElseIf new_fset_status = "30" AND new_abawd_status = "10" Then
        other_notes = other_notes & " * " & MAXIS_footer_month & "/" & MAXIS_footer_year & " for MEMB " & HH_member_number & " is REGULAR ABAWD.###"
    ElseIf new_fset_status <> "30" Then
        Select Case new_fset_status

        Case "03"
            clt_snap_status = "Unfit For Employment"
            exemption_type = "WREG & FSET"
        Case "04"
            clt_snap_status = "Responsible For Care Of Incapacitated Person"
            exemption_type = "WREG & FSET"
        Case "05"
            clt_snap_status = "Age 60 Or Older"
            exemption_type = "WREG & FSET"
        Case "06"
            clt_snap_status = "Under Age 16"
            exemption_type = "WREG & FSET"
        Case "07"
            clt_snap_status = "Age 16-17, Living W/Pare/Crgvr"
            exemption_type = "WREG & FSET"
        Case "08"
            clt_snap_status = "Responsible For Care Of Child < 6 Years Old"
            exemption_type = "WREG & FSET"
        Case "09"
            clt_snap_status = "Earnings At Least Min Wage X 30 Hrs/Wk"
            exemption_type = "WREG & FSET"
        Case "10"
            clt_snap_status = "Matching Grant Participant"
            exemption_type = "WREG & FSET"
        Case "11"
            clt_snap_status = "Receiving Or Applied For Unemployment"
            exemption_type = "WREG & FSET"
        Case "12"
            clt_snap_status = "Enrolled In School, Training, Or Higher Ed"
            exemption_type = "WREG & FSET"
        Case "13"
            clt_snap_status = "Participating In CD Program"
            exemption_type = "WREG & FSET"
        Case "14"
            clt_snap_status = "Receiving MFIP"
            exemption_type = "WREG & FSET"
        Case "20"
            clt_snap_status = "Pending/Receiving DWP"
            exemption_type = "WREG & FSET"
        Case "15"
            clt_snap_status = "Age 16-17 Not Lvg W/Pare/Crgvr"
            exemption_type = "FSET"
        Case "16"
            clt_snap_status = "50-59 Years Old"
            exemption_type = "FSET"
        Case "21"
            clt_snap_status = "Resp For Care Of Child < 18"
            exemption_type = "FSET"
        Case "17"
            clt_snap_status = "Receiving RCA/GA"
            exemption_type = "FSET"
        Case Else
            If new_abawd_status = "10" Then
                clt_snap_status = "Work Reg Exmpt"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "02" Then
                clt_snap_status = "Under Age 18"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "03" Then
                clt_snap_status = "Age 50 Or Over"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "04" Then
                clt_snap_status = "Caregiver Of Minor Child"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "05" Then
                clt_snap_status = "Pregnant"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "06" Then
                clt_snap_status = "Employed Avg Of 20 Hrs/Wk"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "07" Then
                clt_snap_status = "Wrk Experience Participant"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "08" Then
                clt_snap_status = "Other E & T Services"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "09" Then
                clt_snap_status = "Resides In A Waivered Area"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "11" Then
                clt_snap_status = "2nd set ABAWD"
                exemption_type = "ABAWD"
            ElseIf new_abawd_status = "12" Then
                clt_snap_status = "RCA/GA Recipient"
                exemption_type = "ABAWD"
            End If
        End Select

        other_notes = other_notes & " * " & MAXIS_footer_month & "/" & MAXIS_footer_year & " for MEMB " & HH_member_number & " - " & exemption_type & " Exempt ###     -" & clt_snap_status & "-###"

    End If

    this_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
    next_month = DateAdd("m", 1, this_month)

    MAXIS_footer_month = right("00" & DatePart("m", next_month), 2)
    MAXIS_footer_year = right("00" & DatePart("yyyy", next_month), 2)

    Call back_to_SELF

Loop until MAXIS_footer_month = CM_plus_2_mo AND MAXIS_footer_year = CM_plus_2_yr

MAXIS_footer_month = start_month
MAXIS_footer_year = start_year

UPDATED_ARRAY = split(other_notes, "###")

If note_to_make = "WREG Updated - NO APPROVAL" Then

    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("WREG Updated for ABAWD Information for M" & HH_member_number)
    Call write_variable_in_CASE_NOTE("Detail of current WREG Status:")
    For each detail_item in UPDATED_ARRAY
        Call write_variable_in_CASE_NOTE(detail_item)
    Next
    Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_details)
    ' Call write_bullet_and_variable_in_CASE_NOTE("Detail", other_notes)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

End If

If note_to_make = "Approval Done" Then

    If start_month = CM_plus_1_mo AND start_year = CM_plus_1_yr Then
        footer_month = CM_plus_1_mo
        footer_year = CM_plus_1_yr
    Else
        footer_month = CM_mo
        footer_year = CM_yr
    End If
    'setting the variables
    approval_note = ""
    'We are going to loop through each of the months from start month to CM + 1 to gather information from ELIG
    Do
        Call Navigate_to_MAXIS_screen("ELIG", "SUMM")       'Go to ELIG/SUMM
        EmWriteScreen footer_month, 19, 56                  'Go to the SNAP eligibility for the correct month and year
        EMWriteScreen footer_year, 19, 59
        EMWriteScreen "FS  ", 19, 71
        transmit

        elig_row = 7                                        'beginning of the list of members in the case
        list_of_fs_members = ""                             'creating a list of all the members
        Do
            EmReadscreen fs_memb, 2, elig_row, 10           'reading the member number, code and elig status
            EmReadscreen fs_memb_code, 1, elig_row, 35
            EmReadscreen fs_memb_elig, 8, elig_row, 57

            'These are when a member is active and eligible for SNAP on this case
            If fs_memb_code = "A" and fs_memb_elig = "ELIGIBLE" Then list_of_fs_members = list_of_fs_members & "~"& fs_memb

            elig_row = elig_row + 1     'looking at the next member
            EmReadscreen next_member, 2, elig_row, 10   'looking at if there is another member to review
        Loop until next_member = "  "                   'This would be the end of the list of members in ELIG
        'MsgBox "Line 947" & vbNewLine & "List of Members" & list_of_fs_members
        If list_of_fs_members <> "" Then
            list_of_fs_members = right(list_of_fs_members, len(list_of_fs_members)-1)   'This was assembled from reviewing ELIG
            member_array = split(list_of_fs_members, "~")       'making is an ARRAY
        End If

        transmit    'going to FSB1'
        transmit

        EmReadscreen total_earned_income, 9, 8, 32
        EmReadscreen total_unea_income, 9, 18, 32

        total_earned_income = trim(total_earned_income)
        total_unea_income = trim(total_unea_income)

        If total_earned_income = "" Then total_earned_income = 0
        If total_unea_income = "" Then total_unea_income = 0

        total_earned_income = FormatNumber(total_earned_income, 2, -1, 0, -1)
        total_unea_income = FormatNumber(total_unea_income, 2, -1, 0, -1)

        transmit    'going to FSB2'

        EmReadscreen total_shelter_costs, 9, 14, 28
        total_shelter_costs = trim(total_shelter_costs)
        If total_shelter_costs = "" Then total_shelter_costs = 0
        total_shelter_costs = FormatNumber(total_shelter_costs, 2, -1, 0, -1)
        'TODO add format number to each of these

        transmit    'going to FSSM'

        EmReadscreen fs_benefit_amount, 9, 13, 72
        EmReadscreen reporting_status, 9, 8, 31

        fs_benefit_amount = trim(fs_benefit_amount)
        If fs_benefit_amount = "" Then fs_benefit_amount = 0
        fs_benefit_amount = FormatNumber(fs_benefit_amount, 2, -1, 0, -1)
        reporting_status = trim(reporting_status)
        If fs_benefit_amount = 0 Then

            EmReadscreen fs_benefit_amount, 9, 10, 72
            fs_benefit_amount = trim(fs_benefit_amount)
            If fs_benefit_amount = "" Then fs_benefit_amount = 0
            fs_benefit_amount = FormatNumber(fs_benefit_amount, 2, -1, 0, -1)

        End If

        'Creating a list of each line of the case note - created here instead of adding to an array because we don't need it after the note
        approval_note = approval_note & "~!~* SNAP approved for " & footer_month & "/" & footer_year
        approval_note = approval_note & "~!~    Eligible Household Members: "
        For each person in member_array
            approval_note = approval_note & person & ", "
        Next
        approval_note = approval_note & "~!~    Income: Earned: $" & total_earned_income & " Unearned: $" & total_unea_income
        If total_shelter_costs <> "" Then  approval_note = approval_note & "~!~    Shelter Costs: $" & total_shelter_costs
        approval_note = approval_note & "~!~    SNAP BENEFTIT: $" & fs_benefit_amount & " Reporting Status: " & reporting_status

        first_of_footer_month = footer_month & "/01/" & footer_year     'there was no month in the spreadsheet
        next_month = DateAdd("m", 1, first_of_footer_month)                         'the month is advanded by ONE from what the last month we looked at was

        footer_month = DatePart("m", next_month)          'formatting the month and year and setting them for the nav functions to work
        footer_month = right("00"&footer_month, 2)

        footer_year = DatePart("yyyy", next_month)
        footer_year = right(footer_year, 2)

    Loop until footer_month = CM_plus_2_mo and footer_year = CM_plus_2_yr

    ARRAY_OF_NOTE_LINES = split(approval_note, "~!~")       'making this an array

    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("*** SNAP Approved starting in " & start_month & "/" & start_year & " ***")
    For each note_line in ARRAY_OF_NOTE_LINES
        Call write_variable_in_CASE_NOTE(note_line)
    Next
    Call write_variable_in_CASE_NOTE("Detail of current WREG Status:")
    For each detail_item in UPDATED_ARRAY
        Call write_variable_in_CASE_NOTE(detail_item)
    Next
    Call write_bullet_and_variable_in_CASE_NOTE("Notes", other_details)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

End If

script_end_procedure_with_error_report("ALl done")
