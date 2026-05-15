'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - TLR SCREENING.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 750                	'manual run time in seconds
STATS_denomination = "M"       		'M is for Member
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
call changelog_update("05/14/2026", "TLR Screening aligns with WREG and TLR evaluations. This update also aligns with TLR screening in the Interview Script.", "Casey Love, Hennepin County")
call changelog_update("12/10/2024", "Fixed bug in date-based ABAWD evaluation to work dynamically by footer month/year selected.", "Ilse Ferris, Hennepin County")
Call changelog_update("10/07/2024", "Added Age-Based exemption from 53-59 to 55-59 based on 10/2024 policy.", "Ilse Ferris, Hennepin County")
Call changelog_update("06/27/2024", "Added update handling for residents who meet military service ABAWD/TLR exemptions.", "Ilse Ferris, Hennepin County")
Call changelog_update("12/29/2023", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function determine_age (date_of_birth, age)
    If IsDate(date_of_birth) Then
        date_of_birth = DateAdd("d", 0, date_of_birth)

        birthdate_passed_this_year = False                                                          'identify if we need to adjust the age calculation based on if the birthday has passed this year or not
        birth_day = DatePart("d", date_of_birth)                                                    'create a date value for the birthday this year using the birth month and day and the current year
        birth_month = DatePart("m", date_of_birth)
        current_year = DatePart("yyyy", Date)
        this_year_birthday = DateSerial(current_year, birth_month, birth_day)
        ' NOTE THAT DateSerial will error 2/29 to 3/1 for non-leap years but this will create expected behavior and special handling is not needed.
        If DateDiff("d", Date, this_year_birthday) <= 0 Then birthdate_passed_this_year = True      'compare the current date to birthday this year to determine if the birthday has already passed this year or not
        age = DateDiff("yyyy", date_of_birth, Date)                                                 'calculate the age based on the difference in years between the date of birth and today
        If not birthdate_passed_this_year Then age = age - 1                                        'if the birthdate has not yet passed this year, subtract 1 from the age calculation to get the correct age
        date_of_birth = date_of_birth & ""
    End If
end function

'Dialogs===================================================================================================================
EMConnect ""
Call check_for_MAXIS(False)
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
member_number = "01"                        'some default settings
tlr_screen_save_btn = 100
CM_button       = 1000
TLR_CM_button   = 1100
include_panel_code = True
TLR_fixed_clock_string = TLR_fixed_clock_mo & "/" & TLR_fixed_clock_yr
info_evaluated = True

'Case number, member and month calendar dialog
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 181, 110, "ACTIONS - TLR SCREENING"
  Text 20, 15, 50, 10, "Case Number: "
  EditBox 75, 10, 45, 15, MAXIS_case_number
  Text 10, 35, 60, 10, "Member Number:"
  EditBox 75, 30, 30, 15, member_number
  Text 5, 55, 65, 10, "Footer month/year:"
  EditBox 75, 50, 20, 15, MAXIS_footer_month
  EditBox 100, 50, 20, 15, MAXIS_footer_year
  Text 10, 75, 60, 10, "Worker Signature:"
  EditBox 75, 70, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 15, 90, 70, 15, "Script Instructions", script_instructions
    OkButton 90, 90, 40, 15
    CancelButton 135, 90, 40, 15
EndDialog

Do
	Do
	    err_msg = ""
  		Dialog Dialog1
        Cancel_without_confirmation
  		Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If IsNumeric(member_number) = False or len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit member number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
        If ButtonPressed = script_instructions then
            call open_URL_in_browser("https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/_layouts/15/Doc.aspx?sourcedoc=%7B58B17691-CF4B-4EBF-8B97-B556E995F67D%7D&file=ACTIONS%20-%20TLR%20SCREENING.docx")
            err_msg = "LOOP" & err_msg
        End if
		IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call MAXIS_footer_month_confirmation	'making sure we're getting to current month/year
Call MAXIS_background_check
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged, and you do not have access. The script will now end.")

Call write_value_and_transmit(member_number, 20, 76)
EMReadScreen memb_exists_check, 13, 8, 22
If memb_exists_check = "Arrival Date:" Then
    PF3
    PF10
    transmit
    script_end_procedure("The member number (MEMB " & member_number & ") that you entered is not listed on this case.  Please check the member number, and start the script again.")
End If
EmReadScreen first_name, 12, 6, 63
first_name = replace(first_name, "_", "")
EmReadScreen last_name, 25, 6, 30
last_name = replace(last_name, "_", "")
EMReadScreen dob, 10, 8, 42

member_info = member_number & " - " & first_name & " " & last_name

call tlr_screening_read_WREG_details(member_number, wreg_exists, curr_wreg_code, curr_tlr_code, counted_months, second_set_count, counted_months_string, second_set_string, tlr_output, second_set_output)

'filling the drop lists for some screening questions.
'NOTE - there is a function to fill 'all_the_clients' but it does not work for membs_under_6 and we would need to repeat the code anyway so it is done without the function
memb_row = 5
all_the_clients = "Select or Type"
membs_under_6 = "Select or Type"
Call navigate_to_MAXIS_screen ("STAT", "MEMB")
Do
    EMReadScreen ref_numb, 2, memb_row, 3
    If ref_numb = "  " Then Exit Do
    EMWriteScreen ref_numb, 20, 76
    transmit
    EMReadScreen first_name, 12, 6, 63
    EMReadScreen last_name, 25, 6, 30
	EMReadScreen cl_age, 3, 8, 76
    cl_age = trim(cl_age)
    If cl_age = "" then cl_age = 0
    cl_age = cl_age*1
    all_the_clients = all_the_clients+chr(9)+ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
    If cl_age < 6 Then membs_under_6 = membs_under_6+chr(9)+ref_numb & " - " & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
    memb_row = memb_row + 1
Loop until memb_row = 20

'Reviewing member details to determine tlr and wreg status.
'This call is mostly to identify age based exemptions and shorten the screening process for age related exemptions
call tlr_screening_determine_wreg_status(info_evaluated, screening_needed, eval_string, wreg_status, tlr_status, panel_code, dob, disability_info, homeless, dv_victim, schl_training, person_requiring_care, person_care_reason, child_under_6_requiring_care, caregiver_in_home, hrs_per_week, wage_per_week, wage_per_month, other_benefit, unea_income, treatment, children_under_14, pregnant, american_indian, wreg_selection)
If tlr_status = "Exempt" Then
    call determine_age(dob, cl_age)
    age_exempt_msg = member_info & " is currently " & cl_age & " years old and meets an exemption." & vbCr & vbCr
    age_exempt_msg = age_exempt_msg & "WREG - " & wreg_status & ", TLR - " & tlr_status & vbCr
    If right(trim(eval_string), 1) = ";" Then eval_string = left(trim(eval_string), len(trim(eval_string)) - 1)
    If InStr(eval_string, ";") Then
        eval_array = split(eval_string, ";")
        for each eval_info in eval_array
            If trim(eval_info) <> "" Then
                age_exempt_msg = age_exempt_msg & " - " & trim(eval_info) & vbCr
            End If
        next
    Else
        age_exempt_msg = age_exempt_msg & " - " & trim(eval_string) & vbCr
    End If
    age_exempt_msg = age_exempt_msg & vbCr & "Additional screening is not required." & vbCr
    age_exempt_msg = age_exempt_msg & vbCr & "Do you want to CASE/NOTE the age based exemption information?"

    case_note_age_exemption = MsgBox(age_exempt_msg, vbQuestion + vbYesNo, "Member meets age-based exemption criteria.")

    If case_note_age_exemption = vbYes then
        call start_a_blank_CASE_NOTE
        call write_variable_in_CASE_NOTE("*~*~ M" & member_info & " meets age-based TLR exemption criteria. ~*~*")
        call write_variable_in_CASE_NOTE("TLR Screening completed on " & date & ".")
        call tlr_screening_case_note_person_info(member_info, info_evaluated, eval_string, wreg_status, tlr_status, panel_code, wreg_selection, include_panel_code, tlr_output, second_set_output, wreg_exists)
        Call write_variable_in_CASE_NOTE("---")
        Call write_variable_in_CASE_NOTE(worker_signature)
    End if
    Call script_end_procedure_with_error_report("The TLR screening for " & member_info & " is now complete. The member meets age-based exemption criteria, and additional screening is not required.")
End If

'main screening dialog has 2 parts, the questions and screening options, then summary and confirmation with the ability to add notes and select options
big_err_msg = ""
Do
    Do
        If big_err_msg = "" Then
            Do
                err_msg = ""

                BeginDialog Dialog1, 0, 0, 485, 345, "TLR / WREG Screening Questions"
                    ButtonGroup ButtonPressed
                        Text 15, 5, 260, 20, "Record information about the member to have the script evaluate for possible SNAP TLR or WREG exemptions for:"
                        call tlr_screening_wreg_person_info_display(info_evaluated, member_info, all_the_clients, membs_under_6, eval_string, wreg_status, tlr_status, panel_code, dob, disability_info, homeless, dv_victim, schl_training, schl_amt, person_requiring_care, person_care_reason, child_under_6_requiring_care, caregiver_in_home, hrs_per_week, wage_per_week, wage_per_month, other_benefit, unea_income, treatment, children_under_14, pregnant, american_indian, wreg_selection, tlr_screen_save_btn, tlr_output, second_set_output, wreg_exists)
                        CancelButton 400, 320, 75, 15
                        PushButton 375, 5, 85, 15, "Exempt - CM 0028.06.12", CM_button
                        PushButton 375, 20, 85, 15, "TLR - CM 0011.24", TLR_CM_button
                EndDialog

                dialog Dialog1
                cancel_confirmation
                If ButtonPressed = -1 Then ButtonPressed = tlr_screen_save_btn

                hrs_per_week = trim(hrs_per_week)
                If hrs_per_week <> "" then
                    If IsNumeric(hrs_per_week) = False then err_msg = err_msg & vbNewLine & "* Enter a valid number of hours per week."
                End if
                wage_per_week = trim(wage_per_week)
                If wage_per_week <> "" then
                    If IsNumeric(wage_per_week) = False then err_msg = err_msg & vbNewLine & "* Enter a valid number of wage in a week."
                End if
                wage_per_month = trim(wage_per_month)
                If wage_per_month <> "" then
                    If IsNumeric(wage_per_month) = False then err_msg = err_msg & vbNewLine & "* Enter a valid number of wage in a month."
                End if
                If ButtonPressed = CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00280612"
                If ButtonPressed = TLR_CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe	https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_001124"
                If ButtonPressed = CM_button or ButtonPressed = TLR_CM_button Then err_msg = "LOOP"
                If err_msg <> "" and err_msg <> "LOOP" Then MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg

            Loop until err_msg = ""
        End If

        'use the screening details to determine the TLR and WREG status and display the details in a summary dialog.
        If schl_amt <> "" Then schl_training = schl_training & " - " & schl_amt     ' the school type and enrollment amount are separate in the display and this puts them back together.
        call tlr_screening_determine_wreg_status(info_evaluated, screening_needed, eval_string, wreg_status, tlr_status, panel_code, dob, disability_info, homeless, dv_victim, schl_training, person_requiring_care, person_care_reason, child_under_6_requiring_care, caregiver_in_home, hrs_per_week, wage_per_week, wage_per_month, other_benefit, unea_income, treatment, children_under_14, pregnant, american_indian, wreg_selection)
        If right(trim(eval_string), 1) = ";" Then eval_string = left(trim(eval_string), len(trim(eval_string)) - 1)
        dlg_hgt = 155
        If InStr(eval_string, ";") Then
            eval_array = split(eval_string, ";")
            dlg_hgt = dlg_hgt + (10*(ubound(eval_array)+1))
        End if

        'Summary and notes or other option selections
        BeginDialog Dialog1, 0, 0, 450, dlg_hgt, "TLR / WREG Screening Summary"
            ButtonGroup ButtonPressed
                Text 15, 10, 155, 10, member_info
                y_pos = 25
                If panel_code <> "Panel Not Needed" and trim(eval_string) <> "" Then
                    Text 20, y_pos, 250, 10, "Panel Code based on Screening: " & panel_code
                    y_pos = y_pos + 10
                End If
                If wreg_exists Then
                    wreg_display_details = "WREG Panel Counts: "
                    wreg_display_details = wreg_display_details & " Counted Months: " & counted_months
                    If counted_months <> 0 Then wreg_display_details = wreg_display_details & " - " & counted_months_string
                    wreg_display_details = wreg_display_details & ",  Second Set Months: " & second_set_count
                    If second_set_count <> 0 Then wreg_display_details = wreg_display_details & " - " & second_set_string
                    Text 20, y_pos, 400, 10, wreg_display_details
                    y_pos = y_pos + 10
                End If
                y_pos = y_pos + 5
                screening_box_pos = y_pos
                y_pos = y_pos + 10
                If wreg_selection <> "" and wreg_selection <> "Unknown or Not Selected" Then
                    Text 30, y_pos, 350, 10, "WREG/TLR Selected Directly: " & wreg_selection
                    y_pos = y_pos + 10
                End If
                If trim(eval_string) = "" Then
                    Text 30, y_pos, 350, 10, "No WREG/TLR determination details to display."
                    y_pos = y_pos + 10
                Else
                    If InStr(eval_string, ";") Then
                        Text 30, y_pos, 350, 10, "WREG - " & wreg_status & ", TLR - " & tlr_status
                        y_pos = y_pos + 10
                        for each eval_info in eval_array
                            If trim(eval_info) <> "" Then
                                Text 40, y_pos, 400, 10, trim(eval_info)
                                y_pos = y_pos + 10
                            End If
                        next
                    Else
                        Text 30, y_pos, 350, 10, "WREG - " & wreg_status & ", TLR - " & tlr_status & "  --  " & trim(eval_string)
                        y_pos = y_pos + 10
                    End If
                End If
                GroupBox 20, screening_box_pos, 410, y_pos-screening_box_pos+5, "TLR Screening Evaluation Details"
                y_pos = y_pos + 10

                Text 10, y_pos+5, 60, 10, "Exemption notes:"
                EditBox 70, y_pos, 360, 15, exemption_notes
                y_pos = y_pos + 20
                Text 10, y_pos+5, 60, 10, "Exemption basis:"
                ComboBox 70, y_pos, 150, 15, "Select OR Type..."+chr(9)+"Conversation w/ resident"+chr(9)+"Observational"+chr(9)+"Verified", exemption_basis
                y_pos = y_pos + 20
                CheckBox 10, y_pos, 400, 10, "Check here to update STAT/WREG with highest exemption for " & MAXIS_footer_Month & "/" & MAXIS_footer_year & " through current monthly plus 1.", update_wreg_checkbox

                PushButton 360, 5, 85, 15, "Update Screening", update_btn
                CancelButton 300, dlg_hgt-20, 50, 15
                PushButton 355, dlg_hgt-20, 85, 15, "Screening Completed", complete_btn
        EndDialog

        dialog Dialog1
        cancel_confirmation
        If ButtonPressed = -1 Then ButtonPressed = complete_btn

        exemption_basis = trim (exemption_basis)
        exemption_notes = trim(exemption_notes)
        big_err_msg = ""
        If exemption_basis = "" Then big_err_msg = big_err_msg & vbCr & "* Provide detail of how exemption information was gathered in 'Exemption Basis' field."
        If exemption_basis = "Select OR Type..." Then big_err_msg = big_err_msg & vbCr & "* Provide detail of how exemption information was gathered in 'Exemption Basis' field."
        If ButtonPressed = update_btn Then big_err_msg = ""
        If big_err_msg <> "" Then MsgBox "*** NOTICE!!! ***" & vbNewLine & big_err_msg
    Loop until ButtonPressed = complete_btn and big_err_msg = ""
    CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

'Dialog will only show up if NO exemptions have been found AND 3 ABAWD counted months have been used since the start of the fixed TLR calendar.
If (tlr_status <> "Exempt" and counted_months => 3) then
    '----------------------------------------------------------------------------------------------------Second Set TLR
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 326, 180, "Second Set TLR Months"
      GroupBox 5, 5, 315, 55, "TLR Information for " & member_info
      Text 15, 20, 210, 20, "Counted TLR Months: " & tlr_output
      Text 15, 40, 230, 10, "Counted 2nd Set: " & second_set_output
      Text 5, 65, 315, 10, "=============================================================================="
      Text 60, 80, 190, 10, "Select ALL applicable situations for this member below"
      Text 5, 95, 315, 10, "=============================================================================="
      CheckBox 10, 105, 160, 10, "Used all 3 counted TLR months since " & TLR_fixed_clock_string & "?", Used_TLR_checkbox
      CheckBox 10, 120, 260, 10, "Worked at least 80 hours in a month SINCE closing for using 3 TLR months?", worked_80_since_closing
      CheckBox 10, 135, 250, 10, "Work/work activities have ended or reduced to less than 80 hours/month?", work_ended_checkbox
      ButtonGroup ButtonPressed
        'PushButton 10, 155, 110, 15, "View ABAWD Tracking Record", open_ATR_button
        OkButton 210, 155, 50, 15
        CancelButton 265, 155, 50, 15
        PushButton 230, 15, 85, 15, "Exempt - CM 0028.06.12", CM_button
        PushButton 250, 35, 65, 15, "TLR - CM 0011.24", TLR_CM_button
    EndDialog

    Do
	    Do
	        err_msg = ""
  		    Dialog Dialog1
  		    Cancel_confirmation
            If ButtonPressed = CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00280612"
            If ButtonPressed = TLR_CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe	https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_001124"
	    LOOP UNTIL ButtonPressed = -1
	    CALL check_for_password(are_we_passworded_out) 'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					 'loops until user passwords back in
    'determining 2nd set eligibility based on meeting all criteria in "Second Set TLR Months" dialog
End if

'determine if second set is available
second_set_test = True
If Used_TLR_checkbox = 0 then second_set_test = False
If worked_80_since_closing = 0 then second_set_test = False
If work_ended_checkbox = 0 then second_set_test = False
If second_set_test Then panel_code = "30/11"

'If resident doesn't meet any exemptions providing the option to case note.
Do
    If tlr_status <> "Exempt" then
        end_before_note_msg = Member_info & " has been identified as not meeting an exemption, "
        If second_set_test Then end_before_note_msg = end_before_note_msg & "and is eligible for TLR/ABAWD 2nd Set Months."
        If NOT second_set_test Then end_before_note_msg = end_before_note_msg & "and is NOT eligible for TLR/ABAWD 2nd Set Months."
        If update_wreg_checkbox = 1 then end_before_note_msg = end_before_note_msg & vbCr & vbCr & "Cancelling before a CASE/NOTE will skip the WREG update."
        end_before_note_msg = end_before_note_msg & vbCr & vbCr & "Do you want to CASE/NOTE this information?"
        end_before_note_msg = end_before_note_msg & vbCr & vbCr & "Press No to end the script without CASE/NOTE."
        case_note_confirmation = MsgBox(end_before_note_msg, vbInformation + vbYesNo, "Member appears to be a Time-Limited Recipient.")
        If case_note_confirmation = vbNo then script_end_procedure_with_error_report("You have opted out of case noting the TLR screening. The script has ended.")
    End if
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'This will update WREG from MAXIS_footer_month from the initial dialog through the current month plus 1 with the appropriate TLR/ABAWD code based on the screening if the user selected the option to update WREG. This is intended to ensure that the member's WREG panel in MAXIS reflects the most up to date information based on the screening and to avoid having to remember to update WREG separately from this script. It will also add a note in the case note about what was updated in WREG for reference.
If update_wreg_checkbox = 1 then
    Call date_array_generator(MAXIS_footer_month, MAXIS_footer_year, footer_month_array) 'Uses the custom function to create an array of dates from the initial_month and initial_year variables, ends at CM + 1.

    For footer_months = 0 to ubound(footer_month_array)
        call convert_date_into_MAXIS_footer_month(footer_month_array(footer_months), MAXIS_footer_month, MAXIS_footer_year)
        Call MAXIS_background_check

        Call navigate_to_MAXIS_screen("STAT", "WREG")
        Call write_value_and_transmit(member_number, 20, 76)
        EMReadScreen panel_exists, 1, 2, 78
        panel_date = cdate(MAXIS_footer_month & "/01/" & MAXIS_footer_year)
        PWE_col = 68
        ET_col = 80
        If panel_date > cdate("6/30/2025") Then
            PWE_col = 70
            ET_col = 78
        End If

        If panel_exists = "0" then
            Call write_value_and_transmit("NN", 20, 79) 'Adding new WREG panel
            EMWriteScreen "Y", 6, PWE_col 'defaulting PWE to Y if blank panel
        Else
            PF9
        End if

        best_wreg_code = left(panel_code, 2)
        best_tlr_code = right(panel_code, 2)

	    EMWriteScreen best_wreg_code, 8, 50
	    EMWriteScreen best_tlr_code, 13, 50
	    If best_wreg_code = "30" then
            If best_tlr_code = "09" then
                EMWriteScreen "Y", 8, ET_col
            Else
	            EMWriteScreen "N", 8, ET_col
            End if
	    Else
	        EMWriteScreen "_", 8, ET_col
	    End if

	    EMReadScreen orientation_warning, 7, 24, 2 	'reading for orientation date warning message. This message has been causing me TROUBLE!!
	    If orientation_warning = "WARNING" then transmit
	    PF3 'to save and exit to stat/wrap
    Next
End If

call start_a_blank_CASE_NOTE
' call write_variable_in_CASE_NOTE("----- TLR Screening Information for MEMB " & member_number & " -----")          'this is the phrasing used in the interview run. Keeping it here to reference if we want to align them
Call write_variable_in_CASE_NOTE("*~*~ M" & member_info & " TLR Screened: " & tlr_status & " ~*~*")
call write_variable_in_CASE_NOTE("TLR Screening completed on " & date & ".")
call write_variable_in_CASE_NOTE("--- TLR Screening is Based on " & exemption_basis & " ---")

call tlr_screening_case_note_person_info(member_info, info_evaluated, eval_string, wreg_status, tlr_status, panel_code, wreg_selection, include_panel_code, tlr_output, second_set_output, wreg_exists)

Call write_bullet_and_variable_in_CASE_NOTE("Exemption Notes", exemption_notes)
If counted_months => 3 then
    Call write_variable_in_CASE_NOTE("* Member has used all available counted TLR/ABAWD months.")
    If second_set_test Then Call write_variable_in_CASE_NOTE("* Member is eligible for TLR/ABAWD 2nd set months.")
    If NOT second_set_test Then Call write_variable_in_CASE_NOTE("* Member is NOT eligible for TLR/ABAWD 2nd set months.")
End if
If update_wreg_checkbox = 1 then Call write_variable_in_CASE_NOTE("* STAT/WREG panel has been updated with FSET/TLR codes: " & panel_code & ".")

call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

end_msg = "The TLR screening for " & member_info & " is now complete."
end_msg = end_msg & vbCr & vbCr & "Panel Code: " & panel_code & "."
end_msg = end_msg & vbCr & "WREG status: " & wreg_status & "."
end_msg = end_msg & vbCr & "TLR status: " & tlr_status & "."
end_msg = end_msg & vbCr & vbCr & "TLR Months Used: " & tlr_output
end_msg = end_msg & vbCr & "Second Set Months Used: " & second_set_output
If second_set_test Then end_msg = end_msg & vbCr & member_info & " is eligible for Second Set TLR Months"
If update_wreg_checkbox = 1 then end_msg = end_msg & vbCr & "WREG Updated with screening detail."

script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs --------------------------------------------------05/14/2026
'--Tab orders reviewed & confirmed-----------------------------------------------05/14/2026
'--Mandatory fields all present & Reviewed---------------------------------------05/14/2026
'--All variables in dialog match mandatory fields--------------------------------05/14/2026
'Review dialog names for content and content fit in dialog-----------------------05/14/2026
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog--------------------05/14/2026
'--Create a button to reference instructions-------------------------------------05/14/2026
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)----------------------------------05/14/2026
'--CASE:NOTE Header doesn't look funky-------------------------------------------05/14/2026
'--Leave CASE:NOTE in edit mode if applicable------------------------------------05/14/2026
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used------05/14/2026
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed---------------------------------------05/14/2026
'--MAXIS_background_check reviewed (if applicable)-------------------------------05/14/2026
'--PRIV Case handling reviewed --------------------------------------------------05/14/2026
'--Out-of-County handling reviewed-----------------------------------------------05/14/2026---------------------N/A
'--script_end_procedures (w/ or w/o error messaging)-----------------------------05/14/2026
'--BULK - review output of statistics and run time/count (if applicable)---------05/14/2026---------------------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")------------05/14/2026
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed ---------------------------------------------------05/14/2026
'--Incrementors reviewed (if necessary)------------------------------------------05/14/2026
'--Denomination reviewed --------------------------------------------------------05/14/2026
'--Script name reviewed----------------------------------------------------------05/14/2026
'--BULK - remove 1 incrementor at end of script reviewed-------------------------05/14/2026---------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete-----------------------------------------05/14/2026
'--comment Code------------------------------------------------------------------05/14/2026
'--Update Changelog for release/update-------------------------------------------05/14/2026
'--Remove testing message boxes--------------------------------------------------05/14/2026
'--Remove testing code/unnecessary code------------------------------------------05/14/2026
'--Review/update SharePoint instructions-----------------------------------------05/14/2026
'--Other SharePoint sites review (HSR Manual, etc.)------------------------------05/14/2026
'--COMPLETE LIST OF SCRIPTS reviewed---------------------------------------------05/14/2026
'--COMPLETE LIST OF SCRIPTS update policy references-----------------------------05/14/2026
'--Complete misc. documentation (if applicable)----------------------------------05/14/2026
'--Update project team/issue contact (if applicable)-----------------------------05/14/2026
