'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - COLA Review and Approve.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at 1
STATS_manualtime = 90          'manual run time in seconds
STATS_denomination = "C"       'C is for each case

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
call changelog_update("05/31/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT======================================================================================================================
'EMConnect ""
approval_exists = False

'Find the footer month/year
EMReadScreen cola_footer_month, 2, 6, 11
EMReadScreen cola_footer_year, 2, 6, 14

If cola_footer_month = CM_plus_1_mo AND cola_footer_year = CM_plus_1_yr Then
    MAXIS_footer_month = CM_plus_1_mo
    MAXIS_footer_year = CM_plus_1_yr
Else
    MAXIS_footer_month = CM_mo
    MAXIS_footer_year = CM_yr
End If

'possibly add dialog to get footer month and year.
EMWriteScreen "E", 6, 3                         'Navigates to ELIG/HC - maintaining tie to the DAIL for ease of processin
transmit
EMWriteScreen "HC", 20, 71
transmit

EMReadScreen hc_elig_check, 4, 3, 51
If hc_elig_check <> "HHMM" Then approval_exists = FALSE
EMWriteScreen MAXIS_footer_month, 20, 56            'Goes to the next month and checks that elig results exist
EMWriteScreen MAXIS_footer_year,  20, 59
transmit
If hc_elig_check <> "HHMM" Then approval_exists = FALSE

row = 8                                          'Reads each line of Elig HC to find all the approved programs in a case
Do
	save_info = False
    EMReadScreen clt_ref_num, 2, row, 3
    EMReadScreen clt_hc_prog, 4, row, 28
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
            If prog_status <> "APP" Then                        'Finding the approved version
                EMReadScreen total_versions, 2, row, 64
                If total_versions = "01" Then
                    error_processing_msg = error_processing_msg & vbNewLine & "Appears HC eligibility was not approved in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " for " & clt_ref_num & ", please approve HC and rerun script."
                    Exit Do
                Else
                    EMReadScreen current_version, 2, row, 58
                    If current_version = "01" Then
                        error_processing_msg = error_processing_msg & vbNewLine & "Appears HC eligibility was not approved in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " for " & clt_ref_num & ", please approve HC and rerun script."
                        Exit Do
                    Else
                        current_version = abs(current_version)
                        current_version = current_version - 1
                        prev_version = right ("00" & current_version, 2)
                        EMWriteScreen prev_version, row, 58
                        transmit
                    End If
                End If
            Else
                EMReadScreen elig_result, 8, row, 41        'Goes into the elig version to get the major program and elig type
                EMWriteScreen "x", row, 26
                transmit

                EMReadScreen process_date, 8, 2, 73
                If DateDiff("d", process_date, date) > 7 Then error_processing_msg =  error_processing_msg & vbNewLine & "HC Eligibility was not created and approved recently for " & clt_ref_num & "."

				col = 19
				Do
	                EMReadScreen bdgt_month, 2, 6, col
	                EMReadScreen bdgt_year, 2, 6, col + 3
					col = col + 11
					If bdgt_month = MAXIS_footer_month AND bdgt_year = MAXIS_footer_year Then Exit Do
				Loop until col = 85
                elig_col = 19

				' MsgBox "Budg - " & bdgt_month & "/" & bdgt_year & vbCr & "MAXIS - " & MAXIS_footer_month & "/" & MAXIS_footer_year
                If bdgt_month = MAXIS_footer_month AND bdgt_year = MAXIS_footer_year Then
                    approval_exists = TRUE
					save_info = True
                Else
                    Do
                        EMReadScreen elig_mo, 2, 6, elig_col
                        EMReadScreen elig_yr, 2, 6, elig_col + 3

                        If elig_mo = MAXIS_footer_month AND elig_yr = MAXIS_footer_year Then
                            Exit Do
                        Else
                            elig_col = elig_col + 11
                        End If

                    Loop Until elig_col = 85
                End If

                If elig_col < 85 Then
                    EMReadScreen waiver_check, 1, 14, elig_col + 2        'Checking to see if case may be LTC or Waiver'
                    EMReadScreen method_check, 1, 13, elig_col + 2
                    If method_check = "L" or method_check = "S" Then LTC_case = TRUE
                    If method_check = "B" AND waiver_check <> "_" Then LTC_case = TRUE
                    Do
                        transmit
                        EMReadScreen hc_screen_check, 8, 5, 3
                    Loop until hc_screen_check = "Program:"
                    If clt_hc_prog = "SLMB" OR clt_hc_prog = "QMB " Then
                        EMReadScreen elig_type, 2, 13, 78
                        EMReadScreen Majr_prog, 2, 14, 78
                    End If
                    If clt_hc_prog = "MA  " Then
                        EMReadScreen elig_type, 2, 13, 76
                        EMReadScreen Majr_prog, 2, 14, 76
                    End If
                ' Else
                '     approval_exists = FALSE
                End if
                transmit
            End If
            'MsgBox "Current Version:" & current_version & vbNewLine & "Prog Status: " & prog_status
        Loop until current_version = "01" OR prog_status = "APP"
        EMReadScreen panel_name, 4, 3, 51
        Do while panel_name <> "HHMM"
            transmit
            EMReadScreen panel_name, 4, 3, 51
        Loop
        'Adds everything to a varriable so an array can be created
        If save_info = True Then Elig_Info_array = Elig_Info_array & "~Memb " & clt_ref_num & " is approved as " & trim(elig_result) & " for " & trim(clt_hc_prog) & " : " & Majr_prog & "-" & elig_type
    End If
    If LTC_case = TRUE Then                 'LTC/Waiver cases have their own MA Approval script that will run if worker says yes
        run_LTC_Approval = msgbox ("It appears this case is LTC MA or Waiver MA." & vbNewLine & "Would you like to run the NOTES - LTC MA Approval Script for more detailed case noting?", vbYesNo + vbQuestion, "Run LTC Specific Script?")
        If run_LTC_Approval = vbYes Then    'Script will define some variables to carry to the next script for ease of use
            budget_type = method_check
            approved_check = checked
            Exit Do
        Else
            LTC_case = FALSE
        End If
    End If
    row = row + 1
Loop until clt_hc_prog = "    "

If run_LTC_Approval = vbYes Then call run_from_GitHub( script_repository & "notes/ltc-ma-approval.vbs")

'MsgBox approval_exists
If approval_exists = TRUE AND error_processing_msg <> "" Then
	Continue_with_noting = MsgBox("This case does not have approvals completed for all Household Members that appear to have MA Eligibility. The script can still note the approval information for some of the members." & vbCr & vbCr & "Detail about Members that cannot be processed:" & vbCr & error_processing_msg & vbCr & vbCr & "Would you like to continue with processing ONLY the other Members?", vbQuestion + vbYesNo, "Continue for Only Some Household Members")
	If Continue_with_noting = vbNo Then approval_exists = FALSE
End If
PF3             'Back to DAIL
' If error_processing_msg <> "" Then approval_exists = FALSE

If approval_exists = TRUE Then


    'Creates an array of all the HC approvals
    Elig_Info_array = right(Elig_Info_array, len(Elig_Info_array) - 1)
    Elig_Info_array = Split(Elig_Info_array, "~")

    'Array to determine which to case note
    Dim elig_checkbox_array()
    ReDim elig_checkbox_array(0)

    array_counter = 0

    For i = 0 to UBound(Elig_Info_array)
    	ReDim Preserve elig_checkbox_array(i)
    	elig_checkbox_array(i) = checked
    Next

    'Dialog is defined here as it is dynamic
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 286, 135 + (15 * UBound(Elig_Info_array)), "Approval dialog"
      For each elig_approval in Elig_Info_array
        CheckBox 10, 40 + (15 * array_counter), 265, 10, elig_approval, elig_checkbox_array(array_counter)
    	array_counter = array_counter + 1
      Next
      CheckBox 5, 60 + (15 * UBound(Elig_Info_array)), 220, 10, "Check here if you have reviewed/updated MMIS and it is correct", mmis_checkbox
      EditBox 65, 75 + (15 * UBound(Elig_Info_array)), 215, 15, cola_notes
      EditBox 65, 95 + (15 * UBound(Elig_Info_array)), 215, 15, other_notes
      EditBox 65, 115 + (15 * UBound(Elig_Info_array)), 90, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 175, 115 + (15 * UBound(Elig_Info_array)), 50, 15
        CancelButton 230, 115 + (15 * UBound(Elig_Info_array)), 50, 15
      Text 5, 5, 50, 10, "Case Number:"
      Text 60, 5, 30, 10, MAXIS_case_number
      Text 120, 5, 55, 10, "Approval Month:"
      Text 180, 5, 25, 10, MAXIS_footer_month & "/" & MAXIS_footer_year
      Text 5, 25, 275, 10, "Script has identified the following HC Approvals. They will be case noted if checked."
      Text 5, 80 + (15 * UBound(Elig_Info_array)), 55, 10, "COLA Details:"
      Text 5, 100 + (15 * UBound(Elig_Info_array)), 60, 10, "Other Notes:"
      Text 5, 120 + (15 * UBound(Elig_Info_array)), 60, 10, "Worker signature:"
    EndDialog

    Do
        err_msg = ""
        Dialog Dialog1
        cancel_confirmation
        If cola_notes = "" Then err_msg = err_msg & vbNewLine & "* Indicate information about the COLA processing completed."
        If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note"
        if err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""

    EMWriteScreen "N", 6, 3         'Goes to Case Note - maintains tie with DAIL
    transmit

    'Starts a blank case note
    PF9
    EMReadScreen case_note_mode_check, 7, 20, 3
    If case_note_mode_check <> "Mode: A" then script_end_procedure("You are not in a case note on edit mode. You might be in inquiry. Try the script again in production.")

    'Adding information to case note
    Call write_variable_in_CASE_NOTE ("---Approved HC - COLA reviewed---")
    Call write_variable_in_CASE_NOTE ("* Processed HC in " & MAXIS_footer_month & "/" & MAXIS_footer_year & " after reviewing COLA.")
    For array_item = 0 to UBound(Elig_Info_array)
        If elig_checkbox_array(array_item) = checked Then Call write_variable_in_CASE_NOTE ("* " & Elig_Info_array(array_item))
    Next
    Call write_bullet_and_variable_in_CASE_NOTE ("COLA Details", cola_notes)
    Call write_bullet_and_variable_in_CASE_NOTE ("Notes", other_notes)
    call write_variable_in_CASE_NOTE("---")
    call write_variable_in_CASE_NOTE(worker_signature)

    MAXIS_case_number = trim(MAXIS_case_number)

    end_msg = "Case note of HC Approval for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " completed." & vbNewLine & vbNewLine & "The COLA review and approval has been completed and the DAIL can be deleted."
End If

If approval_exists = FALSE Then

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 206, 120, "Program Approved other than HC"
      DropListBox 35, 70, 120, 40, "No other program 'APP'ed"+chr(9)+"Run Approved Programs"+chr(9)+"Run Closed Programs", forward_progress
      ButtonGroup ButtonPressed
        OkButton 150, 100, 50, 15
      Text 10, 10, 165, 25, "This case does not appear to have an approved version of Health Care Eligibility. This script is built for handling Health Care cases. "
      Text 10, 45, 170, 20, "If you have approved a different program, the script can run Approved Programs or Closed Programs."
      Text 10, 90, 135, 25, "Approved Programs or Closed Programs is built to work when a program has been APPed on the same day."
    EndDialog

    Do
        dialog Dialog1
        cancel_confirmation

        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    If forward_progress = "Run Approved Programs" Then Call run_from_GitHub(script_repository & "/notes/approved-programs.vbs")
    If forward_progress = "Run Closed Programs" Then Call run_from_GitHub(script_repository & "/notes/closed-programs.vbs")

    dail_row = 5
    msg_found = FALSE
    Do
        dail_row = dail_row + 1

        If dail_row = 19 Then
            PF8
            dail_row = 6
        End If
        EMReadScreen full_message, 60, dail_row, 20
        full_message = trim(full_message)
        If left(full_message, 42) = "----------------------------------------->"Then
            Call back_to_SELF
            Call navigate_to_MAXIS_screen("STAT", "SUMM")
            dail_row = 0
            Exit Do
        End If

        If InStr(full_message, "COLA UPDATES IN STAT COMPLETED. REVIEW AND APPROVE") Then msg_found = TRUE
        If InStr(full_message, "REVIEW MEDICARE SAVINGS PROGRAM ELIGIBILITY FOR POSSIBLE") Then msg_found = TRUE
        If InStr(full_message, "REVIEW HEALTH CARE ELIGIBILITY FOR POSSIBLE CHANGES DUE TO") Then msg_found = TRUE
        If InStr(full_message, "PERSON DOES NOT HAVE AN APPROVED HEALTH CARE BUDGET") Then msg_found = TRUE
        If InStr(full_message, "PERSON HAS MAINTENANCE NEEDS ALLOWANCE - REVIEW MEDICAL") Then msg_found = TRUE
        If InStr(full_message, "REVIEW MA-EPD FOR POSSIBLE PREMIUM CHANGES DUE TO") Then msg_found = TRUE
        If InStr(full_message, "HEALTH CARE IS IN REINSTATE OR PENDING STATUS - REVIEW") Then msg_found = TRUE

    Loop until msg_found = TRUE

    EMWriteScreen "S", dail_row, 3                         'Navigates to ELIG/HC - maintaining tie to the DAIL for ease of processin
    transmit

    EMWriteScreen "PNLP", 20, 71        'going to PNLP
    transmit

    Do
        EMGetCursor row, col            'seeing where the cursor is to start (it it is at 20 there are no panels on the particular page)
        Do while row < 20               'If the row is above row 20 then we should write a 'V' for view
            EMSendKey "V"               'Sending 'V' will automatically move the cursor to the next line
            EMGetCursor row, col        'seeing where we are now for the next loop
        Loop
        transmit                        'once we get to line 20, we need to transmit to get to the next page

        EMReadScreen first_panel, 4, 2, 44      'reading the panel name at the top - when we get to ADDR, then we've queued up all the panels to view.
    Loop until first_panel = "ADDR"

    end_msg = "REVIEW OF CASE NEEDED." & vbNewLine & vbNewLine & "No approval was made for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". The PNLP was set up to view every panel so the case can be properly reviewed and approved. Once the approval is done, the script can be rerun to case note the approval." & vbNewLine & error_processing_msg
End If

script_end_procedure_with_error_report(end_msg)
