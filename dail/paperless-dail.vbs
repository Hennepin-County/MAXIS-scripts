'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - paperless IR.vbs"
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
call changelog_update("05/01/2019", "Removed the option to delete the dail as there may be issues with it. Will review an return it once rewrite/testing completed.", "Casey Love, Hennepin County")
call changelog_update("09/12/2018", "Bug fixed that was preventing LTC scripts from erroring out.", "Ilse Ferris, Hennepin County")
call changelog_update("05/19/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG
BeginDialog delete_message_dialog, 0, 0, 126, 45, "Double-Check the Computer's Work..."
  ButtonGroup ButtonPressed
    PushButton 10, 25, 50, 15, "YES", delete_button
    PushButton 60, 25, 50, 15, "NO", do_not_delete
  Text 30, 10, 65, 10, "Delete the DAIL??"
EndDialog

'CONNECTS TO DEFAULT SCREEN
EMConnect ""

''CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
'EMReadscreen dail_check, 4, 2, 48
'If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")
'
''TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
'EMSendKey "t"
'transmit
'
''The following reads the message in full for the end part (which tells the worker which message was selected)
'EMReadScreen full_message, 58, 6, 20
'
''FS Eligibility Ending for ABAWD
'EMReadScreen Paperless_tikl_check, 49, 6, 20
'IF Paperless_tikl_check <> "%^% SENT THROUGH BACKGROUND USING BULK SCRIPT %^%" THEN script_end_procedure("This is not the correct kind of DAIL for this script. Run the main DAIL Scrubber for the full supported scripts.")
'
'=========================================================================================
'Everything above this line is a part of the DAIL Scrubber Script if this becomes state supported. Just change the last line to the correct call from github

If run_from_DAIL = TRUE Then
    EMReadScreen Paperless_tikl_check, 49, 6, 20
    'DATE CALCULATIONS'
    next_month = DateAdd("m", 1, date)
    approval_month = DatePart("m", next_month)
    approval_year = DatePart("yyyy", next_month)

    approval_month = right("00" & approval_month, 2)
    approval_year = right(approval_year, 2)

    EMWriteScreen "E", 6, 3                         'Navigates to ELIG/HC - maintaining tie to the DAIL for ease of processin
    transmit
    EMWriteScreen "HC", 20, 71
    transmit
Else
    approval_month = MAXIS_footer_month
    approval_year = MAXIS_footer_year

    Call Navigate_to_MAXIS_screen("ELIG", "HC  ")
End If

EMReadScreen hc_elig_check, 4, 3, 51
If hc_elig_check <> "HHMM" Then script_end_procedure("No HC ELIG results exist, resolve edits and approve new version and run the script again.")
EMWriteScreen approval_month, 20, 56            'Goes to the next month and checks that elig results exist
EMWriteScreen approval_year,  20, 59
transmit
If hc_elig_check <> "HHMM" Then script_end_procedure("No HC ELIG results exist, resolve edits and approve new version and run the script again.")


row = 8                                          'Reads each line of Elig HC to find all the approved programs in a case
Do
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
                EMReadScreen waiver_check, 1, 14, 21        'Checking to see if case may be LTC or Waiver'
                EMReadScreen method_check, 1, 13, 21
                If method_check = "L" or method_check = "S" Then LTC_case = TRUE
                If method_check = "B" AND waiver_check <> "_" Then LTC_case = TRUE
                If method_check = "X" AND waiver_check <> "_" Then LTC_case = TRUE
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
                transmit
            End If
        Loop until current_version = "01" OR prog_status = "APP"
        'Adds everything to a varriable so an array can be created
        Elig_Info_array = Elig_Info_array & "~Memb " & clt_ref_num & " is approved as " & trim(elig_result) & " for " & trim(clt_hc_prog) & " : " & Majr_prog & "-" & elig_type
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

If run_LTC_Approval = vbYes Then                'Defining more variables for the LTC Script and then running it.
    MAXIS_footer_month = approval_month         'The rest of this script will not run if LTC script is selected
    MAXIS_footer_year = approval_year
    special_header_droplist = "Paperless IR"

    call run_from_GitHub( script_repository & "notes/ltc-ma-approval.vbs")
End If

If run_from_DAIL = TRUE Then
    PF3             'Back to DAIL
End If

If error_processing_msg <> "" Then script_end_procedure(error_processing_msg)

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
BeginDialog approval_dialog, 0, 0, 286, 115 + (15 * UBound(Elig_Info_array)), "Approval dialog"
  For each elig_approval in Elig_Info_array
    CheckBox 10, 40 + (15 * array_counter), 265, 10, elig_approval, elig_checkbox_array(array_counter)
	array_counter = array_counter + 1
  Next
  CheckBox 5, 60 + (15 * UBound(Elig_Info_array)), 220, 10, "Check here if you have reviewed/updated MMIS and it is correct", mmis_checkbox
  EditBox 65, 75 + (15 * UBound(Elig_Info_array)), 215, 15, other_notes
  EditBox 65, 95 + (15 * UBound(Elig_Info_array)), 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 95 + (15 * UBound(Elig_Info_array)), 50, 15
    CancelButton 230, 95 + (15 * UBound(Elig_Info_array)), 50, 15
  Text 5, 5, 50, 10, "Case Number:"
  Text 60, 5, 30, 10, MAXIS_case_number
  Text 120, 5, 55, 10, "Approval Month:"
  Text 180, 5, 25, 10, approval_month & "/" & approval_year
  Text 5, 25, 275, 10, "Script has identified the following HC Approvals. They will be case noted if checked."
  Text 5, 80 + (15 * UBound(Elig_Info_array)), 55, 10, "Other Notes:"
  Text 5, 100 + (15 * UBound(Elig_Info_array)), 60, 10, "Worker signature:"
EndDialog

Do
    err_msg = ""
    Dialog approval_dialog
    cancel_confirmation
    If worker_signature = "" then err_msg = err_msg & vbNewLine & "Please sign your case note"
    if err_msg <> "" Then MsgBox err_msg
Loop until err_msg = ""

If run_from_DAIL = TRUE Then
    EMWriteScreen "N", 6, 3         'Goes to Case Note - maintains tie with DAIL
    transmit

    'Starts a blank case note
    PF9
    EMReadScreen case_note_mode_check, 7, 20, 3
    If case_note_mode_check <> "Mode: A" then script_end_procedure("You are not in a case note on edit mode. You might be in inquiry. Try the script again in production.")
Else
    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    PF9 'edit mode
End If

'Adding information to case note
Call write_variable_in_CASE_NOTE ("---Approved HC - IR Waived---")
Call write_variable_in_CASE_NOTE ("* Processed HC for 6 Mo Renewal for " & approval_month & "/" & approval_year)
For array_item = 0 to UBound(Elig_Info_array)
    If elig_checkbox_array(array_item) = checked Then Call write_variable_in_CASE_NOTE ("* " & Elig_Info_array(array_item))
Next
Call write_bullet_and_variable_in_CASE_NOTE ("Notes", other_notes)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

MAXIS_case_number = trim(MAXIS_case_number)
' If run_from_DAIL = TRUE Then
'     DIALOG delete_message_dialog
'     IF ButtonPressed = delete_button THEN
'     	PF3
'     	PF3
'     	DO
'     		dail_read_row = 6
'     		DO
'     			EMReadScreen double_check, 49, dail_read_row, 20
'     			IF double_check = Paperless_tikl_check THEN
'                     EMWriteScreen "T", dail_read_row, 3
'                     EMReadScreen dail_case_number, 8, 5, 73
'                     dail_case_number = trim(dail_case_number)
'                     If dail_case_number = MAXIS_case_number Then EMWriteScreen "D", 6, 3
'     				transmit
'     				EXIT DO
'     			ELSE
'     				dail_read_row = dail_read_row + 1
'     			END IF
'     			IF dail_read_row = 19 THEN PF8
'     		LOOP UNTIL dail_read_row = 19
'     		EMReadScreen others_dail, 13, 24, 2
'     		If others_dail = "** WARNING **" Then transmit
'     	LOOP UNTIL double_check = Paperless_tikl_check
'     END IF
' End If

script_end_procedure("")
