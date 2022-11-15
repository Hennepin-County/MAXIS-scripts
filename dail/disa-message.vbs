'Required for statistical purposes===============================================================================
name_of_script = "DAIL - DISA MESSAGE.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 64          'manual run time in seconds
STATS_denomination = "C"       'C is for Case
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
		FuncLib_URL = FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note. Updated background navigation coding.", "Ilse Ferris")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMConnect ""
EmReadscreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
PF3 'back to DAIL
Call write_value_and_transmit("S", 6, 3)
Call write_value_and_transmit("DISA", 20, 71)

HH_memb = "01"

'HH member dialog to select who's job this is.
dialog1 = ""
BeginDialog Dialog1, 0, 0, 196, 70, "Select DISA HH Member"
  Text 35, 10, 125, 15, "Which HH member is this for? (ex: 01)"
  EditBox 165, 5, 25, 15, HH_memb
  Text 10, 30, 60, 10, "Worker Signature:"
  EditBox 70, 25, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 5, 45, 95, 15, "Temp or Perm Illness - HSR", HSR_manual_button
    OkButton 105, 45, 40, 15
    CancelButton 150, 45, 40, 15
EndDialog

Do
    Do
        err_msg = ""
        Do
            Dialog Dialog1
            Cancel_without_confirmation
            If ButtonPressed = HSR_manual_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Temporary_or_Permanent_Illness.aspx"
        Loop until ButtonPressed = -1
        If (isnumeric(HH_memb) = False and len(HH_memb) <> 2) then err_msg = err_msg & vbcr & "* Enter a 2-digit member number."
    	If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call write_value_and_transmit(HH_memb, 20, 76)
EmReadscreen disa_error, 60, 24, 2
If trim(disa_error) <> "" then script_end_procedure_with_error_report("An error occurred. Review the case and run the script again if applicable. Error messgae: " & disa_error)

'CASH/Housing Supports
If adult_cash_case = True or family_cash_case = True then
    EMReadScreen cash_disa_status, 1, 11, 69
    If  cash_disa_status = "1" or _
        cash_disa_status = "6" or _
        cash_disa_status = "7" then
        cash_request = True
    Else
        cash_request = False
        closing_msg = closing_msg & "This type of Cash DISA status is not yet supported. DISA Code: " & cash_disa_status & ". It could be a SMRT or some other type of verif needed. Process manually at this time." & VbCR & vbCR
    End if
End if
'SNAP
If snap_case = True then
    EMReadScreen snap_disa_status, 1, 12, 69
    IF  snap_disa_status = "1" or _
        snap_disa_status = "5" or _
        snap_disa_status = "6" or _
        snap_disa_status = "7" then
        snap_request = True
    Else
        snap_request = False
        closing_msg = closing_msg & "This type of SNAP DISA status is not yet supported. DISA Code: " & snap_disa_status & ". It could be a SMRT or some other type of verif needed. Process manually at this time." & VbCR & vbCR
    End if
End if
'Health Care Supports
If ma_case = True or msp_case = True then
    EMReadScreen HC_disa_status, 1, 13, 69
    If HC_disa_status = "1" or _
       HC_disa_status = "6" or _
       HC_disa_status = "7" or _
       HC_disa_status = "8" then
        HC_request = True
    Else
        HC_request = False
        closing_msg = closing_msg & "This type of Health Care DISA status is not yet supported. DISA Code: " & hc_disa_status & ". It could be a SMRT or some other type of verif needed. Process manually at this time." & VbCR & vbCR
    End if
End if

'Verif descriptions for case notes
If cash_disa_status = "1" or snap_disa_status = "1" or HC_disa_status = "1" then verif_description = "Dr. Statement/MOF"
If cash_disa_status = "2" or snap_disa_status = "2" or HC_disa_status = "2" then verif_description = "SMRT Certified"
If cash_disa_status = "3" or snap_disa_status = "3" or HC_disa_status = "3" then verif_description = "Certified for RSDI Or SSI"
If cash_disa_status = "4" or snap_disa_status = "4" or HC_disa_status = "4" then verif_description = "Receipt Of HC For Disa/Blind"
If snap_disa_status = "5" then verif_description = "Wrk Judgement"  'SNAP Only: Worker Judgement
If cash_disa_status = "6" or snap_disa_status = "6" or HC_disa_status = "6" then verif_description = "Other Document"
If cash_disa_status = "7" then verif_description = "Professional Stmt of Need"
If snap_disa_status = "7" then verif_description = "Out Of State Ver Pending"
If HC_disa_status = "7" then verif_description = "Case Manager Determination"
If HC_disa_status = "8" then verif_description = "LTC Consult Services" 'HC Only

If cash_request = True or snap_request = True or HC_request = True then
    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    Call create_TIKL("Disability verifs sent 30 days ago. Please review case, case note, and send another request if applicable.", 30, date, False, TIKL_note_text)
    'CASE/NOTE
    Call start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("MEMB " & HH_memb & ": DISABILITY IS ENDING IN 60 DAYS - REVIEW DISA STATUS")
    Call write_variable_in_CASE_NOTE("* DAIL Message receieved and processed for the following programs:")
    'Call write_three_columns_in_CASE_NOTE(col_01_start_point, col_01_variable, col_02_start_point, col_02_variable, col_03_start_point, col_03_variable)
    Call write_three_columns_in_CASE_NOTE(3, "Program", 19, "DISA Description", 50, "DISA Code")
    Call write_variable_in_CASE_NOTE("--------------------------------------------------------")
    If cash_request = True then Call write_three_columns_in_CASE_NOTE(3, "Cash", 19, verif_description, 50, cash_disa_status)
    If snap_request = True then Call write_three_columns_in_CASE_NOTE(3, "SNAP", 19, verif_description, 50, snap_disa_status)
    If HC_request = True then Call write_three_columns_in_CASE_NOTE(3, "Health Care", 19, verif_description, 50, HC_disa_status)
    Call write_variable_in_CASE_NOTE("* Sent verification request and disability forms to resident.")
    Call write_variable_in_CASE_NOTE(TIKL_note_text)
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    closing_msg = closing_msg & "Case note and TIKL created. Send verification request and all applicable disability forms/packet(s) to the resident via ECF."
End if

script_end_procedure_with_error_report(closing_msg)
