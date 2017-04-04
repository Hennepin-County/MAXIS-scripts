'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - DENIED PROGRAMS.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 420          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("04/04/2017", "Added handling for multiple recipient changes to SPEC/WCOM", "David Courtright, St Louis County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'VARIABLE REQUIRED TO RESIZE DIALOG BASED ON A GLOBAL VARIABLE IN FUNCTIONS FILE
If case_noting_intake_dates = False then dialog_shrink_amt = 100

'LOADING SPECIALTY FUNCTIONS----------------------------------------------------------------------------------------------------
function autofill_previous_denied_progs_note_info
  call navigate_to_MAXIS_screen("case", "note")
  MAXIS_row = 1
  MAXIS_col = 1
  EMSearch "---Denied", MAXIS_row, MAXIS_col
  If MAXIS_row = 0 then
    MsgBox "Previous denied progs case note not found."
  Else
    EMWriteScreen "x", MAXIS_row, 3
    transmit
    MAXIS_row = 1                                                              'Scanning for SNAP denial date
    MAXIS_col = 1
    EMSearch "* SNAP denial date: ", MAXIS_row, MAXIS_col
    If MAXIS_row <> 0 then
      SNAP_check = 1
      EMReadScreen SNAP_denial_date, 10, MAXIS_row, 23
    End if
    MAXIS_row = 1                                                              'Scanning for HC denial date
    MAXIS_col = 1
    EMSearch "* HC denial date: ", MAXIS_row, MAXIS_col
    If MAXIS_row <> 0 then
      HC_check = 1
      EMReadScreen HC_denial_date, 10, MAXIS_row, 21
    End if
    MAXIS_row = 1                                                              'Scanning for cash denial date
    MAXIS_col = 1
    EMSearch "* cash denial date: ", MAXIS_row, MAXIS_col
    If MAXIS_row <> 0 then
      cash_check = 1
      EMReadScreen cash_denial_date, 10, MAXIS_row, 23
    End if
    MAXIS_row = 1                                                              'Scanning for emer denial date
    MAXIS_col = 1
    EMSearch "* Emer denial date: ", MAXIS_row, MAXIS_col
    If MAXIS_row <> 0 then
      Emer_check = 1
      EMReadScreen Emer_denial_date, 10, MAXIS_row, 23
    End if
    MAXIS_row = 1                                                              'Scanning for app date
    MAXIS_col = 1
    EMSearch "* Application date: ", MAXIS_row, MAXIS_col
    If MAXIS_row <> 0 then EMReadScreen application_date, 10, MAXIS_row, 23
    MAXIS_row = 1                                                              'Scanning for denial reason
    MAXIS_col = 1
    EMSearch "* Reason for denial: ", MAXIS_row, MAXIS_col
    If MAXIS_row <> 0 then EMReadScreen reason_for_denial, 55, MAXIS_row, 24
    reason_for_denial = trim(reason_for_denial)
    MAXIS_row = 1                                                              'Scanning for verifs needed
    MAXIS_col = 1
    EMSearch "* Verifs needed: ", MAXIS_row, MAXIS_col
    If MAXIS_row <> 0 then EMReadScreen verifs_needed, 59, MAXIS_row, 20
    verifs_needed = trim(verifs_needed)
  End if
End function

Function check_elig_for_verifs
End function

Function check_pnd2_for_denial(coded_denial, SNAP_pnd2_code, cash_pnd2_code, emer_pnd2_code)
	Call navigate_to_MAXIS_screen("REPT", "PND2")
	row = 7
	col = 5
	EMSearch MAXIS_case_number, row, col      'finding correct case to check PND2 codes
	'IF HC_check = checked Then
	'	EMReadScreen HC_pnd2_code, 1, 7, 65
	'	'IF HC_pnd2_code =
	'END IF
	IF SNAP_check = checked Then
		EMReadScreen SNAP_pnd2_code, 1, row, 62
		IF SNAP_pnd2_code = "R" THEN coded_denial = coded_denial & " SNAP withdrawn on PND2."
		IF SNAP_pnd2_code = "I" THEN coded_denial = coded_denial & " SNAP application incomplete, denied on PND2."
		IF SNAP_pnd2_code = "_" THEN
			'If SNAP is selected by the user but the SNAP column is empty on PND2, the script is going to look on the next row for ADDITIONAL APP...
			EMReadScreen additional_maxis_application, 20, row + 1, 16
			additional_maxis_application = trim(additional_maxis_application)
			IF InStr(additional_maxis_application, "ADDITIONAL") <> 0 THEN
				EMReadScreen SNAP_pnd2_code, 1, row + 1, 62
				IF SNAP_pnd2_code = "R" THEN coded_denial = coded_denial & " SNAP withdrawn on PND2."
				IF SNAP_pnd2_code = "I" THEN coded_denial = coded_denial & " SNAP application incomplete, denied on PND2."
			END IF
		END IF
	END IF
	IF cash_check = checked Then
		EMReadScreen cash_pnd2_code, 1, row, 54
		IF cash_pnd2_code = "R" THEN coded_denial = coded_denial & " CASH withdrawn on PND2."
		IF cash_pnd2_code = "I" THEN coded_denial = coded_denial & " CASH application incomplete, denied on PND2."
		IF cash_pnd2_code = "_" THEN
			'If CASH is selected by the user but the CASH column is empty on PND2, the script is going to look on the next row for ADDITIONAL APP...
			EMReadScreen additional_maxis_application, 20, row + 1, 16
			additional_maxis_application = trim(additional_maxis_application)
			IF InStr(additional_maxis_application, "ADDITIONAL") <> 0 THEN
				EMReadScreen cash_pnd2_code, 1, row + 1, 54
				IF cash_pnd2_code = "R" THEN coded_denial = coded_denial & " CASH withdrawn on PND2."
				IF cash_pnd2_code = "I" THEN coded_denial = coded_denial & " CASH application incomplete, denied on PND2."
			END IF
		END IF
	END IF
	IF emer_check = checked Then
		EMReadScreen emer_pnd2_code, 1, row, 68
		IF emer_pnd2_code = "R" THEN coded_denial = coded_denial & " EMER withdrawn on PND2."
		IF emer_pnd2_code = "I" THEN coded_denial = coded_denial & " EMER application incomplete, denied on PND2."
		IF emer_pnd2_code = "_" THEN
			'If EMER is selected by the user but the EMER column is empty on PND2, the script is going to look on the next row for ADDITIONAL APP...
			EMReadScreen additional_maxis_application, 20, row + 1, 16
			additional_maxis_application = trim(additional_maxis_application)
			IF InStr(additional_maxis_application, "ADDITIONAL") <> 0 THEN
				EMReadScreen emer_pnd2_code, 1, row + 1, 68
				IF emer_pnd2_code = "R" THEN coded_denial = coded_denial & " EMER withdrawn on PND2."
				IF emer_pnd2_code = "I" THEN coded_denial = coded_denial & " EMER application incomplete, denied on PND2."
			END IF
		END IF
	END IF
End function

'THE DIALOG----------------------------------------------------------------------------------------------------
'This dialog uses a dialog_shrink_amt variable, along with an if...then which is decided by the global variable case_noting_intake_dates.
BeginDialog denied_dialog, 0, 0, 401, 385 - dialog_shrink_amt, "Denied progs dialog"
  EditBox 65, 5, 55, 15, MAXIS_case_number
  EditBox 185, 5, 55, 15, application_date
  CheckBox 60, 25, 35, 10, "SNAP", SNAP_check
  CheckBox 145, 25, 25, 10, "HC", HC_check
  CheckBox 230, 25, 35, 10, "Cash", cash_check
  CheckBox 315, 25, 40, 10, "Emer", emer_check
  EditBox 60, 40, 55, 15, SNAP_denial_date
  EditBox 145, 40, 55, 15, HC_denial_date
  EditBox 230, 40, 55, 15, cash_denial_date
  EditBox 315, 40, 55, 15, emer_denial_date
  CheckBox 60, 60, 60, 10, "Missing Verifs", missing_verifs_SNAP_checkbox
  CheckBox 145, 60, 60, 10, "Missings Verifs", missing_verifs_HC_checkbox
  CheckBox 230, 60, 60, 10, "Missing Verifs", missing_verifs_CASH_checkbox
  CheckBox 315, 60, 60, 10, "Missing Verifs", missing_verifs_EMER_checkbox
  CheckBox 60, 75, 65, 10, "Denied on Pnd2", denied_pnd2_SNAP_checkbox
  CheckBox 230, 75, 65, 10, "Denied on Pnd2", denied_pnd2_CASH_checkbox
  CheckBox 315, 75, 65, 10, "Denied on Pnd2", denied_pnd2_EMER_checkbox
  CheckBox 60, 90, 75, 10, "Withdrawn on Pnd2", withdraw_pnd2_SNAP_checkbox
  CheckBox 145, 90, 75, 10, "Withdrawn on Pact", withdraw_pact_HC_checkbox
  CheckBox 230, 90, 75, 10, "Withdrawn on Pnd2", withdraw_pnd2_CASH_checkbox
  CheckBox 315, 90, 75, 10, "Withdrawn on Pnd2", withdraw_pnd2_EMER_checkbox
  EditBox 65, 105, 330, 15, reason_for_denial
  EditBox 140, 125, 255, 15, verifs_needed
  Text 30, 145, 350, 25, "Check here to have the script add the verifs needed to denial notices. This will list the contents of the above box on the client denial notice. List each of the specific mandatory verifications that were used for the denial."
  CheckBox 15, 140, 10, 25, "", edit_notice_check
  EditBox 50, 170, 345, 15, other_notes
  If case_noting_intake_dates = True then
    CheckBox 15, 200, 360, 10, "Check here if requested proofs were not provided, interview was completed (if applicable) and this case pended", requested_proofs_not_provided_check
    CheckBox 15, 225, 365, 10, "Denied SNAP for self-declaration of income over 165% FPG (hold for 30 days, with an add'l 30 for proration)", self_declaration_of_income_over_165_FPG
    CheckBox 15, 245, 130, 10, "Client is disabled (60 day HC period)", disabled_client_check
    CheckBox 15, 260, 305, 10, "Check here if there are any programs still open/pending (doesn't become intake again yet)", open_prog_check
    EditBox 105, 275, 235, 15, open_progs
    CheckBox 15, 290, 330, 10, "Check here if there are any HH members still open on HC (won't require a HCAPP to add a member)", HH_membs_on_HC_check
    EditBox 105, 305, 235, 15, HH_membs_on_HC
    GroupBox 5, 190, 390, 140, "Important items that affect the intake date/documentation:"
    Text 40, 210, 300, 10, " the full 30 day period (or 45/60 days for HC). Applies a 30 day reinstate period."
    Text 35, 275, 70, 10, "If so, list them here:"
    Text 35, 310, 70, 10, "If so, list them here:"
  Else
    EditBox 165, 190, 200, 15, open_progs
    EditBox 190, 210, 200, 15, HH_membs_on_HC
    Text 5, 195, 150, 10, "If there are any open programs, list them here: "
    Text 5, 215, 175, 10, "If there are any HH membs open on HC, list them here: "
  End if
  CheckBox 5, 335 - dialog_shrink_amt, 65, 10, "Updated MMIS?", updated_MMIS_check
  CheckBox 80, 335 - dialog_shrink_amt, 155, 10, "Check here if you sent a NOMI to this client.", NOMI_check
  CheckBox 245, 335 - dialog_shrink_amt, 95, 10, "WCOM added to notice?", WCOM_check
  CheckBox 30, 350 - dialog_shrink_amt, 125, 10, "Check here to TIKL to send to CLS.", TIKL_check
  EditBox 75, 365 - dialog_shrink_amt, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 265, 365 - dialog_shrink_amt, 50, 15
    CancelButton 320, 365 - dialog_shrink_amt, 50, 15
    PushButton 250, 5, 145, 15, "Autofill previous denied progs script info", autofill_previous_info_button
    PushButton 345, 335 - dialog_shrink_amt, 50, 10, "SPEC/WCOM", SPEC_WCOM_button
  Text 5, 25, 50, 10, "Denied Progs: "
  Text 5, 10, 50, 10, "Case number:"
  Text 125, 10, 55, 10, "Application date:"
  Text 5, 110, 55, 10, "Other Reasons: "
  Text 5, 130, 130, 10, "Verifs/docs/apps needed (if applicable):"
  Text 5, 175, 45, 10, "Other notes:"
  Text 5, 45, 45, 10, "Denial Date: "
  Text 5, 60, 40, 10, "Reasons:"
  Text 5, 370 - dialog_shrink_amt, 65, 10, "Worker signature: "
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'SCRIPT CONNECTS, THEN FINDS THE CASE NUMBER
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Call check_for_MAXIS(True)

'Resets the check boxes in case this script was run in succession with the closed progs script. In that script, the variables are named the same and when run one
'right after another from the Docs Received headquarters it is autofilling these check boxes.------------------------------------------------------------
SNAP_check = 0
cash_check = 0
HC_check = 0
updated_MMIS_check = 0
WCOM_check = 0


DO
	err_msg = ""
	Dialog denied_dialog
	cancel_confirmation
	If buttonpressed = SPEC_WCOM_button then call navigate_to_MAXIS_screen("spec", "wcom")
	If buttonpressed = autofill_previous_info_button then call autofill_previous_denied_progs_note_info
	If MAXIS_case_number = "" THEN err_msg = err_msg & vbCr & "Please enter a case number."
	If application_date = "" THEN err_msg = err_msg & vbCr & "Please enter an application date."
	If (SNAP_check = checked and SNAP_denial_date = "") or (SNAP_check = unchecked and SNAP_denial_date <> "") THEN err_msg = err_msg & vbCr & "You have checked SNAP but not added a denial date, or vice versa."
	If (HC_check = checked and HC_denial_date = "") or (HC_check = unchecked and HC_denial_date <> "") THEN err_msg = err_msg & vbCr & "You have checked HC but not added a denial date, or vice versa."
	If (cash_check = checked and cash_denial_date = "") or (cash_check = unchecked and cash_denial_date <> "") THEN err_msg = err_msg & vbCr & "You have checked cash but not added a denial date, or vice versa."
	If (emer_check = checked and emer_denial_date = "") or (emer_check = unchecked and emer_denial_date <> "") THEN err_msg = err_msg & vbCr & "You have checked emer but not added a denial date, or vice versa."
	If isdate(SNAP_denial_date) = FALSE and SNAP_check = checked THEN err_msg = err_msg & vbCr & "The date you entered for SNAP denial is not a valid date."
	If isdate(HC_denial_date) = FALSE and HC_check = checked THEN err_msg = err_msg & vbCr & "The date you entered for HC denial is not a valid date."
	If isdate(cash_denial_date) = FALSE and cash_check = checked THEN err_msg = err_msg & vbCr & "The date you entered for CASH denial is not a valid date."
	If isdate(emer_denial_date) = FALSE and emer_check = checked THEN err_msg = err_msg & vbCr & "The date you entered for emer denial is not a valid date."
	If SNAP_check = checked and missing_verifs_SNAP_checkbox = unchecked and denied_pnd2_SNAP_checkbox = unchecked and withdraw_pnd2_SNAP_checkbox = unchecked and reason_for_denial = "" THEN err_msg = err_msg & vbCr & "You selected the SNAP checkbox but did not check a reason or write a reason in other reasons."
	If HC_check = checked and missing_verifs_HC_checkbox = unchecked and withdraw_pact_HC_checkbox = unchecked and reason_for_denial = "" THEN err_msg = err_msg & vbCr & "You selected the HC checkbox but did not check a reason or write a reason in other reasons."
	If cash_check = checked and missing_verifs_cash_checkbox = unchecked and denied_pnd2_cash_checkbox = unchecked and withdraw_pnd2_cash_checkbox = unchecked and reason_for_denial = "" THEN err_msg = err_msg & vbCr & "You selected the CASH checkbox but did not check a reason or write a reason in other reasons."
	If emer_check = checked and missing_verifs_emer_checkbox = unchecked and denied_pnd2_emer_checkbox = unchecked and withdraw_pnd2_emer_checkbox = unchecked and reason_for_denial = "" THEN err_msg = err_msg & vbCr & "You selected the EMER checkbox but did not check a reason or write a reason in other reasons."
	If missing_verifs_SNAP_checkbox = checked and verifs_needed = "" THEN err_msg = err_msg & vbCr & "You checked SNAP missings verifs as a reason but didn't enter verifs needed."
	If missing_verifs_HC_checkbox = checked and verifs_needed = "" THEN err_msg = err_msg & vbCr & "You checked HC missings verifs as a reason but didn't enter verifs needed, or vice versa."
	If missing_verifs_CASH_checkbox = checked and verifs_needed = "" THEN err_msg = err_msg & vbCr & "You checked CASH missings verifs as a reason but didn't enter verifs needed, or vice versa."
	If missing_verifs_EMER_checkbox = checked and verifs_needed = "" THEN err_msg = err_msg & vbCr & "You checked EMER missings verifs as a reason but didn't enter verifs needed, or vice versa."
	If (open_prog_check = checked and open_progs = "") and (open_prog_check = unchecked and open_progs <> "") THEN err_msg = err_msg & vbCr & "You checked that there are open/pending progs but didn't list them, or vice versa."
	If (HH_membs_on_HC_check = checked and HH_membs_on_HC = "") and (HH_membs_on_HC_check = unchecked and HH_membs_on_HC <> "") THEN err_msg = err_msg & vbCr & "You checked that there are members open on HC but didn't list them, or vice versa."
	If worker_signature = "" THEN err_msg = err_msg & vbCr & "Please enter a worker signature."
	coded_denial = "" 			'Reseting this value to make sure we are not duplicating the case note.
	call check_pnd2_for_denial(coded_denial, SNAP_pnd2_code, cash_pnd2_code, emer_pnd2_code)
	If SNAP_pnd2_code = "R" and withdraw_pnd2_SNAP_checkbox = unchecked THEN err_msg = err_msg & vbCr & "Your PND2 has SNAP coded as R. Please select withdraw checkbox."
	If SNAP_pnd2_code = "I" and denied_pnd2_SNAP_checkbox = unchecked THEN err_msg = err_msg & vbCr & "Your PND2 has SNAP coded as I. Please select deny from PND2 checkbox."
	If SNAP_pnd2_code <> "R" and withdraw_pnd2_SNAP_checkbox = checked THEN err_msg = err_msg & vbCr & "Your checked the box indicating SNAP was withdraw but your PND2 is not coded as such Please correct your PND2."
	If SNAP_pnd2_code <> "I" and denied_pnd2_SNAP_checkbox = checked THEN err_msg = err_msg & vbCr & "Your checked the box indicating SNAP was incomplete and denied but your PND2 is not coded as such Please correct your PND2."
	If cash_pnd2_code = "R" and withdraw_pnd2_cash_checkbox = unchecked THEN err_msg = err_msg & vbCr & "Your PND2 has CASH coded as R. Please select withdraw checkbox."
	If cash_pnd2_code = "I" and denied_pnd2_cash_checkbox = unchecked THEN err_msg = err_msg & vbCr & "Your PND2 has CASH coded as I. Please select deny from PND2 checkbox."
	If cash_pnd2_code <> "R" and withdraw_pnd2_cash_checkbox = checked THEN err_msg = err_msg & vbCr & "Your checked the box indicating CASH was withdraw but your PND2 is not coded as such Please correct your PND2."
	If cash_pnd2_code <> "I" and denied_pnd2_cash_checkbox = checked THEN err_msg = err_msg & vbCr & "Your checked the box indicating CASH was incomplete and denied but your PND2 is not coded as such Please correct your PND2."
	If emer_pnd2_code = "R" and withdraw_pnd2_emer_checkbox = unchecked THEN err_msg = err_msg & vbCr & "Your PND2 has EMER coded as R. Please select withdraw checkbox."
	If emer_pnd2_code = "I" and denied_pnd2_emer_checkbox = unchecked THEN err_msg = err_msg & vbCr & "Your PND2 has EMER coded as I. Please select deny from PND2 checkbox."
	If emer_pnd2_code <> "R" and withdraw_pnd2_emer_checkbox = checked THEN err_msg = err_msg & vbCr & "Your checked the box indicating EMER was withdraw but your PND2 is not coded as such Please correct your PND2."
	If emer_pnd2_code <> "I" and denied_pnd2_emer_checkbox = checked THEN err_msg = err_msg & vbCr & "Your checked the box indicating EMER was incomplete and denied but your PND2 is not coded as such Please correct your PND2."
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""


'checking for an active MAXIS session
Call check_for_MAXIS(False)

'IT CONVERTS THE DATE FIELDS TO ACTUAL DATES FOR CALCULATION PURPOSES.
If isdate(SNAP_denial_date) = true then SNAP_denial_date = cdate(SNAP_denial_date)
If isdate(HC_denial_date) = true then HC_denial_date = cdate(HC_denial_date)
If isdate(cash_denial_date) = true then cash_denial_date = cdate(cash_denial_date)
If isdate(emer_denial_date) = true then emer_denial_date = cdate(emer_denial_date)
application_date = cdate(application_date)

'THE DISABLED STATUS AFFECTS THE REAPPLICATION DATE. DISABLED CLIENTS GET 60 DAYS FOR HC, OTHERS GET 45.
If disabled_client_check = 1 then
  HC_intake_date_diff = 60
Else
  HC_intake_date_diff = 45
End if

'NOW THE SCRIPT CALCULATES WHAT THE INTAKE DATES WOULD BE FOR EACH PROGRAM.
If HC_check = 1 then
  If requested_proofs_not_provided_check = 0 or withdraw_pact_HC_checkbox = checked then
    HC_intake_date = dateadd("d", HC_denial_date, 10)
  Else
    If dateadd("d", HC_denial_date, 10) > dateadd("d", application_date, HC_intake_date_diff) then
      HC_intake_date = dateadd("d", HC_denial_date, 10)
    Else
      HC_intake_date = dateadd("d", application_date, HC_intake_date_diff)
    End if
  End if
  progs_denied = progs_denied & "HC/"
  If HH_membs_on_HC_check = 1 then
    HC_last_REIN_date = HC_intake_date & ", new HCAPP not required if other membs are open on HC."
  Else
    HC_last_REIN_date = HC_intake_date & ", after which a new HCAPP is required."
  End if
End if
If SNAP_check = 1 then
  If withdraw_pnd2_SNAP_checkbox = checked Then
	SNAP_intake_date = dateadd("d", SNAP_denial_date, 10)
  ElseIf requested_proofs_not_provided_check = 0 and self_declaration_of_income_over_165_FPG = 0 then
    SNAP_intake_date = SNAP_denial_date
  ElseIf dateadd("d", SNAP_denial_date, 10) > dateadd("d", application_date, 60) then
    SNAP_intake_date = dateadd("d", SNAP_denial_date, 10)
  Else
    SNAP_intake_date = dateadd("d", application_date, 60)
  End if
  progs_denied = progs_denied & "SNAP/"
  SNAP_last_REIN_date = SNAP_intake_date & ", after which a new CAF is required."
End if
If cash_check = 1 then
  If withdraw_pnd2_CASH_checkbox = checked Then
	cash_intake_date = dateadd("d", cash_denial_date, 10)
  ElseIf cash_denial_date > dateadd("d", application_date, 30) then
    cash_intake_date = cash_denial_date
  Else
    cash_intake_date = dateadd("d", application_date, 30)
  End if
  progs_denied = progs_denied & "cash/"
  cash_last_REIN_date = cash_intake_date & ", after which a new CAF is required."
End if
If emer_check = 1 then
  If withdraw_pnd2_EMER_checkbox = checked Then
	emer_intake_date = dateadd("d", emer_denial_date, 10)
  ElseIf emer_denial_date > dateadd("d", application_date, 30) then
    emer_intake_date = emer_denial_date
  Else
    emer_intake_date = dateadd("d", application_date, 30)
  End if
  progs_denied = progs_denied & "emer/"
  emer_last_REIN_date = emer_intake_date & ", after which a new CAF is required."
End if

'deleting last / from progs_denied
progs_denied = left(progs_denied, len(progs_denied) - 1)

'IT HAS TO FIGURE OUT WHICH DATE IS THE LATEST DATE, AS THAT WOULD BE THE DATE THE CLIENT HAS TO BE REASSIGNED TO INTAKE.
If HC_intake_date > SNAP_intake_date then
  If HC_check = 1 then
    intake_date = dateadd("d", HC_intake_date, 1)
  ElseIf SNAP_check = 1 then
    intake_date = dateadd("d", SNAP_intake_date, 1)
  Elseif cash_check = 1 then
    intake_date = dateadd("d", cash_intake_date, 1)
  Elseif emer_check = 1 then
    intake_date = dateadd("d", emer_intake_date, 1)
  End if
Else
  If SNAP_check = 1 then
    intake_date = dateadd("d", SNAP_intake_date, 1)
  ElseIf HC_check = 1 then
    intake_date = dateadd("d", HC_intake_date, 1)
  Elseif cash_check = 1 then
    intake_date = dateadd("d", cash_intake_date, 1)
  Elseif emer_check = 1 then
    intake_date = dateadd("d", emer_intake_date, 1)
  End if
End if
If cash_intake_date > intake_date and cash_check = 1 then intake_date = cash_intake_date
If emer_intake_date > intake_date and emer_check = 1 then intake_date = emer_intake_date

'This section edits the notices if requested.

IF edit_notice_check = 1 THEN
	'This section will check for whether forms go to AREP and SWKR
	call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	EMReadscreen forms_to_arep, 1, 10, 45
	call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	EMReadscreen forms_to_swkr, 1, 15, 63

	notice_edited = false 'Resetting this variable
	call navigate_to_MAXIS_screen("SPEC", "WCOM")
	notice_month = DatePart("m", application_date) 'Entering the benefit month to find notices
	IF len(notice_month) = 1 THEN notice_month = "0" & notice_month
	EMWritescreen notice_month, 3, 46
	EMWriteScreen right(DatePart("yyyy", application_date), 2), 3, 51
	transmit
	row = 6 'setting the variables for EMSearch
	col = 1
	DO 'This loop looks for any waiting notices to edit, and edits them
	EMSearch "Waiting", row, col
	IF row > 6 THEN 'Found a waiting notice, Checking for a match to our denied programs
		EMReadScreen prg_typ, 2, row, 26
		'putting the denied progs into a formatted list
		progs_denied_list = Replace(progs_denied, "SNAP", "FS")
		IF instr(ucase(progs_denied_list), prg_typ) > 0 THEN
			EMWriteScreen "X", row, 13
			Transmit
			'Making sure the notice is actually a denial for verifications
			document_end = "" 'resetting the variable
			DO
				notice_row = 1
				notice_col = 1
				EMSearch "proof", notice_row, notice_col
				IF notice_row = 0 THEN 'It didn't spot the word proofs, checking the next page
					PF8
					EMReadScreen document_end, 3, 24, 13
					IF document_end <> "   " then EXIT DO
				END IF
			LOOP UNTIL notice_row > 1 OR document_end <> "   "
			IF notice_row > 1 THEN	'This means the word "proofs" is contained in the notice, and it should be edited
				PF9
				'The script is now on the recipient selection screen.  Mark all recipients that need NOTICES
				row = 4                             'Defining row and col for the search feature.
				col = 1
				EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
				IF row > 4 THEN  arep_row = row  'locating ALTREP location if it exists'
				row = 4                             'reset row and col for the next search
				col = 1
				EMSearch "SOCWKR", row, col
				IF row > 4 THEN  swkr_row = row     'Logs the row it found the SOCWKR string as swkr_row
				EMWriteScreen "x", 5, 10                                        'We always send notice to client
				IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
				IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
				transmit                                                        'Transmits to start the memo writing process'
				'Writing the verifs needed into the notice
				call write_variable_in_spec_memo("The following verifications were not provided: ")
				call write_variable_in_spec_memo("")
				call write_variable_in_spec_memo(verifs_needed)
				notice_edited = true 'Setting this true lets us know that we successfully edited the notice
				pf4
				pf3
				WCOM_check = 1 'This makes sure to case note that the notice was edited, even if user doesn't check the box.
			ELSE
				pf3
			END IF
		END IF
		row = row + 1 'THis makes the next search start at current line +1
	END IF
	IF row = 0 THEN
		EMReadScreen second_page_check, 1, 18, 77 'looking for a 2nd page of notices
		IF second_page_check = "+" THEN
			PF8
			row = 6 'resetting search variables
			col = 1
		ELSE
			PF8 'this changes to the next benefit month to look for more notices
			row = 6 'resetting search variables
			col = 1
			EMReadScreen last_month_check, 3, 24, 2
			IF last_month_check = "NOT" THEN EXIT DO 'the last month has been reached, exit the loop.
		END IF
	END IF
	LOOP UNTIL row = 18 or last_month_check = "NOT"
END IF

If self_declaration_of_income_over_165_FPG = 1 THEN
	call navigate_to_MAXIS_screen("STAT", "PROG")
	EMReadScreen int_date, 8, 10, 55
          	int_date = replace(int_date, " ", "/")
	call navigate_to_MAXIS_screen("ELIG", "FS")
	transmit
	EMWriteScreen "x", 15, 4
	transmit
	EMReadScreen reported_income, 10, 9, 30
	reported_income = trim(reported_income)
	EMReadScreen max_gross_income, 10, 15, 67
	max_gross_income = trim(max_gross_income)
End if

'NOW IT CASE NOTES THE DATA.
call start_a_blank_case_note
Call write_variable_in_case_note("----Denied " & progs_denied & "----")
call write_bullet_and_variable_in_case_note("SNAP denial date", SNAP_denial_date)
call write_bullet_and_variable_in_case_note("HC denial date", HC_denial_date)
call write_bullet_and_variable_in_case_note("cash denial date", cash_denial_date)
call write_bullet_and_variable_in_case_note("Emer denial date", emer_denial_date)
call write_bullet_and_variable_in_case_note("Application date", application_date)
call write_bullet_and_variable_in_case_note("Reason for denial", reason_for_denial)
call write_bullet_and_variable_in_case_note("Coding for denial", coded_denial)
call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
	'adding case note portion to cover Self Declaration of Over Income Policy
	If self_declaration_of_income_over_165_FPG = 1 THEN
		call write_variable_in_case_note("---")
		call write_variable_in_case_note("   ***Self Declaration of Over Income Policy for SNAP***")
		call write_variable_in_case_note("* Date of Interview: " & int_date)
		call write_variable_in_case_note("* Client's Stated Total Income: $" & reported_income)
		call write_variable_in_case_note("* Max Gross Income 165% of FPG: $" & max_gross_income)
		call write_variable_in_case_note("* Denial Reason: Client stated their income is greater than 165% of FPG")
	End If
If updated_MMIS_check = 1 then call write_variable_in_case_note("* Updated MMIS.")
If disabled_client_check = 1 then call write_variable_in_case_note("* Client is disabled.")
If WCOM_check = 1 then call write_variable_in_case_note("* Added WCOM to notice.")
If NOMI_check = 1 then call write_variable_in_case_note("* Sent NOMI to client.")
If case_noting_intake_dates = True then
	call write_variable_in_case_note("---")
	If HC_check = 1 then call write_bullet_and_variable_in_case_note("Last HC REIN date", HC_last_REIN_date)
	If SNAP_check = 1 then call write_bullet_and_variable_in_case_note("Last SNAP REIN date", SNAP_last_REIN_date)
	If cash_check = 1 then call write_bullet_and_variable_in_case_note("Last cash REIN date", cash_last_REIN_date)
	If emer_check = 1 then call write_bullet_and_variable_in_case_note("Last emer REIN date", emer_last_REIN_date)
	If open_prog_check = 1 or HH_membs_on_HC_check = 1 then
		If open_progs <> "" then call write_bullet_and_variable_in_case_note("Open programs", open_progs)
		If HH_membs_on_HC <> "" then call write_bullet_and_variable_in_case_note("HH members remaining on HC", HH_membs_on_HC)
	Else
		call write_variable_in_case_note("* All programs denied. Case becomes intake again on " & intake_date & ".")
	End if
Else
	If open_progs <> "" then call write_bullet_and_variable_in_case_note("Open programs", open_progs)
	If HH_membs_on_HC <> "" then call write_bullet_and_variable_in_case_note("HH members remaining on HC", HH_membs_on_HC)
End if
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

'defining ending message without te TIKLS
ending_message = "Success! Please remember to check the generated notice to make sure it reads correctly. If not please add WCOMs to make notice read correctly."

'TIKL PORTION -------------------------------------------------------------------------------------------------------------
If TIKL_check = 1 THEN
	'IF PROGRAMS ARE STILL OPEN, BUT THE "TIKL TO SEND TO CLS" PARAMETER WAS SET, THE SCRIPT NEEDS TO STOP, AS THE CASE CAN'T GO TO CLS.
	If open_prog_check = 1 then
		MsgBox "Because you checked the open programs box, the script will not TIKL to send to CLS."
		IF edit_notice_check = checked AND notice_edited = false THEN msgbox "WARNING: You asked the script to edit the eligibilty notices for you, but there were no waiting SNAP/CASH notices showing denied for no proofs.  Please check your denial reasons or edit manually if needed."
		script_end_procedure("")
	End if

	'IT NAVIGATES TO DAIL/WRIT.
	call navigate_to_MAXIS_screen("dail", "writ")

	'DETERMINES THE CORRECT FORMATTING FOR THE DATE CLIENT BECOMES AN INTAKE.
	TIKL_day = datepart("d", intake_date)
	If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day
	TIKL_month = datepart("m", intake_date)
	If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month
	TIKL_year = right(intake_date, 2)

	'WRITES TIKL TO SEND TO CLS
	EMWriteScreen TIKL_month, 5, 18
	EMWriteScreen TIKL_day, 5, 21
	EMWriteScreen TIKL_year, 5, 24
	EMSetCursor 9, 3
	EMSendKey "Case was denied " & denial_date & ". If required proofs have not been received, send to CLS per policy. TIKL auto-generated via script."
	'SAVES THE TIKL
	PF3
	ending_message = "Success! Case noted and TIKL sent. Please remember to check the generated notice to make sure it reads correctly. If not please add WCOMs to make notice read correctly."
END IF

'SUCCESS NOTICE
IF edit_notice_check = checked AND notice_edited = false THEN msgbox "WARNING: You asked the script to edit the eligibilty notices for you, but there were no waiting SNAP/CASH notices showing denied for no proofs.  Please check your denial reasons or edit manually if needed."

script_end_procedure(ending_message) 'the ending message is determined earlier based on if a TIKL was created or not. This makes sure to include the "Success!" which is how we gather stats for scripts that ran to the end.
