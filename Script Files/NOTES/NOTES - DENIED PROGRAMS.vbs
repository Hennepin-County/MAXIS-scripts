'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - DENIED PROGRAMS.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 420          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'VARIABLE REQUIRED TO RESIZE DIALOG BASED ON A GLOBAL VARIABLE IN FUNCTIONS FILE
If case_noting_intake_dates = False then dialog_shrink_amt = 105

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

'THE DIALOG----------------------------------------------------------------------------------------------------
'This dialog uses a dialog_shrink_amt variable, along with an if...then which is decided by the global variable case_noting_intake_dates.
BeginDialog denied_dialog, 0, 0, 401, 360 - dialog_shrink_amt, "Denied progs dialog"
  EditBox 65, 5, 55, 15, case_number
  CheckBox 185, 10, 35, 10, "SNAP", SNAP_check
  CheckBox 230, 10, 25, 10, "HC", HC_check
  CheckBox 265, 10, 35, 10, "cash", cash_check
  CheckBox 310, 10, 40, 10, "Emer", emer_check
  EditBox 45, 35, 55, 15, SNAP_denial_date
  EditBox 130, 35, 55, 15, HC_denial_date
  EditBox 225, 35, 55, 15, cash_denial_date
  EditBox 320, 35, 55, 15, emer_denial_date
  EditBox 65, 60, 55, 15, application_date
  EditBox 75, 80, 320, 15, reason_for_denial
  EditBox 140, 100, 255, 15, verifs_needed
  CheckBox 15, 115, 10, 25, "", edit_notice_check
  Text 30, 120, 350, 25, "Check here to have the script add the verifs needed to denial notices. This will list the contents of the above box on the client denial notice. List each of the specific mandatory verifications that were used for the denial."
  EditBox 50, 145, 345, 15, other_notes
  If case_noting_intake_dates = True then
    CheckBox 15, 175, 355, 10, "Check here if proofs were not provided and this case pended the full 30 day period (or 45/60 days for HC).", requested_proofs_not_provided_check
    CheckBox 15, 200, 365, 10, "Denied SNAP for self-declaration of income over 165% FPG (hold for 30 days, with an add'l 30 for proration)", self_declaration_of_income_over_165_FPG
    CheckBox 15, 220, 130, 10, "Client is disabled (60 day HC period)", disabled_client_check
    CheckBox 15, 235, 305, 10, "Check here if there are any programs still open/pending (doesn't become intake again yet)", open_prog_check
    EditBox 105, 250, 235, 15, open_progs
    CheckBox 15, 265, 330, 10, "Check here if there are any HH members still open on HC (won't require a HCAPP to add a member)", HH_membs_on_HC_check
    EditBox 105, 280, 235, 15, HH_membs_on_HC
    GroupBox 0, 160, 390, 140, "Important items that affect the intake date/documentation:"
    Text 40, 185, 125, 10, "Applies a 30 day reinstate period."
    Text 35, 250, 70, 10, "If so, list them here:"
    Text 35, 285, 70, 10, "If so, list them here:"
  Else
    EditBox 155, 160, 200, 15, open_progs
    EditBox 180, 180, 200, 15, HH_membs_on_HC
    Text 5, 160, 150, 10, "If there are any open programs, list them here: "
    Text 5, 180, 175, 10, "If there are any HH membs open on HC, list them here: "
  End if
  CheckBox 5, 310 - dialog_shrink_amt, 65, 10, "Updated MMIS?", updated_MMIS_check
  CheckBox 80, 310 - dialog_shrink_amt, 155, 10, "Check here if you sent a NOMI to this client.", NOMI_check
  CheckBox 245, 310 - dialog_shrink_amt, 95, 10, "WCOM added to notice?", WCOM_check
  CheckBox 30, 325 - dialog_shrink_amt, 125, 10, "Check here to TIKL to send to CLS.", TIKL_check
  EditBox 75, 340 - dialog_shrink_amt, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 265, 340 - dialog_shrink_amt, 50, 15
    CancelButton 320, 340 - dialog_shrink_amt, 50, 15
    PushButton 125, 60 - dialog_shrink_amt, 175, 15, "Autofill previous denied progs script dates/reasons", autofill_previous_info_button
    PushButton 345, 310 - dialog_shrink_amt, 50, 10, "SPEC/WCOM", SPEC_WCOM_button
  Text 5, 10, 50, 10, "Case number:"
  GroupBox 170, 0, 185, 25, "Progs denied:"
  GroupBox 15, 25, 365, 30, "Denial dates:"
  Text 20, 40, 25, 10, "SNAP:"
  Text 115, 40, 15, 10, "HC:"
  Text 200, 40, 20, 10, "cash:"
  Text 295, 40, 20, 10, "Emer:"
  Text 5, 65, 55, 10, "Application date:"
  Text 5, 85, 70, 10, "Reason for denial:"
  Text 5, 105, 130, 10, "Verifs/docs/apps needed (if applicable):"
  Text 5, 145, 45, 10, "Other notes:"
  Text 5, 345 - dialog_shrink_amt, 65, 10, "Worker signature: "
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'SCRIPT CONNECTS, THEN FINDS THE CASE NUMBER
EMConnect ""
Call MAXIS_case_number_finder(case_number)

'Resets the check boxes in case this script was run in succession with the closed progs script. In that script, the variables are named the same and when run one 
'right after another from the Docs Received headquarters it is autofilling these check boxes.------------------------------------------------------------
SNAP_check = 0
cash_check = 0
HC_check = 0
updated_MMIS_check = 0
WCOM_check = 0

'NOW THE DIALOG STARTS. FIRST IT ALLOWS NAVIGATION TO SPEC/WCOM, THEN IT MAKES SURE PROGRAMS ARE SELECTED FOR DENIAL, AND THAT THE REQUIRED DATE FIELDS FOR THOSE PROGRAMS CONTAIN VALID DATES. 
'  THEN IT CHECKS FOR MAXIS STATUS, AND NAVIGATES TO CASE NOTE.
DO
	Do
	    DO
	        Do
			Do
				Dialog denied_dialog
				cancel_confirmation
				If buttonpressed = SPEC_WCOM_button then call navigate_to_MAXIS_screen("spec", "wcom")
				If buttonpressed = autofill_previous_info_button then call autofill_previous_denied_progs_note_info
			Loop until buttonpressed = -1
			If (isdate(SNAP_denial_date) = False and isdate(HC_denial_date) = False and isdate(cash_denial_date) = False and isdate(emer_denial_date) = False) or isdate(application_date) = False then MsgBox "You need to enter a valid date of denial and application date (MM/DD/YYYY)."
			If isdate(SNAP_denial_date) = False then SNAP_denial_date = ""
			If isdate(HC_denial_date) = False then HC_denial_date = ""
			If isdate(cash_denial_date) = False then cash_denial_date = ""
			If isdate(emer_denial_date) = False then emer_denial_date = ""
			If isdate(application_date) = False then application_date = ""
	        Loop until (isdate(SNAP_denial_date) = True or isdate(HC_denial_date) = True or isdate(cash_denial_date) = True or isdate(emer_denial_date) = True) and isdate(application_date) = True
		If ((SNAP_check = 1 and isdate(SNAP_denial_date) = False) or (SNAP_check = 0 and isdate(SNAP_denial_date) = True)) or ((HC_check = 1 and isdate(HC_denial_date) = False) or (HC_check = 0 and isdate(HC_denial_date) = True)) or ((cash_check = 1 and isdate(cash_denial_date) = False) or (cash_check = 0 and isdate(cash_denial_date) = True)) or ((emer_check = 1 and isdate(emer_denial_date) = False) or (emer_check = 0 and isdate(emer_denial_date) = True)) then MsgBox "It looks like you might have checked a program, but not filled in a date. Or vice versa. Look at the programs selected, and make sure there are dates there."
		Loop until ((SNAP_check = 1 and isdate(SNAP_denial_date) = True) or (SNAP_check = 0 and isdate(SNAP_denial_date) = False)) and ((HC_check = 1 and isdate(HC_denial_date) = True) or (HC_check = 0 and isdate(HC_denial_date) = False)) and ((cash_check = 1 and isdate(cash_denial_date) = True) or (cash_check = 0 and isdate(cash_denial_date) = False)) and ((emer_check = 1 and isdate(emer_denial_date) = True) or (emer_check = 0 and isdate(emer_denial_date) = False))
		If SNAP_check = 0 and HC_check = 0 and cash_check = 0 and emer_check = 0 then MsgBox "You need to select a program to deny."
	Loop until SNAP_check = 1 or HC_check = 1 or cash_check = 1 or emer_check = 1
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false


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
  If requested_proofs_not_provided_check = 0 then 
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
  If requested_proofs_not_provided_check = 0 and self_declaration_of_income_over_165_FPG = 0 then 
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
  If cash_denial_date > dateadd("d", application_date, 30) then
    cash_intake_date = cash_denial_date
  Else
    cash_intake_date = dateadd("d", application_date, 30)
  End if
  progs_denied = progs_denied & "cash/"
  cash_last_REIN_date = cash_intake_date & ", after which a new CAF is required."
End if
If emer_check = 1 then
  If emer_denial_date > dateadd("d", application_date, 30) then
    emer_intake_date = emer_denial_date
  Else
    emer_intake_date = dateadd("d", application_date, 30)
  End if
  progs_denied = progs_denied & "emer/"
  emer_last_REIN_date = emer_intake_date & ", after which a new CAF is required."
End if

'deleting last / from progs_withdrawn
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
	notice_edited = false 'Resetting this variable
	call navigate_to_screen("SPEC", "WCOM")
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

'NOW IT CASE NOTES THE DATA.
call start_a_blank_case_note
Call write_variable_in_case_note("----Denied " & progs_denied & "----")
call write_bullet_and_variable_in_case_note("SNAP denial date", SNAP_denial_date)
call write_bullet_and_variable_in_case_note("HC denial date", HC_denial_date)
call write_bullet_and_variable_in_case_note("cash denial date", cash_denial_date)
call write_bullet_and_variable_in_case_note("Emer denial date", emer_denial_date)
call write_bullet_and_variable_in_case_note("Application date", application_date)
call write_bullet_and_variable_in_case_note("Reason for denial", reason_for_denial)
call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
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
END IF

'SUCCESS NOTICE
IF edit_notice_check = checked AND notice_edited = false THEN msgbox "WARNING: You asked the script to edit the eligibilty notices for you, but there were no waiting SNAP/CASH notices showing denied for no proofs.  Please check your denial reasons or edit manually if needed."

script_end_procedure("Success! Case noted and TIKL sent. Please remember to check the generated notice to make sure it reads correctly. If not please add WCOMs to make notice read correctly.")
