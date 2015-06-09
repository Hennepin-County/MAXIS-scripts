'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - VERIFICATIONS NEEDED.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'DIALOGS--------------------------------------------------------------------------------------------------
BeginDialog verifs_needed_dialog, 0, 0, 351, 360, "Verifs needed"
  EditBox 55, 5, 70, 15, case_number
  EditBox 250, 5, 60, 15, verif_due_date
  EditBox 30, 35, 315, 15, ADDR
  EditBox 70, 55, 275, 15, SCHL
  EditBox 30, 75, 315, 15, DISA
  EditBox 30, 95, 315, 15, JOBS
  EditBox 30, 115, 315, 15, BUSI
  EditBox 30, 135, 315, 15, UNEA
  EditBox 30, 155, 315, 15, ACCT
  EditBox 55, 175, 290, 15, other_assets
  EditBox 30, 195, 315, 15, SHEL
  EditBox 30, 215, 315, 15, INSA
  EditBox 50, 235, 295, 15, other_proofs
  CheckBox 5, 260, 240, 10, "Check here if you sent form DHS-2919A (Verification Request Form - A).", verif_A_check
  CheckBox 5, 275, 240, 10, "Check here if you sent form DHS-2919B (Verification Request Form - B).", verif_B_check
  CheckBox 5, 290, 240, 15, "Check here if these are postponed verifications for expedited SNAP.  ", postponed_check
  CheckBox 5, 310, 175, 10, "Sent form to AREP?", sent_arep_checkbox
  CheckBox 5, 325, 95, 10, "Signature page needed?", signature_page_needed_check
  CheckBox 5, 340, 130, 10, "Check here to TIKL out for this case.", TIKL_check
  EditBox 285, 315, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 240, 340, 50, 15
    CancelButton 295, 340, 50, 15
    PushButton 315, 10, 30, 10, "CD+10", CD_plus_10_button
  Text 5, 10, 50, 10, "Case number:"
  Text 150, 10, 100, 10, "Verifs due by (MM/DD/YYYY):"
  Text 5, 25, 300, 10, "If you aren't requesting something, leave that section blank. That way it doesn't case note."
  Text 5, 40, 25, 10, "ADDR:"
  Text 5, 60, 60, 10, "SCHL/STIN/STEC:"
  Text 5, 80, 25, 10, "DISA:"
  Text 5, 100, 25, 10, "JOBS:"
  Text 5, 120, 20, 10, "BUSI:"
  Text 5, 140, 25, 10, "UNEA:"
  Text 5, 160, 25, 10, "ACCT:"
  Text 5, 180, 50, 10, "Other assets:"
  Text 5, 200, 25, 10, "SHEL:"
  Text 5, 220, 25, 10, "INSA:"
  Text 5, 240, 45, 10, "Other proofs:"
  Text 215, 320, 70, 10, "Sign your case note:"
 EndDialog

BeginDialog verifs_needed_LTC_dialog, 0, 0, 351, 415, "Verifs needed (LTC) dialog"
  EditBox 55, 5, 70, 15, case_number
  EditBox 250, 5, 60, 15, verif_due_date
  EditBox 30, 40, 315, 15, FACI
  EditBox 30, 60, 130, 15, JOBS
  EditBox 205, 60, 140, 15, BUSI_RBIC
  EditBox 45, 80, 300, 15, UNEA_01
  EditBox 75, 100, 270, 15, UNEA_other_membs
  EditBox 45, 120, 300, 15, ACCT_01
  EditBox 75, 140, 270, 15, ACCT_other_membs
  EditBox 45, 160, 300, 15, SECU_01
  EditBox 75, 180, 270, 15, SECU_other_membs
  EditBox 30, 200, 315, 15, CARS
  EditBox 30, 220, 315, 15, REST
  EditBox 30, 240, 315, 15, OTHR
  EditBox 30, 260, 315, 15, SHEL
  EditBox 30, 280, 315, 15, INSA
  EditBox 75, 300, 270, 15, medical_expenses
  EditBox 50, 320, 295, 15, other_proofs
  CheckBox 5, 340, 240, 10, "Check here if you sent form DHS-2919A (Verification Request Form - A).", verif_A_check
  CheckBox 5, 355, 240, 10, "Check here if you sent form DHS-2919B (Verification Request Form - B).", verif_B_check
  CheckBox 5, 370, 165, 10, "Sent form to AREP?", sent_arep_checkbox
  CheckBox 5, 385, 95, 10, "Signature page needed?", signature_page_needed_check
  CheckBox 5, 400, 130, 10, "Check here to TIKL out for this case.", TIKL_check
  EditBox 285, 365, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 240, 385, 50, 15
    CancelButton 295, 385, 50, 15
    PushButton 315, 10, 30, 10, "CD+10", CD_plus_10_button
  Text 5, 10, 50, 10, "Case number:"
  Text 150, 10, 100, 10, "Verifs due by (MM/DD/YYYY):"
  Text 5, 25, 300, 10, "If you aren't requesting something, leave that section blank. That way it doesn't case note."
  Text 5, 45, 25, 10, "FACI:"
  Text 5, 65, 25, 10, "JOBS:"
  Text 165, 65, 40, 10, "BUSI/RBIC:"
  Text 5, 85, 35, 10, "UNEA 01:"
  Text 5, 105, 65, 10, "UNEA other membs:"
  Text 5, 125, 35, 10, "ACCT 01:"
  Text 5, 145, 65, 10, "ACCT other membs:"
  Text 5, 165, 35, 10, "SECU 01:"
  Text 5, 185, 70, 10, "SECU other membs:"
  Text 5, 205, 25, 10, "CARS:"
  Text 5, 225, 25, 10, "REST:"
  Text 5, 245, 25, 10, "OTHR:"
  Text 5, 265, 25, 10, "SHEL:"
  Text 5, 285, 25, 10, "INSA:"
  Text 5, 305, 65, 10, "Medical expenses:"
  Text 5, 325, 45, 10, "Other proofs:"
  Text 205, 370, 70, 10, "Sign your case note:"
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------

'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have a few more options on their dialog.", vbYesNoCancel)
If LTC_case = vbCancel then stopscript

'Connects to BlueZone
EMConnect ""

'Calls a MAXIS case number
call MAXIS_case_number_finder(case_number)

'Shows dialog. Requires a case number, checks for an active MAXIS session, and checks that it can add/update a case note before proceeding.
Do
	Do
		Do
			Do
				If LTC_case = vbYes then dialog verifs_needed_LTC_dialog									'Shows dialog if LTC
				If LTC_case = vbNo then dialog verifs_needed_dialog											'Shows dialog if not LTC
				If buttonpressed = cancel then stopscript													'quits if cancel is pressed
				If buttonpressed = CD_plus_10_button then verif_due_date = dateadd("d", 10, date) & ""		'Fills in current date + 10 if you press the button.
			Loop until buttonpressed = OK																	'Loops until you press OK
			If case_number = "" then MsgBox "You must have a case number to continue!"		'Yells at you if you don't have a case number
		Loop until case_number <> ""														'Loops until that case number exists
		transmit							'transmits to see if there's a password screen
		EMReadScreen MAXIS_check, 5, 1, 39	'Checks to see if MAXIS is up. If MAXIS isn't up, it'll yell at you (next line).
		If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "	'If MAXIS is found, it will move on. Otherwise, it loops.
	call navigate_to_screen("case", "note")			'Goes to case/note.
	PF9												'PF9s, which creates a new case note.
	EMReadScreen mode_check, 7, 20, 3				'Reads the case note "mode". The mode needs to have become A or E (add or edit) to continue (next line).
	If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"		'Exits when mode is A or E.

'Writes a new line, then writes each additional line if there's data in the dialog's edit box (uses if/then statement to decide).
IF postponed_check = checked THEN
	call write_new_line_in_case_note(">>>POSTPONED VERIFICATIONS REQUESTED FOR EXPEDITED SNAP<<<")
ELSE
	call write_new_line_in_case_note(">>>Verifications Requested<<<")
END IF
If verif_due_date <> "" then call write_editbox_in_case_note("Verif due date", verif_due_date, 6)
If ADDR <> "" then call write_editbox_in_case_note("ADDR", ADDR, 6)
If FACI <> "" then call write_editbox_in_case_note("FACI", FACI, 6)
If SCHL <> "" then call write_editbox_in_case_note("SCHL/STIN/STEC", SCHL, 6)
If DISA <> "" then call write_editbox_in_case_note("DISA", DISA, 6)
If JOBS <> "" then call write_editbox_in_case_note("JOBS", JOBS, 6)
If BUSI <> "" then call write_editbox_in_case_note("BUSI", BUSI, 6)
If BUSI_RBIC <> "" then call write_editbox_in_case_note("BUSI/RBIC", BUSI_RBIC, 6)
If UNEA <> "" then call write_editbox_in_case_note("UNEA", UNEA, 6)
If UNEA_01 <> "" then call write_editbox_in_case_note("UNEA (MEMB 01)", UNEA_01, 6)
If UNEA_other_membs <> "" then call write_editbox_in_case_note("UNEA (other membs)", UNEA_other_membs, 6)
If ACCT <> "" then call write_editbox_in_case_note("ACCT", ACCT, 6)
If ACCT_01 <> "" then call write_editbox_in_case_note("ACCT (MEMB 01)", ACCT_01, 6)
If ACCT_other_membs <> "" then call write_editbox_in_case_note("ACCT (other membs)", ACCT_other_membs, 6)
If SECU_01 <> "" then call write_editbox_in_case_note("SECU (MEMB 01)", SECU_01, 6)
If SECU_other_membs <> "" then call write_editbox_in_case_note("SECU (other membs)", SECU_other_membs, 6)
If CARS <> "" then call write_editbox_in_case_note("CARS", CARS, 6)
If REST <> "" then call write_editbox_in_case_note("REST", REST, 6)
If OTHR <> "" then call write_editbox_in_case_note("OTHR", OTHR, 6)
If other_assets <> "" then call write_editbox_in_case_note("Other assets", other_assets, 6)
If SHEL <> "" then call write_editbox_in_case_note("SHEL", SHEL, 6)
If INSA <> "" then call write_editbox_in_case_note("INSA", INSA, 6)
If medical_expenses <> "" then call write_editbox_in_case_note("Medical expenses", medical_expenses, 6)
If other_proofs <> "" then call write_editbox_in_case_note("Other proofs", other_proofs, 6)
If signature_page_needed_check = checked then call write_new_line_in_case_note("* Signature page is needed.")
If verif_A_check = checked then call write_new_line_in_case_note("* DHS-2919A (Verification Request Form - A) sent to client.")
If verif_B_check = checked then call write_new_line_in_case_note("* DHS-2919B (Verification Request Form - B) sent to client.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

'If TIKL_check isn't checked this is the end
If TIKL_check = unchecked then script_end_procedure("")

'Navigating to DAIL/WRIT
call navigate_to_screen("dail", "writ")

'If the date in Verif due date is a date, it'll fill that date in on the TIKL.
If IsDate(verif_due_date) = True then call create_MAXIS_friendly_date(verif_due_date, 0, 5, 18)

'Script ends
script_end_procedure("Success! Case note made. You may TIKL when ready. If you filled in a verif due date, it should be autofilled in this TIKL.")


