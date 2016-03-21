'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - DOCUMENTS RECEIVED.vbs"
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
STATS_manualtime = 180          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DIALOGS--------------------------------------------------------------------------------------------------
BeginDialog documents_rec_GEN_dialog, 0, 0, 351, 405, "Documents received"
  EditBox 55, 5, 70, 15, case_number
  EditBox 225, 5, 60, 15, doc_date_stamp
  EditBox 80, 25, 265, 15, docs_rec
  EditBox 30, 70, 315, 15, ADDR
  EditBox 70, 90, 275, 15, SCHL
  EditBox 30, 110, 315, 15, DISA
  EditBox 30, 130, 315, 15, JOBS
  EditBox 30, 150, 315, 15, BUSI
  EditBox 30, 170, 315, 15, UNEA
  EditBox 30, 190, 315, 15, ACCT
  EditBox 55, 210, 290, 15, other_assets
  EditBox 30, 230, 315, 15, SHEL
  EditBox 30, 250, 315, 15, INSA
  EditBox 50, 270, 295, 15, other_verifs
  EditBox 75, 310, 270, 15, notes
  EditBox 75, 330, 270, 15, actions_taken
  EditBox 75, 350, 270, 15, verifs_needed
  EditBox 155, 375, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 240, 375, 50, 15
    CancelButton 295, 375, 50, 15
  Text 5, 95, 65, 10, "SCHL/STIN/STEC:"
  Text 5, 115, 25, 10, "DISA:"
  Text 5, 135, 25, 10, "JOBS:"
  Text 5, 155, 20, 10, "BUSI:"
  Text 5, 175, 25, 10, "UNEA:"
  Text 5, 195, 25, 10, "ACCT:"
  Text 5, 215, 45, 10, "Other assets:"
  Text 5, 235, 25, 10, "SHEL:"
  Text 5, 255, 25, 10, "INSA:"
  Text 5, 270, 45, 10, "Other verif's:"
  Text 95, 380, 60, 10, "Worker signature:"
  Text 5, 75, 25, 10, "ADDR:"
  Text 5, 315, 70, 10, "Notes on your doc's:"
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 335, 50, 10, "Actions taken:"
  Text 140, 45, 205, 10, "Note: What you enter above will become the case note header."
  Text 5, 30, 70, 10, "Documents received: "
  Text 145, 10, 75, 10, "Document date stamp:"
  Text 5, 355, 65, 10, "Verif's still needed:"
  GroupBox 0, 60, 350, 230, "Breakdown of Documents received"
  GroupBox 0, 300, 345, 70, "Additional information"
EndDialog


BeginDialog documents_received_LTC_dialog, 0, 0, 356, 425, "Documents received LTC"
  EditBox 55, 5, 70, 15, case_number
  EditBox 225, 5, 60, 15, doc_date_stamp
  EditBox 80, 25, 265, 15, docs_rec
  EditBox 30, 60, 315, 15, FACI
  EditBox 30, 80, 135, 15, JOBS
  EditBox 210, 80, 135, 15, BUSI_RBIC
  EditBox 30, 100, 315, 15, UNEA
  EditBox 30, 120, 315, 15, ACCT
  EditBox 30, 140, 315, 15, SECU
  EditBox 30, 160, 315, 15, CARS
  EditBox 30, 180, 315, 15, REST
  EditBox 60, 200, 285, 15, OTHR
  EditBox 30, 220, 315, 15, SHEL
  EditBox 30, 240, 315, 15, INSA
  EditBox 75, 260, 270, 15, medical_expenses
  EditBox 50, 280, 295, 15, veterans_info
  EditBox 50, 300, 295, 15, other_verifs
  EditBox 75, 335, 270, 15, notes
  EditBox 75, 355, 270, 15, actions_taken
  EditBox 75, 375, 270, 15, verifs_needed
  EditBox 155, 400, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 240, 400, 50, 15
    CancelButton 295, 400, 50, 15
  Text 5, 145, 20, 10, "SECU:"
  Text 5, 165, 20, 10, "CARS:"
  Text 5, 185, 20, 10, "REST:"
  Text 5, 205, 50, 10, "BURIAL/OTHR:"
  Text 5, 225, 25, 10, "SHEL:"
  Text 5, 245, 25, 10, "INSA:"
  Text 5, 305, 45, 10, "Other verif's:"
  Text 95, 405, 60, 10, "Worker signature:"
  Text 5, 65, 25, 10, "FACI:"
  Text 5, 340, 70, 10, "Notes on your doc's:"
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 360, 50, 10, "Actions taken:"
  Text 145, 40, 205, 10, "Note: What you enter above will become the case note header."
  Text 5, 30, 70, 10, "Documents received: "
  Text 145, 10, 75, 10, "Document date stamp:"
  Text 5, 380, 70, 10, "Verif's still needed:"
  GroupBox 0, 50, 350, 270, "Breakdown of Documents received"
  Text 5, 125, 20, 10, "ACCT:"
  Text 170, 85, 40, 10, "BUSI/RBIC:"
  Text 5, 105, 25, 10, "UNEA:"
  Text 5, 285, 45, 10, "Veteran info:"
  Text 5, 85, 20, 10, "JOBS:"
  Text 5, 265, 65, 10, "Medical expenses:"
  GroupBox 0, 325, 350, 70, "Additional information"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------
'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have a few more options on their dialog.", vbYesNoCancel or VbDefaultButton2) 'defaults to no since that is most commonly chosen option
If LTC_case = vbCancel then stopscript

'Connects to BlueZone
EMConnect ""
'Calls a MAXIS case number
call MAXIS_case_number_finder(case_number)

'Displays the dialog and navigates to case note
'Shows dialog. Requires a case number, checks for an active MAXIS session, and checks that it can add/update a case note before proceeding.
DO
	Do
		Do
			Do
				If LTC_case = vbYes then dialog documents_received_LTC_dialog					'Shows dialog if LTC
				If LTC_case = vbNo then Dialog documents_rec_GEN_dialog							'Shows dialog if not LTC
				cancel_confirmation																'quits if cancel is pressed
				If worker_signature = "" Then MsgBox "You must sign your case note."
			LOOP until worker_signature <> ""
			If actions_taken = "" Then MsgBox "You must case note your actions taken."
		LOOP until actions_taken <> ""
		If case_number = "" then MsgBox "You must have a case number to continue!"		'Yells at you if you don't have a case number
	Loop until case_number <> ""														'Loops until that case number exists
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false														'Loops until that case number exists	

'checking for an active MAXIS session
Call check_for_MAXIS(FALSE)


'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Writes a new line, then writes each additional line if there's data in the dialog's edit box (uses if/then statement to decide).
call start_a_blank_CASE_NOTE
Call write_variable_in_case_note("Docs Rec'd: " & docs_rec)
call write_bullet_and_variable_in_case_note("Document date stamp", doc_date_stamp)
call write_bullet_and_variable_in_case_note("ADDR", ADDR)
call write_bullet_and_variable_in_case_note("FACI", FACI)
call write_bullet_and_variable_in_case_note("SCHL/STIN/STEC", SCHL)
call write_bullet_and_variable_in_case_note("DISA", DISA)
call write_bullet_and_variable_in_case_note("JOBS", JOBS)
call write_bullet_and_variable_in_case_note("BUSI", BUSI)
call write_bullet_and_variable_in_case_note("BUSI/RBIC", BUSI_RBIC)
call write_bullet_and_variable_in_case_note("UNEA", UNEA)
call write_bullet_and_variable_in_case_note("ACCT", ACCT)
call write_bullet_and_variable_in_case_note("SECU", SECU)
call write_bullet_and_variable_in_case_note("CARS", CARS)
call write_bullet_and_variable_in_case_note("REST", REST)
call write_bullet_and_variable_in_case_note("Burial/OTHR", OTHR)
call write_bullet_and_variable_in_case_note("Other assets", other_assets)
call write_bullet_and_variable_in_case_note("SHEL", SHEL)
call write_bullet_and_variable_in_case_note("INSA", INSA)
call write_bullet_and_variable_in_case_note("Medical expenses", medical_expenses)
call write_bullet_and_variable_in_case_note("Veteran's info", veterans_info)
call write_bullet_and_variable_in_case_note("Other verifications", other_verifs)
Call write_variable_in_case_note("---")
call write_bullet_and_variable_in_case_note("Notes on your doc's", notes)
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
call write_bullet_and_variable_in_case_note("Verifications still needed", verifs_needed)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")
