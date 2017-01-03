'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - DOCUMENTS RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
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
call changelog_update("01/03/2017", "Added HSR scanner option for Hennepin County users only.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS--------------------------------------------------------------------------------------------------
BeginDialog documents_rec_dialog, 0, 0, 366, 395, "Documents received"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 225, 5, 60, 15, doc_date_stamp
  If worker_county_code = "x127" then CheckBox 295, 10, 55, 10, "HSR scanner", HSR_scanner_checkbox
  EditBox 80, 25, 265, 15, docs_rec
  EditBox 35, 70, 315, 15, ADDR
  EditBox 75, 90, 275, 15, SCHL
  EditBox 35, 110, 315, 15, DISA
  EditBox 35, 130, 315, 15, JOBS
  EditBox 35, 150, 315, 15, BUSI
  EditBox 35, 170, 315, 15, UNEA
  EditBox 35, 190, 315, 15, ACCT
  EditBox 60, 210, 290, 15, other_assets
  EditBox 35, 230, 315, 15, SHEL
  EditBox 35, 250, 315, 15, INSA
  EditBox 55, 270, 295, 15, other_verifs
  EditBox 80, 310, 270, 15, notes
  EditBox 80, 330, 270, 15, actions_taken
  EditBox 80, 350, 270, 15, verifs_needed
  EditBox 155, 375, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 240, 375, 50, 15
    CancelButton 295, 375, 50, 15
  Text 10, 115, 25, 10, "DISA:"
  Text 10, 135, 25, 10, "JOBS:"
  Text 10, 155, 20, 10, "BUSI:"
  Text 10, 175, 25, 10, "UNEA:"
  Text 10, 195, 25, 10, "ACCT:"
  Text 10, 215, 45, 10, "Other assets:"
  Text 10, 235, 25, 10, "SHEL:"
  Text 10, 255, 20, 10, "INSA:"
  Text 10, 275, 45, 10, "Other verif's:"
  Text 90, 380, 60, 10, "Worker signature:"
  Text 10, 75, 25, 10, "ADDR:"
  Text 10, 315, 70, 10, "Notes on your doc's:"
  Text 30, 10, 45, 10, "Case number:"
  Text 10, 335, 50, 10, "Actions taken:"
  Text 140, 45, 205, 10, "Note: What you enter above will become the case note header."
  Text 10, 30, 70, 10, "Documents received: "
  Text 150, 10, 75, 10, "Document date stamp:"
  Text 10, 355, 65, 10, "Verif's still needed:"
  GroupBox 5, 55, 350, 235, "Breakdown of Documents received"
  GroupBox 5, 295, 350, 75, "Additional information"
  Text 10, 95, 65, 10, "SCHL/STIN/STEC:"
EndDialog

BeginDialog documents_received_LTC, 0, 0, 361, 425, "Documents received LTC"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 230, 5, 60, 15, doc_date_stamp
  If worker_county_code = "x127" then CheckBox 300, 10, 55, 10, "HSR scanner", HSR_scanner_checkbox
  EditBox 80, 25, 270, 15, docs_rec
  EditBox 35, 65, 315, 15, FACI
  EditBox 35, 85, 135, 15, JOBS
  EditBox 215, 85, 135, 15, BUSI_RBIC
  EditBox 35, 105, 315, 15, UNEA
  EditBox 35, 125, 315, 15, ACCT
  EditBox 35, 145, 315, 15, SECU
  EditBox 35, 165, 315, 15, CARS
  EditBox 35, 185, 315, 15, REST
  EditBox 65, 205, 285, 15, OTHR
  EditBox 35, 225, 315, 15, SHEL
  EditBox 35, 245, 315, 15, INSA
  EditBox 80, 265, 270, 15, medical_expenses
  EditBox 55, 285, 295, 15, veterans_info
  EditBox 55, 305, 295, 15, other_verifs
  EditBox 80, 340, 270, 15, notes
  EditBox 80, 360, 270, 15, actions_taken
  EditBox 80, 380, 270, 15, verifs_needed
  EditBox 160, 405, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 245, 405, 50, 15
    CancelButton 300, 405, 50, 15
  Text 10, 170, 20, 10, "CARS:"
  Text 10, 190, 20, 10, "REST:"
  Text 10, 210, 50, 10, "BURIAL/OTHR:"
  Text 10, 230, 25, 10, "SHEL:"
  Text 10, 250, 25, 10, "INSA:"
  Text 10, 310, 45, 10, "Other verif's:"
  Text 100, 410, 60, 10, "Worker signature:"
  Text 10, 70, 25, 10, "FACI:"
  Text 10, 345, 70, 10, "Notes on your doc's:"
  Text 30, 10, 50, 10, "Case number:"
  Text 10, 365, 50, 10, "Actions taken:"
  Text 145, 40, 205, 10, "Note: What you enter above will become the case note header."
  Text 5, 30, 70, 10, "Documents received: "
  Text 155, 10, 75, 10, "Document date stamp:"
  Text 10, 385, 70, 10, "Verif's still needed:"
  GroupBox 5, 50, 350, 275, "Breakdown of Documents received"
  Text 10, 130, 20, 10, "ACCT:"
  Text 175, 90, 40, 10, "BUSI/RBIC:"
  Text 10, 110, 25, 10, "UNEA:"
  Text 10, 290, 45, 10, "Veteran info:"
  Text 10, 90, 20, 10, "JOBS:"
  Text 10, 270, 65, 10, "Medical expenses:"
  GroupBox 5, 330, 350, 70, "Additional information"
  Text 10, 150, 20, 10, "SECU:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------
'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have a few more options on their dialog.", vbYesNoCancel or VbDefaultButton2) 'defaults to no since that is most commonly chosen option
If LTC_case = vbCancel then stopscript

'Connects to BlueZone
EMConnect ""
'Calls a MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)

'Displays the dialog and navigates to case note
'Shows dialog. Requires a case number, checks for an active MAXIS session, and checks that it can add/update a case note before proceeding.
DO
	Do
		Do
			Do
				If LTC_case = vbYes then dialog documents_received_LTC					'Shows dialog if LTC
				If LTC_case = vbNo then Dialog documents_rec_dialog					'Shows dialog if not LTC
				cancel_confirmation																'quits if cancel is pressed
				If worker_signature = "" Then MsgBox "You must sign your case note."
			LOOP until worker_signature <> ""
			If HSR_scanner_checkbox = 0 and actions_taken = "" Then MsgBox "You must case note your actions taken."
		LOOP until actions_taken <> "" or HSR_scanner_checkbox = 1
		If MAXIS_case_number = "" then MsgBox "You must have a case number to continue!"		'Yells at you if you don't have a case number
	Loop until MAXIS_case_number <> ""														'Loops until that case number exists
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false														'Loops until that case number exists	

'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Writes a new line, then writes each additional line if there's data in the dialog's edit box (uses if/then statement to decide).
call start_a_blank_CASE_NOTE
If HSR_scanner_checkbox = 1 then 
    Call write_variable_in_case_note("Docs Rec'd & scanned: " & docs_rec)
else    
    Call write_variable_in_case_note("Docs Rec'd: " & docs_rec)
END IF
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
IF HSR_scanner_checkbox = 1 then Call write_variable_in_case_note("* Documents imaged to ECF.")
call write_bullet_and_variable_in_case_note("Verifications still needed", verifs_needed)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")
