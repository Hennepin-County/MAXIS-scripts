'STATS GATHERING=============================================================================================================
name_of_script = "NOTICES - VERIFICATIONS STILL NEEDED.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1     	          'sets the stats counter at one
STATS_manualtime = 285            'manual run time in seconds
STATS_denomination = "C"	        'C is for each case
'END OF stats block==========================================================================================================

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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS FOR THE SCRIPT======================================================================================================
''Still Needed Dialog
BeginDialog Verifications_Still_Needed_Dialog, 0, 0, 341, 320, "Verifications Still Needed Dialog"
  EditBox 60, 5, 120, 15, MAXIS_case_number
  EditBox 30, 40, 150, 15, address_verification
  EditBox 70, 60, 110, 15, schl_stin_stec_verification
  EditBox 30, 80, 150, 15, disa_verification
  EditBox 30, 100, 150, 15, jobs_verification
  EditBox 30, 120, 150, 15, busi_verification
  EditBox 30, 140, 150, 15, unea_verification
  EditBox 30, 160, 150, 15, acct_verification
  EditBox 55, 180, 125, 15, other_assets_verification
  EditBox 30, 200, 150, 15, shel_verification
  EditBox 45, 220, 135, 15, subsidy_verification
  EditBox 30, 240, 150, 15, insa_verification
  EditBox 55, 260, 125, 15, other_proof_verification
  CheckBox 5, 280, 325, 15, "Check here to case note that 2919 A/B, or other DHS approved form, was used for initial request.", twenty_nine_nineteen_requested
  EditBox 70, 300, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 240, 300, 50, 15
    CancelButton 290, 300, 50, 15
  Text 5, 5, 50, 15, "Case Number:"
  Text 5, 25, 260, 10, "This script will ultimately put all information entered in to a WORD document. "
  Text 5, 40, 25, 15, "ADDR:"
  Text 5, 60, 65, 15, "SCHL/STIN/STEC:"
  Text 5, 80, 25, 15, "DISA:"
  Text 5, 100, 25, 15, "JOBS:"
  Text 5, 120, 25, 15, "BUSI:"
  Text 5, 140, 25, 15, "UNEA:"
  Text 5, 160, 25, 15, "ACCT:"
  Text 5, 180, 50, 15, "Other Assets:"
  Text 5, 200, 25, 15, "SHEL:"
  Text 5, 220, 40, 15, "*SUBSIDY:"
  Text 5, 240, 20, 15, "INSA:"
  Text 5, 260, 45, 15, "Other Proofs:"
  Text 5, 300, 60, 15, "Worker Signature:"
  GroupBox 190, 35, 145, 215, "NOTES:"
  Text 200, 50, 130, 50, "*2919 IS MANDATORY:                           This script IS NOT a replacement for the DHS-2919 (Verification Request Form A/B, or other DHS approved request form), which must be used to initially request verifications. "
  Text 200, 105, 130, 45, "*REMEMBER:                                       We cannot require a client to provide a specific form of verification. We must accept any form of verification that meets policy requirements."
  Text 200, 155, 125, 40, "*SUBSIDY:                                    Verification of housing subsidy and exceptions to counting the subsidy are mandatory verifications for MFIP."
  Text 200, 200, 130, 45, "*MANDATORY VERIFICATIONS:            For more information about mandatory verifications at application and renewal/recertification refer to CM 0010.18"
EndDialog
'END DIALOGS=================================================================================================================

'THE SCRIPT==================================================================================================================

'Connects to BlueZone
EMConnect ""

'Grabs the MAXIS case number automatically
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------
DO
	err_msg = ""                                       		'Blanks this out every time the loop runs. If mandatory fields aren't entered, this variable is updated below with messages, which then display for the worker.
	Dialog Verifications_Still_Needed_Dialog              'The Dialog command shows the dialog. Replace sample_dialog with your actual dialog pasted above.
	IF ButtonPressed = cancel THEN StopScript          		'If the user pushes cancel, stop the script

'Handling for error messaging (in the case of mandatory fields or fields requiring a specific format)-----------------------------------
'If a condition is met...          ...then the error message is itself, plus a new line, plus an error message...           ...Then add a comment explaining your reason it's mandatory.
	IF IsNumeric(MAXIS_case_number) = FALSE or len(MAXIS_case_number) > 8 	THEN err_msg = err_msg & vbNewLine & "* You must type a valid numeric case number."     'MAXIS_case_number should be mandatory in most cases. Bulk or nav scripts are likely the only exceptions
	IF worker_signature = ""           													THEN err_msg = err_msg & vbNewLine & "* You must sign your case note!"                  'worker_signature is usually also a mandatory field
	IF twenty_nine_nineteen_requested = unchecked 							THEN err_msg = err_msg & vbNewLine & "* If DHS-2919 (or other DHS approved form) was not used for initial verification request, take appropriate action. Do not proceed with this script. Verifications NEED to be requested using DHS-2919 or other DHS approved form."
    '<<Follow the above template to add more mandatory fields!!>>
		call check_for_MAXIS(FALSE)											'Makes sure the user isn't passworded out
		call navigate_to_MAXIS_screen("STAT", "SUMM")		'Navigates to STAT/SUMM
		EMReadScreen summ_check, 4, 2, 46						'Reads to see that you are on STAT/SUMM
	IF summ_check <> "SUMM" THEN err_msg = err_msg & vbNewLine & "* Your case number appears to be invalid. It may be privileged." 'If unable to get to STAT/SUMM gives error message.

'If the error message isn't blank, it'll pop up a message telling you what to do!
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue."     '
LOOP UNTIL err_msg = ""     'It only exits the loop when all mandatory fields are resolved!
'End dialog section-----------------------------------------------------------------------------------------------

'Checks Maxis for password prompt
CALL check_for_MAXIS(FALSE)

'Now it navigates to a blank case note
start_a_blank_case_note

'...and enters a title
CALL write_variable_in_case_note("***Verifications Still Needed***")

'...some editboxes or droplistboxes
CALL write_bullet_and_variable_in_case_note( "ADDR", address_verification)
CALL write_bullet_and_variable_in_case_note( "SCHL/STIN/STEC", schl_stin_stec_verification)
CALL write_bullet_and_variable_in_case_note( "DISA", disa_verification)
CALL write_bullet_and_variable_in_case_note( "JOBS", jobs_verification)
CALL write_bullet_and_variable_in_case_note( "BUSI", busi_verification)
CALL write_bullet_and_variable_in_case_note( "UNEA", unea_verification)
CALL write_bullet_and_variable_in_case_note( "ACCT", acct_verification)
CALL write_bullet_and_variable_in_case_note( "Other Assets", other_assets_verification)
CALL write_bullet_and_variable_in_case_note( "SHEL", shel_verification)
CALL write_bullet_and_variable_in_case_note( "Subsidy", subsidy_verification)
CALL write_bullet_and_variable_in_case_note( "INSA", insa_verification)
CALL write_bullet_and_variable_in_case_note( "Other Proofs", other_proof_verification)

'...checkbox responses
If twenty_nine_nineteen_requested = checked THEN CALL write_variable_in_case_note( "* DHS-2919, or other DHS approved form, was used for initial verification request.")

'...and a worker signature.
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'Jumping to STAT
call navigate_to_MAXIS_screen("stat", "memb")

'Pulling household and worker info for the letter
call navigate_to_MAXIS_screen("stat", "addr") 														'Navigates to STAT/ADDR

EMReadScreen mailing_address_line1, 22, 13, 43																							'Reads first line of mailing address
IF mailing_address_line1 = "______________________" THEN          													'If nothing on first line of mailing address, uses physical address
		EMReadScreen addr_line1, 21, 6, 43																											'Reads first line of physical address
		EMReadScreen addr_line2, 21, 7, 43																											'Reads second line of physical address
		EMReadScreen addr_city, 14, 8, 43																												'Reads city of physical address
		EMReadScreen addr_state, 2, 8, 66																												'Reads state of physical address
		EMReadScreen addr_zip, 5, 9, 43																													'Reads zip of physical address
		hh_address = addr_line1 & " " & addr_line2 																							'Combines first and second lines of physical address in to new variable
		hh_address_line2 = addr_city & " " & addr_state & " " & addr_zip												'Combines city, state, and zip of physical address in to new variable
		hh_address = replace(hh_address, "_", "") & vbCrLf & replace(hh_address_line2, "_", "")	'Cleans up the new variables, makes them pretty for the Word doc
ELSE
		EMReadScreen mailing_address_line1, 21, 13, 43																					'If there is text on the first line of the mailing address then the next few lines do the same as above, only it uses the mailing address info insted.
		EMReadScreen mailing_address_line2, 21, 14, 43
		EMReadScreen mailing_address_city, 15, 15, 43
		EMReadScreen mailing_address_state, 2, 16, 43
		EMReadScreen mailing_address_zip, 5, 16, 52
		hh_address = mailing_address_line1 & " " & mailing_address_line2
		hh_address_line2 = mailing_address_city & " " & mailing_address_state & " " & mailing_address_zip
		hh_address = replace(hh_address, "_", "") & vbCrLf & replace(hh_address_line2, "_", "")
END IF


'Collecting and formatting client name for Word doc
call navigate_to_MAXIS_screen("stat", "memb")
call find_variable("Last: ", last_name, 24)
call find_variable("First: ", first_name, 11)
client_name = first_name & " " & last_name
client_name = replace(client_name, "_", "")


'Writing the Word Doc
Set objWord = CreateObject("Word.Application")			'Creates new/blank Word doc
Const wdDialogFilePrint = 88
Const end_of_doc = 6
objWord.Caption = "Verifications Still Needed"
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Arial"									'Sets the font
objSelection.Font.Size = "14"											'Sets the font size
objSelection.TypeParagraph()
objSelection.TypeText client_name									'Enters the name collected/formatted above
objSelection.TypeParagraph()
objSelection.TypeText hh_address									'Enters the address collected/formatted above
objSelection.TypeParagraph()											'These are the same as return or line down
objSelection.TypeParagraph()

objSelection.TypeText "We recently received and processed several of your requested verifications. Unfortunately there is some information that is still needed/outstanding. Failure to return this information may result in the closure and/or denial of your case. Please provide the following information at your earliest possible convenience: "
objSelection.TypeParagraph()
objSelection.TypeParagraph()

Set objRange = objSelection.Range
objDoc.Tables.Add objRange, 12, 2									'Creates a table that is twelve rows, and two columns
set objTable = objDoc.Tables(1)

'Fills in the table with the information from the dialog box
objTable.Cell(1, 1).Range.Text =	"Address/Residency: "
objTable.Cell(1, 2).Range.Text = 	address_verification
objTable.Cell(2, 1).Range.Text = 	"Financial Aid/Expenses/Student Status: "
objTable.Cell(2, 2).Range.Text = 	schl_stin_stec_verification
objTable.Cell(3, 1).Range.Text = 	"Disability: "
objTable.Cell(3, 2).Range.Text = 	disa_verification
objTable.Cell(4, 1).Range.Text = 	"Earned Income:"
objTable.Cell(4, 2).Range.Text = 	jobs_verification
objTable.Cell(5, 1).Range.Text = 	"Self-Employment:"
objTable.Cell(5, 2).Range.Text = 	busi_verification
objTable.Cell(6, 1).Range.Text = 	"Unearned Income:"
objTable.Cell(6, 2).Range.Text = 	unea_verification
objTable.Cell(7, 1).Range.Text =	"Accounts:"
objTable.Cell(7, 2).Range.Text = 	acct_verification
objTable.Cell(8, 1).Range.Text = 	"Other Assets:"
objTable.Cell(8, 2).Range.Text = 	other_assets_verification
objTable.Cell(9, 1).Range.Text = 	"Shelter Costs:"
objTable.Cell(9, 2).Range.Text = 	shel_verification
objTable.Cell(10, 1).Range.Text = "Housing Subsidy:"
objTable.Cell(10, 2).Range.Text = subsidy_verification
objTable.Cell(11, 1).Range.Text = "Other Health Insurance:"
objTable.Cell(11, 2).Range.Text = insa_verification
objTable.Cell(12, 1).Range.Text = "Other Proofs/Verifications:"
objTable.Cell(12, 2).Range.Text = other_proof_verification

objTable.AutoFormat(16)

objSelection.EndKey end_of_doc
objSelection.TypeParagraph()

'Starts the print dialog
objword.dialogs(wdDialogFilePrint).Show

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")
