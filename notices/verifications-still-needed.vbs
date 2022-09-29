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
CALL changelog_update("09/20/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
'Connects to BlueZone
EMConnect ""

'Grabs the MAXIS case number automatically
CALL MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 366, 300, "Verifications Still Needed"
  EditBox 55, 5, 40, 15, MAXIS_case_number
  CheckBox 15, 20, 320, 15, "Check here to case note that 2919 A/B or other DHS approved form was used for initial request.", twentynine_nineteen_requested_CHECKBOX
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
  EditBox 55, 280, 125, 15, worker_signature
  Text 5, 10, 50, 10, "Case Number:"
  Text 200, 50, 150, 45, "*2919 IS MANDATORY:                               This script is NOT a replacement for the DHS-2919 (Verification Request Form A/B or other DHS approved request form) which must be used to initially request verifications. "
  Text 200, 100, 160, 35, "*REMEMBER:                                                    We cannot require a client to provide a specific form of verification. We must accept any form of verification that meets policy requirements."
  Text 200, 175, 155, 35, "*SUBSIDY:                                               Verification of housing subsidy and exceptions to counting the subsidy are mandatory verifications for MFIP."
  Text 200, 220, 155, 35, "*MANDATORY VERIFICATIONS:                             For more information about mandatory verifications at application and renewal/recertification refer to CM 0010.18"
  Text 105, 10, 245, 10, "This script creates a word document for you to send to the client and ECF."
  Text 5, 45, 25, 10, "ADDR:"
  Text 5, 65, 65, 10, "SCHL/STIN/STEC:"
  Text 5, 85, 25, 10, "DISA:"
  Text 5, 105, 25, 10, "JOBS:"
  Text 5, 125, 25, 10, "BUSI:"
  Text 5, 145, 25, 10, "UNEA:"
  Text 5, 165, 25, 10, "ACCT:"
  Text 5, 185, 50, 10, "Other Assets:"
  Text 5, 205, 25, 10, "SHEL:"
  Text 5, 225, 40, 10, "*SUBSIDY:"
  Text 5, 245, 20, 10, "INSA:"
  Text 5, 265, 45, 10, "Other Proofs:"
  GroupBox 190, 35, 170, 230, "IMPORTANT REMINDERS:"
  Text 5, 285, 45, 10, "Worker Sig:"
  ButtonGroup ButtonPressed
    OkButton 255, 270, 50, 15
    CancelButton 310, 270, 50, 15
EndDialog

'Dialog
DO      'Password DO loop
	DO  'Conditional handling DO loop
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		IF IsNumeric(MAXIS_case_number) = FALSE or len(MAXIS_case_number) > 8 	THEN err_msg = err_msg & vbNewLine & "* You must type a valid numeric case number."     'MAXIS_case_number should be mandatory in most cases. Bulk or nav scripts are likely the only exceptions
		IF twentynine_nineteen_requested_CHECKBOX = unchecked THEN err_msg = err_msg & vbNewLine & "* If DHS-2919 (or other DHS approved form) was not used for initial verification request, take appropriate action. Do not proceed with this script. Verifications NEED to be requested using DHS-2919 or other DHS approved form."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)                                 'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false
'Checks Maxis for password prompt
CALL check_for_MAXIS(FALSE)

'this reads clients current mailing address
Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_zip, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
If mail_line_one = "" then
     client_address = resi_line_one & " " & resi_line_two & " " & resi_city & ", " & resi_state & " " & resi_zip
Else
	client_address =  mail_line_one & " " & mail_line_two & " " & mail_city & ", " & mail_state & " " & mail_zip
End If

'Collecting and formatting client name for Word doc
Call navigate_to_MAXIS_screen("STAT", "MEMB")
call find_variable("Last: ", last_name, 24)
call find_variable("First: ", first_name, 11)
client_name = first_name & " " & last_name
client_name = replace(client_name, "_", "")

'Generates Word Doc Form
Set objWord = CreateObject("Word.Application")
objWord.Caption = "Verifications Still Needed"
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.ParagraphFormat.Alignment = 0
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.ParagraphFormat.SpaceBefore = 0
objSelection.ParagraphFormat.SpaceAfter = 0
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "12"
objSelection.TypeParagraph
objSelection.Font.Bold = True
objSelection.TypeText "Verifications Still Needed"
objSelection.TypeParagraph
objSelection.TypeParagraph
objSelection.TypeText "Hennepin County Human Services & Public Health Department"
objSelection.TypeParagraph
objSelection.TypeText "PO Box 107, Minneapolis, MN 55440-0107"
objSelection.TypeParagraph
objSelection.TypeText "FAX: 612-288-2981"
objSelection.TypeParagraph
objSelection.TypeText "Phone: 612-596-1300"
objSelection.TypeParagraph
objSelection.TypeText "Email: HHSEWS@hennepin.us"
objSelection.TypeParagraph

objSelection.ParagraphFormat.Alignment = 2
objSelection.ParagraphFormat.LineSpacing = 12
objSelection.ParagraphFormat.SpaceBefore = 0
objSelection.ParagraphFormat.SpaceAfter = 0
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "DATE: " & date()

objSelection.TypeParagraph
objSelection.ParagraphFormat.Alignment = 0
objSelection.Font.Size = "10"
objSelection.Font.Bold = True
objSelection.TypeText client_name									'Enters the name collected/formatted above
objSelection.TypeParagraph
objSelection.TypeText client_address
objSelection.TypeParagraph()									'Enters the address collected/formatted above
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.Font.Bold = FALSE
objSelection.TypeText("We recently received and processed some of your required verifications. Unfortunately, there is some information that is still needed. Failure to return this information may result in the closure and/or denial of your case.")
objSelection.TypeParagraph()
	objSelection.TypeText("You now have an option to use an email to return documents to Hennepin County.")
objSelection.TypeParagraph()
	objSelection.TypeText("Email: HHSEWS@hennepin.us")
objSelection.TypeParagraph()
	objSelection.TypeText("Be sure to write the case number and full name associated with the case in the body of the email.")
objSelection.TypeParagraph()
	objSelection.TypeText("Only the following types are accepted PNG, JPG, TIFF, DOC, PDF, and HTML. You will not receive confirmation of receipt or failure.")
objSelection.TypeParagraph()
	objSelection.TypeText("To obtain information about your case please contact Hennepin County.")
objSelection.TypeParagraph()
	objSelection.TypeText("Please provide the following information at your earliest possible convenience:")
objSelection.TypeParagraph()
objSelection.Font.Bold = True										'These are the same as return or line down
IF address_verification <> "" THEN objSelection.TypeText "Address: " & address_verification
objSelection.TypeParagraph()
IF schl_stin_stec_verification <> "" THEN objSelection.TypeText "Financial Aid/Expenses/Student Status: " & schl_stin_stec_verification
objSelection.TypeParagraph()
IF disa_verification <> "" THEN objSelection.TypeText "Disability: " & disa_verification
objSelection.TypeParagraph()
IF jobs_verification <> "" THEN objSelection.TypeText "Earned Income: " & jobs_verification
objSelection.TypeParagraph()
IF busi_verification <> "" THEN objSelection.TypeText "Self-Employment: " & busi_verification
objSelection.TypeParagraph()
IF unea_verification <> "" THEN objSelection.TypeText "Unearned Income: " & unea_verification
objSelection.TypeParagraph
IF acct_verification <> "" THEN objSelection.TypeText "Accounts: " & acct_verification
objSelection.TypeParagraph()
IF other_assets_verification <> "" THEN objSelection.TypeText "Other Assets: " & other_assets_verification
objSelection.TypeParagraph
IF shel_verification <> "" THEN objSelection.TypeText "Shelter Costs: " & shel_verification
objSelection.TypeParagraph()
IF subsidy_verification <> "" THEN objSelection.TypeText  "Housing Subsidy: " & subsidy_verification
objSelection.TypeParagraph
IF insa_verification <> "" THEN objSelection.TypeText "Other Health Insurance: " & insa_verification
objSelection.TypeParagraph()
IF other_proof_verification <> "" THEN objSelection.TypeText "Other Proofs/Verifications: " & other_proof_verification
objSelection.TypeParagraph()
objSelection.TypeText "If you have any questions about this request, please contact Hennepin County at 612-596-1300"
'objSelection.EndKey end_of_doc

'Starts the print dialog
'objword.dialogs(wdDialogFilePrint).Show got an error "the requested member of the collection does not exist"

'...and enters a title
CALL start_a_blank_case_note
CALL write_variable_in_case_note("***Verifications Still Needed***")
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
If twentynine_nineteen_requested_CHECKBOX = checked THEN CALL write_variable_in_case_note( "* DHS-2919 or other DHS approved form was used for initial verification request.")
'...and a worker signature.
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("Success! Your Request for Verification has been generated, please follow up with the next steps to ensure the request is received timely. The verification request must be reflected in ECF.")
