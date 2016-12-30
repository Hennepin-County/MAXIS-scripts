'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CLIENT TRANSPORTATION COSTS.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The Dialog--------------------------------------------------------------

BeginDialog client_transportation_dialog, 0, 0, 146, 95, "Transportation Funds Issued"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 15, 75, 50, 15
    CancelButton 80, 75, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  DropListBox 10, 45, 125, 15, "Select One..."+chr(9)+"Gas Card Issued"+chr(9)+"Mileage Reimbursement"+chr(9)+"Bus Tokens Issued", How_funds_issued_dropbox
  Text 10, 30, 130, 10, "Please select how funds were issued:"
EndDialog



BeginDialog gas_card_dialog, 0, 0, 286, 125, "Gas Card Dialog"
  EditBox 55, 5, 70, 15, MAXIS_case_number
  EditBox 225, 5, 50, 15, date_cards_given
  DropListBox 105, 25, 65, 15, "Select One..."+chr(9)+"10"+chr(9)+"20"+chr(9)+"30"+chr(9)+"40", card_amt_dropbox
  CheckBox 5, 45, 145, 10, "Client Signed Fuel Card Acknowledgement", client_signed_stmt_checkbox
  EditBox 160, 60, 45, 15, amt_given_yr_to_date
  EditBox 100, 80, 105, 15, card_number
  EditBox 80, 100, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 170, 100, 50, 15
    CancelButton 225, 100, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 150, 10, 70, 10, "Gas Cards issued on:"
  Text 5, 30, 95, 10, "Amount of Gas Cards given:"
  Text 5, 65, 150, 10, "Total amount given this year (including today):"
  Text 5, 85, 90, 10, "Gas Card Numbers Given:"
  Text 5, 105, 70, 10, "Gas Card Issued By:"
EndDialog



BeginDialog mileage_dialog, 0, 0, 316, 125, "Mileage Reimbursement"
  EditBox 55, 5, 70, 15, MAXIS_case_number
  EditBox 230, 5, 70, 15, date_docs_recd
  EditBox 55, 25, 70, 15, total_reimbursement
  EditBox 230, 25, 70, 15, date_to_accounting
  EditBox 55, 45, 250, 15, docs_reqd
  EditBox 55, 65, 250, 15, other_notes
  EditBox 55, 85, 245, 15, actions_taken
  EditBox 70, 105, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 105, 50, 15
    CancelButton 250, 105, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 170, 10, 55, 10, "Date Received:"
  Text 5, 30, 45, 10, "Total Amount:"
  Text 165, 30, 60, 10, "Date Sent to Acct:"
  Text 5, 50, 40, 10, "Doc's req'd:"
  Text 5, 70, 45, 10, "Other notes:"
  Text 5, 90, 50, 10, "Actions taken:"
  Text 5, 110, 60, 10, "Worker signature:"
EndDialog

BeginDialog bus_tokens_dialog, 0, 0, 146, 105, "Bus Tokens Issued"
  EditBox 55, 5, 75, 15, MAXIS_case_number
  EditBox 90, 25, 40, 15, date_tokens_issued
  EditBox 90, 45, 40, 15, Amount_tokens_given
  EditBox 70, 65, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 20, 85, 50, 15
    CancelButton 85, 85, 50, 15
  Text 5, 10, 45, 10, "Case number:"
  Text 5, 30, 80, 15, "Date bus tokens issued:"
  Text 5, 50, 80, 10, "Amount of tokens given:"
  Text 5, 70, 60, 10, "Worker Signature:"
EndDialog



'----------------The Script---------------------------------------------------------------------

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'Connects to BlueZone
EMConnect ""

Call check_for_MAXIS(True)

'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(MAXIS_case_number)


'Starting with the 1st dialog box asking how funds were issued
DO
	Err_msg = ""
	Dialog client_transportation_dialog
	cancel_confirmation
	If How_funds_issued_dropbox = "Select one..." THEN err_msg = err_msg & vbNewLine & "*You must select how transportation funds were issued"
	If MAXIS_case_number = "" THEN err_msg = err_msg & vbNewLine & "*You must enter a case number"
	IF err_msg <> "" THEN Msgbox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
Loop until err_msg = ""

'Runs the Gas card dialog box if selected
If How_funds_issued_dropbox = "Gas Card Issued" then
	Do
		err_msg = ""
		Dialog gas_card_dialog
		cancel_confirmation
		IF card_amt_dropbox = "Select one..." THEN err_msg = err_msg & vbNewLine & "*You must select the amount of Gas Cards given"
		If amt_given_yr_to_date = "" THEN err_msg = err_msg & vbNewLine & "*Enter the amount given this year"
		if MAXIS_case_number = "" THEN err_msg = err_msg & vbNewLine & "*You must enter a case number"
		If worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*You must sign your case note"
		'amt_given_yr_to_date = "$" & amt_given_yr_to_date
		'card_amt_dropbox = "$" & card_amt_dropbox
		If err_msg <> "" THEN Msgbox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
	Loop until err_msg = ""
End If

'Runs Mileage reimbursement dialog if selected
If How_funds_issued_dropbox = "Mileage Reimbursement" Then
	Do
		err_msg = ""
		Dialog Mileage_dialog
		cancel_confirmation
		If MAXIS_case_number = "" THEN err_msg = err_msg & vbNewLine & "*You must enter a case number"
		If worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*You must sign your case note"
		If err_msg <> "" Then Msgbox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
	Loop until err_msg = ""
End if

'Runs Bus Tokens dialog if selected
If How_funds_issued_dropbox = "Bus Tokens Issued" THEN
	Do
		err_msg = ""
		Dialog bus_tokens_dialog
		cancel_confirmation
		If MAXIS_case_number = "" THEN err_msg = err_msg & vbNewLine & "*You must enter a case number"
		If worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*You must sign your case note"
		If err_msg <> "" Then Msgbox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
	Loop until err_msg = ""
End if

'Checks for active MAXIS session
CALL check_for_MAXIS(FALSE)

Dim Client_transportation
If How_funds_issued_dropbox = "Gas Card Issued" Then Client_transportation = Client_transportation & "Gas Card Issued"
If How_funds_issued_dropbox = "Mileage Reimbursement" THEN Client_transportation = Client_transportation & "Mileage Reimbursement"
If How_funds_issued_dropbox = "Bus Tokens Issued" THEN Client_transportation = Client_transportation & "Bus Tokens Issued"


'Writes the case note
start_a_blank_CASE_NOTE

CALL write_variable_in_case_note("Client Transportation: " & Client_transportation)   'Writes title in Case note
CALL write_bullet_and_variable_in_case_note("Gas Cards issued on", date_cards_given)                                                'Writes date cards were issued on next line
If card_amt_dropbox <> "" THEN CALL write_bullet_and_variable_in_case_note("Amount of Fuel Cards Given", "$" & card_amt_dropbox)                                         'Write the amt given this
CALL write_bullet_and_variable_in_case_note("Gas Card Numbers", card_number)                                                        'Writes the gas card numbers
If amt_given_yr_to_date <> "" then CALL write_bullet_and_variable_in_case_note("Total Amount Given This Year Including Today", "$" & amt_given_yr_to_date)   	       	    'Writes amt given year to date
IF client_signed_stmt_checkbox = 1 THEN CALL write_variable_in_CASE_NOTE("* Client signed Fuel Card Acknowledgement Form")          'Writes if the client signed stmt if that box was checked
IF card_amt_dropbox >= "40" THEN
	thirty_days_from_now = DateAdd ("d", 30, date)
	CALL write_bullet_and_variable_in_case_note("Next Gas Card Can be Given On", thirty_days_from_now)                             'If $40 is selected, then will write a line telling FW when the next cards can be issued
END IF

call write_bullet_and_variable_in_case_note("Date received", date_docs_recd)
If total_reimbursement <> "" Then call write_bullet_and_variable_in_case_note("Total Amount", "$" & total_reimbursement)
call write_bullet_and_variable_in_case_note("Date Sent to Accounting", date_to_accounting)
call write_bullet_and_variable_in_case_note("Docs requested", docs_reqd)
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
If How_funds_issued_dropbox = "Mileage Reimbursement" AND worker_county_code = "x179" then call write_variable_in_CASE_NOTE("* Please note: DO NOT SCAN!! Accounting will scan into OnBase when processed.")	'Should only do this for Wabasha County, unless other counties request it.

call write_bullet_and_variable_in_case_note("Date bus tokens issued", date_tokens_issued)
call write_bullet_and_variable_in_case_note("Amount of tokens issued", Amount_tokens_given)

CALL write_variable_in_case_note ("---")
CALL write_variable_in_CASE_NOTE(worker_signature)    'Writes worker signature in note

script_end_procedure("")
