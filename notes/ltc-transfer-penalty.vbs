'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - TRANSFER PENALTY.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 480          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'Reference source: http://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=dhs16_150210

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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 146, 70, "Case number dialog"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 80, 25, 25, 15, MAXIS_footer_month
  EditBox 115, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 45, 50, 15
    CancelButton 90, 45, 50, 15
  Text 10, 30, 65, 10, "Footer month/year:"
  Text 10, 10, 45, 10, "Case number: "
EndDialog

BeginDialog LTC_transfer_penalty_dialog, 0, 0, 316, 295, "Dialog"
  EditBox 60, 5, 60, 15, baseline_date
  DropListBox 190, 5, 120, 15, "Select one..."+chr(9)+"Annuity"+chr(9)+"Life Estate"+chr(9)+"Uncompensated Transfer"+chr(9)+"Other", type_of_transfer_list
  EditBox 60, 25, 60, 15, transfer_date
  EditBox 250, 25, 60, 15, date_of_application
  EditBox 60, 45, 60, 15, transfer_amount
  EditBox 250, 45, 60, 15, date_client_was_otherwise_eligible
  EditBox 120, 65, 60, 15, period_begins
  EditBox 120, 85, 60, 15, last_full_month_of_period
  EditBox 120, 105, 60, 15, partial_penalty_amount
  CheckBox 210, 100, 100, 10, "Hardship waiver requested", harship_waiver_requested_check
  CheckBox 210, 115, 95, 10, "Hardship waiver approved", hardship_waiver_approved_check
  EditBox 85, 130, 225, 15, harship_waiver_details
  EditBox 85, 150, 225, 15, other_information
  EditBox 85, 170, 225, 15, case_action
  EditBox 120, 195, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 205, 195, 50, 15
    CancelButton 260, 195, 50, 15
  Text 5, 30, 45, 10, "Transfer date:"
  Text 5, 50, 55, 10, "Transfer amount:"
  Text 130, 30, 115, 10, "Date of application or LTC request:"
  Text 5, 10, 50, 10, "Baseline date:"
  Text 130, 50, 115, 10, "Date client was otherwise eligible:"
  Text 5, 70, 80, 10, "Transfer period begins:"
  Text 5, 90, 110, 10, "Last full month of transfer period:"
  Text 5, 110, 80, 10, "Partial penalty amount:"
  Text 5, 155, 65, 10, "Other transfer info:"
  Text 5, 135, 80, 10, "Hardship waiver details:"
  Text 5, 175, 50, 10, "Actions taken:"
  Text 55, 200, 60, 10, "Worker signature:"
  Text 10, 240, 295, 30, "1. A person is residing in an LTCF or, for a person requesting services through a home and community-based waiver program, the date a screening occurred that indicated a need for services provided through a home and community-based services waiver program AND"
  Text 10, 225, 270, 10, "The BASELINE DATE is the date in which both of the following conditions are met: "
  Text 10, 275, 295, 10, "2. The person's initial request month for MA payment of LTC services"
  GroupBox 0, 215, 310, 75, "Per HCPM 19.40.15:"
  Text 130, 10, 55, 10, "Type of transfer: "
EndDialog

'SCRIPT BODY----------------------------------------------------------------------------------------------------
EMConnect ""														'Connecting to Bluezone
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)		'function autofills the footer month and footer year
call MAXIS_case_number_finder(MAXIS_case_number)							'function autofills case number that worker already has on MAXIS screen
'checking for an active MAXIS session
Call check_for_MAXIS(True)

'calls up dialog for worker to enter case number and applicable month and year.
DO
	Dialog case_number_dialog
	IF buttonPressed = 0 then StopScript
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must enter a valid case number."
LOOP UNTIL MAXIS_case_number <> "" OR IsNumeric(MAXIS_case_number) = True

'information gathering to auto-populate LTC_transfer_penalty_dialog
'grabbing app date from STAT/PROG
Call navigate_to_MAXIS_screen("STAT", "PROG")
EMReadScreen app_month, 2, 12, 33
EMReadScreen app_date, 2, 12, 36
EMReadScreen app_year, 2, 12, 39
'converts date variables into date format
date_of_application = (app_month & "/" & app_date & "/" & app_year)

'navigating to the Uncompensated Transfer Calculation in the ELIG/HC person test
Call navigate_to_MAXIS_screen ("elig", "HC__")
'selects MEMB 01
EMWriteScreen "x", 8, 26
transmit
'selects Person test
EMWriteScreen "x", 7, 17
transmit
'selects Uncompensated transfer calculation
EMWriteScreen "x", 18, 3
transmit
'reading information from Uncompensated Transfer Calculation
EMReadScreen row_1, 71, 5, 6
EMReadScreen row_3, 71, 7, 6
EMReadScreen row_4, 71, 8, 6
EMReadScreen row_5, 71, 9, 6
EMReadScreen row_6, 71, 10, 6
EMReadScreen row_8, 71, 12, 6
EMReadScreen row_9, 71, 13, 6

'Reading specific information from Uncompensated Transfer Calculation screen
'Transfer date
EMReadScreen SAPSNF_check, 8, 7, 46
EMReadScreen SAPSNF_month, 2, 7, 46
EMReadScreen SAPSNF_day, 2, 7, 49
EMReadScreen SAPSNF_year, 2, 7, 52
'converts date variables into date format
If SAPSNF_check <> "" THEN transfer_date = (SAPSNF_month & "/" & SAPSNF_day & "/" & SAPSNF_year)

'Transfer amount
EMReadScreen transfer_amount, 10, 10, 11
transfer_amount = Replace(transfer_amount, "_", "")

'partial penalty amount
EMReadScreen partial_penalty_amount, 11, 10, 66
If partial_penalty_amount <> "          " THEN partial_penalty_amount = trim(partial_penalty_amount)

'Last full month of penalty
EMReadScreen penalty_end_month, 2, 12, 46
EMReadScreen penalty_end_year, 2, 12, 49
If penalty_end_month <> "" and penalty_end_year <> "" Then last_full_month_of_period = penalty_end_month & "/" & penalty_end_year


'Dollar bill symbol will be added to numeric variables----------------------------------------------------------------------------------------------------
IF transfer_amount <> "" THEN transfer_amount = "$" & transfer_amount
IF partial_penalty_amount <> "" THEN partial_penalty_amount = "$" & partial_penalty_amount

'The main transfer dialog----------------------------------------------------------------------------------------------------
DO
	DO
		DO
			DO
				DO
					DO
						DO
							DO
								DO
									DO
										DO
											DO
												Dialog LTC_transfer_penalty_dialog
												cancel_confirmation 	'asks user if they really want to cancel.  If yes, then script stops.  If no, loops back to dialog.
												IF (len(baseline_date) < 8 or IsDate(baseline_date) = FALSE) THEN Msgbox "You must enter a valid baseline date in the MM/DD/YYYY format.  See policy reference at the bottom of the dialog box if you are unsure of how to determined the baseline date."
											LOOP UNTIL len(baseline_date) >= 8 or IsDate(baseline_date) = TRUE
											IF worker_signature = "" THEN MsgBox "You must sign your case note!"
										LOOP UNTIL worker_signature <> ""
										If type_of_transfer_list = "Select one..." then MsgBox "You must select the type of transfer."
									LOOP UNTIL type_of_transfer_list <> "Select one..."
									if (type_of_transfer_list = "Other" AND other_information = "") Then MsgBox "You have selected ""Other"" as your transfer reason.  You must complete the 'other transfer info' field."
								LOOP until (type_of_transfer_list = "Other" AND other_information <> "") OR type_of_transfer_list <> "Other"
								If (transfer_date = "" or IsDate(transfer_date) = FALSE) then MsgBox "You must enter a valid transfer date."
							LOOP UNTIL transfer_date <> "" OR IsDate(transfer_date) = True
							If (date_of_application = "" or IsDate(date_of_application) = FALSE) then MsgBox "You must enter a valid application date."
						LOOP UNTIL date_of_application <> "" OR IsDate(date_of_application) = TRUE
						If (transfer_amount = "" or IsNumeric(transfer_amount) = FALSE) then MsgBox "You must enter a valid transfer penalty amount."
					LOOP UNTIL transfer_amount <> "" or IsNumeric(transfer_amount) = TRUE
					If (date_client_was_otherwise_eligible = "" or IsDate(date_client_was_otherwise_eligible) = FALSE) then MsgBox "You must enter a valid date that the client was otherwise eligible for MA."
				LOOP UNTIL date_client_was_otherwise_eligible <> "" or IsDate(date_client_was_otherwise_eligible) = TRUE
				If period_begins = "" then MsgBox "You must enter the start date of the transfer penalty."
			LOOP UNTIL period_begins <> ""
			IF last_full_month_of_period = "" then MsgBox "You must enter the last full month of the transfer penalty."
		LOOP UNTIL last_full_month_of_period <> ""
		IF (partial_penalty_amount = "" or IsNumeric(partial_penalty_amount) = False) then MsgBox "You must enter the partial penalty amount, even if the amount is 0."
	LOOP UNTIL partial_penalty_amount <> "" or IsNumeric(partial_penalty_amount) = TRUE
	If case_action = "" then MsgBox "You must case note the case action."
LOOP Until case_action <> ""

'ensures that worker has not "passworded" out of MAXIS
Call check_for_MAXIS(false)

'CALCULATIONS AND LOGIC----------------------------------------------------------------------------------------------------
'Autofill for the application_date variable, then determines lookback period based on the info
If baseline_date <> "" then lookback_period = dateadd("m", -60, cdate(baseline_date)) & ""

'Lookback period end date
If baseline_date <> "" then end_of_lookback = dateadd ("d", -1, cdate (baseline_date))

'THE CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE												'navigates to CASE/NOTE and put case into edit mode
Call write_variable_in_case_note ("~~~TRANSFER PENALTY~~~")     			'adding information to case note
Call write_bullet_and_variable_in_case_note ("Type of transfer", type_of_transfer_list )
Call write_bullet_and_variable_in_case_note ("Transfer date", transfer_date)
Call write_bullet_and_variable_in_case_note ("Transfer amount", trim(transfer_amount))
Call write_bullet_and_variable_in_case_note ("Date of application or LTC request", date_of_application)
Call write_bullet_and_variable_in_case_note ("Baseline Date", baseline_date)
Call write_bullet_and_variable_in_case_note ("Date client was otherwise eligible", date_client_was_otherwise_eligible)
Call write_bullet_and_variable_in_case_note ("Lookback period", lookback_period & "-" & end_of_lookback)
Call write_bullet_and_variable_in_case_note ("Transfer period begins", period_begins)
Call write_bullet_and_variable_in_case_note ("Last full month of transfer period", last_full_month_of_period)
Call write_bullet_and_variable_in_case_note ("Partial penalty amount", trim(partial_penalty_amount))
Call write_bullet_and_variable_in_case_note ("Other information", other_information)
IF harship_waiver_requested_check = 1 THEN Call write_variable_in_case_note ("* Hardship waiver requested")
IF hardship_waiver_approved_check = 1 THEN Call write_variable_in_case_note ("* Hardship waiver approved")
Call write_bullet_and_variable_in_case_note ("Hardship waiver details", harship_waiver_details)
Call write_bullet_and_variable_in_case_note ("Case Action", case_action)
Call write_variable_in_case_note ("---")
'writing in the information from the Uncompensated Transfer Calculation in ELIG/HC person test
Call write_variable_in_case_note (row_1)
Call write_variable_in_case_note (row_3)
Call write_variable_in_case_note (row_4)
Call write_variable_in_case_note (row_5)
Call write_variable_in_case_note (row_6)
Call write_variable_in_case_note (row_8)
Call write_variable_in_case_note (row_9)
Call write_variable_in_case_note ("---")
call write_variable_in_case_note (worker_signature)

script_end_procedure ("Success, your penalty case note has been created.  Please ensure that your notices also reflect the transfer information.")							'closing script and writing stats
