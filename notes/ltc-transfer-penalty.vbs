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
call changelog_update("05/03/2023", "Updated dialog to replace waiver checkboxes with drops downs, replaced transfer type dropdown with open text field for assets and added policy buttons", "Megan Geissler, Hennepin County")
call changelog_update("01/09/2020", "Updated dialog err messaging to include all errors in single message box.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT BODY----------------------------------------------------------------------------------------------------
EMConnect ""														'Connecting to Bluezone
Call check_for_MAXIS(False) 'checking for an active MAXIS session
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)		'function autofills the footer month and footer year
call MAXIS_case_number_finder(MAXIS_case_number)							'function autofills case number that worker already has on MAXIS screen
get_county_code

Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 146, 70, "Case number dialog"
  EditBox 80, 5, 60, 15, MAXIS_case_number
  EditBox 80, 25, 25, 15, MAXIS_footer_month
  EditBox 115, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 45, 50, 15
    CancelButton 90, 45, 50, 15
  Text 35, 10, 45, 10, "Case number: "
  Text 15, 30, 65, 10, "Footer Month/Year:"
EndDialog

DO
    DO
        err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
        dialog Dialog1				'main dialog
        cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'information gathering to auto-populate LTC_transfer_penalty_dialog
'grabbing app date from STAT/PROG

Call check_for_MAXIS(False) 'checking for an active MAXIS session
Call navigate_to_MAXIS_screen("STAT", "PROG")
'Verifying case has HC status of PEND or ACTV, otherwise end script.
EMReadScreen HC_status, 4, 12, 74
If (HC_status <> "PEND" AND HC_status <> "ACTV") Then Call script_end_procedure("~LTC Transfer Penalty script cancelled because case does not have HC in PEND or ACTV status.")

'reading HC Appl date
EMReadScreen app_month, 2, 12, 33
EMReadScreen app_date, 2, 12, 36
EMReadScreen app_year, 2, 12, 39
'converts date variables into date format
date_of_application = (app_month & "/" & app_date & "/" & app_year)

'navigating to the Uncompensated Transfer Calculation in the ELIG/HC person test
'Reading Uncompensated Transfer screen for case note
Call navigate_to_MAXIS_screen ("ELIG", "HC__")        'selects MEMB 01
Call write_value_and_transmit("X", 8, 26)             'selects Person test
Call write_value_and_transmit("X", 7, 17)             'selects Uncompensated transfer calculation
Call write_value_and_transmit("X", 18, 3)             'reading information from Uncompensated Transfer Calculation
EMReadScreen row_1, 71, 5, 6 'reading 1st row of text Uncompensated Transfer Calculation
EMReadScreen row_3, 71, 7, 6 'reading 2nd row of text Uncompensated Transfer Calculation
EMReadScreen row_4, 71, 8, 6 'reading 3rd row of text Uncompensated Transfer Calculation
EMReadScreen row_5, 71, 9, 6 'reading 4th row of text Uncompensated Transfer Calculation
EMReadScreen row_6, 71, 10, 6 'reading 5th row of text Uncompensated Transfer Calculation
EMReadScreen row_8, 71, 12, 6 'reading 6th row of text Uncompensated Transfer Calculation
EMReadScreen row_9, 71, 13, 6 'reading 7th row of text Uncompensated Transfer Calculation

'Reading specific information from Uncompensated Transfer Calculation screen
'Transfer date
EMReadScreen SAPSNF_check, 8, 7, 46
EMReadScreen SAPSNF_month, 2, 7, 46
EMReadScreen SAPSNF_day, 2, 7, 49
EMReadScreen SAPSNF_year, 2, 7, 52
'converts date variables into date format
If Trim(SAPSNF_check <> "__ __ __") THEN transfer_date = (SAPSNF_month & "/" & SAPSNF_day & "/" & SAPSNF_year)

'Transfer amount
EMReadScreen transfer_amount, 10, 10, 11
transfer_amount = Replace(transfer_amount, "_", "")

'partial penalty amount
EMReadScreen partial_penalty_amount, 11, 10, 66
If partial_penalty_amount <> "          " THEN partial_penalty_amount = trim(partial_penalty_amount)

'Last full month of penalty
EMReadScreen penalty_end_month, 2, 12, 46
EMReadScreen penalty_end_year, 2, 12, 49
If Trim(penalty_end_month <> "__") and penalty_end_year <> "__" Then last_full_month_of_period = penalty_end_month & "/" & penalty_end_year

'Dollar bill symbol will be added to numeric variables----------------------------------------------------------------------------------------------------
IF transfer_amount <> "" THEN transfer_amount = "$" & transfer_amount
IF partial_penalty_amount <> "" THEN partial_penalty_amount = "$" & partial_penalty_amount

'-------------------------------------------------------------------------------------------------DIALOG
'Define buttons
EPM_baseline = 101 
EPM_transfer_penalty = 102
TE_transfers = 103

DO
	  DO
        err_msg = ""
        Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 331, 270, "LTC Transfer Penalty Information"
          EditBox 65, 45, 60, 15, baseline_date
          EditBox 250, 45, 60, 15, period_begins
          EditBox 65, 65, 60, 15, transfer_date
          EditBox 250, 65, 60, 15, last_full_month_of_period
          EditBox 65, 85, 60, 15, transfer_amount
          EditBox 250, 85, 60, 15, partial_penalty_amount
          EditBox 125, 115, 60, 15, date_of_application
          EditBox 125, 135, 60, 15, date_client_was_otherwise_eligible
          DropListBox 125, 155, 60, 15, "Select one..."+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Requested"+chr(9)+"N/A", waiver_list
          EditBox 100, 180, 215, 15, hardship_waiver_details
          EditBox 100, 200, 215, 15, assets_transfered
          EditBox 100, 220, 215, 15, case_action
          EditBox 120, 245, 80, 15, worker_signature
          ButtonGroup ButtonPressed
            OkButton 205, 245, 50, 15
            CancelButton 260, 245, 50, 15
            PushButton 10, 15, 100, 15, "EPM: Baseline Date Criteria", EPM_baseline
            PushButton 110, 15, 80, 15, "EPM: Transfer Penalty", EPM_transfer_penalty
            PushButton 190, 15, 135, 15, "Uncompensated Transfers TE02.14.27", TE_transfers
          GroupBox 5, 5, 325, 30, "LTC- Transfer Penalty Resources"
          Text 15, 50, 50, 10, "Baseline date:"
          Text 170, 50, 80, 10, " Transfer period begins:"
          Text 15, 70, 45, 10, "Transfer date:"
          Text 140, 70, 110, 10, " Last full month of transfer period:"
          Text 5, 90, 55, 10, "Transfer amount:"
          Text 170, 90, 80, 10, " Partial penalty amount:"
          Text 5, 120, 115, 10, "Date of application or LTC request:"
          Text 5, 140, 115, 10, "Date client was otherwise eligible:"
          Text 35, 160, 80, 10, "Hardship Waiver Status:"
          Text 15, 185, 85, 10, "   Hardship waiver details:"
          Text 5, 205, 95, 10, "Describe Assets Transferred:"
          Text 50, 225, 50, 10, "Actions taken:"
          Text 55, 250, 60, 10, "Worker signature:"
        EndDialog
           
        Dialog Dialog1
        cancel_confirmation 	'asks user if they really want to cancel.  If yes, then script stops.  If no, loops back to dialog.
        
        If ButtonPressed > 100 Then
            err_msg = "Loop"

            If ButtonPressed = EPM_baseline Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://hcopub.dhs.state.mn.us/epm/2_4_1_3_1.htm?rhhlterm=baseline%20date&rhsearch=baseline%20date"
            If ButtonPressed = EPM_transfer_penalty Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://hcopub.dhs.state.mn.us/epm/2_4_1_3_2.htm?rhhlterm=ltc&rhsearch=LTC"
            If ButtonPressed = TE_transfers Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/sites/hs-es-poli-temp/_layouts/15/Doc.aspx?sourcedoc=%7B4980f2be-4d02-498a-a31c-77970e663c6a%7D&action=view&wdAccPdf=0&wdparaid=5A58606E"
        Else
            If (Trim(len(baseline_date) < 8) or IsDate(baseline_date) = FALSE) then err_msg = err_msg & vbNewLine & "You must enter a valid baseline date in the MM/DD/YYYY format.  See policy reference at the bottom of the dialog box if you are unsure of how to determined the baseline date."
            If (Trim(period_begins = "") or IsDate(period_begins) = FALSE) then err_msg = err_msg & vbNewLine & "You must enter the start date of the transfer penalty."
            If (Trim(transfer_date = "") or IsDate(transfer_date) = FALSE) then err_msg = err_msg & vbNewLine & "You must enter a valid transfer date." 
            If (Trim(last_full_month_of_period = "") or IsDate(last_full_month_of_period)= FALSE) then err_msg = err_msg & vbNewLine & "You must enter the last full month of the transfer penalty MM/YY."
            If (Trim(transfer_amount = "") or IsNumeric(transfer_amount) = FALSE) then err_msg = err_msg & vbNewLine & "You must enter a valid transfer penalty amount."
            If (Trim(partial_penalty_amount = "") or IsNumeric(partial_penalty_amount) = False) then err_msg = err_msg & vbNewLine & "You must enter the partial penalty amount, even if the amount is 0."
            If (Trim(date_of_application = "") or IsDate(date_of_application) = FALSE) then err_msg = err_msg & vbNewLine & "You must enter a valid application date."
            If (Trim(date_client_was_otherwise_eligible = "") or IsDate(date_client_was_otherwise_eligible) = FALSE) then err_msg = err_msg & vbNewLine & "You must enter a valid date that the client was otherwise eligible for MA."
            If waiver_list = "Select one..." then err_msg = err_msg & vbNewLine & "You must select the waiver status."
            If Trim(assets_transfered = "") then err_msg = err_msg & vbNewLine & "You must describe the assets transferred."            
            If Trim(case_action = "") then err_msg = err_msg & vbNewLine & "You must case note the case action."
            If Trim(worker_signature = "") then err_msg = err_msg & vbNewLine & "You must sign your case note!"
            If err_msg <> "" AND left(err_msg, 4) <> "Loop" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        End If
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'ensures we're in MAXIS, the case is not PRIV and it's in-county. 
Call check_for_MAXIS(false)
Call navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)
If is_this_priv = True Then script_end_procedure ("This case is privileged and cannot be accessed. The script will now stop.")
EMReadScreen county_code, 4, 21, 14
If county_code <> worker_county_code then script_end_procedure("This case is out-of-county, and cannot access CASE:NOTE. The script will now stop.")

'CALCULATIONS AND LOGIC----------------------------------------------------------------------------------------------------

'Determines lookback period based on baseline date provided
If baseline_date <> "" then 
  lookback_period = dateadd("m", -60, cdate(baseline_date)) & ""
  end_of_lookback = dateadd ("d", -1, cdate (baseline_date))
End If

'THE CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE												'navigates to CASE/NOTE and put case into edit mode
Call write_variable_in_case_note ("~~~TRANSFER PENALTY~~~")     			'adding information to case note
Call write_variable_in_case_note ("Uncompensated Transfer Information")
Call write_bullet_and_variable_in_case_note ("Transfer date", transfer_date)
Call write_bullet_and_variable_in_case_note ("Transfer amount", trim(transfer_amount))
Call write_bullet_and_variable_in_case_note ("Date of application or LTC request", date_of_application)
Call write_bullet_and_variable_in_case_note ("Baseline Date", baseline_date)
Call write_bullet_and_variable_in_case_note ("Date client was otherwise eligible", date_client_was_otherwise_eligible)
Call write_bullet_and_variable_in_case_note ("Lookback period", lookback_period & "-" & end_of_lookback)
Call write_bullet_and_variable_in_case_note ("Transfer period begins", period_begins)
Call write_bullet_and_variable_in_case_note ("Last full month of transfer period", last_full_month_of_period)
Call write_bullet_and_variable_in_case_note ("Partial penalty amount", trim(partial_penalty_amount))
Call write_bullet_and_variable_in_case_note ("Hardship waiver", waiver_list)
Call write_bullet_and_variable_in_case_note ("Hardship waiver details", hardship_waiver_details)
Call write_bullet_and_variable_in_case_note ("Transferred assets described", assets_transfered)
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



'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/03/2023
'--Tab orders reviewed & confirmed----------------------------------------------05/03/2023
'--Mandatory fields all present & Reviewed--------------------------------------05/03/2023
'--All variables in dialog match mandatory fields-------------------------------05/03/2023
'Review dialog names for content and content fit in dialog----------------------05/03/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/03/2023
'--CASE:NOTE Header doesn't look funky------------------------------------------05/03/2023
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/03/2023
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------05/03/2023
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/03/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------NA
'--PRIV Case handling reviewed -------------------------------------------------05/03/2023
'--Out-of-County handling reviewed----------------------------------------------05/03/2023
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/03/2023
'--BULK - review output of statistics and run time/count (if applicable)--------NA
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------05/03/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/03/2023
'--Incrementors reviewed (if necessary)-----------------------------------------05/03/2023
'--Denomination reviewed -------------------------------------------------------05/03/2023
'--Script name reviewed---------------------------------------------------------05/03/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------NA

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/03/2023
'--comment Code-----------------------------------------------------------------05/03/2023
'--Update Changelog for release/update------------------------------------------05/03/2023
'--Remove testing message boxes-------------------------------------------------05/03/2023
'--Remove testing code/unnecessary code-----------------------------------------05/03/2023
'--Review/update SharePoint instructions----------------------------------------05/03/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/03/2023
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/03/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------05/03/2023
'--Complete misc. documentation (if applicable)---------------------------------05/03/2023
'--Update project team/issue contact (if applicable)----------------------------05/03/2023