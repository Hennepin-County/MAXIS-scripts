'Required for statistical purposes==========================================================================================
 name_of_script = "NOTICES - METHOD B WCOM.vbs"
 start_time = timer
 STATS_counter = 1                          'sets the stats counter at one
 STATS_manualtime = 140                      'manual run time in seconds
 STATS_denomination = "C"                   'C is for each case
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
 call changelog_update("12/27/2016", "Script can now write to a MEMO is a waiting notice is not available/found.", "Charles Potter, DHS")
 call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

 'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
 changelog_display
 'END CHANGELOG BLOCK =======================================================================================================

 'Dialogs----------------------------------------------------------------------------------------------------
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

BeginDialog MEMOS_LTC_METHOD_B_dialog, 0, 0, 281, 270, "Method B budget deductions for WCOM"
  EditBox 75, 85, 40, 15, medi_part_a
  EditBox 195, 85, 40, 15, health_insa
  EditBox 75, 105, 40, 15, medi_part_b
  EditBox 195, 110, 40, 15, remedial_care
  EditBox 75, 130, 40, 15, medi_part_d
  EditBox 195, 130, 40, 15, other_deductions
  CheckBox 5, 160, 275, 10, "Check here is client pays for room/ board in addition to spenddown (GRH clients).", GRH_check
  EditBox 50, 190, 40, 15, recipient_amt
  ButtonGroup ButtonPressed
    PushButton 95, 195, 70, 10, "Calculate recip amt", CALC_button
    OkButton 170, 190, 50, 15
    CancelButton 225, 190, 50, 15
    PushButton 130, 20, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 170, 20, 25, 10, "BILS", BILS_button
    PushButton 195, 20, 25, 10, "FACI", FACI_button
    PushButton 220, 20, 25, 10, "HCMI", HCMI_button
    PushButton 245, 20, 25, 10, "UNEA", UNEA_button
  EditBox 75, 20, 40, 15, income
  EditBox 75, 45, 40, 15, income_standard
  EditBox 195, 45, 40, 15, SD
  Text 5, 25, 60, 10, "Budgeted income:"
  Text 135, 135, 60, 10, "Other deductions:"
  Text 30, 135, 40, 10, "Medi part D:"
  Text 30, 110, 40, 10, "Medi part B:"
  Text 30, 90, 40, 10, "Medi part A:"
  GroupBox 125, 10, 150, 25, "STAT based navigation"
  GroupBox 25, 70, 245, 85, "Deductions"
  Text 150, 90, 40, 10, "Health insa:"
  Text 180, 50, 15, 10, "SD:"
  Text 20, 170, 235, 10, "(This will add text on the notice about the additional cost of room/board.)"
  Text 5, 50, 70, 10, "MA income standard:"
  Text 10, 195, 35, 10, "Recip amt:"
  Text 140, 115, 50, 10, "Remedial care:"
  Text 20, 230, 245, 35, "The 'Calculate recip amt' will calculate the recipient amount based on the infromation inputted into the deductions edit boxes. If you calculate the recipeint amount, and add another deduction, please hit the calculate button again. Otherwise the cleint's recipient amount will be incorrect."
  GroupBox 0, 215, 275, 55, "Using the 'Calculate recip amt' button"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabs the case number
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder (MAXIS_footer_month, MAXIS_footer_year)

Call check_for_MAXIS(False)

Do
  err_msg = ""
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
  IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP UNTIL err_msg = ""

'ensures user is in correct footer month'
back_to_SELF
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
transmit

Call navigate_to_MAXIS_screen ("STAT", "MEDI")
'checking for if Medicare is applied to spenddown or not
EMReadscreen medicare_spenddown_check, 1, 11, 71
IF medicare_spenddown_check = "Y" THEN
	EMReadScreen Medicare_A, 8, 7, 46
	EMReadScreen Medicare_B, 8, 7, 73
END IF

'cleaning up and creating variables to be autofilled into the dialog
IF Medicare_A = "________" then Medicare_A = ""
If Medicare_B = "________" then Medicare_B = ""
medi_part_a = (Medicare_A)
medi_part_b = (Medicare_B)

'GETTING INCOME STANDARD AND SPENDOWN AMOUNTS
call navigate_to_MAXIS_screen("ELIG", "HC")
EMSendKey "x"
transmit
EMReadScreen method_type, 1, 13, 21
If method_type <> "B" then script_end_procedure("Your case is not a Method B budget case. The script will now end.")

'finding the correct income, SD and and income standard for footer month selected'
footer_info = MAXIS_footer_month & "/" & MAXIS_footer_year  'turnes footer year and footer month into string'
row = 6
col = 1                'establishes the row to start searching'
EMSearch footer_info, row, col    'searches for footer_info'
EMReadScreen MA_income_standard, 7, row + 10, col
EMReadScreen Income, 7, row + 9, col
EMReadScreen spenddown, 7, row + 11, col
spenddown = Ltrim(spenddown)
If spenddown = "" then script_end_procedure("Your case does not have a spenddown amount. The script will now end.")

'cleaning up the variables for the dialog
income_standard = Ltrim(MA_income_standard)
income = Ltrim(Income)
SD = Ltrim(spenddown)
medi_part_a = Ltrim(medi_part_a)
medi_part_b = Ltrim(medi_part_b)

'Shows the dialog
Do
  err_msg = ""
  Do
		Dialog MEMOS_LTC_METHOD_B_dialog
    cancel_confirmation
    MAXIS_Dialog_navigation
		If ButtonPressed = CALC_button THEN
			'makes the deduction amounts = 0 so the Abs(number) function work
			If medi_part_a = "" THEN medi_part_a = "0"
			If medi_part_b = "" THEN medi_part_b = "0"
			If medi_part_d = "" THEN medi_part_d = "0"
			If health_insa = "" THEN health_insa = "0"
			If remedial_care = "" THEN remedial_care = "0"
			If other_deductions = "" THEN other_deductions = "0"
			recipient_amt = Abs(SD) - Abs(medi_part_a) - Abs(medi_part_b) - Abs(medi_part_d) - Abs(health_insa) - Abs(remedial_care) - Abs(other_deductions) & ""
      If medi_part_a = "0" THEN medi_part_a = ""
      If medi_part_b = "0" THEN medi_part_b = ""
      If medi_part_d = "0" THEN medi_part_d = ""
      If health_insa = "0" THEN health_insa = ""
      If remedial_care = "0" THEN remedial_care = ""
      If other_deductions = "0" THEN other_deductions = ""
	  End if
    Loop until ButtonPressed = -1
  IF IsNumeric(recipient_amt) = False then err_msg = err_msg & vbNewLine & "* Enter the recipient amount."
  IF IsNumeric(income_standard) = False then err_msg = err_msg & vbNewLine & "* Enter the MA income standard."
  IF IsNumeric(income) = False then err_msg = err_msg & vbNewLine & "* Enter the budgeted income."
  IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  Loop until err_msg = ""

recipient_amt = Round(recipient_amt)  'rounds variable to nearest decimal point to clean up for memo'
recipient_amt = recipient_amt & ".00"
'This section will check for whether forms go to AREP and SWKR
call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
EMReadscreen forms_to_arep, 1, 10, 45
call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
EMReadscreen forms_to_swkr, 1, 15, 63


'THE MEMO----------------------------------------------------------------------------------------------------------------
CALL navigate_to_MAXIS_screen("SPEC", "WCOM")
Emwritescreen "Y", 3, 74  'sorts by HC notices
Transmit
'Searching for waiting HC notice
wcom_row = 6
Do
	wcom_row = wcom_row + 1
	Emreadscreen program_type, 2, wcom_row, 26
	Emreadscreen print_status, 7, wcom_row, 71
	If program_type = "HC" then
		If print_status = "Waiting" then
			Emwritescreen "x", wcom_row, 13
			exit Do
		End If
	End If
	If wcom_row = 17 then
		PF8
		Emreadscreen spec_edit_check, 6, 24, 2
		wcom_row = 6
	end if
	If spec_edit_check = "NOTICE" THEN no_hc_waiting = true
Loop until spec_edit_check = "NOTICE"

' If no notice was found then we give the option to write the message in a SPEC MEMO instead
If no_hc_waiting = true then
  swap_to_memo = msgbox ("No waiting HC results were found for the requested month. Would you like to send MEMO in place of WCOM?", vbYesNo)  'fancy message box with yes/no
END IF
'based on output of fancy message box we either end the script or write the WCOM
IF swap_to_memo = vbNo THEN script_end_procedure("No waiting HC results were found for the requested month")
IF swap_to_memo = vbYes THEN
  CALL navigate_to_MAXIS_screen("SPEC","MEMO")
			PF5
			'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
			row = 4                             'Defining row and col for the search feature.
			col = 1
			EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
			IF row > 4 THEN                     'If it isn't 4, that means it was found.
				arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
				call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
				EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
				call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
				PF5                                                     'PF5s again to initiate the new memo process
			END IF
			'Checking for SWKR
			row = 4                             'Defining row and col for the search feature.
			col = 1
			EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
			IF row > 4 THEN                     'If it isn't 4, that means it was found.
				swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
				call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
				EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
				call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
				PF5                                           'PF5s again to initiate the new memo process
			END IF
			EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
			IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
			transmit
      Write_variable_in_SPEC_MEMO("Although your spenddown is $" & spenddown & " your recipient amount (the amount you pay each month) is $" & recipient_amt & ". This is how the recipient amount is determined:")
      Write_variable_in_SPEC_MEMO("Income: $" & income &" - MA Income Standard $" & income_standard & " = $" & spenddown)
      Write_variable_in_SPEC_MEMO("Spenddown:            $" & spenddown)
      If medi_part_a <> "" then Write_variable_in_SPEC_MEMO("Medicare Part A     - $" & medi_part_a)
      If medi_part_b <> "" then Write_variable_in_SPEC_MEMO("Medicare Part B     - $" & medi_part_b)
      If medi_part_d <> "" then Write_variable_in_SPEC_MEMO("Medicare Part D     - $" & medi_part_d)
      If remedial_care <> "" then Write_variable_in_SPEC_MEMO("Remedial care       - $" & remedial_care)
      If other_deductions <> "" then Write_variable_in_SPEC_MEMO("Other deductions    - $" & other_deductions)
      If health_insa <> "" then Write_variable_in_SPEC_MEMO("Health insurance    - $" & health_insa)
      Call Write_variable_in_SPEC_MEMO("Recipient amount:   = $" & recipient_amt)
      If GRH_check = 1 Then Write_variable_in_SPEC_MEMO("This amount is in addition to your room and board.")
      Write_variable_in_SPEC_MEMO("Please contact the agency with any questions. Thank you.")
			PF4
      script_end_procedure("Success! Your MEMO has been written. Please review it for accuracy, and PF4 to save.")
END IF

'transmitting and putting wcom into edit mode
Transmit
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
'Worker Comment Input
Write_variable_in_SPEC_MEMO("Although your spenddown is $" & spenddown & " your recipient amount (the amount you pay each month) is $" & recipient_amt & ". This is how the recipient amount is determined:")
Write_variable_in_SPEC_MEMO("Income: $" & income &" - MA Income Standard $" & income_standard & " = $" & spenddown)
Write_variable_in_SPEC_MEMO("Spenddown:            $" & spenddown)
If medi_part_a <> "" then Write_variable_in_SPEC_MEMO("Medicare Part A     - $" & medi_part_a)
If medi_part_b <> "" then Write_variable_in_SPEC_MEMO("Medicare Part B     - $" & medi_part_b)
If medi_part_d <> "" then Write_variable_in_SPEC_MEMO("Medicare Part D     - $" & medi_part_d)
If remedial_care <> "" then Write_variable_in_SPEC_MEMO("Remedial care       - $" & remedial_care)
If other_deductions <> "" then Write_variable_in_SPEC_MEMO("Other deductions    - $" & other_deductions)
If health_insa <> "" then Write_variable_in_SPEC_MEMO("Health insurance    - $" & health_insa)
Call Write_variable_in_SPEC_MEMO("Recipient amount:   = $" & recipient_amt)
If GRH_check = 1 Then Write_variable_in_SPEC_MEMO("This amount is in addition to your room and board.")
Write_variable_in_SPEC_MEMO("Please contact the agency with any questions. Thank you.")

script_end_procedure("Success! Your WCOM has been written. Please review it for accuracy, and PF4 to save.")
