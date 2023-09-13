'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - BANKED MONTHS UPDATER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 49                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
'END OF stats block=========================================================================================================


'LOADING FUNCTIONS LIBRARY FROM REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		FuncLib_URL = script_repository & "MAXIS FUNCTIONS LIBRARY.vbs"
		critical_error_msgbox = MsgBox ("The Functions Library code was not able to be reached by " &name_of_script & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Send issues to " & contact_admin , _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
	ELSE
		FuncLib_URL = script_repository & "MAXIS FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================


' ASSUMPTIONS
'  The case has a WREG panel​
'  The “FS PWE” and “Defer FSET” fields on STAT/WREG are filled out​
'  The case has no background edits and is otherwise ready to approve​
'  The script is being run on or after [1st of the month after the month where most clients’ ABAWD months will run out post-pandemic]​
'  The script is being run before [date that MNIT shuts off banked months]


'CHANGELOG BLOCK ===========================================================================================================
'("09/01/2023" "Initial version.", "Jared Peterson, DHS")
'END CHANGELOG BLOCK =======================================================================================================


BeginDialog Case_selection_dialog, 0, 0, 156, 80, "Banked Months Updater dialog"
  EditBox 65, 10, 80, 15, MAXIS_case_number
  EditBox 85, 35, 20, 15, MAXIS_footer_month
  EditBox 110, 35, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 15, 55, 50, 15
    CancelButton 80, 55, 50, 15
  Text 5, 15, 50, 10, "Case Number:"
  Text 5, 35, 65, 10, "Footer month/year:"
EndDialog

BeginDialog confirmation_dialog, 0, 0, 156, 80, "Update case confirmation dialog"
  EditBox 70, 5, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 15, 60, 50, 15
    CancelButton 80, 60, 50, 15
  Text 5, 10, 60, 15, "Worker signature:"
  Text 5, 30, 150, 30, "OK to update the tracking record and add a DAIL/WRIT, CASE/NOTE, and SPEC/MEMO?"
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
'Hunts for Maxis case number and footer month/year to autofill
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

DO
	DO
		dialog Case_selection_dialog
		IF buttonpressed = 0 THEN stopscript
		IF MAXIS_case_number = "" THEN MSGBOX "Please enter a case number"
	LOOP UNTIL MAXIS_case_number <> ""
call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

DO
	DO
		err_msg = ""
		'Creating a custom dialog for determining who the HH members are
		call HH_member_custom_dialog(HH_member_array)
		if ((ubound(HH_member_array) >= 1) or (ubound(HH_member_array) = -1)) then err_msg = "You must select exactly one household member"
		if err_msg <> "" then MsgBox err_msg
	LOOP UNTIL  err_msg = ""
call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

'Navigate to STAT/WREG and check for WREG Status codes
CALL navigate_to_MAXIS_screen("STAT", "WREG")
'Checking to see if the case is stuck in background.
row = 1
col = 1
EMSearch "Background", row, col
IF row <> 0 THEN script_end_procedure("The case is stuck in background. Please try again.")


EMReadScreen WREG_STATUS, 2, 8, 50 '  (30)
EMReadScreen TLR_STATUS, 2, 13, 50 '  (10/13) 
IF WREG_STATUS = "30" and (TLR_STATUS = 10 or TLR_STATUS = 13)  THEN 
	'navigate to ABAWD/TLR Tracking panel and check for historical months
	EMWriteScreen "X", 13, 57
	Transmit
	defaultrow = 10
	EMReadScreen yearfindervariable, 2, defaultrow, 15  ' getting the relevant year's location in the Tracking Panel - the rows migrate upwards over time
	row = (defaultrow - (yearfindervariable-MAXIS_footer_year))
	'cols are organized in a grid with 4 spaces between months, starting at position 19
	col = 15 + (MAXIS_footer_month * 4)

	EMReadScreen datapoint, 1, row, col
	if ((datapoint <> "_") and (datapoint<>"D")) then script_end_procedure("It appears that the TLR tracking record for this month has already been updated. Please run this script in another footer month or delete the data in the tracking record and try again.")
	TLR_Months = 0
	Banked_months = 0
	for i = 0 to 36
		EMReadScreen datapoint, 1, row, col
		if  (datapoint = "X" or datapoint = "M") then TLR_Months = TLR_Months + 1
		if (datapoint = "B" or datapoint = "C") then Banked_Months = Banked_Months + 1
		if col = 19 then
			row = row - 1
			col = 63
		else 
			col = col - 4
		end if
	Next
	PF3 	'exit Tracking Record
	If TLR_Months < 3 then script_end_procedure("This client has not yet used all three TLR months. Please assess for TLR months before using banked months.")
	MsgBox("The client has used 3 TLR Months and "&Banked_Months&" Banked months.")
	DO
		DO
			dialog confirmation_dialog
			IF buttonpressed = 0 THEN stopscript
			IF worker_signature = "" THEN MSGBOX "Please enter a worker signature"
		LOOP UNTIL worker_signature <> ""
		call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = false

	TIKL_M = DatePart("M", date) + 1
	if TIKL_M = 13 then TIKL_M = 1
	TIKL_M = right("0" & TIKL_M , 2)
	TIKL_D = "01"
	TIKL_Y = right( DatePart("YYYY",date) , 2)
	if TIKL_M = "01" then TIKL_Y = TIKL_Y + 1

	IF Banked_Months = 2 THEN
		CALL start_a_blank_CASE_NOTE
		CALL write_variable_in_CASE_NOTE("Case no longer eligible for banked months")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE("*This case has used 3 TLR months and 2 banked months.")
		CALL write_variable_in_CASE_NOTE("Case must be closed if no TLR exemptions are met")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE("*Processed using an automated script*")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE(worker_signature)
		Transmit
	end if
	if Banked_Months = 1 then
		PF9
		EMWriteSCreen "13", 13, 50
		EMWriteScreen "2", 14, 50
		Transmit
		EMWriteScreen "X", 13, 57
		Transmit
		defaultrow = 10
		EMReadScreen yearfindervariable, 2, defaultrow, 15  ' getting the relevant year's location in the Tracking Panel - the rows migrate upwards over time
		row = (defaultrow - (yearfindervariable-MAXIS_footer_year))
		'cols are organized in a grid with 4 spaces between months, starting at position 19
		col = 15 + (MAXIS_footer_month * 4)
		PF9
		EMWriteSCreen "C", row,col
		PF3
		Transmit
		CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
		Emwritescreen TIKL_M, 5, 18
		Emwritescreen TIKL_D, 5, 21
		EMWritescreen TIKL_Y, 5, 24
		write_variable_in_TIKL("This case has used both available banked months. Please close the case if the client does not meet any TLR exemptions")
		Back_to_self
		CALL start_a_blank_CASE_NOTE
		CALL write_variable_in_CASE_NOTE("Second Banked Month Used")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE("*This case meets no TLR exemptions and has used 3 TLR months.")
		CALL write_variable_in_CASE_NOTE( "*Banked month approved on:" & date &" for the month of "&MAXIS_footer_month&"/"&MAXIS_footer_year)
		CALL write_variable_in_CASE_NOTE("*This is the last available banked month.")
		CALL write_variable_in_CASE_NOTE( "TIKL set and SPEC/WCOM sent to client.")
		CALL write_variable_in_CASE_NOTE("Case will be closed if no TLR exemptions are met")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE("*Processed using an automated script*")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE(worker_signature)
		Transmit
		CALL start_a_new_spec_memo()
		CALL write_variable_in_spec_memo("You are getting this letter because you or someone in your SNAP unit needs to follow the time-limited work rules and have used all three available months. Unless you or someone in your SNAP unit meet work rules or an exemption, you/they will no longer be eligible for SNAP. However, due to additional funding we are able to approve SNAP benefits for up to 2 more months.	If you/someone in your SNAP unit is not meeting work requirements/meeting an exemption, you/they will no longer receive SNAP after these 2 months. Please contact your worker if you, or someone in your SNAP unit, start meeting work requirements or think that you meet an exemption. If you need help meeting these work requirements, please see the SNAP Time-limited work rules website at: https://mn.gov/dhs/snap-e-and-t/time-limited-work-rules/.")
	end if
	if Banked_Months = 0 then
		PF9
		EMWriteSCreen "13", 13, 50
		EMWriteScreen "1", 14, 50
		Transmit
		EMWriteScreen "X", 13, 57
		Transmit
		defaultrow = 10
		EMReadScreen yearfindervariable, 2, defaultrow, 15  ' getting the relevant year's location in the Tracking Panel - the rows migrate upwards over time
		row = (defaultrow - (yearfindervariable-MAXIS_footer_year))
		'cols are organized in a grid with 4 spaces between months, starting at position 19
		col = 15 + (MAXIS_footer_month * 4)
		PF9
		EMWriteSCreen "C", row,col
		PF3
		Back_to_self
		CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
		Emwritescreen TIKL_M, 5, 18
		Emwritescreen TIKL_D, 5, 21
		EMWritescreen TIKL_Y, 5, 24
		write_variable_in_TIKL("This case has used one banked month. Remember to update and approve the second banked month")
		Back_to_self
		CALL start_a_blank_CASE_NOTE
		CALL write_variable_in_CASE_NOTE("First Banked Month Used")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE("*This case meets no TLR exemptions and has used 3 TLR months.")
		CALL write_variable_in_CASE_NOTE( "*Banked month approved on:" & date &" for the month of "&MAXIS_footer_month&"/"&MAXIS_footer_year)
		CALL write_variable_in_CASE_NOTE("*One more banked month is available.")
		CALL write_variable_in_CASE_NOTE("TIKL set and SPEC/WCOM sent to client.")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE("*Processed using an automated script*")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE(worker_signature)
		Transmit
		CALL start_a_new_spec_memo()
		CALL write_variable_in_spec_memo("You are getting this letter because you or someone in your SNAP unit needs to follow the time-limited work rules and have used all three available months. Unless you or someone in your SNAP unit meet work rules or an exemption, you/they will no longer be eligible for SNAP. However, due to additional funding we are able to approve SNAP benefits for up to 2 more months.	If you/someone in your SNAP unit is not meeting work requirements/meeting an exemption, you/they will no longer receive SNAP after these 2 months. Please contact your worker if you, or someone in your SNAP unit, start meeting work requirements or think that you meet an exemption. If you need help meeting these work requirements, please see the SNAP Time-limited work rules website at: https://mn.gov/dhs/snap-e-and-t/time-limited-work-rules/.")
	end if

else
	script_end_procedure("Based on coding on STAT/WREG, this case does not appear to be eligible for banked months.") 
end if

script_end_procedure("Please remember to APP.")
STATS_counter = STATS_counter - 1			'Removing one instance of the STATS Counter