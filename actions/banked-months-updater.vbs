'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - BANKED MONTHS UPDATER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("09/13/2023", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

' ASSUMPTIONS - Review these as TODO's 
'  The case has a WREG panel​
'  The “FS PWE” and “Defer FSET” fields on STAT/WREG are filled out​
'  The case has no background edits and is otherwise ready to approve​
'  The script is being run on or after [1st of the month after the month where most clients’ ABAWD months will run out post-pandemic]​
'  The script is being run before [date that MNIT shuts off banked months]

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 90, "Case Number/Date Selection Dialog"
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, MAXIS_footer_month
  EditBox 100, 25, 20, 15, MAXIS_footer_year
  EditBox 75, 45, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 65, 50, 15
    CancelButton 120, 65, 50, 15
  Text 10, 30, 65, 10, "Footer month/year:"
  Text 25, 10, 50, 10, "Case Number: "
  Text 10, 50, 60, 10, "Worker Signature:"
EndDialog

DO
	DO
		err_msg = ""
		dialog Dialog1
		cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

'TODO: evaluate to support more than one member. 

DO
	DO
		err_msg = ""
		'Creating a custom dialog for determining who the HH members are
		call HH_member_custom_dialog(HH_member_array)
		if ((ubound(HH_member_array) >= 1) or (ubound(HH_member_array) = -1)) then err_msg = err_msg & vbNewLine & "You must select exactly one household member"
		if err_msg <> "" then MsgBox err_msg
	LOOP UNTIL  err_msg = ""
call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

Call MAXIS_background_check
CALL navigate_to_MAXIS_screen("STAT", "WREG")	'Navigate to STAT/WREG and check for WREG Status codes

EMReadScreen WREG_STATUS, 2, 8, 50 '  (30)
EMReadScreen TLR_STATUS, 2, 13, 50 '  (10/13) 
IF WREG_STATUS = "30" and (TLR_STATUS = 10 or TLR_STATUS = 13) THEN 
	Call write_value_and_transmit("X", 13, 57)		'navigate to ABAWD/TLR Tracking panel and check for historical months

	'todo: update to use fixed clock not 36 month lookback period. 
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

	'todo: 
		'1. add validation point to ensure that 3 months have been used prior to case noting 
		'2. adjust case note verbiage - case might not close, might just be members. 
		'3. Handle for more than the current footer month (initial month)
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
		

		CALL start_a_blank_CASE_NOTE
		CALL write_variable_in_CASE_NOTE("Second Banked Month Used")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE("*This case meets no TLR exemptions and has used 3 TLR months.")
		CALL write_variable_in_CASE_NOTE( "*Banked month approved on:" & date &" for the month of "&MAXIS_footer_month&"/"&MAXIS_footer_year)
		CALL write_variable_in_CASE_NOTE("*This is the last available banked month.")
		CALL write_variable_in_CASE_NOTE("SPEC/WCOM sent to resident about SNAP Banked Months.")
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

		'TODO List 
		'1. make the banked months count more dynamic. 
		'2. Send MEMO prior to Case note. Leave case note in edit mode. 
		'3. Update MEMO to have styling so it's not just a blob of text. Also only call it once. 
		
		'SPEC/MEMO is being sent in leiu of SPEC/WCOM per Bulletin #23-01-02 https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=mndhs-063946

		CALL start_a_blank_CASE_NOTE
		CALL write_variable_in_CASE_NOTE("First Banked Month Used")
		CALL write_variable_in_CASE_NOTE("---")
		CALL write_variable_in_CASE_NOTE("*This case meets no TLR exemptions and has used 3 TLR months.")
		CALL write_variable_in_CASE_NOTE( "*Banked month approved on:" & date &" for the month of "&MAXIS_footer_month&"/"&MAXIS_footer_year)
		CALL write_variable_in_CASE_NOTE("*One more banked month is available.")
		CALL write_variable_in_CASE_NOTE("SPEC/WCOM sent to resident about SNAP Banked Months.")
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

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------