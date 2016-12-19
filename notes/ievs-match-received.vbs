'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - IEVS MATCH RECEIVED.vbs"
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
call changelog_update("11/28/2016", "Changed the name and case note header from IEVS Notice to IEVS Match.", "Casey Love, Ramsey County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog case_number_dlg, 0, 0, 206, 75, "Enter a Case Number"
  EditBox 70, 10, 70, 15, MAXIS_case_number
  EditBox 65, 30, 30, 15, benefit_month
  EditBox 155, 30, 30, 15, benefit_year
  ButtonGroup ButtonPressed
    OkButton 100, 55, 50, 15
    CancelButton 150, 55, 50, 15
  Text 10, 15, 50, 10, "Case Number"
  Text 10, 35, 50, 10, "Benefit Month"
  Text 105, 35, 45, 10, "Benefit Year"
EndDialog

BeginDialog IEVS_Match, 0, 0, 177, 126, "IEVS Match Received"
  DropListBox 56, 4, 100, 14, "Select one"+chr(9)+"Resolved"+chr(9)+"Notice Sent to Client"+chr(9)+"Notice Sent to Employer", OPTIONS
  EditBox 50, 24, 20, 14, MEMB
  DropListBox 112, 24, 48, 14, "Select one"+chr(9)+"1st"+chr(9)+"2nd"+chr(9)+"3rd"+chr(9)+"4th"+chr(9)+"year", Quarter
  EditBox 44, 46, 110, 14, Employer
  EditBox 44, 74, 110, 14, ADDR
  ButtonGroup ButtonPressed
    OkButton 20, 100, 40, 14
    CancelButton 110, 100, 40, 14
  Text 22, 8, 26, 12, "Options:"
  Text 10, 30, 40, 14, "HH Memb:"
  Text 80, 30, 30, 14, "Quarter:"
  Text 10, 52, 34, 14, "Employer:"
  Text 10, 78, 30, 14, "Address:"
EndDialog

BeginDialog Resolved_Non_Cooperation, 0, 0, 137, 76, "IEVS Resolved-Non Cooperation"
  DropListBox 62, 6, 50, 14, "Select one"+chr(9)+"NC"+chr(9)+"CB"+chr(9)+"CC"+chr(9)+"CF"+chr(9)+"CA"+chr(9)+"CI"+chr(9)+"CP"+chr(9)+"BC"+chr(9)+"BN"+chr(9)+"BI"+chr(9)+"BP"+chr(9)+"BU"+chr(9)+"BE"+chr(9)+"BO", code
  EditBox 40, 28, 90, 14, action
  ButtonGroup ButtonPressed
    OkButton 10, 50, 40, 14
    CancelButton 90, 50, 40, 14
  Text 12, 32, 24, 12, "Action:"
  Text 20, 10, 38, 12, "Code used:"
EndDialog

'connects to BlueZone and brings it forward
EMConnect ""
EMFocus

'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
Call check_for_MAXIS(false)

'grabs the case number and benefit month/year that is being worked on
call MAXIS_case_number_finder(MAXIS_case_number)
EMReadScreen at_self, 4, 2, 50
IF at_self = "SELF" THEN
	EMReadScreen benefit_month, 2, 20, 43
	IF len(benefit_month) <> 2 THEN benefit_month = "0" & benefit_month
	EMReadScreen benefit_year, 2, 20, 46
ELSE
	CALL find_variable("Month: ", benefit_month, 2)
	IF benefit_month <> "  " THEN CALL find_variable("Month: " & benefit_month & " ", benefit_year, 2)
END IF

' >>>>> GATHERING & CONFIRMING THE MAXIS CASE NUMBER <<<<<
DO
	err_msg = ""
	DIALOG case_number_dlg
		IF ButtonPressed = 0 THEN stopscript
		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

CALL check_for_MAXIS(False)

' Starts the IEVS Match Received case note dialog
DO
	err_msg = ""
	'starts the IEVS dialog
	Dialog IEVS_Match
	'asks if you want to cancel and if "yes" is selected sends StopScript
	cancel_confirmation
	'checks if an Option has been selected
	IF OPTIONS = "Select one" THEN err_msg = err_msg & vbCr & "You must select an option."
	'checks if a HH Memb has been entered
	IF MEMB = "" THEN err_msg = err_msg & vbCr & "You must enter a HHLD MEMB."
	'checks if Quarter was selected.
	IF Quarter = "Select one" THEN err_msg = err_msg & vbCr & "You must select a time period for the IEVS Match."
	'checks if Employer was entered.
	IF Employer = "" THEN err_msg = err_msg & vbCr & "You must enter an employer."
	'checks if Address was entered.
	IF ADDR = "" THEN err_msg = err_msg & vbCr & "You must enter an address."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'starts Resolved/Non Cooperation dialog if resolved/non coop selected
IF OPTIONS = "Resolved" THEN
	DO
		err_msg = ""
		Dialog Resolved_Non_Cooperation
		'asks if you want to cancel and if "yes" is selected sends StopScript
		cancel_confirmation
		'checks if a code was selected
		IF code = "Select one" THEN err_msg = err_msg & vbCr & "You must select a code used."
		'checks if Action was completed
		IF action = "" THEN err_msg = err_msg & vbCr & "You must enter an action taken."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
End IF

'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
call check_for_MAXIS (false)

'starts a blank case note
call start_a_blank_case_note

'this enters the actual case note info
call write_variable_in_CASE_NOTE("***IEVS Match received: " & OPTIONS & "***")
call write_bullet_and_variable_in_CASE_NOTE("HHLD MEMB", MEMB)
call write_bullet_and_variable_in_CASE_NOTE("Quarter", Quarter)
call write_bullet_and_variable_in_CASE_NOTE("Employer", Employer)
call write_bullet_and_variable_in_CASE_NOTE("Address", ADDR)
IF OPTIONS = "Resolved" THEN call write_bullet_and_variable_in_CASE_NOTE("Code Used", code)
IF OPTIONS = "Resolved" THEN call write_bullet_and_variable_in_CASE_NOTE("Action", action)
call write_variable_in_CASE_NOTE ("---")
call write_variable_in_CASE_NOTE(worker_signature)
'This next line is in as a reminder to the financial worker to not add any other information to the case note to remain in compliance with the FTI rules.
call write_variable_in_CASE_NOTE ("**DO NOT ENTER ANY OTHER INFO**")

IF code = "NC" THEN MsgBox "The client was non-cooperative, remember to add a DISQ panel for this client."

script_end_procedure("")
