'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - BANKED MONTHS WCOMS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

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
CALL changelog_update("07/19/2018", "Added option to inform client of possible Unfit for Employment exemption due to homelessness.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/02/2018", "Added option to inform client of the voluntary E & T option. Removed WCOM option to close for non-cooperation.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC function. Updated script to support change.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC function. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("04/04/2017", "Added handling for multiple recipient changes to SPEC/WCOM", "David Courtright, St Louis County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Dialogs
BeginDialog case_number_dlg, 0, 0, 186, 90, "Case Number Dialog"
  EditBox 115, 10, 60, 15, MAXIS_case_number
  EditBox 130, 30, 20, 15, approval_month
  EditBox 155, 30, 20, 15, approval_year
  EditBox 70, 50, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 70, 50, 15
    CancelButton 125, 70, 50, 15
  Text 50, 35, 70, 10, "Approval Month/Year:"
  Text 60, 15, 50, 10, "Case Number: "
  Text 5, 55, 60, 10, "Worker signature:"
EndDialog

BeginDialog banked_months_menu_dialog, 0, 0, 346, 120, "Banked Months WCOMs"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 90, 10, "All Banked Months Used", banked_months_used_button
    PushButton 10, 40, 90, 10, "Banked Months Notifier", banked_months_notifier
    PushButton 10, 60, 90, 10, "Homeless Exemption Info", homeless_button
    PushButton 10, 80, 90, 10, "Voluntary E and T info", voluntary_button
    CancelButton 295, 100, 45, 15
  Text 110, 40, 200, 10, "-- Notifying the client they may be eligible for banked months."
  Text 110, 80, 200, 10, "-- Informing recipients of the voluntary E and T option."
  GroupBox 5, 5, 335, 90, "Select the following WCOM to add to an approval when:"
  Text 110, 20, 230, 10, "-- SNAP is closing because they used all their eligible banked months."
  Text 110, 60, 230, 10, "-- A recipient may meet Unfit for Employment expansion (homeless)."
EndDialog

'--- The script -----------------------------------------------------------------------------------------------------------------
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
approval_month = CM_plus_1_mo
approval_year = CM_plus_1_yr

Do 
    DO
    	err_msg = ""
    	dialog case_number_dlg
    	If ButtonPressed = cancel then stopscript
    	IF MAXIS_case_number = "" THEN err_msg = "* Please enter a case number" & vbNewLine
    	IF len(approval_month) <> 2 THEN err_msg = err_msg & "* Enter the approval month in MM format." & vbNewLine
    	IF len(approval_year) <> 2 THEN err_msg = err_msg & "* Enter the approval year in YY format." & vbNewLine
        If trim(worker_signature) = "" THEN err_msg = err_msg & "* Enter your worker signature." & vbNewLine
    	IF err_msg <> "" THEN msgbox err_msg
    LOOP until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

CALL check_for_MAXIS(false)

Do 
	DIALOG banked_months_menu_dialog
	If ButtonPressed = cancel then stopscript
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

If ButtonPressed = banked_months_used_button then 
    wcom_text = "You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please contact your team if you have questions."
Elseif ButtonPressed = banked_months_notifier then 
    wcom_text = "You have used all of your available ABAWD months. You may be eligible for SNAP banked months. Please contact your team if you have questions."  
Elseif ButtonPressed = voluntary_button then 
    wcom_text = "You have been approved to receive additional SNAP benefits under the SNAP Banked Months program. Working with Employment and Training is voluntary under this program. If you'd like work with Employment and Training, please contact your team."    
End if 

Call navigate_to_MAXIS_screen("spec", "wcom")
EMWriteScreen approval_month, 3, 46
EMWriteScreen approval_year, 3, 51
transmit

DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
	EMReadScreen more_pages, 8, 18, 72
	IF more_pages = "MORE:  -" THEN PF7
LOOP until more_pages <> "MORE:  -"

read_row = 7
DO
	waiting_check = ""
	EMReadscreen prog_type, 2, read_row, 26
	EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
	If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
		EMSetcursor read_row, 13
		Call write_value_and_transmit("X", read_row, 13)
		pf9
		EMSetCursor 03, 15
        If ButtonPressed = homeless_button then 
            Call write_variable_in_SPEC_MEMO("You receive SNAP under time-limited Banked Months, due to being an ABAWD (Able-bodied adult without dependents). You previously reported to this agency that you are homeless. Based on this information, you may meet an ABAWD exemption and qualify for SNAP benefits that are not time-limited.")
            Call write_variable_in_SPEC_MEMO("One exemption is for a person who is unfit for employment if he or she is currently homeless. Homelessness is specifically defined for this purpose as:")
            Call write_variable_in_SPEC_MEMO("*Lacking a fixed and regular nighttime residence, including temporary housing situations, AND")
            Call write_variable_in_SPEC_MEMO("*Lacking access to work-related necessities (i.e. shower/laundry facilities, etc.).")
            Call write_variable_in_SPEC_MEMO("If you believe you may meet this exemption (or any other potential ABAWD exemption), please contact your team.")      
        Else 
		 CALL write_variable_in_SPEC_MEMO(wcom_text)
        End if 
		PF4
		PF3
		WCOM_count = WCOM_count + 1
		exit do
	ELSE
		read_row = read_row + 1
	END IF
	IF read_row = 18 THEN
		PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
		read_row = 7
	End if
LOOP until prog_type = "  "

'Outcome ---------------------------------------------------------------------------------------------------------------------
If WCOM_count = 0 THEN  'if no waiting FS notice is found
	script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
ELSE 					'If a waiting FS notice is found
	'Case note
	start_a_blank_case_note
	call write_variable_in_CASE_NOTE("---WCOM added regarding banked months---")
	If ButtonPressed = banked_months_used_button then 
		CALL write_variable_in_CASE_NOTE("* WCOM added because client all eligible banked months have been used.")
	ELSEIf ButtonPressed = voluntary_button then 
		CALL write_variable_in_CASE_NOTE("* WCOM added to inform client of voluntary option to work with E & T.")
	ELSEIf ButtonPressed = banked_months_notifier then 
		CALL write_variable_in_CASE_NOTE("* Client has used ABAWD counted months and MAY be eligible for banked months. Eligibility questions should be directed to financial worker.")
    Elseif ButtonPressed = homeless_button then 
        CALL write_variable_in_CASE_NOTE("* Client may meet the unfit for employment exapansion for homeless recipients detailed in CM 11.24.")
	END IF
	call write_variable_in_CASE_NOTE("---")
    call write_variable_in_CASE_NOTE(worker_signature)
END IF

script_end_procedure("Your WCOM has been made for the approval month selected! Please review the notice for any additional clarifications that are needed.")