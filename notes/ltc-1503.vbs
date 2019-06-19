'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - 1503.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 360          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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

'The following code looks to find the user name of the user running the script---------------------------------------------------------------------------------------------
'This is used in arrays that specify functionality to specific workers
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation= ucase(windows_user_ID)
name_for_validation = ""

If user_ID_for_validation = "CALO001" Then name_for_validation = "Casey"
If user_ID_for_validation = "ILFE001" Then name_for_validation = "Ilse"
If user_ID_for_validation = "WFS395" Then name_for_validation = "MiKayla"
If user_ID_for_validation = "WFQ898" Then name_for_validation = "Hannah"
If user_ID_for_validation = "WFK093" Then name_for_validation = "Jessica"
If user_ID_for_validation = "WFM207" Then name_for_validation = "Mandora"
If user_ID_for_validation = "WFP803" Then name_for_validation = "Melissa"
If user_ID_for_validation = "WFC041" Then name_for_validation = "Kerry"
If user_ID_for_validation = "AAGA001" Then name_for_validation = "Aaron"
If user_ID_for_validation = "WFJ454" Then name_for_validation = "True"
If user_ID_for_validation = "WFC719" Then name_for_validation = "Kristen"
If user_ID_for_validation = "WFE269" Then name_for_validation = "Carrie"
If user_ID_for_validation = "WFW682" Then name_for_validation = "Osman"
If user_ID_for_validation = "WFC804" Then name_for_validation = "Shanna"
If user_ID_for_validation = "WFA168" Then name_for_validation = "Michelle"

If name_for_validation <> "" Then
    MsgBox "Hello " & name_for_validation &  ", you have been selected to test the script NOTES - Documents Received. The functionality to run NOTES - LTC - 1503 Script is now added to NOTES - Documents Received. When the script runs, inditate that this is an LTC case and check the box for the 1503 form."  & vbNewLine & vbNewLine & "A testing version of the script will now run.  Thank you for taking your time to review our new scripts and functionality as we strive for Continuous Improvement." & vbNewLine & vbNewLine  & "                                                                                    - BlueZone Script Team"
    testing_run = TRUE
    testing_script_url = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/testing_trial/notes/documents-received.vbs"
    Call run_from_GitHub(testing_script_url)
End if

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/13/2019", "Updated backend process handling. Added additional mandatory fields to main dialog.", "Ilse Ferris, Hennepin County")
call changelog_update("10/12/2017", "Changed ICF-MR to ICF-DD. Updated password handling on back end.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THIS SCRIPT IS BEING USED IN A WORKFLOW SO DIALOGS ARE NOT NAMED
'DIALOGS MAY NOT BE DEFINED AT THE BEGINNING OF THE SCRIPT BUT WITHIN THE SCRIPT FILE

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS & grabs the case number and footer month/year
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'The initial dialog------defined and displayed------------------------------------------------------------------------
BeginDialog , 0, 0, 116, 70, "Case number dialog"
  EditBox 60, 10, 50, 15, MAXIS_case_number
  EditBox 60, 30, 20, 15, MAXIS_footer_month
  EditBox 90, 30, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 50, 40, 15
    CancelButton 70, 50, 40, 15
  Text 10, 35, 45, 10, "Footer month:"
  Text 10, 15, 45, 10, "Case number: "
EndDialog

Do
	DO
        err_msg = ""
		Dialog 					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_without_confirmation
        IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "Enter a valid case number."		'mandatory field
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer month."
        If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit MAXIS footer year."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'THE 1503 MAIN DIALOG---------defined and displayed-----------------------------------------------------------------------
BeginDialog , 0, 0, 366, 285, "1503 Dialog"
  EditBox 55, 5, 135, 15, facility_name
  DropListBox 255, 5, 95, 15, "30 days or less"+chr(9)+"31 to 90 days"+chr(9)+"91 to 180 days"+chr(9)+"over 180 days", length_of_stay
  DropListBox 105, 25, 45, 15, "SNF"+chr(9)+"NF"+chr(9)+"ICF-DD"+chr(9)+"RTC", level_of_care
  DropListBox 215, 25, 135, 15, "acute-care hospital"+chr(9)+"home"+chr(9)+"RTC"+chr(9)+"other SNF or NF"+chr(9)+"ICF-DD", admitted_from
  EditBox 145, 45, 205, 15, hospital_admitted_from
  EditBox 75, 65, 65, 15, admit_date
  EditBox 275, 65, 75, 15, discharge_date
  CheckBox 15, 85, 155, 10, "If you've processed this 1503, check here.", processed_1503_check
  CheckBox 15, 115, 60, 10, "Updated RLVA?", updated_RLVA_check
  CheckBox 85, 115, 60, 10, "Updated FACI?", updated_FACI_check
  CheckBox 150, 115, 50, 10, "Need 3543?", need_3543_check
  CheckBox 205, 115, 55, 10, "Need 3531?", need_3531_check
  CheckBox 265, 115, 95, 10, "Need asset assessment?", need_asset_assessment_check
  EditBox 130, 130, 225, 15, verifs_needed
  CheckBox 15, 150, 85, 10, "Sent 3050 back to LTCF", sent_3050_check
  CheckBox 165, 155, 100, 10, "Sent verif req? If so, to who:", sent_verif_request_check
  ComboBox 270, 150, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"Client & AREP", sent_request_to
  CheckBox 15, 165, 120, 10, "Sent DHS-5181 to Case Manager", sent_5181_check
  EditBox 30, 185, 330, 15, notes
  CheckBox 30, 215, 260, 10, "Check here to have the script TIKL out to contact the FACI re: length of stay.", TIKL_check
  CheckBox 30, 230, 155, 10, "Check here to have the script update HCMI.", HCMI_update_check
  CheckBox 30, 245, 150, 10, "Check here to have the script update FACI.", FACI_update_check
  EditBox 115, 265, 135, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 255, 265, 50, 15
    CancelButton 310, 265, 50, 15
  Text 5, 10, 50, 10, "Facility name:"
  Text 200, 10, 50, 10, "Length of stay:"
  Text 5, 30, 95, 10, "Recommended level of care:"
  Text 160, 30, 50, 10, "Admitted from:"
  Text 5, 50, 135, 10, "If hospital, list name/dates of admission:"
  Text 5, 70, 65, 10, "Date of admission:"
  Text 165, 70, 105, 10, "Date of discharge (if applicible):"
  GroupBox 0, 100, 360, 80, "Actions/Proofs"
  Text 10, 135, 115, 10, "Other proofs needed (if applicable):"
  Text 5, 190, 25, 10, "Notes:"
  GroupBox 5, 205, 355, 55, "Script actions"
  Text 50, 270, 60, 10, "Worker signature:"
EndDialog

Do
	Do
        err_msg = ""
		Dialog  					'Calling a dialog without a assigned variable will call the most recently defined dialog
		cancel_confirmation
		If trim(facility_name) = "" then err_msg = err_msg & vbNewLine & "* Enter the name of the facility."
        If isdate(admit_date) = false then err_msg = err_msg & vbNewLine & "* Enter a valid date of admission."
        If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
    Call check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Checks for an active MAXIS session
Call check_for_MAXIS(False)
Call MAXIS_footer_month_confirmation
Call navigate_to_MAXIS_screen("STAT", "FACI")

'THE TIKL----------------------------------------------------------------------------------------------------
If TIKL_check = 1 then
  If length_of_stay = "30 days or less" then TIKL_multiplier = 30
  If length_of_stay = "31 to 90 days" then TIKL_multiplier = 90
  If length_of_stay = "91 to 180 days" then TIKL_multiplier = 180
  TIKL_date = dateadd("d", TIKL_multiplier, admit_date)
  TIKL_date_DD = datepart("d", TIKL_date)
  If len(TIKL_date_DD) = 1 then TIKL_date_DD = "0" & TIKL_date_DD
  TIKL_date_MM = datepart("m", TIKL_date)
  If len(TIKL_date_MM) = 1 then TIKL_date_MM = "0" & TIKL_date_MM
  TIKL_date_YY = datepart("yyyy", TIKL_date)
  If len(TIKL_date_YY) = 4 then TIKL_date_YY = TIKL_date_YY - 2000
End if

'UPDATING MAXIS PANELS----------------------------------------------------------------------------------------------------
'FACI
If FACI_update_check = 1 then
	call navigate_to_MAXIS_screen("stat", "faci")
	EMReadScreen panel_max_check, 1, 2, 78
	IF panel_max_check = "5" THEN
		script_end_procedure ("This case has reached the maximum amount of FACI panels.  Please review your case, delete an appropriate FACI panel, and run the script again.  Thank you.")
	ELSE
		EMWriteScreen "nn", 20, 79
		transmit
	END IF
	EMWriteScreen facility_name, 6, 43
	If level_of_care = "NF" then EMWriteScreen "42", 7, 43
	If level_of_care = "RTC" THEN EMWriteScreen "47", 7, 43
	If length_of_stay = "30 days or less" and level_of_care = "SNF" then EMWriteScreen "44", 7, 43
	If length_of_stay = "31 to 90 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
	If length_of_stay = "91 to 180 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
	if length_of_stay = "over 180 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
	If length_of_stay = "30 days or less" and level_of_care = "ICF-DD" then EMWriteScreen "44", 7, 43
	If length_of_stay = "31 to 90 days" and level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
	If length_of_stay = "91 to 180 days" and level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
	If length_of_stay = "over 180 days" and level_of_care = "ICF-DD" then EMWriteScreen "41", 7, 43
	EMWriteScreen "n", 8, 43
	Call create_MAXIS_friendly_date_with_YYYY(admit_date, 0, 14, 47)
	If discharge_date<> "" then
		Call create_MAXIS_friendly_date_with_YYYY(discharge_date, 0, 14, 71)
		transmit
		transmit
	End if
End if

'HCMI
If HCMI_update_check = 1 THEN
	call navigate_to_MAXIS_screen("stat", "hcmi")
	EMReadScreen HCMI_panel_check, 1, 2, 78
	IF HCMI_panel_check <> "0" Then
		PF9
	ELSE
		EMWriteScreen "nn", 20, 79
		transmit
	END IF
	EMWriteScreen "dp", 10, 57
	transmit
	transmit
END IF

'THE TIKL----------------------------------------------------------------------------------------------------
If TIKL_check = 1 then
	call navigate_to_MAXIS_screen("dail", "writ")
	EMWriteScreen TIKL_date_MM, 5, 18
	EMWriteScreen TIKL_date_DD, 5, 21
	EMWriteScreen TIKL_date_YY, 5, 24
	EMSetCursor 9, 3
	write_variable_in_TIKL("Have " & worker_signature & " call " & facility_name & " re: length of stay. " & TIKL_multiplier & " days expired.")
	transmit
	PF3
End if

'The CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE

If processed_1503_check = 1 then
  	call write_variable_in_CASE_NOTE("***Processed 1503 from " & facility_name & "***")
Else
  	call write_variable_in_CASE_NOTE("***Rec'd 1503 from " & facility_name & ", DID NOT PROCESS***")
End if
Call write_bullet_and_variable_in_case_note("Length of stay", length_of_stay)
Call write_bullet_and_variable_in_case_note("Recommended level of care", level_of_care)
Call write_bullet_and_variable_in_case_note("Admitted from", admitted_from)
Call write_bullet_and_variable_in_case_note("Hospital admitted from", hospital_admitted_from)
Call write_bullet_and_variable_in_case_note("Admit date", admit_date)
Call write_bullet_and_variable_in_case_note("Discharge date", discharge_date)
Call write_variable_in_CASE_NOTE("---")
If updated_RLVA_check = 1 and updated_FACI_check = 1 then
	Call write_variable_in_CASE_NOTE("* Updated RLVA and FACI.")
Else
  	If updated_RLVA_check = 1 then Call write_variable_in_case_note("* Updated RLVA.")
  	If updated_FACI_check = 1 then Call write_variable_in_case_note("* Updated FACI.")
End if
If need_3543_check = 1 then Call write_variable_in_case_note("* A 3543 is needed for the client.")
If need_3531_check = 1 then call write_variable_in_CASE_NOTE("* A 3531 is needed for the client.")
If need_asset_assessment_check = 1 then call write_variable_in_CASE_NOTE("* An asset assessment is needed before a MA-LTC determination can be made.")
If sent_3050_check = 1 then call write_variable_in_CASE_NOTE("* Sent 3050 back to LTCF.")
If sent_5181_check = 1 then call write_variable_in_CASE_NOTE("* Sent DHS-5181 to Case Manager.")
Call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
If sent_verif_request_check = 1 then Call write_variable_in_case_note("* Sent verif request to " & sent_request_to)
If processed_1503_check = 1 then Call write_variable_in_case_note("* Completed & Returned 1503 to LTCF.")
If TIKL_check = 1 then Call write_variable_in_case_note("* TIKLed to recheck length of stay on " & TIKL_date & ".")
Call write_bullet_and_variable_in_case_note("Notes", notes)
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)

script_end_procedure("")
