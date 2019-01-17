'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CHANGE REPORTED.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds - INCLUDES A POLICY LOOKUP
STATS_denomination = "C"       		'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("01/16/2019", "Updated dialog boxes to prepare for enhancements to script.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Connecting to BlueZone
EMConnect ""

'Finds the case number
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'Initial Dialog Box
BeginDialog change_reported_dialog, 0, 0, 136, 105, "Change Reported"
  EditBox 70, 5, 35, 15, MAXIS_case_number
  EditBox 70, 25, 15, 15, MAXIS_footer_month
  EditBox 90, 25, 15, 15, MAXIS_footer_year
  DropListBox 20, 65, 85, 15, "Select One:"+chr(9)+"Address "+chr(9)+"Baby Born"+chr(9)+"HHLD Comp"+chr(9)+"Income "+chr(9)+"Shelter Cost "+chr(9)+"Other(please specify)", nature_change
  ButtonGroup ButtonPressed
    OkButton 45, 85, 40, 15
    CancelButton 90, 85, 40, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 65, 10, "Footer month/year: "
  Text 5, 50, 130, 10, "Please select the nature of the change:"
EndDialog

BeginDialog HHLD_Comp_Change_Dialog, 0, 0, 161, 200, "Household Comp Change"
  EditBox 80, 5, 20, 15, HH_member
  EditBox 80, 25, 35, 15, date_reported
  EditBox 80, 45, 35, 15, effective_date
  CheckBox 15, 75, 90, 10, "Verifications sent to ECF", Verif_checkbox
  CheckBox 15, 85, 80, 10, "Updated STAT panels", STAT_checkbox
  CheckBox 15, 95, 80, 10, "Approved new results", APP_checkbox
  CheckBox 15, 105, 80, 10, "Notified other agency", notify_checkbox
  EditBox 50, 125, 100, 15, additional_notes
  EditBox 50, 145, 100, 15, worker_signature
  CheckBox 5, 165, 125, 10, "Check if the change is temporary", temporary_change_checkbox
  ButtonGroup ButtonPressed
    OkButton 65, 180, 40, 15
    CancelButton 110, 180, 40, 15
  Text 5, 10, 75, 10, "Member # HH change:"
  Text 30, 50, 50, 10, "Effective date:"
  Text 5, 130, 45, 10, "Other Notes:"
  GroupBox 5, 65, 145, 55, "Action Taken"
  Text 30, 30, 50, 10, "Date reported:"
  Text 5, 150, 40, 10, "Worker Sig:"
EndDialog

BeginDialog change_received_dialog, 0, 0, 376, 280, "Change Report Form Received"
  EditBox 60, 5, 40, 15, MAXIS_case_number
  EditBox 160, 5, 45, 15, date_received
  EditBox 320, 5, 45, 15, effective_date
  EditBox 50, 35, 315, 15, address_notes
  EditBox 50, 55, 315, 15, household_notes
  EditBox 115, 75, 250, 15, asset_notes
  EditBox 50, 95, 315, 15, vehicles_notes
  EditBox 50, 115, 315, 15, income_notes
  EditBox 50, 135, 315, 15, shelter_notes
  EditBox 50, 155, 315, 15, other_change_notes
  EditBox 60, 180, 305, 15, actions_taken
  EditBox 60, 200, 305, 15, other_notes
  EditBox 70, 220, 295, 15, verifs_requested
  DropListBox 270, 240, 95, 20, "Select One:"+chr(9)+"will continue next month"+chr(9)+"will not continue next month", changes_continue
  CheckBox 10, 245, 140, 10, "Check here to navigate to DAIL/WRIT", tikl_nav_check
  EditBox 75, 260, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 260, 260, 50, 15
    CancelButton 315, 260, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 110, 10, 50, 10, "Effective Date:"
  Text 210, 10, 110, 10, "Date Change Reported/Received:"
  GroupBox 5, 25, 365, 150, "Changes Reported:"
  Text 15, 40, 30, 10, "Address:"
  Text 15, 60, 35, 10, "HH Comp:"
  Text 15, 80, 95, 10, "Assets (savings or property):"
  Text 15, 100, 30, 10, "Vehicles:"
  Text 15, 120, 30, 10, "Income:"
  Text 15, 140, 25, 10, "Shelter:"
  Text 15, 160, 20, 10, "Other:"
  Text 10, 185, 45, 10, "Action Taken:"
  Text 10, 205, 45, 10, "Other Notes:"
  Text 10, 225, 60, 10, "Verifs Requested:"
  Text 10, 265, 60, 10, "Worker Signature:"
  Text 180, 245, 90, 10, "The changes client reports:"
EndDialog

'Finds the benefit month
EMReadScreen on_SELF, 4, 2, 50
IF on_SELF = "SELF" THEN
	CALL find_variable("Benefit Period (MM YY): ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" THEN CALL find_variable("Benefit Period (MM YY): " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
ELSE
	CALL find_variable("Month: ", MAXIS_footer_month, 2)
	IF MAXIS_footer_month <> "" THEN CALL find_variable("Month: " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
END IF


'Info to the user of what this script currently covers
MsgBox "This script currently only covers if there is a HHLD Comp Change or a Baby Born. Other reported changes will be covered here in the future."

check_for_maxis(False)

DO
	err_msg = ""
	DIALOG change_reported_dialog
		IF ButtonPressed = 0 THEN stopscript
		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
		IF nature_change = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select the type of change reported."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'this creates the client array for baby_born_dialog dropdown list
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 2, 4, 33
	EMReadscreen last_name_array, 25, 6, 30								'took out clients last name apparently may be too much characters within the form restrictions.
	EMReadscreen first_name_array, 12, 6, 63
	last_name_array = replace(last_name_array, "_", "")
	last_name_array = Lcase(last_name_array)
	last_name_array = UCase(Left(last_name_array, 1)) &  Mid(last_name_array, 2)     	'took out clients last name apparently may be too much characters within the form restrictions.
	first_name_array = replace(first_name_array, "_", "") '& " "
	first_name_array = Lcase(first_name_array)
	first_name_array = UCase(Left(first_name_array, 1)) &  Mid(first_name_array, 2)
	client_string =  "MEMB " & ref_nbr & " - " & first_name_array & " " & last_name_array
	client_array = client_array & client_string & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array
DIM all_client_array()
ReDim all_clients_array(total_clients, 1)
FOR clt_x = 0 to total_clients				'using a dummy array to build list into the array used for the dialog.
	Interim_array = split(client_array, "|")
	all_clients_array(clt_x, 0) = Interim_array(clt_x)
	all_clients_array(clt_x, 1) = 1
NEXT
HH_member_array = ""
FOR i = 0 to total_clients
	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
		IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
			HH_member_array = chr(9) & HH_member_array & chr(9) & all_clients_array(i, 0)
		END IF
	END IF
NEXT
'removes all of the first 'chr(9)'
HH_member_array_dialog = Right(HH_member_array, len(HH_member_array) - total_clients)

'Baby_born Dialog needs to begin here to accept 'HH_member_array_dialog into dropdown list: mothers_name
BeginDialog baby_born_dialog, 0, 0, 186, 265, "BABY BORN"
  EditBox 55, 5, 115, 15, babys_name
  EditBox 55, 25, 40, 15, date_of_birth
  DropListBox 130, 25, 40, 15, "Select One:"+chr(9)+"Male"+chr(9)+"Female", baby_gender
  DropListBox 100, 45, 40, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", parent_in_household
  DropListBox 85, 75, 80, 15, "Select One:" & (HH_member_array_dialog), mothers_name
  EditBox 85, 95, 80, 15, mothers_employer
  EditBox 80, 130, 85, 15, fathers_name
  EditBox 80, 150, 85, 15, fathers_employer
  CheckBox 10, 170, 165, 10, "Newborns MHC plan updated to mother's carrier", MHC_plan_checkbox
  DropListBox 140, 185, 40, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", other_health_insurance
  EditBox 110, 205, 70, 15, OHI_source
  EditBox 50, 225, 130, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 95, 245, 40, 15
    CancelButton 140, 245, 40, 15
  Text 5, 30, 45, 10, "Date of birth:"
  Text 100, 30, 25, 10, "Gender:"
  Text 5, 50, 95, 10, "Other parent in household?"
  Text 15, 135, 50, 10, "Fathers Name:"
  Text 5, 10, 50, 10, "Child's name:"
  Text 15, 155, 65, 10, "Father's Employer:"
  Text 5, 230, 45, 10, "Other Notes:"
  Text 5, 210, 105, 10, "If yes to OHI, source of the OHI:"
  Text 55, 190, 80, 10, "Other Health Insurance?"
  Text 15, 80, 65, 10, "Mother of Newborn: "
  Text 15, 100, 65, 10, "Mother's Employer: "
  GroupBox 5, 120, 175, 50, "Father's Information"
  GroupBox 5, 65, 175, 50, "Mother's Information"
EndDialog

BeginDialog HH_memb_dialog, 0, 0, 371, 220, "HH Comp change"
  EditBox 55, 5, 20, 15, HH_member
  EditBox 135, 5, 40, 15, date_of_birth
  EditBox 55, 25, 120, 15, babys_name
  EditBox 55, 45, 35, 15, date_reported
  DropListBox 135, 45, 40, 15, "Select One:"+chr(9)+"Male"+chr(9)+"Female", baby_gender
  EditBox 55, 65, 35, 15, effective_date
  CheckBox 5, 85, 125, 10, "Check if the change is temporary", temporary_change_checkbox
  CheckBox 15, 110, 90, 10, "Verifications sent to ECF", Verif_checkbox
  CheckBox 15, 120, 80, 10, "Updated STAT panels", STAT_checkbox
  CheckBox 15, 130, 80, 10, "Approved new results", APP_checkbox
  CheckBox 15, 140, 160, 10, "Notified other agency(please advise of name)", notify_checkbox
  EditBox 50, 160, 125, 15, other_notes
  EditBox 50, 180, 125, 15, worker_signature
  DropListBox 320, 5, 40, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", parent_in_household
  DropListBox 255, 35, 105, 15, "Select One: & (HH_member_array_dialog)", mothers_name
  EditBox 270, 55, 90, 15, mothers_employer
  EditBox 250, 85, 110, 15, fathers_name
  EditBox 265, 105, 95, 15, fathers_employer
  CheckBox 195, 140, 165, 10, "Newborns MHC plan updated to mother's carrier", MHC_plan_checkbox
  DropListBox 320, 155, 40, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", other_health_insurance
  EditBox 290, 175, 70, 15, OHI_source
  ButtonGroup ButtonPressed
    OkButton 280, 200, 40, 15
    CancelButton 325, 200, 40, 15
  Text 200, 40, 55, 10, "Mother's name: "
  Text 200, 60, 65, 10, "Mother's employer: "
  GroupBox 190, 75, 175, 50, "Father's Information"
  GroupBox 190, 25, 175, 50, "Mother's Information"
  Text 110, 10, 20, 10, "DOB:"
  Text 235, 160, 80, 10, "Other Health Insurance?"
  Text 105, 50, 25, 10, "Gender:"
  Text 200, 90, 50, 10, "Fathers name:"
  Text 5, 30, 45, 10, "MEMB name:"
  Text 200, 110, 65, 10, "Father's employer:"
  Text 5, 165, 45, 10, "Other Notes:"
  Text 205, 180, 80, 10, "If yes, source of the OHI:"
  Text 225, 10, 95, 10, "Other parent in household?"
  Text 5, 10, 35, 10, "Member #:"
  Text 5, 70, 50, 10, "Effective date:"
  GroupBox 5, 100, 170, 55, "Action Taken"
  Text 5, 50, 50, 10, "Date reported:"
  Text 5, 185, 40, 10, "Worker Sig:"
  GroupBox 190, 130, 175, 65, "Health Care Information"
EndDialog

IF nature_change <> "Baby Born" or nature_change <>"HHLD Comp Change"  THEN
    'Shows dialog
    DO
    	DO
    		DO
    			DO
    				Dialog crf_received_dialog
    				cancel_confirmation
    				IF worker_signature = "" THEN MsgBox "You must sign your case note!"
    			LOOP UNTIL worker_signature <> ""
    			IF IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must type a valid numeric case number."
    		LOOP UNTIL IsNumeric(MAXIS_case_number) = TRUE
    		IF changes_continue = "Select One:" THEN MsgBox "You Must Select 'The changes client reports field'"
    	LOOP UNTIL changes_continue <> "Select One:"
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false


    'Checks Maxis for password prompt
    CALL check_for_MAXIS(FALSE)

    'THE CASENOTE----------------------------------------------------------------------------------------------------
    'Navigates to case note
    Call start_a_blank_case_note
    CALL write_variable_in_case_note ("--CHANGE REPORTED--")
    CALL write_bullet_and_variable_in_case_note("Date Received", date_received)
    CALL write_bullet_and_variable_in_case_note("Date Effective", effective_date)
    CALL write_bullet_and_variable_in_case_note("Address", address_notes)
    CALL write_bullet_and_variable_in_case_note("Household Members", household_notes)
    CALL write_bullet_and_variable_in_case_note("Assets", asset_notes)
    CALL write_bullet_and_variable_in_case_note("Vehicles", vehicles_notes)
    CALL write_bullet_and_variable_in_case_note("Income", income_notes)
    CALL write_bullet_and_variable_in_case_note("Shelter", shelter_notes)
    CALL write_bullet_and_variable_in_case_note("Other", other_change_notes)
    CALL write_bullet_and_variable_in_case_note("Action Taken", actions_taken)
    CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
    CALL write_bullet_and_variable_in_case_note("Verifs Requested", verifs_requested)
    IF changes_continue <> "Select One:" THEN CALL write_bullet_and_variable_in_case_note("The changes client reports", changes_continue)
    CALL write_variable_in_case_note("---")
    CALL write_variable_in_case_note(worker_signature)

    'If we checked to TIKL out, it goes to TIKL and sends a TIKL
    IF tikl_nav_check = 1 THEN
    	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
    	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
    	EMSetCursor 9, 3
    END IF
END IF

IF nature_change = "Baby Born" THEN

'Do loop for Baby Born Dialogbox
DO
	DO
		err_msg = ""
		DIALOG Baby_Born_Dialog
		cancel_confirmation
		IF mothers_name = "Select One:" THEN err_msg = err_msg & vbNewLine & "You must choose newborn's mother"
		IF babys_name = "" THEN err_msg = err_msg & vbNewLine &  "You must enter the babys name"
		IF date_of_birth = "" THEN err_msg = err_msg & vbNewLine &  "You must enter a birth date"
		If parent_in_household = "Select One:" then err_msg = err_msg & vbNewLine &  "You must answer 'Yes' or 'No' if father is listed in the household."
		If parent_in_household = "Yes" and fathers_name = "" then err_msg = err_msg & vbNewLine &  "You must enter Father's name, since he is listed in household."
		IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
	Loop Until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

END IF

IF nature_change = "HHLD Comp Change" THEN

'Do loop for HHLD Comp Change Dialogbox
DO
	DO
		err_msg = ""
		DIALOG HHLD_Comp_Change_Dialog
		cancel_confirmation
		IF HH_Member = "" THEN err_msg = err_msg & vbNewLine & "You must enter a HH Member"
		IF date_reported = "" THEN err_msg = err_msg & vbNewLine & "You must enter date reported"
		IF effective_date = "" THEN err_msg = err_msg & vbNewLine & "You must enter effective date"
		IF notify_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please advise in other notes the name of the agencyt notified"
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "Please sign your note"
		IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false
END IF


'Checks MAXIS for password prompt
CALL check_for_MAXIS(false)

actions_taken = ""
IF Verif_checkbox = CHECKED THEN actions_taken = actions_taken & "Verifications sent to ECF,"
IF STAT_checkbox = CHECKED THEN actions_taken = actions_taken & "Updated STAT panels,"
IF APP_checkbox = CHECKED THEN actions_taken = actions_taken & "Approved new results,"
IF notify_checkbox = CHECKED THEN actions_taken = actions_taken & "Notified other agency,"

start_a_blank_case_note
'writes case note for Baby Born
IF nature_change = "Baby Born" THEN
	CALL write_variable_in_Case_Note("--CHANGE REPORTED - Client reports birth of baby--")
	CALL write_bullet_and_variable_in_Case_Note("Child's's name", babys_name)
	If baby_gender = "Select One:" then									'gender will be listed as unknown if not updated'
		CALL write_bullet_and_variable_in_Case_Note("Gender", "unknown")
	Else
		CALL write_bullet_and_variable_in_Case_Note("Gender", baby_gender)
	End If
	CALL write_bullet_and_variable_in_Case_Note("Date of birth", date_of_birth)
	father_HH = " - not reported in the same household"
	If parent_in_household = "Yes" Then father_HH = " - reported in the same household."
	If fathers_name = "" then fathers_name = "Unknown or not provided"
	CALL write_bullet_and_variable_in_Case_Note("Mother's name", mothers_name)
	CALL write_bullet_and_variable_in_Case_Note("Mother's employer", mothers_employer)
	CALL write_bullet_and_variable_in_Case_Note("Father's name", fathers_name & father_HH)
	CALL write_bullet_and_variable_in_Case_Note("Father's employer", fathers_employer)
	IF other_health_insurance = "Yes" THEN CALL write_bullet_and_variable_in_Case_Note("OHI", OHI_Source)
	IF MHC_plan_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE("* Newborns MHC plan updated to match the mothers.")
	CALL write_bullet_and_variable_in_Case_Note("Other Notes", other_notes)
END IF

'writes case note for HHLD Comp Change
IF nature_change = "HHLD Comp Change" THEN
	CALL write_variable_in_case_note("--CHANGE REPORTED - HH Comp Change--")
	CALL write_bullet_and_variable_in_Case_Note("Unit member HH Member", HH_Member)
	CALL write_bullet_and_variable_in_Case_Note("Date Reported/Addendum", date_reported)
	CALL write_bullet_and_variable_in_Case_Note("Date Effective", effective_date)
	CALL write_bullet_and_variable_in_Case_Note("Actions Taken", actions_taken)
	CALL write_bullet_and_variable_in_Case_Note("Additional Notes", other_notes)
	'case notes if the change is temporary
	IF Temporary_Change_Checkbox = CHECKED THEN CALL write_variable_in_Case_Note("***Change is temporary***")
	IF Temporary_Change_Checkbox = UNCHECKED THEN CALL write_variable_in_Case_Note("***Change is NOT temporary***")
END IF

'signs case note
CALL write_variable_in_Case_Note("----")
CALL write_variable_in_Case_Note(worker_signature)

script_end_procedure ("The case note has been created please be sure to send verifications to ECF or case note how the information was received.")
