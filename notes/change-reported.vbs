'Created by Tim DeLong from Stearns County.

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


'Initial Dialog Box
BeginDialog change_reported_dialog, 0, 0, 171, 105, "Change Reported"
  ButtonGroup ButtonPressed
    OkButton 5, 85, 50, 15
    CancelButton 115, 85, 50, 15
  EditBox 85, 5, 60, 15, MAXIS_case_number
  EditBox 85, 25, 30, 15, MAXIS_footer_month
  EditBox 125, 25, 30, 15, MAXIS_footer_year
  DropListBox 25, 65, 125, 15, "Select One"+chr(9)+"Baby Born"+chr(9)+"HHLD Comp Change", List1
  Text 30, 10, 50, 10, "Case number:"
  Text 15, 30, 65, 10, "Footer month/year: "
  Text 25, 50, 130, 10, "Please select the nature of the change."
EndDialog

BeginDialog HHLD_Comp_Change_Dialog, 0, 0, 291, 175, "Household Comp Change"
  Text 5, 15, 50, 10, "Case Number"
  EditBox 60, 10, 100, 15, MAXIS_case_number
  Text 5, 35, 80, 10, "Unit Member HH Change"
  EditBox 90, 30, 45, 15, HH_member
  Text 5, 55, 85, 10, "Date Reported/Addendum"
  EditBox 95, 50, 60, 15, date_reported
  Text 165, 55, 45, 10, "Effective Date"
  EditBox 215, 50, 70, 15, effective_date
  CheckBox 110, 70, 110, 10, "Check if the change is temporary.", temporary_change_checkbox
  Text 10, 90, 45, 10, "Action Taken"
  EditBox 60, 85, 225, 15, actions_taken
  Text 5, 110, 60, 10, "Additional Notes"
  EditBox 60, 105, 225, 15, additional_notes
  Text 10, 130, 45, 15, "Worker Name"
  EditBox 60, 125, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 15, 150, 50, 15
    CancelButton 230, 150, 50, 15
EndDialog

'Connecting to BlueZone
EMConnect ""

'Finds the case number
Call MAXIS_case_number_finder(MAXIS_case_number)

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
		IF List1 = "Select One" THEN err_msg = err_msg & vbCr & "* Please select the type of change reported."
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
BeginDialog baby_born_dialog, 0, 0, 221, 350, "BABY BORN"
  EditBox 55, 5, 100, 15, MAXIS_case_number
  EditBox 55, 25, 100, 15, babys_name
  EditBox 55, 45, 100, 15, date_of_birth
  DropListBox 55, 65, 100, 15, "Select One"+chr(9)+"Male"+chr(9)+"Female", baby_gender
  DropListBox 85, 85, 70, 15, "Select One"+chr(9)+"Yes"+chr(9)+"No", father_in_household
  EditBox 70, 105, 85, 15, fathers_name
  EditBox 70, 130, 85, 15, fathers_employer
  DropListBox 70, 155, 130, 15, "Select One" & (HH_member_array_dialog), mothers_name
  EditBox 70, 180, 85, 15, mothers_employer
  DropListBox 80, 205, 70, 15, "Select One"+chr(9)+"Yes"+chr(9)+"No", other_health_insurance
  EditBox 115, 230, 80, 15, OHI_source
  EditBox 60, 255, 105, 15, other_notes
  EditBox 60, 275, 105, 15, actions_taken
  CheckBox 20, 295, 165, 10, "Newborns MHC plan updated to mothers carrier.", MHC_plan_checkbox
  EditBox 155, 310, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 5, 330, 50, 15
    CancelButton 165, 330, 50, 15
  Text 5, 235, 110, 15, "If yes to OHI, source of the OHI:"
  Text 10, 260, 45, 15, "Other Notes:"
  Text 5, 280, 50, 15, "Actions Taken:"
  Text 90, 315, 65, 15, "Worker Signature:"
  Text 5, 45, 45, 15, "Date of Birth:"
  Text 20, 105, 50, 10, "Fathers Name:"
  Text 5, 5, 50, 15, "Case Number: "
  Text 5, 25, 50, 15, "Baby's Name:"
  Text 5, 130, 65, 15, "Father's Employer:"
  Text 5, 85, 75, 15, "Father In Household?"
  Text 20, 65, 25, 10, "Gender:"
  Text 55, 205, 20, 10, "OHI?"
  Text 5, 155, 65, 10, "Mother of Newborn: "
  Text 5, 180, 65, 10, "Mother's Employer: "
EndDialog


IF List1 = "Baby Born" THEN

'Do loop for Baby Born Dialogbox
DO
	DO
		err_msg = ""
		DIALOG Baby_Born_Dialog
		cancel_confirmation
		IF mothers_name = "Select One" THEN err_msg = err_msg & vbNewLine & "You must choose newborn's mother"
		IF MAXIS_case_number = "" THEN err_msg = "You must enter case number!"
		IF babys_name = "" THEN err_msg = err_msg & vbNewLine &  "You must enter the babys name"
		IF date_of_birth = "" THEN err_msg = err_msg & vbNewLine &  "You must enter a birth date"
		If father_in_household = "Select One" then err_msg = err_msg & vbNewLine &  "You must answer 'Yes' or 'No' if father is listed in the household."
		If father_in_household = "Yes" and fathers_name = "" then err_msg = err_msg & vbNewLine &  "You must enter Father's name, since he is listed in household."
		'IF fathers_name = "" THEN err_msg = err_msg & vbNewLine &  "You must enter Father's name"
		IF actions_taken = "" THEN err_msg = err_msg & vbNewLine & "You must enter the actions taken"
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "Please sign your note"
		IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
	Loop Until err_msg = ""

	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

END IF

IF List1 = "HHLD Comp Change" THEN

'Do loop for HHLD Comp Change Dialogbox
DO
	DO
		err_msg = ""
		DIALOG HHLD_Comp_Change_Dialog
		cancel_confirmation
		IF MAXIS_case_number = "" THEN err_msg = "You must enter case number!"
		IF HH_Member = "" THEN err_msg = err_msg & vbNewLine & "You must enter a HH Member"
		IF date_reported = "" THEN err_msg = err_msg & vbNewLine & "You must enter date reported"
		IF effective_date = "" THEN err_msg = err_msg & vbNewLine & "You must enter effective date"
		IF actions_taken = "" THEN err_msg = err_msg & vbNewLine & "You must enter the actions taken"
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "Please sign your note"
		IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false
END IF


'Checks MAXIS for password prompt
CALL check_for_MAXIS(false)

'Navigates to case note
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

'Send PF9 to case note
PF9


'writes case note for Baby Born
IF List1 = "Baby Born" THEN

	CALL write_variable_in_Case_Note("--Client reports birth of baby--")
	CALL write_bullet_and_variable_in_Case_Note("Baby's name", babys_name)
	If baby_gender = "Select One" then									'gender will be listed as unknown if not updated'
		CALL write_bullet_and_variable_in_Case_Note("Gender", "unknown")
	Else
		CALL write_bullet_and_variable_in_Case_Note("Gender", baby_gender)
	End If
	CALL write_bullet_and_variable_in_Case_Note("Date of birth", date_of_birth)
	father_HH = " - not reported in the same household"
	If father_in_household = "Yes" Then father_HH = " - reported in the same household."
	If fathers_name = "" then fathers_name = "unknown"
	CALL write_bullet_and_variable_in_Case_Note("Father's name", fathers_name & father_HH)
	CALL write_bullet_and_variable_in_Case_Note("Father's employer", fathers_employer)
	CALL write_bullet_and_variable_in_Case_Note("Mother's name", mothers_name)
	CALL write_bullet_and_variable_in_Case_Note("Mother's employer", mothers_employer)
	IF other_health_insurance = "Yes" THEN CALL write_bullet_and_variable_in_Case_Note("OHI", OHI_Source)
	IF MHC_plan_checkbox = 1 THEN CALL write_variable_in_CASE_NOTE("* Newborns MHC plan updated to match the mothers.")
	CALL write_bullet_and_variable_in_Case_Note("Other Notes", other_notes)
	CALL write_bullet_and_variable_in_Case_Note("Actions Taken", actions_taken)
	CALL write_bullet_and_variable_in_Case_Note("Additional Notes", additional_notes)
END IF

'writes case note for HHLD Comp Change
IF List1 = "HHLD Comp Change" THEN

	CALL write_variable_in_case_note("HH Comp Change Reported")
	CALL write_bullet_and_variable_in_Case_Note("Unit member HH Member", HH_Member)
	CALL write_bullet_and_variable_in_Case_Note("Date Reported/Addendum", date_reported)
	CALL write_bullet_and_variable_in_Case_Note("Date Effective", effective_date)
	CALL write_bullet_and_variable_in_Case_Note("Actions Taken", actions_taken)
	CALL write_bullet_and_variable_in_Case_Note("Additional Notes", additional_notes)

	'case notes if the change is temporary
	IF Temporary_Change_Checkbox = 1 THEN CALL write_variable_in_Case_Note("***Change is temporary***")
	IF Temporary_Change_Checkbox = 0 THEN CALL write_variable_in_Case_Note("***Change is NOT temporary***")

END IF

'signs case note
CALL write_variable_in_Case_Note("----")
CALL write_variable_in_Case_Note(worker_signature)

script_end_procedure ("")
