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
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("04/23/2019", "Updated other notes field to update to case note. Moved initial informational message box to intial dialog box.", "Ilse Ferris, Hennepin County")
call changelog_update("01/16/2019", "Updated dialog boxes to prepare for enhancements to script.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Connecting to BlueZone
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number) 'Finds the case number
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 161, 150, "Change Reported"
  EditBox 90, 55, 35, 15, MAXIS_case_number
  EditBox 90, 75, 15, 15, MAXIS_footer_month
  EditBox 110, 75, 15, 15, MAXIS_footer_year
  DropListBox 20, 110, 125, 15, "Select One:"+chr(9)+"Baby Born"+chr(9)+"HH Comp Change", nature_change
  ButtonGroup ButtonPressed
    OkButton 60, 130, 40, 15
    CancelButton 105, 130, 40, 15
  Text 30, 60, 50, 10, "Case number:"
  Text 20, 80, 65, 10, "Footer month/year: "
  Text 20, 95, 130, 10, "Please select the nature of the change."
  Text 10, 20, 140, 25, "This script currently only covers HH Comp Changes or a Baby Born. Other reported changes will be added  in the future."
  GroupBox 5, 5, 145, 45, "Change Reported script information"
EndDialog

'Info to the user of what this script currently covers
Do
    Do
	    err_msg = ""
	    DIALOG Dialog1
	    Cancel_without_confirmation
	    IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
	    IF nature_change = "Select One:" THEN err_msg = err_msg & vbCr & "* Please select the type of change reported."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

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

IF nature_change = "Baby Born" THEN
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
    'Baby_born Dialog needs to begin here to accept 'HH_member_array_dialog into dropdown list: mothers_name
    BeginDialog Dialog1, 0, 0, 186, 265, "BABY BORN"
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
'Do loop for Baby Born Dialogbox
    DO
    	DO
		    err_msg = ""
    		DIALOG Dialog1
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

IF nature_change = "HH Comp Change" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 161, 200, "Household Comp Change"
      EditBox 80, 5, 20, 15, HH_member
      EditBox 80, 25, 50, 15, date_reported
      EditBox 80, 45, 50, 15, effective_date
      CheckBox 15, 75, 90, 10, "Verifications sent to ECF", Verif_checkbox
      CheckBox 15, 85, 80, 10, "Updated STAT panels", STAT_checkbox
      CheckBox 15, 95, 80, 10, "Approved new results", APP_checkbox
      CheckBox 15, 105, 80, 10, "Notified other agency", notify_checkbox
      EditBox 50, 125, 100, 15, other_notes
      EditBox 50, 145, 100, 15, worker_signature
      CheckBox 5, 165, 125, 10, "Check if the change is temporary", temporary_change_checkbox
      ButtonGroup ButtonPressed
    	OkButton 65, 180, 40, 15
    	CancelButton 110, 180, 40, 15
      Text 5, 10, 75, 10, "Member # HH change:"
      Text 25, 50, 50, 10, "Effective date:"
      Text 5, 130, 45, 10, "Other Notes:"
      GroupBox 5, 65, 145, 55, "Action Taken"
      Text 25, 30, 50, 10, "Date reported:"
      Text 5, 150, 40, 10, "Worker Sig:"
    EndDialog
    'Do loop for HH Comp Change Dialogbox
	DO
    	DO
		    err_msg = ""
    		DIALOG Dialog1
    		cancel_confirmation
    		IF HH_Member = "" THEN err_msg = err_msg & vbNewLine & "You must enter a HH Member"
    		IF date_reported = "" THEN err_msg = err_msg & vbNewLine & "You must enter date reported"
    		IF effective_date = "" THEN err_msg = err_msg & vbNewLine & "You must enter effective date"
    		IF notify_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Enter the name of the agency notified in the 'other notes' section."
    		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "Please sign your note"
    		IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
    	LOOP UNTIL err_msg = ""
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false
END IF

actions_taken = ""
IF Verif_checkbox = CHECKED THEN actions_taken = actions_taken & "Verifications sent to resident,"
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

'writes case note for HH Comp Change
IF nature_change = "HH Comp Change" THEN
	CALL write_variable_in_case_note("--CHANGE REPORTED - HH Comp Change--")
	CALL write_bullet_and_variable_in_Case_Note("Unit member HH Member", HH_Member)
	CALL write_bullet_and_variable_in_Case_Note("Date Reported/Addendum", date_reported)
	CALL write_bullet_and_variable_in_Case_Note("Date Effective", effective_date)
	CALL write_bullet_and_variable_in_Case_Note("Actions Taken", actions_taken)
	CALL write_bullet_and_variable_in_Case_Note("Other notes", other_notes)
	'case notes if the change is temporary
	IF Temporary_Change_Checkbox = CHECKED THEN CALL write_variable_in_Case_Note("***Change is temporary***")
	IF Temporary_Change_Checkbox = UNCHECKED THEN CALL write_variable_in_Case_Note("***Change is NOT temporary***")
END IF

'signs case note
CALL write_variable_in_Case_Note("----")
CALL write_variable_in_Case_Note(worker_signature)

script_end_procedure ("The case note has been created please be sure to send verifications to ECF or case note how the information was received.")
