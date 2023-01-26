'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - SHELTER-CHANGE REPORTED.vbs"
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
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the resident instead.", "Ilse Ferris, Hennepin County")
call changelog_update("03/07/2022", "Updated Team 601 contact emails to be 603 per De Vang's request.", "Ilse Ferris, Hennepin County")
call changelog_update("03/13/2020", "Updated TIKL Functionality and updates to script to pull a list of contacts for emails outside of the script.", "Ilse Ferris, Hennepin County")
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

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 231, 65, "Change Reported"
  EditBox 65, 5, 60, 15, MAXIS_case_number
  EditBox 185, 5, 15, 15, MAXIS_footer_month
  EditBox 205, 5, 15, 15, MAXIS_footer_year
  DropListBox 130, 25, 90, 15, "Select:"+chr(9)+"Address"+chr(9)+"Baby Born"+chr(9)+"HHLD Comp"+chr(9)+"Income"+chr(9)+"Shelter Cost"+chr(9)+"Other(please specify)", nature_change
  EditBox 65, 45, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 130, 45, 45, 15
    CancelButton 175, 45, 45, 15
  Text 130, 10, 50, 10, "Footer MM/YY: "
  Text 5, 30, 105, 10, "Select the nature of the change:"
  Text 5, 50, 60, 10, "Worker signature:"
  Text 5, 10, 50, 10, "Case number:"
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

DO
	DO
		err_msg = ""
		DIALOG Dialog1
		cancel_without_confirmation
		IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND len(MAXIS_case_number) > 8) OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbCr & "Please enter a valid case number."
        IF len(MAXIS_footer_month) > 2 or isnumeric(MAXIS_footer_month) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit initial month."
		IF len(MAXIS_footer_year) > 2 or isnumeric(MAXIS_footer_year) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit initial year."
        IF nature_change = "Select:" THEN err_msg = err_msg & vbCr & "* Please select the type of change reported."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	Loop Until err_msg = ""
call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'this creates the client array for baby_born_dialog dropdown list
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen worker_number, 7, 21, 21
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

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'Baby_born Dialog needs to begin here to accept 'HH_member_array_dialog into dropdown list: mothers_name
BeginDialog Dialog1, 0, 0, 371, 220, "Baby born or HH Comp change"
  EditBox 55, 5, 20, 15, HH_member
  EditBox 135, 5, 40, 15, date_of_birth
  EditBox 55, 25, 120, 15, babys_name
  EditBox 55, 45, 35, 15, date_received
  DropListBox 135, 45, 40, 15, "Select:"+chr(9)+"Male"+chr(9)+"Female", baby_gender
  EditBox 55, 65, 35, 15, effective_date
  CheckBox 5, 85, 125, 10, "Check if the change is temporary", temporary_change_checkbox
  CheckBox 15, 110, 90, 10, "Verifications sent to ECF", Verif_checkbox
  CheckBox 15, 120, 80, 10, "Updated STAT panels", STAT_checkbox
  CheckBox 15, 130, 80, 10, "Approved new results", APP_checkbox
  CheckBox 15, 140, 160, 10, "Notified other agency(please advise of name)", notify_checkbox
  EditBox 50, 160, 125, 15, other_notes
  EditBox 50, 180, 125, 15, worker_signature
  DropListBox 320, 5, 40, 15, "Select:"+chr(9)+"Yes"+chr(9)+"No", parent_in_household
  DropListBox 255, 35, 105, 15, "Select:" & (HH_member_array_dialog), mothers_name
  EditBox 270, 55, 90, 15, mothers_employer
  EditBox 250, 85, 110, 15, fathers_name
  EditBox 265, 105, 95, 15, fathers_employer
  CheckBox 195, 140, 165, 10, "Newborns MHC plan updated to mother's carrier", MHC_plan_checkbox
  DropListBox 320, 155, 40, 15, "Select:"+chr(9)+"Yes"+chr(9)+"No", other_health_insurance
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

'Checks MAXIS for password prompt
CALL check_for_MAXIS(false)

actions_taken = ""
IF Verif_checkbox = CHECKED THEN actions_taken = actions_taken & "Verifications sent to resident,"
IF STAT_checkbox = CHECKED THEN actions_taken = actions_taken & "Updated STAT panels,"
IF APP_checkbox = CHECKED THEN actions_taken = actions_taken & "Approved new results,"
IF notify_checkbox = CHECKED THEN actions_taken = actions_taken & "Notified other agency,"

IF nature_change = "Baby Born" or nature_change = "HHLD Comp" THEN
    DO
    	DO
    		err_msg = ""
    		DIALOG Dialog1
    		cancel_without_confirmation
    		IF babys_name = "" THEN err_msg = err_msg & vbNewLine &  "You must enter the new HH member name"
			IF date_received = "" THEN err_msg = err_msg & vbNewLine & "You must enter date reported"
			IF effective_date = "" THEN err_msg = err_msg & vbNewLine & "You must enter effective date"
    		IF date_of_birth = "" THEN err_msg = err_msg & vbNewLine &  "You must enter a birth date"
    		'If parent_in_household = "Select:" then err_msg = err_msg & vbNewLine &  "You must answer 'Yes' or 'No' if father is listed in the household."
    		If parent_in_household = "Yes" and fathers_name = "" then err_msg = err_msg & vbNewLine &  "You must enter Father's name, since he is listed in household."
			'IF mothers_name = "Select:" THEN err_msg = err_msg & vbNewLine & "You must choose newborn's mother"
			IF HH_Member = "" THEN err_msg = err_msg & vbNewLine & "You must enter a HH Member"
			IF notify_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please advise in other notes the name of the agency notified"
    		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
            IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
    	Loop Until err_msg = ""
    	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false
ELSE
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 376, 280, "Change Report Received"
        EditBox 60, 5, 40, 15, MAXIS_case_number
        EditBox 160, 5, 45, 15, effective_date
        EditBox 320, 5, 45, 15, date_received
        EditBox 50, 35, 315, 15, address_notes
        EditBox 50, 55, 315, 15, household_notes
        EditBox 115, 75, 250, 15, asset_notes
        EditBox 50, 95, 315, 15, vehicles_notes
        EditBox 50, 115, 315, 15, income_notes
        EditBox 50, 135, 315, 15, shelter_notes
        EditBox 50, 155, 315, 15, other_change_notes
        EditBox 70, 180, 295, 15, actions_taken
        EditBox 70, 200, 295, 15, other_notes
        EditBox 70, 220, 295, 15, verifs_requested
        CheckBox 5, 245, 125, 10, "Check if the change is temporary", temporary_change_checkbox
        CheckBox 5, 255, 130, 10, "Check here to navigate to set a TIKL", tikl_nav_check
        CheckBox 5, 265, 155, 10, "Create email to the team -  reporting change", send_email_checkbox
        EditBox 280, 240, 85, 15, worker_signature
        ButtonGroup ButtonPressed
    OkButton 260, 260, 50, 15
    CancelButton 315, 260, 50, 15
        Text 215, 245, 60, 10, "Worker Signature:"
        Text 15, 40, 30, 10, "Address:"
        Text 15, 60, 35, 10, "HH Comp:"
        GroupBox 5, 25, 365, 150, "Changes Reported:"
        Text 15, 80, 95, 10, "Assets (savings or property):"
        Text 15, 100, 30, 10, "Vehicles:"
        Text 15, 120, 30, 10, "Income:"
        Text 15, 140, 25, 10, "Shelter:"
        Text 15, 160, 20, 10, "Other:"
        Text 20, 185, 45, 10, "Action Taken:"
        Text 25, 205, 45, 10, "Other Notes:"
        Text 10, 225, 60, 10, "Verifs Requested:"
        Text 210, 10, 110, 10, "Date Change Reported/Received:"
        Text 110, 10, 50, 10, "Effective Date:"
        Text 5, 10, 50, 10, "Case Number:"
    EndDialog

	DO
        DO
            err_msg = ""
		    DIALOG Dialog1
		    cancel_without_confirmation
		    IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Please enter a valid case number."
            If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
            IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
        Loop Until err_msg = ""
        CALL check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
	LOOP UNTIL are_we_passworded_out = false
END IF

'IF nature_change = "Address"
'IF nature_change = "Baby Born"
'IF nature_change = "HHLD Comp"
'IF nature_change = "Income"
'IF nature_change = "Shelter Cost"
IF nature_change = "Other(please specify)" THEN nature_change = "Other"
IF memb_number = "" THEN memb_number = "01"

'create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If tikl_nav_check = 1 then Call create_TIKL("CHANGE REPORTED", 10, date, False, TIKL_note_text)

IF worker_number =	"X127F3P" 	THEN email_address = "HSPH.ES.MA.EPD.Adult@hennepin.us"
IF worker_number =	"X127F3K" 	THEN email_address = "HSPH.ES.MA.EPD.FAM@hennepin.us"
IF worker_number =	"X127F3F"	THEN email_address = "HSPH.ES.MA.EPD.ADS@hennepin.us"
IF worker_number =	"X127EA0" 	THEN email_address = "hsph.es.ea.team@hennepin.us"
IF worker_number =	"X127EAK" 	THEN email_address = "hsph.es.ea.team@hennepin.us"
IF worker_number =	"X127EM3" 	THEN email_address = "hsph.es.extendicare@hennepin.us"
IF worker_number =	"X127EM4" 	THEN email_address = "hsph.es.extendicare@hennepin.us"
IF worker_number =	"X127FG6" 	THEN email_address = "hsph.es.goldenliving@hennepin.us"
IF worker_number =	"X127FG7" 	THEN email_address = "hsph.es.goldenliving@hennepin.us"
IF worker_number =	"X127LE1" 	THEN email_address = "hsph.es.littleearth@hennepin.us"
IF worker_number =	"X127NP0"	THEN email_address = "hsph.es.northpoint@hennepin.us"
IF worker_number =	"X127NPC" 	THEN email_address = "hsph.es.northpoint@hennepin.us"
IF worker_number =	"X127FF4" 	THEN email_address = "hsph.es.northridge@hennepin.us"
IF worker_number =	"X127FF5" 	THEN email_address = "hsph.es.northridge@hennepin.us"
IF worker_number =	"X127ED8" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127EAJ" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127EN1" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127EN2" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127EN3" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127EN4"  	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127ED6" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127ED7" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127FE6" 	THEN email_address = "hsph.es.team.110@hennepin.us"
IF worker_number =	"X127EJ9" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF worker_number =	"X127ER6" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF worker_number =	"X127EE2" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF worker_number =	"X127EE3" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF worker_number =	"X127EE4"	THEN email_address = "hsph.es.team.120@hennepin.us"
IF worker_number =	"X127EE5" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF worker_number =	"X127EG5" 	THEN email_address = "hsph.es.team.120@hennepin.us"
IF worker_number =	"X127EQ1" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF worker_number =	"X127EQ2" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF worker_number =	"X127EQ5" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF worker_number =	"X127EK8" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF worker_number =	"X127EQ4" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF worker_number =	"X127FH9" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF worker_number =	"X127EG6" 	THEN email_address = "hsph.es.team.130@hennepin.us"
IF worker_number =	"X127EE1"	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127FB2"	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127EG7" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127ED9"	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127EE0" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127EH4" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127EH5" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127F3D" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127FH8" 	THEN email_address = "hsph.es.team.140@hennepin.us"
IF worker_number =	"X127EQ8" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF worker_number =	"X127EE6" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF worker_number =	"X127EE7" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF worker_number =	"X127ER1" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF worker_number =	"X127EG8" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF worker_number =	"X127FH2" 	THEN email_address = "hsph.es.team.150@hennepin.us"
IF worker_number =	"X127EF8" 	THEN email_address = "hsph.es.team.160@hennepin.us"
IF worker_number =	"X127EF9" 	THEN email_address = "hsph.es.team.160@hennepin.us"
IF worker_number =	"X127EG9" 	THEN email_address = "hsph.es.team.160@hennepin.us"
IF worker_number =	"X127EG0" 	THEN email_address = "hsph.es.team.160@hennepin.us"
IF worker_number =	"X127EP8" 	THEN email_address = "hsph.es.team.170@hennepin.us"
IF worker_number =	"X127EP6" 	THEN email_address = "hsph.es.team.170@hennepin.us"
IF worker_number =	"X127EP7"	THEN email_address = "hsph.es.team.170@hennepin.us"
IF worker_number =	"X127EG4"	THEN email_address = "hsph.es.team.170@hennepin.us"
IF worker_number =	"X127FG8" 	THEN email_address = "hsph.es.team.170@hennepin.us"
IF worker_number =	"X127EH1" 	THEN email_address = "hsph.es.team.251@hennepin.us"
IF worker_number =	"X127EH7" 	THEN email_address = "hsph.es.team.251@hennepin.us"
IF worker_number =	"X127EH2"  	THEN email_address = "hsph.es.team.251@hennepin.us"
IF worker_number =	"X127EH3" 	THEN email_address = "hsph.es.team.251@hennepin.us"
IF worker_number =	"X127FH4" 	THEN email_address = "hsph.es.team.251@hennepin.us"
IF worker_number =	"X127EH8" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127EQ3" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127EJ2" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127EJ3" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127FH1" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127FG4" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127F3C" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127F3G"	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127F3L" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127EJ1" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127EH9" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127EM2" 	THEN email_address = "hsph.es.team.252@hennepin.us"
IF worker_number =	"X127EJ6" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF worker_number =	"X127FE5" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF worker_number =	"X127EK3" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF worker_number =	"X127EK1" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF worker_number =	"X127EK2" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF worker_number =	"X127EJ7" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF worker_number =	"X127EJ8" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF worker_number =	"X127EJ5" 	THEN email_address = "hsph.es.team.253@hennepin.us"
IF worker_number =	"X127EL8" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127EL9" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127FE1" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127EL2" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127EL3" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127EL4" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127EL5" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127EL6" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127EL7" 	THEN email_address = "hsph.es.team.255@hennepin.us"
IF worker_number =	"X127EH6" 	THEN email_address = "hsph.es.team.256@hennepin.us"
IF worker_number =	"X127EM1" 	THEN email_address = "hsph.es.team.256@hennepin.us"
IF worker_number =	"X127FI7" 	THEN email_address = "hsph.es.team.256@hennepin.us"
IF worker_number =	"X127EM7" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF worker_number =	"X127FI2" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF worker_number =	"X127FG3" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF worker_number =	"X127EM8" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF worker_number =	"X127EM9" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF worker_number =	"X127EJ4" 	THEN email_address = "hsph.es.team.257@hennepin.us"
IF worker_number =	"X127EK4" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF worker_number =	"X127EK5" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF worker_number =	"X127FH5" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF worker_number =	"X127EN7"	THEN email_address = "hsph.es.team.258@hennepin.us"
IF worker_number =	"X127EK6" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF worker_number =	"X127EK9" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF worker_number =	"X127EN6" 	THEN email_address = "hsph.es.team.258@hennepin.us"
IF worker_number =	"X127EP3" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF worker_number =	"X127EP4" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF worker_number =	"X127EP5" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF worker_number =	"X127EP9" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF worker_number =	"X127F3U"	THEN email_address = "hsph.es.team.259@hennepin.us"
IF worker_number =	"X127F3V" 	THEN email_address = "hsph.es.team.259@hennepin.us"
IF worker_number =	"X127EF7" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF worker_number =	"X127EN5" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF worker_number =	"X127EF5" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF worker_number =	"X127EK7" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF worker_number =	"X127EF6" 	THEN email_address = "hsph.es.team.260@hennepin.us"
IF worker_number =	"X127EQ9" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF worker_number =	"X127ER2" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF worker_number =	"X127ER3" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF worker_number =	"X127ER4" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF worker_number =	"X127ER5" 	THEN email_address = "hsph.es.team.261@hennepin.us"
IF worker_number =	"X127FF6" 	THEN email_address = "hsph.es.team.262@hennepin.us"
IF worker_number =	"X127FF7" 	THEN email_address = "hsph.es.team.262@hennepin.us"
IF worker_number =	"X127FF8" 	THEN email_address = "hsph.es.team.300@hennepin.us"
IF worker_number =	"X127FF9" 	THEN email_address = "hsph.es.team.300@hennepin.us"
IF worker_number =	"X127FF3" 	THEN email_address = "hsph.es.team.410@hennepin.us"
IF worker_number =	"X127EX3" 	THEN email_address = "hsph.es.team.410@hennepin.us"
IF worker_number =	"X127ES7" 	THEN email_address = "hsph.es.team.410@hennepin.us"
IF worker_number =	"X127EX2" 	THEN email_address = "hsph.es.team.410@hennepin.us"
IF worker_number =	"X127EX1"	THEN email_address = "hsph.es.team.410@hennepin.us"
IF worker_number =	"X127ET9"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127EU4"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127EW2"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127EW3"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127EU1"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127EU3"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127BV2"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127EU2"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127FH7"	THEN email_address = "hsph.es.team.450@hennepin.us"
IF worker_number =	"X127FA1" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF worker_number =	"X127FA4" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF worker_number =	"X127BV1" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF worker_number =	"X127FA2" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF worker_number =	"X127F3R" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF worker_number =	"X127F3Y" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF worker_number =	"X127FA3"	THEN email_address = "hsph.es.team.451@hennepin.us"
IF worker_number =	"X127FJ1" 	THEN email_address = "hsph.es.team.451@hennepin.us"
IF worker_number =	"X127ER8" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127F3B" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127ES1" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127ES3" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127FB6" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127F3H" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127F4E" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127FB4" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127F3A" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127FB5" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127FB3" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127ER9" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127ES2" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127F3M" 	THEN email_address = "hsph.es.team.452@hennepin.us"
IF worker_number =	"X127EY8"	THEN email_address = "hsph.es.team.455@hennepin.us"
IF worker_number =	"X127EY9"	THEN email_address = "hsph.es.team.455@hennepin.us"
IF worker_number =	"X127EZ1"	THEN email_address = "hsph.es.team.455@hennepin.us"
IF worker_number =	"X127EX7" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF worker_number =	"X127EY1" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF worker_number =	"X127FJ5" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF worker_number =	"X127F3Q" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF worker_number =	"X127EX9" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF worker_number =	"X127F3T" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF worker_number =	"X127EX8" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF worker_number =	"X127F3Z" 	THEN email_address = "hsph.es.team.458@hennepin.us"
IF worker_number =	"X127EU5" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127EU6" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127EY2" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127F3W" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127EU8" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127F3X" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127EU7" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127F3S" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127EU9" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127FJ3" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127FJ4" 	THEN email_address = "hsph.es.team.459@hennepin.us"
IF worker_number =	"X127ES4" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF worker_number =	"X127ES8" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF worker_number =	"X127EM6" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF worker_number =	"X127ES5" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF worker_number =	"X127ES6" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF worker_number =	"X127ES9" 	THEN email_address = "hsph.es.team.460@hennepin.us"
IF worker_number =	"X127EV1" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF worker_number =	"X127EV5" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF worker_number =	"X127EV2" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF worker_number =	"X127EV4" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF worker_number =	"X127EV3" 	THEN email_address = "hsph.es.team.461@hennepin.us"
IF worker_number =	"X127ET2"  	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127ET3" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127FJ2" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127ET1" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127EM5" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127EZ2" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127EZ9" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127EZ4" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127EZ3" 	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127EW9"	THEN email_address = "hsph.es.team.462@hennepin.us"
IF worker_number =	"X127EZ5" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF worker_number =	"X127EZ8" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF worker_number =	"X127FH6" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF worker_number =	"X127EZ6" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF worker_number =	"X127EZ7" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF worker_number =	"X127EZ0" 	THEN email_address = "hsph.es.team.463@hennepin.us"
IF worker_number =	"X127FA5" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF worker_number =	"X127FA6" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF worker_number =	"X127FA7" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF worker_number =	"X127FA8" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF worker_number =	"X127FB1" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF worker_number =	"X127FA9" 	THEN email_address = "hsph.es.team.465@hennepin.us"
IF worker_number =	"X127ET4" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127ET6" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127ET8" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127F4C" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127F4F" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127F4D" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127ET7"	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127ET5" 	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127BV3"	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127FB9"  	THEN email_address = "hsph.es.team.467@hennepin.us"
IF worker_number =	"X127FC1" 	THEN email_address = "hsph.es.team.467@hennepin.us"
IF worker_number =	"X127FC2"  	THEN email_address = "hsph.es.team.467@hennepin.us"
IF worker_number =	"X127EL1"	THEN email_address = "hsph.es.team.466@hennepin.us"
IF worker_number =	"X127FB8" 	THEN email_address = "hsph.es.team.467@hennepin.us"
IF worker_number =	"X127FB7" 	THEN email_address = "hsph.es.team.467@hennepin.us"
IF worker_number =	"X127FD4" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF worker_number =	"X127FD5" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF worker_number =	"X127FD8" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF worker_number =	"X127FD6" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF worker_number =	"X127FD9" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF worker_number =	"X127FD7" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF worker_number =	"X127EDD" 	THEN email_address = "hsph.es.team.468@hennepin.us"
IF worker_number =	"X127FG1" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF worker_number =	"X127EW6" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF worker_number =	"X1274EC" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF worker_number =	"X127FG2" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF worker_number =	"X127EW4" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF worker_number =	"X127EW5" 	THEN email_address = "hsph.es.team.469@hennepin.us"
IF worker_number =	"X127FE7" 	THEN email_address = "hsph.es.team.470@hennepin.us"
IF worker_number =	"X127FE8" 	THEN email_address = "hsph.es.team.470@hennepin.us"
IF worker_number =	"X127FE9" 	THEN email_address = "hsph.es.team.470@hennepin.us"
IF worker_number =	"X127EX4" 	THEN email_address = "hsph.es.team.603@hennepin.us"
IF worker_number =	"X127EX5" 	THEN email_address = "hsph.es.team.603@hennepin.us"
IF worker_number =	"X127FF1"	THEN email_address = "hsph.es.team.603@hennepin.us"
IF worker_number =	"X127FF2"	THEN email_address = "hsph.es.team.603@hennepin.us"
IF worker_number =	"X127EN8" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF worker_number =	"X127EN9" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF worker_number =	"X127FH3" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF worker_number =	"X127F3E" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF worker_number =	"X127F3J" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF worker_number =	"X127F3N" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF worker_number =	"X127FI6" 	THEN email_address = "hsph.es.team.602@hennepin.us"
IF worker_number =	"X127F4A" 	THEN email_address = "hsph.es.team.603@hennepin.us"
IF worker_number =	"X127F4B" 	THEN email_address = "hsph.es.team.603@hennepin.us"
IF worker_number =	"X127FI1"	THEN email_address = "hsph.es.team.603@hennepin.us"
IF worker_number =	"X127FI3" 	THEN email_address = "hsph.es.team.603@hennepin.us"
IF worker_number =	"X127EQ6" 	THEN email_address = "hsph.es.team.604@hennepin.us"
IF worker_number =	"X127EQ7" 	THEN email_address = "hsph.es.team.604@hennepin.us"
IF worker_number =	"X127EP1" 	THEN email_address = "hsph.es.team.604@hennepin.us"
IF worker_number =	"X127EP2" 	THEN email_address = "hsph.es.team.604@hennepin.us"
IF worker_number =	"X127FE2" 	THEN email_address = "hsph.es.team.605@hennepin.us"
IF worker_number =	"X127FE3" 	THEN email_address = "hsph.es.team.605@hennepin.us"
IF worker_number =	"X127FG5" 	THEN email_address = "hsph.es.team.605@hennepin.us"
IF worker_number =	"X127FG9" 	THEN email_address = "hsph.es.team.605@hennepin.us"
IF worker_number =	"X127EW7"	THEN email_address = "hsph.es.team.ebenezer@hennepin.us"
IF worker_number =	"X127EW8"	THEN email_address = "hsph.es.team.ebenezer@hennepin.us"
IF worker_number =	"X127ER7" 	THEN email_address = "hsph.es.team.mhc@hennepin.us"
IF worker_number =	"X127SH1" 	THEN email_address = "hsph.es.shelter.team@hennepin.us"
IF worker_number =	"X127AN1" 	THEN email_address = "hsph.es.shelter.team@hennepin.us"
IF worker_number =	"X127EHD" 	THEN email_address = "hsph.es.shelter.team@hennepin.us"

'----------------------------------------------------------------------------------------------------THE CASENOTE
Call start_a_blank_case_note
CALL write_variable_in_case_note("--CHANGE REPORTED - " & nature_change & "--")
CALL write_bullet_and_variable_in_Case_Note("Date Reported/Addendum", date_received)
CALL write_bullet_and_variable_in_Case_Note("Date Effective", effective_date)
CALL write_bullet_and_variable_in_case_note("Address", address_notes)
CALL write_bullet_and_variable_in_case_note("Household Members", household_notes)
CALL write_bullet_and_variable_in_case_note("Assets", asset_notes)
CALL write_bullet_and_variable_in_case_note("Vehicles", vehicles_notes)
CALL write_bullet_and_variable_in_case_note("Income", income_notes)
CALL write_bullet_and_variable_in_case_note("Shelter", shelter_notes)
CALL write_bullet_and_variable_in_case_note("Verifs Requested", verifs_requested)
'CALL write_variable_in_case_note("--CHANGE REPORTED -' HH Comp Change--")
CALL write_bullet_and_variable_in_Case_Note("Unit member HH Member", HH_Member)
IF Temporary_Change_Checkbox = CHECKED THEN CALL write_variable_in_Case_Note("* Change is temporary")
IF Temporary_Change_Checkbox = UNCHECKED THEN CALL write_variable_in_Case_Note("* Change is NOT temporary")
'babyborn casenote'
CALL write_bullet_and_variable_in_Case_Note("Child's's name", babys_name)
CALL write_bullet_and_variable_in_Case_Note("Gender", baby_gender)
CALL write_bullet_and_variable_in_Case_Note("Date of birth", date_of_birth)
IF nature_change = "Baby Born" or nature_change = "HHLD Comp" THEN
	IF fathers_name = "" THEN fathers_name = "Unknown or not provided"
	IF parent_in_household = "Yes" THEN
		father_HH = " - reported in the same household."
	ELSE
	   father_HH = " - not reported in the same household"
	END IF
END IF
CALL write_bullet_and_variable_in_Case_Note("Mother's name", mothers_name)
CALL write_bullet_and_variable_in_Case_Note("Mother's employer", mothers_employer)
CALL write_bullet_and_variable_in_Case_Note("Father's name", fathers_name & father_HH)
CALL write_bullet_and_variable_in_Case_Note("Father's employer", fathers_employer)
IF other_health_insurance = "Yes" THEN CALL write_bullet_and_variable_in_Case_Note("OHI", OHI_Source)
IF MHC_plan_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE("* Newborns MHC plan updated to match the mothers.")
CALL write_bullet_and_variable_in_case_note("Other Change Reported", other_change_notes)
IF changes_continue <> "Select:" THEN CALL write_bullet_and_variable_in_case_note("The changes client reports", changes_continue)
CALL write_bullet_and_variable_in_case_note("Action Taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

IF send_email_checkbox = CHECKED THEN
	EMWriteScreen "x", 5, 3
	Transmit
	note_row = 4			'Beginning of the case notes
	Do 						'Read each line
		EMReadScreen note_line, 76, note_row, 3
		note_line = trim(note_line)
		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
		message_array = message_array & note_line & vbcr		'putting the lines together
		note_row = note_row + 1
		If note_row = 18 then 									'End of a single page of the case note
			EMReadScreen next_page, 7, note_row, 3
			If next_page = "More: +" Then 						'This indicates there is another page of the case note
				PF8												'goes to the next line and resets the row to read'\
				note_row = 4
			End If
		End If
	Loop until next_page = "More:  " OR next_page = "       "	'No more pages

	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	CALL create_outlook_email(email_address, "","Change Reported For #" &  MAXIS_case_number & " Member # " & memb_number & " Date Change Reported " & date_received, "CASE NOTE" & vbcr & message_array,"", False)
END IF

script_end_procedure ("The case note has been created please be sure to send verifications to ECF and/or case note how the information was received.")
