'Required for statistical purposes==========================================================================================
name_of_script = "DAIL - INFC CSIA.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 240          'manual run time in seconds
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/16/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK ======================================================================================================
'For the dail scrubber REFERRAL/AB PARENT
'NON-EDITABLE FIELDS MUST BE UPDATED ON THE ABPS PANEL'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

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
'EMReadScreen rel_to_applicant, 2, 10, 42
'EMReadScreen MEMB_gender, 1, 9, 42
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
	'----------------------------------------------------------------------------------------------------ABPS panel
	Call MAXIS_footer_month_confirmation
	Call navigate_to_MAXIS_screen("STAT", "ABPS")
DO
	EMReadScreen child_ref_numb, 2, row, 35
Loop Until child_ref_numb = ""
	EMReadScreen parental_status, 1, 15, 53	'making sure ABPS is not unknown.
	msgbox "What is the status"
	IF parental_status = "2" THEN
		ABPS_client_name = "Unknown"
	ELSEIF parental_status = "3" THEN
		ABPS_client_name = "ABPS deceased"
	ELSEIF parental_status = "4" THEN
		ABPS_client_name = "Rights Severed"
	ELSEIF parental_status = "7" THEN
		ABPS_client_name = "HC No Order Sup"
	ELSEIF parental_status = "1" THEN
		EMReadScreen custodial_status, 1, 15, 57
		EMReadScreen first_name, 12, 10, 63
		EMReadScreen last_name, 24, 10, 30
		first_name = trim(first_name)
		last_name = trim(last_name)
		first_name = replace(first_name, "_", "")
		last_name = replace(last_name, "_", "")
		ABPS_client_name = first_name & " " & last_name
		Call fix_case_for_name(ABPS_client_name)
		EMReadScreen ABPS_gender, 1, 11, 80	'reading the ssn
		EMReadScreen ABPS_SSN, 11, 11, 30	'reading the ssn
		EMReadScreen ABPS_DOB, 10, 11, 60	'reading the DOB
		EMReadScreen ABPS_parent_ID, 10, 13, 40	'making sure ABPS is not unknown.
		ABPS_parent_ID = trim(ABPS_parent_ID)
		EMReadScreen HC_ins_order, 1, 12, 44	'making sure ABPS is not unknown.
		EMReadScreen HC_ins_compliance, 1, 12, 80
	END IF
	'24, 02"THIS DATA WILL EXPIRE ON --/--/--"
DO
	EMReadScreen panel_number, 1, 2, 78
	If panel_number = "0" then script_end_procedure("An ABPS panel does not exist. Please create the panel before running the script again. ")
	Do
		EMReadScreen current_panel_number, 1, 2, 73
		ABPS_check = MsgBox("Is this the right ABPS?  " & ABPS_parent_ID, vbYesNo + vbQuestion, "Confirmation")
		If ABPS_check = vbYes then exit do
		If ABPS_check = vbNo then TRANSMIT
		If (ABPS_check = vbNo AND current_panel_number = panel_number) then	script_end_procedure("Unable to find another ABPS. Please review the case, and run the script again if applicable.")
	Loop until current_panel_number = panel_number

	EMReadScreen ABPS_screen, 4, 2, 50		'if inhibiting error exists, this will catch it and instruct the user to update ABPS
	'msgbox ABPS_screen
	'If ABPS_screen = "ABPS" then script_end_procedure("An error occurred on the ABPS panel. Please update the panel before using the script with the absent parent information.")
	'seting variables for the programs included
	If good_cause_droplist = "Change/exemption ending" then
  	Do
  		Do
  			err_msg = ""
  			dialog change_exemption_dialog
  			cancel_confirmation
  			If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
  		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
  		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  	Loop until are_we_passworded_out = false					'loops until user passwords back in
	END IF

	'trims excess spaces of programs
	programs_included  = trim(programs_included )
	'takes the last comma off of programs
	If right(programs_included, 1) = "," THEN programs_included  = left(programs_included, len(programs_included) - 1)
CSIA
'Making sure we have the correct CSIA
EMReadScreen current_panel_number, 1, 2, 73
EMReadScreen total_panel_number, 1, 2, 78
If current_panel_number = "0" then script_end_procedure("An CSIA panel does not exist.")
If current_panel_number = total_panel_number THEN
Do
	EMReadScreen current_panel_number, 1, 2, 73
	CSIA_check = MsgBox("Is this the right CSIA?", vbYesNo + vbQuestion, "Confirmation")
	If CSIA_check = vbYes then exit do
	If CSIA_check = vbNo then TRANSMIT
	If (CSIA_check = vbNo AND current_panel_number = panel_number) then	script_end_procedure("Unable to find another ABPS. Please review the case, and run the script again if applicable.")
Loop Until
'IF rel_to_applicant = "01" and 'MEMB_gender = "F" THEN Crgvr_Rel_to_Child = "M"
ABPS_parentage = "N"
EmWriteScreen Crgvr_Rel_to_Child, 4, 79
EmWriteScreen ABPS_Deceased, 11, 44
EmWriteScreen Name_Known, 12, 44
EmWriteScreen Mult_Alleged_Fathers, 12, 75
Row = 14
DO
	EMReadScreen CSIA_child_ref, 2, row, 04
	IF CSIA_child_ref <> "" THEN EmWriteScreen ABPS_parentage, row, 75
	TRANSMIT 'CSIB'
	TRANSMIT 'CSIC'
	TRANSMIT 'CSID'
	EMReadScreen panel_number, 1, 2, 78
	If panel_number = "1" then exit DO
Loop

script_end_procedure("Success! CSIA has been updated.")
