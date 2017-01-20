'Required for statistical purposes===============================================================================
name_of_script = "DAIL - CITIZENSHIP VERIFIED.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 55          'manual run time in seconds
STATS_denomination = "M"       'M is for each MEMBER
'END OF stats block==============================================================================================

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

'custom function for this script---------------------------------------------------------------------------------
Function HH_member_custom_dialog_cit_id_ver(HH_member_array)

	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 3, 4, 33
		EMReadscreen last_name, 5, 6, 30
		EMReadscreen first_name, 7, 6, 63
		EMReadscreen Mid_intial, 1, 6, 79
		last_name = replace(last_name, "_", "") & " "
		first_name = replace(first_name, "_", "") & " "
		mid_initial = replace(mid_initial, "_", "")
		client_string = ref_nbr & last_name & first_name & mid_intial
		client_array = client_array & client_string & "|"
		transmit
		Emreadscreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	client_array = TRIM(client_array)
	test_array = split(client_array, "|")
	total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

	DIM all_client_array()
	ReDim all_clients_array(total_clients, 1)

	FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
		Interim_array = split(client_array, "|")
		all_clients_array(x, 0) = Interim_array(x)
		all_clients_array(x, 1) = 0			'Defaulting to update none persons so the user has to update them thar persons
	NEXT

	BEGINDIALOG HH_memb_dialog, 0, 0, 191, (35 + (total_clients * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
		Text 10, 5, 105, 10, "Household members to update:"
		FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
			IF all_clients_array(i, 0) <> "" THEN checkbox 10, (20 + (i * 15)), 120, 10, all_clients_array(i, 0), all_clients_array(i, 1)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
		NEXT
		ButtonGroup ButtonPressed
		OkButton 135, 10, 50, 15
		CancelButton 135, 30, 50, 15
	ENDDIALOG
													'runs the dialog that has been dynamically created. Streamlined with new functions.
	Call navigate_to_MAXIS_screen("DAIL","DAIL")

	'Sticking a do/loop around the dialog call to verify that the user has selected some household members.
	DO
		Dialog HH_memb_dialog
		If buttonpressed = 0 then stopscript
		check_for_maxis(True)

		HH_member_array = ""

		FOR i = 0 to total_clients
			IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
				IF all_clients_array(i, 1) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
					'msgbox all_clients_
					HH_member_array = HH_member_array & left(all_clients_array(i, 0), 2) & " "
				END IF
			END IF
		NEXT

		'If the user has not selected any household members, they will receive a msgbox informing them of that, requesting that they either try again or go home
		IF HH_member_array = "" THEN
			nobody_selected = MsgBox("You have not selected any household members to update. Press OK to try again. Press CANCEL to stop the script.", vbOKCancel)
			IF nobody_selected = vbCancel THEN stopscript
		END IF
	LOOP UNTIL HH_member_array <> ""

	HH_member_array = TRIM(HH_member_array)							'Cleaning up array for ease of use.
	HH_member_array = SPLIT(HH_member_array, " ")
END Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting
EMConnect ""

'Setting variables
row = 1
col = 1

'Finding case number
EMSearch "CASE NBR:", row, col
If row <> 0 then
  EMReadScreen MAXIS_case_number, 8, row, col + 10
  MAXIS_case_number = replace(MAXIS_case_number, "_", "")
  MAXIS_case_number = trim(MAXIS_case_number)
End if

'Error out in case it can't find the case number
If row = 0 then script_end_procedure("A case number could not be found on this DAIL message. Use the ''MAXIS notes'' version of the script at this time.")

Call HH_member_custom_dialog_cit_id_ver(HH_member_array)

'Updated MEMI section-------------------------------------------------------------------------------------------------------
Call navigate_to_MAXIS_screen("STAT","MEMI")
For Each HH_memb in HH_member_array
	EMWriteScreen HH_memb, 20, 76
	Transmit
	PF9
	EMWriteScreen "OT", 10, 78			'writing OT verif since verif is based on automated dail message.
	call create_MAXIS_friendly_date_with_YYYY(date, 0, 6, 35)   'writing actual date of change based on current date.
	Transmit
	Transmit 'second transmit to get past if you enter an actual date in another footer month
	membs_to_case_note = membs_to_case_note & HH_memb & ", "
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
Next

'Dialog to get worker signature----------------------------------------------------------------------------------------------
BeginDialog workersig_dlg, 0, 0, 201, 50, "Dialog"
  EditBox 50, 10, 150, 15, worker_signature
  Text 5, 5, 40, 20, "Worker Signature:"
  ButtonGroup ButtonPressed
    OkButton 85, 30, 50, 15
    CancelButton 140, 30, 50, 15
EndDialog

dialog workersig_dlg
cancel_confirmation
STATS_counter = STATS_counter - 1 'Had to -1 at the end of the script because the counter starts at 1 and Veronica has reasons why we should not change it to 0.
'Msgbox STATS_counter

'Case note section-----------------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
call write_variable_in_CASE_NOTE("***CITIZENSHIP/IDENTITY***")
Call write_variable_in_CASE_NOTE("Automated script has updated MEMI with OT for clients selected by worker. Information was provided to worker via Citizenship/ID Dail")
Call write_variable_in_CASE_NOTE("Members updated: " & membs_to_case_note)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

'Offers worker option to navigate back to DAIL message-----------------------------------------------------------------------------------------------------------
Navigate_Choice = MsgBox("Would you like to navigate back to the DAIL message? Press YES to navigate to DAIL, press NO to stay in the case note.", vbYesNo, "Navigate back to DAIL?")
If Navigate_Choice = vbYes then
	PF3 'to save casenote'
	Call navigate_to_MAXIS_screen("DAIL", "DAIL")
End if

script_end_procedure("")
