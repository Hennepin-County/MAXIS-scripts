'Gathering stats=========================================
name_of_script = "ACTIONS - SHELTER EXPENSE VERIF RECEIVED.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 125          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'THE DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 146, 70, "Case number dialog"
  EditBox 80, 5, 60, 15, case_number
  EditBox 80, 25, 25, 15, MAXIS_footer_month
  EditBox 115, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 45, 50, 15
    CancelButton 90, 45, 50, 15
  Text 10, 30, 65, 10, "Footer month/year:"
  Text 10, 10, 45, 10, "Case number: "
EndDialog

'>>>>>THE NEW DIALOG<<<<<
BeginDialog shelter_form_received_dialog, 0, 0, 276, 395, "Shelter Expenses Dialog"
  EditBox 105, 10, 60, 15, agency_received_date
  EditBox 245, 10, 20, 15, hh_member
  EditBox 45, 35, 45, 15, unit_rent
  EditBox 195, 35, 45, 15, client_share
  CheckBox 10, 60, 155, 10, "Check here if the client's rent is subsidized.", subsidy_check
  EditBox 230, 55, 40, 15, subsidy_amount
  CheckBox 10, 80, 115, 10, "Check here if garage is required.", garage_check
  EditBox 230, 75, 40, 15, garage_amount
  CheckBox 10, 100, 145, 10, "Check here if this is for Room and Board.", room_and_board_check
  EditBox 115, 115, 155, 15, room_board_notes
  DropListBox 105, 135, 90, 15, "Select one..."+chr(9)+"Heat/AC"+chr(9)+"Phone/Electric"+chr(9)+"Electric ONLY"+chr(9)+"Phone ONLY"+chr(9)+"All utilities included in rent"+chr(9)+"None", utilities_paid_listbox
  EditBox 95, 175, 55, 15, move_in_date
  EditBox 90, 195, 130, 15, new_address
  EditBox 90, 215, 65, 15, new_addr_city
  EditBox 155, 215, 25, 15, new_addr_state
  EditBox 180, 215, 40, 15, new_addr_zip
  EditBox 85, 235, 25, 15, county_code
  EditBox 110, 255, 25, 15, num_of_residents
  CheckBox 10, 290, 170, 10, "Check here if the form was signed by the client.", signed_by_client_check
  CheckBox 10, 300, 205, 10, "Check here if the form was signed by the landlord/manager.", signed_by_landlord_check
  EditBox 60, 320, 210, 15, other_notes
  EditBox 65, 340, 205, 15, actions_taken
  EditBox 75, 375, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 170, 375, 50, 15
    CancelButton 220, 375, 50, 15
  Text 10, 15, 90, 10, "Date Received by Agency:"
  Text 35, 200, 45, 10, "New address:"
  Text 35, 260, 70, 10, "Number of residents"
  Text 10, 345, 50, 10, "Actions taken:"
  Text 30, 120, 80, 10, "Room and Board Notes"
  Text 10, 40, 30, 10, "Unit rent"
  Text 35, 180, 55, 10, "Date moved in:"
  Text 100, 40, 90, 10, "Client's amount (if different)"
  Text 175, 60, 50, 10, "Subsidy Amt:"
  Text 175, 80, 45, 10, "Garage Amt:"
  GroupBox 25, 160, 230, 115, "Client Move Information -- Use these fields if the client has moved"
  Text 10, 140, 85, 10, "Utilities paid by resident:"
  Text 10, 325, 45, 10, "Other notes:"
  Text 10, 380, 60, 10, "Worker signature:"
  Text 35, 240, 50, 10, "County Code:"
  Text 195, 15, 40, 10, "HH Member"
EndDialog


'>>>>> DLG FOR ADDITIONAL SHEL INFO <<<<<
BeginDialog shel_dlg, 0, 0, 141, 150, "Additional SHEL Info Required"
  ButtonGroup ButtonPressed
    OkButton 35, 125, 50, 15
    CancelButton 85, 125, 50, 15
  Text 10, 15, 55, 10, "Landlord Name:"
  EditBox 70, 10, 65, 15, landlord_name
  Text 20, 55, 45, 10, "Rent"
  Text 20, 75, 45, 10, "Garage"
  Text 20, 95, 45, 10, "Subsidy"
  EditBox 75, 50, 50, 15, retro_rent
  EditBox 75, 70, 50, 15, retro_garage
  EditBox 75, 90, 50, 15, retro_subsidy
  GroupBox 10, 35, 125, 80, "Retro Information"
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to Bluezone & grabbing case number and footer year/month
EMConnect ""
CALL MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
hh_member = "01"

DO
	DO
		Dialog case_number_dialog																'calls up dialog for worker to enter case number and applicable month and year.	 Script will 'loop' 
		IF buttonpressed = 0 THEN StopScript						   'and verbally request the worker to enter a case number until the worker enters a case number.
		IF case_number = "" THEN MsgBox "You must enter a case number"
	LOOP UNTIL case_number <> ""
	
	'Getting to the correct benefit month
	CALL find_variable("Month: ", benefit_month, 5)
	benefit_month = replace(benefit_month, " ", "/")
	IF benefit_month <> MAXIS_footer_month & "/" & MAXIS_footer_year THEN 
		back_to_SELF
		EMWriteScreen "STAT", 16, 43
		EMWriteScreen "________", 18, 43
		EMWriteScreen case_number, 18, 43
		EMWriteScreen MAXIS_footer_month, 20, 43
		EMWriteScreen MAXIS_footer_year, 20, 46
		transmit
		
		'Checking to see if the case is stuck in background.
		row = 1
		col = 1
		EMSearch "BACKGROUND", row, col
		IF row <> 0 THEN script_end_procedure("The case is stuck in background. Please try again.")	
		'Checking to see if the case is privileged.
		EMReadScreen privileged_check, 60, 24, 2
		IF InStr(privileged_check, "PRIVILEGE") <> 0 THEN script_end_procedure("You do not have access to this case. The script will now stop.")
	END IF	
	valid_case_number = True
	row = 1
	col = 1
	EMSearch "INVALID CASE NUMBER", row, col
	IF row <> 0 THEN 
		MsgBox "The case number you entered is not a valid MAXIS case number. Please try again."
		valid_case_number = False
	END IF
LOOP UNTIL valid_case_number = True

DO
	err_msg = ""
	DIALOG shelter_form_received_dialog
		cancel_confirmation
		'Enforcing required fields
		IF agency_received_date = "" THEN err_msg = err_msg & vbCr & "* Please indicate the date the document was received by the agency."
		IF agency_received_date <> "" AND IsDate(agency_received_date) = TRUE THEN
			IF DateDiff("D", agency_received_date, date) < 0 THEN err_msg = err_msg & vbCr & "* You may not enter a future date for the date the form was received by the agency."
		ELSEIF agency_received_date <> "" AND IsDate(agency_received_date) = FALSE THEN 
			err_msg = err_msg & vbCr & "* Please enter a valid date for the date the form was received by the agency."
		END IF
		IF hh_member = "" THEN err_msg = err_msg & vbCr & "* Please indicate the household member."
		IF unit_rent = "" THEN err_msg = err_msg & vbCr & "* Please indicate the unit rent."
		IF subsidy_check = 1 AND subsidy_amount = "" THEN err_msg = err_msg & vbCr & "* Please indicate the amount of the client's subisdy."
		IF subsidy_check = 1 AND subsidy_amount <> "" AND client_share = "" THEN err_msg = err_msg & vbCr & "* Please indicate what the client's share is after the subsidy is applied."
		IF garage_check = 1 AND garage_amount = "" THEN err_msg = err_msg & vbCr & "* Please indicate the amount for the garage."
		IF room_and_board_check = 1 AND room_board_notes = "" THEN err_msg = err_msg & vbCr & "* You indicated that the client's rent is room and board. Please provide detailed notes about the room and board situation."
		IF utilities_paid_listbox = "Select one..." THEN err_msg = err_msg & vbCr & "* Please indicate which, if any, utilities are paid by the client."
		IF move_in_date <> "" THEN 
			IF IsDate(move_in_date) = False THEN err_msg = err_msg & vbCr & "* Please enter the client move-in date as a valid date."
		END IF
		IF actions_taken = "" THEN err_msg = err_msg & vbCr & "* Please indicate the actions you have taken."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'Going through and updating MAXIS...
'Not updating if room and board. That is more complex than we are taking on at this moment and should
'probably receive more worker scrutiny
IF room_and_board_check = 1 THEN 
	MsgBox "The script will not update SHEL because this client is indicating room and board. Please update SHEL manually. Press OK to continue."
ELSE
	'Going to SHEL
	CALL navigate_to_MAXIS_screen("STAT", "SHEL")
	'...for the specific HH member
	CALL write_value_and_transmit(hh_member, 20, 76)
	EMReadScreen num_of_SHEL_panels, 1, 2, 78
	IF num_of_SHEL_panels = "0" THEN 
		'If needed, creating a new panel...
		EMWriteScreen hh_member, 20, 76
		CALL write_value_and_transmit("NN", 20, 79)
	ELSE
		'...or just editing the one we already have.
		PF9
	END IF
	
	'If client_share is blank, the script will assume that the client pays the entire unit amount and convert that value to client_share.
	IF client_share = "" THEN client_share = unit_rent
	
	'Writing the prospective subsidy amount.
	IF subsidy_check = 1 THEN 
		EMWriteScreen "Y", 6, 46
		EMWriteScreen "________", 18, 56
		EMWriteScreen subsidy_amount, 18, 56
		EMWriteScreen "SF", 18, 67
	ELSE
		EMWriteScreen "N", 6, 46
		EMWriteScreen "________", 18, 56
		EMWriteScreen "________", 18, 67
	END IF
	
	'Writing the prospective garage amount.
	IF garage_check = 1 THEN 
		EMWriteScreen "________", 17, 56
		EMWriteScreen garage_amount, 17, 56
		EMWriteScreen "SF", 17, 67
	ELSE
		EMWriteScreen "________", 17, 56
		EMWriteScreen "________", 17, 67
	END IF
	
	'Writing the prospective rent amount.
	EMWriteScreen "________", 11, 56
	EMWriteScreen client_share, 11, 56
	EMWriteScreen "SF", 11, 67
	
	'Grabbing the existing landlord name, the retrospective rent, garage and subsidy amounts.
	EMReadScreen landlord_name, 25, 7, 50
	landlord_name = replace(landlord_name, "_", "")
	EMReadScreen retro_rent, 8, 11, 37
		retro_rent = replace(retro_rent, "_", "")
		retro_rent = trim(retro_rent)
	EMReadScreen retro_garage, 8, 17, 37
		retro_garage = replace(retro_garage, "_", "")
		retro_garage = trim(retro_garage)
	EMReadScreen retro_subsidy, 8, 18, 37
		retro_subsidy = replace(retro_subsidy, "_", "")
		retro_subsidy = trim(retro_subsidy)
	
	'Running the retro dialog
	DO
		err_msg = ""
		DIALOG shel_dlg
			cancel_confirmation
			IF landlord_name = "" THEN err_msg = err_msg & vbCr & "* Please enter a landlord name."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	
	CALL check_for_MAXIS(false)
	
	'Writing the landlord and retrospective information
	EMWriteScreen "____________________", 7, 50
	EMWriteScreen landlord_name, 7, 50
	EMWriteScreen "N", 6, 64
	
	IF retro_rent <> "" THEN 
		EMWriteScreen "________", 11, 37
		EMWriteScreen retro_rent, 11, 37
		EMWriteScreen "SF", 11, 48
	ELSE
		EMWriteScreen "________", 11, 37
		EMWriteScreen "__", 11, 48
	END IF

	IF retro_garage <> "" THEN 
		EMWriteScreen "________", 17, 37
		EMWriteScreen retro_garage, 17, 37
		EMWriteScreen "SF", 17, 48
	ELSE
		EMWriteScreen "________", 17, 37
		EMWriteScreen "__", 17, 48
	END IF
	
	IF retro_subsidy <> "" THEN 
		EMWriteScreen "________", 18, 37
		EMWriteScreen retro_subsidy, 18, 37
		EMWriteScreen "SF", 18, 48
	ELSE
		EMWriteScreen "________", 18, 37
		EMWriteScreen "__", 18, 48
	END IF
	
	'Closing SHEL
	transmit
	transmit
	transmit
	
END IF

EMWriteScreen "HEST", 20, 71
CALL write_value_and_transmit(hh_member, 20, 76)
'>>>>> DETERMINING WHETHER HEST EXISTS <<<<<
EMReadScreen num_of_HEST, 1, 2, 78
IF num_of_HEST = "0" THEN 
	CALL write_value_and_transmit("NN", 20, 79)
ELSE
	PF9
END IF

EMWriteScreen hh_member, 06, 40
CALL create_MAXIS_friendly_date(date, 0, 7, 40)
IF utilities_paid_listbox = "Heat/AC" THEN 
	EMWriteScreen "Y", 13, 60
	EMWriteScreen "01", 13, 68
	EMWriteScreen "_", 14, 60
	EMWriteScreen "__", 14, 68
	EMWriteScreen "_", 15, 60
	EMWriteScreen "__", 15, 68
ELSEIF utilities_paid_listbox = "Phone/Electric" THEN 
	EMWriteScreen "_", 13, 60
	EMWriteScreen "__", 13, 68
	EMWriteScreen "Y", 14, 60
	EMWriteScreen "01", 14, 68
	EMWriteScreen "Y", 15, 60
	EMWriteScreen "01", 15, 68
ELSEIF utilities_paid_listbox = "Electric ONLY" THEN 
	EMWriteScreen "_", 13, 60
	EMWriteScreen "__", 13, 68
	EMWriteScreen "Y", 14, 60
	EMWriteScreen "01", 14, 68
	EMWriteScreen "_", 15, 60
	EMWriteScreen "__", 15, 68
ELSEIF utilities_paid_listbox = "Phone ONLY" THEN 
	EMWriteScreen "_", 13, 60
	EMWriteScreen "__", 13, 68
	EMWriteScreen "_", 14, 60
	EMWriteScreen "__", 14, 68
	EMWriteScreen "Y", 15, 60
	EMWriteScreen "01", 15, 68
ELSEIF utilities_paid_listbox = "None" OR utilities_paid_listbox = "All utilities included in rent" THEN
	'Deleting HEST if the client is not paying any utilities
	EMWriteScreen "DEL", 20, 71
END IF
transmit

'If the client has a new address reported on the form...
IF new_address <> "" THEN 
	CALL write_value_and_transmit("ADDR", 20, 71)
	PF9
	
	EMReadScreen benefit_month, 2, 20, 55
	EMReadScreen benefit_year, 2, 20, 58
	EMWriteScreen benefit_month, 4, 43
	EMWriteScreen "01", 4, 46
	EMWriteScreen benefit_year, 4, 49
	
	EMWriteScreen "______________________", 6, 43
	EMWriteScreen "______________________", 7, 43
	EMWriteScreen new_address, 6, 43
	EMWriteScreen "_______________", 8, 43
	EMWriteScreen new_addr_city, 8, 43
	EMWriteScreen new_addr_state, 8, 66
	EMWriteScreen "________", 9, 43
	EMWriteScreen new_addr_zip, 9, 43
	EMWriteScreen county_code, 9, 66
	EMWriteScreen "SF", 9, 74
	'Making sure that "homeless y/n" is "N"
	EMWriteScreen "N", 10, 43
	
	transmit
	
	'Checking to see if the address can be standardized.
	'The script will not accept addresses that are not standardized.
	EMReadScreen standardized, 33, 24, 2
	IF standardized <> "RESIDENCE ADDRESS IS STANDARDIZED" THEN 
		row = 1
		col = 1
		EMSearch "RESIDENCE ADDRESS, CITY, STATE, AND ZIP ARE REQUIRED", row, col
		IF row <> 0 THEN 
			PF10
			MsgBox "The address is not acceptable to MAXIS. The script has undid the update to ADDR. Press OK to continue to case note."
			valid_addr = False
		ELSE
			row = 1
			col = 1
			EMSearch "Warning: Mail to this Residence address will not be", row, col
			IF row <> 0 THEN 
				PF10
				PF10
				valid_addr = False
				MsgBox "This address is not standardized. Your client would not have received mail at this address. The script has undid the update to ADDR." & vbCr & vbCr & "Press OK for the script to continue to case noting."
			END IF
		END IF
	ELSE
		valid_addr = True
	END IF
END IF

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_case_note ("~~~ Shelter Expense Verif Received on " & agency_received_date & " ~~~")
CALL write_bullet_and_variable_in_case_note ("Unit Rent", FormatCurrency(unit_rent))
IF client_share <> unit_rent THEN CALL write_bullet_and_variable_in_case_note("Client's Share", FormatCurrency(client_share))
IF subsidy_check = 1 THEN CALL write_bullet_and_variable_in_case_note("Unit is subsidized. Subsidy Amount", FormatCurrency(subsidy_amount))
CALL write_bullet_and_variable_in_case_note("Utilities Paid by Client", utilities_paid_listbox)
'Case noting information about client move
IF move_in_date <> "" THEN 
	CALL write_variable_in_case_note("---")
	CALL write_bullet_and_variable_in_case_note("Client reports moving. Move date", move_in_date)
	CALL write_bullet_and_variable_in_case_note("New Address", new_address)
	CALL write_variable_in_case_note           ("               " & new_addr_city & ", " & new_addr_state & " " & new_addr_zip)
	IF landlord_name <> "" THEN CALL write_bullet_and_variable_in_case_note("Landlord/Ppty Owner", landlord_name)
	CALL write_bullet_and_variable_in_case_note("County of Residence Code", county_code)
	IF valid_addr = False AND update_MAXIS_check = 1 THEN 
		CALL write_variable_in_case_note("**ADDRESS COULD NOT BE STANDARDIZED IN MAXIS.**")
		CALL write_variable_in_case_note("**THE SCRIPT DID NOT UPDATE ADDR.**")
	END IF
	Call write_bullet_and_variable_in_case_note ("Number of Occupants", num_of_residents)
END IF
CALL write_variable_in_case_note("---")

'Case noting specificically what the script did depending on the outcome of the actions taken updating MAXIS
IF valid_addr = False AND room_and_board_check = 0 THEN 
	CALL write_variable_in_case_note("* SHEL and HEST updated with script.")
ELSEIF valid_addr = True AND room_and_board_check = 0 THEN 
	CALL write_variable_in_case_note("* SHEL, HEST, and ADDR updated with script.")
ELSEIF valid_addr = False AND room_and_board_check = 1 THEN 
	CALL write_variable_in_case_note("* HEST updated with script.")
ELSEIF valid_addr = True AND room_and_board_check = 1 THEN 
	CALL write_variable_in_case_note("* HEST and ADDR updated with script.")
END IF

Call write_bullet_and_variable_in_case_note ("Other Notes", other_notes)
CALL write_bullet_and_variable_in_case_note("Actions Taken", actions_taken)
IF signed_by_landlord_check = 1 THEN Call write_variable_in_case_note ("* Form signed by landlord.")
IF signed_by_client_check = 1 THEN Call write_variable_in_case_note ("* Form signed by client.")
IF update_MAXIS_check = 1 THEN CALL write_bullet_and_variable_in_case_note("Panels updated via script", updated_panels)
Call write_variable_in_case_note ("---")
call write_variable_in_case_note (worker_signature)

script_end_procedure ("Success!!")
