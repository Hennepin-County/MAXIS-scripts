'GATHERING STATS===========================================================================================
name_of_script = "NOTES - SHELTER-MONEY MISMANAGEMENT.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
call changelog_update("06/26/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'--------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

when_contact_was_made = date
date_requested = date & ""
income_checkbox = checked
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 346, 105, " Money Mismanagement "
  EditBox 55, 5, 45, 15, maxis_case_number
  ComboBox 160, 5, 60, 15, "Phone call"+chr(9)+"Voicemail"+chr(9)+"Email"+chr(9)+"Office visit"+chr(9)+"Letter", contact_type
  DropListBox 225, 5, 45, 10, "to"+chr(9)+"from", contact_direction
  ComboBox 275, 5, 65, 15, "Client"+chr(9)+"Other HH Memb"+chr(9)+"AREP", who_contacted
  EditBox 100, 25, 45, 15, date_requested
  EditBox 185, 25, 85, 15, phone_number
  DropListBox 50, 45, 145, 15, "1st Instance of Money Mismanagement"+chr(9)+"2nd Instance of Money Mismanagement"+chr(9)+"Grant Management", Occurrence_droplist
  CheckBox 210, 45, 100, 10, "Income no longer available", income_checkbox
  EditBox 50, 65, 220, 15, Comments_notes
  ButtonGroup ButtonPressed
    OkButton 235, 85, 50, 15
    CancelButton 290, 85, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 110, 10, 50, 10, "Contact Type:"
  Text 5, 30, 90, 10, "Client requested shelter on: "
  Text 155, 30, 25, 10, "Phone:"
  Text 5, 50, 40, 10, "Occurrence:"
  Text 5, 70, 40, 10, "Comments:"
EndDialog

DO
	Do
		Dialog Dialog1
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'1st Instance of Money Mismanagement = occurrence_MM1--------SAVE for ENHANCEMNT
'2nd Instance of Money Mismanagement = occurrence_MM2
'Grant Management = occurrence_MM3

back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

months_variable = CM_mo
IF CM_MO = "01" THEN months_variable = "January, February, March"
IF CM_MO = "02" THEN months_variable = "February, March, April"
IF CM_MO = "03" THEN months_variable = "March, April, May"
IF CM_MO = "04" THEN months_variable = "April, May, June"
IF CM_MO = "05" THEN months_variable = "May, June, July"
IF CM_MO = "06" THEN months_variable = "June, July, August"
IF CM_MO = "07" THEN months_variable = "July, August, September"
IF CM_MO = "08" THEN months_variable = "August, September, October"
IF CM_MO = "09" THEN months_variable = "September, October, November"
IF CM_MO = "10" THEN months_variable = "October, November, December"
IF CM_MO = "11" THEN months_variable = "November, December, January"
IF CM_MO = "12" THEN months_variable = "December, January, February"

'----------------------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("### " & Occurrence_droplist & " ###")
CALL write_variable_in_CASE_NOTE("* Contacted " & who_contacted & "on " & when_contact_was_made & " by " & contact_type & " " & contact_direction & " "& phone_number & " ")
CALL write_variable_in_CASE_NOTE("* Client requested shelter on " & date_requested & " and all GA/SSI is gone." )
Call write_bullet_and_variable_in_CASE_NOTE("Comments", Comments_notes)
Call write_variable_in_CASE_NOTE("---")
IF Occurrence_droplist = "2nd Instance of Money Mismanagement" THEN  CALL write_variable_in_CASE_NOTE("* 1st money mismanagement was XX/XX.")
IF Occurrence_droplist = "Grant Management" THEN
    CALL write_variable_in_CASE_NOTE("*** GA=$97.00 FOR 3 MONTHS/Grant Management/NO MATTER WHERE CLIENT LIVES ***")
    CALL write_variable_in_CASE_NOTE("* 1st money mismanagement was XX/XX.")
    CALL write_variable_in_CASE_NOTE("* Second money mismanagement was XX/XX.")
    CALL write_variable_in_CASE_NOTE("* No matter where client lives the grant will be $97.00 for three months. ")
    CALL write_variable_in_CASE_NOTE("* Grant reduced to $97.00 effective: " & months_variable)
END IF
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
