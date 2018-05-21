'STATS GATHERING-----------------------------------------------------1-----------------------------------------------
name_of_script = "NOTES - SHELTER-EA EXTENSION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
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
CALL changelog_update("02/09/2018", "Updated for requested extension number.", "MiKayla Handley, Hennepin County")
call changelog_update("11/20/2016", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog ea_extension, 0, 0, 296, 95, "EA Extension"
  EditBox 55, 10, 45, 15, MAXIS_case_number
  DropListBox 150, 10, 45, 15, "Select one..."+chr(9)+"1st"+chr(9)+"2nd", approval_number
  DropListBox 245, 10, 45, 15, "Select one..."+chr(9)+"1"+chr(9)+"2", ext_number
  EditBox 105, 30, 45, 15, start_date
  EditBox 165, 30, 45, 15, end_date
  EditBox 55, 50, 235, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 185, 75, 50, 15
    CancelButton 240, 75, 50, 15
  CheckBox 225, 35, 60, 10, "HSS approved", Check1
  Text 5, 15, 45, 10, "Case number:"
  Text 110, 15, 35, 10, "Approval #"
  Text 205, 15, 40, 10, "Extension #"
  Text 5, 35, 100, 10, "EA extended for 30 days from:"
  Text 155, 35, 10, 10, "to"
  Text 5, 55, 40, 10, "Other notes:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog ea_extension
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF approval_number = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select the approval number."
		If start_date = "" then err_msg = err_msg & vbNewLine & "* Enter start date."
		If end_date = "" then err_msg = err_msg & vbNewLine & "* Enter end date"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

approval_dates = start_date & "-" & end_date

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("###" & approval_number & " EA Extension(# " & ext_number & ") for: " & approval_dates & "###")
Call write_bullet_and_variable_in_CASE_NOTE("EA extended for 30 days from " & start_date & " through", end_date & ".")
Call write_variable_in_CASE_NOTE("* Client stay in shelter beyond the first 30 days.")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
