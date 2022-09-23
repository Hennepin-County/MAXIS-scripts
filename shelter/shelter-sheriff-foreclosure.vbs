'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-SHERIFF FORECLOSURE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
CALL changelog_update("09/21/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("06/19/2017", "Initial version.", "MiKayla Handley")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'autofilling the review_date variable with the current date
date_checked = date & ""
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 286, 145, "Sheriff Foreclosure"
  EditBox 55, 5, 55, 15, MAXIS_case_number
  EditBox 230, 5, 45, 15, date_checked
  EditBox 70, 25, 210, 15, property_address
  EditBox 70, 45, 100, 15, owner_name
  EditBox 235, 45, 45, 15, foreclosure_date
  EditBox 70, 65, 100, 15, occupant_name
  EditBox 95, 85, 185, 15, occupants_whereabouts
  EditBox 50, 105, 230, 15, other_notes
  EditBox 70, 125, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 125, 50, 15
    CancelButton 230, 125, 50, 15
  Text 5, 50, 55, 10, "Owner(s) name:"
  Text 5, 70, 60, 10, "Occupant(s) name:"
  Text 120, 10, 80, 10, "Date of property review:"
  Text 5, 30, 60, 10, "Property address:"
  Text 5, 110, 40, 10, "Other notes: "
  Text 175, 50, 55, 10, "Foreclosure  date:"
  Text 5, 90, 85, 10, "Occupant(s) whereabouts:"
  Text 5, 10, 45, 10, "Case number:"
  Text 5, 130, 60, 10, "Worker Signature:"
EndDialog

'commented out the foreclosure_date test at reqwust of hennepin shelter Team'
DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If IsDate(date_checked) = False then err_msg = err_msg & vbNewLine & "* Enter the property review date."
		If property_address = "" then err_msg = err_msg & vbNewLine & "* Enter the property address."
		If owner_name = "" then err_msg = err_msg & vbNewLine & "* Enter the property owner's name."
		If occupant_name = "" then err_msg = err_msg & vbNewLine & "* Enter the occupant's name."
		If occupants_whereabouts = "" then err_msg = err_msg & vbNewLine & "* Enter the occupant's current whereabouts."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'adding the case number
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("### Sheriff foreclosure website checked on: " & date_checked & " ###"   )
Call write_bullet_and_variable_in_CASE_NOTE("Property address", property_address)
Call write_bullet_and_variable_in_CASE_NOTE("Owner(s) name", owner_name)
Call write_bullet_and_variable_in_CASE_NOTE("Foreclosure date", foreclosure_date)
Call write_bullet_and_variable_in_CASE_NOTE("Representative name", rep_name)
Call write_bullet_and_variable_in_CASE_NOTE("Occupant(s) name", occupant_name)
Call write_bullet_and_variable_in_CASE_NOTE("Occupant(s) current whereabouts", occupants_whereabouts)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
