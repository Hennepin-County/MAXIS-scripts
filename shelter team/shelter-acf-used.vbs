'GATHERING STATS===========================================================================================
name_of_script = "SHELTER-ACF USED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 60
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
call changelog_update("08/01/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
 
CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"

BeginDialog ACF_used, 0, 0, 231, 90, "ACF Used for Shelter Stay"
  EditBox 110, 10, 50, 15, Shelter_stay_bgn
  EditBox 175, 10, 50, 15, Shelter_stay_end
  EditBox 75, 30, 50, 15, EA_avail_date
  EditBox 45, 50, 180, 15, Comments_notes
  ButtonGroup ButtonPressed
    OkButton 120, 70, 50, 15
    CancelButton 175, 70, 50, 15
  Text 5, 35, 70, 10, "EA will be available:"
  Text 165, 15, 10, 10, "to"
  Text 5, 15, 105, 10, "Use ACF for Shelter Stay Dates:"
  Text 5, 55, 40, 10, "Comments: "
EndDialog

Do
	Do
		err_msg = ""
		dialog ACF_used
		IF buttonpressed = 0 then stopscript 
		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
		IF isdate(Shelter_stay_bgn) = False then err_msg = err_msg & vbnewline & "* Enter a valid date of for the start of shelter stay."
		IF isdate(shelter_stay_end) = False then err_msg = err_msg & vbnewline & "* Enter a valid date of for the end of shelter stay."
        IF isdate(EA_avail_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid date of for EA availablity."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE ("### ACF Used for Shelter Stay " & Shelter_stay_bgn & "-" & Shelter_stay_end & " ###")
	Call write_bullet_and_variable_in_CASE_NOTE("EA will be available", EA_avail_date)
	Call write_bullet_and_variable_in_case_note("Comments", Comments_Notes)
	Call write_variable_in_CASE_NOTE ("-----")
	Call write_variable_in_CASE_NOTE ("Hennepin County Shelter Team")

script_end_procedure ("")
