'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-PERSONAL NEEDS.vbs"
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
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("09/21/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("10/31/2019", "Updated script as the footer month and year were having issues populating correctly. The script will now use current month plus one to determine footer month and year for the dialog title and case note.", "Casey Love, Hennepin County")
CALL changelog_update("07/29/2019", "Updated script per request. Removed ELIG vs INELIG and updated case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/14/2017", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 171, 125, "Personal Needs for " & CM_plus_1_mo & "/" & CM_plus_1_yr
  EditBox 65, 5, 40, 15, MAXIS_case_number
  EditBox 150, 5, 15, 15, HH_size
  EditBox 65, 25, 40, 15, amt_issued
  DropListBox 65, 45, 100, 15, "Select One:"+chr(9)+"CS"+chr(9)+"DWP"+chr(9)+"Earned Income"+chr(9)+"MFIP"+chr(9)+"Per Capita"+chr(9)+"RSDI"+chr(9)+"SSI"+chr(9)+"Other(please explain)", income_source
  EditBox 65, 65, 100, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 65, 105, 50, 15
    CancelButton 115, 105, 50, 15
  Text 5, 30, 55, 10, "Amount Eligible: "
  Text 5, 70, 45, 10, "Other Notes: "
  Text 120, 10, 30, 10, "HH Size: "
  Text 5, 50, 50, 10, "Income Source: "
  Text 5, 10, 50, 10, "Case Number: "
  Text 5, 90, 60, 10, "Worker signature:"
  EditBox 65, 85, 100, 15, worker_signature
EndDialog

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
        cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF HH_size = "" then err_msg = err_msg & vbNewLine & "* Please enter the HH size."
		IF amt_issued = "" then err_msg = err_msg & vbNewLine & "* Please enter the amount issued."
		IF income_source = "Select One:" then err_msg = err_msg & vbNewLine & "* Please select the client's source of income."
		IF income_source = "Other" Then
			If other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please complete the other notes section to explain the income source."
		END IF
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

amt_issued = "$" & amt_issued

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### Personal needs " & CM_plus_1_mo & "/" & CM_plus_1_yr & " ###")
Call write_bullet_and_variable_in_CASE_NOTE("HH size", HH_size)
Call write_bullet_and_variable_in_CASE_NOTE("Amount eligible", amt_issued)
Call write_bullet_and_variable_in_CASE_NOTE("Source of income", income_source)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure_with_error_report("Success! The case note has been entered.")
