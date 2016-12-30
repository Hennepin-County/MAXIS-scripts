'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LEP - SAVE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
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

'DIALOG PORTION----------------------------------------------------------------------------------------------------------------------------------------------

BeginDialog SAVE_dialog, 0, 0, 206, 355, "SAVE Dialog"
  EditBox 65, 5, 85, 15, MAXIS_case_number
  OptionGroup RadioGroup1
    RadioButton 10, 30, 45, 10, "SAVE 1", SAVE_1
    RadioButton 60, 30, 45, 10, "SAVE 2", SAVE_2
  EditBox 65, 55, 130, 15, current_status
  EditBox 80, 75, 115, 15, LPR_adjusted_from
  EditBox 60, 95, 135, 15, date_of_entry
  EditBox 70, 115, 125, 15, country_of_origin
  CheckBox 10, 135, 75, 10, "SAVE 2 requested?", SAVE_2_requested_check
  OptionGroup RadioGroup2
    RadioButton 15, 180, 35, 10, "No", not_sponsored
    RadioButton 15, 195, 75, 10, "Yes, sponsored by:", sponsored
  EditBox 95, 190, 100, 15, sponsor_name
  EditBox 95, 215, 100, 15, imig_doc_received
  EditBox 45, 235, 40, 15, exp_date
  CheckBox 10, 255, 170, 10, "TIKL out to re-request 90 days before expiration.", TIKL_check
  EditBox 60, 270, 135, 15, other_notes
  EditBox 65, 290, 130, 15, actions_taken
  EditBox 80, 310, 115, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 80, 335, 50, 15
    CancelButton 135, 335, 50, 15
  Text 10, 100, 50, 10, "Date of entry:"
  Text 10, 275, 45, 10, "Other Notes:"
  GroupBox 5, 155, 185, 55, "SAVE 2"
  Text 10, 165, 135, 10, "Sponsored on I-864 Affidavit of Support?"
  Text 10, 220, 85, 10, "Imigration doc received:"
  Text 10, 240, 35, 10, "Exp date:"
  Text 10, 315, 70, 10, "Sign your case note:"
  Text 10, 80, 65, 10, "LPR adjusted from:"
  Text 10, 120, 60, 10, "Country of origin:"
  Text 10, 60, 50, 10, "Current status:"
  GroupBox 5, 45, 185, 105, "SAVE 1"
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 295, 50, 10, "Actions Taken:"
EndDialog



'THE SCRIPT PORTION----------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

Call MAXIS_case_number_finder(MAXIS_case_number)      'finding case number

Call check_for_MAXIS(true)						'making sure that person is in MAXIS and logged in

Do
	err_msg = ""						'error message handling to keep dialog looping until completed correctly.
	Dialog SAVE_dialog
	cancel_confirmation
	If MAXIS_case_number = "" THEN err_msg = err_msg & "You must enter a Case number." & vbNewLine
	If TIKL_check = 1 and IsDate(exp_date) = False then err_msg = err_msg & "You must enter a proper date (MM/DD/YYYY) if you want the script to TIKL out." & vbNewLine
	If imig_doc_received = "" THEN err_msg = err_msg & "Please enter a immigration doc received." & vbNewLine
	If worker_sig = "" THEN err_msg = err_msg & "You must sign your case note." & vbNewLine
	If err_msg <> "" THEN msgbox err_msg
Loop until err_msg = ""

Call check_for_MAXIS(false)						'making sure that person is in MAXIS and logged in


'CASE NOTE PORTION----------------------------------------------------------------------------------------------------------------------------------------------
start_a_blank_case_note

IF SAVE_1 = 1 then 											'case notes the save 1 portion if that option is selected
	Call write_variable_in_CASE_NOTE("**SAVE 1**")
	Call write_bullet_and_variable_in_CASE_NOTE("Current status", current_status)
	Call write_bullet_and_variable_in_CASE_NOTE("LPR adjusted from", LPR_adjusted_from)
	Call write_bullet_and_variable_in_CASE_NOTE("Date of entry", date_of_entry)
	Call write_bullet_and_variable_in_CASE_NOTE("Country of origin", country_of_origin)
End If

IF SAVE_2 = 1 then 											'case notes the save 2 portion if that option is selected
	Call write_variable_in_CASE_NOTE("**SAVE 2**")
	If not_sponsored = 1 then Call write_variable_in_CASE_NOTE("* No sponsor indicated on SAVE.")
	If sponsored = 1 then Call write_variable_in_CASE_NOTE("* Client is sponsored. Sponsor is indicated as " & sponsor_name & ".")
End If

															'Case notes portion shared by SAVE 1 and SAVE 2
Call write_bullet_and_variable_in_CASE_NOTE("Immigration document received", imig_doc_received)
Call write_bullet_and_variable_in_CASE_NOTE("Expiration date", exp_date)
If TIKL_check = 1 then Call write_variable_in_CASE_NOTE("* TIKLed to re-request " & dateadd("d", -90, exp_date) & ".")  'subtracting 90 days from expiration date to write a TIKL to request updated information.
If SAVE_2_requested_check = checked then Call write_variable_in_CASE_NOTE("* SAVE 2 requested.")
Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_sig)

'TIKL PORTION----------------------------------------------------------------------------------------------------------------------------------------------

If TIKL_check = checked then
	Call navigate_to_MAXIS_screen("DAIL","WRIT")
	Call create_MAXIS_friendly_date(exp_date, -90, 5, 18)  'subtracting 90 days from expiration date to write a TIKL to request updated information.
	Call write_variable_in_TIKL("Check on immigration documentation. If it hasn't been updated, request updated info, as what we have expires " & exp_date & ". TIKL generated via script.")
	script_end_procedure("Success! TIKL sent for " & dateadd("d", -90, exp_date) & ", 90 days prior to document expiration.")   'subtracting 90 days from expiration date to write a TIKL to request updated information.
END IF

script_end_procedure("")
