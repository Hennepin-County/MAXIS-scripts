'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - 12 MO CONTACT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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
call changelog_update("05/21/2021", "Added supports to pull the case number when using the script from the DAIL scrubber.", "Ilse Ferris, Hennepin County")
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'grabbing case number & connecting to MAXIS
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'Adding case number if using the script from the DAIL scrubber
If MAXIS_case_number = "" then
    EmReadscreen DAIL_panel, 4, 2, 48
    If DAIL_panel = "DAIL" then
        EmReadscreen MAXIS_case_number, 8, 5, 73
        MAXIS_case_number = trim(MAXIS_case_number)
    End if
End if

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, MAXIS_case_number
  Text 5, 25, 70, 10, "Sign your case note:"
  EditBox 80, 20, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

Do
	dialog Dialog1
	Cancel_without_confirmation
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'THE MEMO----------------------------------------------------------------------------------------------------
Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)	'navigates to spec/memo and opens into edit mode

Call write_variable_in_SPEC_MEMO("************************************************************")
Call write_variable_in_SPEC_MEMO("This notice is to remind you to report changes to your county worker by the 10th of the month following the month of the change. Changes that must be reported are address, people in your household, income, shelter costs and other changes such as legal obligation to pay child support. If you don't know whether to report a change, contact your county worker.")
CALL write_variable_in_SPEC_MEMO("")
CALL write_variable_in_SPEC_MEMO("*** Submitting Documents:")
CALL write_variable_in_SPEC_MEMO("- Online at infokeep.hennepin.us or MNBenefits.mn.gov")
CALL write_variable_in_SPEC_MEMO("  Use InfoKeep to upload documents directly to your case.")
CALL write_variable_in_SPEC_MEMO("- Mail, Fax, or Drop Boxes at Service Centers.")
CALL write_variable_in_SPEC_MEMO("  More Info: https://www.hennepin.us/economic-supports")
Call write_variable_in_SPEC_MEMO("************************************************************")
PF4

'THE CASE NOTE
Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("Sent 12 month contact letter via SPEC/MEMO on " & date)

Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
