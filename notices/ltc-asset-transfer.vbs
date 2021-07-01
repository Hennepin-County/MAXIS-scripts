'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - LTC - ASSET TRANSFER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 70                               'manual run time in seconds
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
CALL changelog_update("03/19/2018", "Updated text regarding client's name.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC function. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The script------------------------
'connecting to MAXIS
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 110, "LTC asset transfer dialog"
  EditBox 60, 5, 60, 15, MAXIS_case_number
  EditBox 35, 30, 85, 15, client
  EditBox 35, 50, 85, 15, spouse
  EditBox 70, 70, 50, 15, renewal_footer_month_year
  ButtonGroup LTC_asset_transfer_dialog_ButtonPressed
    OkButton 10, 90, 50, 15
    CancelButton 65, 90, 50, 15
  Text 5, 55, 30, 10, "Spouse:"
  Text 5, 75, 65, 10, "ER date (MM/YY):"
  Text 5, 35, 30, 10, "Client:"
  Text 5, 10, 45, 10, "Case number:"
EndDialog

Do
    Dialog Dialog1
    If LTC_asset_transfer_dialog_ButtonPressed = 0 then stopscript
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)	'navigates to spec/memo and opens into edit mode

Call write_variable_in_SPEC_MEMO("The ownership of " & client & "'s assets must be transferred to " & spouse & " to avoid having them counted in future eligibility determinations. You are encouraged to do this as soon as possible. This transfer of assets must be done before " & client & "'s first annual renewal for " & renewal_footer_month_year & ". Verification of the transfer can be provided at any time.")
Call write_variable_in_SPEC_MEMO("At the first annual renewal in " & renewal_footer_month_year & ", the value of all assets that list " & client & " as an owner or co-owner will be applied towards the Medical Assistance Asset limit of $3,000.00. If the total value of all countable assets for " & client & " is more than $3,000.00, Medical Assistance may be closed for " & renewal_footer_month_year & ".")

script_end_procedure("Please review your notice before pressing PF4 to send.")
