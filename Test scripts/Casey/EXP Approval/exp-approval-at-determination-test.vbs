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


'FUNCTIONS

'gather_maxis_detail_about_expedited
	'Read MAXIS panels to autofill the asset, income, and expense information in determining expedited

'ask_initial_expedited_questions_dlg
	'The dialog code of the initial questions to move to the next steps of expedited determination. This one may not be necessary in some instances (ie Interview)

'ask_income_detail_questions_dlg
	'The dialog code of the income details for Expedited Specifically

'ask_asset_detail_questions_dlg
	'The dialog code of the asset details for Expedited Specifically

'ask_housing_expense_detail_questions_dlg
	'The dialog code of the housing expense details for Expedited Specifically

'ask_utility_detail_questions_dlg
	'The dialog code of the utility details for Expedited Specifically

'Update_income_panels_for_expedited
	'MAXIS coding in the income panels for expedited. We will used EIB functionality for when verifications are received

'Update_asset_panels_for_expedited
	'MAXIS coding in the asset panels for expedited. possibly change this to a regular 'access' functionality for ACCT and CASH

'read_elig_information_for_SNAP
	'Used to find ELIG and pull the information

'display_snap_elig_dlg
	'The dialog code to view the details of elig

'NAVIGATE FUNCTIONS FOR ALL THE BUTTONS AND ERR HANDLING
'navigate_initial_expedited_questions
'navigate_income_expedited_questions
'navigate_asset_expedited_questions
'navigate_housing_expedited_questions
'navigate_utility_expedited_questions
'navigate_snap_elig_info
'

'Update the functions to UPDATE SHEL, HEST to have specific expedited coding
