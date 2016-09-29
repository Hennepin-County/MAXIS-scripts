'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - BILS UPDATER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 30                	'manual run time in seconds
STATS_denomination = "I"       		'I is for each ITEM
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog EMPS_case_number_dialog, 0, 0, 136, 80, "EMPS case number dialog"
  EditBox 60, 5, 70, 15, MAXIS_case_number
  CheckBox 5, 25, 130, 10, "Check here to update existing EMPS.", updating_existing_EMPS_check
  EditBox 85, 40, 15, 15, footer_month
  EditBox 100, 40, 15, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 30, 60, 50, 15
    CancelButton 80, 60, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 45, 75, 10, "Footer month and year"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""

'Run the case number dialog

'Determine what programs the case is active or pending

'IF SNAP only - error out as case is not MFIP/DWP or pending

'Get client age and children's ages 

'Check EMPS for errors (Fin Orient Date missine, ES Ref date missing for active cases. ES option missing for 18/19 yr old)

'Ask worker what process they  want to update 
'Child under 1 - begin, end, get MFIP results
'enter or end sanction
''

script_end_procedure("")