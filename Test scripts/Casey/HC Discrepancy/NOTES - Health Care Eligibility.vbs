'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - HC Eligibility.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 1                       'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
'END OF stats block==============================================================================================

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
call changelog_update("09/30/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS
EMConnect ""
'Finds the case number
call MAXIS_case_number_finder(MAXIS_case_number)

BeginDialog elig_dlg, 0, 0, 121, 70, "Case Information"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  EditBox 80, 30, 15, 15, start_mo
  EditBox 100, 30, 15, 15, start_yr
  ButtonGroup ButtonPressed
    OkButton 45, 50, 35, 15
    CancelButton 80, 50, 35, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 10, 35, 70, 10, "First month of action:"
EndDialog

Do
	Do
		'Adding err_msg handling
		err_msg = ""

        Dialog elig_dlg

        If len(MAXIS_case_number) > 7 Then err_msg = err_msg & vbNewLine & "* Review the case number, it appears to be too long."
        If trim(MAXIS_case_number) = "" Then err_msg = err_msg & vbNewLine & "* Enter a case number."
        If IsNumeric(MAXIS_case_number) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter a valid case number"
        If len(start_mo) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
        If len(start_yr) <> 2 Then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."

        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
Loop until are_we_passworded_out = false

MAXIS_case_number = trim(MAXIS_case_number)
MAXIS_footer_month = start_mo
MAXIS_footer_year = start_yr

Call navigate_to_MAXIS_screen("ELIG", "HC  ")

EMReadScreen hc_elig_check, 4, 3, 51
If hc_elig_check <> "HHMM" Then script_end_procedure("No HC ELIG results exist, resolve edits and approve new version and run the script again.")
EMWriteScreen approval_month, 20, 56            'Goes to the next month and checks that elig results exist
EMWriteScreen approval_year,  20, 59
transmit
If hc_elig_check <> "HHMM" Then script_end_procedure("No HC ELIG results exist, resolve edits and approve new version and run the script again.")

'Read for each person on HC in the start month and year - any approval done in the current day
'Identify elig or inelig to determine approval vs closure vs denial
'TODO figure out how approval/denial/closure look different'
'create dynamic dialog for EACH client and have it specific to the elig information found
'Look at the the following months to ensure nothing has changed.

'case note detail of approval
