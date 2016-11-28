'Required for statistical purposes===============================================================================
name_of_script = "UTILITIES - MOVE PRODUCTION SCREEN TO INQUIRY.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 40                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block==============================================================================================

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

EMConnect "A"
row = 1
col = 1
EMSearch "Function: ", row, col
If row = 0 then
  MsgBox "Function not found."
  StopScript
End if
EMReadScreen MAXIS_function, 4, row, col + 10
If MAXIS_function = "____" then
  MsgBox "Function not found."
  StopScript
End if

row = 1
col = 1
EMSearch "Case Nbr: ", row, col
If row = 0 then
  MsgBox "Case number not found."
  StopScript
End if
EMReadScreen MAXIS_case_number, 8, row, col + 10

row = 1
col = 1
EMSearch "Month: ", row, col
If row = 0 then
  MsgBox "Footer month not found."
  StopScript
End if
EMReadScreen MAXIS_footer_month, 2, row, col + 7
EMReadScreen MAXIS_footer_year, 2, row, col + 10

row = 1
col = 1
EMSearch "(", row, col
If row = 0 then
  MsgBox "Command not found."
  StopScript
End if
EMReadScreen MAXIS_command, 4, row, col + 1
If MAXIS_command = "NOTE" then MAXIS_function = "CASE"

EMConnect "B"
EMFocus

attn
EMReadScreen inquiry_check, 7, 7, 15
If inquiry_check <> "RUNNING" then
  MsgBox "Inquiry not found. The script will now stop."
  StopScript
End if

EMWriteScreen "FMPI", 2, 15
transmit

back_to_self

EMWriteScreen MAXIS_function, 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
EMWriteScreen MAXIS_command, 21, 70
transmit

script_end_procedure("")
