'Required for statistical purposes===============================================================================
name_of_script = "DAIL - FMED DEDUCTION.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 127         'manual run time in seconds
STATS_denomination = "C"       'C is for case
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

'<<<<<GO THROUGH THE SCRIPT AND REMOVE REDUNDANT FUNCTIONS, THANKS TO CUSTOM FUNCTIONS THEY ARE NOT REQUIRED.

EMConnect ""

BeginDialog worker_sig_dialog, 0, 0, 141, 46, "Worker signature"
  EditBox 15, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 85, 5, 50, 15
    CancelButton 85, 25, 50, 15
  Text 5, 10, 75, 10, "Sign your case note."
EndDialog

Dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMWriteScreen "p", 6, 3
EMSendKey "<enter>"
EMWaitReady 0, 0

EMWriteScreen "memo", 20, 70
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "<PF5>"
EMWaitReady 0, 0

EMWriteScreen "x", 5, 10
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "You are turning 60 next month, so you may be eligible for a new deduction for SNAP." + "<newline>" + "<newline>"
EMSendKey "Clients who are over 60 years old may receive increased SNAP benefits if they have recurring medical bills over $35 each month." + "<newline>" + "<newline>"
EMSendKey "If you have medical bills over $35 each month, please contact your worker to discuss adjusting your benefits. You will need to send in proof of the medical bills, such as pharmacy receipts, an explanation of benefits, or premium notices." + "<newline>" + "<newline>"
EMSendKey "Please call your worker with questions."
EMSendKey "<PF4>"
EMWaitReady 0, 0

EMWriteScreen "case", 19, 22
EMWriteScreen "note", 19, 70
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "<PF9>"
EMWaitReady 0, 0

EMSendKey "MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION" + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey "* Sent MEMO to client about FMED deductions." + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_sig + ", using automated script."

EMSendKey "<PF3>"
EMWaitReady 0, 0

EMSendKey "<PF3>"
EMWaitReady 0, 0

MsgBox "The script has sent a MEMO to the client about the possible FMED deduction, and case noted the action."

script_end_procedure("")
