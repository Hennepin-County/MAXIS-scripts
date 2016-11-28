'Required for statistical purposes===============================================================================
name_of_script = "DAIL - DISA MESSAGE.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 64          'manual run time in seconds
STATS_denomination = "C"       'C is for Case
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


'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.


EMConnect ""

EMSendKey "s"
transmit

EMSendKey "disa"
transmit

'HH member dialog to select who's job this is.
BeginDialog HH_memb_dialog, 0, 0, 191, 52, "HH member"
  EditBox 50, 25, 25, 15, HH_memb
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 5, 10, 125, 15, "Which HH member is this for? (ex: 01)"
EndDialog
HH_memb = "01"
dialog HH_memb_dialog
If ButtonPressed = 0 then stopscript

EMWriteScreen HH_memb, 20, 76
transmit

EMReadScreen cash_disa_status, 1, 11, 69
If cash_disa_status <> "1" then
  MsgBox "This type of DISA status is not yet supported. It could be a SMRT or some other type of verif needed. Process manually at this time."
  stopscript
End if

PF4

PF9

EMSendKey "<home>" + "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS" + "<newline>"
If cash_disa_status = 1 then EMSendKey "* Client needs a new Medical Opinion Form. Created using " & EDMS_choice & " and sent to client. TIKLed for 30-day return." & "<newline>"
EMSendKey "---" + "<newline>"

BeginDialog worker_sig_dialog, 0, 0, 191, 57, "Worker signature"
  EditBox 35, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 25, 10, 75, 10, "Sign your case note."
EndDialog

dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMSendKey worker_sig
PF3
PF3
PF3

EMSendKey "w"
transmit

'The following will generate a TIKL formatted date for 30 days from now.
TIKL_month = datepart("m", dateadd("d", 30, date))
If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month
TIKL_day = datepart("d", dateadd("d", 30, date))
If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day
TIKL_year = datepart("yyyy", dateadd("d", 30, date))
TIKL_year = TIKL_year - 2000

EMSetCursor 5, 18
EMSendKey TIKL_month & TIKL_day & TIKL_year
EMSetCursor 9, 3
EMSendKey "Medical Opinion Form sent 30 days ago. If not responded to, send another, and TIKL to close in 30 additional days."
transmit
PF3


MsgBox "Case note and TIKL made. Send a Medical Opinion Form using " & EDMS_choice & "."
script_end_procedure("")
