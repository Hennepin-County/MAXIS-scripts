'===========================================================================================STATS
name_of_script = "Quick Transfer.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 500
STATS_denominatinon = "C"
'===========================================================================================END OF STATS BLOCK

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
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
CALL navigate_to_MAXIS_screen ("STAT", "ADDR")
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 171, 70, "Transfer"
  EditBox 55, 5, 35, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    PushButton 110, 5, 50, 15, "Geocoder", Geo_coder_button
  EditBox 55, 25, 20, 15, spec_xfer_worker
  ButtonGroup ButtonPressed
    OkButton 65, 50, 45, 15
    CancelButton 115, 50, 45, 15
  Text 5, 30, 40, 10, "Transfer to:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 80, 30, 60, 10, " (last 3 digit of X#)"
EndDialog

'--------------------------------------------------------------------------------------------------script
DO 'Password DO loop
    DO  'External resource DO loop
       dialog Dialog1
       cancel_confirmation
       If ButtonPressed = Geo_coder_button then CreateObject("WScript.Shell").Run("https://hcgis.hennepin.us/agsinteractivegeocoder/default.aspx")
    Loop until ButtonPressed = -1
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = FALSE

'-------------------------------------------------------------------------------------Transfers the case to the assigned worker if this was selected in the second dialog box
'Determining if a case will be transferred or not. All cases will be transferred except addendum app types. THIS IS NOT CORRECT AND NEEDS TO BE DISCUSSED WITH QI
CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
EMWriteScreen "x", 7, 16
TRANSMIT
PF9
EMWriteScreen "X127" & spec_xfer_worker, 18, 61
TRANSMIT
EMReadScreen worker_check, 9, 24, 2
IF worker_check = "SERVICING" THEN
	action_completed = False
	PF10
END IF
EMReadScreen transfer_confirmation, 16, 24, 2
IF transfer_confirmation = "CASE XFER'D FROM" then
	action_completed = True
Else
	action_completed = False
End if

PF3

'CALL navigate_to_MAXIS_screen ("DAIL", "DAIL")
'EMWriteScreen "x127d5x", 21, 06
'TRANSMIT

script_end_procedure("CASE HAS BEEN UPDATED.")
