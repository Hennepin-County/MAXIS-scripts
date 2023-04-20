'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - PROJECT NOOB SCRIPT.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone



Dialog1 = "" 'blanking out dialog name
'Add dialog here: Add the dialog just before calling the dialog below unless you need it in the dialog due to using COMBO Boxes or other looping reasons. Blank out the dialog name with Dialog1 = "" before adding dialog.
'Shows dialog -----------------------------------------------------------------------------------------------------

' Add dialog to collect case number, footer month, and footer year. Include field validation.

BeginDialog Dialog1, 0, 0, 191, 105, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 80, 85, 50, 15
    CancelButton 140, 85, 50, 15
  Text 10, 10, 80, 15, "Enter the case number:"
  Text 10, 35, 95, 15, "Enter the footer month (MM):"
  Text 10, 55, 95, 15, "Enter the footer year (YY):"
  EditBox 95, 5, 40, 15, MAXIS_case_number
  EditBox 110, 30, 20, 15, footer_month
  EditBox 110, 50, 20, 15, footer_year
EndDialog


DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		'Add in all of your mandatory field handling from your dialog here.

		If footer_month < 1 OR footer_month > 12 THEN err_msg = err_msg & "* The footer month must be a 2-digit number between 01 and 12"
		' If footer_month_number + footer_month > = true AND footer_month > 12 THEN err_msg = err_msg & "cannot be more than 12"
		If IsNumeric(MAXIS_case_number) = false OR Len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* The case number must be numeric and 8 digits or less." 
		If IsNumeric(footer_month) = false OR Len(footer_month) <> 2 THEN err_msg = err_msg & vbNewLine & "* The footer month must be a 2 digit number. Be sure to include a 0 before single digit years." 
		If IsNumeric(footer_year) = false OR Len(footer_year) <> 2 THEN err_msg = err_msg & vbNewLine & "* The footer year must be a 2 digit number. Be sure to include a 0 before single digit years." 
		If err_msg <> "" THEN MsgBox "FORM ERROR(S)!" & vbNewLine & err_msg

	Loop UNTIL err_msg = ""

    'Add to all dialogs where you need to work within BLUEZONE
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'End dialog section-----------------------------------------------------------------------------------------------
PF3
PF3
PF3
EMWriteScreen "STAT", 16, 43
EMWriteScreen MAXIS_case_number, 18, 43
footer_month_year = footer_month & footer_year
EMWriteScreen footer_month_year, 20, 43
transmit
EMWriteScreen "JOBS", 20, 71
transmit






'code snippet example---------------------------------------------------------------------------------------------
'This is here to show you how we might use the advanced automation library to do something in MAXIS.
'Feel free to build from this or just take the parts that are helpful.

'We have now made sure we are at SELF in MAXIS

'now we are going to STAT/SUMM for a specific case
' EMWriteScreen "STAT", 16, 43				'writing the MAXIS function to enter in the correct place in MAXIS
' EMWriteScreen MAXIS_case_number, 18, 43		'entering  case number in the 'case number' line
' 'TODO - should I be concerned if there is already information on this line?
' EMWriteScreen "SUMM", 21, 70				'writing the MAXIS command to enter in the correct place in MAXIS

' transmit									'function to move in MAXIS

'TODO - how do I make sure that I actually got to STAT/SUMM

















'leave the case note open and in edit mode unless you have a business reason not to (BULK scripts, multiple case notes, etc.)

'End the script. Put any success messages in between the quotes
script_end_procedure("")
