'LOADING ROUTINE FUNCTIONS-------------------------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER FUNCTIONS LIBRARY.vbs"
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF


'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog MEMOS_scripts_main_menu_dialog, 0, 0, 456, 145, "Memos scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 400, 125, 50, 15
    PushButton 10, 25, 65, 10, "12 month contact", TWELVE_MONTH_CONTACT_button
    PushButton 10, 50, 65, 10, "Appointment letter", APPOINTMENT_LETTER_button
    PushButton 10, 65, 70, 10, "LTC - Asset transfer", LTC_ASSET_TRANSFER_button
    PushButton 10, 80, 60, 10, "MFIP orientation", MFIP_ORIENTATION_button
    PushButton 10, 95, 55, 10, "MNsure memo", MNSURE_MEMO_button
    PushButton 10, 110, 25, 10, "NOMI", NOMI_button
  Text 5, 5, 235, 10, "Memos scripts main menu: select the script to run from the choices below."
  Text 80, 25, 370, 20, "--- Sends a MEMO to the client reminding them of their reporting responsibilities (required for SNAP 2-year certification periods, per POLI/TEMP TE02.08.165)."
  Text 80, 50, 300, 10, "--- Sends a MEMO containing the appointment letter (with text from POLI/TEMP TE02.05.15)."
  Text 85, 65, 200, 10, "--- Sends a MEMO to a LTC client regarding asset transfers."
  Text 75, 80, 185, 10, "--- Sends a MEMO to a client regarding MFIP orientation."
  Text 70, 95, 160, 10, "--- Sends a MEMO to a client regarding MNsure."
  Text 40, 110, 375, 10, "--- Sends the SNAP notice of missed interview (NOMI) letter, following rules set out in POLI/TEMP TE02.05.15."
EndDialog




'VARIABLES TO DECLARE
all_case_numbers_array = " "					'Creating blank variable for the future array
call worker_county_code_determination(worker_county_code, two_digit_county_code)	'Determines worker county code
is_not_blank_excel_string = Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34)	'This is the string required to tell excel to ignore blank cells in a COUNTIFS function


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows report scanning dialog, which asks user which report to generate.
dialog MEMOS_scripts_main_menu_dialog
If buttonpressed = cancel then stopscript

'Connecting to BlueZone
EMConnect ""

IF ButtonPressed = TWELVE_MONTH_CONTACT_button 	THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - 12 MONTH CONTACT.vbs")
IF ButtonPressed = APPOINTMENT_LETTER_button 	THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - APPOINTMENT LETTER.vbs")
IF ButtonPressed = LTC_ASSET_TRANSFER_button 	THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - LTC - ASSET TRANSFER.vbs")
IF ButtonPressed = MFIP_ORIENTATION_button 		THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - MFIP ORIENTATION.vbs")
IF ButtonPressed = MNSURE_MEMO_button 			THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - MNSURE MEMO.vbs")
IF ButtonPressed = NOMI_button 					THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - NOMI.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")