'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - MAIN MENU.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS-------------------------------------------------------------------------------------------
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if
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
BeginDialog MEMOS_scripts_main_menu_dialog, 0, 0, 456, 175, "Memos scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 400, 155, 50, 15
    PushButton 375, 10, 65, 10, "SIR instructions", SIR_instructions_button
    PushButton 10, 25, 65, 10, "12 month contact", TWELVE_MONTH_CONTACT_button
    PushButton 10, 50, 65, 10, "Appointment letter", APPOINTMENT_LETTER_button
    PushButton 10, 65, 70, 10, "LTC - Asset transfer", LTC_ASSET_TRANSFER_button
    PushButton 10, 80, 60, 10, "MFIP orientation", MFIP_ORIENTATION_button
    PushButton 10, 95, 55, 10, "MNsure memo", MNSURE_MEMO_button
    PushButton 10, 110, 25, 10, "NOMI", NOMI_button
    PushButton 10, 125, 55, 10, "Overdue baby", OVERDUE_BABY_button
  Text 5, 5, 235, 10, "Memos scripts main menu: select the script to run from the choices below."
  Text 80, 25, 370, 20, "--- Sends a MEMO to the client reminding them of their reporting responsibilities (required for SNAP 2-year certification periods, per POLI/TEMP TE02.08.165)."
  Text 80, 50, 300, 10, "--- Sends a MEMO containing the appointment letter (with text from POLI/TEMP TE02.05.15)."
  Text 85, 65, 200, 10, "--- Sends a MEMO to a LTC client regarding asset transfers."
  Text 75, 80, 185, 10, "--- Sends a MEMO to a client regarding MFIP orientation."
  Text 70, 95, 160, 10, "--- Sends a MEMO to a client regarding MNsure."
  Text 40, 110, 375, 10, "--- Sends the SNAP notice of missed interview (NOMI) letter, following rules set out in POLI/TEMP TE02.05.15."
  Text 70, 125, 365, 20, "--- NEW 04/2015!!! Sends a MEMO informing client that they need to report information regarding the birth of their child, and/or pregnancy end date, within 10 days or their case may close."
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows main menu dialog, which asks user which memo to generate. Loops until a button other than the SIR instructions button is clicked.
Do
	dialog MEMOS_scripts_main_menu_dialog
	If buttonpressed = cancel then stopscript
	If buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/scriptwiki/Wiki%20Pages/Memos%20scripts.aspx")
Loop until buttonpressed <> SIR_instructions_button


'Connecting to BlueZone
EMConnect ""

'Hennepin handling (they don't use the Appt Letter or NOMI scripts because they have permission (at least temporarily) to schedule using a time range instead of a single time. Because of this, the NOMI and Appt letter scripts would technically cause incorrect information to be sent to the clients. This is a simple solution until their procedures are updated.)
If ucase(worker_county_code) = "X127" then
	IF ButtonPressed = APPOINTMENT_LETTER_button 	THEN script_end_procedure("The Appointment Letter script is not available to Hennepin users at this time. Contact an alpha user or your supervisor if you have questions.")
	IF ButtonPressed = NOMI_button 					THEN script_end_procedure("The NOMI script is not available to Hennepin users at this time. Contact an alpha user or your supervisor if you have questions.")
End if

IF ButtonPressed = TWELVE_MONTH_CONTACT_button 	THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - 12 MONTH CONTACT.vbs")
IF ButtonPressed = APPOINTMENT_LETTER_button 	THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - APPOINTMENT LETTER.vbs")
IF ButtonPressed = LTC_ASSET_TRANSFER_button 	THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - LTC - ASSET TRANSFER.vbs")
IF ButtonPressed = MFIP_ORIENTATION_button 		THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - MFIP ORIENTATION.vbs")
IF ButtonPressed = MNSURE_MEMO_button 			THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - MNSURE MEMO.vbs")
IF ButtonPressed = NOMI_button 					THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - NOMI.vbs")
IF ButtonPressed = OVERDUE_BABY_button			THEN CALL run_from_GitHub(script_repository & "MEMOS/MEMOS - OVERDUE BABY.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")