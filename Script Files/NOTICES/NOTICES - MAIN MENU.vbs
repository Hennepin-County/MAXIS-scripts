'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - MAIN MENU.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
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
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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


'A dynamic dialog is created below using a function. This is so we don't waste spaces towards the dialog limit looping a menu. 
'You must DIM any button you add to the menus.
'If you add a new section/menu you will need to add a new ELSE IF statement for that menu.

'dimming all buttons
DIM ButtonPressed
DIM SIR_instructions_button, dialog_name, back_button, SNAP_WCOM_button
DIM TWELVE_MONTH_CONTACT_button, APPOINTMENT_LETTER_button, CS_DISREGARD_button, GRH_OP_CL_LEFT_FACI_button
DIM METHOD_B_WCOM_button, LTC_ASSET_TRANSFER_button, MAEPD_NO_PREMIUM_button, MFIP_ORIENTATION_button, MNSURE_MEMO_button, NOMI_button, OVERDUE_BABY_button, SNAP_E_AND_T_LETTER_button
DIM ABAWD_WITH_CHILD_IN_HH_WCOM_button, DUPLICATE_ASSISTANCE_button, POSTPONED_WREG_button 

'function to dynamically create dialog based on what is stored in the variable entered into it. 
FUNCTION create_NOTICES_main_menu(dialog_name)
	IF dialog_name = "HOME MENU" Then
		BeginDialog dialog_name, 0, 0, 451, 285, "Memos scripts main menu dialog"
		ButtonGroup ButtonPressed
			PushButton 15, 30, 65, 15, "SNAP WCOMS", SNAP_WCOM_button
			PushButton 5, 65, 65, 10, "12 month contact", TWELVE_MONTH_CONTACT_button
			PushButton 5, 85, 65, 10, "Appointment letter", APPOINTMENT_LETTER_button
			PushButton 5, 100, 110, 10, "DWP/MFIP CS Disregard WCOM", CS_DISREGARD_button
			PushButton 5, 115, 125, 10, "GRH overpayment (client left facility)", GRH_OP_CL_LEFT_FACI_button
			PushButton 5, 130, 60, 10, "Method B WCOM", METHOD_B_WCOM_button
			PushButton 5, 145, 70, 10, "LTC - Asset transfer", LTC_ASSET_TRANSFER_button
			PushButton 5, 160, 115, 10, "MAEPD - No initial premium paid", MAEPD_NO_PREMIUM_button
			PushButton 5, 175, 60, 10, "MFIP orientation", MFIP_ORIENTATION_button
			PushButton 5, 190, 55, 10, "MNsure memo", MNSURE_MEMO_button
			PushButton 5, 205, 25, 10, "NOMI", NOMI_button
			PushButton 5, 220, 55, 10, "Overdue baby", OVERDUE_BABY_button
			PushButton 5, 245, 70, 10, "SNAP E and T letter", SNAP_E_AND_T_LETTER_button
			PushButton 375, 5, 65, 10, "SIR instructions", SIR_instructions_button
			CancelButton 395, 265, 50, 15
		Text 75, 65, 375, 20, "--- Sends a MEMO to the client reminding them of their reporting responsibilities (required for SNAP 2-year certification periods, per POLI/TEMP TE02.08.165)."
		Text 75, 85, 300, 10, "--- Sends a MEMO containing the appointment letter (with text from POLI/TEMP TE02.05.15)."
		Text 120, 100, 320, 10, "--- NEW 01/2016!! Adds required WCOM to a notice when applying the CS Disregard to DWP/MFIP."
		Text 135, 115, 310, 10, "--- Sends a MEMO to a facility indicating that an overpayment is due because a client left."
		Text 70, 130, 360, 10, "--- NEW 01/2016!!! Makes detailed WCOM regarding spenddown vs. recipient amount for method B HC cases."
		Text 80, 145, 200, 10, "--- Sends a MEMO to a LTC client regarding asset transfers."
		Text 130, 160, 225, 10, "--- Sends a WCOM on a denial for no initial MA-EPD premium."
		Text 70, 175, 185, 10, "--- Sends a MEMO to a client regarding MFIP orientation."
		Text 65, 190, 160, 10, "--- Sends a MEMO to a client regarding MNsure."
		Text 35, 205, 375, 10, "--- Sends the SNAP notice of missed interview (NOMI) letter, following rules set out in POLI/TEMP TE02.05.15."
		Text 65, 220, 355, 20, "--- Sends a MEMO informing client that they need to report information regarding the birth of their child, and/or pregnancy end date, within 10 days or their case may close."
		Text 80, 245, 315, 10, "--- Sends a SPEC/LETR informing client that they have an Employment and Training appointment."
		Text 5, 5, 255, 10, "Memos scripts main menu: select the script to run from the choices below."
		GroupBox 5, 20, 265, 30, "Other Sections"
		EndDialog
	Else IF dialog_name = "SNAP WCOM MENU" THEN	
		BeginDialog dialog_name, 0, 0, 451, 150, "SNAP Related WCOM Scripts"
		ButtonGroup ButtonPressed
			PushButton 15, 30, 50, 15, "Back", back_button
			PushButton 5, 55, 115, 10, "ABAWD with child in HH WCOM", ABAWD_WITH_CHILD_IN_HH_WCOM_button
			PushButton 5, 70, 100, 10, "Duplicate assistance WCOM", DUPLICATE_ASSISTANCE_button
			PushButton 5, 85, 80, 10, "Postponed WREG verif", POSTPONED_WREG_button
			PushButton 375, 5, 65, 10, "SIR instructions", SIR_instructions_button
			CancelButton 380, 130, 50, 15
		Text 125, 55, 320, 10, "---NEW 01/2016 Adds a WCOM to a notice for an ABAWD adult receiving child under 18 exemption."
		Text 110, 70, 305, 10, "--- Adds a WCOM to a notice for duplicate assistance explaining why the client was ineligible."
		Text 95, 85, 345, 20, "--- Sends a WCOM informing the client of postponed verifications that MAXIS won't add to notice correctly by itself."
		GroupBox 5, 20, 265, 30, "Other Sections"
		Text 5, 5, 270, 10, "SNAP WCOM scripts main menu: select the script to run from the choices below."
		EndDialog

		END IF
	END IF
	'calls dialog inside of function so it can be created dynamically based on the variable placed into it. 
	DIALOG dialog_name
END Function

'Variables to declare
IF script_repository = "" THEN script_repository = "https://raw.githubusercontent.com/MN-Script-Team/DHS-MAXIS-Scripts/master/Script Files"		'If it's blank, we're assuming the user is a scriptwriter, ergo, master branch.

'THE SCRIPT----------------------------------------------------------------------------------------------------
'setting default dialot as home dialog
dialog_name = "HOME MENU"

Do
	'Calling the function that loads the dialogs
	CALL create_NOTICES_main_menu(dialog_name)
		IF ButtonPressed = 0 THEN stopscript
		'Opening the SIR Instructions
		IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Notices%20scripts.aspx")

		'If the user selects the other sub-menu, the script do-loops with the new dialog_name
		'Here is where you can add additional menus with more else if statements. 
		IF ButtonPressed = SNAP_WCOM_button THEN
			dialog_name = "SNAP WCOM MENU"
		ELSEIF ButtonPressed = back_button THEN
			dialog_name = "HOME MENU"
		END IF

		'If the user selects a script button, the script will exit the do-loop, you must add any menu navigation buttons you add to this statement. 
Loop until buttonpressed <> SIR_instructions_button AND buttonpressed <> SNAP_WCOM_button AND buttonpressed <> back_button

'Connecting to BlueZone
EMConnect ""

'Hennepin handling (they don't use the Appt Letter script because they have permission to schedule using a time range instead of a single time. Because of this, the Appt letter script would technically cause incorrect information to be sent to the clients. This is a simple solution until the script is updated to allow for time ranges.)
If ucase(worker_county_code) = "X127" then
	IF ButtonPressed = APPOINTMENT_LETTER_button 	THEN script_end_procedure("The Appointment Letter script is not available to Hennepin users at this time. Contact an alpha user or your supervisor if you have questions.")
End if

If ButtonPressed = TWELVE_MONTH_CONTACT_button				then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - 12 MONTH CONTACT.vbs")
If ButtonPressed = ABAWD_WITH_CHILD_IN_HH_WCOM_button		then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - ABAWD WITH CHILD IN HH WCOM.vbs")
If ButtonPressed = APPOINTMENT_LETTER_button				then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - APPOINTMENT LETTER.vbs")
If ButtonPressed = DUPLICATE_ASSISTANCE_button				then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - DUPLICATE ASSISTANCE WCOM.vbs")
If ButtonPressed = CS_DISREGARD_button						then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - CS DISREGARD WCOM.vbs")
If ButtonPressed = GRH_OP_CL_LEFT_FACI_button				then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - GRH OP CL LEFT FACI.vbs")
If ButtonPressed = LTC_ASSET_TRANSFER_button				then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - LTC - ASSET TRANSFER.vbs")
If ButtonPressed = MAEPD_NO_PREMIUM_button					then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - MA-EPD NO INITIAL PREMIUM.vbs")
If ButtonPressed = METHOD_B_WCOM_button						then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - METHOD B WCOM.vbs")
If ButtonPressed = MFIP_ORIENTATION_button					then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - MFIP ORIENTATION.vbs")
If ButtonPressed = MNSURE_MEMO_button						then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - MNSURE MEMO.vbs")
If ButtonPressed = NOMI_button								then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - NOMI.vbs")
If ButtonPressed = OVERDUE_BABY_button						then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - OVERDUE BABY.vbs")
If ButtonPressed = POSTPONED_WREG_button					then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - POSTPONED WREG VERIFS.vbs")
If ButtonPressed = SNAP_E_AND_T_LETTER_button				then call run_from_GitHub(script_repository & "/NOTICES/NOTICES - SNAP E AND T LETTER.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
