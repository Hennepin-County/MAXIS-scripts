'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
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
BeginDialog BULK_scripts_main_menu_dialog, 0, 0, 456, 325, "Bulk scripts main menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 400, 305, 50, 15
    PushButton 10, 50, 25, 10, "ACTV", ACTV_LIST_button
    PushButton 35, 50, 25, 10, "EOMC", EOMC_LIST_button
    PushButton 85, 50, 25, 10, "PND2", PND2_LIST_button
    PushButton 110, 50, 25, 10, "REVS", REVS_LIST_button
    PushButton 135, 50, 25, 10, "REVW", REVW_LIST_button
    PushButton 160, 50, 25, 10, "MFCM", MFCM_LIST_button
    PushButton 10, 80, 25, 10, "ARST", ARST_LIST_button
    PushButton 10, 95, 80, 10, "LTC-GRH list generator", LTC_GRH_LIST_GENERATOR_button
    PushButton 10, 120, 75, 10, "SWKR list generator", SWKR_LIST_GENERATOR_button
    PushButton 10, 155, 45, 10, "Bulk TIKLer", BULK_TIKLER_button
    PushButton 10, 170, 95, 10, "CASE/NOTE from Excel list", CASE_NOTE_FROM_EXCEL_LIST_button
    PushButton 10, 195, 70, 10, "CEI premium noter", CEI_PREMIUM_NOTER_button
    PushButton 10, 210, 55, 10, "INAC scrubber", INAC_SCRUBBER_button
    PushButton 10, 250, 55, 10, "Returned mail", RETURNED_MAIL_button
    PushButton 10, 275, 80, 10, "REVW/MONT closures", REVW_MONT_CLOSURES_button
  Text 5, 5, 235, 10, "Bulk scripts main menu: select the script to run from the choices below."
  GroupBox 5, 20, 215, 45, "Case lists"
  Text 10, 35, 205, 10, "Case list scripts pull a list of cases into an Excel spreadsheet."
  GroupBox 5, 70, 445, 65, "Other bulk lists"
  Text 40, 80, 250, 10, "--- Caseload stats by worker. Includes cash/SNAP/HC/emergency/GRH stats."
  Text 95, 95, 345, 20, "--- Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload (or group of caseloads)."
  Text 90, 120, 315, 10, "--- Creates a list of SWKRs assigned to the various cases in a caseload (or group of caseloads)."
  GroupBox 5, 140, 445, 160, "Other bulk scripts"
  Text 60, 155, 175, 10, "--- Creates the same TIKL on up to 60 cases at once."
  Text 110, 170, 335, 20, "--- Creates the same CASE/NOTE on potentially hundreds of cases listed on an Excel spreadsheet of your choice."
  Text 85, 195, 240, 10, "--- Case notes recurring CEI premiums on multiple cases simultaneously."
  Text 70, 210, 375, 35, "--- Checks all cases on REPT/INAC (in the month before the current footer month, or prior) for MMIS discrepancies, active claims, DAIL messages, and ABPS panels in need of update (for Good Cause status), and adds them to a Word document. After that, it case notes all of the cases without DAIL messages or MMIS discrepancies. If your agency uses a closed-file worker number, it will SPEC/XFER the cases from your number into that number."
  Text 70, 250, 375, 20, "--- Case notes that returned mail (without a forwarding address) was received for up to 60 cases simultaneously, and TIKLs for 10-day return of proofs."
  Text 95, 275, 350, 20, "--- Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER. Case notes ''last day of REIN'' as well as ''date case becomes an intake.''"
  ButtonGroup ButtonPressed
    PushButton 60, 50, 25, 10, "PND1", PND1_LIST_button
EndDialog




'VARIABLES TO DECLARE
all_case_numbers_array = " "					'Creating blank variable for the future array
call worker_county_code_determination(worker_county_code, two_digit_county_code)	'Determines worker county code
is_not_blank_excel_string = Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34)	'This is the string required to tell excel to ignore blank cells in a COUNTIFS function


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Shows report scanning dialog, which asks user which report to generate.
dialog BULK_scripts_main_menu_dialog
If buttonpressed = cancel then stopscript

'Connecting to BlueZone
EMConnect ""

If buttonpressed = ACTV_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - REPT-ACTV LIST.vbs")
If buttonpressed = ARST_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - REPT-ARST LIST.vbs")
If buttonpressed = EOMC_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - REPT-EOMC LIST.vbs")
If buttonpressed = PND1_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - REPT-PND1 LIST.vbs")
If buttonpressed = PND2_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - REPT-PND2 LIST.vbs")
If buttonpressed = REVS_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - REPT-REVS LIST.vbs")
If buttonpressed = REVW_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - REPT-REVW LIST.vbs")
If buttonpressed = MFCM_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - REPT-MFCM LIST.vbs")
If buttonpressed = LTC_GRH_LIST_GENERATOR_button then call run_from_GitHub(script_repository & "BULK/BULK - LTC-GRH LIST GENERATOR.vbs")
If buttonpressed = SWKR_LIST_GENERATOR_button then call run_from_GitHub(script_repository & "BULK/BULK - SWKR LIST GENERATOR.vbs")
If ButtonPressed = BULK_TIKLER_button then call run_from_GitHub(script_repository & "BULK/BULK - BULK TIKLER.vbs")
If ButtonPressed = CASE_NOTE_FROM_EXCEL_LIST_button then call run_from_GitHub(script_repository & "BULK/BULK - CASE NOTE FROM EXCEL LIST.vbs")
If ButtonPressed = CEI_PREMIUM_NOTER_button then call run_from_GitHub(script_repository & "BULK/BULK - CEI PREMIUM NOTER.vbs")
If ButtonPressed = INAC_SCRUBBER_button then call run_from_GitHub(script_repository & "BULK/BULK - INAC SCRUBBER.vbs")
If ButtonPressed = RETURNED_MAIL_button then call run_from_GitHub(script_repository & "BULK/BULK - RETURNED MAIL.vbs")
If ButtonPressed = REVW_MONT_CLOSURES_button then call run_from_GitHub(script_repository & "BULK/BULK - REVW-MONT CLOSURES.vbs")

'Logging usage stats
script_end_procedure("If you see this, it's because you clicked a button that, for some reason, does not have an outcome in the script. Contact your alpha user to report this bug. Thank you!")
