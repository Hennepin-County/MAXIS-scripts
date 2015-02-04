'GRABBING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - HC RENEWAL.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
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
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000


'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_and_footer_month_dialog, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog


BeginDialog HC_ER_dialog, 0, 0, 456, 300, "HC ER dialog"
  EditBox 75, 50, 50, 15, recert_datestamp
  DropListBox 185, 50, 75, 15, " "+chr(9)+"complete"+chr(9)+"incomplete", recert_status
  EditBox 325, 50, 125, 15, HH_comp
  EditBox 60, 70, 390, 15, earned_income
  EditBox 70, 90, 380, 15, unearned_income
  EditBox 40, 110, 410, 15, assets
  EditBox 60, 130, 95, 15, COEX_DCEX
  CheckBox 180, 135, 205, 10, "Check here if you used this HC ER as a SNAP CSR as well.", SNAP_CSR_check
  EditBox 100, 150, 350, 15, FIAT_reasons
  EditBox 50, 170, 400, 15, other_notes
  EditBox 45, 190, 405, 15, changes
  EditBox 60, 210, 390, 15, verifs_needed
  EditBox 55, 230, 395, 15, actions_taken
  EditBox 60, 260, 90, 15, MAEPD_premium
  CheckBox 10, 280, 65, 10, "Emailed MADE?", MADE_check
  EditBox 400, 250, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 345, 270, 50, 15
    CancelButton 400, 270, 50, 15
    PushButton 10, 20, 25, 10, "BUSI", BUSI_button
    PushButton 35, 20, 25, 10, "JOBS", JOBS_button
    PushButton 10, 30, 25, 10, "RBIC", RBIC_button
    PushButton 35, 30, 25, 10, "UNEA", UNEA_button
    PushButton 75, 20, 25, 10, "ACCT", ACCT_button
    PushButton 100, 20, 25, 10, "CARS", CARS_button
    PushButton 125, 20, 25, 10, "CASH", CASH_button
    PushButton 150, 20, 25, 10, "OTHR", OTHR_button
    PushButton 75, 30, 25, 10, "REST", REST_button
    PushButton 100, 30, 25, 10, "SECU", SECU_button
    PushButton 125, 30, 25, 10, "TRAN", TRAN_button
    PushButton 190, 20, 25, 10, "MEMB", MEMB_button
    PushButton 215, 20, 25, 10, "MEMI", MEMI_button
    PushButton 240, 20, 25, 10, "REVW", REVW_button
    PushButton 285, 20, 35, 10, "HC", ELIG_HC_button
    PushButton 340, 20, 45, 10, "prev. panel", prev_panel_button
    PushButton 340, 30, 45, 10, "next panel", next_panel_button
    PushButton 400, 20, 45, 10, "prev. memb", prev_memb_button
    PushButton 400, 30, 45, 10, "next memb", next_memb_button
    PushButton 5, 135, 25, 10, "COEX/", COEX_button
    PushButton 30, 135, 25, 10, "DCEX:", DCEX_button
    PushButton 85, 280, 65, 10, "SIR mail", SIR_mail_button
  GroupBox 5, 5, 60, 40, "Income panels"
  GroupBox 70, 5, 110, 40, "Asset panels"
  GroupBox 185, 5, 85, 30, "other STAT panels:"
  GroupBox 275, 5, 55, 30, "ELIG panels:"
  GroupBox 335, 5, 115, 40, "STAT-based navigation"
  Text 5, 55, 65, 10, "Recert datestamp:"
  Text 135, 55, 50, 10, "Recert status:"
  Text 280, 55, 35, 10, "HH comp:"
  Text 5, 75, 55, 10, "Earned income:"
  Text 5, 95, 65, 10, "Unearned income:"
  Text 5, 115, 30, 10, "Assets:"
  Text 5, 155, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 175, 45, 10, "Other notes:"
  Text 5, 195, 35, 10, "Changes?:"
  Text 5, 215, 50, 10, "Verifs needed:"
  Text 5, 235, 50, 10, "Actions taken:"
  GroupBox 5, 250, 150, 45, "If MA-EPD..."
  Text 10, 265, 50, 10, "New premium:"
  Text 335, 255, 65, 10, "Worker signature:"
EndDialog

BeginDialog case_note_dialog, 0, 0, 136, 51, "Case note dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 105, 10, "Yes, take me to case note.", yes_case_note_button
    PushButton 5, 35, 125, 10, "No, take me back to the script dialog.", no_case_note_button
  Text 10, 5, 125, 10, "Are you sure you want to case note?"
EndDialog

BeginDialog cancel_dialog, 0, 0, 141, 51, "Cancel dialog"
  Text 5, 5, 135, 10, "Are you sure you want to end this script?"
  ButtonGroup ButtonPressed
    PushButton 10, 20, 125, 10, "No, take me back to the script dialog.", no_cancel_button
    PushButton 20, 35, 105, 10, "Yes, close this script.", yes_cancel_button
EndDialog

'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col
HC_check = 1 'This is so the functions will work without having to select a program. It uses the same dialogs as the CSR, which can look in multiple places. This is HC only, so it doesn't need those.


'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Grabbing the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Grabbing the footer month/year
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

'Showing the case number dialog 
Do
  Dialog case_number_and_footer_month_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then call script_end_procedure("You are not in MAXIS or you are locked out of your case.")

'Navigating to STAT, checking for error prone
call navigate_to_screen("stat", "memb")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then call script_end_procedure("Can't get into STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background email your script administrator the case number and footer month.")
EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then transmit 'For error prone cases.

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling info from STAT
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "REVW", recert_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Creating variable for recert_month
recert_month = footer_month & "/" & footer_year

'Showing case note dialog, with navigation and required answers logic, and it navigates to the case note
Do
  Do
    Do
      Do
        Do
          Dialog HC_ER_dialog
          If ButtonPressed = 0 then 
            dialog cancel_dialog
            If ButtonPressed = yes_cancel_button then stopscript
          End if
        Loop until ButtonPressed <> no_cancel_button
        EMReadScreen STAT_check, 4, 20, 21
        If STAT_check = "STAT" then
          If ButtonPressed = prev_panel_button then call panel_navigation_prev
          If ButtonPressed = next_panel_button then call panel_navigation_next
          If ButtonPressed = prev_memb_button then call memb_navigation_prev
          If ButtonPressed = next_memb_button then call memb_navigation_next
        End if
        transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
        EMReadScreen MAXIS_check, 5, 1, 39
        If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
      Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
      If ButtonPressed = BUSI_button then call navigate_to_screen("stat", "BUSI")
      If ButtonPressed = JOBS_button then call navigate_to_screen("stat", "JOBS")
      If ButtonPressed = RBIC_button then call navigate_to_screen("stat", "RBIC")
      If ButtonPressed = UNEA_button then call navigate_to_screen("stat", "UNEA")
      If ButtonPressed = ACCT_button then call navigate_to_screen("stat", "ACCT")
      If ButtonPressed = CARS_button then call navigate_to_screen("stat", "CARS")
      If ButtonPressed = CASH_button then call navigate_to_screen("stat", "CASH")
      If ButtonPressed = OTHR_button then call navigate_to_screen("stat", "OTHR")
      If ButtonPressed = REST_button then call navigate_to_screen("stat", "REST")
      If ButtonPressed = SECU_button then call navigate_to_screen("stat", "SECU")
      If ButtonPressed = TRAN_button then call navigate_to_screen("stat", "TRAN")
      If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
      If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
      If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
      If ButtonPressed = COEX_button then call navigate_to_screen("stat", "COEX")
      If ButtonPressed = DCEX_button then call navigate_to_screen("stat", "DCEX")
      If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
      If ButtonPressed = SIR_mail_button then run "C:\Program Files\Internet Explorer\iexplore.exe https://www.dhssir.cty.dhs.state.mn.us/Pages/Default.aspx"
    Loop until ButtonPressed = -1
    If recert_status = " " or actions_taken = "" or recert_datestamp = "" or worker_signature = "" then MsgBox "You need to fill in the datestamp, recert status, and actions taken sections, as well as sign your case note. Check these items after pressing ''OK''."
  Loop until recert_status <> " " and actions_taken <> "" and recert_datestamp <> "" and worker_signature <> "" 
  If ButtonPressed = -1 then dialog case_note_dialog
  If buttonpressed = yes_case_note_button then
    call navigate_to_screen("case", "note")
    PF9
    EMReadScreen case_note_check, 17, 2, 33
    EMReadScreen mode_check, 1, 20, 09
    If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then MsgBox "The script can't open a case note. Are you in inquiry? Check MAXIS and try again."
  End if
Loop until case_note_check = "Case Notes (NOTE)" and mode_check = "A"

'The case note
EMSendKey "<home>" & "***" & recert_month & " HC ER received " & recert_datestamp & ": " & recert_status & "***" & "<newline>"
If SNAP_CSR_check = 1 then call write_new_line_in_case_note("* Used HC ER as SNAP CSR.")
If HH_comp <> "" then call write_editbox_in_case_note("HH comp", HH_comp, 6)
If earned_income <> "" then call write_editbox_in_case_note("Earned income", earned_income, 6)
If unearned_income <> "" then call write_editbox_in_case_note("Unearned income", unearned_income, 6)
If assets <> "" then call write_editbox_in_case_note("Assets", assets, 6)
If COEX_DCEX <> "" then call write_editbox_in_case_note("COEX/DCEX", COEX_DCEX, 6)
if FIAT_reasons <> "" then call write_editbox_in_case_note("FIAT reasons", FIAT_reasons, 6)
If other_notes <> "" then call write_editbox_in_case_note("Other notes", other_notes, 6)
If changes <> "" then call write_editbox_in_case_note("Changes", changes, 6)
if verifs_needed <> "" then call write_editbox_in_case_note("Verifs needed", verifs_needed, 6)
call write_editbox_in_case_note("Actions taken", actions_taken, 6)
call write_new_line_in_case_note("---")
If MAEPD_premium <> "" then call write_editbox_in_case_note("MA-EPD premium", MAEPD_premium, 6)
If MADE_check = 1 then call write_new_line_in_case_note("* Emailed MADE.")
If MAEPD_premium <> "" or MADE_check = 1 then call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

call script_end_procedure("")






