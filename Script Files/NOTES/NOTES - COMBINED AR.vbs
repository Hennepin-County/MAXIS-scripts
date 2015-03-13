'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - COMBINED AR.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
If beta_agency = "" or beta_agency = True then
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
Else
	url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
End if
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
BeginDialog case_number_dialog, 0, 0, 181, 120, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  CheckBox 10, 80, 30, 10, "GRH", GRH_check
  CheckBox 50, 80, 30, 10, "MSA", cash_check
  CheckBox 95, 80, 35, 10, "SNAP", SNAP_check
  CheckBox 145, 80, 30, 10, "HC", HC_check
  ButtonGroup ButtonPressed
    OkButton 35, 100, 50, 15
    CancelButton 95, 100, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 65, 170, 30, "Programs recertifying"
EndDialog

BeginDialog Combined_AR_dialog, 0, 0, 441, 335, "Combined AR dialog"
  EditBox 70, 35, 50, 15, recert_datestamp
  EditBox 230, 35, 40, 15, recert_month
  EditBox 60, 55, 50, 15, interview_date
  EditBox 45, 75, 165, 15, HH_comp
  EditBox 265, 75, 170, 15, US_citizen
  EditBox 35, 95, 210, 15, AREP
  EditBox 40, 155, 395, 15, income
  EditBox 35, 175, 400, 15, assets
  EditBox 65, 195, 370, 15, SHEL
  EditBox 100, 215, 335, 15, FIAT_reasons
  EditBox 60, 235, 375, 15, verifs_needed
  EditBox 55, 255, 380, 15, actions_taken
  EditBox 50, 275, 385, 15, other_notes
  CheckBox 5, 300, 65, 10, "R/R explained?", R_R_explained
  DropListBox 135, 295, 60, 15, ""+chr(9)+"complete"+chr(9)+"incomplete", review_status
  EditBox 275, 295, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 335, 315, 50, 15
    CancelButton 385, 315, 50, 15
    PushButton 20, 15, 25, 10, "HCRE", HCRE_button
    PushButton 45, 15, 25, 10, "MEMB", MEMB_button
    PushButton 70, 15, 25, 10, "MEMI", MEMI_button
    PushButton 95, 15, 25, 10, "REVW", REVW_button
    PushButton 185, 15, 20, 10, "FS", ELIG_FS_button
    PushButton 205, 15, 20, 10, "GA", ELIG_GA_button
    PushButton 225, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 245, 15, 20, 10, "MSA", ELIG_MSA_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 390, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 390, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 100, 25, 10, "AREP:", AREP_button
    PushButton 10, 130, 25, 10, "BUSI", BUSI_button
    PushButton 35, 130, 25, 10, "JOBS", JOBS_button
    PushButton 60, 130, 25, 10, "RBIC", RBIC_button
    PushButton 85, 130, 25, 10, "UNEA", UNEA_button
    PushButton 125, 130, 25, 10, "ACCT", ACCT_button
    PushButton 150, 130, 25, 10, "CARS", CARS_button
    PushButton 175, 130, 25, 10, "CASH", CASH_button
    PushButton 200, 130, 25, 10, "OTHR", OTHR_button
    PushButton 225, 130, 25, 10, "REST", REST_button
    PushButton 250, 130, 25, 10, "SECU", SECU_button
    PushButton 275, 130, 25, 10, "TRAN", TRAN_button
    PushButton 5, 200, 25, 10, "SHEL/", SHEL_button
    PushButton 30, 200, 25, 10, "HEST", HEST_button
  Text 5, 40, 65, 10, "Recert datestamp:"
  Text 160, 40, 70, 10, "Recert footer month:"
  Text 5, 80, 40, 10, "HH Comp:"
  Text 220, 80, 40, 10, "US citizen?:"
  GroupBox 5, 120, 110, 25, "Income panels"
  GroupBox 120, 120, 185, 25, "Asset panels"
  Text 5, 160, 30, 10, "Income:"
  Text 5, 180, 25, 10, "Assets:"
  Text 5, 220, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 240, 50, 10, "Verifs needed:"
  Text 5, 260, 50, 10, "Actions taken:"
  Text 5, 280, 40, 10, "Other notes:"
  Text 80, 300, 50, 10, "Review status:"
  Text 210, 300, 65, 10, "Sign the case note:"
  GroupBox 180, 5, 90, 25, "ELIG panels:"
  GroupBox 15, 5, 110, 25, "STAT panels:"
  GroupBox 330, 5, 110, 35, "STAT-based navigation"
  Text 5, 60, 55, 10, "Interview Date:"
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

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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

'Shows case number dialog
Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checks for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then call script_end_procedure("You are not in MAXIS, or you are locked out of your case.")

'Navigates to STAT
call navigate_to_screen("STAT", "REVW")

'Checks for error prone, and moves past it
ERRR_screen_check

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofill info
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", income)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMI", US_citizen)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "REVW", recert_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", income)
CALL autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL)
CALL autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL)

'Determines recert month
recert_month = footer_month & "/" & footer_year


'Showing the case note dialog
Do
  Do
    Do
      Do
        Do
          Dialog combined_AR_dialog
            If ButtonPressed = 0 then 
              dialog cancel_dialog
              If ButtonPressed = yes_cancel_button then stopscript
            End if
          Loop until ButtonPressed <> no_cancel_button
        EMReadScreen STAT_check, 4, 20, 21
        If STAT_check = "STAT" then
          If ButtonPressed = prev_panel_button then call prev_panel_navigation
          If ButtonPressed = next_panel_button then call next_panel_navigation
          If ButtonPressed = prev_memb_button then call prev_memb_navigation
          If ButtonPressed = next_memb_button then call next_memb_navigation
        End if
        transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
        EMReadScreen MAXIS_check, 5, 1, 39
        If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You do not appear to be in MAXIS. Are you passworded out? Or in MMIS? Check these and try again."
      Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS " 
      If ButtonPressed = AREP_button then call navigate_to_screen("stat", "arep")
      If ButtonPressed = FACI_button then call navigate_to_screen("stat", "FACI")
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
      If ButtonPressed = HCRE_button then call navigate_to_screen("stat", "HCRE")
      If ButtonPressed = REVW_button then call navigate_to_screen("stat", "REVW")
      If ButtonPressed = MEMB_button then call navigate_to_screen("stat", "MEMB")
      If ButtonPressed = MEMI_button then call navigate_to_screen("stat", "MEMI")
	  IF ButtonPressed = SHEL_button THEN CALL navigate_to_screen("STAT", "SHEL")
	  IF ButtonPressed = HEST_button THEN CALL navigate_to_screen("STAT", "HEST")
      If ButtonPressed = ELIG_HC_button then call navigate_to_screen("elig", "HC__")
      If ButtonPressed = ELIG_FS_button then call navigate_to_screen("elig", "FS__")
      If ButtonPressed = ELIG_GA_button then call navigate_to_screen("elig", "GA__")
      If ButtonPressed = ELIG_MSA_button then call navigate_to_screen("elig", "MSA_")
    Loop until ButtonPressed = -1
    If worker_signature = "" or review_status = " " or actions_taken = "" or recert_datestamp = "" then MsgBox "You must sign your case note and update the datestamp, actions taken, and review status sections."
  Loop until worker_signature <> "" and review_status <> " " and actions_taken <> "" and recert_datestamp <> ""
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
CALL write_variable_in_case_note("***Combined AR received " & recert_datestamp & " for " & recert_month & ": " & review_status & "***")
CALL write_bullet_and_variable_in_case_note("Interview Date", interview_date)
If HH_comp <> "" then call write_bullet_and_variable_in_case_note("HH comp", HH_comp)
If US_citizen <> "" then call write_bullet_and_variable_in_case_note("Citizenship", US_citizen)
If AREP <> "" then call write_bullet_and_variable_in_case_note("AREP", AREP)
If FACI <> "" then call write_bullet_and_variable_in_case_note("FACI", FACI)
If income <> "" then call write_bullet_and_variable_in_case_note("Income", income)
If assets <> "" then call write_bullet_and_variable_in_case_note("Assets", assets)
IF SHEL <> "" THEN CALL write_bullet_and_variable_in_case_note("SHEL/HEST", SHEL)
if FIAT_reasons <> "" then call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
If verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
If actions_taken <> "" then call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
If R_R_explained = 1 then call write_variable_in_case_note("* R/R explained.")
If other_notes <> "" then call write_bullet_and_variable_in_case_note("Notes", other_notes)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

call script_end_procedure("")






