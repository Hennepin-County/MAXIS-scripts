'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - 1503.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog case_number_dialog, 0, 0, 161, 60, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, case_number
  Text 15, 25, 50, 10, "Footer month:"
  EditBox 65, 20, 25, 15, footer_month
  Text 95, 25, 20, 10, "Year:"
  EditBox 120, 20, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

BeginDialog DHS_1503_dialog, 0, 0, 366, 275, "1503 Dialog"
  Text 5, 10, 20, 10, "FACI:"
  EditBox 30, 5, 135, 15, FACI
  Text 5, 30, 55, 10, "Length of stay:"
  DropListBox 60, 25, 70, 15, "30 days or less"+chr(9)+"31 to 90 days"+chr(9)+"91 to 180 days"+chr(9)+"over 180 days", length_of_stay
  Text 140, 30, 95, 10, "Recommended level of care:"
  DropListBox 240, 25, 45, 15, "SNF"+chr(9)+"NF"+chr(9)+"ICF-MR"+chr(9)+"RTC", level_of_care
  Text 5, 50, 50, 10, "Admitted from:"
  DropListBox 60, 45, 125, 15, "acute-care hospital"+chr(9)+"home"+chr(9)+"RTC"+chr(9)+"other SNF or NF"+chr(9)+"ICF-MR", admitted_from
  Text 190, 50, 70, 10, "If hospital, list here:"
  EditBox 260, 45, 95, 15, hospital_admitted_from
  Text 5, 70, 40, 10, "Admit date:"
  EditBox 45, 65, 65, 15, admit_date
  Text 120, 70, 100, 10, "Discharge date (if applicible):"
  EditBox 225, 65, 65, 15, discharge_date
  CheckBox 15, 85, 155, 10, "If you've processed this 1503, check here.", processed_1503_check
  GroupBox 5, 100, 355, 75, "actions/proofs"
  CheckBox 15, 115, 65, 10, "Updated RLVA?", updated_RLVA_check
  CheckBox 90, 115, 65, 10, "Updated FACI?", updated_FACI_check
  CheckBox 165, 115, 55, 10, "Need 3543?", need_3543_check
  CheckBox 230, 115, 100, 10, "Need asset assessment?", need_asset_assessment_check
  Text 10, 135, 115, 10, "Other proofs needed (if applicable):"
  EditBox 130, 130, 225, 15, verifs_needed
  CheckBox 15, 155, 50, 10, "Sent 3050?", sent_3050_check
  CheckBox 165, 155, 105, 10, "Sent verif req? If so, to who:", sent_verif_request_check
  ComboBox 275, 150, 80, 15, "client"+chr(9)+"AREP", sent_request_to
  Text 5, 185, 25, 10, "Notes:"
  EditBox 30, 180, 330, 15, notes
  Text 5, 205, 75, 10, "Sign your case note:"
  EditBox 85, 200, 75, 15, worker_sig
  CheckBox 10, 230, 260, 10, "Check here to have the script TIKL out to contact the FACI re: length of stay.", TIKL_check
  CheckBox 10, 245, 155, 10, "Check here to have the script update HCMI.", HCMI_update_check
  CheckBox 10, 260, 150, 10, "Check here to have the script update FACI.", FACI_update_check
  ButtonGroup ButtonPressed
    OkButton 200, 210, 50, 15
    CancelButton 260, 210, 50, 15
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""


'Grabs the case number
call find_variable("Case Nbr: ", case_number, 8)
Dialog case_number_dialog
If buttonpressed = 0 then stopscript

'Checks for MAXIS
transmit 'to check for MAXIS
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS is not found on this screen. The script will now stop")

'Navigates to STAT to make sure the case is out of background
call navigate_to_screen("stat", "____")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Unable to get to STAT. Your case could be in background. Wait a few moments then try again.")

'THE DIALOG

Do
  Do
    Do
      Dialog DHS_1503_dialog
      If buttonpressed = 0 then stopscript
      If isdate(admit_date) = False then MsgBox "You did not type a valid date (MM/DD/YYYY) in the admit date box. This is required for the script to work correctly."
    Loop until isdate(admit_date) = True
    PF3
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then MsgBox "You are not in MAXIS. Navigate your ''S1'' screen to MAXIS and try again. You might be passworded out."
  Loop until MAXIS_check = "MAXIS" 
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen NOTE_mode_check, 7, 20, 3
  If NOTE_mode_check <> "Mode: A" then MsgBox "A valid case note could not be found. Are you in the right case number? Did you accidentally start the script in inquiry? Check the case number, the screen, and try again."
Loop until NOTE_mode_check = "Mode: A"

'DATE/TIME EQUATIONS
admit_date_DD = datepart("d", admit_date)
If len(admit_date_DD) = 1 then admit_date_DD = "0" & admit_date_DD
admit_date_MM = datepart("m", admit_date)
If len(admit_date_MM) = 1 then admit_date_MM = "0" & admit_date_MM
admit_date_YY = datepart("yyyy", admit_date)
If len(admit_date_YY) = 4 then admit_date_YYYY = admit_date_YYYY - 2000

If TIKL_check = 1 then
  If length_of_stay = "30 days or less" then TIKL_multiplier = 30
  If length_of_stay = "31 to 90 days" then TIKL_multiplier = 90
  If length_of_stay = "91 to 180 days" then TIKL_multiplier = 180
  TIKL_date = dateadd("d", TIKL_multiplier, admit_date)
  TIKL_date_DD = datepart("d", TIKL_date)
  If len(TIKL_date_DD) = 1 then TIKL_date_DD = "0" & TIKL_date_DD
  TIKL_date_MM = datepart("m", TIKL_date)
  If len(TIKL_date_MM) = 1 then TIKL_date_MM = "0" & TIKL_date_MM
  TIKL_date_YY = datepart("yyyy", TIKL_date)
  If len(TIKL_date_YY) = 4 then TIKL_date_YY = TIKL_date_YY - 2000
End if


'CASE NOTING


If processed_1503_check = 1 then 
  EMSendKey "***Processed 1503 from " & FACI & "***" & "<newline>"
Else
  EMSendKey "***Rec'd 1503 from " & FACI & ", DID NOT PROCESS***" & "<newline>"
End if
Call write_editbox_in_case_note("Length of stay", length_of_stay, 6)
Call write_editbox_in_case_note("Recommended level of care", level_of_care, 6)
Call write_editbox_in_case_note("Admitted from", admitted_from, 6)
If hospital_admitted_from <> "" then Call write_editbox_in_case_note("Hospital admitted from", hospital_admitted_from, 6)
Call write_editbox_in_case_note("Admit date", admit_date, 6)
If discharge_date <> "" then Call write_editbox_in_case_note("Discharge date", discharge_date, 6)
Call write_new_line_in_case_note("---")
If updated_RLVA_check = 1 and updated_FACI_check = 1 then 
Call write_new_line_in_case_note("* Updated RLVA and FACI.")
Else
  If updated_RLVA_check = 1 then Call write_new_line_in_case_note("* Updated RLVA.")
  If updated_FACI_check = 1 then Call write_new_line_in_case_note("* Updated FACI.")
End if
If need_3543_check = 1 then Call write_new_line_in_case_note("* A 3543 is needed.")
If sent_3050_check = 1 then call write_new_line_in_case_note("* Sent 3050.")
If verifs_needed <> "" then Call write_editbox_in_case_note("Verifs needed", verifs_needed, 6)
If sent_verif_request_check = 1 then Call write_editbox_in_case_note("Sent verif request to", sent_request_to, 6)
If processed_1503_check = 1 then Call write_new_line_in_case_note("* Completed & Returned 1503 to LTCF.")
If TIKL_check = 1 then Call write_new_line_in_case_note("* TIKLed to recheck length of stay on " & TIKL_date & ".")
Call write_new_line_in_case_note("---")
Call write_editbox_in_case_note("Notes", notes, 6)
Call write_new_line_in_case_note("---")
Call write_new_line_in_case_note(worker_sig)
transmit

'TIKLING

If TIKL_check = 1 then
  call navigate_to_screen("dail", "writ")
  EMWriteScreen TIKL_date_MM, 5, 18
  EMWriteScreen TIKL_date_DD, 5, 21
  EMWriteScreen TIKL_date_YY, 5, 24
  EMSetCursor 9, 3
  EMSendKey "Have " & worker_sig & " call " & FACI & " re: length of stay. " & TIKL_multiplier & " days expired."
  transmit
  PF3
End if

'UPDATING FACI

If FACI_update_check = 1 then
  call navigate_to_screen("stat", "faci")
  EMReadScreen ERRR_check, 4, 2, 52
  If ERRR_check = "ERRR" then transmit
  EMWriteScreen "nn", 20, 79
  transmit
  EMWriteScreen FACI, 6, 43
  If length_of_stay = "30 days or less" and level_of_care = "SNF" then EMWriteScreen "44", 7, 43
  If length_of_stay = "31 to 90 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
  If level_of_care = "NF" then EMWriteScreen "42", 7, 43
  EMWriteScreen "n", 8, 43
  EMWriteScreen admit_date_MM, 14, 47
  EMWriteScreen admit_date_DD, 14, 50
  EMWriteScreen datepart("yyyy", admit_date), 14, 53
  If isdate(discharge_date) = "True" then
    discharge_date_month = datepart("m", discharge_date)
    If len(discharge_date_month) = 1 then discharge_date_month = "0" & discharge_date_month
    EMWriteScreen discharge_date_month, 14, 71
    discharge_date_day = datepart("d", discharge_date)
    If len(discharge_date_day) = 1 then discharge_date_day = "0" & discharge_date_day
    EMWriteScreen discharge_date_day, 14, 74
    EMWriteScreen datepart("yyyy", discharge_date), 14, 77
    transmit
  End if
End if

'UPDATING HCMI

If HCMI_update_check = 1 then 
  If FACI_update_check = 1 then
    EMWriteScreen "hcmi", 20, 71
    transmit
  Else
    call navigate_to_screen("stat", "hcmi")
  End if
  EMReadScreen ERRR_check, 4, 2, 52
  If ERRR_check = "ERRR" then transmit
  EMReadScreen current_panel_number, 1, 2, 78
  If current_panel_number = "0" then
    EMWriteScreen "nn", 20, 79
    transmit
  Else
    PF9
  End if
  EMWriteScreen "dp", 10, 57
  transmit
  transmit
End if

script_end_procedure("")







