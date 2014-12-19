'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - APPLICATION RECEIVED.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 115, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 95, 50, 15
    CancelButton 95, 95, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
EndDialog

BeginDialog LTC_app_recd_dialog, 0, 0, 286, 335, "LTC application received dialog"
  EditBox 45, 35, 65, 15, appl_date
  EditBox 75, 55, 205, 15, appl_type
  EditBox 160, 75, 120, 15, forms_needed
  EditBox 30, 95, 30, 15, CFR
  EditBox 110, 95, 170, 15, HH_comp
  EditBox 70, 115, 210, 15, pre_FACI_ADDR
  EditBox 60, 135, 220, 15, basis_of_elig
  EditBox 35, 155, 245, 15, FACI
  EditBox 60, 175, 220, 15, retro_request
  EditBox 35, 195, 245, 15, AREP
  EditBox 60, 215, 220, 15, SWKR
  EditBox 60, 235, 220, 15, INSA
  EditBox 65, 255, 215, 15, adult_signatures
  EditBox 50, 275, 230, 15, LTCC
  EditBox 55, 295, 225, 15, actions_taken
  EditBox 75, 315, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 170, 315, 50, 15
    CancelButton 230, 315, 50, 15
    PushButton 15, 15, 25, 10, "TYPE", TYPE_button
    PushButton 40, 15, 25, 10, "PROG", PROG_button
    PushButton 65, 15, 25, 10, "HCRE", HCRE_button
    PushButton 90, 15, 25, 10, "REVW", REVW_button
    PushButton 115, 15, 25, 10, "MEMB", MEMB_button
    PushButton 180, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 180, 25, 45, 10, "next panel", next_panel_button
    PushButton 230, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 230, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 160, 25, 10, "FACI:", FACI_button
    PushButton 5, 200, 25, 10, "AREP:", AREP_button
    PushButton 25, 220, 30, 10, "SWKR:", SWKR_button
    PushButton 5, 240, 25, 10, "INSA/", INSA_button
    PushButton 30, 240, 25, 10, "MEDI:", MEDI_button
  GroupBox 10, 5, 135, 25, "General STAT navigation:"
  GroupBox 175, 5, 105, 35, "STAT-based navigation"
  Text 5, 40, 35, 10, "Appl date: "
  Text 5, 60, 65, 10, "Appl type received:"
  Text 5, 80, 150, 10, "Forms needed? 1503, 3543, 3050, 5181, etc?:"
  Text 5, 100, 20, 10, "CFR:"
  Text 70, 100, 40, 10, "HH Comp:"
  Text 5, 120, 60, 10, "Pre FACI address:"
  Text 5, 140, 50, 10, "Basis of elig:"
  Text 5, 180, 50, 10, "Retro request:"
  Text 5, 220, 20, 10, "PHN/"
  Text 5, 260, 60, 10, "Adult signatures:"
  Text 5, 280, 40, 10, "LTCC info:"
  Text 5, 300, 50, 10, "Actions taken:"
  Text 5, 320, 65, 10, "Worker signature:"
EndDialog



'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
HH_memb_row = 05

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------

footer_month = datepart("m", date)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", date)
footer_year = "" & footer_year - 2000

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Searching for case number.
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Showing the case number dialog, transmits to check for MAXIS.
Do
  Dialog case_number_dialog
  If buttonpressed = 0 then stopscript
  If case_number = "" then MsgBox "You must type a case number!"
Loop until case_number <> ""

'Now it checks to make sure MAXIS is running on this screen.
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("MAXIS is not found. The script will now exit. Make sure you start this script on the window that has MAXIS.")

'Navigating to STAT/HCRE so we can grab the app date
call navigate_to_screen("stat", "hcre")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Grabs autofill info from STAT
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", appl_date)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", MEDI)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)

'Now, because INSA and MEDI will go on the same variable, we're going to add INSA to MEDI. To separate them in the case note, we have to add a semicolon (assuming both have data).
If INSA <> "" and MEDI <> "" then
  INSA = INSA & "; " & MEDI
Else
  INSA = INSA & MEDI
End if

'The dialog
Do
  Do
    Do
      Do
        Do
          Dialog LTC_app_recd_dialog
          If ButtonPressed = 0 then stopscript
          If buttonpressed <> -1 then call navigation_buttons
          If buttonpressed = prev_panel_button then call panel_navigation_prev
          If buttonpressed = next_panel_button then call panel_navigation_next
          If buttonpressed = prev_memb_button then call memb_navigation_prev
          If buttonpressed = next_memb_button then call memb_navigation_next
          If buttonpressed = prev_panel_button or buttonpressed = next_panel_button or buttonpressed = prev_memb_button or buttonpressed = next_memb_button then transmit 'it won't transmit otherwise
        Loop until buttonpressed = -1
        If isdate(appl_date) = False then MsgBox "You must enter a valid APPL date (MM/DD/YYYY). Please try again."
      Loop until isdate(appl_date) = True
      If len(actions_taken) < 3 then MsgBox "You must fill in the actions taken section. Please try again."
    Loop until worker_signature <> ""
    If worker_signature = "" then MsgBox "You must sign your case note!"
  Loop until worker_signature <> ""
  transmit 'to check for password and MAXIS
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "MAXIS appears to have passworded out, or you navigated away from it. Navigate back to MAXIS before trying again."
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

'Navigating to case/note and starting a fresh note. Will close if unable to get to edit mode (in case worker is in inquiry for instance).
call navigate_to_screen("case", "note")
PF9
EMReadScreen edit_mode_check, 1, 20, 09
If edit_mode_check <> "A" then script_end_procedure("Unable to get to edit mode in this case note. Are you in inquiry? Is the case out of county? Resolve these issues and try the script again.")

'Writing the case note
EMSendKey "***LTC intake***" + "<newline>"
If appl_date <> "" then call write_editbox_in_case_note("Application date", appl_date, 6)
If appl_type <> "" then call write_editbox_in_case_note("Application type received", appl_type, 6)
If forms_needed <> "" then call write_editbox_in_case_note("Forms Needed", forms_needed, 6)
If HH_comp <> "" then call write_editbox_in_case_note("HH comp", HH_comp, 6)
If CFR <> "" then call write_editbox_in_case_note("CFR", CFR, 6)
If pre_FACI_ADDR <> "" then call write_editbox_in_case_note("Pre FACI address", pre_FACI_ADDR, 20)
If basis_of_elig <> "" then call write_editbox_in_case_note("Basis of eligibility", basis_of_elig, 6)
If FACI <> "" then call write_editbox_in_case_note("FACI", FACI, 6)
If retro_request <> "" then call write_editbox_in_case_note("Retro request", retro_request, 6)
If AREP <> "" then call write_editbox_in_case_note("AREP", AREP, 6)
If SWKR <> "" then call write_editbox_in_case_note("PHN/SWKR", SWKR, 6)
If INSA <> "" then call write_editbox_in_case_note("INSA/MEDI", INSA, 6)
If adult_signatures <> "" then call write_editbox_in_case_note("Adult signatures", adult_signatures, 6)
If LTCC <> "" then call write_editbox_in_case_note("LTCC info", LTCC, 6)
call write_editbox_in_case_note("Actions taken", actions_taken, 6)
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

script_end_procedure("")






