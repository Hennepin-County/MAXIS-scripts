'GRABBING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - BABY BORN.vbs"
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

'DIALOGS-------------------------------------------------------------------------------------------------------------
BeginDialog baby_born_dialog, 0, 0, 211, 250, "NOTES - BABY BORN"
  EditBox 60, 5, 80, 15, case_number
  EditBox 60, 25, 95, 15, babys_name
  EditBox 50, 45, 80, 15, date_of_birth
  DropListBox 85, 65, 70, 15, "Select One"+chr(9)+"Yes"+chr(9)+"No", father_in_household
  EditBox 75, 85, 80, 15, fathers_employer
  EditBox 75, 105, 80, 15, mothers_employer
  DropListBox 35, 125, 70, 15, "Select One"+chr(9)+"Yes"+chr(9)+"No", other_health_insurance
  EditBox 115, 145, 80, 15, OHI_source
  EditBox 50, 165, 105, 15, other_notes
  EditBox 55, 185, 105, 15, actions_taken
  EditBox 155, 210, 40, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 95, 230, 50, 15
    CancelButton 150, 230, 50, 15
  Text 5, 5, 55, 15, "Case Number: "
  Text 5, 25, 55, 15, "Baby's Name:"
  Text 5, 45, 45, 15, "Date of Birth:"
  Text 5, 65, 75, 15, "Father In Household?"
  Text 5, 85, 70, 15, "Father's Employer:"
  Text 5, 105, 70, 10, "Mother's Employer: "
  Text 5, 125, 25, 15, "OHI?"
  Text 5, 145, 110, 15, "If yes to OHI, source of the OHI:"
  Text 5, 165, 45, 15, "Other Notes:"
  Text 5, 185, 50, 15, "Actions Taken:"
  Text 90, 210, 65, 15, "Worker Signature:"
EndDialog



'THE SCRIPT---------------------------------------------------------------------------------------------------------------
EMConnect ""


call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""



  Do
    Do
      Dialog baby_born_dialog
      If buttonpressed = 0 then stopscript
      If case_number = "" then MsgBox "You must have a case number to continue!"
	If worker_signature = "" then MsgBox "You must sign your case note!"
    Loop until case_number <> "" & worker_signature <> ""
    call check_for_MAXIS(True)
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

call write_variable_in_CASE_NOTE("***Baby Born***")
call write_bullet_and_variable_in_CASE_NOTE("Baby's Name", babys_name)
call write_bullet_and_variable_in_CASE_NOTE("Date of Birth", date_of_birth)
call write_bullet_and_variable_in_CASE_NOTE("Father in Household?", father_in_household)
call write_bullet_and_variable_in_CASE_NOTE("Father's Employer", fathers_employer)
call write_bullet_and_variable_in_CASE_NOTE("Mother's Employer", mothers_employer)
call write_bullet_and_variable_in_CASE_NOTE("Other OHI?", other_health_insurance)
call write_bullet_and_variable_in_CASE_NOTE("Source of OHI", OHI_source)
call write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
call write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)



script_end_procedure("")
