'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - OVERDUE BABY.vbs"
start_time = timer


'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'DIALOG---------------------------------------------------------------------------------------------------------------------
BeginDialog memos_overdue_baby_dialog, 0, 0, 141, 85, "MEMOS - OVERDUE BABY"
  EditBox 60, 5, 60, 15, case_number
  EditBox 70, 25, 60, 15, worker_signature
  CheckBox 5, 45, 100, 15, "TIKL for ten day follow up?", tikl_for_ten_day_follow_up_checkbox
  ButtonGroup ButtonPressed
    OkButton 30, 65, 50, 15
    CancelButton 80, 65, 50, 15
  Text 5, 5, 50, 15, "Case Number:"
  Text 5, 25, 60, 15, "Worker Signature:"
EndDialog

EndDialog

'THE SCRIPT------------------------------------------------------------------------------------------------------------------

'Connects to BlueZone default screen
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(case_number)

Do
  
	Dialog memos_overdue_baby_dialog
	If ButtonPressed = 0 then stopscript
	If case_number = ""  or isnumeric(case_number) = false then MsgBox "You did not enter a valid case number. Please try again."
	If worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
Loop until case_number <> "" and isnumeric(case_number) = true and worker_signature <> ""
transmit
call check_for_MAXIS(True)


'Navigates into SPEC/MEMO
	call navigate_to_screen("SPEC", "MEMO")

'Checks to make sure we're past the SELF menu
	EMReadScreen still_self, 27, 2, 28 
	If still_self = "Select Function Menu (SELF)" then script_end_procedure("Script was not able to get past SELF menu. Is case in background?")

'Creates a new MEMO. If it's unable the script will stop.
	PF5
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
	EMWriteScreen "x", 5, 10
	transmit

'Writes the info into the MEMO
EMSetCursor 3, 15
call write_variable_in_SPEC_MEMO("Our records indicate your due date has passed and you did not report the birth of your child or the pregnancy end date. Please contact us within 10 days of this notice with the following information or your case may close:")
call write_variable_in_SPEC_MEMO("")
call write_variable_in_SPEC_MEMO("* Date of the birth or pregnancy end date.")  
call write_variable_in_SPEC_MEMO("* Baby's sex and full name.")
call write_variable_in_SPEC_MEMO("* Baby's social security number.") 
call write_variable_in_SPEC_MEMO("* Full name of the baby's father.") 
call write_variable_in_SPEC_MEMO("* Does the baby's father live in your home?") 
call write_variable_in_SPEC_MEMO("* If so, does the father have a source of income?")
call write_variable_in_SPEC_MEMO("  (If so, what is the source of income?)")
call write_variable_in_SPEC_MEMO("* Is there other health insurance available through any       household member's employer, or privately?")
call write_variable_in_SPEC_MEMO("")
call write_variable_in_SPEC_MEMO("Thank you,")
PF4

'Navigates to blank case note
call navigate_to_screen("CASE", "NOTE")
PF9

'Writes the case note
call write_variable_in_CASE_NOTE("***Overdue Baby***")
call write_variable_in_CASE_NOTE("* SPEC/MEMO sent this date informing client that they need to report ")
call write_variable_in_CASE_NOTE("      information regarding the birth of their child, and/or pregnancy end ")
call write_variable_in_CASE_NOTE("      date, within 10 days or their case may close.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)


'Navigates to TIKL (if selected)

If tikl_for_ten_day_follow_up_checkbox = checked then 
	call navigate_to_screen("DAIL", "WRIT")
	call create_MAXIS_friendly_date(date, 10, 5, 18)
	call write_variable_in_TIKL("Has information on new baby/end of pregnancy been received? If not, consider case closure/take appropriate action.")
	transmit
	PF3
End If

script_end_procedure("Success! The script has case noted the overdue baby info, sent a SPEC/MEMO to the client, and TIKLed for 10-day return of information.")
