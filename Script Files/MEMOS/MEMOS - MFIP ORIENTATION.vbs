'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - MFIP ORIENTATION.vbs"
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
'Must modify county_office_list manually each time to force recognition of variable from functions file. In other words, remove it from quotes.
BeginDialog MFIP_orientation_dialog, 0, 0, 366, 125, "MFIP orientation letter"
  EditBox 60, 5, 55, 15, case_number
  EditBox 185, 5, 55, 15, orientation_date
  EditBox 310, 5, 55, 15, orientation_time
  DropListBox 245, 25, 60, 15, county_office_list, interview_location
  EditBox 80, 45, 270, 15, MFIP_address_line_01
  EditBox 80, 65, 270, 15, MFIP_address_line_02
  EditBox 65, 85, 55, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 255, 85, 50, 15
    CancelButton 310, 85, 50, 15
    PushButton 315, 25, 50, 15, "refresh", refresh_button
  Text 5, 10, 50, 10, "Case Number:"
  Text 125, 10, 60, 10, "Orientation Date:"
  Text 250, 10, 60, 10, "Orientation Time:"
  Text 5, 30, 235, 10, "Location (select from dropdown and click ''refresh'', or fill in manually):"
  Text 20, 50, 55, 10, "Address line 01:"
  Text 20, 70, 55, 10, "Address line 02:"
  Text 5, 90, 60, 10, "Worker Signature:"
  Text 15, 105, 340, 20, "Please note: the dropdown above automatically fills in from your agency office/intake locations. It may not match your MFIP orientation locations. Please double check the address before pressing OK."
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(case_number)

'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
Do
	Do
		Do 
			Do 
				Do
					Do
						Do
							Dialog MFIP_orientation_dialog
							If ButtonPressed = cancel then stopscript
							If buttonPressed = refresh_button then
								IF interview_location <> "" then 
									call assign_county_address_variables(county_address_line_01, county_address_line_02)
									MFIP_address_line_01 = county_address_line_01
									MFIP_address_line_02 = county_address_line_02
								End if
							End if
						Loop until ButtonPressed = OK
						If isnumeric(case_number) = False or len(case_number) > 8 then MsgBox "You must fill in a valid case number. Please try again."
					Loop until isnumeric(case_number) = True and len(case_number) <= 8
					If isdate(orientation_date) = False then MsgBox "You did not enter a valid  date (MM/DD/YYYY format). Please try again."
				Loop until isdate(orientation_date) = True 
				If orientation_time = "" then MsgBox "You must type an interview time. Please try again."
			Loop until orientation_time <> ""
			If worker_signature = "" then MsgBox "You must provide a signature for your case note."
		Loop until worker_signature <> ""
		If MFIP_address_line_01 = "" or MFIP_address_line_02 = "" then MsgBox "You must enter an orientation address. Select one from the list, or enter one manually. Please note that the list fills in from intake locations, and may not be accurate in all agencies."
	Loop until MFIP_address_line_01 <> "" and MFIP_address_line_02 <> ""
	transmit
	EMReadScreen MAXIS_check, 5, 1, 39
	IF MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You need to be in MAXIS for this to work. Please try again."
Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "

'Using custom function to assign addresses to the selected office


'Navigating to SPEC/MEMO
call navigate_to_screen("SPEC", "MEMO")

'This checks to make sure we've moved passed SELF.
EMReadScreen SELF_check, 27, 2, 28
If SELF_check = "Select Function Menu (SELF)" then script_end_procedure("An error has occurred preventing the script from moving past the SELF menu. Your case might be in background. Check for errors and try again.")

'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
EMWriteScreen "x", 5, 10
transmit

'Writes the MEMO.

EMWriteScreen "************************************************************", 3, 15
EMSetCursor 4, 15	'Does this after the stars, because the stars shouldn't carry into the next line.
call write_new_line_in_SPEC_MEMO("You are required to attend a Financial Orientation for MFIP. Your orientation is scheduled on " & orientation_date & " at " & orientation_time & ".")
call write_new_line_in_SPEC_MEMO("")
call write_new_line_in_SPEC_MEMO("Your orientation is scheduled at the " & interview_location & " office located at: ")
call write_new_line_in_SPEC_MEMO("     " & MFIP_address_line_01)
call write_new_line_in_SPEC_MEMO("     " & MFIP_address_line_02)
call write_new_line_in_SPEC_MEMO("")
call write_new_line_in_SPEC_MEMO("If you cannot attend this orientation, please contact the agency office to reschedule. Failure to attend an orientation will result in a sanction (reduction) of your MFIP benefits.")
EMSendKey "************************************************************"

stopscript
'Exits the MEMO
PF4


'Navigates to CASE/NOTE
call navigate_to_screen("case", "note")
PF9

'Writes the case note
EMSendKey "<home>" & "***MFIP orientation scheduled***" & "<newline>"
call write_new_line_in_case_note("* Appt letter sent via SPEC/MEMO.")
call write_new_line_in_case_note("* Orientation is scheduled on " & orientation_date & " at " & orientation_time & ".")
call write_editbox_in_case_note("Location", interview_location, 6)
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

'Script ends
script_end_procedure("")





