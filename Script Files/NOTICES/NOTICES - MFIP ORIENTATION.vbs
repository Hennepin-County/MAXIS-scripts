'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - MFIP ORIENTATION.vbs"
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

STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each MEMBER
'END OF stats block==============================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
'Must modify county_office_list manually each time to force recognition of variable from functions file. In other words, remove it from quotes.
BeginDialog MFIP_orientation_dialog, 0, 0, 366, 155, "MFIP orientation letter"
  EditBox 60, 5, 55, 15, case_number
  EditBox 185, 5, 55, 15, orientation_date
  EditBox 310, 5, 55, 15, orientation_time
  EditBox 205, 25, 55, 15, member_list
  DropListBox 245, 40, 60, 15, county_office_list, interview_location
  EditBox 80, 60, 270, 15, MFIP_address_line_01
  EditBox 80, 75, 270, 15, MFIP_address_line_02
  EditBox 65, 95, 55, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 255, 135, 50, 15
    CancelButton 310, 135, 50, 15
    PushButton 315, 40, 50, 15, "refresh", refresh_button
  Text 5, 10, 50, 10, "Case Number:"
  Text 125, 10, 60, 10, "Orientation Date:"
  Text 250, 10, 60, 10, "Orientation Time:"
  Text 5, 25, 195, 10, "Enter HH member numbers to attend, separated by commas:"
  Text 5, 40, 235, 10, "Location (select from dropdown and click ''refresh'', or fill in manually):"
  Text 20, 60, 55, 10, "Address line 01:"
  Text 20, 75, 55, 10, "Address line 02:"
  Text 5, 95, 60, 10, "Worker Signature:"
  Text 15, 115, 340, 20, "Please note: the dropdown above automatically fills in from your agency office/intake locations. It may not match your MFIP orientation locations. Please double check the address before pressing OK."
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
							cancel_confirmation
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
			If member_list = "" then MsgBox "You must enter at least one household member to attend the interview."
		Loop until member_list <> ""
		If worker_signature = "" then MsgBox "You must provide a signature for your case note."
	Loop until worker_signature <> ""
	If MFIP_address_line_01 = "" or MFIP_address_line_02 = "" then MsgBox "You must enter an orientation address. Select one from the list, or enter one manually. Please note that the list fills in from intake locations, and may not be accurate in all agencies."
Loop until MFIP_address_line_01 <> "" and MFIP_address_line_02 <> ""
transmit

'checking for active MAXIS session
Call check_for_MAXIS(False)

'Creating an array from the member number list to get names for notice
member_array = split(member_list, ",")
	for each member in member_array
		call navigate_to_MAXIS_screen("STAT", "MEMB")
		member = replace(member, " ", "")
		if len(member) = 1 then member = "0" & member
		EMWriteScreen member, 20, 76
		transmit
		EMReadScreen name_long, 44, 6, 30
		member_name = replace(name_long, "_", "")
		member_name = replace(member_name, " First:", ",")
		if members_to_attend <> "" then members_to_attend = members_to_attend & "; " & member_name
		if members_to_attend = "" then members_to_attend = member_name
	next

'Navigating to SPEC/MEMO
call navigate_to_MAXIS_screen("SPEC", "MEMO")

'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
EMWriteScreen "x", 5, 10
transmit

''Writes the MEMO.
EMSetCursor 3, 15
call write_variable_in_SPEC_MEMO("************************************************************")
call write_variable_in_SPEC_MEMO("You are required to attend a Minnesota Family Investment Program financial orientation. The following members of your household need to attend this appointment: " & members_to_attend)
call write_variable_in_SPEC_MEMO("Your orientation is scheduled on " & orientation_date & " at " & orientation_time & ".")
call write_variable_in_SPEC_MEMO("Your orientation is scheduled at the " & interview_location & " office located at: ")
call write_variable_in_SPEC_MEMO(county_address_line_01)
call write_variable_in_SPEC_MEMO(county_address_line_02)
call write_variable_in_SPEC_MEMO("If you cannot attend this orientation, please contact the agency office to reschedule.  Failure to attend an orientation will result in a sanction of your MFIP benefits.")
call write_variable_in_SPEC_MEMO("************************************************************")
'Exits the MEMO
PF4

'Writes the case note----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
call write_variable_in_case_note("* Financial Orientation letter sent via SPEC/MEMO. *")
call write_variable_in_case_note("Orientation is scheduled on: " & orientation_date & " at " & orientation_time)
call write_variable_in_case_note("Location: " & interview_location)
call write_bullet_and_variable_in_case_note("Household members needing to attend: ", members_to_attend)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")	'Script ends
