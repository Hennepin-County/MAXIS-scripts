'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - MNSURE MEMO.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog MNsure_info_dialog, 0, 0, 196, 120, "MNsure Info Dialog"
  EditBox 60, 5, 70, 15, case_number
  DropListBox 110, 25, 75, 15, "denied"+chr(9)+"closed", how_case_ended
  EditBox 110, 45, 70, 15, denial_effective_date
  OptionGroup RadioGroup1
    RadioButton 20, 80, 35, 10, "WCOM", WCOM_check
    RadioButton 65, 80, 35, 10, "MEMO", MEMO_check
  EditBox 70, 100, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 140, 80, 50, 15
    CancelButton 140, 100, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 100, 10, "Was case closed or denied?:"
  Text 5, 50, 100, 10, "Denial/closure effective date:"
  GroupBox 10, 70, 100, 25, "How are you sending this?"
  Text 5, 105, 60, 10, "Worker signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabs case number
EMConnect ""
call MAXIS_case_number_finder(case_number)

'Shows dialog, checks for MAXIS or WCOM status.
Do
	Dialog MNsure_info_dialog
	cancel_confirmation
	If isdate(denial_effective_date) = False then MsgBox "You must put in a valid denial effective date (MM/DD/YYYY)."
Loop until isdate(denial_effective_date) = True

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'For the WCOM option it needs to go to SPEC/WCOM. Otherwise it goes to MEMO.
If radiogroup1 = 0 then
  'Navigating to SPEC/WCOM
  call navigate_to_MAXIS_screen("SPEC", "WCOM")  
  'This checks to make sure we've moved passed SELF.
  EMReadScreen SELF_check, 27, 2, 28
  If SELF_check = "Select Function Menu (SELF)" then script_end_procedure("Unable to get past SELF menu. Check for error messages and try again.")   
  'Updates to show HC only memos
  EMWriteScreen "Y", 3, 74
  transmit
  'Checks to make sure there's a waiting notice
  EMReadScreen waiting_check, 7, 7, 71
  If waiting_check <> "Waiting" then script_end_procedure("No waiting notice was found. You might be in the wrong footer month. If you still have this problem email your script administrator your footer month and case number. Also include a description of what's wrong.")
  'Creates a new WCOM. If it's unable the script will stop.
  EMWriteScreen "x", 7, 13
  transmit
  PF9
  EMReadScreen client_copy_check, 11, 1, 38
  If client_copy_check = "Client Copy" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
Else
  'Navigating to SPEC/MEMO
  call navigate_to_MAXIS_screen("SPEC", "MEMO")  
  'puts MEMO into edit mode
  PF5
  EMReadScreen memo_display_check, 12, 2, 33
  If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
  EMWriteScreen "x", 5, 10
  transmit
End if

'Enters different text for denials vs closures. This adds the different text to the first line
If how_case_ended = "denied" then write_variable_in_SPEC_MEMO("Your application was denied effective " & denial_effective_date & ". You may be able to purchase medical insurance through MNsure. If your family is under an income limit you may get financial help to purchase insurance. You can apply online at www.mnsure.org. If you have questions or need help to apply you can call the MNsure Call Center at 1-855-366-7873.")
If how_case_ended = "closed" then write_variable_in_SPEC_MEMO("Your case was closed effective " & denial_effective_date & ". You may be able to purchase medical insurance through MNsure. If your family is under an income limit you may get financial help to purchase insurance. You can apply online at www.mnsure.org. If you have questions or need help to apply you can call the MNsure Call Center at 1-855-366-7873.")
'Now it sends the rest of the memo, saves the memo and exits the memo screen
PF4
PF3

'Enters case note
Call start_a_blank_CASE_NOTE
If radiogroup1 = 0 then call write_variable_in_CASE_NOTE("Added MNsure info to client notice via WCOM. -" & worker_signature)
If radiogroup1 = 1 then call write_variable_in_CASE_NOTE("Sent client MNsure info via MEMO. -" & worker_signature)

script_end_procedure("")