'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMO - DUPLICATE ASSISTANCE WCOM.vbs"
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

BeginDialog dup_dlg, 0, 0, 141, 95, "Duplicate Assistance WCOM"
  EditBox 75, 5, 60, 15, case_number
  EditBox 75, 25, 60, 15, worker_signature
  EditBox 75, 45, 25, 15, MAXIS_footer_month
  EditBox 110, 45, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 30, 70, 50, 15
    CancelButton 85, 70, 50, 15
  Text 5, 30, 60, 10, "Worker Signature: "
  Text 5, 50, 65, 10, "Footer month/year:"
  Text 5, 10, 50, 10, "Case Number: "
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS and grabbing footer month/year
EMConnect ""
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'warning box
Msgbox "WARNING: If you have multiple waiting SNAP results this script may be unable to find the most recent one. Please process manually in those instances." & vbNewLine & vbNewLine & vbNewLine &_
		"- If this case includes members who are residing in a battered women's shelter please review approval." & vbNewLine & vbNewLine &_
		"- If this was an expedited case where client reported they did not receive benefits in another state please review approval" & vbNewLine & vbNewLine &_
		"- See CM 001.21 for more details on these two situations and how they qualify for duplicate assistance."
		
'the dialog
Do	
	Do
		Do
			dialog dup_dlg
			cancel_confirmation
			If MAXIS_footer_month = "" or MAXIS_footer_year = "" THEN Msgbox "Please fill in footer month and year (MM YY format)."
			If case_number = "" THEN MsgBox "Please enter a case number."
			If worker_signature = "" THEN MsgBox "Please sign your note."
		Loop until MAXIS_footer_month <> "" & MAXIS_footer_year <> ""
	Loop until case_number <> ""
Loop until worker_signature <> ""

'Converting dates into useable forms
If len(MAXIS_footer_month) < 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
If len(MAXIS_footer_year) > 2 THEN MAXIS_footer_year = right(MAXIS_footer_year, 2)


'Navigating to the spec wcom screen
CALL Check_for_MAXIS(False)
back_to_self
Emwritescreen case_number, 18, 43
Emwritescreen MAXIS_footer_month, 20, 43
Emwritescreen MAXIS_footer_year, 20, 46
transmit
CALL navigate_to_MAXIS_screen("SPEC", "WCOM")

'Searching for waiting SNAP notice
wcom_row = 6
Do
	wcom_row = wcom_row + 1
	Emreadscreen program_type, 2, wcom_row, 26
	Emreadscreen print_status, 7, wcom_row, 71
	If program_type = "FS" then
		If print_status = "Waiting" then
			Emwritescreen "x", wcom_row, 13
			exit Do
		End If
	End If
	If wcom_row = 17 then
		PF8
		Emreadscreen spec_edit_check, 6, 24, 2
		wcom_row = 6
	end if
	If spec_edit_check = "NOTICE" THEN no_fs_waiting = true
Loop until spec_edit_check = "NOTICE"

If no_fs_waiting = true then script_end_procedure("No waiting FS results were found for the requested month")

'writing the WCOM
Transmit
PF9
CALL write_variable_in_SPEC_MEMO("******************************************************")
CALL write_variable_in_SPEC_MEMO("Dear Client,")
CALL write_variable_in_SPEC_MEMO("")
CALL write_variable_in_SPEC_MEMO("You will not be eligible for SNAP benefits this month since you have received SNAP benefits on another case for the same month.")
CALL write_variable_in_SPEC_MEMO("Per program rules SNAP participants are not eligible for duplicate benefits in the same benefit month.")
CALL write_variable_in_SPEC_MEMO("")
CALL write_variable_in_SPEC_MEMO("If you have any questions or concerns please feel free to contact your worker.")
CALL write_variable_in_SPEC_MEMO("---")
CALL write_variable_in_SPEC_MEMO(worker_signature)
CALL write_variable_in_SPEC_MEMO("")
CALL write_variable_in_SPEC_MEMO("******************************************************")
PF4

script_end_procedure("WCOM has been added to the first found waiting SNAP notice for the month and case selected. Please feel free to review the notice.")