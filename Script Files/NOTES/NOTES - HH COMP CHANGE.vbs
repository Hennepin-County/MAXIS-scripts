'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - HH COMP CHANGE.vbs"
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 150          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'----DIALOGS------------------------------------------------------------------------------------------------------------------------------ 
'dialog box for HH comp change
BeginDialog HH_Comp_Change_Dialog, 0, 0, 301, 175, "Household Comp Change"
  Text 5, 15, 50, 10, "Case Number"
  EditBox 60, 10, 100, 15, Case_Number
  Text 5, 35, 80, 10, "Unit Member HH Change"
  EditBox 90, 30, 45, 15, HH_Member
  Text 5, 55, 85, 10, "Date Reported/Addendum"
  EditBox 95, 50, 60, 15, Date_Reported
  Text 165, 55, 55, 15, "Effective Date"
  EditBox 215, 50, 70, 15, Effective_Date
  CheckBox 10, 70, 100, 10, "Is the change temporary?", Temporary_Change_Checkbox
  Text 5, 90, 50, 10, "Action Taken"
  EditBox 55, 85, 230, 15, Action_Taken
  Text 5, 110, 60, 10, "Additional Notes"
  EditBox 60, 105, 225, 15, Additional_Notes
  CheckBox 15, 125, 130, 15, "Is client reporting the birth of a baby?", Baby_Born_Checkbox
  Text 5, 150, 50, 15, "Worker Name"
  EditBox 60, 145, 100, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 150, 50, 15
    CancelButton 230, 150, 50, 15
EndDialog

'dialog box for baby born
BeginDialog Baby_Born_Dialog, 0, 0, 311, 165, "Client Reports Birth of Baby"
  Text 5, 10, 50, 10, "Name of baby"
  EditBox 60, 5, 150, 15, Baby_Name
  Text 5, 30, 45, 10, "Date of Birth"
  EditBox 55, 25, 60, 15, Date_Birth
  Text 130, 30, 45, 10, "Baby's SSN"
  EditBox 180, 25, 85, 15, Baby_SSN
  CheckBox 10, 45, 100, 10, "Father is in the household.", Father_HH_Checkbox
  Text 5, 65, 50, 10, "Father's Name"
  EditBox 60, 60, 150, 15, Father_Name
  Text 5, 85, 65, 10, "Father's Employer"
  EditBox 70, 80, 225, 15, Father_Employer
  Text 5, 105, 65, 10, "Mother's Employer"
  EditBox 70, 100, 225, 15, Mother_Employer
  CheckBox 5, 120, 80, 15, "The baby has OHI", OHI_Checkbox
  Text 110, 125, 55, 10, "Source of OHI"
  EditBox 165, 120, 115, 15, OHI_Source
  ButtonGroup ButtonPressed
    OkButton 180, 140, 50, 15
    CancelButton 235, 140, 50, 15
EndDialog


'---SCRIPTS--------------------------------------------------------------------------------------------------------------------------------------------



'Connecting to BlueZone
EMConnect ""

'autofill case number function 
CALL MAXIS_case_number_finder(case_number)



'run dialog - need to do the DO loops right after running dialog
DO
	err_msg = ""
	DIALOG HH_Comp_Change_Dialog
	cancel_confirmation
	IF Case_Number = "" THEN err_msg = "You must enter case number!"
	IF HH_Member = "" THEN err_msg = err_msg & vbNewLine & "You must enter a HH Member"
	IF Date_Reported = "" THEN err_msg = err_msg & vbNewLine & "You must enter date reported"
	IF Effective_Date = "" THEN err_msg = err_msg & vbNewLine & "You must enter effective date"
	IF Action_Taken = "" THEN err_msg = err_msg & vbNewLine & "You must enter actions taken"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "Please sign your note"
	IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
LOOP UNTIL err_msg = ""


IF Baby_Born_Checkbox = 1 THEN 

'Do loop for Baby Born Dialogbox
DO
	err_msg = ""
	DIALOG Baby_Born_Dialog
	cancel_confirmation
	IF Date_Birth = "" THEN err_msg = err_msg & vbNewLine &  "You must enter a birth date"
	IF Baby_SSN = "" THEN err_msg = err_msg & vbNewLine &  "You must enter baby's Social Security Number"
	IF Father_Name = "" THEN err_msg = err_msg & vbNewLine &  "You must enter Father's name"
	IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
LOOP UNTIL err_msg = ""
	
END IF

'Checks MAXIS for password prompt
CALL check_for_MAXIS(false)

'Navigates to case note
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

'Send PF9 to case note
PF9

'Writes case note
CALL write_variable_in_case_note("HH Comp Change Reported")
CALL write_bullet_and_variable_in_Case_Note("Unit member HH Member", HH_Member)
CALL write_bullet_and_variable_in_Case_Note("Date Reported/Addendum", Date_Reported)
CALL write_bullet_and_variable_in_Case_Note("Date effective", Effective_Date)
CALL write_bullet_and_variable_in_Case_Note("Action Taken", Action_Taken)
CALL write_bullet_and_variable_in_Case_Note("Additional Notes", Additional_Notes)

'checkboxes
IF Temporary_Change_Checkbox = 1 THEN CALL write_variable_in_Case_Note("* Change is temporary")

'writes case note for baby born
IF Baby_Born_Checkbox = 1 THEN

	CALL write_variable_in_Case_Note("--Client reports birth of baby--")
	CALL write_bullet_and_variable_in_Case_Note("Baby's name", Baby_Name)	
	CALL write_bullet_and_variable_in_Case_Note("Date of birth", Date_birth)
	CALL write_bullet_and_variable_in_Case_Note("Baby's SSN", Baby_SSN)
	CALL write_bullet_and_variable_in_Case_Note("Father's name", Father_Name)
	CALL write_bullet_and_variable_in_Case_Note("Father's employer", Father_Employer)
	CALL write_bullet_and_variable_in_Case_Note("Mother's employer", Mother_Employer)
	IF OHI_Checkbox = 1 THEN CALL write_bullet_and_variable_in_Case_Note("OHI", OHI_Source)
END IF

'signs case note
CALL write_variable_in_Case_Note("----")
CALL write_variable_in_Case_Note(worker_signature)

Script_End_Procedure ""
