'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - HC ICAMA.vbs"
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

'Required for statistical purposes===========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block==========================================================================================================
 
'THE DIALOG----------------------------------------------------------------------------------------------------------

BeginDialog HC_ICAMA_dialog, 0, 0, 286, 380, "HC ICAMA"
  EditBox 80, 10, 65, 15, case_number
  EditBox 140, 30, 75, 15, icama_recd
  EditBox 60, 55, 160, 15, state
  DropListBox 120, 80, 100, 15, "Select One..."+chr(9)+"Adoption"+chr(9)+"Foster Care", type_dropdown
  EditBox 105, 105, 115, 15, fc100
  EditBox 105, 130, 115, 15, fcarep
  EditBox 125, 155, 95, 15, ma_requested
  EditBox 140, 180, 80, 15, aa_payment
  EditBox 100, 205, 120, 15, ohc
  EditBox 90, 230, 130, 15, pmap_ex
  EditBox 125, 255, 95, 15, faxed_date
  CheckBox 35, 280, 95, 15, "MA Coverage Form Sent", coverage_checkbox
  CheckBox 35, 300, 245, 15, "Navigate to DAIL/WRIT to Create a TIKL to Approve next 6 Month Budget", tikl_checkbox
  EditBox 90, 325, 130, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 355, 50, 15
    CancelButton 220, 355, 50, 15
  Text 10, 15, 70, 10, "Maxis Case Number:"
  Text 35, 210, 65, 10, "Other Health Care:"
  Text 35, 85, 85, 10, "Adoption or Foster Care:"
  Text 35, 235, 55, 10, "PMAP Excluded:"
  Text 35, 160, 90, 10, "Date Requested MA Open:"
  Text 35, 260, 90, 10, "Faxed ICAMA 6.03 to DHS:"
  Text 35, 110, 65, 10, "FC 100 A/B Rec'd:"
  Text 35, 35, 100, 10, "Health Care ICAMA 6.01 Rec'd:"
  Text 15, 330, 75, 10, "Sign Your Case Note:"
  Text 35, 135, 65, 10, "AREP (Foster Care):"
  Text 35, 185, 100, 10, "Adoption Assistance Payment:"
  Text 35, 60, 20, 10, "State:"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------

'Connects to BLUEZONE
EMConnect ""

'Grabs the MAXIS case number            
CALL MAXIS_case_number_finder(case_number)

'Shows dialog
DO
	err_msg = ""		
	Dialog HC_ICAMA_dialog
	IF ButtonPressed = 0 THEN StopScript
	IF IsNumeric(case_number) = FALSE THEN err_msg = err_msg & vbCr & "* You must type a valid numeric case number."
	IF type_dropdown = "Select One..." THEN err_msg = err_msg & vbCr & "* You must select Adoption or Foster Care!"
	IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your case note!"
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Opens new case note
start_a_blank_case_note

'Writes the Case Note
CALL write_variable_in_case_note("*** HC ICAMA ***")
CALL write_bullet_and_variable_in_case_note("HC ICAMA 6.01 Rec'd", icama_recd)
CALL write_bullet_and_variable_in_case_note("State", state)
CALL write_bullet_and_variable_in_case_note("ADOPT or FC", type_dropdown)
CALL write_bullet_and_variable_in_case_note("FC 100 A & B Rec'd", fc100)
CALL write_bullet_and_variable_in_case_note("AREP (FC)", fcarep)
CALL write_bullet_and_variable_in_case_note("Date Requested MA Open", ma_requested)
CALL write_bullet_and_variable_in_case_note("AA Payment", aa_payment)
CALL write_bullet_and_variable_in_case_note("OHC", ohc)
CALL write_bullet_and_variable_in_case_note("PMAP Excluded", pmap_ex)
CALL write_variable_in_case_note("---------------------------------")
CALL write_bullet_and_variable_in_case_note("Faxed ICAMA 6.03 to DHS", faxed_date)
IF coverage_checkbox = checked THEN CALL write_variable_in_case_note("* MA Coverage form was Sent")        
IF tikl_checkbox = checked THEN CALL write_variable_in_case_note("* TIKL created to approve next 6 Month Budget")
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'If we checked to TIKL out, it goes to DAIL/WRIT and pulls up a blank TIKL
IF tikl_checkbox = 1 THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
END IF

script_end_procedure("")
