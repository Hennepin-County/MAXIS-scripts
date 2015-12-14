'Created by Tim DeLong from Stearns County.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - HOUSING GRANT FIATER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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

'DIALOG
'===========================================================================================================================
BeginDialog HG_Fiater, 0, 0, 161, 95, "Housing Grant FIATer"
  EditBox 60, 5, 90, 15, case_number
  EditBox 60, 25, 20, 15, footer_month
  EditBox 130, 25, 20, 15, footer_year
  EditBox 75, 50, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 35, 70, 50, 15
    CancelButton 90, 70, 50, 15
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 30, 50, 10, "Footer Month:"
  Text 85, 30, 45, 10, "Footer Year:"
  Text 10, 55, 60, 10, "Worker Signature:"
EndDialog


'============================================================================================================================

EMConnect ""

'Finds the case number
Call MAXIS_case_number_finder(case_number)

'Finds the benefit month
EMReadScreen on_SELF, 4, 2, 50
IF on_SELF = "SELF" THEN
	CALL find_variable("Benefit Period (MM YY): ", footer_month, 2)
	IF footer_month <> "" THEN CALL find_variable("Benefit Period (MM YY): " & footer_month & " ", footer_year, 2)
ELSE
	CALL find_variable("Month: ", footer_month, 2)
	IF footer_month <> "" THEN CALL find_variable("Month: " & footer_month & " ", footer_year, 2)
END IF

'Warning/instruction box
MsgBox "This script is intended for use for MFIP cases only." & vbNewLine & vbNewLine &_
		"To be eligible for the Housing Grant the case must meet one of the criteria listed below:" & vbNewLine & vbNewline &_
		vbTab & "EMPS status of:" & vbNewLine &_
		vbtab & "      02 - Age > or = 60" & vbNewLine &_
		vbtab & "      07 - Ill/Incap > 30 days" & vbNewLine &_
		vbTab & "      08 - Care of Ill/Incapacitated Family Member" & vbNewLine &_
		vbTab & "      12 - Special Med Criteria" & vbNewLine &_
		vbTab & "      21 - Age 60 or Older (UP)" & vbNewLine &_
		vbTab & "      23 - Ill/Incap > 30 Days (UP)" & vbNewLine &_
		vbTab & "      24 - Care Ill/Incap Family Member (UP)" & vbNewLine &_
		vbTab & "      27 - Special Med Criteria(UP)" & vbNewLine & vbNewLine &_
		"For the following EMPS statuses, confirm that DISA has been coded as greater than 30 days." & vbNewLine & vbNewLine &_
		vbTab & "EMPS status of:" & vbNewLine &_
		vbTab & "      15 - Mentally Ill" & vbNewLine &_
		vbTab & "      18 - SSI/RSDI Pending" & vbNewLine &_
		vbTab & "      30 - Mentally Ill (UP)" & vbNewLine &_
		vbTab & "      33  SSI/RSDI Pending (UP)" 

check_for_maxis(False)

DO
	DO
		err_msg = ""
		'starts the Housing Grant FIATer dialog
		Dialog HG_FIATer
		'asks if you want to cancel and if "yes" is selected sends StopScript
		cancel_confirmation 
		'checks that there is a case number
		IF case_number = FALSE THEN err_msg = err_msg & vbCr & "You must enter a case number."
		'checks if the footer month has been entered
		IF footer_month = "" THEN err_msg = err_msg & vbCr & "You must enter the footer month."
		'checks if the footer year has been entered
		IF footer_year = "" THEN err_msg = err_msg & vbCr & "You must enter the footer year."
		'checks that the case note was signed
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "You must sign your case note!" 
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false			

check_for_maxis(False)

back_to_self

'starting at requested month
EMwritescreen footer_month, 20, 43
EMwritescreen footer_year, 20, 46

'Navigates to FIAT and selects MFIP.
CALL navigate_to_MAXIS_screen("FIAT", "____")
	EMwritescreen "03", 4, 34
	EMwritescreen "x", 9, 22
	transmit

'Selects View Case Budget.
	EMwritescreen "x", 18, 4
	transmit

'Selects the Subsidy/Tribal popup then the Housing Subsidy sub-popup
	EMwritescreen "x", 17, 5
	transmit
	EMwritescreen "x", 8, 13
	transmit

'Changes the prospective column to $0
	EMwritescreen "0       ", 8, 51
	transmit
	transmit
	transmit

'script ends where the worker can see if the housing grant is showing as eligible and pops up a msg box to do so.	
script_end_procedure ("Verify that the results showing are what were expected." & vbNewline & vbNewline &_
	"If results are correct, PF3 twice to exit FIAT then retain results." & vbNewline & vbNewline &_
	"Run the script for any other months needed and approve.")
