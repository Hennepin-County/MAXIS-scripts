'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - CS DISREGARD WCOM.vbs"
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
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 64                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

'>>>>> DIALOG <<<<<
BeginDialog case_number_dlg, 0, 0, 136, 80, "Case Number"
  EditBox 65, 10, 65, 15, case_number
  EditBox 85, 30, 20, 15, benefit_month
  EditBox 110, 30, 20, 15, benefit_year
  ButtonGroup ButtonPressed
    OkButton 30, 60, 50, 15
    CancelButton 80, 60, 50, 15
  Text 10, 15, 50, 10, "Case Number"
  Text 10, 35, 70, 10, "Benefit Month/Year"
EndDialog

'>>>>> THE SCRIPT <<<<<
EMConnect ""

'>>>>> CHECKING FOR MAXIS <<<<<
CALL check_for_MAXIS(True)

EMReadScreen at_self, 4, 2, 50
IF at_self = "SELF" THEN 
	EMReadScreen case_number, 8, 18, 43
	case_number = replace(case_number, "_", "")
	case_number = trim(case_number)
	EMReadScreen benefit_month, 2, 20, 43
	EMReadScreen benefit_year, 2, 20, 46
ELSE
	CALL find_variable("Case Nbr: ", case_number, 8)
	case_number = replace(case_number, "_", "")
	case_number = trim(case_number)
	CALL find_variable("Month: ", benefit_month, 2)
	CALL find_variable("Month: " & benefit_month & " ", benefit_year, 2)
END IF

DO
	err_msg = ""
	Dialog case_number_dlg
		cancel_confirmation
		IF case_number = "" 	THEN err_msg = err_msg & vbCr & "* Please enter a case number."
		IF benefit_month = "" 	THEN err_msg = err_msg & vbCr & "* Please enter a benefit month."
		IF benefit_year = "" 	THEN err_msg = err_msg & vbCr & "* Please enter a benefit year."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

call navigate_to_screen("spec", "wcom")
EMWriteScreen benefit_month, 3, 46
EMWriteScreen benefit_year, 3, 51
transmit

DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
	EMReadScreen more_pages, 8, 18, 72
	IF more_pages = "MORE:  -" THEN PF7
LOOP until more_pages <> "MORE:  -"

read_row = 7
DO
	EMReadscreen cash_program, 2, read_row, 26 
	EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
	If waiting_check = "Waiting" and cash_program = "MF" THEN 'checking program type and if it's been printed, needs more fool proofing
		EMWriteScreen "X", read_row, 13
		Transmit
		PF9
		EMWriteScreen "STARTING OCTOBER 1, 2015, A NEW LAW BEGINS THAT ALLOWS", 3, 15
      	EMWriteScreen "US TO NOT COUNT SOME OF THE CHILD SUPPORT YOU GET", 4, 15
		EMWriteScreen "WHEN DETERMINING YOUR MONTHLY MFIP/DWP BENEFIT AMOUNT:", 5, 15
		EMWriteScreen "", 6, 15
		EMWriteScreen "  - Up to $100 for an assistance unit with one child", 7, 15
		EMWriteScreen "  - Up to $200 for an assistance unit with two or more", 8, 15
		EMWriteScreen "     children", 9, 15
		EMWriteScreen "", 10, 15
		EMWriteScreen "BECAUSE OF THIS CHANGE, YOU MAY SEE AN INCREASE IN", 11, 15
		EMWriteScreen "YOUR BENEFIT AMOUNT.", 12, 15
	    PF4
		PF3
	END IF
	read_row = read_row + 1
	IF read_row = 18 THEN
		PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18??
		read_row = 7
	End if
LOOP until cash_program = "  "

script_end_procedure("")
