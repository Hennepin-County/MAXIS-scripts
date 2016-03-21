'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - HOUSING GRANT MONY CHCK ISSUANCE.vbs"
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
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 180                	'manual run time in seconds
STATS_denomination = "I"       			'I is for item
'END OF stats block=========================================================================================================			
							
'DIALOG===========================================================================================================================
BeginDialog housing_grant_MONY_CHCK_issuance_dialog, 0, 0, 311, 200, "MFIP Housing Grant MONY/CHCK issuance "
  EditBox 65, 10, 60, 15, case_number
  EditBox 195, 10, 25, 15, member_number
  EditBox 70, 30, 25, 15, initial_month
  EditBox 100, 30, 25, 15, initial_year
  EditBox 65, 180, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 180, 50, 15
    CancelButton 215, 180, 50, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 15, 80, 240, 30, "Snappy warning text to come later"
  Text 10, 35, 55, 10, "month/year:"
  Text 10, 180, 50, 10, "Worker signature:"
  GroupBox 5, 60, 260, 60, "MFIP Housing Grant MONY/CHCK Issuance:"
  Text 145, 15, 40, 10, "Member #:"
EndDialog

'The script============================================================================================================================
'Connects to MAXIS, grabbing the case case_number
EMConnect ""
Call MAXIS_case_number_finder(case_number)


'Main dialog: user will input case number and initial month/year if not already auto-filled 
DO
	DO
		err_msg = ""							'establishing value of varaible, this is necessary for the Do...LOOP
		dialog housing_grant_MONY_CHCK_issuance_dialog				'main dialog'
		If buttonpressed = 0 THEN stopscript	'script ends if cancel is selected'
		IF len(case_number) > 8 or isnumeric(case_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid case number."					'mandatory field
		IF len(member_number) > 2 or isnumeric(member_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit member number."	
		IF len(initial_month) > 2 or isnumeric(initial_month) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit first month."	'mandatory field
		IF len(initial_year) > 2 or isnumeric(initial_year) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit first year."		'mandatory field
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'checking the MFIP ELIG 
back_to_self
EMWritescreen "________", 18, 43
EMWritescreen initial_month, 20, 43
EMWritescreen initial_year, 20, 46

'searching for the housing grant issued on the INQD screen(s)
Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWritescreen initial_month, 6, 38
EMWritescreen initial_year, 6, 41
EMWritescreen initial_month, 6, 53
EMwritescreen initial_year, 6, 56
EMWriteScreen "x", 10, 5		'selecting MFIP
transmit
	
DO
	row = 6
	DO
		EMReadScreen housing_grant, 2, row, 19		'searching for housing grant issuance
		If housing_grant = "__" then exit do
		IF housing_grant = "HG" then
			'reading the housing grant information
			EMReadScreen HG_amt_issued, 7, row, 40
			EMReadScreen HG_month, 2, row, 73
			EMReadScreen HG_year, 2, row, 77
			INQD_issuance = HG_month & HG_year
			month_of_issuance = initial_month & initial_year
			If month_of_issuance = INQD_issuance then script_end_procedure("Issuance has already been made on at least one month selected. Please review your case, and update manually.")
			Else
				row = row + 1
			END IF		
	Loop until row = 18				'repeats until the end of the page
		PF8
		EMReadScreen last_page_check, 21, 24, 2
		If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

'goes into ELIG/MFIP and checks for sanctions and a FIATED version of the month selected'
Call navigate_to_MAXIS_screen("ELIG", "MFIP")
DO 
	MAXIS_row = 7
		EMReadscreen memb_number, 2, MAXIS_row, 6
	IF memb_number = member_number then 
		exit do
	ELSE 
		MAXIS_row = MAXIS_row + 1
	END IF 
	If member_number = "" then script_end_procedure("The member number you entered does not appear to be valid. Please check your member number and try again.")
LOOP until memb_number = member_number

EMWritescreen "x", MAXIS_row, 64
transmit
EMReadscreen emps_status, 2, 9, 22

Call navigate_to_MAXIS_screen("ELIG", "MFBF")
EMReadscreen cash_portion, 1, MAXIS_row, 37
EMReadScreen food_portion, 1, MAXIS_row, 45
EMReadScreen state_portion, 1, MAXIS_row, 
'checking for sanctions
EMReadScreen MFIP_sanction, 1, MAXIS_row, 68
If MFIP_sanction = "Y" then	script_end_procedure("A sanction exist for this member. Please check sanction for accuracy, and process manually.")
	
Call navigate_to_MAXIS_screen("ELIG", "MFSM")
EMReadScreen fiat_check, 4, 9, 31
EMReadScreen housing_grant_issued, 6, 16, 75
IF fiat_check <> "FIAT" and housing_grant_issued <> "110.00" then 
	script_end_procedure("You must FIAT this case prior to issuing the MONY/CHCK. Please FIAT, then try again")
ELSE 
	

	
	




	



'The following loop will take the script through each month in the package, from appl month. to CM+1
For i = 0 to ubound(footer_month_array)				'array of footer months
	MAXIS_footer_month = datepart("m", footer_month_array(i)) 'Need to assign footer month / year each time through
	if len(MAXIS_footer_month) = 1 THEN MAXIS_footer_month = "0" & MAXIS_footer_month		'adds a 0 if footer month is a single digit
	MAXIS_footer_year = right(datepart("YYYY", footer_month_array(i)), 2)			'users the last 2 digits of the footer year

	'-----------------GO TO FIAT!---------------------------------
	back_to_self						'entering the footer month/year and navigating to FIAT'
	EMwritescreen "FIAT", 16, 43
	EMWritescreen case_number, 18, 43
	EMwritescreen MAXIS_footer_month, 20, 43
	EMWritescreen MAXIS_footer_year, 20, 46
	transmit
	EMReadscreen results_check, 4, 9, 46 'We need to make sure results exist, otherwise stop.
	IF results_check = "    " THEN script_end_procedure("The script was unable to find unapproved MFIP results for the benefit month, please check your case and try again.")
	EMWritescreen "03", 4, 34 'entering the FIAT reason
	EMWritescreen "x", 9, 22
	transmit 'This should take us to FMSL

	'Selects View Case Budget.
	EMwritescreen "x", 18, 4
	transmit
	'Selects the Subsidy/Tribal pop-up then the Housing Subsidy sub-pop-up
	EMwritescreen "x", 17, 5
	transmit
	EMwritescreen "x", 8, 13
	transmit
	'Changes the prospective column to $0
	EMwritescreen "0       ", 8, 51
	transmit
	transmit
	transmit
	'Reading to ensure the housing grant is in budget
	EMReadScreen MFIP_grant_confirmation, 6, 15, 75
	If MFIP_grant_confirmation <> "110.00" then 
		script_end_procedure("An issued occurred during the FIAT process. Please process manually.") 
	ELSE
		PF3
		PF3
		EMWritescreen "Y", 13, 41
		transmit
		STATS_counter = STATS_counter + 1  'adds one instance to the stats counter, counting each month as it's own run
	END IF
	EMReadscreen final_month_check, 4, 10, 53 'This looks for a pop-up that only comes up in the final month, and clears it.
	IF final_month_check = "ELIG" THEN
		EMWritescreen "Y", 11, 52
		EMWritescreen initial_month, 13, 37
		EMWritescreen right(initial_year, 2), 13, 40
		transmit
	END IF
NEXT

STATS_counter = STATS_counter - 1 	'removes one instance since one is counted at the start
msgbox STATS_counter
script_end_procedure("Success. The FIAT results have been generated. Please review before approving.")