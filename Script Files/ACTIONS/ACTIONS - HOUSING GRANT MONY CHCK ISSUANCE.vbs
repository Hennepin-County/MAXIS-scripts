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
STATS_manualtime = 269                	'manual run time in seconds
STATS_denomination = "C"       			' is for case
'END OF stats block=========================================================================================================			
	
'Date variables for current month -11'
CM_minus_11_mo =  left("0" &            DatePart("m",           DateAdd("m", -11, date)           ), 2)
CM_minus_11_yr =  right(                 DatePart("yyyy",        DateAdd("m", -11, date)           ), 2)

'DIALOG===========================================================================================================================
BeginDialog housing_grant_MONY_CHCK_issuance_dialog, 0, 0, 311, 200, "MFIP Housing Grant MONY/CHCK issuance "
  EditBox 60, 10, 55, 15, case_number
  EditBox 165, 10, 25, 15, member_number
  EditBox 245, 10, 25, 15, initial_month
  EditBox 275, 10, 25, 15, initial_year
  EditBox 55, 105, 245, 15, other_notes
  EditBox 75, 130, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 130, 50, 15
    CancelButton 250, 130, 50, 15
  Text 25, 55, 215, 30, "Snappy warning text to come later"
  Text 200, 15, 40, 10, "month/year:"
  Text 10, 135, 60, 10, "Worker signature:"
  GroupBox 10, 35, 290, 60, "MFIP Housing Grant MONY/CHCK Issuance:"
  Text 125, 15, 35, 10, "Member #:"
  Text 10, 15, 50, 10, "Case Number:"
  Text 10, 110, 45, 10, "Other notes:"
EndDialog

'The script============================================================================================================================
'Connects to MAXIS, grabbing the case case_number
EMConnect ""
Call MAXIS_case_number_finder(case_number) 
member_number = "01"	'defaults the member number to 01
initial_month = CM_mo  'defaulting date to current month and year
initial_year = CM_yr

'Main dialog: user will input case number and initial month/year if not already auto-filled 
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog housing_grant_MONY_CHCK_issuance_dialog				'main dialog'
		If buttonpressed = 0 THEN stopscript	'script ends if cancel is selected'
		IF len(case_number) > 8 or isnumeric(case_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid case number."					'mandatory field
		IF len(member_number) > 2 or isnumeric(member_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit member number."	'mandatory field'
		IF len(initial_month) > 2 or isnumeric(initial_month) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit month."	'mandatory field
		IF len(initial_year) > 2 or isnumeric(initial_year) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit year."		'mandatory field
		IF worker_signature = ""  then err_msg = err_msg & vbCr & "You must sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Clears out case number and enters the selected footer month/year
back_to_self
EMWritescreen "________", 18, 43
EMWritescreen case_number, 18, 43
EMWritescreen initial_month, 20, 43
EMWritescreen initial_year, 20, 46

'searching for the housing grant issued on the INQD screen(s) for the most current year
Call navigate_to_MAXIS_screen("MONY", "INQX")
EMWritescreen CM_minus_11_mo, 6, 38
EMWritescreen CM_minus_11_yr, 6, 41
EMWritescreen CM_plus_1_mo, 6, 53		
EMwritescreen CM_plus_1_yr, 6, 56
EMWriteScreen "x", 10, 5		'selecting MFIP
transmit

'checking to see if HG has been issued for the month selected	
DO
	row = 6				'establishing the row to start searching for issuance'
	DO
		EMReadScreen housing_grant, 2, row, 19		'searching for housing grant issuance
		If housing_grant = "  " then exit do
		IF housing_grant = "HG" then
			'reading the housing grant information
			EMReadScreen HG_amt_issued, 7, row, 40
			EMReadScreen HG_month, 2, row, 73
			EMReadScreen HG_year, 2, row, 79
			INQD_issuance = HG_month & HG_year
			month_of_issuance = initial_month & initial_year
			If month_of_issuance = INQD_issuance then script_end_procedure("Issuance has already been made on the month selected. Please review your case, and update manually.")
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
		EMReadscreen memb_number, 2, MAXIS_row, 6		'searching for member number
	IF memb_number = member_number then 				'exits do if member number matches
		exit do
	ELSE 
		MAXIS_row = MAXIS_row + 1	'otherwise it searches again on the next row
	END IF 
	If member_number = "  " then script_end_procedure("The member number you entered does not appear to be valid. Please check your member number and try again.")
LOOP until memb_number = member_number

EMWritescreen "x", MAXIS_row, 64			'selects the member number'
transmit
EMReadscreen emps_status, 2, 9, 22			'grabs the EMPS status code'
transmit

'grabs the coding to input in MONY/CHCK
Call navigate_to_MAXIS_screen("ELIG", "MFBF")
EMReadscreen member_code, 1, MAXIS_row, 27		
EMReadscreen cash_portion, 1, MAXIS_row, 37
EMReadScreen state_portion, 1, MAXIS_row, 54

'checking for sanctions, user will have to process manually if there's a sanction
EMReadScreen MFIP_sanction, 1, MAXIS_row, 68
If MFIP_sanction = "Y" then	script_end_procedure("A sanction exist for this member. Please check sanction for accuracy, and process manually.")

'checking for FIAT'd version that shows case is elig for the $110 housing grant
Call navigate_to_MAXIS_screen("ELIG", "MFSM")
EMReadScreen fiat_check, 4, 9, 31
EMReadScreen housing_grant_issued, 6, 16, 75
IF fiat_check <> "FIAT" and housing_grant_issued <> "110.00" then script_end_procedure("You must FIAT this case prior to issuing the MONY/CHCK. Please FIAT, then try again")

'navigates to MONY/CHCK and inputs codes into 1st screen
Call navigate_to_MAXIS_screen("MONY", "CHCK")
EMWriteScreen "MF", 5, 17
EMWriteScreen "MF", 5, 21
EMWriteScreen "31", 5, 32		'restored payment code per the HG bulletin
EMWriteScreen member_number, 7, 27
transmit 	

'now we're on the MFIP issuance detail pop-up screen
EMWriteScreen "01", 10, 6
EMWriteScreen member_code, 10, 14		'adds coding from MFBF into issuance detail screen
EMWriteScreen cash_portion, 10, 23 
EMWriteScreen state_portion, 10, 33
EMwritescreen "110.00", 10, 53
transmit
EMReadScreen ID_10_T_error_check, 7, 17, 4			'checking to make sure that 
IF ID_10_T_error_check = "HOUSING" then script_end_procedure ("Housing grant may have already been issued. Please recheck your case, and try again.")
EMWriteScreen "Y", 15, 52
transmit
EMWriteScreen "Y", 15, 29	
transmit
transmit 
transmit	'transmit three times to get to the restoration of benefits screen '
'writes in the manual check reason per the bulletin on the Housing Grant
EMWriteScreen "You meet one of the exceptions", 13, 18
EMWriteScreen "listed in CM 13.03.09 for families", 14, 18
EMWriteScreen "with an adult MFIP unit member(s)", 15, 18
EMWriteScreen "who get Section 8/HUD funded subsidy:", 16, 18
If emps_status = "02" or emps_status = "07" or emps_status = "12" or emps_status = "23" or emps_status = "27" or emps_status = "15" or emps_status = "18" or _
   emps_status = "30" or emps_status = "33" then
	EMWriteScreen "Caregivers who are elderly/disabled", 17, 18		'writes in disa/elderly if the codes above are the client's emps_status code
Else 
	EMWriteScreen "Caregivers caring for a disabled member", 17, 18
END IF 

'updating emps_status coding for case note'
If  emps_status = "02" then emps_status = "Age 60 or older"
If emps_status = "08" or emps_status = "24" then emps_status = "Care for Ill/incapacitated family member"
If emps_status = "07" or emps_status = "23" then emps_status = "Ill/incapacitated > 30 days" 
If emps_status = "12" or emps_status = "27" then emps_status = "Special medical criteria"
If emps_status = "15" or emps_status = "30" then emps_status = "Mentally Ill"
If emps_status = "18" or emps_status = "33" then emps_status = "SSI/RSDI pending"

'Case noting the MONY/CHCK info'
Call start_a_blank_case_note
Call write_variable_in_case_note("**MONY/CHCK ISSUED FOR HOUSING GRANT for " & initial_month & "/" & initial_year& "**")
Call write_variable_in_case_note("* Housing grant issued due to family meeting an exemption per CM.13.03.09.")
Call write_variable_in_case_note("* Member " & member_number & " exemption is: " & emps_status & ".")
Call write_variable_in_case_note("* Results for " & initial_month & "/" & initial_year & " have been FIATed in ELIG/MFIP.")
Call write_bullet_and_variable_in_case_note("Other notes", other_notes)
Call write_variable_in_case_note("--")
Call write_variable_in_case_note(worker_signature)

script_end_procedure("Success. The FIAT results have been generated. Please review before approving.")