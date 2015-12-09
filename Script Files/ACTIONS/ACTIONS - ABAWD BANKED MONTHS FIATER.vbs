'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - ABAWD BANKED MONTHS FIATER.vbs"
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

'Defining variables----------------------------------------------------------------------------------------------------
'Dim gross_wages, busi_income, gross_RSDI, gross_SSI, gross_VA, gross_UC, gross_CS, gross_other
'Dim deduction_FMED, deduction_DCEX, deduction_COEX

'Dialogs----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 251, 230, "ABAWD BANKED MONTHS FIATER"
  EditBox 105, 10, 60, 15, case_number
  EditBox 105, 30, 25, 15, initial_month
  EditBox 140, 30, 25, 15, initial_year
  ButtonGroup ButtonPressed
    OkButton 75, 50, 50, 15
    CancelButton 130, 50, 50, 15
  Text 50, 15, 50, 10, "Case Number:"
  Text 5, 35, 100, 10, "Initial month/year of package:"
  Text 30, 90, 200, 25, "This script will FIAT eligibility results, income and deductions for each HH member with pending SNAP results for months where ABAWD banked months are being used. "
  GroupBox 20, 75, 215, 70, "Per Bulletin #15-01-01 SNAP banked month policy/procedures:"
  Text 30, 170, 200, 10, "* All STAT panels must be updated before using this script."
  Text 30, 190, 200, 20, "* Do NOT mark partial counted months with an "M". Partial months are not counted, only full months are counted."
  Text 30, 125, 200, 20, "If you are unsure of how/why/when you should be applying this process, please refer to the Bulletin."
  GroupBox 20, 155, 215, 60, "Before you begin:"
EndDialog

'----------------------DEFINING CLASSES WE'LL NEED FOR THIS SCRIPT
class ABAWD_month_data
	public gross_Wages
	public BUSI_income
	public gross_RSDI
	public gross_SSI
	public gross_VA
	public gross_UC
	public gross_CS
	public gross_other
	public deduction_FMED
	public deduction_DCEX
	public deduction_COEX
	public SHEL_rent
	public SHEL_tax
	public SHEL_insa
	public HEST_elect
	public HEST_heat
	public HEST_phone
end class

'-------------------------END CLASSES

'VARIABLES WE'LL NEED TO DECLARE (NOTE, IT'S LIKELY THESE WILL NEED TO MOVE FURTHER DOWN IN THE SCRIPT)----------------------------
ABAWD_counted_months = 1	'<<<<<<<<<<<THIS IS TEMPORARY AND SHOULD BE READ ELSEWHERE, TO FIGURE OUT HOW MANY MONTHS WE NEED

'Create an array of all the counted months
DIM ABAWD_months_array()	'Minus one because arrays
REDIM ABAWD_months_array(ABAWD_counted_months - 1)	'Minus one because arrays

'The script----------------------------------------------------------------------------------------------------
EMConnect ""
call check_for_maxis(false)

call maxis_case_number_finder(case_number)

DO
	err_msg = ""
	dialog case_number_dialog
	If buttonpressed = 0 THEN stopscript
	IF isnumeric(case_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid case number."
	IF len(initial_month) > 2 or isnumeric(initial_month) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit initial month."
	IF len(initial_year) > 2 or isnumeric(initial_year) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit initial year."
	IF err_msg <> "" THEN msgbox err_msg & vbCr & "Please resolve to continue."
LOOP UNTIL err_msg = ""


check_for_maxis(true)
'Create hh_member_array
call HH_member_custom_dialog(HH_member_array)


'defining necessary dates
initial_date = initial_month & "/01/" & initial_year
current_month = initial_date
current_month_plus_one = dateadd("m", 1, date)
maxis_background_check
'The following loop will take the script throught each month in the package, from appl month. to CM+1
Do
	footer_month = datepart("m", current_month)
	if len(footer_month) = 1 THEN footer_month = "0" & footer_month
	footer_year = right(datepart("YYYY", current_month), 2)


	'for each member in hh_member_array
		'go to UNEA and read SNAP PIC for each thing






		'<<<<<<<<<<<<<SAMPLE IDEA FOR ARRAY'
		For i = 0 to ubound(ABAWD_counted_months)
			'Defines the ABAWD_months_array as an obejct of ABAWD month data'
			set ABAWD_months_array(i) = new ABAWD_month_data
			'>>>>NAVIGATE TO WHERE YOU NEED TO GO'
			EMReadScreen x, 8, 18, 56	'<<<<READ THE STUFF'
			ABAWD_months_array(i).gross_RSDI = x	'<<<<ADD THE STUFF TO THE ARRAY'
			'>>>>>>DO THE ABOVE TWO LINES OVER AND OVER AGAIN UNTIL YOU HAVE ALL THE STUFF FOR THIS MONTH'
			'//// <<<<<<GET TO THE NEXT MONTH AT THE END'
		Next
		'<<<<<<<<<<<<<<<<<END SAMPLE'





		'if member > 18 go to JOBS and read SNAP PIC
		'if member > 18 go to BUSI and read SNAP PIC
		'if member > 18 go to RBIC and read SNAP PIC
		'go to COEX and read deductions
		'go to DCEX and read deductions

	'Sum up gross income
	'background check
	'Go to FIAT
	back_to_self
	EMwritescreen "FIAT", 16, 43
	EMWritescreen case_number, 18, 43
	EMwritescreen footer_month, 20, 43
	EMWritescreen footer_year, 20, 46
	transmit
	EMReadscreen results_check, 4, 14, 46 'We need to make sure results exist, otherwise stop.
	IF results_check = "    " THEN script_end_procedure("The script was unable to find unapproved SNAP results for the benefit month, please check your case and try again.")
	EMWritescreen "03", 4, 34 'entering the FIAT reason
	EMWritescreen "x", 14, 22
	transmit 'This should take us to FFSL
	'The following loop will enter person tests screen and pass for each member on grant
	For each member in hh_member_array
		row = 6
		col = 1
		EMSearch member, row, col 'Finding the row this member is on
		EMWritescreen "x", row, 5
		transmit 'Now on FFPR
		EMWritescreen "PASSED", 9, 12
		transmit
		PF3 'back to FFSL
	Next
	'Ready to head into case test / budget screens
	DO 'This is in a loop, because sometimes FIAT has a glitch that won't let it exit.
		EMWritescreen "x", 16, 5
		EMWritescreen "x", 17, 5
		Transmit
		'Passing all case tests
		EMWritescreen "PASSED", 10, 7
		EMWritescreen "PASSED", 13, 7
		EMWritescreen "PASSED", 14, 7
		PF3
		'Now the BUDGET (FFB1) NO
		EMWritescreen gross_wages, 5, 32
		EMWritescreen busi_income, 6, 32
		EMWritescreen gross_RSDI, 11, 32
		EMWritescreen gross_SSI, 12, 32
		EMWritescreen gross_VA, 13, 32
		EMWritescreen gross_UC, 14, 32
		EMWritescreen gross_CS, 15, 32
		EMWritescreen gross_other, 16, 32
		EMWritescreen deduction_FMED, 12, 72
		EMWritescreen deduction_DCEX, 13, 72
		EMWritescreen deduction_COEX, 14, 72
		transmit
		'Now on FFB2
		EMWritescreen SHEL_rent, 5, 29
		EMWritescreen SHEL_tax, 6, 29
		EMWritescreen SHEL_insa, 7, 29
		EMWritescreen HEST_elect, 8, 29
		EMWritescreen HEST_heat, 9, 29
		EMWritescreen HEST_phone, 10, 29
		'Does hennepin cashout matter?
		transmit
		'Now on SUMM screen, which shouldn't matter
		PF3 'back to FFSL
		PF3 'This should bring up the "do you want to retain" popup
		EMReadscreen budget_error_check, 6, 24, 2 'This will be "budget" if MAXIS had a glitch, and will need to loop through again.
	LOOP Until budget_error_check = ""
	EMWritescreen "Y", 13, 41
	transmit
	EMReadscreen final_month_check, 4, 10, 53 'This looks for a popup that only comes up in the final month, and clears it.
	IF final_month_check = "ELIG" THEN
		EMWritescreen "Y", 11, 52
		EMWritescreen initial_month, 13, 37
		EMWritescreen initial_year, 13, 40
		transmit
		Exit DO
	END IF
	'IF datepart("m", current_month) = datepart("m", current_month_plus_one) THEN exit DO
	current_month = dateadd("m", 1, current_month)
	msgbox datediff("m", current_month_plus_one, current_month)
Loop Until datediff("m", current_month_plus_one, current_month) > 0

script_end_procedure("Success. The FIAT results have been generated. Please review before approving.")
