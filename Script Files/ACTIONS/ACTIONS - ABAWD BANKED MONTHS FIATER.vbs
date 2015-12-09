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

BeginDialog income_deductions_dialog, 0, 0, 326, 280, "ABAWD banked months income and deductions dialog"
  ButtonGroup ButtonPressed
    OkButton 260, 155, 50, 15
    CancelButton 260, 175, 50, 15
  EditBox 55, 45, 50, 15, gross_wages
  EditBox 55, 65, 50, 15, busi_income
  EditBox 55, 85, 50, 15, gross_RSDI
  EditBox 55, 105, 50, 15, gross_SSI
  EditBox 55, 125, 50, 15, gross_VA
  EditBox 55, 145, 50, 15, gross_UC
  EditBox 55, 165, 50, 15, gross_CS
  EditBox 55, 185, 50, 15, gross_other
  EditBox 185, 45, 35, 15, SHEL_rent
  EditBox 185, 65, 35, 15, SHEL_tax
  EditBox 185, 85, 35, 15, SHEL_insa
  EditBox 185, 105, 35, 15, SHEL_other
  EditBox 185, 125, 35, 15, deduction_FMED
  EditBox 275, 45, 35, 15, HEST_elec
  EditBox 275, 65, 35, 15, HEST_heat
  EditBox 275, 85, 35, 15, HEST_phone
  EditBox 275, 105, 35, 15, deduction_COEX
  EditBox 275, 125, 35, 15, deduction_DCEX
  ButtonGroup ButtonPressed
    PushButton 20, 15, 25, 10, "BUSI",  BUSI_button
    PushButton 45, 15, 25, 10, "JOBS", JOBS_button
    PushButton 70, 15, 25, 10, "RBIC", RBIC_button
    PushButton 95, 15, 25, 10, "SPON", SPON_button
    PushButton 120, 15, 25, 10, "UNEA", UNEA_button
    PushButton 175, 15, 25, 10, "COEX", COEX_button
    PushButton 200, 15, 25, 10, "DCEX", DCEX_button
    PushButton 225, 15, 25, 10, "FMED", FMED_button
    PushButton 250, 15, 25, 10, "HEST", HEST_button
    PushButton 275, 15, 25, 10, "SHEL", SHEL_button
    PushButton 130, 165, 45, 10, "prev. panel", prev_panel_button
    PushButton 130, 175, 45, 10, "next panel", next_panel_button
    PushButton 185, 165, 45, 10, "prev. memb", prev_memb_button
    PushButton 185, 175, 45, 10, "next memb", next_memb_button
  Text 35, 150, 15, 10, "UC:"
  Text 245, 130, 25, 10, "DCEX:"
  Text 30, 190, 20, 10, "Other:"
  Text 35, 170, 15, 10, "CS:"
  Text 155, 130, 25, 10, "FMED:"
  Text 240, 50, 30, 10, "Electric:"
  Text 245, 110, 25, 10, "COEX:"
  Text 35, 110, 15, 10, "SSI:"
  Text 230, 70, 40, 10, "Heating/air:"
  Text 125, 90, 60, 10, "House insurance:"
  Text 230, 90, 40, 10, "Telephone:"
  Text 130, 50, 50, 10, "Mortgage/rent:"
  Text 160, 110, 20, 10, "Other:"
  Text 30, 70, 20, 10, "BUSI:"
  Text 135, 70, 45, 10, "Property tax:"
  GroupBox 125, 155, 110, 35, "STAT-based navigation"
  GroupBox 170, 5, 135, 25, "Deduction based MAXIS panels:"
  Text 60, 35, 50, 10, "Gross Amount"
  Text 25, 35, 25, 10, "UI type"
  GroupBox 15, 5, 135, 25, "Income based MAXIS panels:"
  Text 30, 90, 20, 10, "RSDI:"
  GroupBox 15, 210, 300, 50, "BEFORE YOU HIT THE OK BUTTON"
  Text 20, 220, 285, 35, "The information pulled into the editboxes above are the amounts that are being FIATed into the SNAP budget in the selected budget month. Please use the navigation buttons on this dialog if you want to check what is listed on your MAXIS panels. If this informaiton is not corret, please press cancel now, and review your case.  "
  Text 35, 130, 10, 10, "VA:"
  Text 20, 50, 30, 10, "WAGES:"
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
	public SHEL_other
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


'The following performs case accuracy checks.
call navigate_to_maxis_screen("ELIG", "FS")
redim ABAWD_member_array(0)

For each member in hh_member_array
	row = 6
	col = 1
	EMSearch member, row, col 'Finding the row this member is on
	EMWritescreen "x", row, 5
	transmit 'Now on FFPR
	EMReadscreen inelig_test, 6, 6, 20 'This reads the ABAWD 3/36 month test
	IF inelig_test = "FAILED" THEN 'This member is failing this test, add them to the ABAWD member array
		If ABAWD_member_array(0) <> "" Then ReDim Preserve ABAWD_member_array(UBound(ABAWD_member_array)+1) 
		ABAWD_member_array(UBound(ABAWD_member_array)) = member
	END IF
	transmit
Next
IF ABAWD_member_array(0) = "" THEN script_end_procedure("ERROR: There are no members on this case with ineligible ABAWDs.  The script will stop.")

err_msg = ""
For each member in ABAWD_member_array 'This loop will check that WREG is coded correctly
	call navigate_to_maxis_screen("STAT", "WREG")
	EMWritescreen member, 20, 76
	Transmit
	EMReadscreen wreg_status, 2, 8, 50
	IF wreg_status <> "30" THEN err_msg = err_msg & vbCr & "Member " & member & " does not have FSET code 30."
	EMReadscreen abawd_status, 2, 13, 50
	IF abawd_status <> "10" THEN err_msg = err_msg & vbCr & "Member " & member & " does not have ABAWD code 10."
	'This section pulls up the counted months popup and checks for 3 months counted before Jan. 16
	EmWriteScreen "x", 13, 57 
	transmit
	bene_mo_col = 55
	bene_yr_row = 8
    abawd_counted_months = 0
    second_abawd_period = 0
 	DO 'This loop actually reads every month in the time period
  	    EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
  		IF is_counted_month = "X" or is_counted_month = "M" THEN abawd_counted_months = abawd_counted_months + 1
		IF is_counted_month = "Y" or is_counted_month = "N" THEN second_abawd_period = second_abawd_period + 1
   		bene_mo_col = bene_mo_col + 4
    		IF bene_mo_col > 63 THEN
        		bene_yr_row = bene_yr_row + 1
   	     		bene_mo_col = 19
   	   	    END IF
   	LOOP until bene_yr_row = 11 'Stops when it reaches 2016
  	IF abawd_counted_months < 3 THEN err_msg = err_msg & vbCr & "Member " & member & " does not have 3 ABAWD months coded before 01/2016"
	row = 11
	col = 19
	EMSearch "M", row, col 'This looks to make sure there is an intial banked month coded on WREG.
	IF row > 11 THEN err_msg = err_msg & vbCr & "Member " & member & " does not have an initial banked month coded on WREG."
	PF3
Next

IF err_msg <> "" THEN 'This means the WREG panel(s) are coded incorrectly.
	msgbox "Please resolve the following errors before continuing. The script will now stop." & vBcr & err_msg
	script_end_procedure("")
END IF

	

'The following loop will take the script throught each month in the package, from appl month. to CM+1
Do
	footer_month = datepart("m", current_month)
	if len(footer_month) = 1 THEN footer_month = "0" & footer_month
	footer_year = right(datepart("YYYY", current_month), 2)


	'for each member in hh_member_array
		'go to UNEA and read SNAP PIC for each thing
	For i = 0 to ubound(HH_member_array)

	For i = 0 to ubound(ABAWD_counted_months)
		Set ABAWD_months_array(i) = new ABAWD_month_data
		Call navigate_to_MAXIS_screen("STAT", "SHEL")		'<<<<< Goes to SHEL for this person
		EMWriteScreen HH_member_array(i), 20, 76 
		EMReadScreen rent_verif, 2, 11, 67
		If rent_verif <> "__" and rent_verif <> "NO" and rent_verif <> "?_" then EMReadScreen rent, 8, 11, 56
		If rent_verif = "__" or rent_verif = "NO" or rent_verif = "?_" then rent = "0"		'<<<<< Gets rent amount
		EMReadScreen lot_rent_verif, 2, 12, 67
		If lot_rent_verif <> "__" and lot_rent_verif <> "NO" and lot_rent_verif <> "?_" then EMReadScreen lot_rent, 8, 12, 56
		If lot_rent_verif = "__" or lot_rent_verif = "NO" or lot_rent_verif = "?_" then lot_rent = "0"		'<<<<< gets Lot Rent amount
		EMReadScreen mortgage_verif, 2, 13, 67
		If mortgage_verif <> "__" and mortgage_verif <> "NO" and mortgage_verif <> "?_" then EMReadScreen mortgage, 8, 13, 56
		If mortgage_verif = "__" or mortgage_verif = "NO" or mortgage_verif = "?_" then mortgage = "0"		'<<<<<< gets Mortgage amount
		EMReadScreen insurance_verif, 2, 14, 67
		If insurance_verif <> "__" and insurance_verif <> "NO" and insurance_verif <> "?_" then EMReadScreen insurance, 8, 14, 56
		If insurance_verif = "__" or insurance_verif = "NO" or insurance_verif = "?_" then SHEL_insa = "0"	'<<<<<< gets insurance amount and adds it to the class property
		EMReadScreen taxes_verif, 2, 15, 67
		If taxes_verif <> "__" and taxes_verif <> "NO" and taxes_verif <> "?_" then EMReadScreen taxes, 8, 15, 56
		If taxes_verif = "__" or taxes_verif = "NO" or taxes_verif = "?_" then SHEL_taxes = "0"				'<<<<<<< gets taxes amount and adds it to the class property
		EMReadScreen room_verif, 2, 16, 67
		If room_verif <> "__" and room_verif <> "NO" and room_verif <> "?_" then EMReadScreen room, 8, 16, 56
		If room_verif = "__" or room_verif = "NO" or room_verif = "?_" then room = "0"						'<<<<<<< gets room/board amount
		EMReadScreen garage_verif, 2, 17, 67
		If garage_verif <> "__" and garage_verif <> "NO" and garage_verif <> "?_" then EMReadScreen garage, 8, 17, 56
		If garage_verif = "__" or garage_verif = "NO" or garage_verif = "?_" then garage = "0"				'<<<<<<< gets garage amount
		SHEL_rent = cint(rent) + cint(mortgage)						'<<<<<<  Adds rent amount and mortage amount together to get the Rent line for elig and adds to Class property 
		SHEL_other = cint(lot_rent) + cint(room) + cint(garage) 	'<<<<<<  Adds lot rent, room, and garage amounts together to get the Other line for elig and adds to Class property
		'///// Needs to navigate to next month
	Next

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
