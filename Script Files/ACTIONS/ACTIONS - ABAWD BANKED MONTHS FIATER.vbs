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

'Required for statistical purposes===========================================================================================	
STATS_counter = 1				'sets the stats counter at one	
STATS_manualtime = 225			'manual run time in seconds	
STATS_denomination = "C"        'C is for each case	
'END OF stats block==========================================================================================================	
	
'-------------------------------FUNCTIONS WE INVENTED THAT WILL SOON BE ADDED TO FUNCLIB
FUNCTION date_array_generator(initial_month, initial_year, date_array)
	'defines an intial date from the initial_month and initial_year parameters
	initial_date = initial_month & "/1/" & initial_year
	'defines a date_list, which starts with just the initial date
	date_list = initial_date

	'This loop creates a list of dates
	Do
		If datediff("m", date, initial_date) = 1 then exit do		'if initial date is the current month plus one then it exits the do as to not loop for eternity'
		working_date = dateadd("m", 1, right(date_list, len(date_list) - InStrRev(date_list,"|")))	'the working_date is the last-added date + 1 month. We use dateadd, then grab the rightmost characters after the "|" delimiter, which we determine the location of using InStrRev
		date_list = date_list & "|" & working_date	'Adds the working_date to the date_list
	Loop until datediff("m", date, working_date) = 1	'Loops until we're at current month plus one

	'Splits this into an array
	date_array = split(date_list, "|")
End function

'This prompts the user to determine whether an item with verification code ? should be counted on expedited cases
'verif is the variable that was read from maxis, verif_name is the name that will be displayed to user '
FUNCTION verif_confirm_message(verif, verif_name)
	IF verif = "?_" or verif = "?" THEN verif_confirm =  msgbox("The " & verif_name & " verification is marked '?' Do you wish to count this amount? Click yes if this is an expedited case and the unverified amount should be budgeted.", vbYesNo)
	IF verif_confirm = vbYes THEN verif = "OT"
END FUNCTION
'-------------------------------END FUNCTIONS

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
  Text 30, 190, 200, 20, "* Do NOT mark partial counted months with an ""M"". Partial months are not counted, only full months are counted."
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

'The script----------------------------------------------------------------------------------------------------
EMConnect ""
call check_for_maxis(false)
call maxis_case_number_finder(case_number)

If worker_county_code = "x101" OR _
		worker_county_code = "x115" OR _
		worker_county_code = "x129" OR _
		worker_county_code = "x133" OR _
		worker_county_code = "x136" OR _
		worker_county_code = "x139" THEN
		script_end_procedure ("Your agency is exempt from ABAWD work requirements. SNAP banked months are not available to your recipients.")
END IF

DO
	err_msg = ""
	dialog case_number_dialog
	If buttonpressed = 0 THEN stopscript
	IF left(case_number, 10) <> "UUDDLRLRBA" AND isnumeric(case_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid case number."
	IF len(initial_month) > 2 or isnumeric(initial_month) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit initial month."
	IF len(initial_year) > 2 or isnumeric(initial_year) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit initial year."
	IF err_msg <> "" THEN msgbox err_msg & vbCr & "Please resolve to continue."
LOOP UNTIL err_msg = ""

'THIS ACTIVATES DEVELOPER MODE if Konami code is entered left of the case number
IF left(case_number, 10) = "UUDDLRLRBA" THEN developer_mode = true
case_number = replace(case_number, "UUDDLRLRBA", "") 'removing it so the case number works in the rest of the script

'Uses the custom function to create an array of dates from the initial_month and initial_year variables, ends at CM + 1.
	'We will need to remove the string "/1/" from each element in the array
call date_array_generator(initial_month, initial_year, footer_month_array)

'Create an array of all the counted months
DIM ABAWD_months_array()	'Minus one because arrays
REDIM ABAWD_months_array(ubound(footer_month_array))	'Minus one because arrays

check_for_maxis(true)

'Need to make sure we start in the correct year for maxis'
footer_month = initial_month
footer_year = initial_year

'Create hh_member_array
call HH_member_custom_dialog(HH_member_array)

maxis_background_check

'Needs the elig begin date for proration reasons, collect it from PROG
call navigate_to_maxis_screen("STAT", "PROG")
EMReadscreen proration_date, 8, 10, 44

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
IF developer_mode = false THEN
	IF ABAWD_member_array(0) = "" THEN script_end_procedure("ERROR: There are no members on this case with ineligible ABAWDs.  The script will stop.")
ELSE
	IF ABAWD_member_array(0) = "" THEN msgbox "ERROR: There are no members on this case with ineligible ABAWDs.  The script would stop in production mode."
END IF
err_msg = ""
For each member in ABAWD_member_array 'This loop will check that WREG is coded correctly
	call navigate_to_maxis_screen("STAT", "WREG")
	EMWritescreen member, 20, 76
	Transmit
	EMReadscreen wreg_status, 2, 8, 50
	IF wreg_status <> "30" THEN err_msg = err_msg & vbCr & "Member " & member & " does not have FSET code 30."
	EMReadscreen abawd_status, 2, 13, 50
	IF abawd_status <> "10" THEN err_msg = err_msg & vbCr & "Member " & member & " does not have ABAWD code 10."
	'-----This section pulls up the counted months popup and checks for 3 months counted before Jan. 16----
	EmWriteScreen "x", 13, 57
	transmit
	'This finds the current month location '
	current_month_col = 15 + 4*cint(initial_month)

	'This finds the current year row
	current_yr_row = 6
	col = 9
	if len(initial_year) = 2 THEN initial_year = "20" & initial_year
	EMSearch initial_year, current_yr_row, col
	'Setting the initial month to look at as current month - 36'
	bene_mo_col = current_month_col
	bene_yr_row = current_yr_row - 3
    WREG_months = 0
    second_abawd_period = 0
 	DO 'This loop actually reads every month in the time period
  	    EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
  	IF is_counted_month = "X" or is_counted_month = "M" THEN WREG_months = WREG_months + 1
		IF is_counted_month = "Y" or is_counted_month = "N" THEN second_abawd_period = second_abawd_period + 1
   		bene_mo_col = bene_mo_col + 4
    IF bene_mo_col > 63 THEN
        		bene_yr_row = bene_yr_row + 1
   	     		bene_mo_col = 19
   	END IF
  LOOP until bene_yr_row = current_yr_row AND bene_mo_col = current_month_col'Stops when it reaches current year & month.
  IF WREG_months < 3 THEN err_msg = err_msg & vbCr & "Member " & member & " has not used 3 ABAWD months in the past 3 years."
	row = current_yr_row
	col = current_month_col
	EMSearch "M", row, col 'This looks to make sure there is an initial banked month coded on WREG on or after the initial month
	IF row < current_yr_row OR row > current_yr_row + 1 THEN err_msg = err_msg & vbCr & "Member " & member & " does not have an initial banked month coded on WREG."
	PF3
Next

IF err_msg <> "" THEN 'This means the WREG panel(s) are coded incorrectly.
	IF developer_mode = true THEN
		msgbox "The following errors were found on the WREG panel. In production mode the script would stop." & vBcr & err_msg
	ELSE
		msgbox "Please resolve the following errors before continuing. The script will now stop." & vBcr & err_msg
		script_end_procedure("")
	END IF
END IF

'The following loop will take the script through each month in the package, from appl month. to CM+1
For i = 0 to ubound(footer_month_array)
	footer_month = datepart("m", footer_month_array(i)) 'Need to assign footer month / year each time through
	if len(footer_month) = 1 THEN footer_month = "0" & footer_month
	footer_year = right(datepart("YYYY", footer_month_array(i)), 2)

	Set ABAWD_months_array(i) = new ABAWD_month_data
	
	Call navigate_to_MAXIS_screen("STAT", "RBIC")	'error handling for RBIC cases 
	EMWriteScreen HH_member, 20, 76					
	transmit
	EMReadScreen RBIC_total, 1, 2, 78				'reading to see if RBIC screens exist
	If RBIC_total <> "0" then 
		script_end_procedure("The script does not currently support cases with RBIC income. Please process this case manually.")
	END IF

	Call navigate_to_MAXIS_screen("STAT", "HEST")		'<<<<< Navigates to STAT/HEST
	EMReadScreen HEST_heat, 6, 13, 75 					'<<<<< Pulls information from the prospective side of HEAT/AC standard allowance
	IF HEST_heat <> "      " then						'<<<<< If there is an amount on the hest line then the electric and phone allowances are not used
		HEST_elect = ""
		HEST_phone = ""				'<<<<< Ignores the electric and phone standards if HEAT/AC is used
	Else
		EMReadScreen HEST_elect, 6, 14, 75				'<<<<< Pulls information from prospective side of Electric standard if HEAT/AC is not used
		EMReadScreen HEST_phone, 6, 15, 75				'<<<<< Pulls information from prospective side of Phone standard if HEAT/AC is not used
	End If

	For each hh_member in HH_member_array
		Call navigate_to_MAXIS_screen("STAT", "SHEL")		'<<<<< Goes to SHEL for this person
		EMWriteScreen hh_member, 20, 76
		transmit
		EMReadScreen rent_verif, 2, 11, 67
		call verif_confirm_message(rent_verif, "rent")
		If rent_verif <> "__" and rent_verif <> "NO" and rent_verif <> "?_" then EMReadScreen rent, 8, 11, 56
		If rent_verif = "__" or rent_verif = "NO" or rent_verif = "?_" then rent = "0"		'<<<<< Gets rent amount
		EMReadScreen lot_rent_verif, 2, 12, 67
		call verif_confirm_message(lot_rent_verif, "lot rent")
		If lot_rent_verif <> "__" and lot_rent_verif <> "NO" and lot_rent_verif <> "?_" then EMReadScreen lot_rent, 8, 12, 56
		If lot_rent_verif = "__" or lot_rent_verif = "NO" or lot_rent_verif = "?_" then lot_rent = "0"		'<<<<< gets Lot Rent amount
		EMReadScreen mortgage_verif, 2, 13, 67
		call verif_confirm_message(mortgage_verif, "mortgage")
		If mortgage_verif <> "__" and mortgage_verif <> "NO" and mortgage_verif <> "?_" then EMReadScreen mortgage, 8, 13, 56
		If mortgage_verif = "__" or mortgage_verif = "NO" or mortgage_verif = "?_" then mortgage = "0"		'<<<<<< gets Mortgage amount
		EMReadScreen insurance_verif, 2, 14, 67
		call verif_confirm_message(insurance_verif, "insurance")
		If insurance_verif <> "__" and insurance_verif <> "NO" and insurance_verif <> "?_" then EMReadScreen insurance, 8, 14, 56
		If insurance_verif = "__" or insurance_verif = "NO" or insurance_verif = "?_" then insurance = "0"	'<<<<<< gets insurance amount and adds it to the class property
		EMReadScreen taxes_verif, 2, 15, 67
		call verif_confirm_message(taxes_verif, "taxes")
		If taxes_verif <> "__" and taxes_verif <> "NO" and taxes_verif <> "?_" then EMReadScreen taxes, 8, 15, 56
		If taxes_verif = "__" or taxes_verif = "NO" or taxes_verif = "?_" then taxes = "0"				'<<<<<<< gets taxes amount and adds it to the class property
		EMReadScreen room_verif, 2, 16, 67
		call verif_confirm_message(room_verif, "room")
		If room_verif <> "__" and room_verif <> "NO" and room_verif <> "?_" then EMReadScreen room, 8, 16, 56
		If room_verif = "__" or room_verif = "NO" or room_verif = "?_" then room = "0"						'<<<<<<< gets room/board amount
		EMReadScreen garage_verif, 2, 17, 67
		call verif_confirm_message(garage_verif, "garage")
		If garage_verif <> "__" and garage_verif <> "NO" and garage_verif <> "?_" then EMReadScreen garage, 8, 17, 56
		If garage_verif = "__" or garage_verif = "NO" or garage_verif = "?_" then garage = "0"				'<<<<<<< gets garage amount
		total_taxes = total_taxes + abs(taxes)
		total_insurance = abs(total_insurance) +	abs(insurance)
		total_rent = total_rent +	abs(rent) + abs(mortgage)						'<<<<<<  Adds rent amount and mortage amount together to get the Rent line for elig and adds to Class property
		shel_other = shel_other + abs(lot_rent) + abs(room) + abs(garage) 	'<<<<<<  Adds lot rent, room, and garage amounts together to get the Other line for elig and adds to Class property
	Next

	'COEX expenses/deductions
	For each HH_member in HH_member_array
		Call navigate_to_MAXIS_screen("STAT", "COEX")	'navs to COEX
		EMWriteScreen HH_member, 20, 76					'writes the HH_member variable
		transmit
		EMReadScreen COEX_total, 1, 2, 78				'reading to see if COEX screens exist
		If COEX_total = "0" then
			total_COEX_deduction = "0"							'if not, sets the variable to blank
		ELSEIF COEX_total <> "0" then
			'Support
			EMReadScreen support_amt, 8, 10, 63				'repeats the above steps for support_amt
			EMReadScreen support_ver, 1, 10, 36
			If support_ver = "?" or support_ver = "N" then support_amt = "0"
			If support_amt = "________" then support_amt = "0"
			support_amt = replace(support_amt, "_", "")
			'Alimony
			EMReadScreen alimony_amt, 8, 11, 63				'repeats the above steps for alimony_amt
			EMReadScreen alimony_ver, 1, 11, 36
			If alimony_ver = "?" or alimony_ver = "N" then alimony_amt = "0"
			If alimony_amt = "________" then alimony_amt = "0"
			alimony_amt = replace(alimony_amt, "_", "")
			'tax dependent
			EMReadScreen tax_dep_amt, 8, 12, 63				'repeats the above steps for tax_dep_amt
			EMReadScreen tax_dep_ver, 1, 12, 36
			If tax_dep_ver = "?" or tax_dep_ver = "N" then tax_dep_amt = "0"
			If tax_dep_amt = "________" then tax_dep_amt = "0"
			tax_dep_amt = replace(tax_dep_amt, "_", "")
			'Other COEX
			EMReadScreen other_COEX_amt, 8, 13, 63				'repeats the above steps for other_COEX_amt
			EMReadScreen other_COEX_ver, 1, 13, 36
			If tax_dep_ver = "?" or tax_dep_ver = "N" then other_COEX_amt = "0"
			If other_COEX_amt = "________" then other_COEX_amt = "0"
			other_COEX_amt = replace(other_COEX_amt, "_", "")
		END IF
		total_COEX_deduction = abs(support_amt) + abs(alimony_amt) + abs(tax_dep_amt) + abs(other_COEX_amt)
		ABAWD_months_array(i).deduction_COEX = ABAWD_months_array(i).deduction_COEX + total_COEX_deduction 'adds the current total_COEX_deduction amt to the array
	NEXT

	'FMED expenses/deductions
	For each HH_member in HH_member_array
		Call navigate_to_MAXIS_screen("STAT", "FMED")	'navs to FMED
		EMWriteScreen HH_member, 20, 76				'writes the HH_member variable
		transmit
		fmed_row = 9 									'Setting this variable for the do...loop
		EMReadScreen fmed_total, 1, 2, 78				'reading to see if fmed screens exist
		If fmed_total = "0" then
			fmed_total_amt = "0"	'if not, sets the variable to blank
		ELSE
		fmed_total_amt = "0"
		Do
			use_expense = False				'Used to determine if an FMED expense that has an end date is going to be counted.
			EMReadScreen fmed_type, 2, fmed_row, 25	'reading FMED information
			EMReadScreen fmed_proof, 2, fmed_row, 32
			EMReadScreen fmed_amt, 8, fmed_row, 70
			EMReadScreen fmed_end_date, 5, fmed_row, 60		'reading end date to see if this one even gets added
			IF fmed_end_date <> "__ __" THEN
				EMReadScreen fmed_footer_month, 2, 20, 55
				EMReadScreen fmed_footer_year, 2, 20, 58
				fmed_current_date = fmed_footer_month & "/01/" & fmed_footer_year
				fmed_end_date = replace(fmed_end_date, " ", "/01/")
				If fmed_end_date = fmed_current_date then
					use_expense = True
				ElseIf fmed_end_date < fmed_current_date then
					use_expense = False
				END IF
			END IF
			If fmed_end_date = "__ __" OR use_expense = TRUE then	'Skips entries with an end date, or end dates in the past.
				If fmed_proof <> "__" or fmed_proof <> "?_" or fmed_proof <> "NO" then
					use_expense = True
				ELSE
					fmed_amt = trim(fmed_amt)
				END IF
			End if
			If fmed_type = "12" then 'for mileage rate deduction information
				EmReadscreen milage_rate, 8, 17, 70
				milage_rate = trim(milage_rate)
				If milage_rate <> "" then fmed_amt = milage_rate
			END IF
			fmed_amt = replace(fmed_amt, "_", "")
			If fmed_amt = "" then fmed_amt = "0"
			If use_expense = False then fmed_amt = "0"
			fmed_row = fmed_row + 1			're-establishing this variable for the next do...loop
			fmed_total_amt = fmed_total_amt + abs(fmed_amt)	'adds the fmed_amt to the fmed_total_amt
			If fmed_row = 15 then
				PF20							'if at the end of the screen PF20 will shift PF8 to the next FMED screen
				fmed_row = 9					're-established this variable if PF20 was used.
				EMReadScreen last_page_check, 21, 24, 2		'checking for the last page edit
				If last_page_check <> "THIS IS THE LAST PAGE" then last_page_check = ""
			End if
		Loop until fmed_type = "__" or last_page_check = "THIS IS THE LAST PAGE"	'repeats all actions within the loop until the conditions are met
		End if
	  fmed_total_amt = abs(fmed_total_amt) - "35"		'remove $35 benchmark from total
		If fmed_total_amt < 0 Then fmed_total_amt = 0 'makes sure we don't end up with a negative deduction, as it becomes positive in the budget.
		ABAWD_months_array(i).deduction_FMED = ABAWD_months_array(i).deduction_FMED	= fmed_total_amt 'creates the total to FIAT into elig for FMED
	Next

'//////////// Going to pull UNEA information
	For each HH_member in HH_member_array
		Call navigate_to_MAXIS_screen("STAT", "UNEA")	'<<<<< Goes to UNEA for this person
		EMWriteScreen HH_member, 20, 76
		EMReadScreen number_of_unea_panels, 1, 2, 78
		For k = 1 to number_of_unea_panels		'<<<<<< Starting at 1 because this is a panel count and it makes sense to use this as a standard count
			EMWriteScreen "0" & k, 20, 79
			transmit
			EMReadScreen unea_type, 2, 5, 37 	'<<<<<< Reads each type of UNEA panel and adds the amounts together within a type
			EMReadscreen unea_verif, 1, 5, 65
			call verif_confirm_message(unea_verif, "Unearned income")
			IF unea_verif <> "?" THEN
				If unea_type = "01" OR unea_type = "02" then '<<<<<< RSDI
					EMWriteScreen "x", 10, 26
					transmit
					EMReadScreen rsdi_amount, 8, 18, 56
					gross_RSDI = abs(gross_RSDI) + abs(rsdi_amount)
					transmit
					ElseIf unea_type = "03" then 				'<<<<<< SSI
					EMWriteScreen "x", 10, 26
					transmit
					EMReadScreen ssi_amount, 8, 18, 56
					gross_SSI = abs(gross_SSI) + abs(ssi_amount)
					transmit
					ElseIf unea_type = "11" OR unea_type = "12" OR unea_type = "13" OR unea_type = "38" then 	'<<<<<< VA
					EMWriteScreen "x", 10, 26
					transmit
					EMReadScreen va_amount, 8, 18, 56
					gross_VA = abs(gross_VA) + abs(va_amount)
					transmit
					ElseIf unea_type = "14" then 				'<<<<<< UC
					EMWriteScreen "x", 10, 26
					transmit
					EMReadScreen uc_amount, 8, 18, 56
					gross_UC = abs(gross_UC) + abs(uc_amount)
					transmit
					ElseIf unea_type = "08" OR unea_type = "36" OR unea_type = "39" then 	'<<<<<< CS
					EMWriteScreen "x", 10, 26
					transmit
					EMReadScreen cs_amount, 8, 18, 56
					gross_CS = abs(gross_CS) + abs(cs_amount)
					transmit
					ElseIf unea_type = "06" OR unea_type = "15" OR unea_type = "16" OR unea_type = "17" OR unea_type = "18" OR unea_type = "23" OR unea_type = "24" OR unea_type = "25" OR unea_type = "26" OR unea_type = "27" OR unea_type = "28" OR unea_type = "29" OR unea_type = "31" OR unea_type = "35" OR unea_type = "37" OR unea_type = "40" then 	'<<<<<< Other UNEA
					EMWriteScreen "x", 10, 26
					transmit
					EMReadScreen other_unea_amount, 8, 18, 56
					gross_other = abs(gross_other) + abs(other_unea_amount)
					transmit
					End If
				END IF
		Next

		Call navigate_to_MAXIS_screen("STAT", "JOBS")		'<<<<< Goes to JOBS for this person
		EMWriteScreen HH_member, 20, 76
		Transmit
		EMReadScreen number_of_jobs_panels, 1, 2, 78
		IF number_of_jobs_panels <> "0" THEN
			For m = 1 to number_of_jobs_panels					'<<<<<< Starting at 1 because this is a panel count and it makes sense to use this as a standard count
				EMWriteScreen "0" & m, 20, 79
				transmit
				EMReadScreen jobs_type, 1, 5, 38
				EMReadScreen jobs_subsidy, 2, 5, 71
				EMReadScreen jobs_verified, 1, 6, 38
				EMReadScreen job_end_date, 8, 9, 49
				call verif_confirm_message(jobs_verified, "job")
				income_counted = vbYes 'defaults to counted, the next statement will confirm certain income types'
				IF job_end_date <> "__ __ __" THEN 'THere is a job end, determine whether it is prior to current footer month, if yes, don't count it.
					IF datediff("d", footer_month & "/01/" & footer_year, replace(job_end_date, " ", "/")) < 0 THEN income_counted = false
				END IF
				IF jobs_verified = "?" or jobs_verified = "_?" THEN income_counted = false 'this will set to not counted if the user selected no on the verif_confirm popup'
				If jobs_type = "J" OR jobs_type = "E" OR jobs_type = "O" OR jobs_type = "I" OR jobs_type = "M" OR jobs_type = "C" OR jobs_subsidy = "01" OR jobs_subsidy = "02" OR jobs_subsidy = "03" OR jobs_verified = "N" OR jobs_verified = "_" then
					'certain rare income types can not be automatically determined, this prompts the user to confirm yes/no to reduce errors.
					income_counted = MsgBox("This script does not support this type of Job" & vbCr & "Please select whether this income should be included in the FIAT budget for SNAP.", vbYesNo)
				END IF
					IF jobs_type = "F" or jobs_type = "S" or jobs_type = "G" THEN income_counted = vbNo 'Work study is not counted, but listed on PIC, so we don't want to read the pic.
					IF income_counted = vbYes	THEN			'<<<<< Gets prospective income from the SNAP PIC
					EMWriteScreen "x", 19, 38
					transmit
					EMReadScreen income_amount, 8, 18, 56
					PF3
					IF isnumeric(income_amount) = false THEN income_amount = 0
					jobs_income = abs(jobs_income) + income_amount	'<<<<< Combines all jobs income
					End If
				Next
			END IF

		Call navigate_to_MAXIS_screen("STAT", "BUSI")		'<<<<<< Same HH member - checking BUSI
		EMWriteScreen HH_member, 20, 76
		transmit
		EMReadScreen number_of_busi_panels, 1, 2, 78 		'<<<<<< Will go to all BUSI panels for this person
		For n = 1 to number_of_busi_panels
			EMWriteScreen "0" & n, 20, 79
			transmit
			EMWriteScreen "x", 6, 26
			transmit
			EMReadScreen busi_verif, 1, 11, 73
			PF3
			call verif_confirm_message(busi_verif, "Self Employment Income")
			If busi_verif = "?" OR busi_verif = "N" OR busi_verif = "_" then '<<<<< Will not count unverified income
				busi_amount = "0"
			Else
				EMReadScreen busi_amount, 8, 10, 69
			End If
			gross_BUSI = abs(gross_BUSI) + abs(busi_amount)		'<<<<<< Combining all busi income together
		Next
	Next

	'ABAWD_months_array(i).gross_wages = cstr(jobs_income)
	'storing all total amounts / adding trims so they read correctly in dialog
	jobs_income = trim(jobs_income)
	total_taxes = trim(total_taxes)
	total_insurance = trim(total_insurance)
 	total_rent = trim(total_rent)
	shel_other = trim(shel_other)
	gross_RSDI = trim(gross_RSDI)
	gross_SSI = trim(gross_SSI)
	gross_VA = trim(gross_VA)
	gross_UC = trim(gross_UC)
	gross_CS = trim(gross_CS)
	gross_other = trim(gross_other)
	gross_BUSI = trim(gross_BUSI)
	total_COEX_deduction = trim(total_COEX_deduction)
	fmed_total_amt = trim(fmed_total_amt)

 '------INCOME and deductions dialog, created here so that the class/properties carry into the dialog each month.-------- '
		BeginDialog income_deductions_dialog, 0, 0, 326, 280, "ABAWD banked months income and deductions dialog"
	  ButtonGroup ButtonPressed
	    OkButton 260, 155, 50, 15
	    CancelButton 260, 175, 50, 15
	  EditBox 55, 45, 50, 15, jobs_income
	  EditBox 55, 65, 50, 15, gross_BUSI
	  EditBox 55, 85, 50, 15, gross_RSDI
	  EditBox 55, 105, 50, 15, gross_SSI
	  EditBox 55, 125, 50, 15, gross_VA
	  EditBox 55, 145, 50, 15, gross_UC
	  EditBox 55, 165, 50, 15, gross_CS
	  EditBox 55, 185, 50, 15, gross_other
	  EditBox 185, 45, 35, 15, total_rent
	  EditBox 185, 65, 35, 15, total_taxes
	  EditBox 185, 85, 35, 15, total_insurance
	  EditBox 185, 105, 35, 15, shel_other
	  EditBox 185, 125, 35, 15, fmed_total_amt
	  EditBox 275, 45, 35, 15, HEST_elect
	  EditBox 275, 65, 35, 15, HEST_heat
	  EditBox 275, 85, 35, 15, HEST_phone
	  EditBox 275, 105, 35, 15, total_COEX_deduction
	 	ButtonGroup ButtonPressed
	    PushButton 20, 15, 25, 10, "BUSI",  BUSI_button
	    PushButton 45, 15, 25, 10, "JOBS", JOBS_button
	    PushButton 70, 15, 25, 10, "RBIC", RBIC_button
	    PushButton 95, 15, 25, 10, "SPON", SPON_button
	    PushButton 120, 15, 25, 10, "UNEA", UNEA_button
	    PushButton 175, 15, 25, 10, "COEX", COEX_button
	    PushButton 200, 15, 25, 10, "FMED", FMED_button
	    PushButton 225, 15, 25, 10, "HEST", HEST_button
	    PushButton 250, 15, 25, 10, "SHEL", SHEL_button
	    PushButton 130, 165, 45, 10, "prev. panel", prev_panel_button
	    PushButton 130, 175, 45, 10, "next panel", next_panel_button
	    PushButton 185, 165, 45, 10, "prev. memb", prev_memb_button
	    PushButton 185, 175, 45, 10, "next memb", next_memb_button
	  Text 35, 150, 15, 10, "UC:"
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
	  Text 20, 220, 285, 35, "The information pulled into the edit boxes above are the amounts that are being FIATed into the SNAP budget in the selected budget month. Please use the navigation buttons on this dialog if you want to check what is listed on your MAXIS panels. If this informaiton is not corret, please press cancel now, and review your case.  "
	  Text 35, 130, 10, 10, "VA:"
	  Text 20, 50, 30, 10, "WAGES:"
	EndDialog

  'This calls the dialog to allow worke r to confirm
	DO
		dialog income_deductions_dialog
		cancel_confirmation
		MAXIS_dialog_navigation
	Loop until buttonpressed = ok

	ABAWD_months_array(i).SHEL_rent = total_rent
	ABAWD_months_array(i).gross_wages = jobs_income
	ABAWD_months_array(i).gross_RSDI = gross_RSDI	'<<<<<<< Stores variables in the class property
	ABAWD_months_array(i).gross_SSI = gross_SSI
	ABAWD_months_array(i).gross_VA = gross_VA
	ABAWD_months_array(i).gross_UC = gross_UC
	ABAWD_months_array(i).gross_CS = gross_CS
	ABAWD_months_array(i).gross_other = gross_other
	ABAWD_months_array(i).BUSI_income = gross_BUSI
	ABAWD_months_array(i).deduction_COEX = total_COEX_deduction
	ABAWD_months_array(i).deduction_FMED = fmed_total_amt
 	ABAWD_months_array(i).SHEL_tax = total_taxes
 	ABAWD_months_array(i).SHEL_insa = total_insurance
 	ABAWD_months_array(i).HEST_elect = HEST_elect
 	ABAWD_months_array(i).HEST_heat = HEST_heat
 	ABAWD_months_array(i).HEST_phone = HEST_phone

	jobs_income = 0
	gross_BUSI = 0
	gross_RSDI = 0
	gross_SSI = 0
	gross_VA = 0
	gross_UC = 0
	gross_CS = 0
	gross_other = 0
	total_rent = 0
	total_taxes = 0
	total_insurance = 0
	shel_other = 0
	fmed_total_amt = 0
	HEST_elect = 0
	HEST_heat = 0
	HEST_phone = 0
	total_COEX_deduction = 0	'<<<<<< Resets the variables for the next cycle of this function

	'-----------------GO TO FIAT!---------------------------------
	back_to_self
	EMwritescreen "FIAT", 16, 43
	EMWritescreen case_number, 18, 43
	EMwritescreen footer_month, 20, 43
	EMWritescreen footer_year, 20, 46
	transmit
	EMReadscreen results_check, 4, 14, 46 'We need to make sure results exist, otherwise stop.
	'IF results_check = "    " THEN script_end_procedure("The script was unable to find unapproved SNAP results for the benefit month, please check your case and try again.")
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
		EMReadscreen state_food_check, 1, 7, 58 'We need to enter something here if it is blank'
		IF state_food_check <> "N" or state_food_check <> "Y" THEN EMwritescreen "N", 7, 58
		transmit
		PF3 'back to FFSL
	Next
	'Ready to head into case test / budget screens

		EMWritescreen "x", 16, 5
		EMWritescreen "x", 17, 5
		Transmit
		'Passing all case tests
		EMWritescreen "PASSED", 10, 7
		EMWritescreen "PASSED", 13, 7
		Transmit
		EMReadscreen net_check, 3, 24, 40 'sometimes this test needs to be passed, sometimes n/a.  the transmit triggers an error msg if it needs to pass this'
		IF net_check = "NET" THEN EMWritescreen "PASSED", 14, 7
		PF3
		'Now the BUDGET (FFB1) NO
		'First, blank out existing values to avoid an error from existing info
		EMWriteScreen "         ", 5, 32
		EMWriteScreen "         ", 6, 32
		EMWriteScreen "         ", 11, 32
		EMWriteScreen "         ", 12, 32
		EMWriteScreen "         ", 13, 32
		EMWriteScreen "         ", 14, 32
		EMWriteScreen "         ", 15, 32
		EMWriteScreen "         ", 16, 32
		EMWriteScreen "         ", 12, 72
		EMWriteScreen "         ", 13, 72
		EMWriteScreen "         ", 14, 72
		EMWritescreen ABAWD_months_array(i).gross_wages, 5, 32
		EMWritescreen ABAWD_months_array(i).busi_income, 6, 32
		EMWritescreen ABAWD_months_array(i).gross_RSDI, 11, 32
		EMWritescreen ABAWD_months_array(i).gross_SSI, 12, 32
		EMWritescreen ABAWD_months_array(i).gross_VA, 13, 32
		EMWritescreen ABAWD_months_array(i).gross_UC, 14, 32
		EMWritescreen ABAWD_months_array(i).gross_CS, 15, 32
		EMWritescreen ABAWD_months_array(i).gross_other, 16, 32
		EMWritescreen ABAWD_months_array(i).deduction_FMED, 12, 72
		EMWritescreen ABAWD_months_array(i).deduction_COEX, 14, 72
		transmit
		EMReadScreen warning_check, 4, 18, 9 'We need to check here for a warning on potential expedited cases..
		IF warning_check = "FIAT" Then 'and enter two extra transmits to bypass.
			transmit
			transmit
		END IF

		'Now on FFB2
		EMWriteScreen "         ", 5, 29
		EMWriteScreen "         ", 6, 29
		EMWriteScreen "         ", 7, 29
		EMWriteScreen "         ", 8, 29
		EMWriteScreen "         ", 9, 29
		EMWriteScreen "         ", 10, 29
		EMWritescreen ABAWD_months_array(i).SHEL_rent, 5, 29
		EMWritescreen ABAWD_months_array(i).SHEL_tax, 6, 29
		EMWritescreen ABAWD_months_array(i).SHEL_insa, 7, 29
		EMWritescreen ABAWD_months_array(i).HEST_elect, 8, 29
		EMWritescreen ABAWD_months_array(i).HEST_heat, 9, 29
		EMWritescreen ABAWD_months_array(i).HEST_phone, 10, 29
		EMWriteScreen ABAWD_months_array(i).SHEL_other, 12, 29
		'this enters the proration date in the initial month'
		IF abs(footer_month) = abs(left(proration_date, 2)) THEN
			EMWriteScreen left(proration_date, 2), 11, 56
			EMWriteScreen mid(proration_date, 4, 2), 11, 59
			EMWriteScreen right(proration_date, 2), 11, 62
		END IF
		transmit
		EMReadScreen warning_check, 4, 18, 9 'We need to check here for a warning on potential expedited cases..
		IF warning_check = "FIAT" Then 'and enter two extra transmits to bypass.
			transmit
			transmit
		END IF
		'Now on SUMM screen, which shouldn't matter
		PF3 'back to FFSL
		PF3 'This should bring up the "do you want to retain" popup
		EMReadScreen income_cap_check, 11, 24, 2
		If income_cap_check = "PROSP GROSS" then script_end_procedure("Prospective gross income is over the income standard. THE FIAT cannot be saved. Please review case and budget for potential errors.")
		EMWritescreen "Y", 13, 41
		transmit
		EMReadscreen final_month_check, 4, 10, 53 'This looks for a pop-up that only comes up in the final month, and clears it.
		IF final_month_check = "ELIG" THEN
			EMWritescreen "Y", 11, 52
			EMWritescreen initial_month, 13, 37
			EMWritescreen right(initial_year, 2), 13, 40
			transmit
		END IF
next

script_end_procedure("Success. The FIAT results have been generated. Please review before approving.")
