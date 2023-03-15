'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - FIAT GA-RCA INTO SNAP.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 225                	'manual run time in seconds
STATS_denomination = "C"       			'C is for each Case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("02/03/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The script----------------------------------------------------------------------------------------------------
'Connecting to BlueZone and finding the MAXIS case number
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(initial_month, initial_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 226, 115, "FIAT GA/RCA into SNAP"
  EditBox 105, 10, 60, 15, MAXIS_case_number
  EditBox 105, 30, 25, 15, initial_month
  EditBox 140, 30, 25, 15, initial_year
  ButtonGroup ButtonPressed
    OkButton 75, 50, 50, 15
    CancelButton 130, 50, 50, 15
  Text 50, 15, 50, 10, "Case Number:"
  Text 5, 35, 100, 10, "Initial month/year of package:"
  GroupBox 5, 75, 215, 35, "FIAT GA/RCA into SNAP:"
  Text 10, 90, 210, 20, "This FIATer is to be used when a case is active on GA or RCA, and the monthly grant needs to be FIATed into the SNAP budget. "
EndDialog

DO
	DO
		err_msg = ""
		dialog Dialog1
		Cancel_without_confirmation
		If IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
		IF len(initial_month) > 2 or isnumeric(initial_month) = FALSE THEN err_msg = err_msg & vbCr & "* You must enter a valid 2 digit initial month."
		IF len(initial_year) > 2 or isnumeric(initial_year) = FALSE THEN err_msg = err_msg & vbCr & "* You must enter a valid 2 digit initial year."
		IF err_msg <> "" THEN msgbox err_msg & vbCr & "Please resolve to continue."
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Uses the custom function to create an array of dates from the initial_month and initial_year variables, ends at CM + 1.
	'We will need to remove the string "/1/" from each element in the array
call date_array_generator(initial_month, initial_year, footer_month_array)

'Need to make sure we start in the correct year for maxis
MAXIS_footer_month = initial_month
MAXIS_footer_year = initial_year

maxis_background_check				'ensures that case is out of background

'----------------------------------------------------------------------------------------------------STAT/PROG
'Needs the elig begin date for proration reasons, collect it from PROG
call navigate_to_maxis_screen("STAT", "PROG")

'Checking for PRIV cases.
EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case, script will end.
IF priv_check = "PRIVIL" THEN script_end_procedure("This case is a privliged case. You do not have access to this case.")

EMReadscreen proration_date, 8, 10, 44		'reading the SNAP proration date for ELIG/FS

'checking for active cash program and type of cash program
MAXIS_row = 6
DO
	EMReadScreen cash_prog, 2, MAXIS_row, 67
	If (cash_prog = "GA" or cash_prog = "RC") then
		exit do
	Else
		MAXIS_row = MAXIS_row + 1
	End if
LOOP until MAXIS_row = 8

If (cash_prog = "GA" or cash_prog = "RC") then
	EMReadScreen cash_status, 4, MAXIS_row, 74
Else
	script_end_procedure("This case is not active on GA or RCA. Please review case.")
End if

If cash_status <> "ACTV" then script_end_procedure("This case is not active on GA or RCA. Please review case.")
IF cash_prog = "RC" then cash_prog = "RCA"		'changes the program to the full orogram name

'Goes to STAT/REVW to check for a SNAP review date. If nissing this will cause an error in the FIAT----------------------------------------------------------------------------------------------------
Call navigate_to_maxis_screen("STAT", "REVW")
EMReadscreen REVW_date, 8, 9, 57
If REVW_date = "__ 01 __" then script_end_procedure("A SNAP review date is required. Please update STAT/REVW, then run the script again.")

'Reads the Hennepin Cash out code for x127 users. Elderly/SSI clients have the option to cash out SNAP benefits.
If worker_county_code = "x127" then
	Call navigate_to_maxis_screen("MONY", "DISB")
	EMReadscreen cash_out, 1, 14, 53
END IF

'The following loop will take the script through each month in the package, from appl month. to CM+1
For i = 0 to ubound(footer_month_array)
	MAXIS_footer_month = datepart("m", footer_month_array(i)) 'Need to assign footer month / year each time through
	if len(MAXIS_footer_month) = 1 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
	MAXIS_footer_year = right(datepart("YYYY", footer_month_array(i)), 2)

	'----------------------------------------------------------------------------------------------------ELIG
	Call navigate_to_MAXIS_screen("ELIG", cash_prog)
	transmit

	EMReadScreen no_CASH, 10, 24, 2
	If no_CASH = "NO VERSION" then script_end_procedure("There are no " & cash_prog & " results for this case. Please review case.")
	If cash_prog = "GA" then
		EMWriteScreen "99", 20, 78
	Else
		EMWriteScreen "99", 19, 78
	END IF
	transmit
	'This brings up the cash versions of eligibilty results to search for approved versions
	status_row = 7
	Do
		EMReadScreen app_status, 8, status_row, 50
		If trim(app_status) = "" then script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
		If app_status = "UNAPPROV" Then status_row = status_row + 1
	Loop until  app_status = "APPROVED" or trim(app_status) = ""

	If app_status <> "APPROVED" then
		script_end_procedure("No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". Please review case.")
	Elseif app_status = "APPROVED" then
		EMReadScreen vers_number, 1, status_row, 23
		EMWriteScreen vers_number, 18, 54
		transmit
	END IF

	IF cash_prog = "RCA" then
		EMWriteScreen "RCSM", 19, 70
		transmit
		EMReadscreen grant_amt, 7, 12, 74
	Else
		EMWriteScreen "GASM", 20, 70
		transmit
		EMReadscreen issue_header, 21, 12, 45
		If issue_header = "Amount Already Issued" then
			EMReadscreen grant_amt, 7, 12, 74
		Else
			EMReadscreen grant_amt, 7, 14, 74
		End if
	END IF

	grant_amt = trim(grant_amt)		'cleans up grant amount

	'----------------------------------------------------------------------------------------------------The FIAT
	back_to_self
	EMwritescreen "FIAT", 16, 43
	EMWritescreen MAXIS_case_number, 18, 43
	EMwritescreen MAXIS_footer_month, 20, 43
	EMWritescreen MAXIS_footer_year, 20, 46
	transmit

	'Checking for cases that are out of county
	EMReadScreen error_check, 11, 24, 2
	If error_check = "YOU ARE NOT" then script_end_procedure("You do not have access to update this case. The script will now end.")

	EMReadscreen results_check, 4, 14, 46 'We need to make sure results exist, otherwise stop.
	'IF results_check = "    " THEN script_end_procedure("The script was unable to find unapproved SNAP results for the benefit month, please check your case and try again.")
	EMWritescreen "22", 4, 34 'entering the FIAT reason
	EMWritescreen "x", 14, 22
	transmit 'This should take us to FFSL

	EMWriteScreen "x", 17, 5	'Going into 'VEIW BUDGET'
	transmit
	EMWriteScreen "x", 10, 5	'Going into the PA GRANT
	transmit

	If cash_prog = "GA" then
		EMWriteScreen "_________", 9, 23
		EMWriteScreen grant_amt, 9, 23
	Elseif cash_prog = "RCA" then
		EMWriteScreen "_________", 8, 23
		EMWriteScreen grant_amt, 8, 23
	END IF
	Transmit 'to get back to FFB1

	grant_amt = ""	'clears out the grant amt

	EMReadScreen warning_check, 4, 18, 9 'We need to check here for a warning on potential expedited cases..
	IF warning_check = "FIAT" Then 'and enter two extra transmits to bypass.
		transmit
		transmit
	END IF

	EMwritescreen "FFB2", 20, 70 'This is to make sure we end up in the right place'
	transmit

	'this enters the proration date in the initial month'
	IF abs(MAXIS_footer_month) = abs(left(proration_date, 2)) THEN
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

	EMwritescreen "FFSM", 20, 70 'Goes to the last screen in ELIG as this is a requirement of the FIAT
	transmit
	'Updates the Hennepin County Cashout (Y/N) field if cash out in MONY/DISB is coded as a Y
	If cash_out <> "N" then EMwritescreen cash_out, 19, 72

	'Exiting the FIAT
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

script_end_procedure("Success, the FIAT results have been generated. Please review before approving." & vbcr & vbcr & "Update the applicable client's WREG panel using the FSET/ABAWD coding hierarchy.")
