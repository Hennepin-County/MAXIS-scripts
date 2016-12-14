'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - HG EXPANSION MONY-CHCK.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 240               	'manual run time in seconds
STATS_denomination = "C"       			' is for case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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

call changelog_update("12/14/2016", "Updated handling for signficant change cases, and for cases that have exceed the issuance amount (and require a supervisor to approve the housing grant supplment.)", "Ilse Ferris, Hennepin County")
call changelog_update("12/08/2016", "Updated handling for exiting the TIME panel, confirming version number and MFBF panel, added handling for migrant indicator on MONY/CHCK. Also added comments to code, and removed outdated coding.", "Ilse Ferris, Hennepin County")
call changelog_update("12/01/2016", "Added ACTIONS script that will create a MONY/CHCK for cases that meet the Housing Grant expansion criteria.", "Ilse Ferris, Hennepin County")
call changelog_update("12/01/2016", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Date variables
'current month -1
CM_minus_1_mo =  right("0" &          	 DatePart("m",           DateAdd("m", -1, date)            ), 2)
CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
'current month - 11
CM_minus_11_mo =  left("0" &            DatePart("m",           DateAdd("m", -11, date)           ), 2)
CM_minus_11_yr =  right(                 DatePart("yyyy",        DateAdd("m", -11, date)           ), 2)

'DIALOG===========================================================================================================================
BeginDialog housing_grant_MONY_CHCK_issuance_dialog, 0, 0, 351, 90, "Housing grant Expansion MONY-CHCK"
  EditBox 65, 70, 55, 15, MAXIS_case_number
  EditBox 175, 70, 25, 15, initial_month
  EditBox 205, 70, 25, 15, initial_year
  ButtonGroup ButtonPressed
    OkButton 240, 70, 50, 15
    CancelButton 295, 70, 50, 15
  Text 15, 45, 320, 10, "Before you use the script, please review the case for eligibility for the MFIP housing grant."
  Text 130, 75, 40, 10, "month/year:"
  GroupBox 10, 5, 335, 55, "Housing grant Expansion:"
  Text 15, 75, 50, 10, "Case Number:"
  Text 15, 20, 325, 20, "This script should be used when the MFIP housing grant should have been issued on an eligible case for months prior to the current month or current month plus one. "
EndDialog

'The script============================================================================================================================
'Connects to MAXIS, grabbing the case MAXIS_case_number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
initial_month = CM_minus_1_mo	'defaulting to current month - 1 
initial_year = CM_minus_1_yr

'Main dialog: user will input case number and initial month/year will default to current month - 1 and member 01 as member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog housing_grant_MONY_CHCK_issuance_dialog				'main dialog
		If buttonpressed = 0 THEN stopscript	'script ends if cancel is selected
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid case number."		'mandatory field
		IF len(initial_month) <> 2 or isnumeric(initial_month) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit month."	'mandatory field
		IF len(initial_year) <> 2 or isnumeric(initial_year) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a valid 2 digit year."		'mandatory field
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Clears out case number and enters the selected footer month/year
back_to_self
EMWritescreen "________", 18, 43
EMWritescreen MAXIS_case_number, 18, 43
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

'Checking for PRIV cases.
EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case, script will end.
IF priv_check = "PRIVIL" THEN script_end_procedure("This case is a privliged case. You do not have access to this case.")

'checking to see if HG has been issued for the month selected: MONY/INQX----------------------------------------------------------------------------------------------------
DO
	row = 6				'establishing the row to start searching for issuance'
	DO
		EMReadScreen housing_grant, 2, row, 19		'searching for housing grant issuance
		If housing_grant = "  " then exit do		'reached the end of the issuance amounts
		IF housing_grant = "HG" then
			'reading the housing grant information
			EMReadScreen HG_amt_issued, 7, row, 40
			EMReadScreen HG_month, 2, row, 73
			EMReadScreen HG_year, 2, row, 79
			INQD_issuance = HG_month & HG_year			'creates a new variable for housing grant month & year
			month_of_issuance = initial_month & initial_year	'creates a new variable with footer month & footer year from dialog
			'if an issuance is found that matches the month/year selected by the user, the script will stop
			If month_of_issuance = INQD_issuance then script_end_procedure("Issuance has already been made on the month selected. Please review your case, and update manually.")
		END IF
		row = row + 1
	Loop until row = 18				'repeats until the end of the page
		PF8
		EMReadScreen last_page_check, 21, 24, 2
		If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

'navigates to STAT/MEMB to check for PARIS matches for all people on the case----------------------------------------------------------------------------------------------------
back_to_SELF
EMWritescreen initial_month, 20, 43			'enters footer month/year user selected since you have to be in the same footer month/year as the CHCK is being issued for
EMWritescreen initial_year, 20, 46

CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the SSN number to check PARIS.

DO								'reads the SSN number and puts it into a single string to convern into an array
	EMReadscreen SSN_number_read, 11, 7, 42
	SSN_number_read = replace(SSN_number_read, "_", "")  'replacing blank SSN underscores with nothing.
	client_array = client_array & replace(SSN_number_read, " ", "") & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = TRIM(client_array)					'converting into array'
HH_member_array = split(client_array, "|")

CALL Navigate_to_MAXIS_screen("INFC", "")

FOR each SSN_for_PARIS in HH_member_array						'for each person who we found on memb we enter their SSN into INFC INTM
	IF SSN_for_PARIS <> "" THEN												'if the SSN_for_PARIS spot on that array isn't blank we need to evaluate it
		EMWriteScreen SSN_for_PARIS, 3, 63
		EMWriteScreen "INTM", 20, 71
		transmit
		INTM_row = 8													'setting variable for the top of the current read row
		DO
			PARIS_edit_check = ""								'resetting varibles for each iteration of the loop'
			PARIS_resolution = ""
			EMReadScreen PARIS_month, 5, INTM_row, 59							'reading current row's paris month
			EMReadScreen PARIS_resolution, 2, INTM_row, 73				'reading current row's paris resolution code'
			IF TRIM(PARIS_resolution) = "" THEN Exit DO
			IF PARIS_resolution <> "RV" THEN											'If the match is marked as ANYTHING other than RV we need to evaluate it.
				paris_date = left(PARIS_month, 2) & "/01/" & right(PARIS_month, 2)		'putting dates into MM/DD/YY format for easier use
				initial_date = initial_month & "/01/" & initial_year									'putting dates into MM/DD/YY format for easier use
				IF TRIM(PARIS_month) = "" THEN																'if we find a blank row we will try to PF8
					PF8
					EMReadScreen PARIS_edit_check, 4, 24, 14										'checking if the PF causes the message THIS IS THE LAST PAGE as that will let us out of loop.
					INTM_row = 8																								'resetting row in case there are multiple pages of PARIS matches.
				ELSE																													'Otherwise if PARIS_month wasn't blank we compare it against the last 12 months from selected issuance date
					IF DateDiff("m", paris_date, initial_date) <= 12 THEN script_end_procedure("This SSN has a PARIS match from within the last 12 months. Please review and process manually.")
				END IF
			END IF
			INTM_row = INTM_row + 1																		'iterating the row so next loop through can check following row.
		LOOP until PARIS_edit_check = "LAST"												'exiting if we haven't script ended already and we reached the last page of results for this person
		PF3																													'PF3ing back to INTM
	END IF
NEXT																														'looping back with the for next to check the next SSN.

'navigates to ELIG/MFIP once the footer month and date are the selected dates: ELIG/MFIP----------------------------------------------------------------------------------------------------
Call navigate_to_MAXIS_screen("ELIG", "MFIP")
'Ensures that users is in the most recently approved version of MFIP
EMReadScreen no_MFIP, 10, 24, 2
If no_MFIP = "NO VERSION" then script_end_procedure("There are no eligibilty results for this case. Please check your case number/case for accuracy.")

'Signficant change cases do not automatically open to the MFPR panel. This ensures that we get there. 
Do 
	EMReadscreen MFPR_panel_check, 4, 3, 47
	If MFPR_panel_check <> "MFPR" then 
		EMWritescreen "MFPR", 20, 71
		transmit
	END IF 
LOOP until MFPR_panel_check = "MFPR"

'Script will check for Fraud on most recent unappproved version that may have been added after report was generated as you cannot approve negative actions in previous months
fraud_row = 7											'dummy variable to count what row we are on for do loop
DO
	EMReadScreen fraud_member_check, 1, fraud_row, 53 'Reading a spot on elig status to determine if there is even an entry for this row
	IF fraud_member_check <> " " THEN  								'if there is something on this row, check it
		EMWriteScreen "X", fraud_row, 3  'places x on member's person test
		Transmit													'transmits to open person test for that member
		EMReadScreen fraud_status, 1, 13, 17 'Reading one character as it will only either be FAILED or PASSED
		IF fraud_status = "F" THEN						'if Fraud is FAILED then we must quit script, otherwise we can check
			script_end_procedure("Fraud was found FAILED on MFIP person test. Please email this case to Rita Galindre at Rita.Galindre@state.mn.us ")
		ELSE
			Transmit													'transmitting out of person test
		END IF
	END IF
	fraud_row = fraud_row + 1   'incrementing to the next row
	IF fraud_row = 18 THEN
		PF8      'if we've reading the
		EMReadScreen fraud_edit_check, 2, 24, 5 'Reading for NO MORE MEMBERS TO DISPLAY edit message.
		fraud_row = 7 'resetting fraud row
	END IF
LOOP until fraud_edit_check = "NO"

'navigating back to intial MFIP elig now that we've checked for fraud.
Call navigate_to_MAXIS_screen("ELIG", "MFIP")
'Ensures that users is in the most recently approved version of MFIP
EMReadScreen no_MFIP, 10, 24, 2
If no_MFIP = "NO VERSION" then script_end_procedure("There are no eligibilty results for this case. Please check your case number/case for accuracy.")

'if case is signficatnt change, then user will need to transmit past to the MFPR
EMReadScreen sign_change, 6, 4, 15
If sign_change = "CHANGE" then 
	EMReadScreen app_version, 8, 3, 3
	IF app_version = "APPROVED" then 
		transmit
	Else 
		'If the most recent version is not approved, then the worker should be reviewing and processing this case manually
		script_end_procedure("Case has significant change, but version is not approved. Process manually.")
	END IF 
Else 
	EMWriteScreen "99", 20, 79 		'this is the most amount of eligibility results that elig can contain, so all versions appear in the next pop up
	transmit
    'This brings up the MFIP versions of eligibilty results to search for approved versions
    MAXIS_row = 7
    Do
    	EMReadScreen app_status, 8, MAXIS_row, 50
    	If trim(app_status) = "" then exit do 	'if end of the list is reached then exits the do loop
    	If app_status = "UNAPPROV" Then MAXIS_row = MAXIS_row + 1
    	If app_status = "APPROVED" then
    		EMReadScreen elig_status, 8, MAXIS_row, 37
    		If elig_status = "ELIGIBLE" then
    			EMReadScreen vers_number, 1, MAXIS_row, 23
    			EMWriteScreen vers_number, 18, 54
    			transmit		'transmits to approved and eligible veresion of MFIP
    			exit do
    		ELSE
    		 	MAXIS_row = MAXIS_row + 1
    		END IF
    	END IF
    Loop until app_status = "APPROVED" or trim(app_status) = ""
    'If no elig results are found, then the script ends.
    If trim(app_status) = "" then script_end_procedure("Eligible and approved MFIP results were not found. Please check your case for accuracy.")
End if 

'Signficant change cases do not automatically open to the MFPR panel. This ensures that we get there. 
Do 
	EMReadscreen MFPR_panel_check, 4, 3, 47
	If MFPR_panel_check <> "MFPR" then 
		EMWritescreen "MFPR", 20, 71
		transmit
	END IF 
LOOP until MFPR_panel_check = "MFPR"

EMWritescreen "x", 7, 3			'selects the member number to navigate to the MFIP Person Test Results
transmit
'The recipient isevaluated as meeting one of the 2 newly added population inelgible codes
'Checking FAILED reason for newly added population (SSI recipients and undocumented non-citizens with eligible children)
issuance_reason = ""	'issuance_reason = "" will determine what path the script takes. If "" then case is an emps exempt person, if not person is newly added population person
EMReadscreen cit_test_status, 6, 9, 17
EMReadscreen SSI_test_status, 6, 9, 52
If cit_test_status = "FAILED" then
	issuance_reason = "is an undocumented non-citizen with eligible children"
ElseIf SSI_test_status = "FAILED" then
	issuance_reason = "receives federal SSI due to disability that prevents work"
END IF

'If no EMPS exclusion exists, or one of the applicable tests are not failed, then case is not elig for HG supplement.
If issuance_reason = "" then script_end_procedure("Case does not meet criteria for a Housing Grant supplement. Please review the case for accuracy.")

transmit  'Transmits to exit the MFIP Person Test Results back to MFPR
Call navigate_to_MAXIS_screen("ELIG", "MFBF")
'ensures the user is indeed on MFBF otherwise the array will not be filled and the script will suffer from an epic fail
DO
	EMReadScreen MFBF_check, 4, 3, 47
	If MFBF_check <> "MFBF" then
		EMWriteScreen "MFBF", 20, 71
		transmit
	END IF
LOOP until MFBF_check = "MFBF"

'If case is signifcant change, then it does not enter the version number since the approved version is the current version. Otherwise, the version # needs to be selected.
If sign_change <> "CHANGE" then  
	EMWriteScreen vers_number, 20, 79 'enters the version number of the elig and approved version of the script once it's confirmed that we're back in MFBF
	transmit
END IF 

'establishes values for variables and declaring the arrays for newly added population cases
number_eligible_members = 0
entry_record = 0

DIM MFIP_member_array()
Redim MFIP_member_array(3, 0)

'constants for array
const member_code 		= 0
const adult_child_code	= 1
const cash_code 		= 2
const state_food_code 	= 3

'Gathers information for the array (member code, adult_child_code, cash_code, state_food_code)
MAXIS_row = 7	'establishing the row to start searching for members
DO
	add_to_array = ""
	EMReadscreen ref_num, 2, MAXIS_row, 3		'searching for member number
	If ref_num = "  " then exit do				'exits do if member number matches
	EMReadScreen member_elig_status, 1, MAXIS_row, 27
	'Adding members to array to gather information for the MONY/CHCK (member number, adult vs child, cash and state food coding)
	If ref_num = "01" then
		add_to_array = True						'MEMB 01 needs to be added to MONY/CHCK weather they are eligible or not
	Elseif trim(member_elig_status) = "A" then
		add_to_array = True						'all eligible HH members need to be added to MONY/CHCK
	Else
		add_to_array = False 					'Anyone who is not MEMB 01 or is INELIGIBLE is not added to the array
	End if

	If add_to_array = True then
		EMReadScreen cash, 	 1, MAXIS_row, 37		'reads cash and state_food coding
		EMReadScreen state_food,  1, MAXIS_row, 54

		ReDim Preserve MFIP_member_array(3,  entry_record)				'This resizes the array based on the number of members being added to the array
		MFIP_member_array (member_code,      entry_record) = ref_num	'The client member # is added to the array
		MFIP_member_array (cash_code,    	 entry_record) = cash		'inputs the cash code into the array
		MFIP_member_array (state_food_code,  entry_record) = state_food	'inputs the state food code into the array

		entry_record = entry_record + 1
		If trim(member_elig_status) = "A" then number_eligible_members = number_eligible_members + 1	'adds up the total number of eligible members to be inputted into MONY/CHCK
	END IF

	MAXIS_row = MAXIS_row + 1	'otherwise it searches again on the next row
	If MAXIS_row = 16 then
		PF8
	END IF
LOOP until trim(ref_num) = ""

'ensures that number_eligible_members is a two-digit number to be inputted into MONY/CHCK
number_eligible_members = "0" & number_eligible_members
number_eligible_members = right(number_eligible_members, 2)

'goes into CASE/PERS and grabs the adult_child_code to be inputted into the MONY/CHCK
Call navigate_to_MAXIS_screen("CASE", "PERS")
For item = 0 to Ubound(MFIP_member_array, 2)
	MAXIS_row = 10
	Do
		EMReadScreen pers_ref_number, 2, MAXIS_row, 3
		IF trim(pers_ref_number) = "" then exit do
		IF MFIP_member_array(member_code, item) = pers_ref_number then
			EMReadScreen relationship_status, 10, MAXIS_row + 1, 18 	'relationship_status is found one line down from the member number
			IF 	trim(relationship_status) = "Child" or _
				trim(relationship_status) = "Step Child" or _
				trim(relationship_status) = "Grandchild" or _
				trim(relationship_status) = "Niece" or _
				trim(relationship_status) = "Nephew" then
				relationship_status = "C"
			else
				relationship_status = "A"			'defaults all non-child relationships to adults
			END IF
			MFIP_member_array (adult_child_code, item) = relationship_status 'The client adult_child_code is added to the array
			exit do
		Else
			MAXIS_row = MAXIS_row + 3			'information is 3 rows apart
			If MAXIS_row = 19 then
				PF8
				MAXIS_row = 10					'changes MAXIS row if more than one page exists
			END if
		END if
		EMReadScreen last_PERS_page, 21, 24, 2
	LOOP until last_PERS_page = "THIS IS THE LAST PAGE"
Next

'MONY/CHCK----------------------------------------------------------------------------------------------------
'navigates to MONY/CHCK and inputs codes into 1st screen:
back_to_SELF
EMWritescreen initial_month, 20, 43			'enters footer month/year user selected since you have to be in the same footer month/year as the CHCK is being issued for
EMWritescreen initial_year, 20, 46

Call navigate_to_MAXIS_screen("MONY", "CHCK")
'error handling if a worker does not have access to a specific case (out of county, etc.)
EMReadscreen auth_error, 8, 24, 2
If auth_error = "YOUR ARE" then script_end_procedure("You are not authorized to issue a MONY/CHCK on this case. The script will now end.")

EMWriteScreen "MF", 5, 17		'enters mandatory codes per HG instruction
EMWriteScreen "MF", 5, 21		'enters mandatory codes per HG instruction
EMWriteScreen "31", 5, 32		'restored payment code per the HG instruction
EMWriteScreen "N", 8, 27		'enters N for migrant status for cases that are now inactive, and prog has been cleared. 

'total # eligible house hold members from MFBF needs to be inputted
EMWriteScreen number_eligible_members, 7, 27			'enters the number of eligible HH members
transmit

'Ensures that cases that have exceeded the issuance cannot continue. 
EMReadScreen issuance_exceeded, 5, 24, 2
IF issuance_exceeded = "TOTAL" then script_end_procedure("Total issuance exceeds monthly maximum of $1500 for this case. Contact your supervisor to approve issuance.")
	
EMReadScreen future_month_check, 6, 24, 2		'ensuring that issuances for current or future months are not being made
IF future_month_check = "REASON" then script_end_procedure("You cannot issue a MONY/CHCK for the current or future month. Approve results in ELIG/MFIP.")

'now we're on the MFIP issuance detail pop-up screen
MAXIS_row = 10
For item = 0 to UBound(MFIP_member_array, 2)
	'writing in each member's member, adult/child, cash and state food codes from ELIG/MFIP
	EMWriteScreen MFIP_member_array(member_code,		item), MAXIS_row, 6
	EMWriteScreen MFIP_member_array(adult_child_code, 	item), MAXIS_row, 14
	EMWriteScreen MFIP_member_array(cash_code, 			item), MAXIS_row, 23
	EMWriteScreen MFIP_member_array(state_food_code,	item), MAXIS_row, 33
	MAXIS_row = MAXIS_row + 1
	If MAXIS_row = 15 then
		PF8			'accounting for more than one page of members to input
		MAXIS_row = 10
	End if
NEXT

EMwritescreen "110.00", 10, 53			'enters the housing grant amount
transmit

EMReadScreen extra_error_check, 7, 17, 4			'double-checking that a duplicate issuance has not been made
IF extra_error_check = "HOUSING" then script_end_procedure ("Housing grant may have already been issued. Please recheck your case, and try again.")
EMReadscreen REI_issue, 3, 15, 6
If REI_issue = "REI" then
	EMWriteScreen "N", 15, 52	'N to REI issuance per instruction from DHS
Else
	transmit
END IF
Transmit
EMWriteScreen "Y", 15, 29	'Y to confirm approval
transmit
transmit 'transmits twice to get to the restoration of benefits screen

'some cases need to have the TIME panel completed
EMReadScreen update_TIME_panel_check, 4, 14, 32
If update_TIME_panel_check = "TIME" then
	transmit
	Do 
		PF10
		PF3
		EMReadScreen TIME_panel, 4, 2, 46
	LOOP until TIME_panel <> "TIME"
END IF
PF3
PF3 	'PF3's twice to NOT send the notice

'Ensuring that issuance made by checking the automated case note
back_to_SELF
Call navigate_to_MAXIS_screen("CASE", "NOTE")
EMWriteScreen "X", 5, 3
transmit		'Entering the 1st case note which is the case note made by the system automatically after issuance.
EMReadScreen payment_month, 2, 5, 28
EMReadScreen payment_year, 2, 5, 34
'Created new variables for confirming issuance
payment_date = payment_month & "/" & payment_year
issuance_month = initial_month & "/" & initial_year

If payment_date <> issuance_month then
	script_end_procedure("WARNING!" & vbNewLine & VbnewLine & " Issuance for " & issuance_month & " may not have occurred. Please check the case to ensure that issuance has been made.")
Else
	script_end_procedure("Success! A MONY/CHCK has been issued.")
END IF