'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - HG SUPPLEMENT.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 269                	'manual run time in seconds
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

'Date variables 
'current month -1
CM_minus_1_mo =  right("0" &          	 DatePart("m",           DateAdd("m", -1, date)            ), 2)
CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
'current month - 11
CM_minus_11_mo =  left("0" &            DatePart("m",           DateAdd("m", -11, date)           ), 2)
CM_minus_11_yr =  right(                 DatePart("yyyy",        DateAdd("m", -11, date)           ), 2)

'DIALOG===========================================================================================================================
BeginDialog housing_grant_MONY_CHCK_issuance_dialog, 0, 0, 311, 135, "Housing grant supplement"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  EditBox 165, 10, 25, 15, member_number
  EditBox 245, 10, 25, 15, initial_month
  EditBox 275, 10, 25, 15, initial_year
  EditBox 80, 110, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 110, 50, 15
    CancelButton 250, 110, 50, 15
  Text 15, 80, 280, 20, "Before you use the script, please review the case for eligibility for the MFIP housing grant."
  Text 200, 15, 40, 10, "month/year:"
  Text 15, 115, 60, 10, "Worker signature:"
  GroupBox 10, 35, 290, 70, "Housing grant supplement:"
  Text 125, 15, 35, 10, "Member #:"
  Text 10, 15, 50, 10, "Case Number:"
  Text 15, 55, 280, 20, "This script should be used when the MFIP housing grant should have been issued on an eligible case for months prior to the current month or current month plus one. "
EndDialog

'The script============================================================================================================================
'Connects to MAXIS, grabbing the case MAXIS_case_number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number) 
member_number = "01"	'defaults the member number to 01
initial_month = CM_minus_1_mo  'defaulting date to current month - one
initial_year = CM_minus_1_yr

'Main dialog: user will input case number and initial month/year will default to current month - 1 and member 01 as member number
DO
	DO
		err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
		dialog housing_grant_MONY_CHCK_issuance_dialog				'main dialog
		If buttonpressed = 0 THEN stopscript	'script ends if cancel is selected
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid case number."		'mandatory field
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

'navigates to ELIG/MFIP once the footer month and date are the selected dates: ELIG/MFIP----------------------------------------------------------------------------------------------------
back_to_SELF
EMWritescreen initial_month, 20, 43			'enters footer month/year user selected since you have to be in the same footer month/year as the CHCK is being issued for
EMWritescreen initial_year, 20, 46
Call navigate_to_MAXIS_screen("ELIG", "MFIP")	

'Ensures that users is in the most recently approved version of MFIP
EMReadScreen no_MFIP, 10, 24, 2
If no_MFIP = "NO VERSION" then script_end_procedure("There are no eligibilty results for this case. Please check your case number/case for accuracy.")
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

msgbox "Are we on the most up-to-date approved version?"

'goes into ELIG/MFIP, and checks for the reason for the manual MONY/CHCK: either emps exempt populations or the newly added populations 
MAXIS_row = 7	'establishing the row to start searching
DO 
	EMReadscreen memb_number, 2, MAXIS_row, 6		'searching for member number from initial dialog
	If memb_number = "  " then script_end_procedure("The member number you entered does not appear to be valid. Please check your member number and try again.")
	IF member_number = memb_number then exit do				'exits do if member number matches
	MAXIS_row = MAXIS_row + 1	'otherwise it searches again on the next row 	
LOOP until MAXIS_row = 18

'If the member number is found, script reads the EMPS coding to case note and fill out the MONY/CHCK verbiage
EMWritescreen "x", MAXIS_row, 64			'selects the member number at EMPS indicator
transmit
EMReadscreen emps_status_error, 19, 24, 2
'If there is an EMPS coding, then the emps coding and the cash_portion code and the state_portion code are gathered
If trim(emps_status_error) = "" then 
	EMReadscreen emps_status, 2, 9, 22			'grabs the EMPS status code'
	transmit
	'grabs the coding to input in MONY/CHCK
	Call navigate_to_MAXIS_screen("ELIG", "MFBF")
	EMReadscreen cash_portion, 1, MAXIS_row, 37
	EMReadScreen state_portion, 1, MAXIS_row, 54
END IF 

'If the error code exist it means that there is no EMPS code, and the recipient needs to be evaluated as meeting one of the 2 newly added population inelgible codes
If emps_status_error = "EMPS DOES NOT EXIST" then 
	msgbox "EMPS does not exist"
	EMWritescreen "_", MAXIS_row, 64
	EMWritescreen "x", MAXIS_row, 3			'selects the member number to navigate to the MFIP Person Test Results
	transmit
	'Checking FAILED reason for newly added population (SSI recipients and undocumented non-citizens with eligible children) 
	issuance_reason = ""	'issuance_reason = "" will determine what path the script takes. If "" then case is an emps exempt person, if not person is newly added population person 
	EMReadscreen cit_test_status, 6, 9, 17
	EMReadscreen SSI_test_status, 6, 9, 52
	If cit_test_status = "FAILED" then 
		issuance_reason = "is an undocumented non-citizen with eligible children"	
	ElseIf SSI_test_status = "FAILED" then 
		issuance_reason = "receives federal SSI due to disability that prevents work" 		 
	END IF
	
	msgbox "issuance reason" & issuance_reason
	
	'If no EMPS exclusion exists, or one of the applicable tests are not failed, then case is not elig for HG supplement.
	If issuance_reason = "" then script_end_procedure("Case does not meet criteria for a Housing Grant supplement. Please review the case for accuracy.")
	
	'establishes values for variables and declaring the arrays for newly added population cases
	number_eligible_members = 0
	entry_record = 0
	
	DIM MFIP_member_array()
	Redim MFIP_member_array(3, 0)
	
	'constants for array
	const member_code 		= 0
	const adult_child_code		= 1
	const cash_code 		= 2
	const state_food_code 	= 3 
	
	transmit  'Transmits to exit the MFIP Person Test Results back to MFPR
	MAXIS_row = 7	'establishing the row to start searching at MEMB 01
	DO 
		add_to_array = ""
		EMReadscreen ref_num, 2, MAXIS_row, 6		'searching for member number
		If ref_num = "  " then exit do				'exits do if member number matches
		EMReadScreen member_elig_status, 10, MAXIS_row, 53
		'Adding members to array to gather information for the MONY/CHCK (member number, adult vs child, cash and state food coding)
		If ref_num = "01" then 
			add_to_array = True						'MEMB 01 needs to be added to MONY/CHCK weather they are eligible or not
		Elseif trim(member_elig_status) = "ELIGIBLE" then 
			add_to_array = True						'all eligible HH members need to be added to MONY/CHCK
		Else
			add_to_array = False 					'Anyone who is not MEMB 01 or is INELIGIBLE is not added to the array 
		End if 
		msgbox ref_num & vbcr & member_elig_status & vbcr & add_to_array
		If add_to_array = True then 	
			ReDim Preserve MFIP_member_array(3,  entry_record)	'This resizes the array based on the number of members being added to the array
			MFIP_member_array (member_code,      entry_record) = ref_num			'The client member # is added to the array
			entry_record = entry_record + 1
			If trim(member_elig_status) = "ELIGIBLE" then number_eligible_members = number_eligible_members + 1	'adds up the total number of eligible members to be inputted into MONY/CHCK
		END IF 	
		MAXIS_row = MAXIS_row + 1	'otherwise it searches again on the next row 	
		If MAXIS_row = 19 then 
			PF8
			MAXIS_row = 7
		END IF 
	LOOP until MAXIS_row = 19		
	
	'ensures that number_eligible_members is a two-digit number to be inputted into MONY/CHCK
	number_eligible_members = "0" & number_eligible_members
	number_eligible_members = right(number_eligible_members, 2)
	msgbox "# of eligible members: " & number_eligible_members
	
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
		msgbox pers_ref_number & " " & relationship_status
	Next
	'Cannot navigate directly to ELIG/MFBF, so needs to go back to ELIG/MFIP 1st
	Call navigate_to_MAXIS_screen("ELIG", "MFIP")
END IF 	

Call navigate_to_MAXIS_screen("ELIG", "MFBF")
If issuance_reason = "" then 
	'checking for sanctions, user will have to process manually if there's a sanction
	EMReadScreen MFIP_sanction, 1, MAXIS_row, 68	'checking for SANCTION for selected HH member
	If MFIP_sanction = "Y" then	script_end_procedure("A sanction exist for this member. Please check sanction for accuracy, and process manually.")
	'checking for elig for the $110 housing grant if exmeption is based on eligible emps codes
	Call navigate_to_MAXIS_screen("ELIG", "MFSM")
	EMReadScreen housing_grant_issued, 6, 16, 75
	IF housing_grant_issued <> "110.00" then script_end_procedure("This case does not have the housing grant issued in the eligibility results. Please review the case for eligibility. You may need to run this case through background. You will need to populate housing grant results prior to issuing the MONY/CHCK.")
Else
	'get the hh member information for member 01 and all eligible HH member_elig_status
	For item = 0 to UBound(MFIP_member_array, 2)
		MAXIS_row = 7
		DO 
			EMReadScreen reference_number, 2, MAXIS_row, 3
			IF trim(reference_number) = "" then exit do
			IF MFIP_member_array(member_code, item) = reference_number then 	
				EMReadScreen cash, 	 1, MAXIS_row, 37		'reads cash and state_food coding
				EMReadScreen state_food,  1, MAXIS_row, 54
				msgbox reference_number & vbcr & cash & vbcr & state_food
				MFIP_member_array (cash_code,    	item) = cash		'inputs the cash and state_food codes into the array for each member
				MFIP_member_array (state_food_code, item) = state_food
				exit do
			Else 
 				MAXIS_row = MAXIS_row + 1
				If MAXIS_row = 16 then 
					PF8					'changes MAXIS row if more than one page exists
					MAXIS_row = 7
				END IF
			END IF 
		LOOP until trim(reference_number) = ""
	NEXT
END IF 

'navigates to MONY/CHCK and inputs codes into 1st screen: MONY/CHCK----------------------------------------------------------------------------------------------------
back_to_SELF
EMWritescreen initial_month, 20, 43			'enters footer month/year user selected since you have to be in the same footer month/year as the CHCK is being issued for
EMWritescreen initial_year, 20, 46

Call navigate_to_MAXIS_screen("MONY", "CHCK")
'error handling if a worker does not have access to a specific case.
EMReadscreen auth_error, 8, 24, 2
If auth_error = "YOUR ARE" then script_end_procedure("You are not authoriszed to issue a MONY/CHCK on this case. The script will now end.")
 
EMWriteScreen "MF", 5, 17		'enters mandatory codes per HG instruction
EMWriteScreen "MF", 5, 21		'enters mandatory codes per HG instruction
EMWriteScreen "31", 5, 32		'restored payment code per the HG bulletin
'If newly added population eligble, then total # eligible house hold members needs to be inputted
If issuance_reason <> "" then 
	EMWriteScreen number_eligible_members, 7, 27			'enters the number of eligible HH members
	msgbox "# of eligible members: " & number_eligible_members
Else 
	EMWriteScreen member_number, 7, 27
End if
transmit 
EMReadScreen future_month_check, 7, 24, 2		'ensuring that issuances for current or future months are not being made
IF future_month_check = "PERIOD" then script_end_procedure("You cannot issue a MONY/CHCK for the current or future month. Approve results in ELIG/MFIP.")	

'now we're on the MFIP issuance detail pop-up screen
If issuance_reason <> "" then 
	'enter everyone's member, cash and state_food codes************************************************************************************************************************************************************************************************
	MAXIS_row = 10
	For item = 0 to UBound(MFIP_member_array, 2)
		'writing in each member's member, adult/child, cash and state food codes from ELIG
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
Else 
	'writing in coding for emps exempt population
	EMWriteScreen "01", 10, 6
	EMWriteScreen "A", 10, 14			'adds coding from MFBF into issuance detail screen
	EMWriteScreen cash_portion, 10, 23 
	EMWriteScreen state_portion, 10, 33
END IF 
EMwritescreen "110.00", 10, 53			'enters the housing grant amount

'This is here temporarily until testing is completed. Testers will need to PF3 to exit the MONY/CHCK function, then MONY/CHCK's will not be sent to recipients.  
msgbox "All eligible members and MEMB 01 added and Hg issuance ready. Stop script will occur once message box is closed."
stopscript

transmit
EMReadScreen extra_error_check, 7, 17, 4			'double-checking that a duplicate issuance has not been made
IF extra_error_check = "HOUSING" then script_end_procedure ("Housing grant may have already been issued. Please recheck your case, and try again.")

EMWriteScreen "N", 15, 52	'N to REI issuance per instruction from DHS
transmit
EMWriteScreen "Y", 15, 29	'Y to confirm approval
transmit
transmit 'transmits twice to get to the restoration of benefits screen

'some cases need to have the TIME panel completed
EMReadScreen update_TIME_panel_check, 4, 14, 32
If update_TIME_panel_check = "TIME" then 
	transmit
	time_panel_confirmation = MsgBox("You must update the time panel for " & initial_month & "/" & initial_year & ". Please update the TIME panel, or PF10 if it does not need to be updated, and press OK when complete.", vbOk, "Update the TIME panel")
	DO
		EMReadScreen time_panel_complete_check, 7, 24, 2 
	LOOP until time_panel_check <> "NO DATA"
	If time_panel_confirmation = vbOK then PF3
END IF 

'writes in the manual check reason per the bulletin on the Housing Grant for emps_exmption reason cases only!
If issuance_reason = "" then 
	EMWriteScreen "You meet one of the exceptions", 13, 18
	EMWriteScreen "listed in CM 13.03.09 for families", 14, 18
	EMWriteScreen "with an adult MFIP unit member(s)", 15, 18
	If emps_status = "02" or emps_status = "07" or emps_status = "12" or emps_status = "23" or emps_status = "27" or emps_status = "15" or emps_status = "18" or emps_status = "30" or emps_status = "33" then
   		EMWriteScreen "who get Section 8/HUD funded subsidy:", 16, 18
		EMWriteScreen "Caregivers who are elderly/disabled", 17, 18		'writes in disa/elderly if the codes above are the client's emps_status code
	Else 
		EMWriteScreen "who get Section 8/HUD funded subsidy:", 16, 18
		EMWriteScreen "Caregivers caring for a disabled member", 17, 18
	END IF 
	PF4  'sends the restoration letter

    'updating emps_status coding for case note
    If emps_status = "02" then 
    	emps_status = "Age 60 or older"
    Elseif emps_status = "08" or emps_status = "24" then 
    	emps_status = "Care for Ill/incapacitated family member"
    Elseif emps_status = "07" or emps_status = "23" then 
    	emps_status = "Ill/incapacitated > 30 days" 
    ElseIf emps_status = "12" or emps_status = "27" then 
    	emps_status = "Special medical criteria"
    ElseIf emps_status = "15" or emps_status = "30" then 
    	emps_status = "Mentally Ill"
    ElseIf emps_status = "18" or emps_status = "33" then 
    	emps_status = "SSI/RSDI pending"
    Else 
    	emps_status = "Other reason"
    END IF
    
    'Case noting the MONY/CHCK info
    Call start_a_blank_case_note
    Call write_variable_in_case_note("**MONY/CHCK ISSUED FOR HOUSING GRANT for " & initial_month & "/" & initial_year& "**")
    if emps_status = "Other reason" then 
    	Call write_variable_in_case_note("* Member " & member_number & " meets criteria to receive the housing grant.")
    Else
    	Call write_variable_in_case_note("* Housing grant issued due to family meeting an exemption per CM.13.03.09.")
    	Call write_variable_in_case_note("* Member " & member_number & " exemption is: " & emps_status & ".")
    END IF 
    Call write_variable_in_case_note("--")
    Call write_variable_in_case_note(worker_signature)
END IF 

script_end_procedure("Success! A MONY/CHCK has been issued. Please review the case to ensure that all housing grant issuances have been made.")