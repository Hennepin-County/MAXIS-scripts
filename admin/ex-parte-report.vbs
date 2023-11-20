'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - EX PARTE REPORT.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	100			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

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
call changelog_update("05/10/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================

function find_unea_information()
'This function is to find UNEA information from VA, UC, and RR benefits and add them to arrays for these specific income types
	Call navigate_to_MAXIS_screen("STAT", "UNEA")									'navigate to STAT/UNEA
	For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)								'Loop through each member listed on the case
		EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76		'Navigate to the MEMB on UNEA
		EMWriteScreen "01", 20, 79													'Make sure we are at the first instance of UNEA for this member
		transmit
		MEMBER_INFO_ARRAY(unea_VA_exists, each_memb) = False						'Defaulting the income types
		MEMBER_INFO_ARRAY(unea_UC_exists, each_memb) = False
		MEMBER_INFO_ARRAY(unea_RR_exists, each_memb) = False

		EMReadScreen unea_vers, 1, 2, 78											'reading to see if there is at least 1 UNEA panel for this member
		If unea_vers <> "0" Then
			Do
				EMReadScreen claim_num, 15, 6, 37									'reading the panel information for the claim number and income type
				EMReadScreen income_type_code, 2, 5, 37
				If income_type_code = "01" or income_type_code = "02" Then			'These are RSDI types
					If left(start_of_claim, 9) <> MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) Then	'saving if there is a claim number that is NOT from the SSN of the member
						MEMBER_INFO_ARRAY(unmatched_claim_numb, each_memb) = claim_num
					End If
				End if
				claim_num = replace(claim_num, "_", "")								'removing extra underscores from the claim number

				'These are all income types associated with VA income.
				If income_type_code = "11" or income_type_code = "12" or income_type_code = "13" or income_type_code = "38" Then
					MEMBER_INFO_ARRAY(unea_VA_exists, each_memb) = True				'setting the array for the member to indicate that there is VA income
					ReDim Preserve VA_INCOME_ARRAY(va_last_const, va_count)			'sizing up the array for all VA income

					'Saving all the information about VA income for this member into the VA income array
					VA_INCOME_ARRAY(va_case_numb_const, va_count) = MAXIS_case_number
					VA_INCOME_ARRAY(va_ref_numb_const, va_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					VA_INCOME_ARRAY(va_pers_name_const, va_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
					VA_INCOME_ARRAY(va_pers_ssn_const, va_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
					VA_INCOME_ARRAY(va_pers_pmi_const, va_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
					VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = income_type_code
					If income_type_code = "11" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Disability"
					If income_type_code = "12" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Pension"
					If income_type_code = "13" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Other"
					If income_type_code = "38" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Aid & Attendance"
					VA_INCOME_ARRAY(va_claim_numb_const, va_count) = claim_num
					EMReadScreen VA_INCOME_ARRAY(va_prosp_inc_const, va_count), 8, 18, 68
					VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = trim(VA_INCOME_ARRAY(va_prosp_inc_const, va_count))
					If VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = "" Then VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = "0.00"

					va_count = va_count + 1			'incrementing the number of va incomes
				End If

				'This is UC income
				If income_type_code = "14" Then
					MEMBER_INFO_ARRAY(unea_UC_exists, each_memb) = True				'setting the array for the member to indicate that there is UC income
					ReDim Preserve UC_INCOME_ARRAY(uc_last_const, uc_count)			'sizing up the array for all UC income

					'Saving all the information about UC income for this member into the UC income array
					UC_INCOME_ARRAY(uc_case_numb_const, uc_count) = MAXIS_case_number
					UC_INCOME_ARRAY(uc_ref_numb_const, uc_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					UC_INCOME_ARRAY(uc_pers_name_const, uc_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
					UC_INCOME_ARRAY(uc_pers_ssn_const, uc_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
					UC_INCOME_ARRAY(uc_pers_pmi_const, uc_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
					UC_INCOME_ARRAY(uc_inc_type_code_const, uc_count) = income_type_code
					UC_INCOME_ARRAY(uc_inc_type_info_const, uc_count) = "Unemployment"
					UC_INCOME_ARRAY(uc_claim_numb_const, uc_count) = claim_num
					EMReadScreen UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count), 8, 13, 68
					UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = trim(UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count))
					If UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = "________" Then UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = "0.00"

					uc_count = uc_count + 1			'incrementing the number of uc incomes
				End If

				'This is railroad income
				If income_type_code = "16" Then
					MEMBER_INFO_ARRAY(unea_RR_exists, each_memb) = True				'setting the array for the member to indicate that there is RR income
					ReDim Preserve RR_INCOME_ARRAY(rr_last_const, rr_count)			'sizing up the array for all RR income

					'Saving all the information about the RR income for this member into the RR income array
					RR_INCOME_ARRAY(rr_case_numb_const, rr_count) = MAXIS_case_number
					RR_INCOME_ARRAY(rr_ref_numb_const, rr_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					RR_INCOME_ARRAY(rr_pers_name_const, rr_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
					RR_INCOME_ARRAY(rr_pers_ssn_const, rr_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
					RR_INCOME_ARRAY(rr_pers_pmi_const, rr_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
					RR_INCOME_ARRAY(rr_inc_type_code_const, rr_count) = income_type_code
					RR_INCOME_ARRAY(rr_inc_type_info_const, rr_count) = "Railroad Retirement"
					RR_INCOME_ARRAY(rr_claim_numb_const, rr_count) = claim_num
					EMReadScreen RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count), 8, 13, 68
					RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count) = trim(RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count))
					If RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count) = "________" Then RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count) = "0.00"

					rr_count = rr_count + 1			'incrementing the number of rr incomes
				End If

				transmit								'this moves to the next UNEA panel for this member
				EMReadScreen next_unea_nav, 7, 24, 2	'checking to see if we are at the last UNEA panel for this member
			Loop until next_unea_nav = "ENTER A"
		End If
	Next
end function

function find_UNEA_panel(MEMB_reference_number, UNEA_type_code, UNEA_instance, UNEA_claim_number, panel_found)
'function used to find the correct UNEA panel that matches member, UNEA type and specific claim number.
'This is currently only tested and vetted on SSA income panels, particularly Type 01, 02, 03
	ReDim unea_panel_array(last_panel_const, 0)					'Reset the array that will gather the information of all the unea panels for the given member
	unea_panel_counter = 0										'this array will need to iterate, this will set the counter to 0
	array_index = ""											'blanking out the variable that will capture the array index that the correct panel is at

	panel_found = False											'defaulting the boolean variables
	type_code_found = False
	UNEA_claim_number = replace(UNEA_claim_number, " ", "")		'formatting the claim number
	SSN_UNEA_claim_portion = left(UNEA_claim_number, 9)			'creating an array of the SSN portion of the claim number

	'here we navigate to the UNEA panel
	EMWriteScreen "UNEA", 20, 71
	transmit
	EMReadScreen unea_check, 4, 2, 48
	Do While unea_check <> "UNEA"								'making sure we've made it.
		Call navigate_to_MAXIS_screen("STAT", "UNEA")
		EMReadScreen unea_check, 4, 2, 48
	Loop
	EMWriteScreen MEMB_reference_number, 20, 76 				'Navigating to the right member of unea
	EMWriteScreen "01", 20, 79 									'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	transmit

	EMReadScreen vers_count, 1, 2, 78							'reading the number of versions of UNEA for this member
	If vers_count <> "0" Then									'if there are none, the rest of the functionality will be skipped, if there are any, we have to read them
		Do
			ReDim Preserve unea_panel_array(last_panel_const, unea_panel_counter)		'resizing the array of panels
			EMReadScreen panel_instance, 1, 2, 73										'reading the specific panel information
			EMReadScreen panel_type_code, 2, 5, 37
			EMReadScreen panel_claim_number, 15, 6, 37
			panel_claim_number = replace(panel_claim_number, "_", "")					'formatting the claim number
			panel_claim_number = replace(panel_claim_number, " ", "")
			If panel_type_code = UNEA_type_code Then type_code_found = True				'if the type code on the panel matches the type code as a parameter, we identify that we've found a panel

			unea_panel_array(panel_type_const, unea_panel_counter) = panel_type_code	'adding the information we have read to the array
			unea_panel_array(panel_claim_const, unea_panel_counter) = panel_claim_number
			unea_panel_array(panel_claim_left_9_const, unea_panel_counter) = left(panel_claim_number, 9)
			unea_panel_array(panel_instance_const, unea_panel_counter) = "0" & panel_instance

			unea_panel_counter = unea_panel_counter + 1			'incrementing the array counter

			transmit											'go to the next UNEA panel
			EMReadScreen end_of_UNEA_panels, 7, 24, 2			'read the warning/error message to see if we could not move to the next panel
		Loop Until end_of_UNEA_panels = "ENTER A"
	End If

	'THIS IS TESTING FUNCTIONALITY TO CHECK THE FUNCTION
	' For known_panel = 0 to UBound(unea_panel_array, 2)
	' 	claim_match = False
	' 	If unea_panel_array(panel_claim_const, known_panel) = UNEA_claim_number Then claim_match = True
	' 	ssn_match = False
	' 	If unea_panel_array(panel_claim_const, known_panel) = UNEA_claim_number Then ssn_match = True
	' 	MsgBox 	"unea_panel_array(panel_type_const, known_panel) - " & unea_panel_array(panel_type_const, known_panel) & vbCr &_
	' 	 		"unea_panel_array(panel_claim_const, known_panel) - " & unea_panel_array(panel_claim_const, known_panel) & vbCr &_
	' 	 		"unea_panel_array(panel_claim_left_9_const, known_panel) - " & unea_panel_array(panel_claim_left_9_const, known_panel) & vbCr &_
	' 	 		"unea_panel_array(panel_instance_const, known_panel) - " & unea_panel_array(panel_instance_const, known_panel) & vbCr & vbCr &_
	' 			"panel_type_code - " & panel_type_code & vbCr &_
	' 			"UNEA_type_code - " & UNEA_type_code & vbCr &_
	' 			"UNEA_claim_number - " & UNEA_claim_number & vbCr & vbCr &_
	' 			"claim numb match? " & claim_match & vbCr &_
	' 			"claim ssn numb match? " & ssn_match
	' Next

	'This part of the function will only trigger if the reading of the panels found a panel type that matches the parameter passed through
	If type_code_found = True Then
		For known_panel = 0 to UBound(unea_panel_array, 2)				'Loop through all the known panels
			'Now we need to find if an SSI panel is found or and RSDI panel is found.
			'RSDI can allow for 01 to match with 02 - this way the UNEA type can be corrected.
			'We were finding duplicates if the existing panel was the wrong UNEA type - we will still be trying to match on the claim number
			panel_type_matches = False																									'using a variable to identify a 'qualified match'
			If unea_panel_array(panel_type_const, known_panel) = UNEA_type_code Then panel_type_matches = True							'if there is an exact match, then it is matching
			If unea_panel_array(panel_type_const, known_panel) = "02" and UNEA_type_code = "01" Then panel_type_matches = True			'either RSDI type is a match if the panel is an RSDI type
			If unea_panel_array(panel_type_const, known_panel) = "01" and UNEA_type_code = "02" Then panel_type_matches = True

			If panel_type_matches = True Then
				If UNEA_type_code = "03" Then 			'If the call is for an SSI panel type, we do not need to match the claim number, just the panel type since SSI is only on the persons SSN
					panel_found = True					'setting the parameter to true
					array_index = known_panel			'defining which instance of the array matched the provided criteria
				End If
				If unea_panel_array(panel_claim_const, known_panel) = UNEA_claim_number Then		'If the claim number on the panel matches the claim number passed through, WE FOUND A MATCH
					panel_found = True					'setting the parameter to true
					array_index = known_panel			'defining which instance of the array matched the provided criteria
				End If
			End If
		Next
		'This functionality will only run if the correct panel was not found in the above loop
		If panel_found = False Then
			For known_panel = 0 to UBound(unea_panel_array, 2)			'Loop through all the known panels
				'Now we need to find if an SSI panel is found or and RSDI panel is found.
				'RSDI can allow for 01 to match with 02 - this way the UNEA type can be corrected.
				'We were finding duplicates if the existing panel was the wrong UNEA type - we will still be trying to match on the claim number
				panel_type_matches = False																									'using a variable to identify a 'qualified match'
				If unea_panel_array(panel_type_const, known_panel) = UNEA_type_code Then panel_type_matches = True							'if there is an exact match, then it is matching
				If unea_panel_array(panel_type_const, known_panel) = "02" and UNEA_type_code = "01" Then panel_type_matches = True			'either RSDI type is a match if the panel is an RSDI type
				If unea_panel_array(panel_type_const, known_panel) = "01" and UNEA_type_code = "02" Then panel_type_matches = True

				If panel_type_matches = True Then
					If unea_panel_array(panel_claim_left_9_const, known_panel) = SSN_UNEA_claim_portion Then	'If the SSN portion of the panel claim number matches the SSN portion of the claim passed through the function parameter
						panel_found = True				'setting the parameter to true
						array_index = known_panel		'defining which instance of the array matched the provided criteria
					End If
				End If
			Next
		End If

		'If the panel was found (either by exact match OR by SSN match)
		If panel_found = True Then
			EMWriteScreen MEMB_reference_number, 20, 76 									'Navigating to the right member
			EMWriteScreen unea_panel_array(panel_instance_const, array_index), 20, 79 		'enter the instance number from the panel
			transmit																		'navigating to the panel information entered

			UNEA_instance = unea_panel_array(panel_instance_const, array_index)				'passing the correct instance through the parameter
		End If
	End If
end function

function get_list_of_members()
'this function will get all of the members on the case into the member array. This will NOT duplicate members already read from the data table.
	client_count = UBound(MEMBER_INFO_ARRAY, 2) + 1											'setting the incrementer to the next instance of the member array
	If MEMBER_INFO_ARRAY(memb_pmi_numb_const, 0) = "" Then client_count = 0					'setting back to 0 if the first instance of the array is blank
	EMWriteScreen "01", 20, 76																'make sure to start at Memb 01
	transmit
	loop_count = 0

	'reads the reference number, last name, first name, and then puts it into a single string then into the array
	DO
		EMReadScreen client_PMI, 8, 4, 46													'reading the PMI from the panel we are currently on
		client_PMI = trim(client_PMI)														'formatting and adding the leading 0s to the PMI because that is how it is recorded in the data table
		client_PMI = RIGHT("00000000" & client_PMI, 8)

		client_found = False																'looking to see if this member is already listed in the member array (from the data table elig list)
		For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
			If client_PMI = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then		'if there is a match, we read some specific information
				client_found = True															'setting that the client was found to a boolean for this loop
				EMReadScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs), 2, 4, 33	'reading the members reference number
				EMReadScreen clt_age, 3, 8, 76												'reading the client age and setting it into the array
				MEMBER_INFO_ARRAY(memb_age_const, known_membs) = trim(clt_age)
				EMReadScreen MEMBER_INFO_ARRAY(memb_smi_numb_const, known_membs), 9, 5, 46	'reading the SMI for the member
				Exit For		'if we found the client, we can leave the loop
			End If
		Next

		If client_found = False Then														'if the member is not already in the member array, we will need to resize the array and add the person
			ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, client_count)					'resize the array

			EMReadScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, client_count), 2, 4, 33		'gather the member information and add it to the array
			MEMBER_INFO_ARRAY(memb_pmi_numb_const, client_count) = client_PMI
			EMReadScreen SSN1, 3, 7, 42
			EMReadScreen SSN2, 2, 7, 46
			EMReadScreen SSN3, 4, 7, 49
			MEMBER_INFO_ARRAY(memb_ssn_const, client_count) = SSN1 & SSN2 & SSN3
			EMReadScreen clt_age, 3, 8, 76
			MEMBER_INFO_ARRAY(memb_age_const, client_count) = trim(clt_age)
			EMReadScreen last_name, 25, 6, 30
			EMReadScreen first_name, 12, 6, 63
			last_name = trim(replace(last_name, "_", ""))
			first_name = trim(replace(first_name, "_", ""))
			MEMBER_INFO_ARRAY(memb_name_const, client_count) = last_name & ", " & first_name
			EMReadScreen MEMBER_INFO_ARRAY(memb_smi_numb_const, client_count), 9, 5, 46
			MEMBER_INFO_ARRAY(memb_active_hc_const, client_count)	= False					'default the active hc boolean in the array as false

			client_count = client_count + 1			'increment up the counter for the member array
		End If
		transmit								'go to the next member
		EMReadScreen edit_check, 7, 24, 2
		loop_count = loop_count + 1
		If loop_count > 40 Then Exit Do
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
end function

function send_sves_qury(ssn_or_claim, qury_finish)
'this function will send a SVES QURY for a person
	qury_finish = ""																	'blanking out the return parameter to output the result of the qury attempt
	Call navigate_to_MAXIS_screen("INFC", "SVES")										'navigate to INFC/SVES
	EMWriteScreen MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 68					'enter the member's SSN
	EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68				'enter the member's PMI number
	Call write_value_and_transmit("QURY", 20, 70)										'Now we will enter the QURY screen to type the case number.

	If ssn_or_claim = "CLAIM" Then														'If it is indicated that this is a QURY based on claim number, the qury entry need to be adjusted
		Call clear_line_of_text(5, 38)													'removing the SSN for the qury
		EMWriteScreen MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 7, 38			'enter the secondary claim number into the qury details
	End If
	EMWriteScreen MAXIS_case_number, 	11, 38											'enter the case number - QUESTION - if we do NOT enter this, would the DAIL not happen?
	EMWriteScreen "Y", 					14, 38											'confirm entry of this qury
	transmit  																			'Now it sends the SVES.

	EMReadScreen duplicate_SVES, 	    7, 24, 2										'reading to see if there is a duplicate SVES warning - then transmit by it
	If duplicate_SVES = "WARNING" then transmit
	EMReadScreen confirm_SVES, 			6, 24, 2										'reading the message from submitting the QURY to see if it was successful
	if confirm_SVES = "RECORD" then
		qury_finish = date							'output the date if the record was submitted
	Else
		qury_finish = "FAILED"						'output the failure if the record was not submitted
	END IF
end function

function update_stat_budg()
'function purpose is to update STAT/BUDG to align with the Ex Parte Renewal month
'This is a necessary step because Ex Parte process in MAXIS will only correctly message IF the ELIG Budget starts with the Ex parte month
	Call navigate_to_MAXIS_screen("STAT", "BUDG")		'get to STAT/BUDG
	EMReadScreen budg_begin_mo, 2, 10, 35				'Read the current budget month and year
	EMReadScreen budg_begin_yr, 2, 10, 38
	EMReadScreen budg_end_mo, 2, 10, 46
	EMReadScreen budg_end_yr, 2, 10, 49

	'this version of the functionality only gives one attempt
	'BUDG is fussy and won't allow for any gaps in buget months, and also doesn't allow for multiple changes very easily.
	If budg_begin_mo <> ep_revw_mo Then					'If it does not already match we are going to try to change it
		PF9 											'put the panel in update mode.
		EMWriteScreen ep_revw_mo, 5, 64					'Entering the
		EMWriteScreen ep_revw_yr, 5, 67
		EMWriteScreen ep_end_budg_revw_mo, 5, 72		'We have to enter in the last budget month and year as well
		EMWriteScreen ep_end_budg_revw_yr, 5, 75
		transmit										'save the updates

		EMReadScreen edit_message, 56, 24, 2			'check to see if there is an edit in the panel
		edit_message = trim(edit_message)				'trim the edit message

		If edit_message <> "" Then						'If there is an edit message - we will save the case and undo the changes
			objTextStream.WriteLine "Case: " & MAXIS_case_number & " - BUDG not updated"
			PF10
		End If
	End If
end function


function update_unea_pane(panel_found, unea_type, income_amount, claim_number, start_date, end_date, last_pay)
'this function will upate the information in the UNEA panel with new information.
'This function requires the function 'find_UNEA_panel' before calling this one because it identifies if a pannel has been found
	panel_in_edit_mode = False								'defaulting the boolean to identify that this case is not in edit mode
	If panel_found = False and end_date = "" Then			'if the panel has not been found using the function 'find_UNEA_panel' and the income does not appear ended - we make a new panel
		Call write_value_and_transmit("NN", 20, 79)
		panel_in_edit_mode = True
	ElseIf panel_found = True Then							'if the correct panel was found, the script will put it in edit mode
		PF9
		panel_in_edit_mode = True
	End If
	If panel_in_edit_mode = True Then						'Once the panel is created or in edit mode, we can add information
		If claim_number <> "" Then							'If a claim number was entered into the function call we will update the claim number information
			Call clear_line_of_text(6, 37)					'delete the existing claim number
			EMWriteScreen claim_number, 6, 37				'enter the claim number from the function call
		End If
		EMWriteScreen unea_type, 5, 37						'writing the unea type into the panel (this usually won't change if it was existing but it is necessary for new panels)
		EMWriteScreen "7", 5, 65							'Write Verification Worker Initiated Verfication "7"
		'NOTE - THIS FUNCTION IS NOT IN USE FOR THESE INCOME TYPES AT THIS TIME. TODO - additional review and testing would be needed to use this function for these types
		If unea_type = "11" or unea_type = "12" or unea_type = "13" or unea_type = "38" Then EMWriteScreen "6", 5, 65

		'Handling of the Start date will need to be dealth with differently depending on if the panel is new or is being updated
		If panel_found = False Then							'add a known start date if the panel is new
			Call create_mainframe_friendly_date(start_date, 7, 37, "YY") 	'income start date (SSI: ssi_SSP_elig_date, RSDI: intl_entl_date)
		Else												'if the panel is being updated, we need to read the start date
			EMReadScreen start_date, 8, 7, 37
			start_date = replace(start_date, " ", "/")
			start_date = DateAdd("d", 0 , start_date)
		End If

		'Now we clear the information for COLA and end date information
		Call clear_line_of_text(10, 67)		'clear the COLA disregard - TODO - update this for Jan - June to not remove this
		' Call clear_line_of_text(7, 68)
		' Call clear_line_of_text(7, 71)
		' Call clear_line_of_text(7, 74)

		'Clear amount details
		row = 13									'row 13 is the top of the income information
		DO
			EMWriteScreen "__", row, 25				'remove the information that was previously in the panel for each date and amount field
			EMWriteScreen "__", row, 28
			EMWriteScreen "__", row, 31
			EMWriteScreen "________", row, 39

			EMWriteScreen "__", row, 54
			EMWriteScreen "__", row, 57
			EMWriteScreen "__", row, 60
			EMWriteScreen "________", row, 68
			row = row + 1
		Loop until row = 18							'loop through removing the detail in each row from 13 through 17, leave the loop once we get to 18

		'creating a date that is the start of the retro month
		retro_date = CM_minus_1_mo & "/1/" & CM_minus_1_yr
		retro_date = DateAdd("d", 0, retro_date)

		If end_date <> "" Then						'if there is an end date, we need to enter the end date and be careful about if we enter income information
			EMReadScreen curr_end_date, 8, 7, 68
			If curr_end_date = "__ __ __" Then
				'enter the end date
				Call create_mainframe_friendly_date(end_date, 7, 68, "YY")	'income end date (SSI: ssi_denial_date, RSDI: susp_term_date)
				'determine the footer month of the end date
				Call convert_date_into_MAXIS_footer_month(end_date, footer_month_end, footer_year_end)

				If footer_month_end = CM_plus_1_mo and footer_year_end = CM_plus_1_yr Then		'if the footer month of the end date is the same as CM plus 1, we should enter the last pay information
					Call create_mainframe_friendly_date(end_date, 13, 54, "YY")					'enter the last pay date
					EMWriteScreen last_pay, 13, 68												'enter the last pay amount
				End If
			End If
		Else										'if there is not an end date, we need to code in the current payement amounts
			Call clear_line_of_text(7, 68)
			Call clear_line_of_text(7, 71)
			Call clear_line_of_text(7, 74)
			If DateDiff("d", start_date, retro_date) >= 0 Then								'This ensures the start date is before the retro month. If it is, income is entered in the retro month fields
				'the month and year are hardcoded using the CM_minus_1 global variable
				EMWriteScreen CM_minus_1_mo, 13, 25
				EMWriteScreen "01", 13, 28					'TODO - maybe we should look at changing this for some income types
				EMWriteScreen CM_minus_1_yr, 13, 31
				EMWriteScreen income_amount, 13, 39			'this is the income amount passed through the function call as an argument
			End If
			'the month and year are hardcoded using the CM_plus_1 global variable
			EMWriteScreen CM_plus_1_mo, 13, 54
			EMWriteScreen "01", 13, 57						'TODO - maybe we should look at changing this for some income types
			EMWriteScreen CM_plus_1_yr, 13, 60
			EMWriteScreen income_amount, 13, 68				'this is the income amount passed through the function call as an argument

			'This part of the code is for the HC Income Estimate and will open the pop-up
			Call write_value_and_transmit("X", 6, 56)
			Call clear_line_of_text(9, 65)					'empty the current information in the pop-up
			EMWriteScreen income_amount, 9, 65				'enter the current income amount passed through the function call as an argument
			EMWriteScreen "1", 10, 63						'code for pay frequency
			Do
				transmit									'this should return to the main panel
				EMReadScreen HC_popup, 9, 7, 41				'check to make sure that the pop-up closed
			Loop until HC_popup <> "HC Income"
		End If

		'this part will save the information entered into the panel
		transmit
		EMReadScreen cola_warning, 29, 24, 2
		If cola_warning = "WARNING: ENTER COLA DISREGARD" then transmit
		EMReadScreen HC_income_warning, 25, 24, 2
		If HC_income_warning = "WARNING: UPDATE HC INCOME" then transmit
	End If
end function
'END FUNCTIONS BLOCK =======================================================================================================


'DECLARATIONS ==============================================================================================================

Const memb_ref_numb_const 	= 0
Const memb_pmi_numb_const 	= 1
Const memb_ssn_const 		= 2
Const memb_age_const 		= 3
Const memb_name_const 		= 4
Const memb_active_hc_const	= 5
Const table_prog_1			= 6
Const table_type_1			= 7
Const table_prog_2			= 8
Const table_type_2			= 9
Const table_prog_3			= 10
Const table_type_3			= 11
Const memb_smi_numb_const	= 12

Const unea_type_01_esists	= 20
Const unea_type_02_esists	= 21
Const unea_type_03_esists	= 22
Const unea_type_16_esists	= 23
Const unmatched_claim_numb	= 24
Const unea_VA_exists		= 25
Const unea_UC_exists		= 26
Const unea_RR_exists		= 27

Const sves_qury_sent		= 35
Const second_qury_sent		= 36
Const sves_tpqy_response	= 37
Const sql_uc_income_exists	= 38
Const sql_va_income_exists	= 39
Const sql_rr_income_exists	= 40
Const tpqy_date_of_death	= 41

Const tpqy_rsdi_record 				= 45
Const tpqy_ssi_record 				= 46
Const tpqy_rsdi_claim_numb 			= 47
Const tpqy_dual_entl_nbr 			= 48
Const tpqy_rsdi_status_code 		= 49
Const tpqy_rsdi_gross_amt 			= 50
Const tpqy_rsdi_net_amt 			= 51
Const tpqy_railroad_ind 			= 52
Const tpqy_intl_entl_date 			= 53
Const tpqy_susp_term_date 			= 54
Const tpqy_rsdi_disa_date 			= 55
Const tpqy_medi_claim_num 			= 56
Const tpqy_part_a_premium 			= 57
Const tpqy_part_a_start 			= 58
Const tpqy_part_a_stop 				= 59
Const tpqy_part_a_buyin_ind 		= 60
Const tpqy_part_a_buyin_code 		= 61
Const tpqy_part_a_buyin_start_date 	= 62
Const tpqy_part_a_buyin_stop_date 	= 63
Const tpqy_part_b_premium 			= 64
Const tpqy_part_b_start 			= 65
Const tpqy_part_b_stop 				= 66
Const tpqy_part_b_buyin_ind 		= 67
Const tpqy_Part_b_buyin_code 		= 68
Const tpqy_part_b_buyin_start_date 	= 69
Const tpqy_part_b_buyin_stop_date 	= 70
Const tpqy_ssi_claim_numb 			= 71
Const tpqy_ssi_recip_code 			= 72
Const tpqy_ssi_recip_desc 			= 73
Const tpqy_fed_living 				= 74
Const tpqy_ssi_pay_code 			= 75
Const tpqy_ssi_pay_desc 			= 76
Const tpqy_cit_ind_code 			= 77
Const tpqy_ssi_denial_code 			= 78
Const tpqy_ssi_denial_desc 			= 79
Const tpqy_ssi_denial_date 			= 80
Const tpqy_ssi_disa_date 			= 81
Const tpqy_ssi_SSP_elig_date 		= 82
Const tpqy_ssi_appeals_code 		= 83
Const tpqy_ssi_appeals_date 		= 84
Const tpqy_ssi_appeals_dec_code 	= 85
Const tpqy_ssi_appeals_dec_date 	= 86
Const tpqy_ssi_disa_pay_code 		= 87
Const tpqy_ssi_pay_date 			= 88
Const tpqy_ssi_gross_amt 			= 89
Const tpqy_ssi_over_under_code 		= 90
Const tpqy_ssi_pay_hist_1_date 		= 91
Const tpqy_ssi_pay_hist_1_amt 		= 92
Const tpqy_ssi_pay_hist_1_type 		= 93
Const tpqy_ssi_pay_hist_2_date 		= 94
Const tpqy_ssi_pay_hist_2_amt 		= 95
Const tpqy_ssi_pay_hist_2_type 		= 96
Const tpqy_ssi_pay_hist_3_date 		= 97
Const tpqy_ssi_pay_hist_3_amt 		= 98
Const tpqy_ssi_pay_hist_3_type 		= 99
Const tpqy_gross_EI 				= 100
Const tpqy_net_EI 					= 101
Const tpqy_rsdi_income_amt 			= 102
Const tpqy_pass_exclusion 			= 103
Const tpqy_inc_inkind_start 		= 104
Const tpqy_inc_inkind_stop 			= 105
Const tpqy_rep_payee 				= 106
Const tpqy_ssi_last_pay_date		= 107
Const tpqy_ssi_is_ongoing			= 108
COnst tpqy_ssi_last_pay_amt			= 109

Const tpqy_memb_has_ssi				= 110
Const tpqy_memb_has_rsdi			= 111
Const tpqy_rsdi_has_disa			= 112
Const created_medi					= 113
Const updated_medi_a				= 114
Const updated_medi_b				= 115


Const memb_last_const 		= 120

Dim MEMBER_INFO_ARRAY()


Const va_case_numb_const 		= 0
Const va_ref_numb_const 		= 1
Const va_pers_name_const		= 2
Const va_pers_pmi_const			= 3
Const va_pers_ssn_const			= 4
Const va_inc_type_code_const 	= 5
Const va_inc_type_info_const	= 6
Const va_claim_numb_const 		= 7
Const va_prosp_inc_const 		= 8
Const va_end_date_const			= 9
Const va_panel_updated_const 	= 10
Const va_last_const 			= 11

Dim VA_INCOME_ARRAY()
ReDim VA_INCOME_ARRAY(va_last_const, 0)

Const uc_case_numb_const 		= 0
Const uc_ref_numb_const 		= 1
Const uc_pers_name_const		= 2
Const uc_pers_pmi_const			= 3
Const uc_pers_ssn_const			= 4
Const uc_inc_type_code_const 	= 5
Const uc_inc_type_info_const	= 6
Const uc_claim_numb_const 		= 7
Const uc_prosp_inc_const 		= 8
Const uc_end_date_const			= 9
Const uc_panel_updated_const 	= 10
Const uc_last_const 			= 11

Dim UC_INCOME_ARRAY()
ReDim UC_INCOME_ARRAY(uc_last_const, 0)

Const rr_case_numb_const 		= 0
Const rr_ref_numb_const 		= 1
Const rr_pers_name_const		= 2
Const rr_pers_pmi_const			= 3
Const rr_pers_ssn_const			= 4
Const rr_inc_type_code_const 	= 5
Const rr_inc_type_info_const	= 6
Const rr_claim_numb_const 		= 7
Const rr_prosp_inc_const 		= 8
Const rr_end_date_const			= 9
Const rr_panel_updated_const 	= 10
Const rr_last_const 			= 11

Dim RR_INCOME_ARRAY()
ReDim RR_INCOME_ARRAY(rr_last_const, 0)

const panel_type_const			= 0
const panel_claim_const			= 1
const panel_claim_left_9_const	= 2
const panel_instance_const 		= 3
const last_panel_const			= 4

dim unea_panel_array()

'Setting constants
Const adOpenStatic = 3
Const adLockOptimistic = 3

'These are the constants that we need to create tables in Excel
Const xlSrcRange = 1
Const xlYes = 1

'END DECLARATIONS BLOCK ====================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

Confirm_Process_to_Run_btn	= 200		'setting the button values
incorrect_process_btn		= 100
end_msg = "DONE"						'making sure there is something in the end message for when the script run completes. This is often overwritten.

MAXIS_footer_month = CM_plus_1_mo		'We are always operating in Current Month plus 1 while runing this script
MAXIS_footer_year = CM_plus_1_yr

'This is the file path for the Excel files that are created/updated during the different script run options.
ex_parte_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte"

'Dialog to select which script operation we need to run
DO
	DO
		DO
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 401, 300, "Ex Parte Report"
				DropListBox 200, 25, 190, 45, "Select one..."+chr(9)+"Prep 1"+chr(9)+"Prep 2"+chr(9)+"Phase 1"+chr(9)+"Phase 2"+chr(9)+"ADMIN Review"+chr(9)+"FIX LIST"+chr(9)+"Check REVW information on Phase 1 Cases"+chr(9)+"DHS Data Validation"+chr(9)+"Ex Parte Eval Case Review"+chr(9)+"Evaluate DHS Error List", ex_parte_function
				ButtonGroup ButtonPressed
					OkButton 290, 280, 50, 15
					CancelButton 345, 280, 50, 15
				Text 5, 10, 400, 10, "This script will connect to the SQL Table to pull a list of cases to operate on based on the Ex Parte functionality selected."
				Text 100, 30, 95, 10, "Selection Ex Parte Function:"
				Text 10, 45, 35, 10, "Prep"
				Text 50, 45, 150, 10, "PREP requires TWO runs."
				Text 50, 55, 200, 10, "Timing - prior to the last week of the PREP Month"
				Text 50, 65, 190, 10, "Review any Case Criteria not available in Info Store."
				Text 50, 75, 175, 10, "Send SVES/QURY for all members on all cases."
				Text 50, 85, 200, 10, "Generate a UC and VA Verif Report for OS Staff completion."
				Text 50, 95, 200, 10, "Makes the initial determination of Ex Parte"
				Text 10, 110, 35, 10, "Phase 1"
				Text 50, 110, 200, 10, "Timing - last couple days of the PREP Month"
				Text 50, 120, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
				Text 50, 130, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
				Text 50, 140, 125, 10, "Run each case through Background."
				Text 10, 155, 35, 10, "Phase 2"
				Text 50, 155, 160, 10, "Timing - 1st Day of the PROCESSING Month"
				Text 50, 165, 285, 10, "Update STAT/BUDG to start with the Renewal Month."
				Text 50, 175, 145, 10, "Run each case through Background."
				Text 50, 190, 150, 10, "--- Additional Administrative Runs ---"
				Text 10, 205, 270, 10, "ADMIN Review  -  Display of current Ex Parte progress counts."
				Text 10, 220, 270, 10, "FIX LIST  -  MANAU CODE UPDATE REQUIRED. Used to resolve the SQL Table."
				Text 10, 235, 270, 10, "Check REVW information on Phase 1 Cases  -  Report on Ex Parte and REVW."
				Text 10, 250, 270, 10, "DHS Data Validation  -  Compare SQL Ex Parte list to the DHS Ex Parte List."
				Text 10, 265, 350, 10, "Ex Parte Eval Case Review  -  Displays the evaluation of Ex Parte details for a single case."
				Text 10, 280, 205, 10, "* * * * * THIS SCRIPT MUST BE RUN IN PRODUCTION * * * * *"
			EndDialog

			err_msg = ""
			Dialog Dialog1
			cancel_without_confirmation
			If ex_parte_function = "Select one..." then err_msg = err_msg & vbNewLine & "* Select an Ex Parte Function."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP until err_msg = ""

		'Once a selection is made, for the BULK runs, we need to confirm the details about the script run selected.
		If ex_parte_function <> "ADMIN Review" and ex_parte_function <> "Ex Parte Eval Case Review" and ex_parte_function <>"Evaluate DHS Error List" Then
			allow_bulk_run_use = False												'At this time the BULK runs can only be completed by AIT due to database access.
			If user_ID_for_validation = "CALO001" Then allow_bulk_run_use = True
			If user_ID_for_validation = "ILFE001" Then allow_bulk_run_use = True
			If user_ID_for_validation = "MARI001" Then allow_bulk_run_use = True
			If user_ID_for_validation = "MEGE001" Then allow_bulk_run_use = True
			If user_ID_for_validation = "DACO003" Then allow_bulk_run_use = True

			'stopping the script run if someone else tries to run the script for the bulk options.
			If allow_bulk_run_use = False Then script_end_procedure("Ex Parte Report functionality for completing Ex Parte actions and list review is locked. The script will now end.")

			'This part sets the footer month to run based on the selection made
			If ex_parte_function = "Prep 1" or ex_parte_function = "Prep 2" or ex_parte_function = "FIX LIST" or ex_parte_function = "DHS Data Validation" Then
				ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 3, date)), 2)
				ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 3, date)), 2)
			End If
			If ex_parte_function = "Phase 1" Then
				ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 3, date)), 2)
				ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 3, date)), 2)
			End If
			If ex_parte_function = "Check REVW information on Phase 1 Cases" Then
				ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 2, date)), 2)
				ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 2, date)), 2)
			End If
			If ex_parte_function = "Phase 2" Then											'setting the dates for a Phase 2 BULK run
				ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 1, date)), 2)			'This is the ex parte renewal month when running Phase 2 BULK
				ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 1, date)), 2)
				ep_end_budg_revw_mo = right("00" & DatePart("m",	DateAdd("m", 6, date)), 2)	'this sets the last month of the budget period. We need this to update STAT/BUDG
				ep_end_budg_revw_yr = right(DatePart("yyyy",	DateAdd("m", 6, date)), 2)
			End If

			'Dialog to confirm the BULK run selected
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 341, 165, "Confirm Ex Parte process"
				EditBox 600, 700, 10, 10, fake_edit_box
				If ex_parte_function <> "DHS Data Validation" and ex_parte_function <> "Check REVW information on Phase 1 Cases" and ex_parte_function <> "FIX LIST" Then Checkbox 10, 115, 330, 10, "Check here to clear any previous 'In Progress' statuses on cases in the Data Table.", reset_in_Progress
				ButtonGroup ButtonPressed
					PushButton 10, 145, 210, 15, "CONFIRMED! This is the correct Process and Review Month", Confirm_Process_to_Run_btn
					PushButton 230, 145, 100, 15, "Incorrect Process/Month", incorrect_process_btn
				Text 10, 10, 275, 10, "You are running the Ex Parte Function " & ex_parte_function
				Text 10, 25, 190, 10, "This will run for the Ex Parte Review month of " & ep_revw_mo & "/" & ep_revw_yr

				If ex_parte_function = "Prep 1" Then
					GroupBox 5, 40, 240, 60, "Tasks to be Completed:"
					Text 20, 55, 190, 10, "Collect any Case Criteria not available in Info Store."
					Text 20, 65, 175, 10, "Send SVES/QURY for all members on all cases."
					Text 20, 75, 200, 10, "Generate a UC, VA, and RR Verif Report for OS Staff completion."
					Text 20, 85, 200, 10, "Create a list of SMRT ending members."
				End If
				If ex_parte_function = "DHS Data Validation" Then
					GroupBox 5, 40, 240, 75, "Tasks to be Skipped:"
					Text 20, 55, 190, 10, "Compare Hennepin Ex Parte list to the cases from the DHS list."
					CheckBox 25, 65, 150, 10, "First SQL review is done", sql_reviewed_checkbox
					CheckBox 25, 75, 150, 10, "HC Details already gathered", hc_elig_reviewed_checkbox
					CheckBox 25, 85, 150, 10, "Missing cases have been added", missing_cases_added_checkbox
					Text 25, 100, 110, 10, "What row should we start at?"
					EditBox 135, 95, 30, 15, excel_starting_row
				End If
				If ex_parte_function = "Prep 2" Then
					GroupBox 5, 40, 320, 50, "Tasks to be Completed:"
					Text 20, 55, 270, 10, "Read SVES/TPQY and update UNEA with the response information."
					Text 20, 65, 300, 20, "Send SVES/QURY for all members whose TPQY indicates a second associated RSDI Claim."
					Text 20, 75, 270, 10, "Generate a list of all members with a Date of Death in TPQY."
				End If
				If ex_parte_function = "FIX LIST" Then
					Text 20, 55, 270, 10, "FIX LIST HAS NO SET FUNCTIONALITY"
				End If
				If ex_parte_function = "Evaluate DHS Error List" Then
					Text 20, 55, 270, 10, "Built to review our information and compare it to a DHS error list."
					Text 20, 65, 270, 10, "REVIEW FUNCTIONALITY BEFORE USE."
					Text 20, 75, 270, 10, "Functionality may need alteration based on the list provided"
				End If
				If ex_parte_function = "Phase 1" Then
					GroupBox 5, 40, 295, 70, "Tasks to be Completed:"
					Text 20, 55, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
					Text 20, 65, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
					Text 20, 75, 125, 10, "Run each case through Background."
					Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
					Text 20, 95, 225, 10, "Read and Record in the SQL Table the detail of MMIS Open Spans."
				End If
				If ex_parte_function = "Check REVW information on Phase 1 Cases" Then
					GroupBox 5, 40, 240, 50, "Tasks to be Completed:"
					Text 20, 55, 270, 10, "Pull a list of all cases for the Ex Parte month from the SQL Table."
					Text 20, 65, 270, 10, "Pull all cases from REPT/REVS for the Ex Parte month from MAXIS."
					Text 20, 75, 270, 10, "Read Ex Parte information from STAT/REVW."
					Text 20, 85, 270, 10, "Connect cases on both lists and output everything to Excel"
				End If
				If ex_parte_function = "Phase 2" Then
					GroupBox 5, 40, 305, 60, "Tasks to be Completed:"
					Text 20, 55, 285, 10, "Update STAT/BUDG to align with the Ex Parte Renewal Month"
					Text 20, 65, 145, 10, "Run each case through Background."
				End If

				Text 10, 130, 330, 10, "Review the process datails and ex parte review month to confirm this is the correct run to complete."
			EndDialog

			Dialog Dialog1
			cancel_without_confirmation

			If ButtonPressed = OK Then ButtonPressed = Confirm_Process_to_Run_btn
		Else
			ButtonPressed = Confirm_Process_to_Run_btn
		End If

	Loop until ButtonPressed = Confirm_Process_to_Run_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Display the details used in Prep 1 to determine if a case is Ex Parte or not. For information/review only.
If ex_parte_function = "Ex Parte Eval Case Review" Then
	'This functionality is locked down and only available for use by certain staff.
	allow_admin_use = False
	If user_ID_for_validation = "CALO001" Then allow_admin_use = True
	If user_ID_for_validation = "ILFE001" Then allow_admin_use = True
	If user_ID_for_validation = "MARI001" Then allow_admin_use = True
	If user_ID_for_validation = "MEGE001" Then allow_admin_use = True
	If user_ID_for_validation = "LALA004" Then allow_admin_use = True
	If user_ID_for_validation = "WFX901" Then allow_admin_use = True
	If user_ID_for_validation = "BETE001" Then allow_admin_use = True
	If user_ID_for_validation = "DACO003" Then allow_bulk_run_use = True

	'Ending the script run if someone else tries to run it
	If allow_admin_use = False Then script_end_procedure("ADMIN function for reviewing Ex Parte Functionality is locked. The script will now end.")

	Call MAXIS_case_number_finder(MAXIS_case_number)						'attempting to pull the Case Number from MAXIS
	Do
		Do
			'dialog to confirm the case number of the case to review. It will default to whatever case is currently entered in MAXIS
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 261, 100, "Case Number Selection"
				ButtonGroup ButtonPressed
					OkButton 150, 75, 50, 15
					CancelButton 200, 75, 50, 15
				Text 10, 10, 245, 20, "This functionality will review a single case to display the evaluation used in the initial PREP run to identify why it was selected or Ex Parte or not."
				Text 15, 40, 50, 10, "Case Number:"
				EditBox 70, 35, 50, 15, MAXIS_case_number
				Text 15, 60, 200, 10, "Ente the CASE NUMBER you want to have the script check."
			EndDialog

			dialog Dialog1
			cancel_without_confirmation

			Call validate_MAXIS_case_number(err_msg, "*")

			iF ERR_MSG <> "" Then MsgBox "* * * * NOTICE * * * *" & vbCr & err_msg

		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in


	MAXIS_case_number = right("00000000"&MAXIS_case_number, 8)
	ex_parte_renewal_date = ""			'blank out these variables to ensure there is no carry over data
	SQL_case_status = ""
	SQL_select_ex_parte = ""
	SQL_prep_complete = ""
	SQL_phase1_complete = ""
	SQL_ex_parte_after_phase1 = ""
	SQL_phase1_cancel_reason = ""
	SQL_phase2_complete = ""
	SQL_ex_parte_after_phase2 = ""
	SQL_phase2_cancel_reason = ""
	SQL_all_HC_is_ABD = ""
	SQL_ssa_income_exists = ""
	SQL_wages_exist = ""
	SQL_va_inc_exists = ""
	SQL_self_emp_exists = ""
	SQL_no_income = ""
	SQL_EPD_on_case = ""
	SQL_year_month = ""
	SQL_eval_year_month = ""
	SQL_app_year_month = ""



	'This is opening the Ex Parte Case List data table so we can loop through it.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE CaseNumber = '" & MAXIS_case_number & "'"		'we only need to look at the cases for the specific review month

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	Do While NOT objRecordSet.Eof 					'Loop through each item on the CASE LIST Table

		ex_parte_renewal_date 		= objRecordSet("HCEligReviewDate")
		SQL_case_status 			= objRecordSet("CaseStatus")
		SQL_select_ex_parte 		= objRecordSet("SelectExParte")
		SQL_prep_complete 			= objRecordSet("PREP_Complete")
		SQL_phase1_complete 		= objRecordSet("Phase1Complete")
		SQL_ex_parte_after_phase1 	= objRecordSet("ExParteAfterPhase1")
		SQL_phase1_cancel_reason 	= objRecordSet("Phase1ExParteCancelReason")
		SQL_phase2_complete 		= objRecordSet("Phase2Complete")
		SQL_ex_parte_after_phase2 	= objRecordSet("ExParteAfterPhase2")
		SQL_phase2_cancel_reason 	= objRecordSet("Phase2ExParteCancelReason")
		SQL_all_HC_is_ABD 			= objRecordSet("AllHCisABD")
		SQL_ssa_income_exists 		= objRecordSet("SSAIncomExist")
		SQL_wages_exist 			= objRecordSet("VAIncomeExist")
		SQL_va_inc_exists 		= objRecordSet("VAIncomeExist")
		SQL_self_emp_exists 		= objRecordSet("SelfEmpExists")
		SQL_no_income 				= objRecordSet("NoIncome")
		SQL_EPD_on_case 			= objRecordSet("EPDonCase")
		SQL_year_month 				= objRecordSet("YearMonth")
		SQL_eval_year_month 		= objRecordSet("EvaluationYearMonth")
		SQL_app_year_month 			= objRecordSet("ApprovalYearMonth")



		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
		EMReadScreen case_pw, 7, 21, 14

		Call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)
		If is_this_priv = True Then appears_ex_parte = False						'excluding cases that are privileged
		If is_this_priv = False Then
			Call write_value_and_transmit("X", 5, 71)
			EMReadScreen STAT_HC_ER_mo, 2, 8, 27
			EMReadScreen STAT_HC_ER_yr, 2, 8, 33
			If ep_revw_mo <> STAT_HC_ER_mo or ep_revw_yr <> STAT_HC_ER_yr Then  appears_ex_parte = False		'if this does not have the correct renewal month, we will exclude it from Ex Parte
		End If


		ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)			'This is defined here without a preserve to blank it out at the beginning of every loop with a new case
		memb_count = 0										'resetting the counting variable to size the member array
		list_of_membs_on_hc = " "							'we need to keep a list members by pmi to know if a person is already accounted for as we find all the members and programs

		'We need to pull all of the instances from the ELIG table for the currently defined case number
		'This will list the HH member and eligibility program for HC. We will use this to start to determine if the case can be processed as Ex Parte
		objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & MAXIS_case_number & "'"

		Set objELIGConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
		Set objELIGRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objELIGRecordSet.Open objELIGSQL, objELIGConnection

		person_found = False		'setting the default of if we have found a person in the list
		Do While NOT objELIGRecordSet.Eof
			list_of_membs_on_hc = list_of_membs_on_hc & objELIGRecordSet("PMINumber") & " "		'adding the PMI to the list of all PMIs known on the case
			person_found = True																	'indicating that there was a person in the list for this case
			memb_known = False																	'sets that we don't know if we have already looked at this person
			'now we loop through all of the people we have already found for this case - we only want 1 array instance per person.
			For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				If trim(objELIGRecordSet("PMINumber")) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then		'If the PMI matches one in the array, we are going to set the information to that array instance
					memb_known = True															'identifies that we know about this person and they are already in the array

					'figuring out which program type location the information should be saved in for this table data
					'each person on a case may have up to three different lines for different programs
					If MEMBER_INFO_ARRAY(table_prog_1, known_membs) = "" Then
						MEMBER_INFO_ARRAY(table_prog_1, known_membs) 		= objELIGRecordSet("MajorProgram")
						MEMBER_INFO_ARRAY(table_type_1, known_membs) 		= objELIGRecordSet("EligType")
					ElseIf MEMBER_INFO_ARRAY(table_prog_2, known_membs) = "" Then
						MEMBER_INFO_ARRAY(table_prog_2, known_membs) 		= objELIGRecordSet("MajorProgram")
						MEMBER_INFO_ARRAY(table_type_2, known_membs) 		= objELIGRecordSet("EligType")
					ElseIf MEMBER_INFO_ARRAY(table_prog_3, known_membs) = "" Then
						MEMBER_INFO_ARRAY(table_prog_3, known_membs) 		= objELIGRecordSet("MajorProgram")
						MEMBER_INFO_ARRAY(table_type_3, known_membs) 		= objELIGRecordSet("EligType")
					End If
				End If
			Next

			'If this is an unknown member, and has not been added to the array already, we need to add it
			If memb_known = False Then
				ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)								'resizing the array

				'setting personal information to the array
				MEMBER_INFO_ARRAY(memb_pmi_numb_const, memb_count) 	= trim(objELIGRecordSet("PMINumber"))
				MEMBER_INFO_ARRAY(memb_ssn_const, memb_count) 		= trim(objELIGRecordSet("SocialSecurityNbr"))
				name_var									 		= trim(objELIGRecordSet("Name"))		'we want to format the name corectly.
				name_array = split(name_var)
				MEMBER_INFO_ARRAY(memb_name_const, memb_count) = name_array(UBound(name_array))
				For name_item = 0 to UBound(name_array)-1
					MEMBER_INFO_ARRAY(memb_name_const, memb_count) = MEMBER_INFO_ARRAY(memb_name_const, memb_count) & " " & name_array(name_item)
				Next
				MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
				MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(objELIGRecordSet("MajorProgram"))	'setting the program information
				MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(objELIGRecordSet("EligType"))

				memb_count = memb_count + 1		'incrementing the array counter up for the next loop
			End if
			objELIGRecordSet.MoveNext			'going to the next record
		Loop
		objELIGRecordSet.Close			'Closing all the data connections
		objELIGConnection.Close
		Set objELIGRecordSet=nothing
		Set objELIGConnection=nothing



		'If the case still appears Ex Parte, we are going to check if we are missing people, and check income for further determination of Ex Parte
		If appears_ex_parte = True Then
			'If we did not find people in the ELIG list, we are going to check ELIG/HC
			If person_found = False Then
				Call navigate_to_MAXIS_screen("STAT", "SUMM")		'Creating new ELIG results
				Call write_value_and_transmit("BGTX", 20, 71)

				Call MAXIS_background_check

				Call navigate_to_MAXIS_screen("ELIG", "HC  ")		'Navigate to ELIG/HC
				'Here we start at the top of ELIG/HC and read each row to find HC information
				hc_row = 8
				Do
					pers_type = ""		'blanking out variables so they don't carry over from loop to loop
					std = ""
					meth = ""
					waiv = ""

					'reading the main HC Elig information - member, program, status
					EMReadScreen read_ref_numb, 2, hc_row, 3
					EMReadScreen clt_hc_prog, 4, hc_row, 28
					EMReadScreen hc_prog_status, 6, hc_row, 50
					ref_row = hc_row
					Do while read_ref_numb = "  "				'this will read for the reference number if there are multiple programs for a single member
						ref_row = ref_row - 1
						EMReadScreen read_ref_numb, 2, ref_row, 3
					Loop

					If hc_prog_status = "ACTIVE" Then			'If HC is currently active, we need to read more details about the program/eligibility
						clt_hc_prog = trim(clt_hc_prog)			'formatting this to remove whitespace
						If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "" Then		'these are non-hc persons

							Call write_value_and_transmit("X", hc_row, 26)									'opening the ELIG detail spans
							If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then		'If it is an MSP, we want to read the type only from a specific place
								elig_msp_prog = clt_hc_prog
								EMReadScreen pers_type, 2, 6, 56
							Else																			'These are MA type programs (not MSP)
								'Now we have to fund the current month in elig to get the current elig type
								col = 19
								Do
									EMReadScreen span_month, 2, 6, col										'reading the month in ELIG
									EMReadScreen span_year, 2, 6, col+3

									'if the span month matchest current month plus 1, we are going to grab elig from that month
									If span_month = MAXIS_footer_month and span_year = MAXIS_footer_year Then
										EMReadScreen pers_type, 2, 12, col - 2								'reading the ELIG TYPE
										EMReadScreen std, 1, 12, col + 3
										EMReadScreen meth, 1, 13, col + 2
										EMReadScreen waiv, 1, 17, col + 2
										Exit Do																'leaving once we've found the information for this elig
									End If
									col = col + 11			'this goes to the next column
								Loop until col = 85			'This is off the page - if we hit this, we did NOT find the elig type in this elig result

								'If we hit 85, we did not get the information. So we are going to read it from the last budget month (most current)
								If col = 85 Then
									EMReadScreen pers_type, 2, 12, 72										'reading the ELIG TYPE
									EMReadScreen std, 1, 12, 77
									EMReadScreen meth, 1, 13, 76
									EMReadScreen waiv, 1, 17, 76
								End If
							End If
							PF3			'leaving the elig detail information

							'now we need to add the information we just read to the member array
							memb_known = False										'default that the member know is false
							For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)								'Looking at all the members known in the array
								If MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs) = read_ref_numb Then	'if the member reference from ELIG matches the ARRAY, we are going to add more elig details
									memb_known = True														'look we found a person
									If MEMBER_INFO_ARRAY(table_prog_1, known_membs) = "" Then				'finding which area of the array is blank to save the elig information there
										MEMBER_INFO_ARRAY(table_prog_1, known_membs) 		= clt_hc_prog
										MEMBER_INFO_ARRAY(table_type_1, known_membs) 		= pers_type
									ElseIf MEMBER_INFO_ARRAY(table_prog_2, known_membs) = "" Then
										MEMBER_INFO_ARRAY(table_prog_2, known_membs) 		= clt_hc_prog
										MEMBER_INFO_ARRAY(table_type_2, known_membs) 		= pers_type
									ElseIf MEMBER_INFO_ARRAY(table_prog_3, known_membs) = "" Then
										MEMBER_INFO_ARRAY(table_prog_3, known_membs) 		= clt_hc_prog
										MEMBER_INFO_ARRAY(table_type_3, known_membs) 		= pers_type
									End If
								End If
							Next

							'If this is an unknown member, and has not been added to the array already, we need to add it
							If memb_known = False Then
								ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)								'resizing the array

								'setting personal information to the array
								MEMBER_INFO_ARRAY(memb_ref_numb_const, memb_count) = read_ref_numb
								MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
								MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(clt_hc_prog)
								MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(pers_type)

								memb_count = memb_count + 1 	'incrementing the array counter up for the next loop
							End If

						End If
					End If
					hc_row = hc_row + 1												'now we go to the next row
					EMReadScreen next_ref_numb, 2, hc_row, 3						'read the next HC information to find when we've reeached the end of the list
					EMReadScreen next_maj_prog, 4, hc_row, 28
				Loop until next_ref_numb = "  " and next_maj_prog = "    "

				CALL back_to_SELF()													'going to STAT/MEMB - because there is misssing personal information for the members discovered in this way
				Do
					CALL navigate_to_MAXIS_screen("STAT", "MEMB")
					EMReadScreen memb_check, 4, 2, 48
				Loop until memb_check = "MEMB"

				at_least_one_hc_active = False										'this is a default to identify if HC is active on the case
				For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)					'loop through the member array
					Call write_value_and_transmit(MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs), 20, 76)		'navigate to the member for this instance of the array
					EMReadscreen last_name, 25, 6, 30								'read and cormat the name from MEMB
					EMReadscreen first_name, 12, 6, 63
					last_name = trim(replace(last_name, "_", "")) & " "
					first_name = trim(replace(first_name, "_", "")) & " "
					MEMBER_INFO_ARRAY(memb_name_const, known_membs) = first_name & " " & last_name
					EMReadScreen PMI_numb, 8, 4, 46									'capturing the PMI number
					PMI_numb = trim(PMI_numb)
					MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) = right("00000000" & PMI_numb, 8)			'we have to format the pmi to match the data list format (8 digits with leading 0s included)
					EMReadScreen MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), 11, 7, 42							'catpturing the SSN
					MEMBER_INFO_ARRAY(memb_ssn_const, known_membs) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), " ", "")
					MEMBER_INFO_ARRAY(memb_ssn_const, known_membs) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), "_", "")
					If MEMBER_INFO_ARRAY(table_prog_1, known_membs) <> "" Then at_least_one_hc_active = True		'setting the variable that identifies there is HC active based on the ELIG read from HC/ELIG
					If MEMBER_INFO_ARRAY(table_prog_2, known_membs) <> "" Then at_least_one_hc_active = True
					If MEMBER_INFO_ARRAY(table_prog_3, known_membs) <> "" Then at_least_one_hc_active = True
					If MEMBER_INFO_ARRAY(table_prog_1, known_membs) <> "" or MEMBER_INFO_ARRAY(table_prog_2, known_membs) <> "" or MEMBER_INFO_ARRAY(table_prog_3, known_membs) <> "" Then
						list_of_membs_on_hc = list_of_membs_on_hc & MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) & " "		'adding individuals to our list of members on HC
					End If

				Next
				If at_least_one_hc_active = False Then appears_ex_parte = False			'if no one is on HC, this cannot be Ex Parte
			End If
		End If

		Call navigate_to_MAXIS_screen("STAT", "MEMB")		'now we go find all the HH members
		Call get_list_of_members

		'Now we are going to start looking at income information to remove any cases that have income thant disqualifies it from Ex parte
		SSA_income_exists = False				'setting these variables to false at the beginning of each loop through
		RR_income_exists = False
		VA_income_exists = False
		UC_income_exists = False
		PRISM_income_exists = False
		Other_UNEA_income_exists = False
		JOBS_income_exists = False
		BUSI_income_exists = False

		'Pulling all rows from the INCOME list for the case number we are currently processing
		objIncomeSQL = "SELECT * FROM ES.ES_ExParte_IncomeList WHERE [CaseNumber] = '" & MAXIS_case_number & "'"

		Set objIncomeConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
		Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

		'This will create an array of all of the income listed in SQL - which is what this report uses to complete the evaluation.
		Const sql_pers_name 	= 0
		Const sql_pers_pmi 		= 1
		Const sql_ref_numb 		= 2
		Const sql_inc_panel 	= 3
		Const sql_inc_type 		= 4
		Const sql_inc_desc 		= 5
		Const sql_clm_numb 		= 6
		Const sql_prosp_amt 	= 7
		Const sql_qury_sent 	= 8
		Const sql_tpqy_resp		= 9
		Const sql_tpqy_grs_amt	= 10
		Const sql_tpqy_net_amt 	= 11
		Const sql_tpqy_end_dt	= 12
		Const last_sql_const 	= 30

		Dim SQL_INCOME_ARRAY()
		ReDim SQL_INCOME_ARRAY(last_sql_const, 0)
		inc_count = 0

		'looping through each row in this case
		Do While NOT objIncomeRecordSet.Eof
			ReDim Preserve SQL_INCOME_ARRAY(last_sql_const, inc_count)			'saving each item from SQL Income list for this case into the array

			SQL_INCOME_ARRAY(sql_pers_name, 	inc_count) = objIncomeRecordSet("PersName")
			SQL_INCOME_ARRAY(sql_pers_pmi,  	inc_count) = objIncomeRecordSet("PersonID")
			SQL_INCOME_ARRAY(sql_ref_numb,  	inc_count) = trim(objIncomeRecordSet("RefNumb"))
			SQL_INCOME_ARRAY(sql_inc_panel,  	inc_count) = objIncomeRecordSet("IncExpTypeCode")
			SQL_INCOME_ARRAY(sql_inc_type,  	inc_count) = objIncomeRecordSet("IncomeTypeCode")
			SQL_INCOME_ARRAY(sql_inc_desc,  	inc_count) = objIncomeRecordSet("Description")
			SQL_INCOME_ARRAY(sql_clm_numb,  	inc_count) = objIncomeRecordSet("ClaimNbr")
			SQL_INCOME_ARRAY(sql_prosp_amt,  	inc_count) = objIncomeRecordSet("ProspAmount")
			SQL_INCOME_ARRAY(sql_qury_sent,  	inc_count) = objIncomeRecordSet("QURY_Sent")
			SQL_INCOME_ARRAY(sql_tpqy_resp, 	inc_count) = objIncomeRecordSet("TPQY_Response")
			SQL_INCOME_ARRAY(sql_tpqy_grs_amt,	inc_count) = objIncomeRecordSet("GrossAmt")
			SQL_INCOME_ARRAY(sql_tpqy_net_amt,	inc_count) = objIncomeRecordSet("NetAmt")
			SQL_INCOME_ARRAY(sql_tpqy_end_dt,	inc_count) = objIncomeRecordSet("EndDate")
			If DateDiff("d", SQL_INCOME_ARRAY(sql_tpqy_end_dt, inc_count), #1/1/1900#) = 0 Then SQL_INCOME_ARRAY(sql_tpqy_end_dt, inc_count) = ""

			SQL_INCOME_ARRAY(sql_ref_numb,  	inc_count) = right(SQL_INCOME_ARRAY(sql_ref_numb,  	inc_count), 2)

			If objIncomeRecordSet("IncomeTypeCode") = "16" Then RR_income_exists = True				'identifying some income types
			If objIncomeRecordSet("IncomeTypeCode") = "14" Then UC_income_exists = True

			If objIncomeRecordSet("IncomeTypeCode") = "36" Then PRISM_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "37" Then PRISM_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "39" Then PRISM_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "40" Then PRISM_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "36" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "37" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "39" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "40" Then Other_UNEA_income_exists = True

			If objIncomeRecordSet("IncomeTypeCode") = "06" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "15" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "17" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "18" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "23" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "24" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "25" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "26" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "27" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "28" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "29" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "08" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "35" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "43" Then Other_UNEA_income_exists = True
			If objIncomeRecordSet("IncomeTypeCode") = "47" Then Other_UNEA_income_exists = True

			inc_count = inc_count + 1
			objIncomeRecordSet.MoveNext		'move to the next Income row
		Loop
		objIncomeRecordSet.Close			'Closing all the data connections
		objIncomeConnection.Close
		Set objIncomeRecordSet=nothing
		Set objIncomeConnection=nothing

		objRecordSet.MoveNext			'now we go to the next case
	Loop

	'stopping the script if the case was not found.
	If ex_parte_renewal_date = "" Then call script_end_procedure("The case " & MAXIS_case_number & " is not listed on the SQL Ex Parte Data Table.")

	'Here we show the information about the case that was found in SQL
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 450, 350, "Case Detials"
		Text 10, 10, 130, 10, "Case Number: " & MAXIS_case_number
		Text 20, 20, 130, 10, "CASE/CURR X Numb: " & case_pw
		Text 10, 35, 120, 10, "Appears Ex Parte: " & SQL_select_ex_parte
		Text 35, 45, 105, 10, "All HC is ABD: " & SQL_all_HC_is_ABD
		Text 30, 55, 105, 10, " Case has EPD: " & SQL_EPD_on_case
		Text 15, 70, 105, 10, "SQL HC ER: " & ex_parte_renewal_date
		Text 10, 80, 105, 10, "STAT HC ER: " & STAT_HC_ER_mo & "/" & STAT_HC_ER_yr

		Text 10, 95, 90, 10, "Case Active: " & case_active
		Text 20, 105, 90, 10, "MA Status: " & ma_status
		Text 15, 115, 90, 10, "MSP Status: " & msp_status
		Text 15, 125, 90, 10, "MFIP Status: " & mfip_status
		Text 10, 135, 90, 10, " SNAP Status: " & snap_status
		Text 120, 95, 90, 10, " SSA Income: " & SQL_ssa_income_exists
		Text 125, 105, 90, 10, "RR Income: " & RR_income_exists
		Text 125, 115, 90, 10, " VA Income: " & SQL_va_inc_exists
		Text 125, 125, 90, 10, " UC Income: " & UC_income_exists
		Text 230, 95, 90, 10, " PRISM Income: " & PRISM_income_exists
		Text 240, 105, 90, 10, "Other UNEA: " & Other_UNEA_income_exists
		Text 235, 115, 90, 10, " JOBS Income: " & SQL_wages_exist
		Text 235, 125, 90, 10, "  BUSI Income: " & SQL_self_emp_exists
		Text 150, 10, 50, 10, "Persons"
		y_pos = 25
		For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
			If MEMBER_INFO_ARRAY(table_prog_1, each_memb) <> "" Then
				Text 150, y_pos, 105, 10, MEMBER_INFO_ARRAY(memb_name_const, each_memb)
				Text 255, y_pos, 35, 10, MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
				Text 310, y_pos, 30, 10, MEMBER_INFO_ARRAY(table_prog_1, each_memb) & "-" & MEMBER_INFO_ARRAY(table_type_1, each_memb)
				y_pos = y_pos + 10
				If MEMBER_INFO_ARRAY(table_prog_2, each_memb) <> "" Then
					Text 310, y_pos, 30, 10, MEMBER_INFO_ARRAY(table_prog_2, each_memb) & "-" & MEMBER_INFO_ARRAY(table_type_2, each_memb)
					y_pos = y_pos + 10
				End If
				If MEMBER_INFO_ARRAY(table_prog_3, each_memb) <> "" Then
					Text 310, y_pos, 30, 10, MEMBER_INFO_ARRAY(table_prog_3, each_memb) & "-" & MEMBER_INFO_ARRAY(table_type_3, each_memb)
					y_pos = y_pos + 10
				End If
			End If
		Next
		y_pos = y_pos + 10

		y_pos = 150
		Text 155, y_pos, 50, 10, "SQL Income"
		y_pos = y_pos + 10
		for each_inc = 0 to UBound(SQL_INCOME_ARRAY, 2)
			Text 15, y_pos, 200, 10, "MEMB " & SQL_INCOME_ARRAY(sql_ref_numb, each_inc) & " - " & SQL_INCOME_ARRAY(sql_inc_panel, each_inc) & " - " & SQL_INCOME_ARRAY(sql_inc_type, each_inc) & " (" & SQL_INCOME_ARRAY(uc_inc_type_code_const, each_inc) &")"
			Text 220, y_pos, 100, 10, "SQL amount $ " & SQL_INCOME_ARRAY(sql_prosp_amt, each_inc)
			If SQL_INCOME_ARRAY(sql_clm_numb, each_inc) <> "" Then Text 325, y_pos, 150, 10, "Claim Number: " &  SQL_INCOME_ARRAY(sql_clm_numb, each_inc)
			y_pos = y_pos + 10
			If SQL_INCOME_ARRAY(sql_qury_sent, each_inc) <> "" or SQL_INCOME_ARRAY(sql_tpqy_resp, each_inc) <> "" Then
				Text 25, y_pos, 185, 10, "QURY Date: " & SQL_INCOME_ARRAY(sql_qury_sent, each_inc) & " - TPQY Date: " & SQL_INCOME_ARRAY(sql_tpqy_resp, each_inc)
				Text 200, y_pos, 140, 10, "TPQY Gross: $ " & SQL_INCOME_ARRAY(sql_tpqy_grs_amt, each_inc) & ", Net: $ " & SQL_INCOME_ARRAY(sql_tpqy_net_amt, each_inc)
				If SQL_INCOME_ARRAY(sql_tpqy_end_dt, each_inc) <> "" Then Text 340, y_pos, 120, 10, "TPQY End Date: " & SQL_INCOME_ARRAY(sql_tpqy_end_dt, each_inc)
				y_pos = y_pos + 10
			End If
		next
		ButtonGroup ButtonPressed
			OkButton 390, 330, 50, 15
	EndDialog

	Dialog Dialog1			'show the dialog
	'there are no inputs or other loops in this script, it will just end.

	Call script_end_procedure("")
End If

If ex_parte_function = "ADMIN Review" Then
	'this functionality is meant to review the status of cases on the SQL data list. This can help track the progress on Ex Parte cases.

	'This functionality is locked down and only available for use by certain staff.
	allow_admin_use = False
	If user_ID_for_validation = "CALO001" Then allow_admin_use = True
	If user_ID_for_validation = "ILFE001" Then allow_admin_use = True
	If user_ID_for_validation = "MARI001" Then allow_admin_use = True
	If user_ID_for_validation = "MEGE001" Then allow_admin_use = True
	If user_ID_for_validation = "LALA004" Then allow_admin_use = True
	If user_ID_for_validation = "WFX901" Then allow_admin_use = True
	If user_ID_for_validation = "BETE001" Then allow_admin_use = True
	If user_ID_for_validation = "DACO003" Then allow_bulk_run_use = True

	If allow_admin_use = False Then script_end_procedure("ADMIN function for reviewing Ex Parte Functionality is locked. The script will now end.")

	'First we need to set the dates for each phase of Ex Parte.
	current_month_revw = CM_mo & "/1/" & CM_yr
	next_month_revw = CM_plus_1_mo & "/1/" & CM_plus_1_yr
	month_after_next_revw = CM_plus_2_mo & "/1/" & CM_plus_2_yr
	phase_one_hard_stop_date = CM_mo & "/15/" & CM_yr
	prep_month_revw = CM_plus_3_mo & "/1/" & CM_plus_3_yr
	current_month_revw = DateAdd("d", 0, current_month_revw)
	next_month_revw = DateAdd("d", 0, next_month_revw)
	month_after_next_revw = DateAdd("d", 0, month_after_next_revw)
	phase_one_hard_stop_date = DateAdd("d", 0, phase_one_hard_stop_date)
	prep_month_revw = DateAdd("d", 0, prep_month_revw)

	PREP_PHASE_MO = CM_plus_3_mo & "/" & CM_plus_3_yr		'now we are creating strings with the months to display which month is in which phase during the dialog
	PHASE_ONE_MO = CM_plus_2_mo & "/" & CM_plus_2_yr
	PHASE_TWO_MO = CM_plus_1_mo & "/" & CM_plus_1_yr
	COMPLETED_MO = CM_mo & "/" & CM_yr

	Phase_one_hard_stop_passed = False						'defining if we have passed er cut off for Phase 1 work
	If DateDiff("d", phase_one_hard_stop_date, date) > 0 Then Phase_one_hard_stop_passed = True

	'declare the SQL statement that will query the database - we need to pull cases from 3 different review months.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE HCEligReviewDate = '" & next_month_revw & "' or HCEligReviewDate = '" & month_after_next_revw & "' or HCEligReviewDate = '" & prep_month_revw & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'Opening the SQL data path for Ex Parte
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'This array is to maintain a list of workers that are processing Ex Parte
	const worker_numb_const 		= 0
	const worker_name_const 		= 1
	const case_complete_p1_count 	= 2
	const case_complete_p2_count	= 3
	const case_phase_const 			= 4
	Dim HSR_WORK_ARRAY()
	ReDim HSR_WORK_ARRAY(case_phase_const, 0)

	'Setting intial numbers as counts for Ex parte case scenarios and types.
	next_month_er_count = 0
	next_month_still_expt = 0
	next_month_hsr_phase2_complete_count = 0
	next_month_need_to_work = 0
	next_month_need_more_review = 0
	next_month_app_count = 0
	next_month_rescheduled_count = 0
	next_month_closed_xfer_count = 0

	month_after_next_er_count = 0
	month_after_next_expt_at_prep = 0
	month_after_next_hsr_phase1_complete_count = 0
	month_after_next_complete_and_expt = 0
	month_after_next_still_expt = 0
	month_after_next_need_to_work = 0

	prep_month_er_count = 0
	prep_month_still_need_eval = 0
	prep_month_expt_at_prep = 0

	'Starting values for looking through all of the cases.
	list_of_hsrs = " "
	hsr_count = 0
	Do While NOT objRecordSet.Eof		'here is where we look at all of the cases to count where everything is at and determine the workers
		'PHASE 2 CASE INFORMATION
		If DateDiff("d", objRecordSet("HCEligReviewDate"), next_month_revw) = 0 Then
			next_month_er_count = next_month_er_count + 1
			If objRecordSet("SelectExParte") = True Then next_month_still_expt = next_month_still_expt + 1
			If IsNull(objRecordSet("Phase2HSR")) = False and trim(objRecordSet("Phase2HSR")) <> "" Then
				next_month_hsr_phase2_complete_count = next_month_hsr_phase2_complete_count + 1
				If objRecordSet("ExParteAfterPhase2") = "REVIEW" Then next_month_need_more_review = next_month_need_more_review + 1
				If objRecordSet("ExParteAfterPhase2") = "Approved as Ex Parte" Then next_month_app_count = next_month_app_count + 1
				If objRecordSet("ExParteAfterPhase2") = "Closed HC" Then next_month_closed_xfer_count = next_month_closed_xfer_count + 1
				If objRecordSet("ExParteAfterPhase2") = "Case not in 27" Then next_month_closed_xfer_count = next_month_closed_xfer_count + 1
				If InStr(objRecordSet("ExParteAfterPhase2"), "ER Scheduled") <> 0 Then next_month_rescheduled_count = next_month_rescheduled_count + 1

				case_phase_two_hsr = objRecordSet("Phase2HSR")
				If InStr(list_of_hsrs, case_phase_two_hsr) = 0 Then
					ReDim Preserve HSR_WORK_ARRAY(case_phase_const, hsr_count)
					HSR_WORK_ARRAY(worker_numb_const, hsr_count) = case_phase_two_hsr
					HSR_WORK_ARRAY(case_complete_p1_count, hsr_count) = 0
					HSR_WORK_ARRAY(case_complete_p2_count, hsr_count) = 1
					hsr_count = hsr_count + 1
					list_of_hsrs = list_of_hsrs & case_phase_two_hsr & " "
				Else
					For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
						If HSR_WORK_ARRAY(worker_numb_const, each_worker) = case_phase_two_hsr Then HSR_WORK_ARRAY(case_complete_p2_count, each_worker) = HSR_WORK_ARRAY(case_complete_p2_count, each_worker) + 1
					Next
				End If
			Else
				If objRecordSet("SelectExParte") = True Then next_month_need_to_work = next_month_need_to_work + 1
			End If
			' If objRecordSet("SelectExParte") = True Then month_after_next_still_expt = month_after_next_still_expt + 1
		End If

		'PHASE 1 CASE INFORMATION
		If DateDiff("d", objRecordSet("HCEligReviewDate"), month_after_next_revw) = 0 Then
			month_after_next_er_count = month_after_next_er_count + 1
			If IsDate(objRecordSet("PREP_Complete")) = True and IsDate(objRecordSet("Phase1Complete")) = True Then month_after_next_expt_at_prep = month_after_next_expt_at_prep + 1
			If IsNull(objRecordSet("Phase1HSR")) = False and trim(objRecordSet("Phase1HSR")) <> "" Then
				month_after_next_hsr_phase1_complete_count = month_after_next_hsr_phase1_complete_count + 1
				If objRecordSet("SelectExParte") = True Then month_after_next_complete_and_expt = month_after_next_complete_and_expt + 1
				case_phase_one_hsr = objRecordSet("Phase1HSR")
				If InStr(list_of_hsrs, case_phase_one_hsr) = 0 Then
					ReDim Preserve HSR_WORK_ARRAY(case_phase_const, hsr_count)
					HSR_WORK_ARRAY(worker_numb_const, hsr_count) = case_phase_one_hsr
					HSR_WORK_ARRAY(case_complete_p1_count, hsr_count) = 1
					HSR_WORK_ARRAY(case_complete_p2_count, hsr_count) = 0
					hsr_count = hsr_count + 1
					list_of_hsrs = list_of_hsrs & case_phase_one_hsr & " "
				Else
					For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
						If HSR_WORK_ARRAY(worker_numb_const, each_worker) = case_phase_one_hsr Then HSR_WORK_ARRAY(case_complete_p1_count, each_worker) = HSR_WORK_ARRAY(case_complete_p1_count, each_worker) + 1
					Next
				End If
			Else
				If objRecordSet("SelectExParte") = True Then month_after_next_need_to_work = month_after_next_need_to_work + 1
			End If
			If objRecordSet("SelectExParte") = True Then month_after_next_still_expt = month_after_next_still_expt + 1
		End If

		'PREP CASE INFORMATION
		If DateDiff("d", objRecordSet("HCEligReviewDate"), prep_month_revw) = 0 Then
			prep_month_er_count = prep_month_er_count + 1
			If objRecordSet("SelectExParte") = True Then prep_month_expt_at_prep = prep_month_expt_at_prep + 1
			If IsNull(objRecordSet("PREP_Complete")) = True or objRecordSet("PREP_Complete") = "" Then prep_month_still_need_eval = prep_month_still_need_eval + 1
		End If

		objRecordSet.MoveNext		'go to the next case
	Loop
    objRecordSet.Close				'close the data connection
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	'here we calculate the percentage of cases for a number of the counts
	function calculate_percent(numerator, denominator, percent)
		percent = numerator/denominator
		percent = percent * 100
		percent = FormatNumber(percent, 2, -1, 0, -1)
	end function
	If prep_month_er_count <> 0 Then call calculate_percent(prep_month_expt_at_prep, prep_month_er_count, prep_month_percent_ex_parte_pcnt)
	call calculate_percent(month_after_next_expt_at_prep, month_after_next_er_count, month_after_next_initially_expt_pcnt)
	call calculate_percent(month_after_next_hsr_phase1_complete_count, month_after_next_expt_at_prep, month_after_next_processed_pcnt)
	call calculate_percent(month_after_next_need_to_work, month_after_next_expt_at_prep, month_after_next_waiting_pcnt)
	If month_after_next_hsr_phase1_complete_count <> 0 Then
		call calculate_percent(month_after_next_complete_and_expt, month_after_next_hsr_phase1_complete_count, month_after_next_complete_and_expt_pcnt)
	End If

	call calculate_percent(next_month_still_expt, next_month_er_count, next_month_initially_expt_pcnt)
	call calculate_percent(next_month_hsr_phase2_complete_count, next_month_still_expt, next_month_processed_pcnt)
	call calculate_percent(next_month_need_to_work, next_month_still_expt, next_month_waiting_pcnt)

	If next_month_hsr_phase2_complete_count <> 0 Then
		call calculate_percent(next_month_app_count, next_month_hsr_phase2_complete_count, next_month_app_pcnt)
		call calculate_percent(next_month_rescheduled_count, next_month_hsr_phase2_complete_count, next_month_rescheduled_pcnt)
		call calculate_percent(next_month_closed_xfer_count, next_month_hsr_phase2_complete_count, next_month_closed_xfer_pcnt)
		call calculate_percent(next_month_need_more_review, next_month_hsr_phase2_complete_count, next_problem_pcnt)
	End If

	'Now we are going to put a name to the worker ID that is entered int he Ex parte list
	SQL_table = "SELECT * from ES.V_ESAllStaff"				'identifying the table that stores the ES Staff user information

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path the data tables
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open SQL_table, objConnection							'Here we connect to the data tables

	Do While NOT objRecordSet.Eof										'now we will loop through each item listed in the table of ES Staff
		Name_array = ""
		For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
			If HSR_WORK_ARRAY(worker_numb_const, each_worker) = objRecordSet("EmpLogOnID") Then		'If the ID number is found, we will get the name
				HSR_WORK_ARRAY(worker_name_const, each_worker) = objRecordSet("EmpFullName")
				If InStr(HSR_WORK_ARRAY(worker_name_const, each_worker), ",") <> 0 Then				'this will format the name to be easier to read in the dialog display
					Name_array = split(HSR_WORK_ARRAY(worker_name_const, each_worker), ",")
					HSR_WORK_ARRAY(worker_name_const, each_worker) = trim(Name_array(1)) & " " & trim(Name_array(0))
				End If
			End If
			If HSR_WORK_ARRAY(worker_numb_const, each_worker) = "BULK Script" Then HSR_WORK_ARRAY(worker_name_const, each_worker) = "BULK Script"	'this is for the cases that were updated by a BULK run
		Next
		objRecordSet.MoveNext											'Going to the next row in the table
	Loop

	'Now we disconnect from the table and close the connections
	objRecordSet.Close
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing

	'Now we need to resize the dialog
	phase_1_factor = 0
	phase_2_factor = 0
	For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
		If HSR_WORK_ARRAY(case_complete_p1_count, each_worker) <> 0 Then phase_1_factor = phase_1_factor + 1
		If HSR_WORK_ARRAY(case_complete_p2_count, each_worker) <> 0 Then phase_2_factor = phase_2_factor + 1
	Next
	If phase_1_factor < 10 Then phase_1_factor = 10
	If phase_2_factor < 6 Then phase_2_factor = 6
	If next_month_hsr_phase2_complete_count <> 0 Then
		If phase_2_factor < 14 Then phase_2_factor = 14
	End If
	If phase_1_factor mod 2 = 1 Then phase_1_factor = phase_1_factor + 1
	If phase_2_factor mod 2 = 1 Then phase_2_factor = phase_2_factor + 1
	' MsgBox "phase_1_factor - " & phase_1_factor & vbCr & "phase_2_factor - " & phase_2_factor
	dlg_len = 180 + (phase_1_factor/2)*10 + (phase_2_factor/2)*10

	If dlg_len < 180 Then dlg_len = 185

	'display the counts and information gathered from the data list in a dialog
	BeginDialog Dialog1, 0, 0, 500, dlg_len, "Ex Parte Work Details"
		ButtonGroup ButtonPressed
			'PREP PHASE
			GroupBox 5, 10, 250, 50, "PREP Phase - " & PREP_PHASE_MO
			Text 15, 25, 155, 10, "Total Cases with HC ER in " & PREP_PHASE_MO & ": " & prep_month_er_count
			If prep_month_expt_at_prep = 0 Then Text 15, 40, 155, 10, "PREP Run not completed."
			If prep_month_expt_at_prep <> 0 Then
				Text 15, 35, 155, 10, "Case that appear Ex Parte Eligible: " & prep_month_expt_at_prep
				Text 175, 35, 75, 10, "Percent: " & prep_month_percent_ex_parte_pcnt & " %"
				If prep_month_still_need_eval <> 0 Then
					Text 25, 45, 155, 10, "Cases that still need PREP run: " & prep_month_still_need_eval
				End If
			End If

			'PHASE 1
			Text 15, 75, 155, 10, "Total Cases with HC ER in " & PHASE_ONE_MO & ": " & month_after_next_er_count
			Text 15, 90, 175, 10, "Cases that appeared Ex Parte at PREP: " & month_after_next_expt_at_prep		'" - XX%"
			Text 115, 100, 175, 10, "Percent: " & month_after_next_initially_expt_pcnt & " %"
			Text 15, 115, 165, 10, "Cases with Phase 1 completed by HSR: " & month_after_next_hsr_phase1_complete_count
			Text 115, 125, 165, 10, "Percent: " & month_after_next_processed_pcnt & " %"
			Text 15, 140, 205, 10, "Cases processed and passed: " & month_after_next_complete_and_expt & "    ( " & month_after_next_complete_and_expt_pcnt & " % )"
			' Text 145, 150, 75, 10, "( " & month_after_next_complete_and_expt_pcnt & " % )"
			If Phase_one_hard_stop_passed = True Then Text 15, 155, 165, 10, "Phase One Processing has stopped."

			y_pos = 85
			x_pos = 185
			For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
				If HSR_WORK_ARRAY(case_complete_p1_count, each_worker) <> 0 Then
					If len(HSR_WORK_ARRAY(case_complete_p1_count, each_worker)) = 4 Then Text x_pos, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p1_count, each_worker)
					If len(HSR_WORK_ARRAY(case_complete_p1_count, each_worker)) = 3 Then Text x_pos+5, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p1_count, each_worker)
					If len(HSR_WORK_ARRAY(case_complete_p1_count, each_worker)) = 2 Then Text x_pos+10, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p1_count, each_worker)
					If len(HSR_WORK_ARRAY(case_complete_p1_count, each_worker)) = 1 Then Text x_pos+15, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p1_count, each_worker)
					Text x_pos+25, y_pos, 115, 10, HSR_WORK_ARRAY(worker_name_const, each_worker)
					If x_pos = 185 Then
						x_pos = 350
					Else
						x_pos = 185
						y_pos = y_pos + 10
					End If
				End If
			Next
			If x_pos = 350 Then y_pos = y_pos + 10
			y_pos = y_pos + 5
			If month_after_next_need_to_work <> 0 Then
				Text 180, y_pos, 130, 10, "Cases to Still Process in Phase 1: " & month_after_next_need_to_work
				Text 350, y_pos, 130, 10, "Percent: " & month_after_next_waiting_pcnt & " %"
				If month_after_next_need_to_work < 30 Then
					PushButton 15, y_pos-3, 150, 13, "Export list of Unprocessed Phase 1", export_list_of_phase_1
				End If
			Else
				Text 180, y_pos, 260, 10, "All Ex Parte Evaluation cases for " & PHASE_ONE_MO & " have been completed."
			End If
			GroupBox 180, 75, 306, y_pos-75, "Count"
			Text 210, 75, 20, 10, "Name"
			Text 350, 75, 20, 10, "Count"
			Text 375, 75, 20, 10, "Name"

			y_pos = y_pos + 15
			If y_pos < 155 Then y_pos = 155
			' MsgBox "1 - y_pos - " & y_pos
			' If y_pos = 125 Then y_pos = 165
			GroupBox 5, 65, 485, y_pos-65, "PHASE ONE - " & PHASE_ONE_MO

			y_pos = y_pos + 15

			'PHASE 2
			start_y_pos = y_pos
			set_y_pos = y_pos - 10
			Text 15, y_pos, 155, 10, "Total Cases with HC ER in " & PHASE_TWO_MO & ": " & next_month_er_count
			y_pos = y_pos + 15
			Text 15, y_pos, 175, 10, "Cases Ex Parte after Phase ONE: " & next_month_still_expt		'" - XX%"
			y_pos = y_pos + 10
			Text 95, y_pos, 175, 10, "Percent: " & next_month_initially_expt_pcnt & " %"
			y_pos = y_pos + 15
			Text 15, y_pos, 165, 10, "Cases with Phase 2 completed by HSR: " & next_month_hsr_phase2_complete_count
			y_pos = y_pos + 10
			Text 115, y_pos, 165, 10, "Percent: " & next_month_processed_pcnt & " %"
			y_pos = y_pos + 15
			If next_month_hsr_phase2_complete_count <> 0 Then
				Text 15, y_pos, 205, 10, "Cases Approved for " & PHASE_TWO_MO & ": " & next_month_app_count & "    ( " & next_month_app_pcnt & " % )"
				y_pos = y_pos + 10
				Text 15, y_pos, 205, 10, "Cases with ER Rescheduled : " & next_month_rescheduled_count & "    ( " & next_month_rescheduled_pcnt & " % )"
				y_pos = y_pos + 10
				Text 15, y_pos, 205, 10, "Cases closed/transferred : " & next_month_closed_xfer_count  & "    ( " & next_month_closed_xfer_pcnt & " % )"
				y_pos = y_pos + 10
				Text 15, y_pos, 205, 10, "Cases to on PROBLEM list: " & next_month_need_more_review & "    ( " & next_problem_pcnt & " % )"
			End If

			y_pos = set_y_pos
			y_pos = y_pos + 20
			x_pos = 185
			For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
				If HSR_WORK_ARRAY(case_complete_p2_count, each_worker) <> 0 Then
					If len(HSR_WORK_ARRAY(case_complete_p2_count, each_worker)) = 4 Then Text x_pos, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p2_count, each_worker)
					If len(HSR_WORK_ARRAY(case_complete_p2_count, each_worker)) = 3 Then Text x_pos+5, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p2_count, each_worker)
					If len(HSR_WORK_ARRAY(case_complete_p2_count, each_worker)) = 2 Then Text x_pos+10, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p2_count, each_worker)
					If len(HSR_WORK_ARRAY(case_complete_p2_count, each_worker)) = 1 Then Text x_pos+15, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p2_count, each_worker)
					Text x_pos+25, y_pos, 115, 10, HSR_WORK_ARRAY(worker_name_const, each_worker)
					If x_pos = 185 Then
						x_pos = 350
					Else
						x_pos = 185
						y_pos = y_pos + 10
					End If
				End If
			Next
			If x_pos = 350 Then y_pos = y_pos + 10
			y_pos = y_pos + 5
			If next_month_need_to_work <> 0 Then
				Text 180, y_pos, 130, 10, "Cases to Still Process in Phase 2: " & next_month_need_to_work
				Text 350, y_pos, 130, 10, "Percent: " & next_month_waiting_pcnt & " %"
				If next_month_need_to_work < 30 Then
					PushButton 15, y_pos-3, 150, 13, "Export list of Unprocessed Phase 2", export_list_of_phase_2
				End If
			Else
				Text 180, y_pos, 260, 10, "All Ex Parte Approval cases for " & PHASE_TWO_MO & " have been completed."
			End If
			GroupBox 180, set_y_pos+10, 306, y_pos-set_y_pos-10, "Count"
			Text 210, set_y_pos+10, 20, 10, "Name"
			Text 350, set_y_pos+10, 20, 10, "Count"
			Text 375, set_y_pos+10, 20, 10, "Name"
			y_pos = y_pos + 5
			If next_month_hsr_phase2_complete_count <> 0 Then
				If y_pos < start_y_pos+100 Then y_pos = start_y_pos+100
			Else
				If y_pos < start_y_pos+55 Then y_pos = start_y_pos+55
			End If
			' MsgBox "2 - y_pos - " & y_pos

			' If y_pos < 310 then y_pos = 310
			' If y_pos = 125 Then y_pos = 165
			GroupBox 5, set_y_pos, 485, y_pos-set_y_pos+10, "PHASE TWO - " & PHASE_TWO_MO

			OkButton 440, y_pos+15, 50, 15
	EndDialog

	Do
		Dialog Dialog1		'There is no looping and the dialog shows until the user presses OK or Cancel
		cancel_without_confirmation

		If ButtonPressed = export_list_of_phase_1 or ButtonPressed = export_list_of_phase_2 Then
			'Opening the Excel file
			Set objExcel = CreateObject("Excel.Application")
			objExcel.Visible = True
			Set objWorkbook = objExcel.Workbooks.Add()
			objExcel.DisplayAlerts = True

			'Setting the first 4 col as worker, case number, name, and APPL date
			ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
			If ButtonPressed = export_list_of_phase_1 Then ObjExcel.Cells(1, 3).Value = "Phase 1 Cases Not Completed in the Data Table - Ex Parte Month " & month_after_next_revw
			If ButtonPressed = export_list_of_phase_2 Then ObjExcel.Cells(1, 3).Value = "Phase 2 Cases Not Completed in the Data Table - Ex Parte Month " & next_month_revw
			ObjExcel.columns(1).AutoFit()
			ObjExcel.columns(3).AutoFit()
			excel_row = 2

			'declare the SQL statement that will query the database - we need to pull cases from 3 different review months.
			objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE HCEligReviewDate = '" & next_month_revw & "' or HCEligReviewDate = '" & month_after_next_revw & "' or HCEligReviewDate = '" & prep_month_revw & "'"

			'Creating objects for Access
			Set objConnection = CreateObject("ADODB.Connection")
			Set objRecordSet = CreateObject("ADODB.Recordset")

			'Opening the SQL data path for Ex Parte
			objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objRecordSet.Open objSQL, objConnection

			Do While NOT objRecordSet.Eof		'here is where we look at all of the cases to count where everything is at and determine the workers
				If ButtonPressed = export_list_of_phase_1 Then
					'PHASE 1 CASE INFORMATION
					If DateDiff("d", objRecordSet("HCEligReviewDate"), month_after_next_revw) = 0 Then
						If IsNull(objRecordSet("Phase1HSR")) = False and trim(objRecordSet("Phase1HSR")) <> "" Then
						Else
							If objRecordSet("SelectExParte") = True Then
								ObjExcel.Cells(excel_row, 1).Value = objRecordSet("CaseNumber")
								excel_row = excel_row + 1
							End If
						End If
					End If
				End If

				If ButtonPressed = export_list_of_phase_2 Then
					'PHASE 2 CASE INFORMATION
					If DateDiff("d", objRecordSet("HCEligReviewDate"), next_month_revw) = 0 Then
						If IsNull(objRecordSet("Phase2HSR")) = False and trim(objRecordSet("Phase2HSR")) <> "" Then
						Else
							If objRecordSet("SelectExParte") = True Then
								ObjExcel.Cells(excel_row, 1).Value = objRecordSet("CaseNumber")
								excel_row = excel_row + 1
							End If
						End If
					End If
				End If
				objRecordSet.MoveNext		'go to the next case
			Loop
			objRecordSet.Close				'close the data connection
			objConnection.Close
			Set objRecordSet=nothing
			Set objConnection=nothing
		End If
	Loop until ButtonPressed = -1

	end_msg = ""
	Call script_end_procedure(end_msg)		'That's all in the ADMIN run
End If

bz_user = False
If user_ID_for_validation = "CALO001" Then bz_user = True
If user_ID_for_validation = "ILFE001" Then bz_user = True
If user_ID_for_validation = "MARI001" Then bz_user = True
If user_ID_for_validation = "MEGE001" Then bz_user = True
If user_ID_for_validation = "DACO003" Then bz_user = True
If bz_user = False Then script_end_procedure("This script functionality can only be operated by the BlueZone Script Team. The script will now end.")

If ex_parte_function = "FIX LIST" Then
	Call script_end_procedure("There is no fix currently established.")
	fix_report_out = ""

	'THIS FIX IS TO FIND DUPLICATE SSA PANELS AND DELETE THEM
	' review_date = "2023-11-01"
	' 'This is opening the Ex Parte Case List data table so we can loop through it.
	' objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "' and [SelectExParte] = '1'"

	' Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	' Set objRecordSet = CreateObject("ADODB.Recordset")

	' 'opening the connections and data table
	' objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	' objRecordSet.Open objSQL, objConnection

	' 'Loop through each item on the CASE LIST Table
	' Do While NOT objRecordSet.Eof
	' 	'For each case that is indicated as Ex parte, we are going to update the case information
	' 	MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

	' 	'Here is functionality to be sure the case is able to be updated
	' 	case_is_in_henn = False					'default this to false

	' 	'reading case program information and PW
	' 	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
	' 	EMReadScreen case_pw, 7, 21, 14									'reading the curent PW for the case
	' 	If left(case_pw, 4) = "X127" Then case_is_in_henn = True		'identifying if the case is not in HENN
	' 	kick_it_off_reason = ""											'create an explanation of why the case is being removed form the Ex Parte list
	' 	If case_is_in_henn = False Then kick_it_off_reason = "Case not in 27"
	' 	If case_active = False Then kick_it_off_reason = "Case not Active"
	' 	If (case_active = False and case_pending = False and case_rein = False) or case_is_in_henn = False Then
	' 		'WE ARE NOT going to update this here for now
	' 		' select_ex_parte = False
	' 		' objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & select_ex_parte & "', PREP_Complete = '" & kick_it_off_reason & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

	' 		' Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
	' 		' Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

	' 		' 'opening the connections and data table
	' 		' objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	' 		' objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
	' 	Else
	' 		ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)							'Reset this array to blank at the beginning of each loop for each case.
	' 		Do
	' 			Call navigate_to_MAXIS_screen("STAT", "MEMB")					'making suyre we get to STAT MEMB
	' 			EMReadScreen memb_check, 4, 2, 48
	' 		Loop until memb_check = "MEMB"
	' 		Call get_list_of_members											'get a list of all the HH memebers on the case

	' 		'Read SVES/TPQY for all persons on a case
	' 		For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
	' 			MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False			'defaulting these to false
	' 			MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = False
	' 			MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = False

	' 			Call navigate_to_MAXIS_screen("INFC", "SVES")					'navigate to SVES
	' 			EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68				'Enter the PMI for the current member and open the TPQY
	' 			Call write_value_and_transmit("TPQY", 20, 70)

	' 			Do
	' 				EMReadScreen check_TPQY_panel, 4, 2, 53 						'Reads for TPQY panel
	' 				If check_TPQY_panel <> "TPQY" Then Call write_value_and_transmit("TPQY", 20, 70)
	' 			Loop until check_TPQY_panel = "TPQY"

	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb), 		1, 8, 39		'saving all tpqy information into the member array
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb), 		1, 8, 65
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), 		10, 6, 61
	' 			MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), " ", "/")
	' 			EMReadScreen sves_response, 8, 7, 22 		'Return Date
	' 			sves_response = replace(sves_response," ", "/")

	' 			transmit

	' 			Do
	' 				EMReadScreen check_BDXP_panel, 4, 2, 53 						'Reads fro BDXP panel\
	' 				If check_BDXP_panel <> "BDXP" Then
	' 					row = 1
	' 					col = 1
	' 					EMSearch "Command:", row, col
	' 					Call write_value_and_transmit("BDXP", row, col++9)
	' 				End If
	' 			Loop until check_BDXP_panel = "BDXP"

	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 	12, 5, 40		'saving all tpqy information into the member array
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 		12, 5, 69
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb), 	2, 6, 19
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), 	8, 8, 16
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb), 		8, 8, 32
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb), 		1, 8, 69
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), 	5, 11, 69
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), 	5, 14, 69
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), 	10, 15, 69
	' 			MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), " ", "")
	' 			MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = Trim (MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), " ", "/1/")
	' 			MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), " ", "/1/")
	' 			MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), " ", "/")

	' 			transmit

	' 			Do
	' 				EMReadScreen check_BDXM_panel, 4, 2, 53 						'Reads for BDXM panel
	' 				If check_BDXM_panel <> "BDXM" Then
	' 					row = 1
	' 					col = 1
	' 					EMSearch "Command:", row, col
	' 					Call write_value_and_transmit("BDXM", row, col++9)
	' 				End If
	' 			Loop until check_BDXM_panel = "BDXM"

	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb), 			13, 4, 29		'saving all tpqy information into the member array
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb), 			7, 6, 64
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 				5, 7, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), 				5, 7, 63
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_ind, each_memb), 			1, 8, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb), 			3, 8, 63
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb), 	5, 9, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb), 	5, 9, 63
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb), 			7, 12, 64
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 				5, 13, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), 				5, 13, 63
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_ind, each_memb), 			1, 14, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb), 			3, 14, 63
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), 	5, 15, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb), 	5, 15, 63
	' 			MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), " ", "/01/")
	' 			MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), " ", "/01/")
	' 			MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb), " ", "/01/")
	' 			MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb), " ", "/01/")
	' 			MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), " ", "/01/")
	' 			MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), " ", "/01/")
	' 			MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), " ", "/01/")
	' 			MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb), " ", "/01/")

	' 			transmit

	' 			Do
	' 				EMReadScreen check_SDXE_panel, 4, 2, 53 						'Reads for SDXE panel
	' 				If check_SDXE_panel <> "SDXE" Then
	' 					row = 1
	' 					col = 1
	' 					EMSearch "Command:", row, col
	' 					Call write_value_and_transmit("SDXE", row, col++9)
	' 				End If
	' 			Loop until check_SDXE_panel = "SDXE"

	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb), 		12, 5, 36		'saving all tpqy information into the member array
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), 		2, 7, 21
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb), 		22, 7, 24
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_fed_living, each_memb), 			1, 6, 70
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 			3, 8, 21
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb), 			30, 8, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_cit_ind_code, each_memb), 			1, 7, 70
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_code, each_memb), 		3, 10, 26
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb), 		40, 10, 30
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb), 		8, 11, 26
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb), 			8, 12, 26
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), 		8, 13, 26
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_code, each_memb), 		1, 11, 65
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb), 		8, 12, 65
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_code, each_memb), 	2, 13, 65
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb), 	8, 14, 65
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_disa_pay_code, each_memb), 		1, 15, 65
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb), " ", "/")
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb), " ", "/")
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), " ", "/")
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb), " ", "/")
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb), " ", "/")

	' 			transmit

	' 			Do
	' 				EMReadScreen check_SDXP_panel, 4, 2, 50 							'Reads for SDXP panel
	' 				If check_SDXP_panel <> "SDXP" Then
	' 					row = 1
	' 					col = 1
	' 					EMSearch "Command:", row, col
	' 					Call write_value_and_transmit("SDXP", row, col++9)
	' 				End If
	' 			Loop until check_SDXP_panel = "SDXP"

	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), 			5, 4, 16		'saving all tpqy information into the member array
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), 			7, 4, 42
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_over_under_code, each_memb), 	1, 4, 73
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb), 	5, 8, 3
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb), 	6, 8, 13
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_type, each_memb), 	1, 8, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb), 	5, 9, 3
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb), 	6, 9, 13
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_type, each_memb), 	1, 9, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb), 	5, 10, 3
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb), 	6, 10, 13
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_type, each_memb), 	1, 10, 25
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb), 				8, 5, 66
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb), 				8, 6, 66
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb), 		8, 7, 66
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb), 		8, 8, 66
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb), 		8, 9, 66
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb), 		8, 10, 66
	' 			EMReadScreen MEMBER_INFO_ARRAY(tpqy_rep_payee, each_memb), 				1, 11, 66

	' 			If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) <> "C01" Then
	' 				last_payment_date = ""
	' 				sdx_row = 8
	' 				Do
	' 					EMReadScreen sdx_payment_type, 1, sdx_row, 25
	' 					If sdx_payment_type <> "0" and sdx_payment_type <> " " Then
	' 						EMReadScreen last_payment_date, 5, sdx_row, 3
	' 						EMReadScreen last_payment_amt, 9, sdx_row, 13
	' 						Exit Do
	' 					End If
	' 					sdx_row = sdx_row + 1
	' 				Loop until sdx_payment_type = " "
	' 				If last_payment_date <> "" Then
	' 					MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb) = replace(last_payment_date, " ", "/1/")
	' 					MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb) = DateAdd("d", 0, MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb))
	' 					MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_amt, each_memb) = trim(last_payment_amt)
	' 				End If
	' 			End If

	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb))
	' 			If MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "" Then MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "0"
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb))
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), " ", "/01/")
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb), " ", "/")
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb), " ", "/")
	' 			MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb), " ", "/")
	' 			MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb), " ", "/")
	' 			MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb), " ", "/")

	' 			transmit

	' 			If MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb) = "Y" Then
	' 				MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True
	' 				MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb)= False
	' 				If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) = "C01" Then MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb) = True
	' 				If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "E" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
	' 				If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "H" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
	' 				If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "M" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
	' 				If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "P" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
	' 				If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "S" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
	' 			End If
	' 			If MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb) = "Y" Then
	' 				If MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "C" or MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "E" Then
	' 					MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True
	' 					If IsDate(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb)) = True Then MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True
	' 				End If
	' 			End If

	' 			Call back_to_SELF
	' 		Next

	' 		'navigating into STAT
	' 		Do
	' 			Call navigate_to_MAXIS_screen("STAT", "SUMM")
	' 			EMReadScreen summ_check, 4, 2, 46
	' 		Loop until summ_check = "SUMM"
	' 		verif_types = ""						'blanking out the list of verifications for the CASE/NOTE

	' 		'here we attempt to go update STAT with the information gathered from TPQY
	' 		For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)									'looping thorugh each HH Member
	' 			If IsDate(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb)) = True Then 		'If there is a date of dealth listed, for now we are just going to add them to a list
	' 				' ObjExcel.Cells(excel_row, 1).Value = MAXIS_case_number
	' 				' ObjExcel.Cells(excel_row, 2).Value = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
	' 				' ObjExcel.Cells(excel_row, 3).Value = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
	' 				' ObjExcel.Cells(excel_row, 4).Value = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
	' 				' ObjExcel.Cells(excel_row, 5).Value = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
	' 				' ObjExcel.Cells(excel_row, 6).Value = MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb)
	' 				' excel_row = excel_row + 1													'counting to increment to the next excel row
	' 			Else 	'If there is no date of death, we are going to try to update UNEA for SSI/RSDI
	' 				'Update MAXIS UNEA panels with information from TPQY
	' 				If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True Then				'Member with SSI
	' 					If MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb) = True Then		'If SSI appears to be ongoing (Current Pay)
	' 						MEMB_reference_number = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
	' 						UNEA_type_code = "03"
	' 						UNEA_claim_number = MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb)&MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb)

	' 						ReDim unea_panel_array(last_panel_const, 0)
	' 						unea_panel_counter = 0

	' 						EMWriteScreen "UNEA", 20, 71
	' 						transmit
	' 						EMReadScreen unea_check, 4, 2, 48
	' 						Do While unea_check <> "UNEA"
	' 							Call navigate_to_MAXIS_screen("STAT", "UNEA")
	' 							EMReadScreen unea_check, 4, 2, 48
	' 						Loop
	' 						EMWriteScreen MEMB_reference_number, 20, 76 		'Navigating to STAT/UNEA
	' 						EMWriteScreen "01", 20, 79 		'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	' 						transmit

	' 						list_of_panels_that_match = " "
	' 						EMReadScreen vers_count, 1, 2, 78
	' 						If vers_count <> "0" Then
	' 							Do
	' 								ReDim Preserve unea_panel_array(last_panel_const, unea_panel_counter)
	' 								EMReadScreen panel_instance, 1, 2, 73
	' 								EMReadScreen panel_type_code, 2, 5, 37
	' 								EMReadScreen panel_claim_number, 15, 6, 37
	' 								panel_claim_number = replace(panel_claim_number, "_", "")
	' 								panel_claim_number = replace(panel_claim_number, " ", "")
	' 								If panel_type_code = UNEA_type_code Then type_code_found = True
	' 								' MsgBox "panel_type_code - " & panel_type_code & vbCr & "UNEA_type_code - " & UNEA_type_code & vbCr & "type_code_found - " & type_code_found
	' 								If panel_type_code = UNEA_type_code and panel_claim_number = UNEA_claim_number Then list_of_panels_that_match = list_of_panels_that_match & "0"&panel_instance & " "

	' 								unea_panel_array(panel_type_const, unea_panel_counter) = panel_type_code
	' 								unea_panel_array(panel_claim_const, unea_panel_counter) = panel_claim_number
	' 								unea_panel_array(panel_claim_left_9_const, unea_panel_counter) = left(panel_claim_number, 9)
	' 								unea_panel_array(panel_instance_const, unea_panel_counter) = "0" & panel_instance

	' 								unea_panel_counter = unea_panel_counter + 1

	' 								transmit
	' 								EMReadScreen end_of_UNEA_panels, 7, 24, 2
	' 							Loop Until end_of_UNEA_panels = "ENTER A"
	' 						End If

	' 						' MsgBox "All panels that have the same detail: " & list_of_panels_that_match
	' 						list_of_panels_that_match = trim(list_of_panels_that_match)
	' 						If len(list_of_panels_that_match) > 2 Then
	' 							matching_panels_array = split(list_of_panels_that_match)

	' 							for the_matched_panel = 0 to UBound(matching_panels_array)-1
	' 								EMWriteScreen MEMB_reference_number, 20, 76 		'Navigating to STAT/UNEA
	' 								EMWriteScreen matching_panels_array(the_matched_panel), 20, 79 		'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	' 								transmit

	' 								EMWriteScreen "DEL", 20, 71
	' 								PF9
	' 								transmit

	' 								EMWaitReady 0, 0
	' 								transmit
	' 							next
	' 							' MsgBox "Did we delete?"
	' 						End If

	' 					ElseIf isDate(MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb)) = True Then	'If SSI has an end date listed
	' 						MEMB_reference_number = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
	' 						UNEA_type_code = "03"
	' 						UNEA_claim_number = MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb)&MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb)

	' 						ReDim unea_panel_array(last_panel_const, 0)
	' 						unea_panel_counter = 0

	' 						EMWriteScreen "UNEA", 20, 71
	' 						transmit
	' 						EMReadScreen unea_check, 4, 2, 48
	' 						Do While unea_check <> "UNEA"
	' 							Call navigate_to_MAXIS_screen("STAT", "UNEA")
	' 							EMReadScreen unea_check, 4, 2, 48
	' 						Loop
	' 						EMWriteScreen MEMB_reference_number, 20, 76 		'Navigating to STAT/UNEA
	' 						EMWriteScreen "01", 20, 79 		'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	' 						transmit

	' 						list_of_panels_that_match = " "
	' 						EMReadScreen vers_count, 1, 2, 78
	' 						If vers_count <> "0" Then
	' 							Do
	' 								ReDim Preserve unea_panel_array(last_panel_const, unea_panel_counter)
	' 								EMReadScreen panel_instance, 1, 2, 73
	' 								EMReadScreen panel_type_code, 2, 5, 37
	' 								EMReadScreen panel_claim_number, 15, 6, 37
	' 								panel_claim_number = replace(panel_claim_number, "_", "")
	' 								panel_claim_number = replace(panel_claim_number, " ", "")
	' 								If panel_type_code = UNEA_type_code Then type_code_found = True
	' 								' MsgBox "panel_type_code - " & panel_type_code & vbCr & "UNEA_type_code - " & UNEA_type_code & vbCr & "type_code_found - " & type_code_found
	' 								If panel_type_code = UNEA_type_code and panel_claim_number = UNEA_claim_number Then list_of_panels_that_match = list_of_panels_that_match & "0"&panel_instance & " "

	' 								unea_panel_array(panel_type_const, unea_panel_counter) = panel_type_code
	' 								unea_panel_array(panel_claim_const, unea_panel_counter) = panel_claim_number
	' 								unea_panel_array(panel_claim_left_9_const, unea_panel_counter) = left(panel_claim_number, 9)
	' 								unea_panel_array(panel_instance_const, unea_panel_counter) = "0" & panel_instance

	' 								unea_panel_counter = unea_panel_counter + 1

	' 								transmit
	' 								EMReadScreen end_of_UNEA_panels, 7, 24, 2
	' 							Loop Until end_of_UNEA_panels = "ENTER A"
	' 						End If

	' 						' MsgBox "All panels that have the same detail: " & list_of_panels_that_match
	' 						list_of_panels_that_match = trim(list_of_panels_that_match)
	' 						If len(list_of_panels_that_match) > 2 Then
	' 							matching_panels_array = split(list_of_panels_that_match)

	' 							for the_matched_panel = 0 to UBound(matching_panels_array)-1
	' 								EMWriteScreen MEMB_reference_number, 20, 76 		'Navigating to STAT/UNEA
	' 								EMWriteScreen matching_panels_array(the_matched_panel), 20, 79 		'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	' 								transmit

	' 								EMWriteScreen "DEL", 20, 71
	' 								PF9
	' 								transmit

	' 								EMWaitReady 0, 0
	' 								transmit
	' 							next
	' 							' MsgBox "Did we delete?"
	' 						End If

	' 					End If
	' 				End If

	' 				If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then				'Member with RSDI
	' 					'TODO - this functionality might need revision - not sure if the amount matching is the way to go
	' 					If MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) <> MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) Then
	' 						rsdi_type = "02"
	' 						If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then rsdi_type = "01"

	' 						MEMB_reference_number = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
	' 						UNEA_type_code = rsdi_type
	' 						UNEA_claim_number = MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb)

	' 						ReDim unea_panel_array(last_panel_const, 0)
	' 						unea_panel_counter = 0

	' 						EMWriteScreen "UNEA", 20, 71
	' 						transmit
	' 						EMReadScreen unea_check, 4, 2, 48
	' 						Do While unea_check <> "UNEA"
	' 							Call navigate_to_MAXIS_screen("STAT", "UNEA")
	' 							EMReadScreen unea_check, 4, 2, 48
	' 						Loop
	' 						EMWriteScreen MEMB_reference_number, 20, 76 		'Navigating to STAT/UNEA
	' 						EMWriteScreen "01", 20, 79 		'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	' 						transmit

	' 						list_of_panels_that_match = " "
	' 						EMReadScreen vers_count, 1, 2, 78
	' 						If vers_count <> "0" Then
	' 							Do
	' 								ReDim Preserve unea_panel_array(last_panel_const, unea_panel_counter)
	' 								EMReadScreen panel_instance, 1, 2, 73
	' 								EMReadScreen panel_type_code, 2, 5, 37
	' 								EMReadScreen panel_claim_number, 15, 6, 37
	' 								panel_claim_number = replace(panel_claim_number, "_", "")
	' 								panel_claim_number = replace(panel_claim_number, " ", "")
	' 								If panel_type_code = UNEA_type_code Then type_code_found = True
	' 								' MsgBox "panel_type_code - " & panel_type_code & vbCr & "UNEA_type_code - " & UNEA_type_code & vbCr & "type_code_found - " & type_code_found
	' 								If panel_type_code = UNEA_type_code and panel_claim_number = UNEA_claim_number Then list_of_panels_that_match = list_of_panels_that_match & "0"&panel_instance & " "

	' 								unea_panel_array(panel_type_const, unea_panel_counter) = panel_type_code
	' 								unea_panel_array(panel_claim_const, unea_panel_counter) = panel_claim_number
	' 								unea_panel_array(panel_claim_left_9_const, unea_panel_counter) = left(panel_claim_number, 9)
	' 								unea_panel_array(panel_instance_const, unea_panel_counter) = "0" & panel_instance

	' 								unea_panel_counter = unea_panel_counter + 1

	' 								transmit
	' 								EMReadScreen end_of_UNEA_panels, 7, 24, 2
	' 							Loop Until end_of_UNEA_panels = "ENTER A"
	' 						End If

	' 						list_of_panels_that_match = trim(list_of_panels_that_match)
	' 						If len(list_of_panels_that_match) > 2 Then
	' 							' MsgBox "RSDI - All panels that have the same detail: " & list_of_panels_that_match
	' 							matching_panels_array = split(list_of_panels_that_match)

	' 							for the_matched_panel = 0 to UBound(matching_panels_array)-1
	' 								EMWriteScreen MEMB_reference_number, 20, 76 		'Navigating to STAT/UNEA
	' 								EMWriteScreen matching_panels_array(the_matched_panel), 20, 79 		'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	' 								transmit

	' 								EMWriteScreen "DEL", 20, 71
	' 								PF9
	' 								transmit

	' 								EMWaitReady 0, 0
	' 								transmit
	' 							next
	' 							' MsgBox "Did we delete?"
	' 						End If

	' 					End If
	' 				End If
	' 			End If
	' 		Next
	' 	End If
	' 	objRecordSet.MoveNext			'now we go to the next case
	' Loop
    ' objRecordSet.Close			'Closing all the data connections
    ' objConnection.Close
    ' Set objRecordSet=nothing
    ' Set objConnection=nothing

	'-------------------------------------------------------

	' 'THIS IS FOR CREATING A LIST OF CASES APPROVED FOR THE REVIEW MONTH THAT HAVE MSA and/or GRH
	' review_date = "8/1/2023"			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	' review_date = DateAdd("d", 0, review_date)

	' 'Opening a spreadsheet to capture the cases with a SMRT ending soon
	' Set ObjExcel = CreateObject("Excel.Application")
	' ObjExcel.Visible = True
	' Set objSMRTWorkbook = ObjExcel.Workbooks.Add()
	' ObjExcel.DisplayAlerts = True

	' 'Setting the first 4 col as worker, case number, name, and APPL date
	' ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
	' ObjExcel.Cells(1, 2).Value = "APPROVAL WORKER"
	' ObjExcel.Cells(1, 3).Value = "MSA Status"
	' ObjExcel.Cells(1, 4).Value = "GRH Status"

	' FOR i = 1 to 8		'formatting the cells'
	' 	ObjExcel.Cells(1, i).Font.Bold = True		'bold font'
	' NEXT

	' excel_row = 2		'initializing the counter to move through the excel lines

	' objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

	' Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	' Set objRecordSet = CreateObject("ADODB.Recordset")

	' 'opening the connections and data table
	' objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	' objRecordSet.Open objSQL, objConnection

	' Do While NOT objRecordSet.Eof 					'Loop through each item on the CASE LIST Table
	' 	If objRecordSet("ExParteAfterPhase2") = "Approved as Ex Parte" Then
	' 		MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER
	' 		Call back_to_SELF

	' 		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

	' 		If msa_case = True or grh_case = True Then
	' 			ObjExcel.Cells(excel_row, 1).Value = MAXIS_case_number
	' 			ObjExcel.Cells(excel_row, 2).Value = objRecordSet("Phase2HSR")
	' 			ObjExcel.Cells(excel_row, 3).Value = msa_status
	' 			ObjExcel.Cells(excel_row, 4).Value = grh_status
	' 			excel_row = excel_row + 1
	' 		End If
	' 	End If
	' 	objRecordSet.MoveNext			'now we go to the next case
	' Loop
	' objRecordSet.Close			'Closing all the data connections
	' objConnection.Close
	' Set objRecordSet=nothing
	' Set objConnection=nothing

	' For col_to_autofit = 1 to 4
	' 	ObjExcel.columns(col_to_autofit).AutoFit()
	' Next
	'----------------------------------------------------------------

	'This area is here in the event that we need to create an update process to the Ex Parte data list on a large number of cases.
	'This will need to be defined on a case-by-case scenario.

	Call script_end_procedure("Fix completed." & fix_report_out)

End If

'This is the first functionality to run after the data list is created. It will likely be run some time between the 6th and the 15th.
'This will evaluate the cases for Ex Parte, send the initial SVES QURY and create the other verification lists
If ex_parte_function = "Prep 1" Then
	review_date = ep_revw_mo & "/1/" & ep_revw_yr			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	review_date = DateAdd("d", 0, review_date)
	smrt_cut_off = DateAdd("m", 1, review_date)				'This is the cutoff date for SMRT ending to identify which ones we want to have evaluated

	'This functionality will remove any 'holds' with 'In Progress' marked. This is to make sure no cases get left behind if a script fails
	If reset_in_Progress = checked Then
		'This is opening the Ex Parte Case List data table so we can loop through it.
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

		Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		'Loop through each item on the CASE LIST Table
		Do While NOT objRecordSet.Eof
			If objRecordSet("PREP_Complete") = "In Progress" Then			'If the case is marked as 'In Progress' - we are going to remove it
				MAXIS_case_number = objRecordSet("CaseNumber") 				'SET THE MAXIS CASE NUMBER

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = '" & NULL & "'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"	'removing the 'In Progress' indicator and blanking it out

				Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			End If
			objRecordSet.MoveNext			'move to the next item in the table
		Loop
		objRecordSet.Close			'Closing all the data connections
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
	End If

	va_count = 0		'initializing these counter variables at 0
	uc_count = 0
	rr_count = 0

	Set ObjFSO = CreateObject("Scripting.FileSystemObject")

	'If the file exists we open it and set to add to it
	If ObjFSO.FileExists(ex_parte_folder & "\VA Income Verifications\VA Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx") Then
		Call excel_open(ex_parte_folder & "\VA Income Verifications\VA Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx", True, False, objVAExcel, objVAWorkbook)
		va_excel_row = 2
		va_inc_count = 0
		Do
			listed_case_numb = trim(objVAExcel.Cells(va_excel_row, 1).value)
			If listed_case_numb <> "" Then
				va_excel_row = va_excel_row + 1
				va_inc_count = va_inc_count + 1
			End If
		Loop until listed_case_numb = ""
	Else												'If the file does not exists, we create it and set to writing the file
		'set the Excel sheet up for VA
		Set objVAExcel = CreateObject("Excel.Application")				'opening a new Excel sheet
		objVAExcel.Visible = True
		Set objVAWorkbook = objVAExcel.Workbooks.Add()
		objVAExcel.DisplayAlerts = True

		objVAExcel.Cells(1, 1).Value = "CASE NUMBER"					'Putting the headers in place for the Excel sheet
		objVAExcel.Cells(1, 2).Value = "REF"
		objVAExcel.Cells(1, 3).Value = "NAME"
		objVAExcel.Cells(1, 4).Value = "PMI NUMBER"
		objVAExcel.Cells(1, 5).Value = "SSN"
		objVAExcel.Cells(1, 6).Value = "VA INC TYPE"
		objVAExcel.Cells(1, 7).Value = "VA CLAIM NUMB"
		objVAExcel.Cells(1, 8).Value = "CURR VA INCOME"
		objVAExcel.Cells(1, 9).Value = "Verified VA Income"
		objVAExcel.columns(2).NumberFormat = "@" 		'formatting as text

		FOR i = 1 to 9		'formatting the cells'
			objVAExcel.Cells(1, i).Font.Bold = True		'bold font'
		NEXT
		va_excel_row = 2
		va_inc_count = 0

		objVAExcel.ActiveSheet.ListObjects.Add(xlSrcRange, objVAExcel.Range("A1:I" & va_excel_row - 1), xlYes).Name = "Table1"
		objVAExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
		objVAExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\VA Income Verifications\VA Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	End If

	If ObjFSO.FileExists(ex_parte_folder & "\UC Income Verifications\UC Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx") Then
		Call excel_open(ex_parte_folder & "\UC Income Verifications\UC Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx", True, False, objUCExcel, objUCWorkbook)
		uc_excel_row = 2
		uc_inc_count = 0
		Do
			listed_case_numb = trim(objUCExcel.Cells(uc_excel_row, 1).value)
			If listed_case_numb <> "" Then
				uc_excel_row = uc_excel_row + 1
				uc_inc_count = uc_inc_count + 1
			End If
		Loop until listed_case_numb = ""
	Else
		'set the Excel sheet up for UC
		Set objUCExcel = CreateObject("Excel.Application")				'opening a new Excel sheet
		objUCExcel.Visible = True
		Set objUCWorkbook = objUCExcel.Workbooks.Add()
		objUCExcel.DisplayAlerts = True

		objUCExcel.Cells(1, 1).Value = "CASE NUMBER"					'Putting the headers in place for the Excel sheet
		objUCExcel.Cells(1, 2).Value = "REF"
		objUCExcel.Cells(1, 3).Value = "NAME"
		objUCExcel.Cells(1, 4).Value = "PMI NUMBER"
		objUCExcel.Cells(1, 5).Value = "SSN"
		objUCExcel.Cells(1, 6).Value = "UC INC TYPE"
		objUCExcel.Cells(1, 7).Value = "UC CLAIM NUMB"
		objUCExcel.Cells(1, 8).Value = "CURR UC INCOME"
		objUCExcel.Cells(1, 9).Value = "Verified UC Income"
		objUCExcel.columns(2).NumberFormat = "@" 		'formatting as text

		FOR i = 1 to 9		'formatting the cells'
			objUCExcel.Cells(1, i).Font.Bold = True		'bold font'
		NEXT
		uc_excel_row = 2
		uc_inc_count = 0

		objUCExcel.ActiveSheet.ListObjects.Add(xlSrcRange, objUCExcel.Range("A1:I" & uc_excel_row - 1), xlYes).Name = "Table1"
		objUCExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
		objUCExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\UC Income Verifications\UC Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	End If

	If ObjFSO.FileExists(ex_parte_folder & "\RR Income Verifications\RR Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx") Then
		Call excel_open(ex_parte_folder & "\RR Income Verifications\RR Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx", True, False, objRRExcel, objRRWorkbook)
		rr_excel_row = 2
		rr_inc_count = 0
		Do
			listed_case_numb = trim(objRRExcel.Cells(rr_excel_row, 1).value)
			If listed_case_numb <> "" Then
				rr_excel_row = rr_excel_row + 1
				rr_inc_count = rr_inc_count + 1
			End If
		Loop until listed_case_numb = ""
	Else
		'set the Excel sheet up for RR
		Set objRRExcel = CreateObject("Excel.Application")				'opening a new Excel sheet
		objRRExcel.Visible = True
		Set objRRWorkbook = objRRExcel.Workbooks.Add()
		objRRExcel.DisplayAlerts = True

		objRRExcel.Cells(1, 1).Value = "CASE NUMBER"					'Putting the headers in place for the Excel sheet
		objRRExcel.Cells(1, 2).Value = "REF"
		objRRExcel.Cells(1, 3).Value = "NAME"
		objRRExcel.Cells(1, 4).Value = "PMI NUMBER"
		objRRExcel.Cells(1, 5).Value = "SSN"
		objRRExcel.Cells(1, 6).Value = "RR INC TYPE"
		objRRExcel.Cells(1, 7).Value = "RR CLAIM NUMB"
		objRRExcel.Cells(1, 8).Value = "CURR RR INCOME"
		objRRExcel.Cells(1, 9).Value = "Verified RR Income"
		objRRExcel.columns(2).NumberFormat = "@" 		'formatting as text

		FOR i = 1 to 9		'formatting the cells'
			objRRExcel.Cells(1, i).Font.Bold = True		'bold font'
		NEXT
		rr_excel_row = 2
		rr_inc_count = 0

		objRRExcel.ActiveSheet.ListObjects.Add(xlSrcRange, objRRExcel.Range("A1:I" & rr_excel_row - 1), xlYes).Name = "Table1"
		objRRExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
		objRRExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\RR Income Verifications\RR Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	End If


	'This is opening the Ex Parte Case List data table so we can loop through it.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	Do While NOT objRecordSet.Eof 					'Loop through each item on the CASE LIST Table
		'Pulling any case where the PREP_complete is null or blank
		If IsNull(objRecordSet("PREP_Complete")) = True or objRecordSet("PREP_Complete") = "" Then
			all_hc_is_ABD = ""				'resetting all these variables to blank at the beginning of each loop so information doesn't carry over from one case to another
			SSA_income_exists = ""
			JOBS_income_exists = ""
			VA_income_exists = ""
			BUSI_income_exists = ""
			case_has_no_income = ""
			case_has_EPD = ""

			appears_ex_parte = True			'we default this to true and find reasons that exclude the case from Ex Parte as we look at case data.
			all_hc_is_ABD = True
			case_has_EPD = False
			case_is_in_henn = False
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			'Here we are setting the PREP_Complete to 'In Progress' to hold the case as being worked.
			'This portion of the script is required to be able to have more than one person operating the BULK run at the same time.
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = 'In Progress'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)			'This is defined here without a preserve to blank it out at the beginning of every loop with a new case
			memb_count = 0										'resetting the counting variable to size the member array
			list_of_membs_on_hc = " "							'we need to keep a list members by pmi to know if a person is already accounted for as we find all the members and programs

			'We need to pull all of the instances from the ELIG table for the currently defined case number
			'This will list the HH member and eligibility program for HC. We will use this to start to determine if the case can be processed as Ex Parte
			objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & MAXIS_case_number & "'"

			Set objELIGConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objELIGRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objELIGRecordSet.Open objELIGSQL, objELIGConnection
			' MsgBox "We are at 1"

			person_found = False		'setting the default of if we have found a person in the list
			Do While NOT objELIGRecordSet.Eof
				list_of_membs_on_hc = list_of_membs_on_hc & objELIGRecordSet("PMINumber") & " "		'adding the PMI to the list of all PMIs known on the case
				person_found = True																	'indicating that there was a person in the list for this case
				memb_known = False																	'sets that we don't know if we have already looked at this person
				'now we loop through all of the people we have already found for this case - we only want 1 array instance per person.
				For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					If trim(objELIGRecordSet("PMINumber")) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then		'If the PMI matches one in the array, we are going to set the information to that array instance
						memb_known = True															'identifies that we know about this person and they are already in the array

						'figuring out which program type location the information should be saved in for this table data
						'each person on a case may have up to three different lines for different programs
						If MEMBER_INFO_ARRAY(table_prog_1, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_1, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_1, known_membs) 		= objELIGRecordSet("EligType")
						ElseIf MEMBER_INFO_ARRAY(table_prog_2, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_2, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_2, known_membs) 		= objELIGRecordSet("EligType")
						ElseIf MEMBER_INFO_ARRAY(table_prog_3, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_3, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_3, known_membs) 		= objELIGRecordSet("EligType")
						End If
						'This will read the ELIG type information and use it to identify if a case does NOT appear Ex Parte
						If objELIGRecordSet("MajorProgram") = "EH" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "AX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "AA" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "DP" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CK" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CB" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CM" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "13" Then appears_ex_parte = False 	'TYMA
						If objELIGRecordSet("EligType") = "14" Then appears_ex_parte = False 	'TYMA
						If objELIGRecordSet("EligType") = "09" Then appears_ex_parte = False 	'Adoption Assistance
						If objELIGRecordSet("EligType") = "11" Then appears_ex_parte = False 	'Auto Newborn
						If objELIGRecordSet("EligType") = "10" Then appears_ex_parte = False 	'Adoption Assistance
						If objELIGRecordSet("EligType") = "25" Then appears_ex_parte = False 	'Foster Care
						If objELIGRecordSet("EligType") = "PX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "PC" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "BC" Then appears_ex_parte = False

						If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False		'identifying if the case has ABD basis or not
						If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True										'identifying if the case has MA-EPD
					End If
				Next

				'If this is an unknown member, and has not been added to the array already, we need to add it
				If memb_known = False Then
					ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)								'resizing the array

					'setting personal information to the array
					MEMBER_INFO_ARRAY(memb_pmi_numb_const, memb_count) 	= trim(objELIGRecordSet("PMINumber"))
					MEMBER_INFO_ARRAY(memb_ssn_const, memb_count) 		= trim(objELIGRecordSet("SocialSecurityNbr"))
					name_var									 		= trim(objELIGRecordSet("Name"))		'we want to format the name corectly.
					name_array = split(name_var)
					MEMBER_INFO_ARRAY(memb_name_const, memb_count) = name_array(UBound(name_array))
					For name_item = 0 to UBound(name_array)-1
						MEMBER_INFO_ARRAY(memb_name_const, memb_count) = MEMBER_INFO_ARRAY(memb_name_const, memb_count) & " " & name_array(name_item)
					Next
					MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
					MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(objELIGRecordSet("MajorProgram"))	'setting the program information
					MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(objELIGRecordSet("EligType"))

					'This will read the ELIG type information and use it to identify if a case does NOT appear Ex Parte
					If objELIGRecordSet("MajorProgram") = "EH" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "AX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "AA" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "DP" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CK" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CB" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CM" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "13" Then appears_ex_parte = False 	'TYMA
					If objELIGRecordSet("EligType") = "14" Then appears_ex_parte = False 	'TYMA
					If objELIGRecordSet("EligType") = "09" Then appears_ex_parte = False 	'Adoption Assistance
					If objELIGRecordSet("EligType") = "11" Then appears_ex_parte = False 	'Auto Newborn
					If objELIGRecordSet("EligType") = "10" Then appears_ex_parte = False 	'Adoption Assistance
					If objELIGRecordSet("EligType") = "25" Then appears_ex_parte = False 	'Foster Care
					If objELIGRecordSet("EligType") = "PX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "PC" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "BC" Then appears_ex_parte = False

					If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False		'identifying if the case has ABD basis or not
					If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True										'identifying if the case has MA-EPD

					MEMBER_INFO_ARRAY(sql_rr_income_exists, memb_count) = False		'defaulting the income types for this case to false
					MEMBER_INFO_ARRAY(sql_va_income_exists, memb_count) = False
					MEMBER_INFO_ARRAY(sql_uc_income_exists, memb_count) = False

					memb_count = memb_count + 1		'incrementing the array counter up for the next loop
				End if
				objELIGRecordSet.MoveNext			'going to the next record
			Loop
			objELIGRecordSet.Close			'Closing all the data connections
			objELIGConnection.Close
			Set objELIGRecordSet=nothing
			Set objELIGConnection=nothing
			' MsgBox "We are at 2"
			'If the ELIG types still indicate that the case is Ex Parte, we are going to check REVW to make sure the case meets renewal requirements
			If appears_ex_parte = True Then
				'check HC ER date in STAT/REVW
				' MsgBox "We are at 3"
				Call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)
				If is_this_priv = True Then appears_ex_parte = False						'excluding cases that are privileged
				If is_this_priv = False Then
					Call write_value_and_transmit("X", 5, 71)
					EMReadScreen STAT_HC_ER_mo, 2, 8, 27
					EMReadScreen STAT_HC_ER_yr, 2, 8, 33
					If ep_revw_mo <> STAT_HC_ER_mo or ep_revw_yr <> STAT_HC_ER_yr Then  appears_ex_parte = False		'if this does not have the correct renewal month, we will exclude it from Ex Parte
				End If
			End If

			' MsgBox "We are at 4"
			'If the case still appears Ex Parte, we are going to check if we are missing people, and check income for further determination of Ex Parte
			If appears_ex_parte = True Then
				'If we did not find people in the ELIG list, we are going to check ELIG/HC
				' MsgBox "We are at 5"
				If person_found = False Then
					Call navigate_to_MAXIS_screen("STAT", "SUMM")		'Creating new ELIG results
					'Send the case through background
					Call write_value_and_transmit("BGTX", 20, 71)					'Enter the command to force the case through background
					EMReadScreen wrap_check, 4, 2, 46								'Making sure we are at STAT/WRAP
					If wrap_check = "WRAP" Then transmit							'If we are at WRAP, transmit through
					EMWaitReady 0, 0												'give a pause here
					EMReadScreen database_busy, 23, 4, 44							'Sometimes, when we send a case through background a database record error raises
					If database_busy = "A MAXIS database record" Then transmit  	'we need to transmit past it
					'TODO - there may be a NAT error being raised here, but I don't know what that might be from or if we need to resolve it - there does not seem to be any impact to running the script
					Call back_to_SELF												'Need to get to SELF

					Call MAXIS_background_check

					Call navigate_to_MAXIS_screen("ELIG", "HC  ")		'Navigate to ELIG/HC
					'Here we start at the top of ELIG/HC and read each row to find HC information
					hc_row = 8
					Do
						pers_type = ""		'blanking out variables so they don't carry over from loop to loop
						std = ""
						meth = ""
						waiv = ""

						'reading the main HC Elig information - member, program, status
						EMReadScreen read_ref_numb, 2, hc_row, 3
						EMReadScreen clt_hc_prog, 4, hc_row, 28
						EMReadScreen hc_prog_status, 6, hc_row, 50
						ref_row = hc_row
						Do while read_ref_numb = "  "				'this will read for the reference number if there are multiple programs for a single member
							ref_row = ref_row - 1
							EMReadScreen read_ref_numb, 2, ref_row, 3
						Loop

						If hc_prog_status = "ACTIVE" Then			'If HC is currently active, we need to read more details about the program/eligibility
							clt_hc_prog = trim(clt_hc_prog)			'formatting this to remove whitespace
							If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "" Then		'these are non-hc persons

								Call write_value_and_transmit("X", hc_row, 26)									'opening the ELIG detail spans
								If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then		'If it is an MSP, we want to read the type only from a specific place
									elig_msp_prog = clt_hc_prog
									EMReadScreen pers_type, 2, 6, 56
								Else																			'These are MA type programs (not MSP)
									'Now we have to fund the current month in elig to get the current elig type
									col = 19
									Do
										EMReadScreen span_month, 2, 6, col										'reading the month in ELIG
										EMReadScreen span_year, 2, 6, col+3

										'if the span month matchest current month plus 1, we are going to grab elig from that month
										If span_month = MAXIS_footer_month and span_year = MAXIS_footer_year Then
											EMReadScreen pers_type, 2, 12, col - 2								'reading the ELIG TYPE
											EMReadScreen std, 1, 12, col + 3
											EMReadScreen meth, 1, 13, col + 2
											EMReadScreen waiv, 1, 17, col + 2
											Exit Do																'leaving once we've found the information for this elig
										End If
										col = col + 11			'this goes to the next column
									Loop until col = 85			'This is off the page - if we hit this, we did NOT find the elig type in this elig result

									'If we hit 85, we did not get the information. So we are going to read it from the last budget month (most current)
									If col = 85 Then
										EMReadScreen pers_type, 2, 12, 72										'reading the ELIG TYPE
										EMReadScreen std, 1, 12, 77
										EMReadScreen meth, 1, 13, 76
										EMReadScreen waiv, 1, 17, 76
									End If
								End If
								PF3			'leaving the elig detail information

								'now we need to add the information we just read to the member array
								memb_known = False										'default that the member know is false
								For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)								'Looking at all the members known in the array
									If MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs) = read_ref_numb Then	'if the member reference from ELIG matches the ARRAY, we are going to add more elig details
										memb_known = True														'look we found a person
										If MEMBER_INFO_ARRAY(table_prog_1, known_membs) = "" Then				'finding which area of the array is blank to save the elig information there
											MEMBER_INFO_ARRAY(table_prog_1, known_membs) 		= clt_hc_prog
											MEMBER_INFO_ARRAY(table_type_1, known_membs) 		= pers_type
										ElseIf MEMBER_INFO_ARRAY(table_prog_2, known_membs) = "" Then
											MEMBER_INFO_ARRAY(table_prog_2, known_membs) 		= clt_hc_prog
											MEMBER_INFO_ARRAY(table_type_2, known_membs) 		= pers_type
										ElseIf MEMBER_INFO_ARRAY(table_prog_3, known_membs) = "" Then
											MEMBER_INFO_ARRAY(table_prog_3, known_membs) 		= clt_hc_prog
											MEMBER_INFO_ARRAY(table_type_3, known_membs) 		= pers_type
										End If

										'This will read the ELIG type information and use it to identify if a case does NOT appear Ex Parte
										If clt_hc_prog = "EH" Then appears_ex_parte = False
										If pers_type = "AX" Then appears_ex_parte = False
										If pers_type = "AA" Then appears_ex_parte = False
										If pers_type = "DP" Then appears_ex_parte = False
										If pers_type = "CK" Then appears_ex_parte = False
										If pers_type = "CX" Then appears_ex_parte = False
										If pers_type = "CB" Then appears_ex_parte = False
										If pers_type = "CM" Then appears_ex_parte = False
										If pers_type = "13" Then appears_ex_parte = False 	'TYMA
										If pers_type = "14" Then appears_ex_parte = False 	'TYMA
										If pers_type = "09" Then appears_ex_parte = False 	'Adoption Assistance
										If pers_type = "11" Then appears_ex_parte = False 	'Auto Newborn
										If pers_type = "10" Then appears_ex_parte = False 	'Adoption Assistance
										If pers_type = "25" Then appears_ex_parte = False 	'Foster Care
										If pers_type = "PX" Then appears_ex_parte = False
										If pers_type = "PC" Then appears_ex_parte = False
										If pers_type = "BC" Then appears_ex_parte = False

										If appears_ex_parte = False AND pers_type <> "DP" Then all_hc_is_ABD = False		'identifying if the case has ABD basis or not
										If pers_type = "DP" Then case_has_EPD = True										'identifying if the case has MA-EPD
									End If
								Next

								'If this is an unknown member, and has not been added to the array already, we need to add it
								If memb_known = False Then
									ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)								'resizing the array

									'setting personal information to the array
									MEMBER_INFO_ARRAY(memb_ref_numb_const, memb_count) = read_ref_numb
									MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
									MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(clt_hc_prog)
									MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(pers_type)

									'This will read the ELIG type information and use it to identify if a case does NOT appear Ex Parte
									If clt_hc_prog = "EH" Then appears_ex_parte = False
									If pers_type = "AX" Then appears_ex_parte = False
									If pers_type = "AA" Then appears_ex_parte = False
									If pers_type = "DP" Then appears_ex_parte = False
									If pers_type = "CK" Then appears_ex_parte = False
									If pers_type = "CX" Then appears_ex_parte = False
									If pers_type = "CB" Then appears_ex_parte = False
									If pers_type = "CM" Then appears_ex_parte = False
									If pers_type = "13" Then appears_ex_parte = False 	'TYMA
									If pers_type = "14" Then appears_ex_parte = False 	'TYMA
									If pers_type = "09" Then appears_ex_parte = False 	'Adoption Assistance
									If pers_type = "11" Then appears_ex_parte = False 	'Auto Newborn
									If pers_type = "10" Then appears_ex_parte = False 	'Adoption Assistance
									If pers_type = "25" Then appears_ex_parte = False 	'Foster Care
									If pers_type = "PX" Then appears_ex_parte = False
									If pers_type = "PC" Then appears_ex_parte = False
									If pers_type = "BC" Then appears_ex_parte = False

									If appears_ex_parte = False AND pers_type <> "DP" Then all_hc_is_ABD = False		'identifying if the case has ABD basis or not
									If pers_type = "DP" Then case_has_EPD = True										'identifying if the case has MA-EPD

									MEMBER_INFO_ARRAY(sql_rr_income_exists, memb_count) = False		'defaulting the income types for this case to false
									MEMBER_INFO_ARRAY(sql_va_income_exists, memb_count) = False
									MEMBER_INFO_ARRAY(sql_uc_income_exists, memb_count) = False

									memb_count = memb_count + 1 	'incrementing the array counter up for the next loop
								End If

							End If
						End If
						hc_row = hc_row + 1												'now we go to the next row
						EMReadScreen next_ref_numb, 2, hc_row, 3						'read the next HC information to find when we've reeached the end of the list
						EMReadScreen next_maj_prog, 4, hc_row, 28
					Loop until next_ref_numb = "  " and next_maj_prog = "    "

					CALL back_to_SELF()													'going to STAT/MEMB - because there is misssing personal information for the members discovered in this way
					Do
						CALL navigate_to_MAXIS_screen("STAT", "MEMB")
						EMReadScreen memb_check, 4, 2, 48
					Loop until memb_check = "MEMB"

					at_least_one_hc_active = False										'this is a default to identify if HC is active on the case
					For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)					'loop through the member array
						Call write_value_and_transmit(MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs), 20, 76)		'navigate to the member for this instance of the array
						EMReadscreen last_name, 25, 6, 30								'read and cormat the name from MEMB
						EMReadscreen first_name, 12, 6, 63
						last_name = trim(replace(last_name, "_", "")) & " "
						first_name = trim(replace(first_name, "_", "")) & " "
						MEMBER_INFO_ARRAY(memb_name_const, known_membs) = first_name & " " & last_name
						EMReadScreen PMI_numb, 8, 4, 46									'capturing the PMI number
						PMI_numb = trim(PMI_numb)
						MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) = right("00000000" & PMI_numb, 8)			'we have to format the pmi to match the data list format (8 digits with leading 0s included)
						EMReadScreen MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), 11, 7, 42							'catpturing the SSN
						MEMBER_INFO_ARRAY(memb_ssn_const, known_membs) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), " ", "")
						MEMBER_INFO_ARRAY(memb_ssn_const, known_membs) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), "_", "")
						If MEMBER_INFO_ARRAY(table_prog_1, known_membs) <> "" Then at_least_one_hc_active = True		'setting the variable that identifies there is HC active based on the ELIG read from HC/ELIG
						If MEMBER_INFO_ARRAY(table_prog_2, known_membs) <> "" Then at_least_one_hc_active = True
						If MEMBER_INFO_ARRAY(table_prog_3, known_membs) <> "" Then at_least_one_hc_active = True
						If MEMBER_INFO_ARRAY(table_prog_1, known_membs) <> "" or MEMBER_INFO_ARRAY(table_prog_2, known_membs) <> "" or MEMBER_INFO_ARRAY(table_prog_3, known_membs) <> "" Then
							list_of_membs_on_hc = list_of_membs_on_hc & MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) & " "		'adding individuals to our list of members on HC
						End If

					Next
					If at_least_one_hc_active = False Then appears_ex_parte = False			'if no one is on HC, this cannot be Ex Parte
				End If
			End If

			' MsgBox "We are at 6"
			If is_this_priv = False and appears_ex_parte = True Then
				Call navigate_to_MAXIS_screen("STAT", "MEMB")		'now we go find all the HH members
				Call get_list_of_members
			End if

			'Now we are going to start looking at income information to remove any cases that have income thant disqualifies it from Ex parte
			SSA_income_exists = False				'setting these variables to false at the beginning of each loop through
			RR_income_exists = False
			VA_income_exists = False
			UC_income_exists = False
			PRISM_income_exists = False
			Other_UNEA_income_exists = False
			JOBS_income_exists = False
			BUSI_income_exists = False

			'Pulling all rows from the INCOME list for the case number we are currently processing
			objIncomeSQL = "SELECT * FROM ES.ES_ExParte_IncomeList WHERE [CaseNumber] = '" & MAXIS_case_number & "'"

			Set objIncomeConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

			'looping through each row in this case
			Do While NOT objIncomeRecordSet.Eof
				income_for_person_is_on_HC = False			'default this variable to false, indicating if the income is for a person on HC
				If InStr(list_of_membs_on_hc, objIncomeRecordSet("PersonID")) <> 0 Then income_for_person_is_on_HC = True		'this compares the PMI for the income to the list of PMIS discovered in finding HC elig information

				'If this income is for someone on HC, we are going to assess the income detail to determine if the case should still be Ex Parte
				If income_for_person_is_on_HC = True Then
					If objIncomeRecordSet("IncExpTypeCode") = "UNEA" Then									'UNEA income exists each type code will set the boolean about that income typr for this case
						If objIncomeRecordSet("IncomeTypeCode") = "01" Then SSA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "02" Then SSA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "03" Then SSA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "16" Then RR_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "11" Then VA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "12" Then VA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "13" Then VA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "38" Then VA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "14" Then UC_income_exists = True

						If objIncomeRecordSet("IncomeTypeCode") = "36" Then PRISM_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "37" Then PRISM_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "39" Then PRISM_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "40" Then PRISM_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "36" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "37" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "39" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "40" Then Other_UNEA_income_exists = True

						If objIncomeRecordSet("IncomeTypeCode") = "06" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "15" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "17" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "18" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "23" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "24" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "25" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "26" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "27" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "28" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "29" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "08" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "35" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "43" Then Other_UNEA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "47" Then Other_UNEA_income_exists = True
					End If
					If objIncomeRecordSet("IncExpTypeCode") = "JOBS" Then JOBS_income_exists = True					'we do not need to clarify further for JOBS or BUSI income, just indicate if these incomes exist.
					If objIncomeRecordSet("IncExpTypeCode") = "BUSI" Then BUSI_income_exists = True
				End If

				'Here we set if there is certain types of income on the case for any member. This will information the creation of verification lists
				For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)			'loop through all the members
					If trim(objIncomeRecordSet("PersonID")) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then						'if the PMI matches
						If objIncomeRecordSet("IncomeTypeCode") = "16" Then MEMBER_INFO_ARRAY(sql_rr_income_exists, known_membs) = True		'if the income type is any of the specified, identify that the income exists
						If objIncomeRecordSet("IncomeTypeCode") = "11" Then MEMBER_INFO_ARRAY(sql_va_income_exists, known_membs) = True
						If objIncomeRecordSet("IncomeTypeCode") = "12" Then MEMBER_INFO_ARRAY(sql_va_income_exists, known_membs) = True
						If objIncomeRecordSet("IncomeTypeCode") = "13" Then MEMBER_INFO_ARRAY(sql_va_income_exists, known_membs) = True
						If objIncomeRecordSet("IncomeTypeCode") = "38" Then MEMBER_INFO_ARRAY(sql_va_income_exists, known_membs) = True
						If objIncomeRecordSet("IncomeTypeCode") = "14" Then MEMBER_INFO_ARRAY(sql_uc_income_exists, known_membs) = True
					End If
				Next

				objIncomeRecordSet.MoveNext		'move to the next Income row
			Loop
			objIncomeRecordSet.Close			'Closing all the data connections
			objIncomeConnection.Close
			Set objIncomeRecordSet=nothing
			Set objIncomeConnection=nothing

			'This part is for logic to help us determine if the income impacts the Ex parte option
			case_has_no_income = False			'start at false for 'no income' basically false here means the case has income
			'If every income type is false, then the case has no income and the variable, 'case_has_no_income' is set to True, because it is true that there is no income.
			If SSA_income_exists = False and RR_income_exists = False and VA_income_exists = False and UC_income_exists = False and PRISM_income_exists = False and Other_UNEA_income_exists = False and JOBS_income_exists = False and BUSI_income_exists = False Then case_has_no_income = True

			'If the case apears Ex Parte at this point, we are going to do another assessment
			If appears_ex_parte = True Then
				'reading case program information and PW
				Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
				EMReadScreen case_pw, 7, 21, 14
				If left(case_pw, 4) = "X127" Then case_is_in_henn = True
				'This would exclude cases that are not in Hennepin or are Closed
				' If case_is_in_henn = False then  appears_ex_parte = False				'we are not going to exclude for inactive or out of county uuntil Phase 1 at this point
				' If case_active = False Then appears_ex_parte = False
				' If ma_status <> "ACTIVE" and msp_status <> "ACTIVE" Then appears_ex_parte = False

				'Any other UNEA, or JOBS/BUSI income requires the case be on SNAP or MFIP at this point
				If Other_UNEA_income_exists = True OR JOBS_income_exists = True OR BUSI_income_exists = True Then
					appears_ex_parte = False									'if there is JOBS/BUSI/Other UNEA - this cannot be ex parte
					If mfip_status = "ACTIVE" Then appears_ex_parte = True		'unless MFIP or SNAP is active
					If snap_status = "ACTIVE" Then appears_ex_parte = True
				End If
			End If

			'If the case still appears Ex Parte at this point, we need to start the verifications
			If appears_ex_parte = True Then
				'For each case that is indicated as potentially ExParte, we are going to take preperation actions
				last_va_count = va_count			'These are counting variables to set for each loop
				last_uc_count = uc_count
				last_rr_count = rr_count

				Call find_unea_information			'Now we are reading UNEA information for all the HH members

				Call back_to_SELF

				'Send a SVES/CURY for all persons on a case
				Call navigate_to_MAXIS_screen("INFC", "SVES")
				'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
				EMReadScreen agreement_check, 9, 2, 24
				IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

				'We need to loop through each HH Member on the case and send a QURY for every one.
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					Call send_sves_qury("SSN", qury_finish)							'function to send a SVES/QURY
					MEMBER_INFO_ARRAY(sves_qury_sent, each_memb) = qury_finish		'set the output of the qury attempt to the member array

					'we are trying to find and update any rows in the INCOME list where the case number and pmi match exactly and the claim number is close to the SSN to se the QURY information
					objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET QURY_Sent = '" & qury_finish & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) & "%'"

					Set objIncomeConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
					Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

					'opening the connections and data table
					objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
					objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					'If there is RR income listed from the SQL table and NOT from UNEA - it is going to save any member with RR income listed on SQL to the RR array for the verif list
					If MEMBER_INFO_ARRAY(sql_rr_income_exists, each_memb) = True and MEMBER_INFO_ARRAY(unea_RR_exists, each_memb) = False Then
						ReDim Preserve RR_INCOME_ARRAY(rr_last_const, rr_count)

						RR_INCOME_ARRAY(rr_case_numb_const, rr_count) = MAXIS_case_number
						RR_INCOME_ARRAY(rr_ref_numb_const, rr_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
						RR_INCOME_ARRAY(rr_pers_name_const, rr_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						RR_INCOME_ARRAY(rr_pers_ssn_const, rr_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
						RR_INCOME_ARRAY(rr_pers_pmi_const, rr_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						RR_INCOME_ARRAY(rr_inc_type_code_const, rr_count) = income_type_code
						RR_INCOME_ARRAY(rr_inc_type_info_const, rr_count) = "Railroad Retirement"
						RR_INCOME_ARRAY(rr_claim_numb_const, rr_count) = ""
						RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count) = "Unknown"

						rr_count = rr_count + 1
					End If

					'If there is VA income listed from the SQL table and NOT from UNEA - it is going to save any member with VA income listed on SQL to the VA array for the verif list
					If MEMBER_INFO_ARRAY(sql_va_income_exists, each_memb) = True and MEMBER_INFO_ARRAY(unea_VA_exists, each_memb) = False Then
						ReDim Preserve VA_INCOME_ARRAY(va_last_const, va_count)

						VA_INCOME_ARRAY(va_case_numb_const, va_count) = MAXIS_case_number
						VA_INCOME_ARRAY(va_ref_numb_const, va_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
						VA_INCOME_ARRAY(va_pers_name_const, va_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						VA_INCOME_ARRAY(va_pers_ssn_const, va_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
						VA_INCOME_ARRAY(va_pers_pmi_const, va_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = ""
						VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Income"
						VA_INCOME_ARRAY(va_claim_numb_const, va_count) = ""
						VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = "Unknown"

						va_count = va_count + 1
					End If

					'If there is UC income listed from the SQL table and NOT from UNEA - it is going to save any member with UC income listed on SQL to the UC array for the verif list
					If MEMBER_INFO_ARRAY(sql_uc_income_exists, each_memb) = True and MEMBER_INFO_ARRAY(unea_UC_exists, each_memb) = False Then
						ReDim Preserve UC_INCOME_ARRAY(uc_last_const, uc_count)

						UC_INCOME_ARRAY(uc_case_numb_const, uc_count) = MAXIS_case_number
						UC_INCOME_ARRAY(uc_ref_numb_const, uc_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
						UC_INCOME_ARRAY(uc_pers_name_const, uc_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						UC_INCOME_ARRAY(uc_pers_ssn_const, uc_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
						UC_INCOME_ARRAY(uc_pers_pmi_const, uc_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						UC_INCOME_ARRAY(uc_inc_type_code_const, uc_count) = ""
						UC_INCOME_ARRAY(uc_inc_type_info_const, uc_count) = "Unemployment"
						UC_INCOME_ARRAY(uc_claim_numb_const, uc_count) = ""
						UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = "Unknown"

						uc_count = uc_count + 1
					End If
				Next

				'Now that we have an array saved, we are going to add it to the Excel sheet right away for UC, VA, or RR income.
				'We do it all at once because if we have a script error, this way we don't lose the information

				If va_count <> 0 and last_va_count <> va_count Then	'if there is VA income found and the va income found has incremented up since the last loop
					va_excel_created = True						'Identifying that the VA excel list was created

					'adding any va income from the array to the spreadsheet
					Do
						objVAExcel.Cells(va_excel_row, 1).value = VA_INCOME_ARRAY(va_case_numb_const, va_inc_count)
						objVAExcel.Cells(va_excel_row, 2).value = VA_INCOME_ARRAY(va_ref_numb_const, va_inc_count)
						objVAExcel.Cells(va_excel_row, 3).value = VA_INCOME_ARRAY(va_pers_name_const, va_inc_count)
						objVAExcel.Cells(va_excel_row, 4).value = VA_INCOME_ARRAY(va_pers_pmi_const, va_inc_count)
						objVAExcel.Cells(va_excel_row, 5).value = VA_INCOME_ARRAY(va_pers_ssn_const, va_inc_count)
						If VA_INCOME_ARRAY(va_inc_type_code_const, va_inc_count) <> "" Then objVAExcel.Cells(va_excel_row, 6).value = VA_INCOME_ARRAY(va_inc_type_code_const, va_inc_count) & " - " & VA_INCOME_ARRAY(va_inc_type_info_const, va_inc_count)
						If VA_INCOME_ARRAY(va_inc_type_code_const, va_inc_count) = "" Then objVAExcel.Cells(va_excel_row, 6).value = VA_INCOME_ARRAY(va_inc_type_info_const, va_inc_count)
						objVAExcel.Cells(va_excel_row, 7).value = VA_INCOME_ARRAY(va_claim_numb_const, va_inc_count)
						objVAExcel.Cells(va_excel_row, 8).value = VA_INCOME_ARRAY(va_prosp_inc_const, va_inc_count)
						objVAWorkbook.Save()

						va_inc_count = va_inc_count + 1			'going to the next array item
						va_excel_row = va_excel_row + 1			'going to the next row
					Loop until va_inc_count = va_count			'loop until the income count gets to the total of va counted
				End If

				If uc_count <> 0 and last_uc_count <> uc_count Then	'If there is UC income found and the UC income found has incremented up since the last loop
					uc_excel_created = True						'Identifying that the UC excel list was created

					'adding any uc income from the array to the spreadsheet
					Do
						objUCExcel.Cells(uc_excel_row, 1).value = UC_INCOME_ARRAY(uc_case_numb_const, uc_inc_count)
						objUCExcel.Cells(uc_excel_row, 2).value = UC_INCOME_ARRAY(uc_ref_numb_const, uc_inc_count)
						objUCExcel.Cells(uc_excel_row, 3).value = UC_INCOME_ARRAY(uc_pers_name_const, uc_inc_count)
						objUCExcel.Cells(uc_excel_row, 4).value = UC_INCOME_ARRAY(uc_pers_pmi_const, uc_inc_count)
						objUCExcel.Cells(uc_excel_row, 5).value = UC_INCOME_ARRAY(uc_pers_ssn_const, uc_inc_count)
						If UC_INCOME_ARRAY(uc_inc_type_code_const, uc_inc_count) <> "" Then objUCExcel.Cells(uc_excel_row, 6).value = UC_INCOME_ARRAY(uc_inc_type_code_const, uc_inc_count) & " - " & UC_INCOME_ARRAY(uc_inc_type_info_const, uc_inc_count)
						If UC_INCOME_ARRAY(uc_inc_type_code_const, uc_inc_count) = "" Then objUCExcel.Cells(uc_excel_row, 6).value = UC_INCOME_ARRAY(uc_inc_type_info_const, uc_inc_count)
						objUCExcel.Cells(uc_excel_row, 7).value = UC_INCOME_ARRAY(uc_claim_numb_const, uc_inc_count)
						objUCExcel.Cells(uc_excel_row, 8).value = UC_INCOME_ARRAY(uc_prosp_inc_const, uc_inc_count)
						objUCWorkbook.Save()

						uc_inc_count = uc_inc_count + 1			'going to the next array item
						uc_excel_row = uc_excel_row + 1			'going to the next row
					Loop until uc_inc_count = uc_count			'loop until the income count gets to the total of uc counted
				End If


				If rr_count <> 0 and last_rr_count <> rr_count Then		'If there is RR income found and the RR income found has incremented up since the last loop
					rr_excel_created = True							'Identifying that the RR excel list was created

					'adding any rr income from the array to the spreadsheet
					Do
						objRRExcel.Cells(rr_excel_row, 1).value = RR_INCOME_ARRAY(rr_case_numb_const, rr_inc_count)
						objRRExcel.Cells(rr_excel_row, 2).value = RR_INCOME_ARRAY(rr_ref_numb_const, rr_inc_count)
						objRRExcel.Cells(rr_excel_row, 3).value = RR_INCOME_ARRAY(rr_pers_name_const, rr_inc_count)
						objRRExcel.Cells(rr_excel_row, 4).value = RR_INCOME_ARRAY(rr_pers_pmi_const, rr_inc_count)
						objRRExcel.Cells(rr_excel_row, 5).value = RR_INCOME_ARRAY(rr_pers_ssn_const, rr_inc_count)
						If RR_INCOME_ARRAY(rr_inc_type_code_const, rr_inc_count) <> "" Then objRRExcel.Cells(rr_excel_row, 6).value = RR_INCOME_ARRAY(rr_inc_type_code_const, rr_inc_count) & " - " & RR_INCOME_ARRAY(rr_inc_type_info_const, rr_inc_count)
						If RR_INCOME_ARRAY(rr_inc_type_code_const, rr_inc_count) = "" Then objRRExcel.Cells(rr_excel_row, 6).value = RR_INCOME_ARRAY(rr_inc_type_info_const, rr_inc_count)
						objRRExcel.Cells(rr_excel_row, 7).value = RR_INCOME_ARRAY(rr_claim_numb_const, rr_inc_count)
						objRRExcel.Cells(rr_excel_row, 8).value = RR_INCOME_ARRAY(rr_prosp_inc_const, rr_inc_count)
						objRRWorkbook.Save()

						rr_inc_count = rr_inc_count + 1			'going to the next array item
						rr_excel_row = rr_excel_row + 1			'going to the next row
					Loop until rr_inc_count = rr_count			'loop until the income count gets to the total of rr counted
				End If
			End If

			Call back_to_SELF				'getting back to base

			'Now we are going to update the case list with the Ex parte evaluation done. This also removes the 'In Progress' marker
			prep_status = date														'prep status should be a date
			If appears_ex_parte = False Then prep_status = "Not Ex Parte"			'if this case is not ex parte, the prep status is reset

			'here is the update statement. setting the exparte eval and the income/case information for the case running
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & appears_ex_parte & "', PREP_Complete = '" & prep_status & "', AllHCisABD = '" & all_hc_is_ABD & "', SSAIncomExist = '" & SSA_income_exists & "', WagesExist = '" & JOBS_income_exists & "', VAIncomeExist = '" & VA_income_exists & "', SelfEmpExists = '" & BUSI_income_exists & "', NoIncome = '" & case_has_no_income & "', EPDonCase = '" & case_has_EPD & "' WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
		End If
		objRecordSet.MoveNext			'now we go to the next case
	Loop

	'now we format and save the verification lists
	For col_to_autofit = 1 to 9
		If va_excel_created = True Then objVAExcel.columns(col_to_autofit).AutoFit()
		If uc_excel_created = True Then objUCExcel.columns(col_to_autofit).AutoFit()
		If rr_excel_created = True Then objRRExcel.columns(col_to_autofit).AutoFit()
	Next

	If va_excel_created = True Then
		objVAWorkbook.Save()
		objVAExcel.ActiveWorkbook.Close
		objVAExcel.Application.Quit
		objVAExcel.Quit
	End If
	If uc_excel_created = True Then
		objUCWorkbook.Save()
		objUCExcel.ActiveWorkbook.Close
		objUCExcel.Application.Quit
		objUCExcel.Quit
	End If
	If rr_excel_created = True Then
		objRRWorkbook.Save()
		objRRExcel.ActiveWorkbook.Close
		objRRExcel.Application.Quit
		objRRExcel.Quit
	End If

    objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	'TODO Add Automation to send the emails with the Excel files attached.

	'We are going to set the display message for the end of the script run
	end_msg = "BULK Prep 1 Run has been completed."

	'declare the SQL statement that will query the database for all cases with the review month we are evaluating
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1															'counting all the cases
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1	'counting all the ex parte cases
		objRecordSet.MoveNext		'go to the next case
	Loop
    objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	percent_ex_parte = ex_parte_count/case_count						'doing some calculations for see percentages
	percent_ex_parte = percent_ex_parte * 100
	percent_ex_parte = FormatNumber(percent_ex_parte, 2, -1, 0, -1)

	'Creating an end message to display the case list counts
	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count
	end_msg = end_msg & vbCr & "This appears to be " & percent_ex_parte & "% of cases."

	'This is the end of the fucntionality and will just display the end message at the end of this script file.
End If

'This functionality will be run about 5 days after the first PREP run.
'This will read the SVES TPQY information that was received from the QURY in PREP 1 and update the STAT panels
If ex_parte_function = "Prep 2" Then
	'this is for testing - we want to know
	'Creating a txt file output of cases in where there is a second TPQY.
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	tracking_doc_file = user_myDocs_folder & "ExParte Tracking Lists/" & ep_revw_mo & "-" & ep_revw_yr & " - prep 2 sept second tpqy list.txt"
	If ObjFSO.FileExists(tracking_doc_file) Then		'If the file exists we open it and set to add to it
		Set objTextStream = ObjFSO.OpenTextFile(tracking_doc_file, ForAppending, true)
	Else												'If the file does not exists, we create it and set to writing the file
		Set objTextStream = ObjFSO.CreateTextFile(tracking_doc_file, ForWriting, true)
	End If
	objTextStream.WriteLine "LIST START"		'This is going to head each start of the script run.

	review_date = ep_revw_mo & "/1/" & ep_revw_yr			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	review_date = DateAdd("d", 0, review_date)

	'This functionality will remove any 'holds' with 'In Progress' marked. This is to make sure no cases get left behind if a script fails
	If reset_in_Progress = checked Then
		'This is opening the Ex Parte Case List data table so we can loop through it.
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

		Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		'Loop through each item on the CASE LIST Table
		Do While NOT objRecordSet.Eof
			If objRecordSet("PREP_Complete") = "In Progress" Then			'If the case is marked as 'In Progress' - we are going to remove it
				MAXIS_case_number = objRecordSet("CaseNumber") 				'SET THE MAXIS CASE NUMBER

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = '" & NULL & "'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"	'removing the 'In Progress' indicator and blanking it out

				Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			End If
			objRecordSet.MoveNext			'move to the next item in the table
		Loop
		objRecordSet.Close			'Closing all the data connections
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
	End If

	If ObjFSO.FileExists(ex_parte_folder & "\MEMBS with TPQY Date of Death - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx") Then
		Call excel_open(ex_parte_folder & "\MEMBS with TPQY Date of Death - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx", True, False, ObjExcel, objWorkbook)
		excel_row = 2
		Do
			listed_case_numb = trim(ObjExcel.Cells(excel_row, 1).value)
			If listed_case_numb <> "" Then
				excel_row = excel_row + 1
			End If
		Loop until listed_case_numb = ""
	Else
		'Opening a spreadsheet to capture the cases with a SMRT ending soon
		Set ObjExcel = CreateObject("Excel.Application")
		ObjExcel.Visible = True
		Set objWorkbook = ObjExcel.Workbooks.Add()
		ObjExcel.DisplayAlerts = True

		'Setting the first 4 col as worker, case number, name, and APPL date
		ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
		ObjExcel.Cells(1, 2).Value = "REF"
		ObjExcel.Cells(1, 3).Value = "NAME"
		ObjExcel.Cells(1, 4).Value = "PMI NUMBER"
		ObjExcel.Cells(1, 5).Value = "SSN"
		ObjExcel.Cells(1, 6).Value = "Date of Death"

		FOR i = 1 to 6		'formatting the cells'
			ObjExcel.Cells(1, i).Font.Bold = True		'bold font'
		NEXT

		'Formatting the table created in the list of date of death that is listed
		ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, ObjExcel.Range("A1:H" & excel_row - 1), xlYes).Name = "Table1"
		ObjExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
		ObjExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\MEMBS with TPQY Date of Death - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"

		excel_row = 2		'initializing the counter to move through the excel lines
	End If

	yesterday = DateAdd("d", -1, date)		'defining yesterday

	'This is opening the Ex Parte Case List data table so we can loop through it.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Loop through each item on the CASE LIST Table
	Do While NOT objRecordSet.Eof
		'We are selecting cases that are indicated as Ex Parte
		'We need to determine if the information in the table necessitates the functionality be run as a separate logic statement
		work_this_case = True
		If IsDate(objRecordSet("PREP_Complete")) = True Then
			prep_complete_date = objRecordSet("PREP_Complete")			'pulling this into a seperate variable allows us to do things with it, like MAKE SURE it is treated as a date
			prep_complete_date = DateAdd("d", 0, prep_complete_date)	'force it to be a date
			If prep_complete_date = date Then work_this_case = False	'if this was already completed today or yesterday, we do not need to run the functionality again
			If prep_complete_date = yesterday Then work_this_case = False
		Else
			'if this is not a date, then we only work it if it null or blank
			If objRecordSet("PREP_Complete") = "In Progress" Then work_this_case = False
			If IsNull(objRecordSet("PREP_Complete")) = False and objRecordSet("PREP_Complete") <> "" Then work_this_case = False
		End If

		'determining which case on this list we should work.
		If objRecordSet("SelectExParte") = True and work_this_case = True Then
			'For each case that is indicated as Ex parte, we are going to update the case information
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			'Here we are setting the PREP_Complete to 'In Progress' to hold the case as being worked.
			'This portion of the script is required to be able to have more than one person operating the BULK run at the same time.
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = 'In Progress'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			'Here is functionality to be sure the case is able to be updated
			case_is_in_henn = False					'default this to false

			'reading case program information and PW
			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
			EMReadScreen case_pw, 7, 21, 14									'reading the curent PW for the case
			If left(case_pw, 4) = "X127" Then case_is_in_henn = True		'identifying if the case is not in HENN
			kick_it_off_reason = ""											'create an explanation of why the case is being removed form the Ex Parte list
			If case_is_in_henn = False Then kick_it_off_reason = "Case not in 27"
			If case_active = False Then kick_it_off_reason = "Case not Active"
			If (case_active = False and case_pending = False and case_rein = False) or case_is_in_henn = False Then
				'WE ARE NOT going to update this here for now
				' select_ex_parte = False
				' objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & select_ex_parte & "', PREP_Complete = '" & kick_it_off_reason & "' WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

				' Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				' Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				' 'opening the connections and data table
				' objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				' objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			Else
				ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)							'Reset this array to blank at the beginning of each loop for each case.
				Do
					Call navigate_to_MAXIS_screen("STAT", "MEMB")					'making suyre we get to STAT MEMB
					EMReadScreen memb_check, 4, 2, 48
				Loop until memb_check = "MEMB"
				Call get_list_of_members											'get a list of all the HH memebers on the case

				'Read SVES/TPQY for all persons on a case
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False			'defaulting these to false
					MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = False
					MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = False

					Call navigate_to_MAXIS_screen("INFC", "SVES")					'navigate to SVES
					EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68				'Enter the PMI for the current member and open the TPQY
					Call write_value_and_transmit("TPQY", 20, 70)

					Do
						EMReadScreen check_TPQY_panel, 4, 2, 53 						'Reads for TPQY panel
						If check_TPQY_panel <> "TPQY" Then Call write_value_and_transmit("TPQY", 20, 70)
					Loop until check_TPQY_panel = "TPQY"

					EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb), 		1, 8, 39		'saving all tpqy information into the member array
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb), 		1, 8, 65
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), 		10, 6, 61
					MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb))
					MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), " ", "/")
					EMReadScreen sves_response, 8, 7, 22 		'Return Date
					sves_response = replace(sves_response," ", "/")

					transmit

					Do
						EMReadScreen check_BDXP_panel, 4, 2, 53 						'Reads for BDXP panel and makes sure we are there
						If check_BDXP_panel <> "BDXP" Then
							row = 1
							col = 1
							EMSearch "Command:", row, col
							Call write_value_and_transmit("BDXP", row, col++9)
						End If
					Loop until check_BDXP_panel = "BDXP"

					EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 	12, 5, 40		'saving all tpqy information into the member array
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 		12, 5, 69
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb), 	2, 6, 19
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), 	8, 8, 16
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb), 		8, 8, 32
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb), 		1, 8, 69
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), 	5, 11, 69
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), 	5, 14, 69
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), 	10, 15, 69
					MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb))
					MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb))
					MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), " ", "")
					MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb))
					MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb))
					MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb))
					MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb))
					MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = Trim (MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), " ", "/1/")
					MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), " ", "/1/")
					MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), " ", "/")

					transmit

					Do
						EMReadScreen check_BDXM_panel, 4, 2, 53 						'Reads for BDXM panel
						If check_BDXM_panel <> "BDXM" Then
							row = 1
							col = 1
							EMSearch "Command:", row, col
							Call write_value_and_transmit("BDXM", row, col++9)
						End If
					Loop until check_BDXM_panel = "BDXM"

					EMReadScreen MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb), 			13, 4, 29		'saving all tpqy information into the member array
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb), 			7, 6, 64
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 				5, 7, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), 				5, 7, 63
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_ind, each_memb), 			1, 8, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb), 			3, 8, 63
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb), 	5, 9, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb), 	5, 9, 63
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb), 			7, 12, 64
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 				5, 13, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), 				5, 13, 63
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_ind, each_memb), 			1, 14, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb), 			3, 14, 63
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), 	5, 15, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb), 	5, 15, 63
					MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb))
					MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb))
					MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb))
					MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb))
					MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb))
					MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb))
					MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb))
					MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb))
					MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb))
					MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), " ", "/01/")
					MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), " ", "/01/")
					MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb), " ", "/01/")
					MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb), " ", "/01/")
					MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), " ", "/01/")
					MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), " ", "/01/")
					MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), " ", "/01/")
					MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb), " ", "/01/")

					transmit

					Do
						EMReadScreen check_SDXE_panel, 4, 2, 53 						'Reads for SDXE panel
						If check_SDXE_panel <> "SDXE" Then
							row = 1
							col = 1
							EMSearch "Command:", row, col
							Call write_value_and_transmit("SDXE", row, col++9)
						End If
					Loop until check_SDXE_panel = "SDXE"

					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb), 		12, 5, 36		'saving all tpqy information into the member array
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), 		2, 7, 21
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb), 		22, 7, 24
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_fed_living, each_memb), 			1, 6, 70
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 			3, 8, 21
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb), 			30, 8, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_cit_ind_code, each_memb), 			1, 7, 70
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_code, each_memb), 		3, 10, 26
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb), 		40, 10, 30
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb), 		8, 11, 26
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb), 			8, 12, 26
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), 		8, 13, 26
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_code, each_memb), 		1, 11, 65
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb), 		8, 12, 65
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_code, each_memb), 	2, 13, 65
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb), 	8, 14, 65
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_disa_pay_code, each_memb), 		1, 15, 65
					MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb), " ", "/")
					MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb), " ", "/")
					MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), " ", "/")
					MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb), " ", "/")
					MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb), " ", "/")

					transmit

					Do
						EMReadScreen check_SDXP_panel, 4, 2, 50 							'Reads for SDXP panel
						If check_SDXP_panel <> "SDXP" Then
							row = 1
							col = 1
							EMSearch "Command:", row, col
							Call write_value_and_transmit("SDXP", row, col++9)
						End If
					Loop until check_SDXP_panel = "SDXP"

					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), 			5, 4, 16		'saving all tpqy information into the member array
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), 			7, 4, 42
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_over_under_code, each_memb), 	1, 4, 73
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb), 	5, 8, 3
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb), 	6, 8, 13
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_type, each_memb), 	1, 8, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb), 	5, 9, 3
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb), 	6, 9, 13
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_type, each_memb), 	1, 9, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb), 	5, 10, 3
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb), 	6, 10, 13
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_type, each_memb), 	1, 10, 25
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb), 				8, 5, 66
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb), 				8, 6, 66
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb), 		8, 7, 66
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb), 		8, 8, 66
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb), 		8, 9, 66
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb), 		8, 10, 66
					EMReadScreen MEMBER_INFO_ARRAY(tpqy_rep_payee, each_memb), 				1, 11, 66

					If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) <> "C01" Then
						last_payment_date = ""
						sdx_row = 8
						Do
							EMReadScreen sdx_payment_type, 1, sdx_row, 25
							If sdx_payment_type <> "0" and sdx_payment_type <> " " Then
								EMReadScreen last_payment_date, 5, sdx_row, 3
								EMReadScreen last_payment_amt, 9, sdx_row, 13
								Exit Do
							End If
							sdx_row = sdx_row + 1
						Loop until sdx_payment_type = " "
						If last_payment_date <> "" Then
							MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb) = replace(last_payment_date, " ", "/1/")
							MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb) = DateAdd("d", 0, MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_amt, each_memb) = trim(last_payment_amt)
						End If
					End If

					MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb))
					If MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "" Then MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "0"
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb))
					MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb))
					MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb))
					MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb))
					MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb))
					MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb))
					MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb))
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), " ", "/01/")
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb), " ", "/")
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb), " ", "/")
					MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb), " ", "/")
					MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb), " ", "/")
					MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb), " ", "/")

					transmit

					If MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb) = "Y" Then
						MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True
						MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb)= False
						If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) = "C01" Then MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb) = True
						If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "E" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
						If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "H" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
						If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "M" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
						If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "P" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
						If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "S" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
					End If
					If MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb) = "Y" Then
						If MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "C" or MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "E" Then
							MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True
							If IsDate(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb)) = True Then MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True
						End If
					End If

					objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [QURY_Sent] != 'NULL'"

					Set objIncomeConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
					Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

					'opening the connections and data table
					objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
					objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					Call back_to_SELF
				Next

				'navigating into STAT
				Do
					Call navigate_to_MAXIS_screen("STAT", "SUMM")
					EMReadScreen summ_check, 4, 2, 46
				Loop until summ_check = "SUMM"
				verif_types = ""						'blanking out the list of verifications for the CASE/NOTE

				'here we attempt to go update STAT with the information gathered from TPQY
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)									'looping thorugh each HH Member
					If IsDate(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb)) = True Then 		'If there is a date of dealth listed, for now we are just going to add them to a list
						ObjExcel.Cells(excel_row, 1).Value = MAXIS_case_number
						ObjExcel.Cells(excel_row, 2).Value = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
						ObjExcel.Cells(excel_row, 3).Value = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						ObjExcel.Cells(excel_row, 4).Value = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						ObjExcel.Cells(excel_row, 5).Value = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
						ObjExcel.Cells(excel_row, 6).Value = MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb)
						objWorkbook.Save()
						excel_row = excel_row + 1													'counting to increment to the next excel row

					Else 	'If there is no date of death, we are going to try to update UNEA for SSI/RSDI
						'Update MAXIS UNEA panels with information from TPQY
						If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True Then				'Member with SSI
							If MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb) = True Then		'If SSI appears to be ongoing (Current Pay)
								Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), "03", SSI_UNEA_instance, "", SSI_panel_found)

								Call update_unea_pane(SSI_panel_found, "03", MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), "", "")
								If InStr(verif_types, "SSI") = 0 Then verif_types = verif_types & "/SSI"
							ElseIf isDate(MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb)) = True Then	'If SSI has an end date listed
								Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), "03", SSI_UNEA_instance, "", SSI_panel_found)

								Call update_unea_pane(SSI_panel_found, "03", MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_amt, each_memb))
								If InStr(verif_types, "SSI End") = 0 Then verif_types = verif_types & "/SSI End"
							'There is no handling for if person appears to have SSI ended but we could not find an end date.
							End If
						End If

						If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then				'Member with RSDI
							If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then
								rsdi_type = "01"
								Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
							Else
								rsdi_type = "02"
								Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
							End If
							Call update_unea_pane(RSDI_panel_found, rsdi_type, MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), "", "")
							If InStr(verif_types, "RSDI") = 0 Then verif_types = verif_types & "/RSDI"
						End If

						'Update MAXIS MEDI panels with information from TPQY
						MEDI_panel_exists = False
						MEMBER_INFO_ARRAY(created_medi, each_memb) = False
						If MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) <> "" Then

							EMWriteScreen "MEDI", 20, 71
							transmit
							EMReadScreen medi_check, 4, 2, 44
							Do while medi_check <> "MEDI"
								Call navigate_to_MAXIS_screen("STAT", "MEDI")
								EMReadScreen medi_check, 4, 2, 44
							Loop

							EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
							transmit

							EMReadScreen total_amt_of_panels, 1, 2, 78			'Checks to make sure there are JOBS panels for this member. If none exists, one will be created
							MEDI_panel_exists = True
							MEDI_active = False
							If total_amt_of_panels = "0" Then MEDI_panel_exists = False
							If (MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) <> "" and MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = "") or (MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) <> "" and MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = "") Then MEDI_active = True
							part_a_ended = False
							If IsDate(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb)) = True Then part_a_ended = True
							part_b_ended = False
							If IsDate(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb)) = True Then part_b_ended = True

							panel_part_a_accurate = False
							panel_part_b_accurate = False
							If MEDI_panel_exists = True Then
								Do
									PF20
									EMReadScreen end_of_list, 9, 24, 14
								Loop Until end_of_list = "LAST PAGE"
								row = 17
								Do
									EMReadScreen begin_dt_a, 8, row, 24 		'reads part a start date
									begin_dt_a = replace(begin_dt_a, " ", "/")	'reformatting with / for date
									If begin_dt_a = "__/__/__" Then begin_dt_a = "" 		'blank out if not a date

									EMReadScreen end_dt_a, 8, row, 35	'reads part a end date
									end_dt_a =replace(end_dt_a , " ", "/")		'reformatting with / for date
									If end_dt_a = "__/__/__" Then end_dt_a = ""					'blank out if not a date

									If part_a_ended = True Then
										If end_dt_a <> "" Then
											panel_part_a_accurate = True
											Exit Do
										End If
										If end_dt_a = "" and begin_dt_a <> "" Then Exit Do
									Else
										If begin_dt_a <> "" and end_dt_a <> "" Then
											Exit Do
										ElseIf begin_dt_a <> "" and end_dt_a = "" Then
											panel_part_a_accurate = True
											Exit Do
										End If
									End If
									row = row - 1

									If row = 14 Then
										PF19
										EMReadScreen begining_of_list, 10, 24, 14
										' MsgBox "begining_of_list - " & begining_of_list & vbcr & "1"
										If begining_of_list = "FIRST PAGE" Then
											Exit Do
										Else
											row = 17
										End If
									End If
								Loop
								Do
									PF19
									EMReadScreen begining_of_list, 10, 24, 14
								Loop Until begining_of_list = "FIRST PAGE"

								Do
									PF20
									EMReadScreen end_of_list, 9, 24, 14
								Loop Until end_of_list = "LAST PAGE"
								row = 17
								Do
									EMReadScreen begin_dt_b, 8, row, 54 		'reads part a start date
									begin_dt_b = replace(begin_dt_b, " ", "/")	'reformatting with / for date
									If begin_dt_b = "__/__/__" Then begin_dt_b = "" 		'blank out if not a date

									EMReadScreen end_dt_b, 8, row, 65	'reads part a end date
									end_dt_b =replace(end_dt_b , " ", "/")		'reformatting with / for date
									If end_dt_b = "__/__/__" Then end_dt_b = ""					'blank out if not a date

									If part_a_ended = True Then
										If end_dt_b <> "" Then
											panel_part_b_accurate = True
											Exit Do
										End If
										If end_dt_b = "" and begin_dt_b <> "" Then Exit Do
									Else
										If begin_dt_b <> "" and end_dt_b <> "" Then
											Exit Do
										ElseIf begin_dt_b <> "" and end_dt_b = "" Then
											panel_part_b_accurate = True
										End If
									End If
									row = row - 1

									If row = 14 Then
										PF19
										EMReadScreen begining_of_list, 10, 24, 14

										If begining_of_list = "FIRST PAGE" Then
											Exit Do
										Else
											row = 17
										End If
									End If
								Loop

							End If
							If MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = "" and MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = "" Then panel_part_a_accurate = True
							If MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = "" and MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = "" Then panel_part_b_accurate = True

							If MEDI_panel_exists = True and (panel_part_a_accurate = False or panel_part_b_accurate = False) Then
								If InStr(verif_types, "Medicare") = 0 Then verif_types = verif_types & "/Medicare"
								PF9
								part_a_error = False
								part_b_error = False
								If panel_part_a_accurate = False Then
									Do
										EMReadScreen begin_date_three, 8, 17, 24
										EMReadScreen end_date_three, 8, 17, 35
										If begin_date_three = "__ __ __" and end_date_three = "__ __ __" Then Exit Do

										PF20
										EMReadScreen end_of_list, 34, 24, 2

										If InStr(end_of_list, "BEGIN DATE IS REQUIRED") <> 0 and InStr(end_of_list, "PART A") <> 0 Then
											If end_date_three <> "__ __ __" and part_a_ended = True Then
												part_a_error = True
												exit Do
											End If
										End If
									Loop Until end_of_list = "COMPLETE THE PAGE BEFORE SCROLLING"

									If part_a_error = False Then
										row = 17

										Do
											EMReadScreen begin_dt_a, 8, row, 24 		'reads part a start date
											begin_dt_a = replace(begin_dt_a, " ", "/")	'reformatting with / for date
											If begin_dt_a = "__/__/__" Then begin_dt_a = "" 		'blank out if not a date

											EMReadScreen end_dt_a, 8, row, 35	'reads part a end date
											end_dt_a =replace(end_dt_a , " ", "/")		'reformatting with / for date
											If end_dt_a = "__/__/__" Then end_dt_a = ""					'blank out if not a date

											If part_a_ended = True Then
												If end_dt_a <> "" Then Exit Do
												If end_dt_a = "" and begin_dt_a <> "" Then
													MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True
													EMReadScreen verif_code, 1, row, 47
													If verif_code <> "V" Then
														Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), row, 35, "YY")
													Else
														If row = 17 Then
															PF20
															Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 15, 24, "YY")
															Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), 15, 35, "YY")
														Else
															Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), row+1, 24, "YY")
															Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), row+1, 35, "YY")
														End If
													End If
													Exit Do
												End If
											Else
												If begin_dt_a <> "" and end_dt_a <> "" Then
													MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True
													If row = 17 Then
														PF20
														Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 15, 24, "YY")
													Else
														Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), row+1, 24, "YY")
													End If
													Exit Do
												ElseIf begin_dt_a <> "" and end_dt_a = "" Then
													Exit Do
												End If
											End If
											row = row - 1

											If row = 14 Then
												PF19
												EMReadScreen begining_of_list, 10, 24, 14
												If InStr(begining_of_list, " DATE IS R") <> 0 Then Exit Do

												If begining_of_list = "FIRST PAGE" Then
													Exit Do
												Else
													row = 17
												End If
											End If
										Loop
										Do
											PF19
											EMReadScreen begining_of_list, 10, 24, 14
											If InStr(begining_of_list, " DATE IS R") <> 0 Then Exit Do
										Loop Until begining_of_list = "FIRST PAGE"
									End If
								End If
								If panel_part_b_accurate = False Then
									Do
										EMReadScreen begin_date_three, 8, 17, 54
										EMReadScreen end_date_three, 8, 17, 65

										If begin_date_three = "__ __ __" and end_date_three = "__ __ __" Then Exit Do
										PF20
										EMReadScreen end_of_list, 34, 24, 2

										If InStr(end_of_list, "BEGIN DATE IS REQUIRED") <> 0 and InStr(end_of_list, "PART B") <> 0 Then
											If end_date_three <> "__ __ __" and part_b_ended = True Then
												part_b_error = True
												exit Do
											End If
										End If
									Loop Until end_of_list = "COMPLETE THE PAGE BEFORE SCROLLING"

									If part_b_error = False Then
										row = 17
										Do
											EMReadScreen begin_dt_b, 8, row, 54 		'reads part a start date
											begin_dt_b = replace(begin_dt_b, " ", "/")	'reformatting with / for date
											If begin_dt_b = "__/__/__" Then begin_dt_b = "" 		'blank out if not a date

											EMReadScreen end_dt_b, 8, row, 65	'reads part a end date
											end_dt_b =replace(end_dt_b , " ", "/")		'reformatting with / for date
											If end_dt_b = "__/__/__" Then end_dt_b = ""					'blank out if not a date

											If part_b_ended = True Then
												If end_dt_b <> "" Then Exit Do
												If end_dt_b = "" and begin_dt_b <> "" Then
													MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True
													EMReadScreen verif_code, 1, row, 77

													If verif_code <> "V" Then
														Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), row, 65, "YY")
													Else
														If row = 17 Then
															PF20
															Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 15, 54, "YY")
															Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), 15, 65, "YY")
														Else
															Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), row+1, 54, "YY")
															Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), row+1, 65, "YY")
														End If
													End If
													Exit Do
												End If
											Else
												If begin_dt_b <> "" and end_dt_b <> "" Then
													MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True
													If row = 17 Then
														PF20
														Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 15, 54, "YY")
													Else
														Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), row+1, 54, "YY")
													End If
													Exit Do
												ElseIf begin_dt_b <> "" and end_dt_b = "" Then
													Exit Do
												End If
											End If
											row = row - 1

											If row = 14 Then
												PF19
												EMReadScreen begining_of_list, 10, 24, 14
												If InStr(begining_of_list, " DATE IS R") <> 0 Then Exit Do
												If begining_of_list = "FIRST PAGE" Then
													Exit Do
												Else
													row = 17
												End If
											End If
										Loop
									End If
								End If
								transmit

								EMReadScreen end_of_list, 34, 24, 2
								If InStr(end_of_list, "BEGIN DATE IS REQUIRED") <> 0 Then
									PF10
									MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = False
									MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = False
								End If
							End If

							If MEDI_panel_exists = False and MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) <> "" Then
								If InStr(verif_types, "Medicare") = 0 Then verif_types = verif_types & "/Medicare"
								If (MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) <> "" and MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = "") or (MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) <> "" and MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = "") Then
									MEMBER_INFO_ARRAY(created_medi, each_memb) = True
									Call write_value_and_transmit("NN", 20, 79)
									medi_claim_array = Null
									medi_claim_array = split(MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb))
									EMWriteScreen medi_claim_array(0), 6, 39
									EMWriteScreen medi_claim_array(1), 6, 43
									EMWriteScreen medi_claim_array(2), 6, 46
									EMWriteScreen left(medi_claim_array(3), 1), 6, 51

									If MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) <> "" Then
										MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True
										Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 15, 24, "YY")
										If MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) <> "" Then Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), 15, 35, "YY")
									End If

									If MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) <> "" Then
										MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True
										Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 15, 54, "YY")
										If MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) <> "" Then
											Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), 15, 65, "YY")
										Else
											If MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) <> "" Then EMWriteScreen MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb), 7, 73
											If MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) = "" Then
												If IsDate(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb)) = True Then Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), 8, 44, "YY")
											End If
										End If
									End If
									transmit
								End If
							End If
						End If
					End If
				Next

				'Send the case through background
				Call write_value_and_transmit("BGTX", 20, 71)
				EMReadScreen wrap_check, 4, 2, 46
				If wrap_check = "WRAP" Then transmit
				EMWaitReady 0, 0												'give a pause here

				EMReadScreen wrap_error, 30, 24, 2
				If wrap_error = "THE COMMAND 'BGTX' NOT ALLOWED" Then transmit

				EMReadScreen database_busy, 23, 4, 44							'Sometimes, when we send a case through background a database record error raises
				If database_busy = "A MAXIS database record" Then transmit  	'we need to transmit past it
				'TODO - there may be a NAT error being raised here, but I don't know what that might be from or if we need to resolve it - there does not seem to be any impact to running the script

				Call back_to_SELF

				'here we are trying to update the INCOME List with the information found in TPQY
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)

					' If ssi_claim_numb <> "" or sves_rsdi_claim_numb <> "" Then
					If MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) <> "" Then
						objIncUpdtSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response &_
										"', GrossAmt = '" & MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) &_
										"', NetAmt = '" & "" &_
										"', EndDate = '" & NULL &_
										"' WHERE [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] = '" & MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb) & "'"

						Set objIncUpdtConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
						Set objIncUpdtRecordSet = CreateObject("ADODB.Recordset")

						'opening the connections and data table
						objIncUpdtConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncUpdtRecordSet.Open objIncUpdtSQL, objIncUpdtConnection

					End If
					If MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) <> "" Then
						objIncUpdtSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response &_
										"', GrossAmt = '" & MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) &_
										"', NetAmt = '" & MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) &_
										"', EndDate = '" & MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) &_
										"' WHERE [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] LIKE '" & left(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 9) & "'"

						Set objIncUpdtConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
						Set objIncUpdtRecordSet = CreateObject("ADODB.Recordset")

						'opening the connections and data table
						objIncUpdtConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncUpdtRecordSet.Open objIncUpdtSQL, objIncUpdtConnection

					End If

					'If TPQY indicates that there may be a secondary clam number, we are going to send a QURY and save that to the INCOME list
					If MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) <> "" Then
						Call send_sves_qury("CLAIM", qury_finish)
						MEMBER_INFO_ARRAY(second_qury_sent, each_memb) = qury_finish
						' MsgBox "qury_finish - " & qury_finish
						objIncUpdtSQL = "UPDATE ES.ES_ExParte_IncomeList SET QURY_Sent = '" & qury_finish & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & left(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 9) & "%'"

						Set objIncUpdtConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
						Set objIncUpdtRecordSet = CreateObject("ADODB.Recordset")

						'opening the connections and data table
						objIncUpdtConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncUpdtRecordSet.Open objIncUpdtSQL, objIncUpdtConnection

						' MsgBox "check INC list - " & MAXIS_case_number
						objTextStream.WriteLine MAXIS_case_number & "| NAME: " & MEMBER_INFO_ARRAY(memb_name_const, each_memb) & "|" & "SSN: " &  MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) & "| CLAIM NUMB: " & MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) & "| QURY FINISH: " & qury_finish
					End If

				Next

				'CASE/NOTE details of the case information
				If left(verif_types, 1) = "/" Then verif_types = right(verif_types, len(verif_types)-1)
				note_title = "Verification of " & verif_types

				If verif_types <> "" Then
					Call navigate_to_MAXIS_screen("CASE", "NOTE")
					EMReadScreen last_note, 55, 5, 25
					EMReadScreen last_note_date, 8, 5, 6
					today_day = right("0"&DatePart("d", date), 2)
					today_mo = right("0"&DatePart("d", date), 2)
					today_yr = right(DatePart("d", date), 2)
					today_as_text = today_mo & "/" & today_day & "/" &today_yr

					last_note = trim(last_note)

					If last_note <> note_title or last_note_date <> today_as_text Then
						start_a_blank_CASE_NOTE
						Call write_variable_in_CASE_NOTE(note_title)
						For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
							If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True or MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
								Call write_variable_in_CASE_NOTE("Income from SSA for MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & " - " & MEMBER_INFO_ARRAY(memb_name_const, each_memb) & ".")
								If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True Then
									Call write_variable_in_CASE_NOTE(" * SSI Income of $ " & MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) & " per month.")
								End If
								If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
									rsdi_inc = "RSDI"
									If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then rsdi_inc = "RSDI, Disa"
									Call write_variable_in_CASE_NOTE(" * " & rsdi_inc & " Income of $ " & MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) & " per month.")
								End If
								Call write_variable_in_CASE_NOTE("   - Verified through worker initiated data match.")
								Call write_variable_in_CASE_NOTE("   - UNEA panel updated eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
							End If

							If IsDate(MEMBER_INFO_ARRAY(second_qury_sent, each_memb)) = True Then
								Call write_variable_in_CASE_NOTE("* Additional QURY sent for Claim numb: XXX-XX-" & right(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), len(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb))-5))
							End If
						Next
						For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
							If MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True or MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True Then
								Call write_variable_in_CASE_NOTE("Medicare for MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & " - " & MEMBER_INFO_ARRAY(memb_name_const, each_memb) & ".")

								If MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True Then
									If MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) <> "" Then
										Call write_variable_in_CASE_NOTE("* Medicare Part A ended " & MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb))
									Else
										Call write_variable_in_CASE_NOTE("* Medicare Part A started " & MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb))
									End If
								End If
								If MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True Then
									If MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) <> "" Then
										Call write_variable_in_CASE_NOTE("* Medicare Part B ended " & MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb))
									Else
										Call write_variable_in_CASE_NOTE("* Medicare Part B started " & MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb))
										If MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) <> "" Then
											Call write_variable_in_CASE_NOTE("  - Part B Premium: $ " &MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb))
										Else
											Call write_variable_in_CASE_NOTE("  - Part B Buy-In Start Date: " & MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb))
										End If

									End If
								End If
								Call write_variable_in_CASE_NOTE("   - Verified through worker initiated data match.")
								Call write_variable_in_CASE_NOTE("   - MEDI panel updated eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
							End If
						Next

						call write_variable_in_case_note("---")
						call write_variable_in_case_note(worker_signature)
						call write_variable_in_case_note("Automated Update")

					End If
				End If

			End If

			'here is the update statement. setting the Phase2 BULK run completion date for the case running
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = '" & date & "' WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
		End If
		objRecordSet.MoveNext			'now we go to the next case
	Loop
    objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	For col_to_autofit = 1 to 6
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next

	objWorkbook.Save()
	ObjExcel.ActiveWorkbook.Close
	ObjExcel.Application.Quit
	ObjExcel.Quit

	'We are going to set the display message for the end of the script run
	end_msg = "BULK Prep 2 Run has been completed for " & review_date & "."

	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1
		If IsNull(objRecordSet("PREP_Complete")) = False Then prep_done_count = prep_done_count + 1
		If objRecordSet("PREP_Complete") = date Then prep_2_count = prep_2_count + 1
		If objRecordSet("PREP_Complete") = yesterday Then prep_2_count = prep_2_count + 1
		objRecordSet.MoveNext
	Loop
    objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	percent_ex_parte = ex_parte_count/case_count						'formatting some percentages
	percent_ex_parte = percent_ex_parte * 100
	percent_ex_parte = FormatNumber(percent_ex_parte, 2, -1, 0, -1)

	'Creating an end message to display the case list counts
	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count
	end_msg = end_msg & vbCr & "This appears to be " & percent_ex_parte & "% of cases."
	end_msg = end_msg & vbCr & vbCr & "Cases with PREP completed: " &  prep_done_count
	end_msg = end_msg & vbCr & "Cases where PREP 2 is completed: " & prep_2_count

	'This is the end of the fucntionality and will just display the end message at the end of this script file.
End If

If ex_parte_function = "Phase 1" Then
	'Creating a txt file output of cases where the income was updated during this run.
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")			'creating the object to connect with the file
	tracking_doc_file = user_myDocs_folder & "ExParte Tracking Lists/Phase 1 " & ep_revw_mo & "-" & ep_revw_yr & " income update list.txt"
	If ObjFSO.FileExists(tracking_doc_file) Then					'If the file exists we open it and set to add to it
		Set objTextStream = ObjFSO.OpenTextFile(tracking_doc_file, ForAppending, true)
	Else															'If the file does not exists, we create it and set to writing the file
		Set objTextStream = ObjFSO.CreateTextFile(tracking_doc_file, ForWriting, true)
	End If
	objTextStream.WriteLine "LIST START"		'This is going to head each start of the script run.

	'TODO - this will likely need to be updated to remove the SMRT list assessment. Need process determination before doing this.
	'loading the excel file paths into the variables based on the naming ocnventions.
	va_excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte\VA Income Verifications\VA Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	uc_excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte\UC Income Verifications\UC Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	rr_excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte\RR Income Verifications\RR Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	smrt_excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex ParteSMRT Ending\SMRT Ending - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"

	'this dialog is necessary for Phase 1 to mark what date the second QURYs were sent and the Excel files with other income information
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 400, 170, "Confirm Ex Parte process"
				Text 10, 15, 75, 10, "Date of Prep 2 Run:"
				EditBox 85, 10, 50, 15, prep_phase_2_run_date
				Text 10, 30, 75, 10, "Load VA List"
				EditBox 10, 40, 325, 15, va_excel_file_path
				Text 10, 60, 75, 10, "Load UC List"
				EditBox 10, 70, 325, 15, uc_excel_file_path
				Text 10, 90, 75, 10, "Load RR List"
				EditBox 10, 100, 325, 15, rr_excel_file_path
				Text 10, 120, 75, 10, "Load SMRT List"
				EditBox 10, 130, 325, 15, smrt_excel_file_path
				ButtonGroup ButtonPressed
					PushButton 185, 150, 210, 15, "Continue, all excel files are accurate.", continue_phase_1_btn
					PushButton 345, 40, 50, 15, "VA BROWSE", va_browse_btn
					PushButton 345, 70, 50, 15, "UC BROWSE", uc_browse_btn
					PushButton 345, 100, 50, 15, "RR BROWSE", rr_browse_btn
					PushButton 345, 130, 50, 15, "SMRT BROWSE", smrt_browse_btn
			EndDialog

			Dialog Dialog1
			cancel_without_confirmation

			'confirming the date was set correctly and that excel files were selected.
			If IsDate(prep_phase_2_run_date) = False Then err_msg = err_msg & vbCr & "* Enter the date the second PREP run was completed. If you are not sure, check the SQL table."
			If right(va_excel_file_path, 5) <> ".xlsx" Then err_msg = err_msg & vbCr & "* Select a valid Excel File for the VA inocme verification."
			If right(uc_excel_file_path, 5) <> ".xlsx" Then err_msg = err_msg & vbCr & "* Select a valid Excel File for the UC inocme verification."

			If err_msg <> "" Then MsgBox "* * * * NOTICE * * * *" & vbCr & err_msg		'displaying possible errors

		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in

	'Here we load the verifications of income that we receive from the Excel files for UC and VA into arrays
	Call excel_open(uc_excel_file_path, True, True, ObjExcel, objWorkbook)			'opening the UC excel

	uc_count = 0			'setting the initial incrementors
	excel_row = 2			'starting at the top of the excel list
	Do
		ReDim Preserve UC_INCOME_ARRAY(uc_last_const, uc_count)						'resize the array

		'adding the information from the Excel to the array
		UC_INCOME_ARRAY(uc_case_numb_const, uc_count) 	= ObjExcel.Cells(1, excel_row).Value
		UC_INCOME_ARRAY(uc_ref_numb_const, uc_count) 	= ObjExcel.Cells(2, excel_row).Value
		UC_INCOME_ARRAY(uc_pers_name_const, uc_count) 	= ObjExcel.Cells(3, excel_row).Value
		UC_INCOME_ARRAY(uc_pers_pmi_const, uc_count) 	= right("00000000" & trim(ObjExcel.Cells(4, excel_row).Value), 8)
		UC_INCOME_ARRAY(uc_pers_ssn_const, uc_count) 	= replace(trim(ObjExcel.Cells(5, excel_row).Value), "-", "")  'left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
		UC_INCOME_ARRAY(uc_inc_type_code_const, uc_count) = "14"
		UC_INCOME_ARRAY(uc_inc_type_info_const, uc_count) = "Unemployment"
		UC_INCOME_ARRAY(uc_claim_numb_const, uc_count) 	= ObjExcel.Cells(7, excel_row).Value
		UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) 	= ObjExcel.Cells(9, excel_row).Value
		If IsDate(trim(ObjExcel.Cells(10, excel_row).Value)) = True then
			UC_INCOME_ARRAY(uc_end_date_const, uc_count) 	= ObjExcel.Cells(10, excel_row).Value
		End if
		If IsNumeric(UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count)) = True Then UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) * 1

		uc_count = uc_count + 1				'count up
		excel_row = excel_row + 1			'go to the next row
		next_case_numb = ObjExcel.Cells(1, excel_row).Value		'loop until there are no more cases
	Loop until next_case_numb = ""

	ObjExcel.ActiveWorkbook.Close		'close the Excel file
	ObjExcel.Application.Quit
	ObjExcel.Quit

	'Now for VA
	Call excel_open(va_excel_file_path, True, True, ObjExcel, objWorkbook)		'Opening the VA excel

	va_count = 0			'setting the initial inrementors
	excel_row = 2
	Do
		ReDim Preserve VA_INCOME_ARRAY(va_last_const, va_count)					'resize the array

		' adding the information from the Excel to the array
		VA_INCOME_ARRAY(va_case_numb_const, va_count) 	= trim(ObjExcel.Cells(1, excel_row).Value)		'MAXIS_case_number
		VA_INCOME_ARRAY(va_ref_numb_const, va_count) 	= trim(ObjExcel.Cells(2, excel_row).Value)		'MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
		VA_INCOME_ARRAY(va_pers_name_const, va_count) 	= trim(ObjExcel.Cells(3, excel_row).Value)		'MEMBER_INFO_ARRAY(memb_name_const, each_memb)
		VA_INCOME_ARRAY(va_pers_pmi_const, va_count) 	= right("00000000" & trim(ObjExcel.Cells(4, excel_row).Value), 8)		'MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
		VA_INCOME_ARRAY(va_pers_ssn_const, va_count) 	= replace(trim(ObjExcel.Cells(5, excel_row).Value), "-", "")		'left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
		VA_INCOME_ARRAY(va_claim_numb_const, va_count) 	= trim(ObjExcel.Cells(7, excel_row).Value)
		VA_INCOME_ARRAY(va_prosp_inc_const, va_count) 	= trim(ObjExcel.Cells(9, excel_row).Value)
		If IsNumeric(VA_INCOME_ARRAY(va_prosp_inc_const, va_count)) = True Then VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = VA_INCOME_ARRAY(va_prosp_inc_const, va_count) * 1

		va_type_from_excel = trim(ObjExcel.Cells(6, excel_row).Value)			'splitting UNEA type code and the information detail
		If InStr(va_type_from_excel, "-") = 0 Then
			VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = va_type_from_excel
		Else
			temp_array = split(va_type_from_excel, "-")
			VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = trim(temp_array(0))
			VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = trim(temp_array(1))
		End If

		va_count = va_count + 1				'count up
		excel_row = excel_row + 1			'go to the next row
		next_case_numb = ObjExcel.Cells(1, excel_row).Value		'Loop until there are no more cases
	Loop until next_case_numb = ""

	ObjExcel.ActiveWorkbook.Close		'close the Excel file
	ObjExcel.Application.Quit
	ObjExcel.Quit

	review_date = ep_revw_mo & "/1/" & ep_revw_yr			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	review_date = DateAdd("d", 0, review_date)				'forcint it to be a date

	'This functionality will remove any 'holds' with 'In Progress' marked. This is to make sure no cases get left behind if a script fails
	If reset_in_Progress = checked Then
		'This is opening the Ex Parte Case List data table so we can loop through it.
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

		Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		'Loop through each item on the CASE LIST Table
		Do While NOT objRecordSet.Eof
			If objRecordSet("Phase1Complete") = "In Progress" Then			'If the case is marked as 'In Progress' - we are going to remove it
				MAXIS_case_number = objRecordSet("CaseNumber") 				'SET THE MAXIS CASE NUMBER

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase1Complete = '" & NULL & "'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"	'removing the 'In Progress' indicator and blanking it out

				Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			End If
			objRecordSet.MoveNext			'move to the next item in the table
		Loop
		objRecordSet.Close			'Closing all the data connections
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
	End If

	'Open The CASE LIST Table
	'This is opening the Ex Parte Case List data table so we can loop through it.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"			'we're pulling all the cases based on the renewal month coded in

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Loop through each item on the CASE LIST Table
	Do While NOT objRecordSet.Eof
		If objRecordSet("SelectExParte") = True and (objRecordSet("Phase1Complete") = "" or IsNull(objRecordSet("Phase1Complete")) = True) Then
			kick_it_off_reason = ""			'resetting variables used to make case asseessments to be sure that information from other cases doesn't carry through the loops
			case_active = ""
			case_is_in_henn = ""
			is_this_priv = ""

			'For each case that is indicated as Ex parte, we are going to update the case information
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			'Here we are setting the Phase1Complete to 'In Progress' to hold the case as being worked.
			'This portion of the script is required to be able to have more than one person operating the BULK run at the same time.
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase1Complete = 'In Progress'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			'first we check for access
			case_is_in_henn = False					'default this to false, we will read the PW to see if it is in hennepin county
			Do
				Call back_to_SELF
				Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)			'CASE/CURR will ahve good case information.
				EMReadScreen curr_check, 4, 2, 55												'making sure the script gets to CASE/CURR and it isn't held up somehow
			Loop until curr_check = "CURR" or is_this_priv = True

			If is_this_priv = True Then kick_it_off_reason = "PRIV case"						'PRIV cases cannot be assessed for Ex Parte correctly
			If is_this_priv = False Then														'If not PRIV, we need case information and some other case details
				Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
				EMReadScreen case_pw, 7, 21, 14													'identifying if the case is in Henn or not
				If left(case_pw, 4) = "X127" Then case_is_in_henn = True
				If case_is_in_henn = False Then kick_it_off_reason = "Case not in 27"			'updating detail to record to the case list in SQL the reason the case cannot be processed as ex parte
				If case_active = False Then kick_it_off_reason = "Case not Active"
			End If

			'if we could a reason that we cannot process ex parte, we need to update the SQL list with this information
			If kick_it_off_reason <> "" Then
				select_ex_parte = False
				'This is opening the Ex Parte Case List data to update the information for the case if the case is not ex parte.
				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & select_ex_parte & "', Phase1Complete = '" & kick_it_off_reason & "' WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

				Set objUpdateConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			Else
				'if the case has not been determined to be 'kicked off' we can look closer into the case information

				ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)										'resetting the person array
				Do																				'making sure to get to STAT/MEMB
					Call back_to_SELF
					Call navigate_to_MAXIS_screen("STAT", "MEMB")
					EMReadScreen memb_check, 4, 2, 48
				Loop until memb_check = "MEMB"
				Call get_list_of_members														'this funciton will fill the MEMBER_INFO_ARRAY with person information.

				'Read SVES/TPQY for all persons on a case
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False						'setting the array with these defaults
					MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = False
					MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = False
					memb_has_railroad = False

					Call navigate_to_MAXIS_screen("INFC", "SVES")								'navigate to the SVES interface to read the TPQY response
					EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68
					Call write_value_and_transmit("TPQY", 20, 70)

					Do
						EMReadScreen check_TPQY_panel, 4, 2, 53 								'Reads for TPQY panel to make sure we've made it to the right place
						If check_TPQY_panel <> "TPQY" Then Call write_value_and_transmit("TPQY", 20, 70)
					Loop until check_TPQY_panel = "TPQY"

					'here we see if the tpqy response is from the Prep 2 run.
					'Not all cases have a new TPQY response at this point (phase 1)
					'we only want to read the onces that were sent during Prep 2 (with a secondary RSDI claim number)
					EMReadScreen tpqy_response_date, 8, 7, 22									'Reading the TPQY response date
					tpqy_response_date = trim(tpqy_response_date)
					If tpqy_response_date <> "" Then											'make sure that a response date exists
						tpqy_response_date = replace(tpqy_response_date, " ", "/")				'making the response date a date
						tpqy_response_date = DateAdd("d", 0, tpqy_response_date)

						If DateDiff("d", prep_phase_2_run_date, tpqy_response_date) > 0 Then	'if the response date is after the Prep 2 run date, then the response should be in regards to the Prep 2 QURY


							EMReadScreen tpqy_name_txt, 40, 4, 10								'Read information from TPQY and save it in the member array
							EMReadScreen tpqy_ssn_txt, 11, 5, 9
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb), 		1, 8, 39
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb), 		1, 8, 65
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 	12, 5, 35
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), 		10, 6, 61
							MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb))
							MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb))
							EMReadScreen sves_response, 8, 7, 22 		'Return Date
							sves_response = replace(sves_response," ", "/")						'this records the information from successful qurys into a txt file for the BZ team to review
							objTextStream.WriteLine MAXIS_case_number & "| NAME: " & tpqy_name_txt & "|" & "SSN: " &  tpqy_ssn_txt & "|" & "SSI record - " & MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb) & "|" & "RSDI record - " & MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb) & "| CLAIM NUMB: " & MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb)

							transmit															'go to the next panel
							Do
								EMReadScreen check_BDXP_panel, 4, 2, 53 						'Makes sure we are at the BDXP panel
								If check_BDXP_panel <> "BDXP" Then
									row = 1
									col = 1
									EMSearch "Command:", row, col
									Call write_value_and_transmit("BDXP", row, col++9)
								End If
							Loop until check_BDXP_panel = "BDXP"

							EMReadScreen MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 		12, 5, 69		'Read SVES information
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb), 	2, 6, 19
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), 	8, 8, 16
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb), 		8, 8, 32
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb), 		1, 8, 69
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), 	5, 11, 69
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), 	5, 14, 69
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), 	10, 15, 69
							MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb))
							MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), " ", "")
							MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb))
							' MEMBER_INFO_ARRAY(tpqy_rsdi_staus_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_staus_desc, each_memb))
							' MEMBER_INFO_ARRAY(tpqy_rsdi_paydate, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_paydate, each_memb))
							MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb))
							MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb))
							MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb))
							MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = Trim (MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), " ", "/1/")
							MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), " ", "/1/")
							MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), " ", "/")

							transmit

							Do
								EMReadScreen check_BDXM_panel, 4, 2, 53 						'Reads for BDXM panel
								If check_BDXM_panel <> "BDXM" Then
									row = 1
									col = 1
									EMSearch "Command:", row, col
									Call write_value_and_transmit("BDXM", row, col++9)
								End If
							Loop until check_BDXM_panel = "BDXM"

							EMReadScreen MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb), 			13, 4, 29
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb), 			7, 6, 64
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 				5, 7, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), 				5, 7, 63
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_ind, each_memb), 			1, 8, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb), 			3, 8, 63
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb), 	5, 9, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb), 	5, 9, 63
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb), 			7, 12, 64
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 				5, 13, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), 				5, 13, 63
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_ind, each_memb), 			1, 14, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb), 			3, 14, 63
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), 	5, 15, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb), 	5, 15, 63
							MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb))
							MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb))
							MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb))
							MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb))
							MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb))
							MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb))
							MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb))
							MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb))
							MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb))
							MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), " ", "/01/")
							MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), " ", "/01/")
							MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb), " ", "/01/")
							MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb), " ", "/01/")
							MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), " ", "/01/")
							MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), " ", "/01/")
							MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), " ", "/01/")
							MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb), " ", "/01/")

							transmit

							Do
								EMReadScreen check_SDXE_panel, 4, 2, 53 						'Reads for SDXE panel
								If check_SDXE_panel <> "SDXE" Then
									row = 1
									col = 1
									EMSearch "Command:", row, col
									Call write_value_and_transmit("SDXE", row, col++9)
								End If
							Loop until check_SDXE_panel = "SDXE"

							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb), 		12, 5, 36
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), 		2, 7, 21
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb), 		22, 7, 24
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_fed_living, each_memb), 			1, 6, 70
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 			3, 8, 21
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb), 			30, 8, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_cit_ind_code, each_memb), 			1, 7, 70
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_code, each_memb), 		3, 10, 26
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb), 		40, 10, 30
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb), 		8, 11, 26
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb), 			8, 12, 26
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), 		8, 13, 26
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_code, each_memb), 		1, 11, 65
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb), 		8, 12, 65
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_code, each_memb), 	2, 13, 65
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb), 	8, 14, 65
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_disa_pay_code, each_memb), 		1, 15, 65
							MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb), " ", "/")
							' MsgBox MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb)

							transmit

							Do
								EMReadScreen check_SDXP_panel, 4, 2, 50 							'Reads for SDXP panel
								If check_SDXP_panel <> "SDXP" Then
									row = 1
									col = 1
									EMSearch "Command:", row, col
									Call write_value_and_transmit("SDXP", row, col++9)
								End If
							Loop until check_SDXP_panel = "SDXP"

							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), 			5, 4, 16
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), 			7, 4, 42
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_over_under_code, each_memb), 	1, 4, 73
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb), 	5, 8, 3
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb), 	6, 8, 13
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_type, each_memb), 	1, 8, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb), 	5, 9, 3
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb), 	6, 9, 13
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_type, each_memb), 	1, 9, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb), 	5, 10, 3
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb), 	6, 10, 13
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_type, each_memb), 	1, 10, 25
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb), 				8, 5, 66
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb), 				8, 6, 66
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb), 		8, 7, 66
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb), 		8, 8, 66
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb), 		8, 9, 66
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb), 		8, 10, 66
							EMReadScreen MEMBER_INFO_ARRAY(tpqy_rep_payee, each_memb), 				1, 11, 66
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb))
							If MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "" Then MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "0"
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb))
							MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb))
							MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb))
							MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb))
							MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb))
							MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb))
							MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb))
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), " ", "/01/")
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb), " ", "/")

							transmit

							If MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb) = "Y" Then
								MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True
								MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb)= False
								If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) = "C01" Then MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb) = True
								If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "E" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
								If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "H" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
								If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "M" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
								If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "P" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
								If left(MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 1) = "S" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
							End If
							If MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb) = "Y" Then
								If MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "C" or MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "E" Then
									MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True
									If IsDate(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb)) = True Then MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True
								End If
							End If

							'This is opening the Ex Parte income list to record the QURY information for a particaular case and person - but only if the qury was sent during the prep 2 run
							objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [QURY_Sent] = '" & prep_phase_2_run_date & "'"

							Set objIncomeConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
							Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

							'opening the connections and data table
							objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
							objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection


						End If
					End If
					Call back_to_SELF		'getting all th way back
				Next

				Do
					Call navigate_to_MAXIS_screen("STAT", "SUMM")					'making sure the case has come through background if it went through
					EMReadScreen summ_check, 4, 2, 46
				Loop until summ_check = "SUMM"

				verif_types = ""													'blanking out the verification types for CASE/NOTE
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					'NOTE that there is no SSI information that should be on a secondary claim so it should not be updated during the Phase 1 run (it was updated during the Prep 2 run.)

					'here we need to see what is already listed as income on the RSDI panels that is NOT in the new TPQY.
					'There has been duplication of
					If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
						'First we need to see what the income is listed for any existant RSDI panel
						'here we navigate to the UNEA panel
						EMWriteScreen "UNEA", 20, 71
						transmit
						EMReadScreen unea_check, 4, 2, 48
						Do While unea_check <> "UNEA"								'making sure we've made it.
							Call navigate_to_MAXIS_screen("STAT", "UNEA")
							EMReadScreen unea_check, 4, 2, 48
						Loop
						EMWriteScreen MEMB_reference_number, 20, 76 				'Navigating to the right member of unea
						EMWriteScreen "01", 20, 79 									'to ensure we're on the 1st instance of UNEA panels for the appropriate member
						transmit

						other_rsdi_amount = ""
						EMReadScreen vers_count, 1, 2, 78							'reading the number of versions of UNEA for this member
						If vers_count <> "0" Then									'if there are none, the rest of the functionality will be skipped, if there are any, we have to read them
							Do
								EMReadScreen panel_instance, 1, 2, 73										'reading the specific panel information
								EMReadScreen panel_type_code, 2, 5, 37

								If panel_type_code = "01" or panel_type_code = "02" Then
									EMReadScreen panel_claim_number, 15, 6, 37
									EMReadScreen panel_prosp_amount, 8, 18, 68
									panel_claim_number = replace(panel_claim_number, "_", "")					'formatting the claim number
									panel_claim_number = replace(panel_claim_number, " ", "")

									If left(panel_claim_number, 9) <> left(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 9) Then
										other_rsdi_amount = trim(panel_prosp_amount)
									End If
								End If

								transmit											'go to the next UNEA panel
								EMReadScreen end_of_UNEA_panels, 7, 24, 2			'read the warning/error message to see if we could not move to the next panel
							Loop Until end_of_UNEA_panels = "ENTER A"
						End If

						rsdi_amount = MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb)
						If other_rsdi_amount = MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) Then rsdi_amount = 0
						objTextStream.WriteLine "        | AMOUNT USED IN UPDATING RSDI: " & rsdi_amount & "| RSDI amount listed in TPQY: " &  MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb)

						If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then
							rsdi_type = "01"
							Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
						Else
							rsdi_type = "02"
							Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
						End If
						If RSDI_panel_found = True or rsdi_amount <> 0 Then
							Call update_unea_pane(RSDI_panel_found, rsdi_type, rsdi_amount, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), "", "")
							If InStr(verif_types, "RSDI") = 0 Then verif_types = verif_types & "/RSDI"
						End If
					End If

					'Additional NOTE that there is no need to update MEDI from a secondary claim number either.
				Next

				For each_uc = 0 to UBound(UC_INCOME_ARRAY, 2)
					If UC_INCOME_ARRAY(uc_case_numb_const, each_uc) = MAXIS_case_number Then

						Call navigate_to_MAXIS_screen("STAT", "UNEA")
						EMWriteScreen UC_INCOME_ARRAY(uc_ref_numb_const, each_uc), 20, 76
						EMWriteScreen "01", 20, 79
						transmit

						Do
							EMReadScreen unea_inc_type, 2, 5, 37
							If unea_inc_type = "14" Then
								EMReadScreen unea_claim_number, 15, 6, 37
								unea_claim_number = replace(unea_claim_number, "_", "")

								If unea_claim_number = UC_INCOME_ARRAY(uc_claim_numb_const, each_uc) or UC_INCOME_ARRAY(uc_claim_numb_const, each_uc) = "" Then
									UC_INCOME_ARRAY(uc_panel_updated_const, each_uc) = "YES"
									PF9
									EMWriteScreen "6", 5, 65		'Write Other Verification Code "6"

									EMReadScreen first_pay_day, 8, 13, 54
									If first_pay_day <> "__ __ __" Then
										first_pay_day = replace(first_pay_day, " ", "/")
										first_pay_day = DateAdd("d", 0, first_pay_day)
										day_of_the_week = Weekday(first_pay_day)
									Else
										first_pay_day = ""
										day_of_the_week = 3
									End If

									first_of_cm_plus_1 = CM_plus_1_mo & "/1/" & CM_plus_1_yr
									first_of_cm_plus_1 = DateAdd("d", 0, first_of_cm_plus_1)
									first_of_cm_minus_1 = CM_minus_1_mo & "/1/" & CM_minus_1_yr
									first_of_cm_minus_1 = DateAdd("d", 0, first_of_cm_minus_1)
									Do While Weekday(first_of_cm_plus_1)<> day_of_the_week
										first_of_cm_plus_1 = DateAdd("d", 1, first_of_cm_plus_1)
									Loop
									Do While Weekday(first_of_cm_minus_1)<> day_of_the_week
										first_of_cm_minus_1 = DateAdd("d", 1, first_of_cm_minus_1)
									Loop

									'Clear amounts
									row = 13
									DO
										EMWriteScreen "__", row, 25
										EMWriteScreen "__", row, 28
										EMWriteScreen "__", row, 31
										EMWriteScreen "________", row, 39

										EMWriteScreen "__", row, 54
										EMWriteScreen "__", row, 57
										EMWriteScreen "__", row, 60
										EMWriteScreen "________", row, 68
										row = row + 1
									Loop until row = 18

									If DateDiff("m", UC_INCOME_ARRAY(uc_end_date_const, each_uc), date) >=6 Then
										Call write_value_and_transmit("DEL", 20, 71)
									ElseIf UC_INCOME_ARRAY(uc_end_date_const, each_uc) <> "" Then
										Call create_mainframe_friendly_date(end_date, 7, 68, "YY")	'income end date (SSI: ssi_denial_date, RSDI: susp_term_date)
										Call write_value_and_transmit("X", 6, 56)
										Call clear_line_of_text(9, 65)
										Do
											transmit
											EMReadScreen HC_popup, 9, 7, 41
											' If HC_popup = "HC Income" then transmit
										Loop until HC_popup <> "HC Income"
									Else
										retro_date = first_of_cm_minus_1
										' MsgBox "retro_date - " & retro_date & vbCr & "start_date - " & start_date & vbCr & "DateDiff - " & DateDiff("d", retro_date, start_date)
										'TODO - this retro date thing failed
										EMReadScreen start_date, 8, 7, 37
										start_date = replace(start_date, " ", "/")
										start_date = DateAdd("d", 0, start_date)
										If DateDiff("d", retro_date, start_date) < 0 Then
											row = 13
											Do
												Call create_mainframe_friendly_date(retro_date, row, 25, "YY")
												EMWriteScreen UC_INCOME_ARRAY(uc_prosp_inc_const, each_uc), row, 39		'TODO: Testing values
												retro_date = DateAdd("w", 1, retro_date)
												row = row + 1
											Loop Until DateDiff("m", first_of_cm_minus_1, retro_date) = 1
										End If

										' MsgBox "STOP and look"
										prosp_date = first_of_cm_plus_1
										row = 13
										Do
											Call create_mainframe_friendly_date(prosp_date, row, 54, "YY")
											EMWriteScreen UC_INCOME_ARRAY(uc_prosp_inc_const, each_uc), row, 68		'TODO: Testing values
											prosp_date = DateAdd("w", 1, prosp_date)
											row = row + 1
										Loop Until DateDiff("m", first_of_cm_plus_1, prosp_date) = 1

										EMWriteScreen CM_plus_1_mo, 13, 54 'hardcoded dates
										EMWriteScreen "01", 13, 57
										EMWriteScreen CM_plus_1_yr, 13, 60 'hardcoded dates
										EMWriteScreen income_amount, 13, 68		'TODO: Testing values (income_amt which = rsdi_gross_amt or ssi_gross_amt )

										Call write_value_and_transmit("X", 6, 56)
										Call clear_line_of_text(9, 65)
										EMWriteScreen UC_INCOME_ARRAY(uc_prosp_inc_const, each_uc), 9, 65		'TODO: Testing values (rsdi_gross_amt or ssi_gross_amt )
										EMWriteScreen "4", 10, 63		'code for pay frequency
										Do
											transmit
											EMReadScreen HC_popup, 9, 7, 41
											' If HC_popup = "HC Income" then transmit
										Loop until HC_popup <> "HC Income"

										' MsgBox "STOP AND LOOK AT THE PANEL"
										' PF10

										transmit
										EMReadScreen cola_warning, 29, 24, 2
										If cola_warning = "WARNING: ENTER COLA DISREGARD" then transmit
										EMReadScreen HC_income_warning, 25, 24, 2
										If HC_income_warning = "WARNING: UPDATE HC INCOME" then transmit
										' MsgBox "Wait"
									End If

									If InStr(verif_types, "UC") = 0 Then verif_types = verif_types & "/UC"
									Exit Do
								End If
							End If
							transmit
							EMReadScreen last_unea, 7, 24, 2
						Loop until last_unea = "ENTER A"

					End If
				Next


				For each_va = 0 to UBound(VA_INCOME_ARRAY, 2)
					If VA_INCOME_ARRAY(va_case_numb_const, eacheach_va_uc) = MAXIS_case_number Then

						Call navigate_to_MAXIS_screen("STAT", "UNEA")
						EMWriteScreen VA_INCOME_ARRAY(va_ref_numb_const, each_va), 20, 76
						transmit

						Do
							EMReadScreen unea_inc_type, 2, 5, 37
							If VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = unea_inc_type or (VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = "" and (unea_inc_type = "11" or unea_inc_type = "12" or unea_inc_type = "13" or unea_inc_type = "38")) Then
								If IsNumeric(VA_INCOME_ARRAY(va_prosp_inc_const, va_count)) = True Then
									VA_INCOME_ARRAY(va_panel_updated_const, va_count) = "YES"
									Call update_unea_pane(True, unea_inc_type, VA_INCOME_ARRAY(va_prosp_inc_const, va_count), VA_INCOME_ARRAY(va_claim_numb_const, va_count), "", "", "")
									If InStr(verif_types, "VA") = 0 Then verif_types = verif_types & "/VA"
									Exit Do
								Else
									PF9
									EMWriteScreen "N", 5, 65

									For unea_row = 13 to 17
										EMReadScreen pay_aount, 8, unea_row, 39
										If pay_aount <> "________" Then
											EMWriteScreen CM_minus_1_mo, unea_row, 25
											EMWriteScreen CM_minus_1_yr, unea_row, 31
										End If
									Next
									For unea_row = 13 to 17
										EMReadScreen pay_aount, 8, unea_row, 68
										If pay_aount <> "________" Then
											EMWriteScreen CM_minus_1_mo, unea_row, 54
											EMWriteScreen CM_minus_1_yr, unea_row, 60
										End If
									Next
								End If

							End If
							transmit
							EMReadScreen last_unea, 7, 24, 2
						Loop until last_unea = "ENTER A"
					End If
				Next





				'Send the case through background
				Call write_value_and_transmit("BGTX", 20, 71)
				EMReadScreen wrap_check, 4, 2, 46
				If wrap_check = "WRAP" Then transmit
				EMWaitReady 0, 0												'give a pause here
				EMReadScreen wrap_error, 30, 24, 2
				If wrap_error = "THE COMMAND 'BGTX' NOT ALLOWED" Then transmit
				EMReadScreen database_busy, 23, 4, 44							'Sometimes, when we send a case through background a database record error raises
				If database_busy = "A MAXIS database record" Then transmit  	'we need to transmit past it
				'TODO - there may be a NAT error being raised here, but I don't know what that might be from or if we need to resolve it - there does not seem to be any impact to running the script

				Call back_to_SELF

				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)

					If MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) <> "" Then
						'This is updating the Ex Parte Income list with details of rsdi information from SVES
						objIncUpdtSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response &_
										"', GrossAmt = '" & MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) &_
										"', NetAmt = '" & MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) &_
										"', EndDate = '" & MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) &_
										"' WHERE [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] LIKE '" & left(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 9) & "'"

						Set objIncUpdtConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
						Set objIncUpdtRecordSet = CreateObject("ADODB.Recordset")

						'opening the connections and data table
						objIncUpdtConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncUpdtRecordSet.Open objIncUpdtSQL, objIncUpdtConnection

					End If
				Next


				'CASE/NOTE details of the case information
				If left(verif_types, 1) = "/" Then verif_types = right(verif_types, len(verif_types)-1)
				note_title = "Verification of " & verif_types

				If verif_types <> "" Then
					Call navigate_to_MAXIS_screen("CASE", "NOTE")
					EMReadScreen last_note, 55, 5, 25
					EMReadScreen last_note_date, 8, 5, 6
					today_day = right("0"&DatePart("d", date), 2)
					today_mo = right("0"&DatePart("d", date), 2)
					today_yr = right(DatePart("d", date), 2)
					today_as_text = today_mo & "/" & today_day & "/" & today_yr

					last_note = trim(last_note)

					If last_note <> note_title or last_note_date <> today_as_text Then
						start_a_blank_CASE_NOTE
						Call write_variable_in_CASE_NOTE(note_title)
						For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
							If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True or MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
								Call write_variable_in_CASE_NOTE("Income from SSA for MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & " - " & MEMBER_INFO_ARRAY(memb_name_const, each_memb) & ".")
								If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True Then
									Call write_variable_in_CASE_NOTE(" * SSI Income of $ " & MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) & " per month.")
								End If
								If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
									rsdi_inc = "RSDI"
									If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then rsdi_inc = "RSDI, Disa"
									Call write_variable_in_CASE_NOTE(" * " & rsdi_inc & " Income of $ " & MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) & " per month.")
								End If
								Call write_variable_in_CASE_NOTE("   - Verified through worker initiated data match.")
								Call write_variable_in_CASE_NOTE("   - UNEA panel updated eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
							End If
						Next
						For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
							If MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True or MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True Then
								Call write_variable_in_CASE_NOTE("Medicare for MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & " - " & MEMBER_INFO_ARRAY(memb_name_const, each_memb) & ".")

								If MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True Then
									If MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) <> "" Then
										Call write_variable_in_CASE_NOTE("* Medicare Part A ended " & MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb))
									Else
										Call write_variable_in_CASE_NOTE("* Medicare Part A started " & MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb))
									End If
								End If
								If MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True Then
									If MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) <> "" Then
										Call write_variable_in_CASE_NOTE("* Medicare Part B ended " & MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb))
									Else
										Call write_variable_in_CASE_NOTE("* Medicare Part B started " & MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb))
										If MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) <> "" Then
											Call write_variable_in_CASE_NOTE("  - Part B Premium: $ " &MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb))
										Else
											Call write_variable_in_CASE_NOTE("  - Part B Buy-In Start Date: " & MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb))
										End If

									End If
								End If
								Call write_variable_in_CASE_NOTE("   - Verified through worker initiated data match.")
								Call write_variable_in_CASE_NOTE("   - MEDI panel updated eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
							End If
						Next
						For each_va = 0 to UBound(VA_INCOME_ARRAY, 2)
							If VA_INCOME_ARRAY(uc_case_numb_const, each_va) = MAXIS_case_number Then
								If VA_INCOME_ARRAY(uc_panel_updated_const, each_va) = "YES"	Then
									Call write_variable_in_CASE_NOTE("Income from Unemployment for MEMB " & VA_INCOME_ARRAY(va_ref_numb_const, each_va) & " - " & VA_INCOME_ARRAY(uc_pers_name_const, each_va) & ".")
									Call write_variable_in_CASE_NOTE(" * Income of $ " & VA_INCOME_ARRAY(va_prosp_inc_const, each_va) & " per month.")
									objTextStream.WriteLine MAXIS_case_number & "| VA - MEMB: " & VA_INCOME_ARRAY(va_ref_numb_const, each_va)
								End If
							End If
						Next
						For each_uc = 0 to UBound(UC_INCOME_ARRAY, 2)
							If UC_INCOME_ARRAY(uc_case_numb_const, each_uc) = MAXIS_case_number Then
								If UC_INCOME_ARRAY(uc_panel_updated_const, each_uc) = "YES"	Then
									Call write_variable_in_CASE_NOTE("Income from Unemployment for MEMB " & UC_INCOME_ARRAY(uc_ref_numb_const, each_uc) & " - " & UC_INCOME_ARRAY(uc_pers_name_const, each_uc) & ".")
									Call write_variable_in_CASE_NOTE(" * Income of $ " & UC_INCOME_ARRAY(uc_prosp_inc_const, each_uc) & " per week.")
									objTextStream.WriteLine MAXIS_case_number & "| UC - MEMB: " & UC_INCOME_ARRAY(uc_ref_numb_const, each_uc)
								End If
							End If
						Next
						call write_variable_in_case_note("---")
						call write_variable_in_case_note(worker_signature)
						call write_variable_in_case_note("Automated Update")

					End If
				End If

				'This is opening the Ex Parte Case List data table so we update the progress on the case for phase 1 run completion. This tracks that the work was done for this case in phase 1.
				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase1Complete = '" & date & "' WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

				Set objUpdateConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			End If
		End If
		objRecordSet.MoveNext
	Loop
	'Close the object
	objTextStream.Close

	objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	end_msg = "BULK Phase 1 Run has been completed for " & review_date & "."



	'This is opening the Ex Parte Case List data table so we can loop through it. We just want to see what all happened.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	phase_1_done_count = 0
	today_phase_1_count = 0
	cases_removed_from_ex_parte_in_phase_1 = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1

		If IsDate(objRecordSet("Phase1Complete")) = True Then
			phase_1_done_count = phase_1_done_count + 1
			phase_1_date = objRecordSet("Phase1Complete")
			phase_1_date = DateAdd("d", 0, phase_1_date)
			If phase_1_date = date Then today_phase_1_count = today_phase_1_count + 1
		End If
		If objRecordSet("Phase1Complete") = "Case not in 27" Then cases_removed_from_ex_parte_in_phase_1 = cases_removed_from_ex_parte_in_phase_1 + 1
		If objRecordSet("Phase1Complete") = "Case not Active" Then cases_removed_from_ex_parte_in_phase_1 = cases_removed_from_ex_parte_in_phase_1 + 1
		objRecordSet.MoveNext
	Loop
    objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	'Creating an end message to display the case list counts
	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count & vbCr
	end_msg = end_msg & vbCr & "Cases that completed PREP but are NOT Ex Parte Now: " & cases_removed_from_ex_parte_in_phase_1 & vbCr

	end_msg = end_msg & vbCr & "Cases with Phase 1 Done: " & phase_1_done_count
	end_msg = end_msg & vbCr & "Cases with Phase 1 Done Today: " & today_phase_1_count

	'This is the end of the fucntionality and will just display the end message at the end of this script file.
End If

'This functionality is run on the 1st of the Processing Month.
'Currently it ONLY updates STAT/BUDG to align the budget to the processing requirements and send the case through background
If ex_parte_function = "Phase 2" Then
	review_date = ep_revw_mo & "/1/" & ep_revw_yr			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	review_date = DateAdd("d", 0, review_date)

	'This functionality will remove any 'holds' with 'In Progress' marked. This is to make sure no cases get left behind if a script fails (Unset In Progress)
	If reset_in_Progress = checked Then
		'This is opening the Ex Parte Case List data table so we can loop through it.
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

		Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		'Loop through each item on the CASE LIST Table
		Do While NOT objRecordSet.Eof
			If objRecordSet("Phase2Complete") = "In Progress" Then			'If the case is marked as 'In Progress' - we are going to remove it
				MAXIS_case_number = objRecordSet("CaseNumber") 				'SET THE MAXIS CASE NUMBER

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase2Complete = '" & NULL & "'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"	'removing the 'In Progress' indicator and blanking it out

				Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			End If
			objRecordSet.MoveNext			'move to the next item in the table
		Loop
		objRecordSet.Close			'Closing all the data connections
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
	End If

	'Creating a txt file output of cases in which the BUDG update did not work or there was another problem.
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")			'creating the object to connect with the file
	tracking_doc_file = user_myDocs_folder & "ExParte Tracking Lists/Phase 2 " & ep_revw_mo & "-" & ep_revw_yr & " budg issues list.txt"		'this is the file path
	If ObjFSO.FileExists(tracking_doc_file) Then
		Set objTextStream = ObjFSO.OpenTextFile(tracking_doc_file, ForAppending, true)		'If the file exists we open it and set to add to it
	Else
		Set objTextStream = ObjFSO.CreateTextFile(tracking_doc_file, ForWriting, true)		'If the file does not exists, we create it and set to writing the file
	End If
	objTextStream.WriteLine "LIST START"		'This is going to head each start of the script run.

	review_date = ep_revw_mo & "/1/" & ep_revw_yr			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	review_date = DateAdd("d", 0, review_date)

	'Open The CASE LIST Table
	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Loop through each item on the CASE LIST Table
	Do While NOT objRecordSet.Eof
		'Only select cases that are Ex Parte and where Phase 2 Complete has not been updated
		'NOTE - we have to use NULL and "" because if the 'unset' and In Progress case, it appears as a "" instead of a NULL
		If objRecordSet("SelectExParte") = True and (IsNull(objRecordSet("Phase2Complete")) = True or objRecordSet("Phase2Complete") = "") Then
			'For each case that is indicated as Ex parte, we are going to update the case information
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			'Here we are setting the Phase1Complete to 'In Progress' to hold the case as being worked.
			'This portion of the script is required to be able to have more than one person operating the BULK run at the same time.
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase2Complete = 'In Progress'  WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			'Need to make sure we get to SUMM
			Do
				Call navigate_to_MAXIS_screen("STAT", "SUMM")
				EMReadScreen summ_check, 4, 2, 46
			Loop until summ_check = "SUMM"

			'Calls the function to update the budget panel
			Call update_stat_budg

			'Send the case through background
			Call write_value_and_transmit("BGTX", 20, 71)					'Enter the command to force the case through background
			EMReadScreen wrap_check, 4, 2, 46								'Making sure we are at STAT/WRAP
			If wrap_check = "WRAP" Then transmit							'If we are at WRAP, transmit through
			EMWaitReady 0, 0												'give a pause here
			EMReadScreen database_busy, 23, 4, 44							'Sometimes, when we send a case through background a database record error raises
			If database_busy = "A MAXIS database record" Then transmit  	'we need to transmit past it
			'TODO - there may be a NAT error being raised here, but I don't know what that might be from or if we need to resolve it - there does not seem to be any impact to running the script
			Call back_to_SELF												'Need to get to SELF

			'here is the update statement. setting the Phase2 BULK run completion date for the case running
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase2Complete = '" & date & "' WHERE CaseNumber = '" & MAXIS_case_number & "' and HCEligReviewDate = '" & review_date & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

		End If
		objRecordSet.MoveNext			'now we go to the next case
	Loop

	objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	'We are going to set the display message for the end of the script run
	end_msg = "BULK Phase 2 Run has been completed for " & review_date & "."

	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	phase_2_done_count = 0
	today_phase_2_count = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1
		If IsDate(objRecordSet("Phase2Complete")) = True Then
			phase_2_done_count = phase_2_done_count + 1
			phase_2_date = objRecordSet("Phase2Complete")
			phase_2_date = DateAdd("d", 0, phase_2_date)
			If phase_2_date = date Then today_phase_2_count = today_phase_2_count + 1
		End If
		objRecordSet.MoveNext
	Loop
    objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	'Creating an end message to display the case list counts
	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count & vbCr
	end_msg = end_msg & vbCr & "Cases with Phase 2 Done: " & phase_2_done_count
	end_msg = end_msg & vbCr & "Cases with Phase 2 Done Today: " & today_phase_2_count

	'This is the end of the fucntionality and will just display the end message at the end of this script file.
End If

If ex_parte_function = "Check REVW information on Phase 1 Cases" Then
	'This should be run for cases at Phase 1 only after ER cutoff
	'Create a spreadsheet and pull all cases in the data table for CM + 2 review into the list
	'Opening a spreadsheet to capture the cases with a SMRT ending soon
	Set ObjExcel = CreateObject("Excel.Application")
	ObjExcel.Visible = True
	Set objSMRTWorkbook = ObjExcel.Workbooks.Add()
	ObjExcel.DisplayAlerts = True

	'Setting the first 4 col as worker, case number, name, and APPL date
	ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
	ObjExcel.Cells(1, 2).Value = "Ex Parte WORKER"
	ObjExcel.Cells(1, 3).Value = "Select Ex Parte"
	ObjExcel.Cells(1, 4).Value = "Phase 1 Ex Parte Eval"
	ObjExcel.Cells(1, 5).Value = "Phase 1 Notes"

	ObjExcel.Cells(1, 6).Value = "REVW HC ER"
	ObjExcel.Cells(1, 7).Value = "REVW Status"
	ObjExcel.Cells(1, 8).Value = "Ex Parte Ind"
	ObjExcel.Cells(1, 9).Value = "Ex Parte REVW Mo"

	FOR i = 1 to 9		'formatting the cells'
		ObjExcel.Cells(1, i).Font.Bold = True		'bold font'
	NEXT


	'Capture the Ex Parte information from the table into the excel.
	excel_row = 2		'initializing the counter to move through the excel lines
	review_date = ep_revw_mo & "/1/" & ep_revw_yr
	review_date = DateAdd("d", 0, review_date)



	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	Do While NOT objRecordSet.Eof 					'Loop through each item on the CASE LIST Table
		MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER
		Do
			If left(MAXIS_case_number, 1) = "0" Then MAXIS_case_number = right(MAXIS_case_number, len(MAXIS_case_number)-1)
		Loop until left(MAXIS_case_number, 1) <> "0"

		ObjExcel.Cells(excel_row, 1).Value = MAXIS_case_number
		ObjExcel.Cells(excel_row, 2).Value = objRecordSet("Phase1HSR")
		ObjExcel.Cells(excel_row, 3).Value = objRecordSet("SelectExParte")
		ObjExcel.Cells(excel_row, 4).Value = objRecordSet("ExParteAfterPhase1")
		ObjExcel.Cells(excel_row, 5).Value = objRecordSet("Phase1ExParteCancelReason")
		excel_row = excel_row + 1

		objRecordSet.MoveNext			'now we go to the next case
	Loop
	objRecordSet.Close			'Closing all the data connections
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing

	'Run to REPT/REVS for CM+2
	back_to_self    'We need to get back to SELF and manually update the footer month
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)

	Call navigate_to_MAXIS_screen("REPT", "REVS")
	EMWriteScreen ep_revw_mo, 20, 55
	EMWriteScreen ep_revw_yr, 20, 58
	transmit

	'Pull all REVS cases into an array
	const case_num_const 			= 0
	const hc_revw_status_const		= 1
	const hc_revw_er_month_const	= 2
	const hc_revw_ex_parte_yn_const	= 3
	const hc_revw_ex_parte_mo_const	= 4
	const hc_on_revs_const			= 5
	const case_found_on_sql			= 6
	const last_expt_const			= 7

	Dim EX_PARTE_REVW_INFO_ARRAY()
	ReDim EX_PARTE_REVW_INFO_ARRAY(last_expt_const, 0)

	case_count = 0

	'start of the FOR...next loop
	For each worker in worker_array
		worker = trim(worker)
		If worker = "" then exit for
		Call write_value_and_transmit(worker, 21, 6)   'writing in the worker number in the correct col

		'Grabbing case numbers from REVS for requested worker
		DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
			row = 7	'Setting or resetting this to look at the top of the list
			DO		'All of this loops until row = 19
				'Reading case information (case number, SNAP status, and cash status)
				EMReadScreen MAXIS_case_number, 8, row, 6
				MAXIS_case_number = trim(MAXIS_case_number)
				EmReadscreen HC_status, 1, row, 49

				'Navigates though until it runs out of case numbers to read
				IF MAXIS_case_number = "" then exit do

				'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
				If HC_status = "-" 		then HC_status = ""

				If HC_status <> "" Then
					ReDim Preserve EX_PARTE_REVW_INFO_ARRAY(last_expt_const, case_count)

					EX_PARTE_REVW_INFO_ARRAY(case_num_const, case_count) = MAXIS_case_number
					EX_PARTE_REVW_INFO_ARRAY(hc_revw_status_const, case_count) = HC_status
					EX_PARTE_REVW_INFO_ARRAY(hc_on_revs_const, case_count) = True
					EX_PARTE_REVW_INFO_ARRAY(case_found_on_sql, case_count) = False

					case_count = case_count + 1
				End if

				row = row + 1    'On the next loop it must look to the next row
				MAXIS_case_number = "" 'Clearing variables before next loop
			Loop until row = 19		'Last row in REPT/REVS
			'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
			'if max reviews are reached, the goes to next worker is applicable
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	next

	'navigate_to STAT to gather REVW information
	For revs_case = 0 to UBound(EX_PARTE_REVW_INFO_ARRAY, 2)
		MAXIS_case_number = EX_PARTE_REVW_INFO_ARRAY(case_num_const, revs_case)
		Call navigate_to_MAXIS_screen("STAT", "REVW")
		Call write_value_and_transmit("X", 5, 71)
		EMReadScreen HC_ER_Date, 8, 8, 27
		EMReadScreen ExPte_Ind, 1, 9, 27
		EMReadScreen ExPte_Mo, 7, 9, 71

		EX_PARTE_REVW_INFO_ARRAY(hc_revw_er_month_const, revs_case) = replace(HC_ER_Date, " ", "/")
		EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_yn_const, revs_case) = ExPte_Ind
		EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_mo_const, revs_case) = replace(ExPte_Mo, " ", "/")

		PF3
		Call back_to_SELF
	Next

	'Match the array cases to the ones on Excel and output the renewal information
	For xl_row = 2 to excel_row-1
		MAXIS_case_number = trim(ObjExcel.Cells(xl_row, 1).Value)
		case_found_on_revw = False
		For revs_case = 0 to UBound(EX_PARTE_REVW_INFO_ARRAY, 2)
			If MAXIS_case_number = EX_PARTE_REVW_INFO_ARRAY(case_num_const, revs_case) Then
				case_found_on_revw = True
				EX_PARTE_REVW_INFO_ARRAY(case_found_on_sql, revs_case) = True

				ObjExcel.Cells(xl_row, 6).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_er_month_const, revs_case)
				ObjExcel.Cells(xl_row, 7).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_status_const, revs_case)
				ObjExcel.Cells(xl_row, 8).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_yn_const, revs_case)
				ObjExcel.Cells(xl_row, 9).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_mo_const, revs_case)

			End If
		Next

		'navigate to STAT/REVW for any case that was not on the list.
		If case_found_on_revw = False Then
			Call navigate_to_MAXIS_screen("STAT", "REVW")
			EMReadScreen hc_revw_status, 1, 7, 73
			Call write_value_and_transmit("X", 5, 71)
			EMReadScreen HC_ER_Date, 8, 8, 27
			EMReadScreen ExPte_Ind, 1, 9, 27
			EMReadScreen ExPte_Mo, 7, 9, 71

			ObjExcel.Cells(xl_row, 6).Value = replace(HC_ER_Date, " ", "/")
			ObjExcel.Cells(xl_row, 7).Value = hc_revw_status
			ObjExcel.Cells(xl_row, 8).Value = ExPte_Ind
			ObjExcel.Cells(xl_row, 9).Value = replace(ExPte_Mo, " ", "/")

			PF3
			Call back_to_SELF
		End If
	Next
	'Add any cases that are in the REVS array to Excel if they were not already there
	For revs_case = 0 to UBound(EX_PARTE_REVW_INFO_ARRAY, 2)
		If EX_PARTE_REVW_INFO_ARRAY(case_found_on_sql, revs_case) = False Then
			ObjExcel.Cells(excel_row, 1).Value = EX_PARTE_REVW_INFO_ARRAY(case_num_const, revs_case)

			ObjExcel.Cells(excel_row, 6).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_er_month_const, revs_case)
			ObjExcel.Cells(excel_row, 7).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_status_const, revs_case)
			ObjExcel.Cells(excel_row, 8).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_yn_const, revs_case)
			ObjExcel.Cells(excel_row, 9).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_mo_const, revs_case)
			excel_row = excel_row + 1
		End If
	Next

	For col_to_autofit = 1 to 9
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next

	objExcel.ActiveSheet.ListObjects.Add(xlSrcRange, objExcel.Range("A1:I" & excel_row - 1), xlYes).Name = "Table1"
	objExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
	objExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\Phase 1 REVS Check - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"

	Call script_end_procedure("We have a list of HC REVWs for " & ep_revw_mo & "/" & ep_revw_yr & ".")
End If

If ex_parte_function = "DHS Data Validation" Then
	data_sheet_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte\DHS Lists\DHS " & ep_revw_mo & ep_revw_yr & " List.xlsx"

	call excel_open(data_sheet_file_path, True, True, ObjExcel, objWorkbook)
	list_of_all_the_cases = " "
	run_time_timer = timer
	ObjExcel.worksheets("Case Numbers").Activate

	skip_finding_case_in_sql = False
	skip_finding_hc_elig_datils = False
	skip_adding_missing_cases = False
	row_to_start_with = 4

	If sql_reviewed_checkbox = checked Then skip_finding_case_in_sql = True
	If hc_elig_reviewed_checkbox = checked Then skip_finding_hc_elig_datils = True
	If missing_cases_added_checkbox = checked Then skip_adding_missing_cases = True

	excel_starting_row = trim(excel_starting_row)
	If IsNumeric(excel_starting_row) = True Then
		If excel_starting_row > 4 Then row_to_start_with = excel_starting_row
	End If

	If skip_finding_case_in_sql = False Then

		'First we run through the existing cases on the list - these are all from the DHS list
		excel_row = row_to_start_with
		Do
			MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)
			MAXIS_case_number = right("00000000" & MAXIS_case_number, 8)
			If timer - run_time_timer  >= 720 Then
				Call navigate_to_MAXIS_screen("STAT", "ADDR")
				call back_to_SELF
				run_time_timer = timer
			End If
			ObjExcel.Cells(excel_row, 2).Value = True

			'declare the SQL statement that will query the database
			objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '20" & ep_revw_yr & "-" & ep_revw_mo & "-01'"
			' objSQL = "SELECT * FROM ES.ES_ExParte_CaseList"

			'Creating objects for Access
			Set objConnection = CreateObject("ADODB.Connection")
			Set objRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
			objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objRecordSet.Open objSQL, objConnection
			found_on_sql = False

			Do While NOT objRecordSet.Eof
				sql_case_number = objRecordSet("CaseNumber")
				If MAXIS_case_number = sql_case_number Then
					found_on_sql = True
					ObjExcel.Cells(excel_row, 15).Value = True
					ObjExcel.Cells(excel_row, 16).Value = objRecordSet("SelectExParte")

					Exit Do
				End If
				objRecordSet.MoveNext
			Loop

			objRecordSet.Close			'Closing all the data connections
			objConnection.Close
			Set objRecordSet=nothing
			Set objConnection=nothing
			If found_on_sql = False Then ObjExcel.Cells(excel_row, 15).Value = False

			excel_row = excel_row + 1
			next_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)

		Loop Until next_case_number = ""
	End If

	PMI_01 = ""
	name_01 = ""
	person_01_ref_number = ""
	MAXIS_MA_prog_01 = ""
	MAXIS_MA_basis_01 = ""
	MAXIS_msp_prog_01 = ""
	name_02 = ""
	PMI_02 = ""
	person_02_ref_number = ""
	MAXIS_MA_prog_02 = ""
	MAXIS_MA_basis_02 = ""
	MAXIS_msp_prog_02 = ""

	If skip_finding_hc_elig_datils = False Then
		excel_row = row_to_start_with
		Do
			MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)
			MAXIS_case_number = right("00000000" & MAXIS_case_number, 8)
			list_of_all_the_cases = list_of_all_the_cases & MAXIS_case_number & " "

			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, 15).Value, on_henn_list)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, 16).Value, henn_appears_ex_parte)
			list_mx_maj_prog = trim(ObjExcel.Cells(excel_row, 7).Value)
			list_mx_msp = trim(ObjExcel.Cells(excel_row, 9).Value)

			If on_henn_list = True and henn_appears_ex_parte = True and list_mx_maj_prog = "" and list_mx_msp = "" Then

				objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & MAXIS_case_number & "'"

				'Creating objects for Access
				Set objELIGConnection = CreateObject("ADODB.Connection")
				Set objELIGRecordSet = CreateObject("ADODB.Recordset")

				'This is the file path for the statistics Access database.
				' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
				objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objELIGRecordSet.Open objELIGSQL, objELIGConnection

				Do While NOT objELIGRecordSet.Eof

					If name_01 = "" Then
						name_01 = trim(objELIGRecordSet("Name"))
						PMI_01 = trim(objELIGRecordSet("PMINumber"))

						If objELIGRecordSet("MajorProgram") = "QM" or objELIGRecordSet("MajorProgram") = "SL"  Then
							MAXIS_msp_prog_01 = objELIGRecordSet("MajorProgram")
							MAXIS_msp_basis_01 = objELIGRecordSet("EligType")
						ElseIf objELIGRecordSet("MajorProgram") <> "MA" or objELIGRecordSet("MajorProgram") <> "IM" or objELIGRecordSet("MajorProgram") <> "EH" or objELIGRecordSet("MajorProgram") <> "NM"  Then
							MAXIS_MA_prog_01 = objELIGRecordSet("MajorProgram")
							MAXIS_MA_basis_01 = objELIGRecordSet("EligType")
						End If
					ElseIf PMI_01 = trim(objELIGRecordSet("PMINumber")) Then
						If objELIGRecordSet("MajorProgram") = "QM" or objELIGRecordSet("MajorProgram") = "SL"  Then
							MAXIS_msp_prog_01 = objELIGRecordSet("MajorProgram")
							MAXIS_msp_basis_01 = objELIGRecordSet("EligType")
						ElseIf objELIGRecordSet("MajorProgram") <> "MA" or objELIGRecordSet("MajorProgram") <> "IM" or objELIGRecordSet("MajorProgram") <> "EH" or objELIGRecordSet("MajorProgram") <> "NM"  Then
							MAXIS_MA_prog_01 = objELIGRecordSet("MajorProgram")
							MAXIS_MA_basis_01 = objELIGRecordSet("EligType")
						End If
					ElseIf name_02 = "" Then
						name_02 = trim(objELIGRecordSet("Name"))
						PMI_02 = trim(objELIGRecordSet("PMINumber"))

						If objELIGRecordSet("MajorProgram") = "QM" or objELIGRecordSet("MajorProgram") = "SL"  Then
							MAXIS_msp_prog_02 = objELIGRecordSet("MajorProgram")
							MAXIS_msp_basis_02 = objELIGRecordSet("EligType")
						ElseIf objELIGRecordSet("MajorProgram") <> "MA" or objELIGRecordSet("MajorProgram") <> "IM" or objELIGRecordSet("MajorProgram") <> "EH" or objELIGRecordSet("MajorProgram") <> "NM"  Then
							MAXIS_MA_prog_02 = objELIGRecordSet("MajorProgram")
							MAXIS_MA_basis_02 = objELIGRecordSet("EligType")
						End If
					ElseIf PMI_02 = trim(objELIGRecordSet("PMINumber")) Then
						If objELIGRecordSet("MajorProgram") = "QM" or objELIGRecordSet("MajorProgram") = "SL"  Then
							MAXIS_msp_prog_02 = objELIGRecordSet("MajorProgram")
							MAXIS_msp_basis_02 = objELIGRecordSet("EligType")
						ElseIf objELIGRecordSet("MajorProgram") <> "MA" or objELIGRecordSet("MajorProgram") <> "IM" or objELIGRecordSet("MajorProgram") <> "EH" or objELIGRecordSet("MajorProgram") <> "NM"  Then
							MAXIS_MA_prog_02 = objELIGRecordSet("MajorProgram")
							MAXIS_MA_basis_02 = objELIGRecordSet("EligType")
						End If
					End If
					objELIGRecordSet.MoveNext
				Loop

				ObjExcel.Cells(excel_row, 18).Value = PMI_01
				ObjExcel.Cells(excel_row, 19).Value = MAXIS_MA_prog_01
				ObjExcel.Cells(excel_row, 20).Value = MAXIS_MA_basis_01
				ObjExcel.Cells(excel_row, 21).Value = MAXIS_msp_prog_01
				ObjExcel.Cells(excel_row, 23).Value = PMI_02
				ObjExcel.Cells(excel_row, 24).Value = MAXIS_MA_prog_02
				ObjExcel.Cells(excel_row, 25).Value = MAXIS_MA_basis_02
				ObjExcel.Cells(excel_row, 26).Value = MAXIS_msp_prog_02

				objELIGRecordSet.Close
				objELIGConnection.Close
				Set objELIGRecordSet=nothing
				Set objELIGConnection=nothing


				PMI_01 = ""
				name_01 = ""
				person_01_ref_number = ""
				MAXIS_MA_prog_01 = ""
				MAXIS_MA_basis_01 = ""
				MAXIS_msp_prog_01 = ""
				name_02 = ""
				PMI_02 = ""
				person_02_ref_number = ""
				MAXIS_MA_prog_02 = ""
				MAXIS_MA_basis_02 = ""
				MAXIS_msp_prog_02 = ""


				Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

				ObjExcel.Cells(excel_row, 3).Value = ma_status
				ObjExcel.Cells(excel_row, 4).Value = msp_status

				Call navigate_to_MAXIS_screen("ELIG", "HC  ")
				hc_row = 8
				Do
					pers_type = ""
					std = ""
					meth = ""
					' elig_result = ""
					' results_created = ""
					waiv = ""
					EMReadScreen read_ref_numb, 2, hc_row, 3
					EMReadScreen clt_hc_prog, 4, hc_row, 28
					prev_row = hc_row
					Do while read_ref_numb = "  "
						prev_row = prev_row - 1
						EMReadScreen read_ref_numb, 2, prev_row, 3
					Loop


					clt_hc_prog = trim(clt_hc_prog)
					If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "" Then


						Call write_value_and_transmit("X", hc_row, 26)
						If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then
							elig_msp_prog = clt_hc_prog
							EMReadScreen pers_type, 2, 6, 56
						Else
							col = 19
							Do									'Finding the current month in elig to get the current elig type
								EMReadScreen span_month, 2, 6, col
								EMReadScreen span_year, 2, 6, col+3

								If span_month = MAXIS_footer_month and span_year = MAXIS_footer_year Then		'reading the ELIG TYPE
									EMReadScreen pers_type, 2, 12, col - 2
									EMReadScreen std, 1, 12, col + 3
									EMReadScreen meth, 1, 13, col + 2
									EMReadScreen waiv, 1, 17, col + 2
									Exit Do
								End If
								col = col + 11
							Loop until col = 85
							If col = 85 Then
								Do
									col = col - 11
									EMReadScreen pers_type, 2, 12, col - 2
								Loop until pers_type <> "__" and pers_type <> "  "
							End If

						End If
						PF3

						If person_01_ref_number = "" Or person_01_ref_number = read_ref_numb Then
							person_01_ref_number = read_ref_numb
						ElseIf person_02_ref_number = "" Then
							person_02_ref_number = read_ref_numb
						End If


						If person_01_ref_number = read_ref_numb Then
							If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then
								MAXIS_msp_prog_01 = clt_hc_prog
								MAXIS_msp_basis_01 = pers_type
							Else
								MAXIS_MA_prog_01 = clt_hc_prog
								MAXIS_MA_basis_01 = pers_type
							End If
						ElseIf person_02_ref_number = read_ref_numb Then
							If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then
								MAXIS_msp_prog_02 = clt_hc_prog
								MAXIS_msp_basis_02 = pers_type
							Else
								MAXIS_MA_prog_02 = clt_hc_prog
								MAXIS_MA_basis_02 = pers_type
							End If
						End If
						' MsgBox "person_01_ref_number - " & person_01_ref_number & vbCr &_
						' 		"MAXIS_MA_prog_01 - " & MAXIS_MA_prog_01 & vbCr &_
						' 		"MAXIS_MA_basis_01 - " & MAXIS_MA_basis_01 & vbCr &_
						' 		"MAXIS_msp_prog_01 - " & MAXIS_msp_prog_01 & vbCr & vbCR &_
						' 		"person_02_ref_number - " & person_02_ref_number & vbCr &_
						' 		"MAXIS_MA_prog_02 - " & MAXIS_MA_prog_02 & vbCr &_
						' 		"MAXIS_MA_basis_02 - " & MAXIS_MA_basis_02 & vbCr &_
						' 		"MAXIS_msp_prog_02 - " & MAXIS_msp_prog_02
					End If
					hc_row = hc_row + 1
					EMReadScreen next_ref_numb, 2, hc_row, 3
					EMReadScreen next_maj_prog, 4, hc_row, 28
				Loop until next_ref_numb = "  " and next_maj_prog = "    "


				CALL back_to_SELF()
				If person_01_ref_number <> "" Then
					CALL navigate_to_MAXIS_screen("STAT", "MEMB")
					Do
						EMReadScreen read_ref_number, 2, 4, 33
						EMReadscreen last_name, 25, 6, 30
						EMReadscreen first_name, 12, 6, 63
						last_name = trim(replace(last_name, "_", "")) & " "
						first_name = trim(replace(first_name, "_", "")) & " "
						If read_ref_number = person_01_ref_number Then
							EMReadScreen PMI_01, 8, 4, 46
							PMI_01 = trim(PMI_01)
							PMI_01 = right("00000000" & PMI_01, 8)
							name_01 = first_name & " " & last_name
						End If
						If read_ref_number = person_02_ref_number Then
							EMReadScreen PMI_02, 8, 4, 46
							PMI_02 = trim(PMI_02)
							PMI_02 = right("00000000" & PMI_02, 8)
							name_02 = first_name & " " & last_name
						End If
						transmit
						EMReadScreen MEMB_end_check, 13, 24, 2
						' MsgBox "PMI_01 - " & PMI_01 & vbCr & "PMI_02 - " & PMI_02 & vbCr & "MEMB_end_check - " & MEMB_end_check
					LOOP Until PMI_01 <> "" AND (PMI_02 <> "" OR MEMB_end_check = "ENTER A VALID")
				End If

				ObjExcel.Cells(excel_row, 5).Value = person_01_ref_number
				ObjExcel.Cells(excel_row, 6).Value = PMI_01
				ObjExcel.Cells(excel_row, 7).Value = MAXIS_MA_prog_01
				ObjExcel.Cells(excel_row, 8).Value = MAXIS_MA_basis_01
				ObjExcel.Cells(excel_row, 9).Value = MAXIS_msp_prog_01
				ObjExcel.Cells(excel_row, 10).Value = person_02_ref_number
				ObjExcel.Cells(excel_row, 11).Value = PMI_02
				ObjExcel.Cells(excel_row, 12).Value = MAXIS_MA_prog_02
				ObjExcel.Cells(excel_row, 13).Value = MAXIS_MA_basis_02
				ObjExcel.Cells(excel_row, 14).Value = MAXIS_msp_prog_02

				PMI_01 = ""
				name_01 = ""
				person_01_ref_number = ""
				MAXIS_MA_prog_01 = ""
				MAXIS_MA_basis_01 = ""
				MAXIS_msp_prog_01 = ""
				name_02 = ""
				PMI_02 = ""
				person_02_ref_number = ""
				MAXIS_MA_prog_02 = ""
				MAXIS_MA_basis_02 = ""
				MAXIS_msp_prog_02 = ""
			End If

			excel_row = excel_row + 1
			next_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)

		Loop Until next_case_number = ""
	End If

	If skip_adding_missing_cases = False Then
		excel_row = 2
		list_of_all_the_cases = ""
		Do
			MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)
			MAXIS_case_number = right("00000000" & MAXIS_case_number, 8)
			list_of_all_the_cases = list_of_all_the_cases & MAXIS_case_number & " "

			excel_row = excel_row + 1
			next_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)
		Loop Until next_case_number = ""

		'declare the SQL statement that will query the database
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '20" & ep_revw_yr & "-" & ep_revw_mo & "-01'"
		' objSQL = "SELECT * FROM ES.ES_ExParte_CaseList"

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'This is the file path for the statistics Access database.
		' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		Do While NOT objRecordSet.Eof
			sql_case_number = objRecordSet("CaseNumber")
			If Instr(list_of_all_the_cases, sql_case_number) = 0 Then
			' If MAXIS_case_number = sql_case_number Then
				sql_appears_ex_parte = False
				' found_on_sql = True
				ObjExcel.Cells(excel_row, 1).Value = sql_case_number
				ObjExcel.Cells(excel_row, 2).Value = False
				ObjExcel.Cells(excel_row, 15).Value = True
				sql_prep_complete = objRecordSet("PREP_Complete")
				If IsDate(sql_prep_complete) = True Then sql_appears_ex_parte = True
				ObjExcel.Cells(excel_row, 16).Value = sql_appears_ex_parte
				excel_row = excel_row + 1

			End If
			objRecordSet.MoveNext
		Loop
		' ObjExcel.Cells(excel_row, 15).Value = found_on_sql
	End If
	end_msg = "DHS Compare is Completed."

End If

If ex_parte_function = "Evaluate DHS Error List" Then
	'THIS FUNCTIONALITY NEEDS TO BE REVIEWED BEFORE USE AND IS FOR A DHS ERROR LIST
	original_user = windows_user_ID

	'dialog to open a file
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 381, 115, "Ex Perte Error Information"
		Text 10, 10, 350, 20, "This functionality will gather information from the Ex Parte Case list data table based on an Excel file and using the case number."
		Text 10, 40, 365, 20, "Select the file you need information added to. The script will ask you to select the Case Number Column and will add any selected column to the right of the data."
		Text 10, 70, 70, 10, "Select an Excel file:"
		EditBox 80, 65, 245, 15, ex_parte_error_list_excel_file_path
		ButtonGroup ButtonPressed
			PushButton 330, 65, 45, 15, "Browse...", select_a_file_button
			OkButton 270, 95, 50, 15
			CancelButton 325, 95, 50, 15
	EndDialog

	'Show initial dialog
	Do
		Do
			Dialog Dialog1
			cancel_without_confirmation
			If ButtonPressed = select_a_file_button then call file_selection_system_dialog(ex_parte_error_list_excel_file_path, ".xlsx")
		Loop until ButtonPressed = OK and ex_parte_error_list_excel_file_path <> ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in

	'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
	call excel_open(ex_parte_error_list_excel_file_path, True, True, ObjExcel, objWorkbook)

	'Finding all of the worksheets available in the file. We will likely open up the main 'Review Report' so the script will default to that one.
	For Each objWorkSheet In objWorkbook.Worksheets
		scenario_list = scenario_list & chr(9) & objWorkSheet.Name
	Next
	scenario_dropdown = report_date & " Review Report"

	'Dialog to select worksheet
	'DIALOG is defined here so that the dropdown can be populated with the above code
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 191, 60, "Select the Worksheet"
		DropListBox 5, 20, 180, 45, "Select One..." & scenario_list, scenario_dropdown
		ButtonGroup ButtonPressed
			OkButton 80, 40, 50, 15
			CancelButton 135, 40, 50, 15
		Text 5, 10, 155, 10, "Select the correct worksheet from the error list:"
	EndDialog


	'Shows the dialog to select the correct worksheet
	Do
		Do
			Dialog Dialog1
			cancel_without_confirmation
		Loop until scenario_dropdown <> "Select One..."
		call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	'Activates worksheet based on user selection
	objExcel.worksheets(scenario_dropdown).Activate

	'dialog to select the case number column and the excel row to start in
	col_to_use = 1
	column_options = "Select One..."
	Do
		col_header = trim(ObjExcel.Cells(1, col_to_use).Value)
		col_letter = convert_digit_to_excel_column(col_to_use)

		column_options = column_options+chr(9)+col_letter & " - " & col_header

		col_to_use = col_to_use + 1
		next_col_header = trim(ObjExcel.Cells(1, col_to_use).Value)
	Loop until next_col_header = ""

	excel_row_to_start = "2"

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 386, 230, "Select the Case Number and Information"
		Text 5, 10, 155, 10, "Select the column the Case Number is listed in:"
		DropListBox 165, 5, 130, 45, column_options, case_number_column
		Text 100, 30, 65, 10, "Excel Row to start:"
		EditBox 165, 25, 30, 15, excel_row_to_start
		GroupBox 10, 45, 370, 160, "Check all the SQL Column Information to Collect"
		CheckBox 20, 65, 115, 10, "Worker ID", worker_id_checkbox
		CheckBox 140, 65, 115, 10, "HC ER Date", hc_er_date_checkbox
		CheckBox 260, 65, 115, 10, "Select Ex Parte", select_ex_parte_checkbox
		CheckBox 20, 80, 115, 10, "PREP Complete", prep_complete_checkbox
		CheckBox 140, 80, 115, 10, "Phase 1 Complete", phase_1_complete_checkbox
		CheckBox 260, 80, 115, 10, "Phase 1 HSR", phase_1_HSR_checkbox
		CheckBox 20, 95, 115, 10, "Ex Parte after Phase 1", ex_parte_after_phase_1_checkbox
		CheckBox 140, 95, 115, 10, "Phase 1 Cancel Reason", phase_1_ex_parte_cancel_checkbox
		CheckBox 260, 95, 115, 10, "Phase 2 Complete", phase_2_complete_checkbox
		CheckBox 20, 110, 115, 10, "Phse 2 HSR", phase_2_hsr_checkbox
		CheckBox 140, 110, 115, 10, "Ex parte after Phase 2", ex_parte_after_phase_2_checkbox
		CheckBox 260, 110, 115, 10, "Phase 2 Cancel Reason", phase_2_ex_parte_cancel_checkbox
		CheckBox 20, 125, 115, 10, "All HC is ABD", all_hc_is_abd_checkbox
		CheckBox 140, 125, 115, 10, "SSA Income", ssa_income_checkbox
		CheckBox 260, 125, 115, 10, "Wages Income", wages_income_checkbox
		CheckBox 20, 140, 115, 10, "VA Income", va_income_checkbox
		CheckBox 140, 140, 115, 10, "Self Emp Income", self_emp_income_checkbox
		CheckBox 260, 140, 115, 10, "No Income", no_income_checkbox
		CheckBox 20, 155, 115, 10, "EPD on Case", epd_on_case_checkbox
		CheckBox 140, 155, 115, 10, "Year Month", year_month_checkbox
		CheckBox 260, 155, 115, 10, "Eval Year Month", eval_year_month_checkbox
		CheckBox 20, 170, 115, 10, "Approval Year Month", approval_year_month_checkbox
		Text 15, 190, 305, 10, "Each SQL Data Information selected will be added as a new column on the Excel Error List."
		ButtonGroup ButtonPressed
			OkButton 275, 210, 50, 15
			CancelButton 330, 210, 50, 15
	EndDialog

	'Shows the dialog to select the correct worksheet
	Do
		Do
			Dialog Dialog1
			cancel_without_confirmation
		Loop until case_number_column <> "Select One..." and IsNumeric(excel_row_to_start) = True
		call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

	case_number_col_letter = left(case_number_column, 2)
	case_number_col_letter = trim(case_number_col_letter)

	case_number_column = Instr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", case_number_col_letter)

	If worker_id_checkbox = checked Then
		ObjExcel.Cells(1, col_to_use).Value = "Worker ID"
		worker_id_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If hc_er_date_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "HC Elig Review Date"
		hc_er_date_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If select_ex_parte_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Select ExParte"
		select_ex_parte_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If prep_complete_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "PREP Complete"
		prep_complete_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If phase_1_complete_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Phase 1 Complete"
		phase_1_complete_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If phase_1_HSR_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Phase 1 HSR"
		phase_1_HSR_col = col_to_use
		col_to_use = col_to_use + 1
    	ObjExcel.Cells(1, col_to_use).Value = "Phase 1 HSR Name"
		phase_1_hsr_name_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If ex_parte_after_phase_1_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Ex Parte After Phase 1"
		ex_parte_after_phase_1_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If phase_1_ex_parte_cancel_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Phase 1 Ex Parte Cancel Reason"
		phase_1_ex_parte_cancel_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If phase_2_complete_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Phase 2 Complete"
		phase_2_complete_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If phase_2_hsr_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Phase 2 HSR"
		phase_2_hsr_col = col_to_use
		col_to_use = col_to_use + 1
    	ObjExcel.Cells(1, col_to_use).Value = "Phase 2 HSR Name"
		phase_2_hsr_name_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If ex_parte_after_phase_2_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Ex Parte After Phase 2"
		ex_parte_after_phase_2_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If phase_2_ex_parte_cancel_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Phase 2 Ex Parte Cancel Reason"
		phase_2_ex_parte_cancel_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If all_hc_is_abd_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "All HC is ABD"
		all_hc_is_abd_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If ssa_income_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "SSA Income Exist"
		ssa_income_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If wages_income_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Wages Exist"
		wages_income_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If va_income_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "VA Income Exist"
		va_income_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If self_emp_income_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Self Emp Exists"
		self_emp_income_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If no_income_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "No Income"
		no_income_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If epd_on_case_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "EPD on Case"
		epd_on_case_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If year_month_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Year Month"
		year_month_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If eval_year_month_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Evaluation Year Month"
		eval_year_month_col = col_to_use
		col_to_use = col_to_use + 1
	End If
	If approval_year_month_checkbox = checked Then
    	ObjExcel.Cells(1, col_to_use).Value = "Approval Year Month"
		approval_year_month_col = col_to_use
		col_to_use = col_to_use + 1
	End If

	For col_to_autofit = 1 to col_to_use
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next

	excel_row = excel_row_to_start * 1
	Do
		MAXIS_case_number = trim(ObjExcel.Cells(excel_row, case_number_column).Value)
		MAXIS_case_number = right("00000000"&MAXIS_case_number, 8)


		'This is opening the Ex Parte Case List data table so we can loop through it.
		objLIST = "SELECT * FROM [ES].[ES_ExParte_CaseList] WHERE CaseNumber = '" & MAXIS_case_number & "'"		'we only need to look at the cases for the specific review month

		Set objConnect = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
		Set objTheRecord = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objConnect.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objTheRecord.Open objLIST, objConnect

		Do While NOT objTheRecord.Eof 					'Loop through each item on the CASE LIST Table

			SQL_worker_id 				= objTheRecord("WorkerID")
			SQL_HCElig 					= objTheRecord("HCEligReviewDate")
			SQL_Select_ExParte 			= objTheRecord("SelectExParte")

			SQL_prep_complete 			= objTheRecord("PREP_Complete")
			SQL_phase1_complete 		= objTheRecord("Phase1Complete")
			SQL_phase1_hsr 				= objTheRecord("Phase1HSR")
			SQL_ex_parte_after_phase1 	= objTheRecord("ExParteAfterPhase1")
			SQL_phase1_cancel_reason 	= objTheRecord("Phase1ExParteCancelReason")
			SQL_phase2_complete 		= objTheRecord("Phase2Complete")
			SQL_phase2_hsr 				= objTheRecord("Phase2HSR")
			SQL_ex_parte_after_phase2 	= objTheRecord("ExParteAfterPhase2")
			SQL_phase2_cancel_reason 	= objTheRecord("Phase2ExParteCancelReason")
			SQL_all_HC_is_ABD 			= objTheRecord("AllHCisABD")
			SQL_ssa_income_exists 		= objTheRecord("SSAIncomExist")
			SQL_wages_exist 			= objTheRecord("WagesExist")
			SQL_va_inc_exists 			= objTheRecord("VAIncomeExist")
			SQL_self_emp_exists 		= objTheRecord("SelfEmpExists")
			SQL_no_income 				= objTheRecord("NoIncome")
			SQL_EPD_on_case 			= objTheRecord("EPDonCase")
			SQL_year_month 				= objTheRecord("YearMonth")
			SQL_eval_year_month 		= objTheRecord("EvaluationYearMonth")
			SQL_app_year_month 			= objTheRecord("ApprovalYearMonth")
			objTheRecord.MoveNext			'now we go to the next case

			If worker_id_checkbox = checked Then ObjExcel.Cells(excel_row, worker_id_col).Value = SQL_worker_id
			If hc_er_date_checkbox = checked Then
				SQL_HCElig = DateAdd("d", 0, SQL_HCElig)
				ObjExcel.Cells(excel_row, hc_er_date_col).Value = SQL_HCElig
			End If
			If select_ex_parte_checkbox = checked Then ObjExcel.Cells(excel_row, select_ex_parte_col).Value = SQL_Select_ExParte


			If prep_complete_checkbox = checked Then ObjExcel.Cells(excel_row, prep_complete_col).Value = SQL_prep_complete
			If phase_1_complete_checkbox = checked Then ObjExcel.Cells(excel_row, phase_1_complete_col).Value = SQL_phase1_complete
			If phase_1_HSR_checkbox = checked Then
				ObjExcel.Cells(excel_row, phase_1_HSR_col).Value = SQL_phase1_hsr
				windows_user_ID = ucase(trim(SQL_phase1_hsr))
				Call find_user_name(phase_1_worker)
				ObjExcel.Cells(excel_row, phase_1_hsr_name_col).Value = phase_1_worker
			End If
			If ex_parte_after_phase_1_checkbox = checked Then ObjExcel.Cells(excel_row, ex_parte_after_phase_1_col).Value = SQL_ex_parte_after_phase1
			If phase_1_ex_parte_cancel_checkbox = checked Then ObjExcel.Cells(excel_row, phase_1_ex_parte_cancel_col).Value = SQL_phase1_cancel_reason
			If phase_2_complete_checkbox = checked Then ObjExcel.Cells(excel_row, phase_2_complete_col).Value = SQL_phase2_complete
			If phase_2_hsr_checkbox = checked Then
				ObjExcel.Cells(excel_row, phase_2_hsr_col).Value = SQL_phase2_hsr
				windows_user_ID = ucase(trim(SQL_phase2_hsr))
				Call find_user_name(phase_2_worker)
				ObjExcel.Cells(excel_row, phase_2_hsr_name_col).Value = phase_2_worker
			End If
			If ex_parte_after_phase_2_checkbox = checked Then ObjExcel.Cells(excel_row, ex_parte_after_phase_2_col).Value = SQL_ex_parte_after_phase2
			If phase_2_ex_parte_cancel_checkbox = checked Then ObjExcel.Cells(excel_row, phase_2_ex_parte_cancel_col).Value = SQL_phase2_cancel_reason
			If all_hc_is_abd_checkbox = checked Then ObjExcel.Cells(excel_row, all_hc_is_abd_col).Value = SQL_all_HC_is_ABD
			If ssa_income_checkbox = checked Then ObjExcel.Cells(excel_row, ssa_income_col).Value = SQL_ssa_income_exists
			If wages_income_checkbox = checked Then ObjExcel.Cells(excel_row, wages_income_col).Value = SQL_wages_exist
			If va_income_checkbox = checked Then ObjExcel.Cells(excel_row, va_income_col).Value = SQL_va_inc_exists
			If self_emp_income_checkbox = checked Then ObjExcel.Cells(excel_row, self_emp_income_col).Value = SQL_self_emp_exists
			If no_income_checkbox = checked Then ObjExcel.Cells(excel_row, no_income_col).Value = SQL_no_income
			If epd_on_case_checkbox = checked Then ObjExcel.Cells(excel_row, epd_on_case_col).Value = SQL_EPD_on_case
			If year_month_checkbox = checked Then ObjExcel.Cells(excel_row, year_month_col).Value = SQL_year_month
			If eval_year_month_checkbox = checked Then ObjExcel.Cells(excel_row, eval_year_month_col).Value = SQL_eval_year_month
			If approval_year_month_checkbox = checked Then ObjExcel.Cells(excel_row, approval_year_month_col).Value = SQL_app_year_month

			windows_user_ID = original_user
			user_ID_for_validation = ucase(windows_user_ID)

		Loop
		objTheRecord.Close			'Closing all the data connections
		objConnect.Close
		Set objTheRecord=nothing
		Set objConnect=nothing

		excel_row = excel_row + 1


		next_case_numb = trim(ObjExcel.Cells(excel_row, case_number_column).Value)

	Loop until next_case_numb = ""




	'find the last filled column, add columns to fill excel information

	'loop through each row on the Excel

	'create a SQL version of the case number
	'select the case from SQL and enter the new information into the spreadsheet
	end_msg = "Info done"
End If

'Loop through all the SQL Items and look for the right revew month and year and phase to determine if it's done.

Call script_end_procedure(end_msg)
