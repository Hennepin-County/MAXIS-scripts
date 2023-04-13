'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - TRAINING CASE CREATOR"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                              'manual run time in seconds  this run time only includes appl'ing the case. it gets time added it to as panels are added and approvals are made.
STATS_denomination = "C"       'I is for each case
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
		FuncLib_URL = "C:\MAXIS-Scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
Call changelog_update("01/03/2020", "Added custom functions to the script file itself.", "Ilse Ferris, Hennepin County")
call changelog_update("08/07/2019", "Updated coding to read Marital status, spouse reference number, last grade completed citizenship, citizenship verif, in MN more than 12 months and residence verif codes at new location due to MEMI panel changes associated with New Spouse Income Policy.", "Ilse Ferris, Hennepin County")
call changelog_update("03/22/2019", "Added handling for FSET sanction reasons, updated input for sanction dates and updated dialog with current contact info.", "Ilse Ferris, Hennepin County")
call changelog_update("10/2/2018", "Fixed bug with creating MEDI panel. Added functionality to add waiver or 1619 status to DISA.", "Casey Love, Hennepin County")
call changelog_update("03/28/2018", "Added handling to send the HRF, and updated REI handling for MFIP cases.", "Ilse Ferris, Hennepin County")
call changelog_update("03/06/2018", "Updated WF1M handling for MFIP cases that require a referral.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'========================================================================TRANSFER CASES========================================================================
'Functions for PROJECT KRABAPPEL (UTILITIES - TRAINING CASE CREATOR)====================================================================================================================================================
'writing in placeholder information for now re: the functions and parameters to be in line with the new documentation process.

function write_panel_to_MAXIS_ABPS(abps_supp_coop,abps_gc_status)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ abps_supp_coop,abps_gc_status: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","PARE")							'Starts by creating an array of all the kids on PARE
	EMReadScreen pare_check, 4, 2, 43
	If pare_check <> "PARE" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "PARE")
	End If
	EMReadScreen abps_pare_check, 1, 2, 78
	If abps_pare_check = "0" then
		MsgBox "No PARE exists. Exiting Creating ABPS."
	ElseIf abps_pare_check <> "0" then
		child_list = ""
		row = 8
		Do
			EMReadScreen child_check, 2, row, 24
			If child_check <> "__" then
				If child_list = "" then
					child_list = child_check
				ElseIf child_list <> "" then
					child_list = child_list & "," & child_check
				End If
			End If
			row = row + 1
			If row = 18 then
				PF8
				row = 8
			End If
		Loop until child_check = "__"
		call navigate_to_MAXIS_screen("STAT","ABPS")						'Navigates to ABPS to enter kids in
		EMReadScreen abps_check, 4, 2, 50
		If abps_check <> "ABPS" Then
			MAXIS_background_check
			call navigate_to_MAXIS_screen("STAT", "ABPS")
		End If
		call create_panel_if_nonexistent
		abps_child_list = split(child_list, ",")
		row = 15
		for each abps_child in abps_child_list
			EMWriteScreen abps_child, row, 35
			EMWriteScreen "2", row, 53
			EMWriteScreen "1", row, 67
			row = row + 1
			If row = 18 then
				PF8
				row = 15
			End If
		next
		IF abps_act_date <> "" THEN call create_MAXIS_friendly_date_with_YYYY(date, 0, 18, 38)
		EMWriteScreen reference_number, 4, 47		'Enters the reference_number
		If abps_supp_coop <> "" then
			abps_supp_coop = ucase(abps_supp_coop)
			abps_supp_coop = left(abps_supp_coop,1)
			EMWriteScreen abps_supp_coop, 4, 73
		End If
		If abps_gc_status <> "" then
			EMWriteScreen abps_gc_status, 5, 47
		End If
		transmit
	End If
end function

function write_panel_to_MAXIS_ACCT(acct_type, acct_numb, acct_location, acct_balance, acct_bal_ver, acct_date, acct_withdraw, acct_cash_count, acct_snap_count, acct_HC_count, acct_GRH_count, acct_IV_count, acct_joint_owner, acct_share_ratio, acct_interest_date_mo, acct_interest_date_yr)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ acct_type, acct_numb, acct_location, acct_balance, acct_bal_ver, acct_date, acct_withdraw, acct_cash_count, acct_snap_count, acct_HC_count, acct_GRH_count, acct_IV_count, acct_joint_owner, acct_share_ratio, acct_interest_date_mo, acct_interest_date_yr: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "ACCT")  'navigates to the stat panel
	EMReadScreen acct_check, 4, 2, 44
	If acct_check <> "ACCT" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "ACCT")
	End If
	call create_panel_if_nonexistent
	Emwritescreen acct_type, 6, 44  'enters the account type code
	Emwritescreen acct_numb, 7, 44  'enters the account number
	Emwritescreen acct_location, 8, 44  'enters the account location

	' >>>>> Comment: Updated 06/22/2016 <<<<<
	' >>>>> Looking for the acct_bal_ver location. This changed with asset unification...
	' >>>>> ... but the location is not the same across all months. It needs to be variable...
	' >>>>> ... so Krabappel knows where to write stuff and junk or whatever ...
	' >>>>> This has been tested on training case 226398 for the benefit months 05/16 and 06/16...
	' >>>>> ... in 05/16 the acct_bal_ver coordinates are 10, 63 and in 06/16, they are 10, 64...
	' >>>>> ... and the code is working in both months.
	' >>> Looking for the balance field and then we will write the verif code on the same line...
	acct_row = 1
	acct_col = 1
	EMSearch "Balance: ", acct_row, acct_col
	EMWriteScreen acct_balance, acct_row, acct_col + 11  'enters the balance
	acct_col = 1
	EMSearch "Ver: ", acct_row, acct_col
	EMWriteScreen acct_bal_ver, acct_row, acct_col + 5  'enters the balance verification

	IF acct_date <> "" THEN call create_MAXIS_friendly_date(acct_date, 0, 11, 44)  'enters the account balance date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen acct_withdraw, 12, 46  'enters the withdrawl penalty
	Emwritescreen acct_cash_count, 14, 50  'enters y/n if counted for cash
	Emwritescreen acct_snap_count, 14, 57  'enters y/n if counted for snap
	Emwritescreen acct_HC_count, 14, 64  'enters y/n if counted for HC
	Emwritescreen acct_GRH_count, 14, 72  'enters y/n if counted for grh
	Emwritescreen acct_IV_count, 14, 80  'enters y/n if counted for IV
	Emwritescreen acct_joint_owner, 15, 44  'enters if it is a jointly owned acct
	Emwritescreen left(acct_share_ratio, 1), 15, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(acct_share_ratio, 1), 15, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
	Emwritescreen acct_interest_date_mo, 17, 57  'enters the next interest date MM format
	Emwritescreen acct_interest_date_yr, 17, 60  'enters the next interest date YY format
	transmit
	transmit
end function

function write_panel_to_MAXIS_ACUT(ACUT_shared, ACUT_heat, ACUT_air, ACUT_electric, ACUT_fuel, ACUT_garbage, ACUT_water, ACUT_sewer, ACUT_other, ACUT_phone, ACUT_heat_verif, ACUT_air_verif, ACUT_electric_verif, ACUT_fuel_verif, ACUT_garbage_verif, ACUT_water_verif, ACUT_sewer_verif, ACUT_other_verif)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ ACUT_shared, ACUT_heat, ACUT_air, ACUT_electric, ACUT_fuel, ACUT_garbage, ACUT_water, ACUT_sewer, ACUT_other, ACUT_phone, ACUT_heat_verif, ACUT_air_verif, ACUT_electric_verif, ACUT_fuel_verif, ACUT_garbage_verif, ACUT_water_verif, ACUT_sewer_verif, ACUT_other_verif: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "ACUT")
	EMReadScreen acut_check, 4, 2, 52
	If acut_check <> "ACUT" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "ACUT")
	End If
	call create_panel_if_nonexistent
		EMWritescreen ACUT_shared, 6, 42
		EMWritescreen ACUT_heat, 10, 61
		EMWritescreen ACUT_air, 11, 61
		EMWritescreen ACUT_electric, 12, 61
		EMWritescreen ACUT_fuel, 13, 61
		EMWritescreen ACUT_garbage, 14, 61
		EMWritescreen ACUT_water, 15, 61
		EMWritescreen ACUT_sewer, 16, 61
		EMWritescreen ACUT_other, 17, 61
		EMWritescreen ACUT_heat_verif, 10, 55
		EMWritescreen ACUT_air_verif, 11, 55
		EMWritescreen ACUT_electric_verif, 12, 55
		EMWritescreen ACUT_fuel_verif, 13, 55
		EMWritescreen ACUT_garbage_verif, 14, 55
		EMWritescreen ACUT_water_verif, 15, 55
		EMWritescreen ACUT_sewer_verif, 16, 55
		EMWritescreen ACUT_other_verif, 17, 55
		EMWritescreen Left(ACUT_phone, 1), 18, 55
	transmit
end function

'---This function writes the information for BILS.
function write_panel_to_MAXIS_BILS(bils_1_ref_num, bils_1_serv_date, bils_1_serv_type, bils_1_gross_amt, bils_1_third_party, bils_1_verif, bils_1_bils_type, bils_2_ref_num, bils_2_serv_date, bils_2_serv_type, bils_2_gross_amt, bils_2_third_party, bils_2_verif, bils_2_bils_type, bils_3_ref_num, bils_3_serv_date, bils_3_serv_type, bils_3_gross_amt, bils_3_third_party, bils_3_verif, bils_3_bils_type, bils_4_ref_num, bils_4_serv_date, bils_4_serv_type, bils_4_gross_amt, bils_4_third_party, bils_4_verif, bils_4_bils_type, bils_5_ref_num, bils_5_serv_date, bils_5_serv_type, bils_5_gross_amt, bils_5_third_party, bils_5_verif, bils_5_bils_type, bils_6_ref_num, bils_6_serv_date, bils_6_serv_type, bils_6_gross_amt, bils_6_third_party, bils_6_verif, bils_6_bils_type, bils_7_ref_num, bils_7_serv_date, bils_7_serv_type, bils_7_gross_amt, bils_7_third_party, bils_7_verif, bils_7_bils_type, bils_8_ref_num, bils_8_serv_date, bils_8_serv_type, bils_8_gross_amt, bils_8_third_party, bils_8_verif, bils_8_bils_type, bils_9_ref_num, bils_9_serv_date, bils_9_serv_type, bils_9_gross_amt, bils_9_third_party, bils_9_verif, bils_9_bils_type)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ bils_1_ref_num, bils_1_serv_date, bils_1_serv_type, bils_1_gross_amt, bils_1_third_party, bils_1_verif, bils_1_bils_type, bils_2_ref_num, bils_2_serv_date, bils_2_serv_type, bils_2_gross_amt, bils_2_third_party, bils_2_verif, bils_2_bils_type, bils_3_ref_num, bils_3_serv_date, bils_3_serv_type, bils_3_gross_amt, bils_3_third_party, bils_3_verif, bils_3_bils_type, bils_4_ref_num, bils_4_serv_date, bils_4_serv_type, bils_4_gross_amt, bils_4_third_party, bils_4_verif, bils_4_bils_type, bils_5_ref_num, bils_5_serv_date, bils_5_serv_type, bils_5_gross_amt, bils_5_third_party, bils_5_verif, bils_5_bils_type, bils_6_ref_num, bils_6_serv_date, bils_6_serv_type, bils_6_gross_amt, bils_6_third_party, bils_6_verif, bils_6_bils_type, bils_7_ref_num, bils_7_serv_date, bils_7_serv_type, bils_7_gross_amt, bils_7_third_party, bils_7_verif, bils_7_bils_type, bils_8_ref_num, bils_8_serv_date, bils_8_serv_type, bils_8_gross_amt, bils_8_third_party, bils_8_verif, bils_8_bils_type, bils_9_ref_num, bils_9_serv_date, bils_9_serv_type, bils_9_gross_amt, bils_9_third_party, bils_9_verif, bils_9_bils_type: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "BILS")
	EMReadScreen bils_check, 4, 2, 54
	If bils_check <> "BILS" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "BILS")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	EMReadScreen num_of_BILS, 1, 2, 78
	IF num_of_BILS = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF

	'---MAXIS will not allow BILS to be updated if HC is inactive. Exiting the function if HC is inactive.
	EMReadScreen hc_inactive, 21, 24, 2
	IF hc_inactive = "HC STATUS IS INACTIVE" THEN Exit function

	BILS_row = 6
	DO
		EMReadScreen available_row, 2, BILS_row, 26
		IF available_row <> "__" THEN BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	LOOP UNTIL available_row = "__"

	IF bils_1_ref_num <> "" THEN
		IF len(bils_1_ref_num) = 1 THEN bils_1_ref_num = "0" & bils_1_ref_num
		EMWriteScreen bils_1_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_1_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_1_serv_type, BILS_row, 40
		EMWriteScreen bils_1_gross_amt, BILS_row, 45
		EMWriteScreen bils_1_third_party, BILS_row, 57
		IF bils_1_verif = "03" AND bils_1_serv_type <> "22" THEN bils_1_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_1_verif, BILS_row, 67
		EMWriteScreen bils_1_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_2_ref_num <> "" THEN
		IF len(bils_2_ref_num) = 1 THEN bils_2_ref_num = "0" & bils_2_ref_num
		EMWriteScreen bils_2_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_2_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_2_serv_type, BILS_row, 40
		EMWriteScreen bils_2_gross_amt, BILS_row, 45
		EMWriteScreen bils_2_third_party, BILS_row, 57
		IF bils_2_verif = "03" AND bils_2_serv_type <> "22" THEN bils_2_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_2_verif, BILS_row, 67
		EMWriteScreen bils_2_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_3_ref_num <> "" THEN
		IF len(bils_3_ref_num) = 1 THEN bils_3_ref_num = "0" & bils_3_ref_num
		EMWriteScreen bils_3_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_3_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_3_serv_type, BILS_row, 40
		EMWriteScreen bils_3_gross_amt, BILS_row, 45
		EMWriteScreen bils_3_third_party, BILS_row, 57
		IF bils_3_verif = "03" AND bils_3_serv_type <> "22" THEN bils_3_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_3_verif, BILS_row, 67
		EMWriteScreen bils_3_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_4_ref_num <> "" THEN
		IF len(bils_4_ref_num) = 1 THEN bils_4_ref_num = "0" & bils_4_ref_num
		EMWriteScreen bils_4_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_4_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_4_serv_type, BILS_row, 40
		EMWriteScreen bils_4_gross_amt, BILS_row, 45
		EMWriteScreen bils_4_third_party, BILS_row, 57
		IF bils_4_verif = "03" AND bils_4_serv_type <> "22" THEN bils_4_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_4_verif, BILS_row, 67
		EMWriteScreen bils_4_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_5_ref_num <> "" THEN
		IF len(bils_5_ref_num) = 1 THEN bils_5_ref_num = "0" & bils_5_ref_num
		EMWriteScreen bils_5_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_5_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_5_serv_type, BILS_row, 40
		EMWriteScreen bils_5_gross_amt, BILS_row, 45
		EMWriteScreen bils_5_third_party, BILS_row, 57
		IF bils_5_verif = "03" AND bils_5_serv_type <> "22" THEN bils_5_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_5_verif, BILS_row, 67
		EMWriteScreen bils_5_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_6_ref_num <> "" THEN
		IF len(bils_6_ref_num) = 1 THEN bils_6_ref_num = "0" & bils_6_ref_num
		EMWriteScreen bils_6_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_6_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_6_serv_type, BILS_row, 40
		EMWriteScreen bils_6_gross_amt, BILS_row, 45
		EMWriteScreen bils_6_third_party, BILS_row, 57
		IF bils_6_verif = "03" AND bils_6_serv_type <> "22" THEN bils_6_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_6_verif, BILS_row, 67
		EMWriteScreen bils_6_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_7_ref_num <> "" THEN
		IF len(bils_7_ref_num) = 1 THEN bils_7_ref_num = "0" & bils_7_ref_num
		EMWriteScreen bils_7_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_7_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_7_serv_type, BILS_row, 40
		EMWriteScreen bils_7_gross_amt, BILS_row, 45
		EMWriteScreen bils_7_third_party, BILS_row, 57
		IF bils_7_verif = "03" AND bils_7_serv_type <> "22" THEN bils_7_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_7_verif, BILS_row, 67
		EMWriteScreen bils_7_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_8_ref_num <> "" THEN
		IF len(bils_8_ref_num) = 1 THEN bils_8_ref_num = "0" & bils_8_ref_num
		EMWriteScreen bils_8_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_8_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_8_serv_type, BILS_row, 40
		EMWriteScreen bils_8_gross_amt, BILS_row, 45
		EMWriteScreen bils_8_third_party, BILS_row, 57
		IF bils_8_verif = "03" AND bils_8_serv_type <> "22" THEN bils_8_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_8_verif, BILS_row, 67
		EMWriteScreen bils_8_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_9_ref_num <> "" THEN
		IF len(bils_9_ref_num) = 1 THEN bils_9_ref_num = "0" & bils_9_ref_num
		EMWriteScreen bils_9_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_9_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_9_serv_type, BILS_row, 40
		EMWriteScreen bils_9_gross_amt, BILS_row, 45
		EMWriteScreen bils_9_third_party, BILS_row, 57
		IF bils_9_verif = "03" AND bils_9_serv_type <> "22" THEN bils_9_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_9_verif, BILS_row, 67
		EMWriteScreen bils_9_bils_type, BILS_row, 71
	END IF
end function

function write_panel_to_MAXIS_BUSI(busi_type, busi_start_date, busi_end_date, busi_cash_total_retro, busi_cash_total_prosp, busi_cash_total_ver, busi_IV_total_prosp, busi_IV_total_ver, busi_snap_total_retro, busi_snap_total_prosp, busi_snap_total_ver, busi_hc_total_prosp_a, busi_hc_total_ver_a, busi_hc_total_prosp_b, busi_hc_total_ver_b, busi_cash_exp_retro, busi_cash_exp_prosp, busi_cash_exp_ver, busi_IV_exp_prosp, busi_IV_exp_ver, busi_snap_exp_retro, busi_snap_exp_prosp, busi_snap_exp_ver, busi_hc_exp_prosp_a, busi_hc_exp_ver_a, busi_hc_exp_prosp_b, busi_hc_exp_ver_b, busi_retro_hours, busi_prosp_hours, busi_hc_total_est_a, busi_hc_total_est_b, busi_hc_exp_est_a, busi_hc_exp_est_b, busi_hc_hours_est)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the busi panel in MAXIS)
'~~~~~ busi_type, busi_start_date, busi_end_date, busi_cash_total_retro, busi_cash_total_prosp, busi_cash_total_ver, busi_IV_total_prosp, busi_IV_total_ver, busi_snap_total_retro, busi_snap_total_prosp, busi_snap_total_ver, busi_hc_total_prosp_a, busi_hc_total_ver_a, busi_hc_total_prosp_b, busi_hc_total_ver_b, busi_cash_exp_retro, busi_cash_exp_prosp, busi_cash_exp_ver, busi_IV_exp_prosp, busi_IV_exp_ver, busi_snap_exp_retro, busi_snap_exp_prosp, busi_snap_exp_ver, busi_hc_exp_prosp_a, busi_hc_exp_ver_a, busi_hc_exp_prosp_b, busi_hc_exp_ver_b, busi_retro_hours, busi_prosp_hours, busi_hc_total_est_a, busi_hc_total_est_b, busi_hc_exp_est_a, busi_hc_exp_est_b, busi_hc_hours_est: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "BUSI")  'navigates to the stat panel
	EMReadScreen busi_check, 4, 2, 51
	If busi_check <> "BUSI" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "BUSI")
	End If
	Emwritescreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_BUSI, 1, 2, 78
	IF num_of_BUSI = "0" THEN
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit

		'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
		If datediff("d", "02/01/2015", BUSI_footer_month) < 0 then
			Emwritescreen busi_type, 5, 37  'enters self employment type
			call create_MAXIS_friendly_date(busi_start_date, 0, 5, 54)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			IF busi_end_date <> "" THEN call create_MAXIS_friendly_date(busi_end_date, 0, 5, 71)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			Emwritescreen "x", 7, 26  'this enters into the gross income calculator
			Transmit
			Do
				Emreadscreen busi_gross_income_check, 12, 06, 35  'This checks to see if the gross income calculator has actually opened.
			LOOP UNTIL busi_gross_income_check = "Gross Income"
			Emwritescreen busi_cash_total_retro, 9, 43  'enters the cash total income retrospective number
			Emwritescreen busi_cash_total_prosp, 9, 59  'enters the cash total income prospective number
			Emwritescreen busi_cash_total_ver, 9, 73    'enters the cash total income verification code
			Emwritescreen busi_IV_total_prosp, 10, 59   'enters the IV total income prospective number
			Emwritescreen busi_IV_total_ver, 10, 73     'enters the IV total income verification code
			Emwritescreen busi_snap_total_retro, 11, 43 'enters the snap total income retro number
			Emwritescreen busi_snap_total_prosp, 11, 59 'enters the snap total income prosp number
			Emwritescreen busi_snap_total_ver, 11, 73   'enters the snap total verification code
			Emwritescreen busi_hc_total_prosp_a, 12, 59 'enters the HC total income prospective number for method a
			Emwritescreen busi_hc_total_ver_a, 12, 73   'enters the HC total income verification code for method a
			Emwritescreen busi_hc_total_prosp_b, 13, 59 'enters the HC total income prospective number for method b
			Emwritescreen busi_hc_total_ver_b, 13, 73   'enters the HC total income verification code for method b
			Emwritescreen busi_cash_exp_retro, 15, 43   'enters the cash expenses retrospective number
			Emwritescreen busi_cash_exp_prosp, 15, 59   'enters the cash expenses prospective number
			Emwritescreen busi_cash_exp_ver, 15, 73     'enters the cash expenses verification code
			Emwritescreen busi_IV_exp_prosp, 16, 59     'enters the IV expenses retro number
			Emwritescreen busi_IV_exp_ver, 9, 73        'enters the IV expenses verification code
			Emwritescreen busi_snap_exp_retro, 17, 43   'enters the snap expenses retro number
			Emwritescreen busi_snap_exp_prosp, 17, 59   'enters the snap expenses prospective number
			Emwritescreen busi_snap_exp_ver, 17, 73     'enters the snap expenses verif code
			Emwritescreen busi_hc_exp_prosp_a, 18, 59   'enters the hc expenses prospective number for method a
			Emwritescreen busi_hc_exp_ver_a, 18, 73    'enters the hc expenses verification code for method a
			Emwritescreen busi_hc_exp_prosp_b, 19, 59   'enters the hc expenses prospective number for method b
			Emwritescreen busi_hc_exp_ver_b, 19, 73	  'enters the hc expenses verification code for method b
			transmit
			PF3
			Emwritescreen busi_retro_hours, 14, 59  'enters the retrospective hours
			Emwritescreen busi_prosp_hours, 14, 73  'enters the prospective hours

		ELSE				'This is the NEW logic for all months after 02/2015
			Emwritescreen busi_type, 5, 37  'enters self employment type
			call create_MAXIS_friendly_date(busi_start_date, 0, 5, 55)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			IF busi_end_date <> "" THEN call create_MAXIS_friendly_date(busi_end_date, 0, 5, 72)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			Emwritescreen "x", 6, 26  'this enters into the gross income calculator
			Transmit
			Do
				Emreadscreen busi_gross_income_check, 12, 06, 35  'This checks to see if the gross income calculator has actually opened.
			LOOP UNTIL busi_gross_income_check = "Gross Income"
			Emwritescreen busi_cash_total_retro, 9, 43  'enters the cash total income retrospective number
			Emwritescreen busi_cash_total_prosp, 9, 59  'enters the cash total income prospective number
			Emwritescreen busi_cash_total_ver, 9, 73    'enters the cash total income verification code
			Emwritescreen busi_IV_total_prosp, 10, 59   'enters the IV total income prospective number
			Emwritescreen busi_IV_total_ver, 10, 73     'enters the IV total income verification code
			Emwritescreen busi_snap_total_retro, 11, 43 'enters the snap total income retro number
			Emwritescreen busi_snap_total_prosp, 11, 59 'enters the snap total income prosp number
			Emwritescreen busi_snap_total_ver, 11, 73   'enters the snap total verification code
			Emwritescreen busi_hc_total_prosp_a, 12, 59 'enters the HC total income prospective number for method a
			Emwritescreen busi_hc_total_ver_a, 12, 73   'enters the HC total income verification code for method a
			Emwritescreen busi_hc_total_prosp_b, 13, 59 'enters the HC total income prospective number for method b
			Emwritescreen busi_hc_total_ver_b, 13, 73   'enters the HC total income verification code for method b
			Emwritescreen busi_cash_exp_retro, 15, 43   'enters the cash expenses retrospective number
			Emwritescreen busi_cash_exp_prosp, 15, 59   'enters the cash expenses prospective number
			Emwritescreen busi_cash_exp_ver, 15, 73     'enters the cash expenses verification code
			Emwritescreen busi_IV_exp_prosp, 16, 59     'enters the IV expenses retro number
			Emwritescreen busi_IV_exp_ver, 9, 73        'enters the IV expenses verification code
			Emwritescreen busi_snap_exp_retro, 17, 43   'enters the snap expenses retro number
			Emwritescreen busi_snap_exp_prosp, 17, 59   'enters the snap expenses prospective number
			Emwritescreen busi_snap_exp_ver, 17, 73     'enters the snap expenses verif code
			Emwritescreen busi_hc_exp_prosp_a, 18, 59   'enters the hc expenses prospective number for method a
			Emwritescreen busi_hc_exp_ver_a, 18, 73    'enters the hc expenses verification code for method a
			Emwritescreen busi_hc_exp_prosp_b, 19, 59   'enters the hc expenses prospective number for method b
			Emwritescreen busi_hc_exp_ver_b, 19, 73	  'enters the hc expenses verification code for method b
			transmit
			PF3
			Emwritescreen busi_retro_hours, 13, 60  'enters the retrospective hours
			Emwritescreen busi_prosp_hours, 13, 74  'enters the prospective hours
			'---Adding Self-Employment Method -- Hard-Coded for now.
			EMWriteScreen "01", 16, 53
			CALL create_MAXIS_friendly_date(#02/01/2015#, 0, 16, 63)
		END IF
	ELSEIF num_of_BUSI <> "0" THEN
		PF9
		'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
		If datediff("d", "02/01/2015", BUSI_footer_month) >= 0 then
			'---Adding Self-Employment Method -- Hard-Coded for now.
			EMWriteScreen "01", 16, 53
			CALL create_MAXIS_friendly_date(#02/01/2015#, 0, 16, 63)
			'---Going into the HC Income Estimate
			EMWriteScreen "X", 17, 27
			transmit
			DO
				EMReadScreen hc_income, 9, 4, 42
			LOOP UNTIL hc_income = "HC Income"
			EMReadScreen current_month_plus_one, 17, 21, 59
			IF current_month_plus_one = "CURRENT MONTH + 1" THEN
				PF3
			ELSE
				Emwritescreen busi_hc_total_est_a, 7, 54                'enters hc total income estimation for method A
				Emwritescreen busi_hc_total_est_b, 8, 54                'enters hc total income estimation for method B
				Emwritescreen busi_hc_exp_est_a, 11, 54                 'enters hc expense estimation for method A
				Emwritescreen busi_hc_exp_est_b, 12, 54                 'enters hc expense estimation for method B
				Emwritescreen busi_hc_hours_est, 18, 58                 'enters hc hours estimation
				transmit
				PF3
			END IF
		END IF
	END IF
end function

function write_panel_to_MAXIS_CARS(cars_type, cars_year, cars_make, cars_model, cars_trade_in, cars_loan, cars_value_source, cars_ownership_ver, cars_amount_owed, cars_amount_owed_ver, cars_date, cars_use, cars_HC_benefit, cars_joint_owner, cars_share_ratio)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ cars_type, cars_year, cars_make, cars_model, cars_trade_in, cars_loan, cars_value_source, cars_ownership_ver, cars_amount_owed, cars_amount_owed_ver, cars_date, cars_use, cars_HC_benefit, cars_joint_owner, cars_share_ratio: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "CARS")  'navigates to the stat screen
	EMReadScreen panel_check, 4, 2, 44
	If panel_check <> "CARS" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "CARS")
	End If
	call create_panel_if_nonexistent
	Emwritescreen cars_type, 6, 43  'enters the vehicle type
	Emwritescreen cars_year, 8, 31  'enters the vehicle year
	Emwritescreen cars_make, 8, 43  'enters the vehicle make
	Emwritescreen cars_model, 8, 66  'enters the vehicle model
	Emwritescreen cars_trade_in, 9, 45  'enters the trade in value
	Emwritescreen cars_loan, 9, 62  'enters the loan value
	Emwritescreen cars_value_source, 9, 80  'enters the source of value information
	Emwritescreen cars_ownership_ver, 10, 60  'enters the ownership verification code
	Emwritescreen cars_amount_owed, 12, 45  'enters the amount owed on vehicle
	Emwritescreen cars_amount_owed_ver, 12, 60  'enters the amount owed verification code
	IF cars_date <> "" THEN call create_MAXIS_friendly_date(cars_date, 0, 13, 43)  'enters the amounted owed as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen cars_use, 15, 43  'enters the use code for the vehicle
	Emwritescreen cars_HC_benefit, 15, 76  'enters if the vehicle is for client benefit
	Emwritescreen cars_joint_owner, 16, 43  'enters if it is a jointly owned car
	Emwritescreen left(cars_share_ratio, 1), 16, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(cars_share_ratio, 1), 16, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
end function

function write_panel_to_MAXIS_CASH(cash_amount)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the cash panel in MAXIS)
'~~~~~ cash_amount: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "CASH")  'navigates to the stat panel
	EMReadScreen panel_check, 4, 2, 42
	If panel_check <> "CASH" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "CASH")
	End If
	call create_panel_if_nonexistent
	Emwritescreen cash_amount, 8, 39
end function

function write_panel_to_MAXIS_COEX(retro_support, prosp_support, support_verif, retro_alimony, prosp_alimony, alimony_verif, retro_tax_dep, prosp_tax_dep, tax_dep_verif, retro_other, prosp_other, other_verif, change_in_circum, hc_exp_support, hc_exp_alimony, hc_exp_tax_dep, hc_exp_other)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the COEX panel in MAXIS.)
'~~~~~ retro_support, prosp_support, support_verif, retro_alimony, prosp_alimony, alimony_verif, retro_tax_dep, prosp_tax_dep, tax_dep_verif, retro_other, prosp_other, other_verif, change_in_circum, hc_exp_support, hc_exp_alimony, hc_exp_tax_dep, hc_exp_other: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "COEX")
	EMReadScreen panel_check, 4, 2, 51
	If panel_check <> "COEX" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "COEX")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	EMWriteScreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_COEX, 1, 2, 78
	IF num_of_COEX = "0" THEN
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit
		'---If the script is creating a new COEX panel, it will enter this information...
		EMWriteScreen support_verif, 10, 36
		EMWriteScreen retro_support, 10, 45
		EMWriteScreen prosp_support, 10, 63
		EMWriteScreen alimony_verif, 11, 36
		EMWriteScreen retro_alimony, 11, 45
		EMWriteScreen prosp_alimony, 11, 63
		EMWriteScreen tax_dep_verif, 12, 36
		EMWriteScreen retro_tax_dep, 12, 45
		EMWriteScreen prosp_tax_dep, 12, 63
		EMWriteScreen other_verif, 13, 36
		EMWriteScreen retro_other, 13, 45
		EMWriteScreen prosp_other, 13, 63
		EMWriteScreen change_in_circum, 17, 61
	ELSEIF num_of_COEX <> "0" THEN
		PF9
		'---...if the script is PF9'ing, it is doing so to enter information into the HC Expense sub-menu
		'Opening the HC Expenses Sub-menu
		EMWriteScreen "X", 18, 44
		transmit

		DO
			EMReadScreen hc_expense_est, 14, 4, 30
		LOOP UNTIL hc_expense_est = "HC Expense Est"

		EMReadScreen current_month_plus_one, 17, 13, 51
		IF current_month_plus_one <> "CURRENT MONTH + 1" THEN
			EMWriteScreen hc_exp_support, 6, 38
			EMWriteScreen hc_exp_alimony, 7, 38
			EMWriteScreen hc_exp_tax_dep, 8, 38
			EMWriteScreen hc_exp_other, 9, 38
			transmit
		END IF
		PF3
	END IF
	transmit
end function

function write_panel_to_MAXIS_DCEX(DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "DCEX")
	EMReadScreen panel_check, 4, 2, 52
	If panel_check <> "DCEX" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "DCEX")
	End If
	EMWriteScreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_DCEX, 1, 2, 78
	IF num_of_DCEX = "0" THEN
		EMWriteScreen "__", 20, 76
		Emwritescreen "NN", 20, 79
		transmit

		'---If the script is creating a new DCEX panel, it is going to enter this information into the DCEX main screen...
		EMWritescreen DCEX_provider, 6, 47
		EMWritescreen DCEX_reason, 7, 44
		EMWritescreen DCEX_subsidy, 8, 44
		IF len(DCEX_child_number1) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number1
		EMWritescreen DCEX_child_number1, 11, 29
		IF len(DCEX_child_number2) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number2
		EMWritescreen DCEX_child_number2, 12, 29
		IF len(DCEX_child_number3) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number3
		EMWritescreen DCEX_child_number3, 13, 29
		IF len(DCEX_child_number4) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number4
		EMWritescreen DCEX_child_number4, 14, 29
		IF len(DCEX_child_number5) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number5
		EMWritescreen DCEX_child_number5, 15, 29
		IF len(DCEX_child_number6) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number6
		EMWritescreen DCEX_child_number6, 16, 29
		EMWritescreen DCEX_child_number1_ver, 11, 41
		EMWritescreen DCEX_child_number2_ver, 12, 41
		EMWritescreen DCEX_child_number3_ver, 13, 41
		EMWritescreen DCEX_child_number4_ver, 14, 41
		EMWritescreen DCEX_child_number5_ver, 15, 41
		EMWritescreen DCEX_child_number6_ver, 16, 41
		EMWritescreen DCEX_child_number1_retro, 11, 48
		EMWritescreen DCEX_child_number2_retro, 12, 48
		EMWritescreen DCEX_child_number3_retro, 13, 48
		EMWritescreen DCEX_child_number4_retro, 14, 48
		EMWritescreen DCEX_child_number5_retro, 15, 48
		EMWritescreen DCEX_child_number6_retro, 16, 48
		EMWritescreen DCEX_child_number1_pro, 11, 63
		EMWritescreen DCEX_child_number2_pro, 12, 63
		EMWritescreen DCEX_child_number3_pro, 13, 63
		EMWritescreen DCEX_child_number4_pro, 14, 63
		EMWritescreen DCEX_child_number5_pro, 15, 63
		EMWritescreen DCEX_child_number6_pro, 16, 63
	ELSE
		PF9
		'---...if the script is PF9'ing, it is ONLY because it is going to enter information in the HC Expense sub-menu.
		'---Writing in the HC Expenses Est
		EMWriteScreen "X", 17, 55
		transmit

		DO			'---Waiting to make sure the HC Expense Est window has opened.
			EMReadScreen hc_expense_est, 10, 4, 41
		LOOP UNTIL hc_expense_est = "HC Expense"

		EMReadScreen hc_month, 17, 18, 62
		IF hc_month = "CURRENT MONTH + 1" THEN
			PF3
		ELSE
			IF len(DCEX_child_number1) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number1
			EMWritescreen DCEX_child_number1, 8, 39
			IF len(DCEX_child_number2) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number2
			EMWritescreen DCEX_child_number2, 9, 39
			IF len(DCEX_child_number3) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number3
			EMWritescreen DCEX_child_number3, 10, 39
			IF len(DCEX_child_number4) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number4
			EMWritescreen DCEX_child_number4, 11, 39
			IF len(DCEX_child_number5) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number5
			EMWritescreen DCEX_child_number5, 12, 39
			IF len(DCEX_child_number6) = 1 THEN DCEX_child_number1 = "0" & DCEX_child_number6
			EMWritescreen DCEX_child_number6, 13, 39
			EMWritescreen DCEX_child_number1_pro, 8, 49
			EMWritescreen DCEX_child_number2_pro, 9, 49
			EMWritescreen DCEX_child_number3_pro, 10, 49
			EMWritescreen DCEX_child_number4_pro, 11, 49
			EMWritescreen DCEX_child_number5_pro, 12, 49
			EMWritescreen DCEX_child_number6_pro, 13, 49
			transmit
			PF3
		END IF
	END IF
	transmit
end function

function write_panel_to_MAXIS_DFLN(conv_dt_1, conv_juris_1, conv_st_1, conv_dt_2, conv_juris_2, conv_st_2, rnd_test_dt_1, rnd_test_provider_1, rnd_test_result_1, rnd_test_dt_2, rnd_test_provider_2, rnd_test_result_2)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ conv_dt_1, conv_juris_1, conv_st_1, conv_dt_2, conv_juris_2, conv_st_2, rnd_test_dt_1, rnd_test_provider_1, rnd_test_result_1, rnd_test_dt_2, rnd_test_provider_2, rnd_test_result_2: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "DFLN")
	EMReadScreen panel_check, 4, 2, 45
	If panel_check <> "DFLN" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "DFLN")
	End If
	EMReadScreen num_of_DFLN, 1, 2, 78
	IF num_of_DFLN = "0" THEN
		EMWriteScreen reference_number, 20, 76
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF

	CALL create_MAXIS_friendly_date(conv_dt_1, 0, 6, 27)
	EMWriteScreen conv_juris_1, 6, 41
	EMWriteScreen conv_st_1, 6, 75
	IF conv_dt_2 <> "" THEN
		CALL create_MAXIS_friendly_date(conv_dt_2, 0, 7, 27)
		EMWriteScreen conv_juris_2, 7, 41
		EMWriteScreen conv_st_2, 7, 75
	END IF
	IF rnd_test_dt_1 <> "" THEN
		CALL create_MAXIS_friendly_date(rnd_test_dt_1, 0, 14, 27)
		EMWriteScreen rnd_test_provider_1, 14, 41
		EMWriteScreen rnd_test_result_1, 14, 75
		IF rnd_test_dt_2 <> "" THEN
			CALL create_MAXIS_friendly_date(rnd_test_dt_2, 0, 15, 27)
			EMWriteScreen rnd_test_provider_2, 15, 41
			EMWriteScreen rnd_test_result_2, 15, 75
		END IF
	END IF
end function

function write_panel_to_MAXIS_DIET(DIET_mfip_1, DIET_mfip_1_ver, DIET_mfip_2, DIET_mfip_2_ver, DIET_msa_1, DIET_msa_1_ver, DIET_msa_2, DIET_msa_2_ver, DIET_msa_3, DIET_msa_3_ver, DIET_msa_4, DIET_msa_4_ver)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ DIET_mfip_1, DIET_mfip_1_ver, DIET_mfip_2, DIET_mfip_2_ver, DIET_msa_1, DIET_msa_1_ver, DIET_msa_2, DIET_msa_2_ver, DIET_msa_3, DIET_msa_3_ver, DIET_msa_4, DIET_msa_4_ver: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "DIET")
	EMReadScreen panel_check, 4, 2, 48
	If panel_check <> "DIET" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "DIET")
	End If
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen "NN", 20, 79
	transmit

	EMWriteScreen DIET_mfip_1, 8, 40
	EMWriteScreen DIET_mfip_1_ver, 8, 51
	EMWriteScreen DIET_mfip_2, 9, 40
	EMWriteScreen DIET_mfip_2_ver, 9, 51
	EMWriteScreen DIET_msa_1, 11, 40
	EMWriteScreen DIET_msa_1_ver, 11, 51
	EMWriteScreen DIET_msa_2, 12, 40
	EMWriteScreen DIET_msa_2_ver, 12, 51
	EMWriteScreen DIET_msa_3, 13, 40
	EMWriteScreen DIET_msa_3_ver, 13, 51
	EMWriteScreen DIET_msa_4, 14, 40
	EMWriteScreen DIET_msa_4_ver, 14, 51
	transmit
end function

function write_panel_to_MAXIS_DISA(disa_begin_date, disa_end_date, disa_cert_begin, disa_cert_end, disa_wavr_begin, disa_wavr_end, disa_grh_begin, disa_grh_end, disa_cash_status, disa_cash_status_ver, disa_snap_status, disa_snap_status_ver, disa_hc_status, disa_hc_status_ver, disa_waiver, disa_1619, disa_drug_alcohol)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the disa panel in MAXIS)
'~~~~~ disa_begin_date, disa_end_date, disa_cert_begin, disa_cert_end, disa_wavr_begin, disa_wavr_end, disa_grh_begin, disa_grh_end, disa_cash_status, disa_cash_status_ver, disa_snap_status, disa_snap_status_ver, disa_hc_status, disa_hc_status_ver, disa_waiver, disa_1619, disa_drug_alcohol: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "DISA")  'navigates to the stat panel
	EMReadScreen panel_check, 4, 2, 45
	If panel_check <> "DISA" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "DISA")
	End If
	call create_panel_if_nonexistent
	IF disa_begin_date <> "" THEN
		call create_MAXIS_friendly_date(disa_begin_date, 0, 6, 47)  'enters the disability begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_begin_date), 6, 53
	END IF
	IF disa_end_date <> "" THEN
		call create_MAXIS_friendly_date(disa_end_date, 0, 6, 69)  'enters the disability end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_end_date), 6, 75
	END IF
	IF disa_cert_begin <> "" THEN
		call create_MAXIS_friendly_date(disa_cert_begin, 0, 7, 47)  'enters the disability certification begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_cert_begin), 7, 53
	END IF
	IF disa_cert_end <> "" THEN
		call create_MAXIS_friendly_date(disa_cert_end, 0, 7, 69)  'enters the disability certification end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_cert_end), 7, 75
	END IF
	IF disa_wavr_begin <> "" THEN
		call create_MAXIS_friendly_date(disa_wavr_begin, 0, 8, 47)  'enters the disability waiver begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_wavr_begin), 8, 53
	END IF
	IF disa_wavr_end <> "" THEN
		call create_MAXIS_friendly_date(disa_wavr_end, 0, 8, 69)  'enters the disability waiver end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_wavr_end), 8, 75
	END IF
	IF disa_grh_begin <> "" THEN
		call create_MAXIS_friendly_date(disa_grh_begin, 0, 9, 47)  'enters the disability grh begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_grh_begin), 9, 53
	END IF
	IF disa_grh_end <> "" THEN
		call create_MAXIS_friendly_date(disa_grh_end, 0, 9, 69)  'enters the disability grh end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_grh_end), 9, 75
	END IF
	Emwritescreen disa_cash_status, 11, 59  'enters status code for cash disa status
	Emwritescreen disa_cash_status_ver, 11, 69  'enters verification code for cash disa status
	Emwritescreen disa_snap_status, 12, 59  'enters status code for snap disa status
	Emwritescreen disa_snap_status_ver, 12, 69  'enters verification code for snap disa status
	Emwritescreen disa_hc_status, 13, 59  'enters status code for hc disa status
	Emwritescreen disa_hc_status_ver, 13, 69  'enters verification code for hc disa status
	Emwritescreen disa_waiver, 14, 59  'enters home and comminuty waiver code
	Emwritescreen disa_1619, 16, 59  'enters 1619 status
	Emwritescreen disa_drug_alcohol, 18, 69  'enters material drug & alcohol verification
end function

function write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, DSTT_HH_income_stop_date, DSTT_income_expected_amt)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ DSTT_ongoing_income, DSTT_HH_income_stop_date, DSTT_income_expected_amt: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "DSTT")
	EMReadScreen panel_check, 4, 2, 48
	If panel_check <> "DSTT" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "DSTT")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen DSTT_ongoing_income, 6, 69
	IF HH_income_stop_date <> "" THEN call create_MAXIS_friendly_date(HH_income_stop_date, 0, 9, 69)
	EMWriteScreen income_expected_amt, 12, 71
end function

function write_panel_to_MAXIS_EATS(eats_together, eats_boarder, eats_group_one, eats_group_two, eats_group_three)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ eats_together, eats_boarder, eats_group_one, eats_group_two, eats_group_three: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	IF reference_number = "01" THEN
		call navigate_to_MAXIS_screen("STAT", "EATS")
		EMReadScreen panel_check, 4, 2, 47
		If panel_check <> "EATS" Then
			MAXIS_background_check
			call navigate_to_MAXIS_screen("STAT", "EATS")
		End If
		call create_panel_if_nonexistent
		EMWriteScreen eats_together, 4, 72
		EMWriteScreen eats_boarder, 5, 72
		IF ucase(eats_together) = "N" THEN
			EMWriteScreen "01", 13, 28
			eats_group_one = replace(eats_group_one, " ", "")
			eats_group_one = split(eats_group_one, ",")
			eats_col = 39
			FOR EACH eats_household_member IN eats_group_one
				EMWriteScreen eats_household_member, 13, eats_col
				eats_col = eats_col + 4
			NEXT
			EMWriteScreen "02", 14, 28
			eats_group_two = replace(eats_group_two, " ", "")
			eats_group_two = split(eats_group_two, ",")
			eats_col = 39
			FOR EACH eats_household_member IN eats_group_two
				EMWriteScreen eats_household_member, 14, eats_col
				eats_col = eats_col + 4
			NEXT
			IF eats_group_three <> "" THEN
				EMWriteScreen "03", 15, 28
				eats_group_three = replace(eats_group_three, " ", "")
				eats_group_three = split(eats_group_three, ",")
				eats_col = 39
				FOR EACH eats_household_member IN eats_group_three
					EMWriteScreen eats_household_member, 15, eats_col
					eats_col = eats_col + 4
				NEXT
			END IF
		END IF
	transmit
	END IF
end function

function write_panel_to_MAXIS_EMMA(EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "EMMA")
	EMReadScreen panel_check, 4, 2, 53
	If panel_check <> "EMMA" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "EMMA")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen EMMA_medical_emergency, 6, 46
	EMWriteScreen EMMA_health_consequence, 8, 46
	EMWriteScreen EMMA_verification, 10, 46
	call create_MAXIS_friendly_date(EMMA_begin_date, 0, 12, 46)
	IF EMMA_end_date <> "" THEN call create_MAXIS_friendly_date(EMMA_end_date, 0, 14, 46)
end function

function write_panel_to_MAXIS_EMPS(EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "EMPS")
	EMReadScreen panel_check, 4, 2, 50
	If panel_check <> "EMPS" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "EMPS")
	End If
	call create_panel_if_nonexistent
	If EMPS_orientation_date <> "" then call create_MAXIS_friendly_date(EMPS_orientation_date, 0, 5, 39) 'enter orientation date
	EMWritescreen left(EMPS_orientation_attended, 1), 5, 65
	EMWritescreen EMPS_good_cause, 5, 79
	If EMPS_sanc_begin <> "" then call create_MAXIS_friendly_date(EMPS_sanc_begin, 1, 6, 39) 'Sanction begin date
	If EMPS_sanc_end <> "" then call create_MAXIS_friendly_date(EMPS_sanc_end, 1, 6, 65) 'Sanction end date
	EMWritescreen left(EMPS_memb_at_home, 1), 8, 76
	EMWritescreen left(EMPS_care_family, 1), 9, 76
	EMWritescreen left(EMPS_crisis, 1), 10, 76
	EMWritescreen EMPS_hard_employ, 11, 76
	EMWritescreen left(EMPS_under1, 1), 12, 76 'child under 1 exemption
	EMWritescreen "n", 13, 76 'enters n for child under 12 weeks
	If EMPS_DWP_date <> "" then call create_MAXIS_friendly_date(EMPS_DWP_date, 1, 17, 40) 'DWP plan date
	'This populates the child under 1 popup if needed
	IF ucase(left(EMPS_under1, 1)) = "Y" THEN
		EMReadScreen month_to_use, 2, 20, 55
		EMReadScreen start_year, 2, 20, 58
		Emwritescreen "x", 12, 39
		Transmit
		EMReadScreen check_for_blank, 2, 7, 22 'makes sure the popup isn't already filled out
		month_to_use = cint(month_to_use)
		start_year = cint("20" & start_year)
		popup_row = 7 'setting initial starting point for the popup
		popup_col = 22
		IF check_for_blank <> "  " THEN 'blank popup, fill it out!
			FOR i = 1 to 12
				IF month_to_use > 12 THEN 'handling the year change
					popup_month = month_to_use - 12
					year_to_use = start_year +1
				ELSE
					popup_month = month_to_use
					year_to_use = start_year
				END IF
				IF len(popup_month) = 1 THEN popup_month = "0" & popup_month 'formatting to two digit month
				Emwritescreen popup_month, popup_row, popup_col
				Emwritescreen year_to_use, popup_row, popup_col + 5
				popup_col = popup_col + 11
				month_to_use = month_to_use + 1
				IF popup_col > 55 THEN 'This moves to the next row if necessary
					popup_col = 22
					popup_row = popup_row + 1
				END IF
			NEXT
			PF3 'closing the popup
		END IF
	END IF
end function

function write_panel_to_MAXIS_FACI(FACI_vendor_number, FACI_name, FACI_type, FACI_FS_eligible, FACI_FS_facility_type, FACI_date_in, FACI_date_out)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ FACI_vendor_number, FACI_name, FACI_type, FACI_FS_eligible, FACI_FS_facility_type, FACI_date_in, FACI_date_out: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "FACI")
	EMReadScreen panel_check, 4, 2, 44
	If panel_check <> "FACI" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "FACI")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen FACI_vendor_number, 5, 43
	EMWriteScreen FACI_name, 6, 43
	EMWriteScreen FACI_type, 7, 43
	EMWriteScreen FACI_FS_eligible, 8, 43
	If FACI_date_in <> "" then
		call create_MAXIS_friendly_date(FACI_date_in, 0, 14, 47)
		EMWriteScreen datepart("YYYY", FACI_date_in), 14, 53
	End if
	If FACI_date_out <> "" then
		call create_MAXIS_friendly_date(FACI_date_out, 0, 14, 71)
		EMWriteScreen datepart("YYYY", FACI_date_out), 14, 77
	End if
	transmit
	transmit
end function

function write_panel_to_MAXIS_FMED(FMED_medical_mileage, FMED_1_type, FMED_1_verif, FMED_1_ref_num, FMED_1_category, FMED_1_begin, FMED_1_end, FMED_1_amount, FMED_2_type, FMED_2_verif, FMED_2_ref_num, FMED_2_category, FMED_2_begin, FMED_2_end, FMED_2_amount, FMED_3_type, FMED_3_verif, FMED_3_ref_num, FMED_3_category, FMED_3_begin, FMED_3_end, FMED_3_amount, FMED_4_type, FMED_4_verif, FMED_4_ref_num, FMED_4_category, FMED_4_begin, FMED_4_end, FMED_4_amount)
'--- This function writes to MAXIS in Krabappel only (pulls FMED information from the Excel file. This function can handle up to 4 FMED rows per client.)
'~~~~~ FMED_medical_mileage, FMED_1_type, FMED_1_verif, FMED_1_ref_num, FMED_1_category, FMED_1_begin, FMED_1_end, FMED_1_amount, FMED_2_type, FMED_2_verif, FMED_2_ref_num, FMED_2_category, FMED_2_begin, FMED_2_end, FMED_2_amount, FMED_3_type, FMED_3_verif, FMED_3_ref_num, FMED_3_category, FMED_3_begin, FMED_3_end, FMED_3_amount, FMED_4_type, FMED_4_verif, FMED_4_ref_num, FMED_4_category, FMED_4_begin, FMED_4_end, FMED_4_amount: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "FMED")
	EMReadScreen panel_check, 4, 2, 58
	If panel_check <> "FMED" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "FMED")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	EMReadScreen num_of_FMED, 1, 2, 78
	IF num_of_FMED = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF

	'Determining where to start writing...
	FMED_row = 9
	DO
		EMReadScreen FMED_available, 2, FMED_row, 25
		IF FMED_available <> "__" THEN FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	LOOP UNTIL FMED_available = "__"

	IF FMED_1_type <> "" THEN
		EMWriteScreen FMED_1_type, FMED_row, 25
			IF FMED_1_type = "12" THEN
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_1_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_1_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_1_verif, FMED_row, 32
		EMWriteScreen FMED_1_ref_num, FMED_row, 38
		EMWriteScreen FMED_1_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_1_begin)			'Turning the value in FMED_1_begin and FMED_1_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_1_begin), 2), FMED_row, 53
		IF FMED_1_end <> "" THEN
			FMED_month = DatePart("M", FMED_1_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_1_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_1_amount, FMED_row, 70

		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF

	IF FMED_2_type <> "" THEN
		EMWriteScreen FMED_2_type, FMED_row, 25
			IF FMED_2_type = "12" THEN
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_2_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_2_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_2_verif, FMED_row, 32
		EMWriteScreen FMED_2_ref_num, FMED_row, 38
		EMWriteScreen FMED_2_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_2_begin)			'Turning the value in FMED_2_begin and FMED_2_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_2_begin), 2), FMED_row, 53
		IF FMED_2_end <> "" THEN
			FMED_month = DatePart("M", FMED_2_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_2_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_2_amount, FMED_row, 70

		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF

	IF FMED_3_type <> "" THEN
		EMWriteScreen FMED_3_type, FMED_row, 25
			IF FMED_3_type = "12" THEN
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_3_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_3_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_3_verif, FMED_row, 32
		EMWriteScreen FMED_3_ref_num, FMED_row, 38
		EMWriteScreen FMED_3_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_3_begin)			'Turning the value in FMED_3_begin and FMED_3_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_3_begin), 2), FMED_row, 53
		IF FMED_3_end <> "" THEN
			FMED_month = DatePart("M", FMED_3_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_3_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_3_amount, FMED_row, 70

		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF

	IF FMED_4_type <> "" THEN
		EMWriteScreen FMED_4_type, FMED_row, 25
			IF FMED_4_type = "12" THEN
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_4_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_4_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_4_verif, FMED_row, 32
		EMWriteScreen FMED_4_ref_num, FMED_row, 38
		EMWriteScreen FMED_4_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_4_begin)			'Turning the value in FMED_4_begin and FMED_4_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_4_begin), 2), FMED_row, 53
		IF FMED_4_end <> "" THEN
			FMED_month = DatePart("M", FMED_4_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_4_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_4_amount, FMED_row, 70
	END IF

	transmit
end function

function write_panel_to_MAXIS_HCRE(hcre_appl_addnd_date_input,hcre_retro_months_input,hcre_recvd_by_service_date_input)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ hcre_appl_addnd_date_input,hcre_retro_months_input,hcre_recvd_by_service_date_input: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","HCRE")
	EMReadScreen panel_check, 4, 2, 50
	If panel_check <> "HCRE" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "HCRE")
	End If
	call create_panel_if_nonexistent
	'Converting the Appl Addendum Date into a usable format
	call MAXIS_dater(hcre_appl_addnd_date_input, hcre_appl_addnd_date_output, "HCRE Addendum Date")
	'Converting the Received by service date into a usable format
	call MAXIS_dater(hcre_recvd_by_service_date_input, hcre_recvd_by_service_date_output, "received by Service Date")
	'Converts Retro Months Input into a negative
	hcre_retro_months_input = (Abs(hcre_retro_months_input)*(-1))
	call add_months(hcre_retro_months_input,hcre_appl_addnd_date_output,hcre_retro_date_output)
	row = 1
	col = 1
	EMSearch "* " & reference_number, row, col
		'Appl Addendum Request Date
	EMWriteScreen left(hcre_appl_addnd_date_output,2)		, row, col + 29
	EMWriteScreen mid(hcre_recvd_by_service_date_input,4,2)	, row, col + 32
	EMWriteScreen right(hcre_appl_addnd_date_output,2)		, row, col + 35
		'Coverage Request Date
	EMWriteScreen left(hcre_retro_date_output,2)	, row, col + 42
	EMWriteScreen right(hcre_retro_date_output,2)	, row, col + 45
		'Recv By Sv Date
	EMWriteScreen left(hcre_recvd_by_service_date_output,2)	, row, col + 51
	EMWriteScreen mid(hcre_recvd_by_service_date_output,4,2), row, col + 54
	EMWriteScreen right(hcre_recvd_by_service_date_output,2), row, col + 57
	transmit
end function

function write_panel_to_MAXIS_HEST(HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "HEST")
	EMReadScreen panel_check, 4, 2, 53
	If panel_check <> "HEST" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "HEST")
	End If
	call create_panel_if_nonexistent
	Emwritescreen "01", 6, 40
	call create_MAXIS_friendly_date(HEST_FS_choice_date, 0, 7, 40)
	EMWritescreen HEST_first_month, 8, 61
	'Filling in the #/FS units field (always 01)
	If ucase(left(HEST_heat_air_retro, 1)) = "Y" then EMWritescreen "01", 13, 42
	If ucase(left(HEST_heat_air_pro, 1)) = "Y" then EMWritescreen "01", 13, 68
	If ucase(left(HEST_electric_retro, 1)) = "Y" then EMWritescreen "01", 14, 42
	If ucase(left(HEST_electric_pro, 1)) = "Y" then EMWritescreen "01", 14, 68
	If ucase(left(HEST_phone_retro, 1)) = "Y" then EMWritescreen "01", 15, 42
	If ucase(left(HEST_phone_pro, 1)) = "Y" then EMWritescreen "01", 15, 68
	EMWritescreen left(HEST_heat_air_retro, 1), 13, 34
	EMWritescreen left(HEST_electric_retro, 1), 14, 34
	EMWritescreen left(HEST_phone_retro, 1), 15, 34
	EMWritescreen left(HEST_heat_air_pro, 1), 13, 60
	EMWritescreen left(HEST_electric_pro, 1), 14, 60
	EMWritescreen left(HEST_phone_pro, 1), 15, 60
	transmit
end function

function write_panel_to_MAXIS_IMIG(IMIG_imigration_status, IMIG_entry_date, IMIG_status_date, IMIG_status_ver, IMIG_status_LPR_adj_from, IMIG_nationality, IMIG_40_soc_sec, IMIG_40_soc_sec_verif, IMIG_battered_spouse_child, IMIG_battered_spouse_child_verif, IMIG_military_status, IMIG_military_status_verif, IMIG_hmong_lao_nat_amer, IMIG_st_prog_esl_ctzn_coop, IMIG_st_prog_esl_ctzn_coop_verif, IMIG_fss_esl_skills_training)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ IMIG_imigration_status, IMIG_entry_date, IMIG_status_date, IMIG_status_ver, IMIG_status_LPR_adj_from, IMIG_nationality, IMIG_40_soc_sec, IMIG_40_soc_sec_verif, IMIG_battered_spouse_child, IMIG_battered_spouse_child_verif, IMIG_military_status, IMIG_military_status_verif, IMIG_hmong_lao_nat_amer, IMIG_st_prog_esl_ctzn_coop, IMIG_st_prog_esl_ctzn_coop_verif, IMIG_fss_esl_skills_training: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "IMIG")
	EMReadScreen panel_check, 4, 2, 49
	If panel_check <> "IMIG" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "IMIG")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	call create_MAXIS_friendly_date(APPL_date, 0, 5, 45)						'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
	EMWriteScreen datepart("yyyy", APPL_date), 5, 51
	EMWriteScreen IMIG_imigration_status, 6, 45							'Writes imig status
	IF IMIG_entry_date <> "" THEN
		call create_MAXIS_friendly_date(IMIG_entry_date, 0, 7, 45)			'Enters year as a 2 digit number, so have to modify manually
		EMWriteScreen datepart("yyyy", IMIG_entry_date), 7, 51
	END IF
	IF IMIG_status_date <> "" THEN
		call create_MAXIS_friendly_date(IMIG_status_date, 0, 7, 71)			'Enters year as a 2 digit number, so have to modify manually
		EMWriteScreen datepart("yyyy", IMIG_status_date), 7, 77
	END IF
	EMWriteScreen IMIG_status_ver, 8, 45								'Enters status ver
	EMWriteScreen IMIG_status_LPR_adj_from, 9, 45						'Enters status LPR adj from
	EMWriteScreen IMIG_nationality, 10, 45								'Enters nationality
	EMwritescreen IMIG_40_soc_sec, 13, 56								'Enters info about Social Security Credits
	EMwritescreen IMIG_40_soc_sec_verif, 13, 71
	EMwritescreen IMIG_battered_spouse_child, 14, 56					'Enters info about Battered Child/Spouse claims
	EMwritescreen IMIG_battered_spouse_child_verif, 14, 71
	EMwritescreen IMIG_military_status, 15, 56 							'Enters info about possible military status
	EMwritescreen IMIG_military_status_verif, 15, 71
	EMwritescreen IMIG_hmong_lao_nat_amer, 16, 56 						'Enters status of particular nationalities/identity
	EMwritescreen IMIG_st_prog_esl_ctzn_coop, 17, 56 					'Enters information about ESL/Citizen cooperation status
	EMwritescreen IMIG_st_prog_esl_ctzn_coop_verif, 17, 71
	EMwritescreen IMIG_fss_esl_skills_training, 18, 56 					'Enters information about ESL Skills course
	transmit
	transmit
end function

function write_panel_to_MAXIS_INSA(insa_pers_coop_ohi, insa_good_cause_status, insa_good_cause_cliam_date, insa_good_cause_evidence, insa_coop_cost_effect, insa_insur_name, insa_prescrip_drug_cover, insa_prescrip_end_date, insa_persons_covered)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ insa_pers_coop_ohi, insa_good_cause_status, insa_good_cause_cliam_date, insa_good_cause_evidence, insa_coop_cost_effect, insa_insur_name, insa_prescrip_drug_cover, insa_prescrip_end_date, insa_persons_covered: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","INSA")
	EMReadScreen panel_check, 4, 2, 49
	If panel_check <> "INSA" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "INSA")
	End If
	call create_panel_if_nonexistent

	EMWriteScreen insa_pers_coop_ohi, 4, 62
	EMWriteScreen insa_good_cause_status, 5, 62
	If insa_good_cause_cliam_date <> "" then CALL create_MAXIS_friendly_date(insa_good_cause_cliam_date, 0, 6, 62)
	EMWriteScreen insa_good_cause_evidence, 7, 62
	EMWriteScreen insa_coop_cost_effect, 8, 62
	EMWriteScreen insa_insur_name, 10, 38
	EMWriteScreen insa_prescrip_drug_cover, 11, 62
	If insa_prescrip_end_date <> "" then CALL create_MAXIS_friendly_date(insa_prescrip_end_date, 0, 12, 62)

	'Adding persons covered
	insa_row = 15
	insa_col = 30

	insa_persons_covered = replace(insa_persons_covered, " ", "")
	insa_persons_covered = split(insa_persons_covered, ",")

	FOR EACH insa_peep IN insa_persons_covered
		EMWriteScreen insa_peep, insa_row, insa_col
		insa_col = insa_col + 4
		IF insa_col = 70 THEN
			insa_col = 30
			insa_row = 16
		END IF
	NEXT
	transmit
end function

function write_panel_to_MAXIS_JOBS(jobs_number, jobs_inc_type, jobs_inc_verif, jobs_employer_name, jobs_inc_start, jobs_wkly_hrs, jobs_hrly_wage, jobs_pay_freq)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ jobs_number, jobs_inc_type, jobs_inc_verif, jobs_employer_name, jobs_inc_start, jobs_wkly_hrs, jobs_hrly_wage, jobs_pay_freq: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "JOBS")
	EMReadScreen panel_check, 4, 2, 45
	If panel_check <> "JOBS" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "JOBS")
	End If
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen jobs_number, 20, 79
	transmit

	EMReadScreen does_not_exist, 14, 24, 13
	IF does_not_exist = "DOES NOT EXIST" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF

	EMWriteScreen jobs_inc_type, 5, 34
	EMWriteScreen jobs_inc_verif, 6, 34

	EMWriteScreen jobs_employer_name, 7, 42
	call create_MAXIS_friendly_date(jobs_inc_start, 0, 9, 35)
	EMWriteScreen jobs_pay_freq, 18, 35

	'===== navigates to the SNAP PIC to update the PIC =====
	EMWriteScreen "X", 19, 38
	transmit
	DO
		EMReadScreen at_snap_pic, 12, 3, 22
	LOOP UNTIL at_snap_pic = "Food Support"
	EMReadScreen jobs_pic_wages_per_pp, 7, 17, 57
	EMReadScreen pic_info_exists, 8, 18, 57
	pic_info_exists = trim(pic_info_exists)
	IF pic_info_exists = "" THEN
		call create_MAXIS_friendly_date(date, 0, 5, 34)
		EMWriteScreen jobs_pay_freq, 5, 64
		EMWriteScreen jobs_wkly_hrs, 8, 64
		EMWriteScreen jobs_hrly_wage, 9, 66
		transmit
		transmit
		EMReadScreen jobs_pic_hrs_per_pp, 6, 16, 51
		EMReadScreen jobs_pic_wages_per_pp, 7, 17, 57
	END IF
	transmit		'<=====navigates out of the PIC

	'=====the following bit is for the retrospective & prospective pay dates=====
	EMReadScreen bene_month, 2, 20, 55
	EMReadScreen bene_year, 2, 20, 58
	benefit_month = bene_month & "/01/" & bene_year
	retro_month = DatePart("M", DateAdd("M", -2, benefit_month))
	IF len(retro_month) <> 2 THEN retro_month = "0" & retro_month
	retro_year = right(DatePart("YYYY", DateAdd("M", -2, benefit_month)), 2)

	EMWriteScreen retro_month, 12, 25
	EMWriteScreen retro_year, 12, 31
	EMWriteScreen bene_month, 12, 54
	EMWriteScreen bene_year, 12, 60

	IF pic_info_exists = "" THEN 		'---If the PIC is blank, the information needs to be added to the main JOBS panel as well.
		EMWriteScreen "05", 12, 28
		EMWriteScreen jobs_pic_wages_per_pp, 12, 38
		EMWriteScreen "05", 12, 57
		EMWriteScreen jobs_pic_wages_per_pp, 12, 67
		EMWriteScreen Int(jobs_pic_hrs_per_pp), 18, 43
		EMWriteScreen Int(jobs_pic_hrs_per_pp), 18, 72
	END IF

	IF jobs_pay_freq = 2 OR jobs_pay_freq = 3 THEN
		EMWriteScreen retro_month, 13, 25
		EMWriteScreen retro_year, 13, 31
		EMWriteScreen bene_month, 13, 54
		EMWriteScreen bene_year, 13, 60

		IF pic_info_exists = "" THEN
			EMWriteScreen "19", 13, 28
			EMWriteScreen jobs_pic_wages_per_pp, 13, 38
			EMWriteScreen "19", 13, 57
			EMWriteScreen jobs_pic_wages_per_pp, 13, 67
			EMWriteScreen Int(2 * jobs_pic_hrs_per_pp), 18, 43
			EMWriteScreen Int(2 * jobs_pic_hrs_per_pp), 18, 72
		END IF
	ELSEIF jobs_pay_freq = 4 THEN
		EMWriteScreen retro_month, 13, 25
		EMWriteScreen retro_year, 13, 31
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen bene_month, 13, 54
		EMWriteScreen bene_year, 13, 60
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen bene_month, 15, 54
		EMWriteScreen bene_year, 15, 60

		IF pic_info_exists = "" THEN
			EMWriteScreen "12", 13, 28
			EMWriteScreen jobs_pic_wages_per_pp, 13, 38
			EMWriteScreen "19", 14, 28
			EMWriteScreen jobs_pic_wages_per_pp, 14, 38
			EMWriteScreen "26", 15, 28
			EMWriteScreen jobs_pic_wages_per_pp, 15, 38
			EMWriteScreen "12", 13, 57
			EMWriteScreen jobs_pic_wages_per_pp, 13, 67
			EMWriteScreen "19", 14, 57
			EMWriteScreen jobs_pic_wages_per_pp, 14, 67
			EMWriteScreen "26", 15, 57
			EMWriteScreen jobs_pic_wages_per_pp, 15, 67
			EMWriteScreen Int(4 * jobs_pic_hrs_per_pp), 18, 43
			EMWriteScreen Int(4 * jobs_pic_hrs_per_pp), 18, 72
		END IF
	END IF

	'=====determines if the benefit month is current month + 1 and dumps information into the HC income estimator
	IF (bene_month * 1) = (datepart("M", DATE) + 1) THEN		'<===== "bene_month * 1" is needed to convert bene_month from a string to numeric.
		EMReadScreen HC_income_est_check, 3, 19, 63 'reading to find the HC income estimator is moving 6/1/16, to account for if it only affects future months we are reading to find the HC inc EST
		IF HC_income_est_check = "Est" Then 'this is the old position
			EMWriteScreen "x", 19, 54
		ELSE								'this is the new position
			EMWriteScreen "x", 19, 48
		END IF
		transmit

		DO
			EMReadScreen hc_inc_est, 9, 9, 43
		LOOP UNTIL hc_inc_est = "HC Income"

		EMWriteScreen jobs_pic_wages_per_pp, 11, 63
		transmit
		transmit
	END IF
end function

function write_panel_to_MAXIS_MEDI(SSN_first, SSN_mid, SSN_last, MEDI_claim_number_suffix, MEDI_part_A_premium, MEDI_part_B_premium, MEDI_part_A_begin_date, MEDI_part_B_begin_date, MEDI_apply_prem_to_spdn, MEDI_apply_prem_end_date)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ SSN_first, SSN_mid, SSN_last, MEDI_claim_number_suffix, MEDI_part_A_premium, MEDI_part_B_premium, MEDI_part_A_begin_date, MEDI_part_B_begin_date, MEDI_apply_prem_to_spdn, MEDI_apply_prem_end_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "MEDI")
	EMReadScreen panel_check, 4, 2, 44
	If panel_check <> "MEDI" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "MEDI")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen SSN_first, 6, 39				'Next three lines pulled
	EMWriteScreen SSN_mid, 6, 43
	EMWriteScreen SSN_last, 6, 46
	EMWriteScreen MEDI_claim_number_suffix, 6, 51
	EMWriteScreen MEDI_part_A_premium, 7, 46
	EMWriteScreen MEDI_part_B_premium, 7, 73
	If MEDI_part_A_begin_date <> "" then call create_MAXIS_friendly_date(MEDI_part_A_begin_date, 0, 15, 24)
	If MEDI_part_B_begin_date <> "" then call create_MAXIS_friendly_date(MEDI_part_B_begin_date, 0, 15, 54)
	EMWriteScreen MEDI_apply_prem_to_spdn, 11, 71
	IF MEDI_apply_prem_end_date <> "" THEN
		EMWriteScreen left(MEDI_apply_prem_end_date, 2), 12, 71
		EMWriteScreen right(MEDI_apply_prem_end_date, 2), 12, 74
	END IF
	transmit
	transmit
end function

function write_panel_to_MAXIS_MMSA(mmsa_liv_arr, mmsa_cont_elig, mmsa_spous_inc, mmsa_shared_hous)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ mmsa_liv_arr, mmsa_cont_elig, mmsa_spous_inc, mmsa_shared_hous: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	IF mmsa_liv_arr <> "" THEN
		call navigate_to_MAXIS_screen("STAT", "MMSA")
		EMReadScreen panel_check, 4, 2, 55
		If panel_check <> "MMSA" Then
			MAXIS_background_check
			call navigate_to_MAXIS_screen("STAT", "MMSA")
		End If
		EMWriteScreen "NN", 20, 79
		transmit
		EMWriteScreen mmsa_liv_arr, 7, 54
		EMWriteScreen mmsa_cont_elig, 9, 54
		EMWriteScreen mmsa_spous_inc, 12, 62
		EMWriteScreen mmsa_shared_hous, 14, 62
		transmit
	END IF
end function

function write_panel_to_MAXIS_MSUR(msur_begin_date)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ msur_begin_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","MSUR")
	EMReadScreen panel_check, 4, 2, 54
	If panel_check <> "MSUR" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "MSUR")
	End If
	call create_panel_if_nonexistent

	'msur_begin_date This is the date MSUR began for this client
	row = 7
	DO
		EMReadScreen available_space, 2, row, 36
		IF available_space = "__" THEN
			row = row + 1
		ELSE
			EXIT DO
		END IF
	LOOP UNTIL available_space <> "__"

	CALL create_MAXIS_friendly_date(msur_begin_date, 0, row, 36)
	Emwritescreen DatePart("YYYY", msure_begin_date), row, 42
	transmit
end function

function write_panel_to_MAXIS_OTHR(othr_type, othr_cash_value, othr_cash_value_ver, othr_owed, othr_owed_ver, othr_date, othr_cash_count, othr_SNAP_count, othr_HC_count, othr_IV_count, othr_joint_owner, othr_share_ratio)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the othr panel in MAXIS)
'~~~~~ othr_type, othr_cash_value, othr_cash_value_ver, othr_owed, othr_owed_ver, othr_date, othr_cash_count, othr_SNAP_count, othr_HC_count, othr_IV_count, othr_joint_owner, othr_share_ratio: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "OTHR")  'navigates to the stat panel
	EMReadScreen panel_check, 4, 2, 46
	If panel_check <> "OTHR" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "OTHR")
	End If
	call create_panel_if_nonexistent
	Emwritescreen othr_type, 6, 40  'enters other asset type
	IF othr_cash_value = "" THEN othr_cash_value = 0
	Emwritescreen othr_cash_value, 8, 40  'enters cash value of asset
	Emwritescreen othr_cash_value_ver, 8, 57  'enters cash value verification code
	IF othr_owed = "" THEN othr_owed = 0
	Emwritescreen othr_owed, 9, 40  'enters amount owed value
	Emwritescreen othr_owed_ver, 9, 57  'enters amount owed verification code
	call create_MAXIS_friendly_date(othr_date, 0, 10, 39)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen othr_cash_count, 12, 50  'enters y/n if counted for cash
	Emwritescreen othr_SNAP_count, 12, 57  'enters y/n if counted for snap
	Emwritescreen othr_HC_count, 12, 64  'enters y/n if counted for hc
	Emwritescreen othr_IV_count, 12, 73  'enters y/n if counted for iv
	Emwritescreen othr_joint_owner, 13, 44  'enters if it is a jointly owned other asset
	Emwritescreen left(othr_share_ratio, 1), 15, 50  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(othr_share_ratio, 1), 15, 54  'enters the ratio of ownership using the right 1 digit of what is entered into the file
end function

function write_panel_to_MAXIS_PARE(appl_date, reference_number, PARE_child_1, PARE_child_1_relation, PARE_child_1_verif, PARE_child_2, PARE_child_2_relation, PARE_child_2_verif, PARE_child_3, PARE_child_3_relation, PARE_child_3_verif, PARE_child_4, PARE_child_4_relation, PARE_child_4_verif, PARE_child_5, PARE_child_5_relation, PARE_child_5_verif, PARE_child_6, PARE_child_6_relation, PARE_child_6_verif)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ appl_date, reference_number, PARE_child_1, PARE_child_1_relation, PARE_child_1_verif, PARE_child_2, PARE_child_2_relation, PARE_child_2_verif, PARE_child_3, PARE_child_3_relation, PARE_child_3_verif, PARE_child_4, PARE_child_4_relation, PARE_child_4_verif, PARE_child_5, PARE_child_5_relation, PARE_child_5_verif, PARE_child_6, PARE_child_6_relation, PARE_child_6_verif: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "PARE")
	EMReadScreen panel_check, 4, 2, 43
	If panel_check <> "PARE" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "PARE")
	End If
	CALL write_value_and_transmit(reference_number, 20, 76)
	EMReadScreen num_of_PARE, 1, 2, 78
	IF num_of_PARE = "0" THEN
		CALL write_value_and_transmit("NN", 20, 79)
	ELSE
		PF9
	END IF
	CALL create_MAXIS_friendly_date(appl_date, 0, 5, 37)
	EMWriteScreen DatePart("YYYY", appl_date), 5, 43

	IF len(PARE_child_1) = 1 THEN PARE_child_1 = "0" & PARE_child_1
	IF len(PARE_child_2) = 1 THEN PARE_child_1 = "0" & PARE_child_2
	IF len(PARE_child_3) = 1 THEN PARE_child_1 = "0" & PARE_child_3
	IF len(PARE_child_4) = 1 THEN PARE_child_1 = "0" & PARE_child_4
	IF len(PARE_child_5) = 1 THEN PARE_child_1 = "0" & PARE_child_5
	IF len(PARE_child_6) = 1 THEN PARE_child_1 = "0" & PARE_child_6
	EMWritescreen PARE_child_1, 8, 24
	EMWritescreen PARE_child_1_relation, 8, 53
	EMWritescreen PARE_child_1_verif, 8, 71
	EMWritescreen PARE_child_2, 9, 24
	EMWritescreen PARE_child_2_relation, 9, 53
	EMWritescreen PARE_child_2_verif, 9, 71
	EMWritescreen PARE_child_3, 10, 24
	EMWritescreen PARE_child_3_relation, 10, 53
	EMWritescreen PARE_child_3_verif, 10, 71
	EMWritescreen PARE_child_4, 11, 24
	EMWritescreen PARE_child_4_relation, 11, 53
	EMWritescreen PARE_child_4_verif, 11, 71
	EMWritescreen PARE_child_5, 12, 24
	EMWritescreen PARE_child_5_relation, 12, 53
	EMWritescreen PARE_child_5_verif, 12, 71
	EMWritescreen PARE_child_6, 13, 24
	EMWritescreen PARE_child_6_relation, 13, 53
	EMWritescreen PARE_child_6_verif, 13, 71
	transmit
end function

function write_panel_to_MAXIS_PBEN(pben_referal_date, pben_type, pben_appl_date, pben_appl_ver, pben_IAA_date, pben_disp)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the pben panel in MAXIS)
'~~~~~ pben_referal_date, pben_type, pben_appl_date, pben_appl_ver, pben_IAA_date, pben_disp: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "PBEN")  'navigates to the stat panel
	EMReadScreen panel_check, 4, 2, 49
	If panel_check <> "PBEN" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "PBEN")
	End If
	call create_panel_if_nonexistent
	Emreadscreen pben_row_check, 2, 8, 24  'reads the MAXIS screen to find out if the PBEN row has already been used.
	If pben_row_check = "__" THEN   'if the row is blank it enters it in the 8th row.
		Emwritescreen pben_type, 8, 24  'enters pben type code
		call create_MAXIS_friendly_date(pben_referal_date, 0, 8, 40)  'enters referal date in MAXIS friendly format mm/dd/yy
		call create_MAXIS_friendly_date(pben_appl_date, 0, 8, 51)  'enters appl date in  MAXIS friendly format mm/dd/yy
		Emwritescreen pben_appl_ver, 8, 62  'enters appl verification code
		call create_MAXIS_friendly_date(pben_IAA_date, 0, 8, 66)  'enters IAA date in MAXIS friendly format mm/dd/yy
		Emwritescreen pben_disp, 8, 77  'enters the status of pben application
	else
		EMreadscreen pben_row_check, 2, 9, 24  'if row 8 is filled already it will move to row 9 and see if it has been used.
		IF pben_row_check = "__" THEN  'if the 9th row is blank it enters the information there.
		'second pben row
			Emwritescreen pben_type, 9, 24
			call create_MAXIS_friendly_date(pben_referal_date, 0, 9, 40)
			call create_MAXIS_friendly_date(pben_appl_date, 0, 9, 51)
			Emwritescreen pben_appl_ver, 9, 62
			call create_MAXIS_friendly_date(pben_IAA_date, 0, 9, 66)
			Emwritescreen pben_disp, 9, 77
		else
		Emreadscreen pben_row_check, 2, 10, 24  'if row 8 is filled already it will move to row 9 and see if it has been used.
			IF pben-row_check = "__" THEN  'if the 9th row is blank it enters the information there.
			'third pben row
				Emwritescreen pben_type, 10, 24
				call create_MAXIS_friendly_date(pben_referal_date, 0, 10, 40)
				call create_MAXIS_friendly_date(pben_appl_date, 0, 10, 51)
				Emwritescreen pben_appl_ver, 10, 62
				call create_MAXIS_friendly_date(pben_IAA_date, 0, 10, 66)
				Emwritescreen pben_disp, 10, 77
			END IF
		END IF
	END IF
end function

function write_panel_to_MAXIS_PDED(PDED_wid_deduction, PDED_adult_child_disregard, PDED_wid_disregard, PDED_unea_income_deduction_reason, PDED_unea_income_deduction_value, PDED_earned_income_deduction_reason, PDED_earned_income_deduction_value, PDED_ma_epd_inc_asset_limit, PDED_guard_fee, PDED_rep_payee_fee, PDED_other_expense, PDED_shel_spcl_needs, PDED_excess_need, PDED_restaurant_meals)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ PDED_wid_deduction, PDED_adult_child_disregard, PDED_wid_disregard, PDED_unea_income_deduction_reason, PDED_unea_income_deduction_value, PDED_earned_income_deduction_reason, PDED_earned_income_deduction_value, PDED_ma_epd_inc_asset_limit, PDED_guard_fee, PDED_rep_payee_fee, PDED_other_expense, PDED_shel_spcl_needs, PDED_excess_need, PDED_restaurant_meals: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","PDED")
	EMReadScreen panel_check, 4, 2, 50
	If panel_check <> "PDED" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "PDED")
	End If
	call create_panel_if_nonexistent

	'Disa Widow/ers Deductionpded_shel_spcl_needs
	If pded_wid_deduction <> "" then
		pded_wid_deduction = ucase(pded_wid_deduction)
		pded_wid_deduction = left(pded_wid_deduction,1)
		EMWriteScreen pded_wid_deduction, 7, 60
	End If

	'Disa Adult Child Disregard
	If pded_adult_child_disregard <> "" then
		pded_adult_child_disregard = ucase(pded_adult_child_disregard)
		pded_adult_child_disregard = left(pded_adult_child_disregard,1)
		EMWriteScreen pded_adult_child_disregard, 8, 60
	End If

	'Widow/ers Disregard
	If pded_wid_disregard <> "" then
		pded_wid_disregard = ucase(pded_wid_disregard)
		pded_wid_disregard = left(pded_wid_disregard,1)
		EMWriteScreen pded_wid_disregard, 9, 60
	End If

	'Other Unearned Income Deduction
	If pded_unea_income_deduction_reason <> "" and pded_unea_income_deduction_value <> "" then
		EMWriteScreen pded_unea_income_deduction_value, 10, 62
		EMWriteScreen "X", 10, 25
		Transmit
		EMWriteScreen pded_unea_income_deduction_reason, 10, 51
		Transmit
		PF3
	End If

	'Other Earned Income Deduction
	If pded_earned_income_deduction_reason <> "" and pded_earned_income_deduction_value <> "" then
		EMWriteScreen pded_earned_income_deduction_value, 11, 62
		EMWriteScreen "X", 11, 27
		Transmit
		EMWriteScreen pded_earned_income_deduction_reason, 10, 51
		Transmit
		PF3
	End If

	'Extend MA-EPD Income/Asset Limits
	If pded_ma_epd_inc_asset_limit <> "" then
		pded_ma_epd_inc_asset_limit = ucase(pded_ma_epd_inc_asset_limit)
		pded_ma_epd_inc_asset_limit = left(pded_ma_epd_inc_asset_limit,1)
		EMWriteScreen pded_ma_epd_inc_asset_limit, 12, 65
	End If

	'Guardianship Fee
	If pded_guard_fee <> "" then
		EMWriteScreen pded_guard_fee, 15, 44
	End If

	'Rep Payee Fee
	If pded_rep_payee_fee <> "" then
		EMWriteScreen pded_guard_fee, 15, 70
	End If

	'Other Expense
	If pded_other_expense <> "" then
		EMWriteScreen pded_other_expense, 18, 41
	End If

	'Shelter Special Needs
	If pded_shel_spcl_needs <> "" then
		pded_shel_spcl_needs = ucase(pded_shel_spcl_needs)
		pded_shel_spcl_needs = left(pded_shel_spcl_needs,1)
		EMWriteScreen pded_shel_spcl_needs, 18, 78
	End If

	'Excess Need
	If pded_excess_need <> "" then
		EMWriteScreen pded_excess_need, 19, 41
	End If

	'Restaurant Meals
	If pded_restaurant_meals <> "" then
		pded_restaurant_meals = ucase(pded_restaurant_meals)
		pded_restaurant_meals = left(pded_restaurant_meals,1)
		EMWriteScreen pded_restaurant_meals, 19, 78
	End If
	Transmit
end function

function write_panel_to_MAXIS_PREG(PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver, PREG_due_date, PREG_multiple_birth)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver, PREG_due_date, PREG_multiple_birth: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "PREG")
	EMReadScreen panel_check, 4, 2, 45
	If panel_check <> "PREG" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "PREG")
	End If
	call create_panel_if_nonexistent
	EMWritescreen "NN", 20, 79
	transmit
	call create_MAXIS_friendly_date(PREG_conception_date, 0, 6, 53)
	third_trimester_date = dateadd("M", 6, PREG_conception_date)
	CALL create_MAXIS_friendly_date(third_trimester_date, 0, 8, 53)
	call create_MAXIS_friendly_date(PREG_due_date, 1, 10, 53)
	EMWritescreen PREG_conception_date_ver, 6, 75
	EMWritescreen PREG_third_trimester_ver, 8, 75
	EMWritescreen PREG_multiple_birth, 14, 53
	transmit
end function

function write_panel_to_MAXIS_RBIC(rbic_type, rbic_start_date, rbic_end_date, rbic_group_1, rbic_retro_income_group_1, rbic_prosp_income_group_1, rbic_ver_income_group_1, rbic_group_2, rbic_retro_income_group_2, rbic_prosp_income_group_2, rbic_ver_income_group_2, rbic_group_3, rbic_retro_income_group_3, rbic_prosp_income_group_3, rbic_ver_income_group_3, rbic_retro_hours, rbic_prosp_hours, rbic_exp_type_1, rbic_exp_retro_1, rbic_exp_prosp_1, rbic_exp_ver_1, rbic_exp_type_2, rbic_exp_retro_2, rbic_exp_prosp_2, rbic_exp_ver_2)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the rbic panel in MAXIS)
'~~~~~ rbic_type, rbic_start_date, rbic_end_date, rbic_group_1, rbic_retro_income_group_1, rbic_prosp_income_group_1, rbic_ver_income_group_1, rbic_group_2, rbic_retro_income_group_2, rbic_prosp_income_group_2, rbic_ver_income_group_2, rbic_group_3, rbic_retro_income_group_3, rbic_prosp_income_group_3, rbic_ver_income_group_3, rbic_retro_hours, rbic_prosp_hours, rbic_exp_type_1, rbic_exp_retro_1, rbic_exp_prosp_1, rbic_exp_ver_1, rbic_exp_type_2, rbic_exp_retro_2, rbic_exp_prosp_2, rbic_exp_ver_2: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "RBIC")  'navigates to the stat panel
	EMReadScreen panel_check, 4, 2, 55
	If panel_check <> "RBIC" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "RBIC")
	End If
	call create_panel_if_nonexistent
	EMwritescreen rbic_type, 5, 44  'enters rbic type code
	call create_MAXIS_friendly_date(rbic_start_date, 0, 6, 44)  'creates and enters a MAXIS friend date in the format mm/dd/yy for rbic start date
	IF rbic_end_date <> "" THEN call create_MAXIS_friendly_date(rbic_end_date, 6, 68)  'creates and enters a MAXIS friend date in the format mm/dd/yy for rbic end date
	rbic_group_1 = replace(rbic_group_1, " ", "")  'this will replace any spaces in the array with nothing removing the spaces.
	rbic_group_1 = split(rbic_group_1, ",")  'this will split up the reference numbers in the array based on commas
	rbic_col = 25                            'this will set the starting column to enter rbic reference numbers
	For each rbic_hh_memb in rbic_group_1    'for each reference number that is in the array for group 1 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 10, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_1, 10, 47  'enters the rbic retro income for group 1
	EMwritescreen rbic_prosp_income_group_1, 10, 62  'enters the rbic prospective income for group 1
	EMwritescreen rbic_ver_income_group_1, 10, 76    'enters the income verification code for group 1
	rbic_group_2 = replace(rbic_group_2, " ", "")
	rbic_group_2 = split(rbic_group_2, ",")
	rbic_col = 25
	For each rbic_hh_memb in rbic_group_2    'for each reference number that is in the array for group 2 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 11, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_2, 11, 47  'enters the rbic retro income for group 2
	EMwritescreen rbic_prosp_income_group_2, 11, 62  'enters the rbic prospective income for group 2
	EMwritescreen rbic_ver_income_group_2, 11, 76    'enters the income verification code for group 2
	rbic_group_3 = replace(rbic_group_3, " ", "")
	rbic_group_3 = split(rbic_group_3, ",")
	rbic_col = 25
	For each rbic_hh_memb in rbic_group_3    'for each reference number that is in the array for group 3 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 10, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_3, 12, 47  'enters the rbic retro income for group 3
	EMwritescreen rbic_prosp_income_group_3, 12, 62  'enters the rbic prospective income for group 3
	EMwritescreen rbic_ver_income_group_3, 12, 76    'enters the income verification code for group 3
	EMwritescreen rbic_retro_hours, 13, 52  'enters the retro hours
	EMwritescreen rbic_prosp_hours, 13, 67  'enters the prospective hours
	EMwritescreen rbic_exp_type_1, 15, 25   'enters the expenses type for group 1
	EMwritescreen rbic_exp_retro_1, 15, 47  'enters the expenses retro for group 1
	EMwritescreen rbic_exp_prosp_1, 15, 62  'enters the expenses prospective for group 1
	EMwritescreen rbic_exp_ver_1, 15, 76    'enters the expenses verification code for group 1
	EMwritescreen rbic_exp_type_2, 16, 25   'enters the expenses type for group 2
	EMwritescreen rbic_exp_retro_2, 16, 47  'enters the expenses retro for group 2
	EMwritescreen rbic_exp_prosp_2, 16, 62  'enters the expenses prospective for group 2
	EMwritescreen rbic_exp_ver_2, 16, 76    'enters the expenses verification code for group 2
end function

function write_panel_to_MAXIS_REST(rest_type, rest_type_ver, rest_market, rest_market_ver, rest_owed, rest_owed_ver, rest_date, rest_status, rest_joint, rest_share_ratio, rest_agreement_date)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the rest panel in MAXIS)
'~~~~~ rest_type, rest_type_ver, rest_market, rest_market_ver, rest_owed, rest_owed_ver, rest_date, rest_status, rest_joint, rest_share_ratio, rest_agreement_date: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "REST")  'navigates to the stat panel
	EMReadScreen panel_check, 4, 2, 46
	If panel_check <> "REST" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "REST")
	End If
	call create_panel_if_nonexistent
	Emwritescreen rest_type, 6, 39  'enters residence type
	Emwritescreen rest_type_ver, 6, 62  'enters verification of residence type
	Emwritescreen rest_market, 8, 41  'enters market value of residence
	Emwritescreen rest_market_ver, 8, 62  'enters market value verification code
	Emwritescreen rest_owed, 9, 41  'enters amount owned on residence
	Emwritescreen rest_owed_ver, 9, 62  'enters amount owed verification code
	call create_MAXIS_friendly_date(rest_date, 0, 10, 39)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen rest_status, 12, 54  'enters property status code
	Emwritescreen rest_joint, 13, 54  'enters if it is a jointly owned home
	Emwritescreen left(rest_share_ratio, 1), 14, 54  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(rest_share_ratio, 1), 14, 58  'enters the ratio of ownership using the right 1 digit of what is entered into the file
	IF rest_agreement_date <> "" THEN call create_MAXIS_friendly_date(rest_agreement_date, 0, 16, 62)
end function

function write_panel_to_MAXIS_SCHL(appl_date, SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ appl_date, SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "SCHL")
	EMReadScreen panel_check, 4, 2, 43
	If panel_check <> "SCHL" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "SCHL")
	End If
	EMWriteScreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_SCHL, 1, 2, 78
	IF num_of_SCHL = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit

		call create_MAXIS_friendly_date(appl_date, 0, 5, 40)						'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
		EMWriteScreen datepart("yyyy", appl_date), 5, 46
		EMWriteScreen SCHL_status, 6, 40
		EMWriteScreen SCHL_ver, 6, 63
		EMWriteScreen SCHL_type, 7, 40
		IF len(SCHL_district_nbr) <> 4 THEN
			DO
				SCHL_district_nbr = "0" & SCHL_district_nbr
			LOOP UNTIL len(SCHL_district_nbr) = 4
		END IF
		EMWriteScreen SCHL_district_nbr, 8, 40
		If SCHL_kindergarten_start_date <> "" then call create_MAXIS_friendly_date(SCHL_kindergarten_start_date, 0, 10, 63)
		EMWriteScreen left(SCHL_grad_date, 2), 11, 63
		EMWriteScreen right(SCHL_grad_date, 2), 11, 66
		EMWriteScreen SCHL_grad_date_ver, 12, 63
		EMWriteScreen SCHL_primary_secondary_funding, 14, 63
		EMWriteScreen SCHL_FS_eligibility_status, 16, 63
		EMWriteScreen SCHL_higher_ed, 18, 63
		transmit
	END IF
end function

function write_panel_to_MAXIS_SECU(secu_type, secu_pol_numb, secu_name, secu_cash_val, secu_date, secu_cash_ver, secu_face_val, secu_withdraw, secu_cash_count, secu_SNAP_count, secu_HC_count, secu_GRH_count, secu_IV_count, secu_joint, secu_share_ratio)
'--- This function writes to MAXIS in Krabappel only (writes using the variables read off of the specialized excel template to the secu panel in MAXIS)
'~~~~~ secu_type, secu_pol_numb, secu_name, secu_cash_val, secu_date, secu_cash_ver, secu_face_val, secu_withdraw, secu_cash_count, secu_SNAP_count, secu_HC_count, secu_GRH_count, secu_IV_count, secu_joint, secu_share_ratio: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	Call navigate_to_MAXIS_screen("STAT", "SECU")  'navigates to the stat panel
	EMReadScreen panel_check, 4, 2, 45
	If panel_check <> "SECU" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "SECU")
	End If
	call create_panel_if_nonexistent
	Emwritescreen secu_type, 6, 50  'enters security type
	Emwritescreen secu_pol_numb, 7, 50  'enters policy number
	Emwritescreen secu_name, 8, 50  'enters name of policy
	Emwritescreen secu_cash_val, 10, 52  'enters cash value of policy
	call create_MAXIS_friendly_date(secu_date, 0, 11, 35)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen secu_cash_ver, 11, 50  'enters cash value verification code
	Emwritescreen secu_face_val, 12, 52  'enters face value of policy
	Emwritescreen secu_withdraw, 13, 52  'enters withdrawl penalty
	Emwritescreen secu_cash_count, 15, 50  'enters y/n if counted for cash
	Emwritescreen secu_SNAP_count, 15, 57  'enters y/n if counted for snap
	Emwritescreen secu_HC_count, 15, 64  'enters y/n if counted for hc
	Emwritescreen secu_GRH_count, 15, 72  'enters y/n if counted for grh
	Emwritescreen secu_IV_count, 15, 80  'enters y/n if counted for iv
	Emwritescreen secu_joint, 16, 44  'enters if it is a jointly owned security
	Emwritescreen left(secu_share_ratio, 1), 16, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(secu_share_ratio, 1), 16, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
end function

function write_panel_to_MAXIS_SHEL(SHEL_subsidized, SHEL_shared, SHEL_paid_to, SHEL_rent_retro, SHEL_rent_retro_ver, SHEL_rent_pro, SHEL_rent_pro_ver, SHEL_lot_rent_retro, SHEL_lot_rent_retro_ver, SHEL_lot_rent_pro, SHEL_lot_rent_pro_ver, SHEL_mortgage_retro, SHEL_mortgage_retro_ver, SHEL_mortgage_pro, SHEL_mortgage_pro_ver, SHEL_insur_retro, SHEL_insur_retro_ver, SHEL_insur_pro, SHEL_insur_pro_ver, SHEL_taxes_retro, SHEL_taxes_retro_ver, SHEL_taxes_pro, SHEL_taxes_pro_ver, SHEL_room_retro, SHEL_room_retro_ver, SHEL_room_pro, SHEL_room_pro_ver, SHEL_garage_retro, SHEL_garage_retro_ver, SHEL_garage_pro, SHEL_garage_pro_ver, SHEL_subsidy_retro, SHEL_subsidy_retro_ver, SHEL_subsidy_pro, SHEL_subsidy_pro_ver)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ SHEL_subsidized, SHEL_shared, SHEL_paid_to, SHEL_rent_retro, SHEL_rent_retro_ver, SHEL_rent_pro, SHEL_rent_pro_ver, SHEL_lot_rent_retro, SHEL_lot_rent_retro_ver, SHEL_lot_rent_pro, SHEL_lot_rent_pro_ver, SHEL_mortgage_retro, SHEL_mortgage_retro_ver, SHEL_mortgage_pro, SHEL_mortgage_pro_ver, SHEL_insur_retro, SHEL_insur_retro_ver, SHEL_insur_pro, SHEL_insur_pro_ver, SHEL_taxes_retro, SHEL_taxes_retro_ver, SHEL_taxes_pro, SHEL_taxes_pro_ver, SHEL_room_retro, SHEL_room_retro_ver, SHEL_room_pro, SHEL_room_pro_ver, SHEL_garage_retro, SHEL_garage_retro_ver, SHEL_garage_pro, SHEL_garage_pro_ver, SHEL_subsidy_retro, SHEL_subsidy_retro_ver, SHEL_subsidy_pro, SHEL_subsidy_pro_ver: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "SHEL")
	EMReadScreen panel_check, 4, 2, 48
	If panel_check <> "SHEL" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "SHEL")
	End If
	call create_panel_if_nonexistent
	EMWritescreen SHEL_subsidized, 6, 46
	EMWritescreen SHEL_shared, 6, 64
	EMWritescreen SHEL_paid_to, 7, 50
	EMWritescreen SHEL_rent_retro, 11, 37
	EMWritescreen SHEL_rent_retro_ver, 11, 48
	EMWritescreen SHEL_rent_pro, 11, 56
	EMWritescreen SHEL_rent_pro_ver, 11, 67
	EMWritescreen SHEL_lot_rent_retro, 12, 37
	EMWritescreen SHEL_lot_rent_retro_ver, 12, 48
	EMWritescreen SHEL_lot_rent_pro, 12, 56
	EMWritescreen SHEL_lot_rent_pro_ver, 12, 67
	EMWritescreen SHEL_mortgage_retro, 13, 37
	EMWritescreen SHEL_mortgage_retro_ver, 13, 48
	EMWritescreen SHEL_mortgage_pro, 13, 56
	EMwritescreen SHEL_mortgage_pro_ver, 13, 67
	EMWritescreen SHEL_insur_retro, 14, 37
	EMWritescreen SHEL_insur_retro_ver, 14, 48
	EMWritescreen SHEL_insur_pro, 14, 56
	EMWritescreen SHEL_insur_pro_ver, 14, 67
	EMWritescreen SHEL_taxes_retro, 15, 37
	EMWritescreen SHEL_taxes_retro_ver, 15, 48
	EMWritescreen SHEL_taxes_pro, 15, 56
	EMWritescreen SHEL_taxes_pro_ver, 15, 67
	EMWritescreen SHEL_room_retro, 16, 37
	EMWritescreen SHEL_room_retro_ver, 16, 48
	EMWritescreen SHEL_room_pro, 16, 56
	EMWritescreen SHEL_room_pro_ver, 16, 67
	EMWritescreen SHEL_garage_retro, 17, 37
	EMWritescreen SHEL_garage_retro_ver, 17, 48
	EMWritescreen SHEL_garage_pro, 17, 56
	EMWritescreen SHEL_garage_pro_ver, 17, 67
	EMWritescreen SHEL_subsidy_retro, 18, 37
	EMWritescreen SHEL_subsidy_retro_ver, 18, 48
	EMWritescreen SHEL_subsidy_pro, 18, 56
	EMWritescreen SHEL_subsidy_pro_ver, 18, 67
	transmit
end function

function write_panel_to_MAXIS_SIBL(SIBL_group_1, SIBL_group_2, SIBL_group_3)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ SIBL_group_1, SIBL_group_2, SIBL_group_3: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "SIBL")
	EMReadScreen panel_check, 4, 2, 47
	If panel_check <> "SIBL" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "SIBL")
	End If
	EMReadScreen num_of_SIBL, 1, 2, 78
	IF num_of_SIBL = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	END IF

	If SIBL_group_1 <> "" then
		EMWritescreen "01", 7, 28
		SIBL_group_1 = replace(SIBL_group_1, " ", "") 'Removing spaces
		SIBL_group_1 = split(SIBL_group_1, ",") 'Splits the sibling group value into an array by commas
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_1 'Writes the member numbers onto the group line
			EMWritescreen SIBL_group_member, 7, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if

	If SIBL_group_2 <> "" then
		EMWritescreen "02", 8, 28
		SIBL_group_2 = replace(SIBL_group_2, " ", "")
		SIBL_group_2 = split(SIBL_group_2, ",")
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_2
			EMWritescreen SIBL_group_member, 8, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if

	If SIBL_group_3 <> "" then
		EMWritescreen "03", 9, 28
		SIBL_group_2 = replace(SIBL_group_3, " ", "")
		SIBL_group_2 = split(SIBL_group_3, ",")
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_3
			EMWritescreen SIBL_group_member, 9, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if
	transmit
end function

function write_panel_to_MAXIS_SPON(SPON_type, SPON_ver, SPON_name, SPON_state)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ SPON_type, SPON_ver, SPON_name, SPON_state: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "SPON")
	EMReadScreen panel_check, 4, 2, 53
	If panel_check <> "SPON" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "SPON")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits
	call create_panel_if_nonexistent
	EMWriteScreen SPON_type, 6, 38
	EMWriteScreen SPON_ver, 6, 62
	EMWriteScreen SPON_name, 8, 38
	EMWriteScreen SPON_state, 10, 62
	transmit
end function

function write_panel_to_MAXIS_STEC(STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "STEC")
	EMReadScreen panel_check, 4, 2, 54
	If panel_check <> "STEC" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "STEC")
	End If
	EMWriteSCreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_STEC, 1, 2, 78
	IF num_of_STEC = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit

		EMWriteScreen STEC_type_1, 8, 25				'STEC 1
		EMWriteScreen STEC_amt_1, 8, 31
		STEC_actual_from_thru_months_1 = replace(STEC_actual_from_thru_months_1, " ", "")
		EMWriteScreen left(STEC_actual_from_thru_months_1, 2), 8, 41
		EMWriteScreen mid(STEC_actual_from_thru_months_1, 4, 2), 8, 44
		EMWriteScreen mid(STEC_actual_from_thru_months_1, 7, 2), 8, 48
		EMWriteScreen right(STEC_actual_from_thru_months_1, 2), 8, 51
		EMWriteScreen STEC_ver_1, 8, 55
		EMWriteScreen STEC_earmarked_amt_1, 8, 59
		STEC_earmarked_from_thru_months_1 = replace(STEC_earmarked_from_thru_months_1, " ", "")
		EMWriteScreen left(STEC_earmarked_from_thru_months_1, 2), 8, 69
		EMWriteScreen mid(STEC_earmarked_from_thru_months_1, 4, 2), 8, 72
		EMWriteScreen mid(STEC_earmarked_from_thru_months_1, 7, 2), 8, 76
		EMWriteScreen right(STEC_earmarked_from_thru_months_1, 2), 8, 79
		EMWriteScreen STEC_type_2, 9, 25				'STEC 1
		EMWriteScreen STEC_amt_2, 9, 31
		STEC_actual_from_thru_months_2 = replace(STEC_actual_from_thru_months_2, " ", "")
		EMWriteScreen left(STEC_actual_from_thru_months_2, 2), 9, 41
		EMWriteScreen mid(STEC_actual_from_thru_months_2, 4, 2), 9, 44
		EMWriteScreen mid(STEC_actual_from_thru_months_2, 7, 2), 9, 48
		EMWriteScreen right(STEC_actual_from_thru_months_2, 2), 9, 51
		EMWriteScreen STEC_ver_2, 9, 55
		EMWriteScreen STEC_earmarked_amt_2, 9, 59
		STEC_earmarked_from_thru_months_2 = replace(STEC_earmarked_from_thru_months_2, " ", "")
		EMWriteScreen left(STEC_earmarked_from_thru_months_2, 2), 9, 69
		EMWriteScreen mid(STEC_earmarked_from_thru_months_2, 4, 2), 9, 72
		EMWriteScreen mid(STEC_earmarked_from_thru_months_2, 7, 2), 9, 76
		EMWriteScreen right(STEC_earmarked_from_thru_months_2, 2), 9, 79
		transmit
	END IF
end function

function write_panel_to_MAXIS_STIN(STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "STIN")
	EMReadScreen panel_check, 4, 2, 47
	If panel_check <> "STIN" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "STIN")
	End If
	EMWriteSCreen reference_number, 20, 76
	transmit

	EMReadScreen num_of_STIN, 1, 2, 78
	IF num_of_STIN = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit

		EMWriteScreen STIN_type_1, 8, 27				'STIN 1
		EMWriteScreen STIN_amt_1, 8, 34
		call create_MAXIS_friendly_date(STIN_avail_date_1, 0, 8, 46)
		STIN_months_covered_1 = replace(STIN_months_covered_1, " ", "")
		EMWriteScreen left(STIN_months_covered_1, 2), 8, 58
		EMWriteScreen mid(STIN_months_covered_1, 4, 2), 8, 61
		EMWriteScreen mid(STIN_months_covered_1, 7, 2), 8, 67
		EMWriteScreen right(STIN_months_covered_1, 2), 8, 70
		EMWriteScreen STIN_ver_1, 8, 76
		EMWriteScreen STIN_type_2, 9, 27				'STIN 2
		EMWriteScreen STIN_amt_2, 9, 34
		STIN_avail_date_2 = replace(STIN_avail_date_2, " ", "")
		IF STIN_avail_date_2 <> "" THEN call create_MAXIS_friendly_date(STIN_avail_date_2, 0, 9, 46)
		EMWriteScreen left(STIN_months_covered_2, 2), 9, 58
		EMWriteScreen mid(STIN_months_covered_2, 4, 2), 9, 61
		EMWriteScreen mid(STIN_months_covered_2, 7, 2), 9, 67
		EMWriteScreen right(STIN_months_covered_2, 2), 9, 70
		EMWriteScreen STIN_ver_2, 9, 76
		transmit
	END IF
end function

function write_panel_to_MAXIS_STWK(STWK_empl_name, STWK_wrk_stop_date, STWK_wrk_stop_date_verif, STWK_inc_stop_date, STWK_refused_empl_yn, STWK_vol_quit, STWK_ref_empl_date, STWK_gc_cash, STWK_gc_grh, STWK_gc_fs, STWK_fs_pwe, STWK_maepd_ext)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ STWK_empl_name, STWK_wrk_stop_date, STWK_wrk_stop_date_verif, STWK_inc_stop_date, STWK_refused_empl_yn, STWK_vol_quit, STWK_ref_empl_date, STWK_gc_cash, STWK_gc_grh, STWK_gc_fs, STWK_fs_pwe, STWK_maepd_ext: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT","STWK")
	EMReadScreen panel_check, 4, 2, 45
	If panel_check <> "STWK" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "STWK")
	End If
	call create_panel_if_nonexistent

	EMWriteScreen stwk_empl_name, 6, 46
	If stwk_wrk_stop_date <> "" then CALL create_MAXIS_friendly_date(stwk_wrk_stop_date, 0, 7, 46)
	EMWriteScreen stwk_wrk_stop_date_verif, 7, 63
	IF stwk_inc_stop_date <> "" THEN CALL create_MAXIS_friendly_date(stwk_inc_stop_date, 0, 8, 46)
	EMWriteScreen stwk_refused_empl_yn, 8, 78
	EMWriteScreen stwk_vol_quit, 10, 46
	If stwk_ref_empl_date <> "" then CALL create_MAXIS_friendly_date(stwk_ref_empl_date, 0, 10, 72)
	EMWriteScreen stwk_gc_cash, 12, 52
	EMWriteScreen stwk_gc_grh, 12, 60
	EMWriteScreen stwk_gc_fs, 12, 67
	EMWriteScreen stwk_fs_pwe, 14, 46
	EMWriteScreen stwk_maepd_ext, 16, 46
	Transmit
end function

function write_panel_to_MAXIS_TYPE_PROG_REVW(appl_date, type_cash_yn, type_hc_yn, type_fs_yn, prog_mig_worker, revw_ar_or_ir, revw_exempt)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ appl_date, type_cash_yn, type_hc_yn, type_fs_yn, prog_mig_worker, revw_ar_or_ir, revw_exempt: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "TYPE")
	EMReadScreen panel_check, 4, 2, 48
	If panel_check <> "TYPE" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "TYPE")
	End If
	IF reference_number = "01" THEN
		EMWriteScreen "NN", 20, 79
		transmit
		EMWriteScreen type_cash_yn, 6, 28
		EMWriteScreen type_hc_yn, 6, 37
		EMWriteScreen type_fs_yn, 6, 46
		EMWriteScreen "N", 6, 55
		EMWriteScreen "N", 6, 64
		EMWriteScreen "N", 6, 73
		type_row = 7
		DO				'<=====this DO/LOOP populates "N" for all other HH members on TYPE so the script can get past TYPE when the reference number = "01"
			EMReadScreen type_does_hh_memb_exist, 2, type_row, 3
			IF type_does_hh_memb_exist <> "  " THEN
				EMWriteScreen "N", type_row, 28
				EMWriteScreen "N", type_row, 37
				EMWriteScreen "N", type_row, 46
				EMWriteScreen "N", type_row, 55
				type_row = type_row + 1
			ELSE
				EXIT DO
			END IF
		LOOP WHILE type_does_hh_memb_exist <> "  "
	ELSE
		PF9
		type_row = 7
		DO
			EMReadScreen type_does_hh_memb_exist, 2, type_row, 3
			IF type_does_hh_memb_exist = reference_number THEN
				EMWriteScreen type_cash_yn, type_row, 28
				EMWriteScreen type_hc_yn, type_row, 37
				EMWriteScreen type_fs_yn, type_row, 46
				EMWriteScreen "N", type_row, 55
				exit do
			ELSE
				type_row = type_row + 1
			END IF
		LOOP UNTIL type_does_hh_memb_exist = reference_number
	END IF
	transmit		'<===== when reference_number = "01" this transmit will navigate to PROG, else, it will navigate to STAT/WRAP

	IF reference_number = "01" THEN		'<===== only accesses PROG & REVW if reference_number = "01"
		call navigate_to_MAXIS_screen("STAT", "PROG")
		EMWriteScreen "NN", 20, 71
		transmit
			IF type_cash_yn = "Y" THEN
				call create_MAXIS_friendly_date(appl_date, 0, 6, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 6, 44)
				call create_MAXIS_friendly_date(appl_date, 0, 6, 55)
			END IF
			IF type_fs_yn = "Y" THEN
				call create_MAXIS_friendly_date(appl_date, 0, 10, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 10, 44)
				call create_MAXIS_friendly_date(appl_date, 0, 10, 55)
			END IF
			IF type_hc_yn = "Y" THEN
				call create_MAXIS_friendly_date(appl_date, 0, 12, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 12, 55)
			END IF
			EMWriteScreen mig_worker, 18, 67
			transmit
			EMWriteScreen mig_worker, 18, 67
			transmit

		call navigate_to_MAXIS_screen("STAT", "REVW")
		EMWriteScreen "NN", 20, 71
		transmit
			IF type_cash_yn = "Y" THEN
				cash_review_date = dateadd("YYYY", 1, appl_date)
				call create_MAXIS_friendly_date(cash_review_date, 0, 9, 37)
			END IF
			IF type_fs_yn = "Y" THEN
				EMWriteScreen "X", 5, 58
				transmit
				DO
					EMReadScreen food_support_reports, 20, 5, 30
				LOOP UNTIL food_support_reports = "FOOD SUPPORT REPORTS"
				fs_csr_date = dateadd("M", 6, appl_date)
				fs_er_date = dateadd("M", 12, appl_date)
				call create_MAXIS_friendly_date(fs_csr_date, 0, 9, 26)
				call create_MAXIS_friendly_date(fs_er_date, 0, 9, 64)
				transmit
			END IF
			IF type_hc_yn = "Y" THEN
				EMWriteScreen "X", 5, 71
				transmit
				DO
					EMReadScreen health_care_renewals, 20, 4, 32
				LOOP UNTIL health_care_renewals = "HEALTH CARE RENEWALS"
				IF revw_ar_or_ir = "AR" THEN
					call create_MAXIS_friendly_date((dateadd("M", 6, appl_date)), 0, 8, 71)
				ELSEIF revw_ar_or_ir = "IR" THEN
					call create_MAXIS_friendly_date((dateadd("M", 6, appl_date)), 0, 8, 27)
				END IF
				call create_MAXIS_friendly_date((dateadd("M", 12, appl_date)), 0, 9, 27)
				EMWriteScreen revw_exempt, 9, 71
				transmit
			END IF
	END IF
end function

function write_panel_to_MAXIS_UNEA(unea_number, unea_inc_type, unea_inc_verif, unea_claim_suffix, unea_start_date, unea_pay_freq, unea_inc_amount, ssn_first, ssn_mid, ssn_last)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ unea_number, unea_inc_type, unea_inc_verif, unea_claim_suffix, unea_start_date, unea_pay_freq, unea_inc_amount, ssn_first, ssn_mid, ssn_last: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "UNEA")
	EMReadScreen unea_check, 4, 2, 48
	If unea_check <> "UNEA" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "UNEA")
	End If
	PF10
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen unea_number, 20, 79
	transmit

	EMReadScreen does_not_exist, 14, 24, 13
	IF does_not_exist = "DOES NOT EXIST" THEN
		EMWriteScreen "NN", 20, 79
		transmit

		'Putting this part in with the NN because otherwise the script will update it in later months and change claim number information.
		EMWriteScreen unea_inc_type, 5, 37
		EMWriteScreen unea_inc_verif, 5, 65
		EMWriteScreen (ssn_first & ssn_mid & ssn_last & unea_claim_suffix), 6, 37
		call create_MAXIS_friendly_date(unea_start_date, 0, 7, 37)
	ELSE
		PF9
	END IF

	'=====Navigates to the PIC for UNEA=====
	EMWriteScreen "X", 10, 26
	transmit
	EMReadScreen pic_info_exists, 6, 18, 58		'---Deteremining if PIC info exists. If it does, the script will just back out.
	pic_info_exists = trim(pic_info_exists)
	IF pic_info_exists = "" THEN
		EMWriteScreen unea_pay_freq, 5, 64
		EMWriteScreen unea_inc_amount, 8, 66
		calc_month = datepart("M", date)
		IF len(calc_month) = 1 THEN calc_month = "0" & calc_month
		calc_day = datepart("D", date)
		IF len(calc_day) = 1 THEN calc_day = "0" & calc_day
        calc_year = right( DatePart("yyyy",date), 2)
		EMWriteScreen calc_month, 5, 34
		EMWriteScreen calc_day, 5, 37
		EMWriteScreen calc_year, 5, 40
        Do              '<=====navigates out of the PIC
            transmit
            EmReadscreen PIC_Check, 16, 3, 28
            IF PIC_check <> "SNAP Prospective" then exit do
        Loop
	ELSE
		PF3
	END IF

	'=====the following bit is for the retrospective & prospective pay dates=====
	EMReadScreen bene_month, 2, 20, 55
	EMReadScreen bene_year, 2, 20, 58
	current_bene_month = bene_month & "/01/" & bene_year
	retro_month = datepart("M", DateAdd("M", -2, current_bene_month))
	IF len(retro_month) <> 2 THEN retro_month = "0" & retro_month
	retro_year = right(datepart("YYYY", DateAdd("M", -2, current_bene_month)), 2)

	EMWriteScreen retro_month, 13, 25
	EMWriteScreen retro_year, 13, 31
	EMWriteScreen bene_month, 13, 54
	EMWriteScreen bene_year, 13, 60

	IF pic_info_exists = "" THEN 	'---Meaning, the case has PIC info...which is to say that this is a PF9 and not a NN
		EMWriteScreen "05", 13, 28
		EMWriteScreen unea_inc_amount, 13, 39
		EMWriteScreen "05", 13, 57
		EMWriteScreen unea_inc_amount, 13, 68
	END IF

	IF unea_pay_freq = "2" OR unea_pay_freq = "3" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen bene_year, 14, 60

		IF pic_info_exists = "" THEN
			EMWriteScreen "19", 14, 28
			EMWriteScreen "19", 14, 57
			EMWriteScreen unea_inc_amount, 14, 39
			EMWriteScreen unea_inc_amount, 14, 68
		END IF
	ELSEIF unea_pay_freq = "4" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen retro_month, 16, 25
		EMWriteScreen retro_year, 16, 31
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen bene_month, 15, 54
		EMWriteScreen bene_year, 15, 60
		EMWriteScreen bene_month, 16, 54
		EMWriteScreen bene_year, 16, 60

		IF pic_info_exists = "" THEN
			EMWriteScreen "12", 14, 28
			EMWriteScreen unea_inc_amount, 14, 39
			EMWriteScreen "19", 15, 28
			EMWriteScreen unea_inc_amount, 15, 39
			EMWriteScreen "26", 16, 28
			EMWriteScreen unea_inc_amount, 16, 39
			EMWriteScreen "12", 14, 57
			EMWriteScreen unea_inc_amount, 14, 68
			EMWriteScreen "19", 15, 57
			EMWriteScreen unea_inc_amount, 15, 68
			EMWriteScreen "26", 16, 57
			EMWriteScreen unea_inc_amount, 16, 68
		END IF
	END IF

	'=====determines if the benefit month is current month + 1 and dumps information into the HC income estimator
	IF (bene_month * 1) = (datepart("M", date) + 1) THEN		'<===== "bene_month * 1" is needed to convert bene_month from a string to a useable number
		EMWriteScreen "X", 6, 56
		transmit
		EMWriteScreen "________", 9, 65
		EMWriteScreen unea_inc_amount, 9, 65
		EMWriteScreen unea_pay_freq, 10, 63
		transmit
		transmit
	END IF
	Transmit
  	EMReadScreen warning_check, 7, 24, 2 'This checks for an error with COLA field being blank
  	IF warning_check = "WARNING" THEN transmit
end function

function write_panel_to_MAXIS_WKEX(program, fed_tax_retro, fed_tax_prosp, fed_tax_verif, state_tax_retro, state_tax_prosp, state_tax_verif, fica_retro, fica_prosp, fica_verif, tran_retro, tran_prosp, tran_verif, tran_imp_rel, meals_retro, meals_prosp, meals_verif, meals_imp_rel, uniforms_retro, uniforms_prosp, uniforms_verif, uniforms_imp_rel, tools_retro, tools_prosp, tools_verif, tools_imp_rel, dues_retro, dues_prosp, dues_verif, dues_imp_rel, othr_retro, othr_prosp, othr_verif, othr_imp_rel, HC_Exp_Fed_Tax, HC_Exp_State_Tax, HC_Exp_FICA, HC_Exp_Tran, HC_Exp_Tran_imp_rel, HC_Exp_Meals, HC_Exp_Meals_Imp_Rel, HC_Exp_Uniforms, HC_Exp_Uniforms_Imp_Rel, HC_Exp_Tools, HC_Exp_Tools_Imp_Rel, HC_Exp_Dues, HC_Exp_Dues_Imp_Rel, HC_Exp_Othr, HC_Exp_Othr_Imp_Rel)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ program, fed_tax_retro, fed_tax_prosp, fed_tax_verif, state_tax_retro, state_tax_prosp, state_tax_verif, fica_retro, fica_prosp, fica_verif, tran_retro, tran_prosp, tran_verif, tran_imp_rel, meals_retro, meals_prosp, meals_verif, meals_imp_rel, uniforms_retro, uniforms_prosp, uniforms_verif, uniforms_imp_rel, tools_retro, tools_prosp, tools_verif, tools_imp_rel, dues_retro, dues_prosp, dues_verif, dues_imp_rel, othr_retro, othr_prosp, othr_verif, othr_imp_rel, HC_Exp_Fed_Tax, HC_Exp_State_Tax, HC_Exp_FICA, HC_Exp_Tran, HC_Exp_Tran_imp_rel, HC_Exp_Meals, HC_Exp_Meals_Imp_Rel, HC_Exp_Uniforms, HC_Exp_Uniforms_Imp_Rel, HC_Exp_Tools, HC_Exp_Tools_Imp_Rel, HC_Exp_Dues, HC_Exp_Dues_Imp_Rel, HC_Exp_Othr, HC_Exp_Othr_Imp_Rel: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	CALL navigate_to_MAXIS_screen("STAT", "WKEX")
	EMReadScreen panel_check, 4, 2, 47
	If panel_check <> "WKEX" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "WKEX")
	End If
	EMReadScreen ERRR_check, 4, 2, 52			'Checking for the ERRR screen
	If ERRR_check = "ERRR" then transmit		'If the ERRR screen is found, it transmits0

	EMWriteScreen reference_number, 20, 76
	transmit

	'Determining the number of WKEX panels so the script knows how to handle the incoming information.
	EMReadScreen num_of_WKEX_panels, 1, 2, 78
	IF num_of_WKEX_panels = "5" THEN		'If there are already 5 WKEX panels, the script will not create a new panel.
		EXIT function
	ELSEIF num_of_WKEX_panels = "0" THEN
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit

		'---When the script needs to generate a new WKEX, it will enter the information for that panel...
		EMWriteScreen program, 5, 33
		EMWriteScreen fed_tax_retro, 7, 43
		EMWriteScreen fed_tax_prosp, 7, 57
		EMWriteScreen fed_tax_verif, 7, 69
		EMWriteScreen state_tax_retro, 8, 43
		EMWriteScreen state_tax_prosp, 8, 57
		EMWriteScreen state_tax_verif, 8, 69
		EMWriteScreen fica_retro, 9, 43
		EMWriteScreen fica_prosp, 9, 57
		EMWriteScreen fica_verif, 9, 69
		EMWriteScreen tran_retro, 10, 43
		EMWriteScreen tran_prosp, 10, 57
		EMWriteScreen tran_verif, 10, 69
		EMWriteScreen tran_imp_rel, 10, 75
		EMWriteScreen meals_retro, 11, 43
		EMWriteScreen meals_prosp, 11, 57
		EMWriteScreen meals_verif, 11, 69
		EMWriteScreen meals_imp_rel, 11, 75
		EMWriteScreen uniforms_retro, 12, 43
		EMWriteScreen uniforms_prosp, 12, 57
		EMWriteScreen uniforms_verif, 12, 69
		EMWriteScreen uniforms_imp_rel, 12, 75
		EMWriteScreen tools_retro, 13, 43
		EMWriteScreen tools_prosp, 13, 57
		EMWriteScreen tools_verif, 13, 69
		EMWriteScreen tools_imp_rel, 13, 75
		EMWriteScreen dues_retro, 14, 43
		EMWriteScreen dues_prosp, 14, 57
		EMWriteScreen dues_verif, 14, 69
		EMWriteScreen dues_imp_rel, 14, 75
		EMWriteScreen othr_retro, 15, 43
		EMWriteScreen othr_prosp, 15, 57
		EMWriteScreen othr_verif, 15, 69
		EMWriteScreen othr_imp_rel, 15, 75
	ELSE
		PF9
		'---If the script is editing an existing WKEX page, it would be doing so ONLY to update the HC Expense sub-menu.
		'---Adding to the HC Expenses
		EMWriteScreen "X", 18, 57
		transmit

		EMReadScreen current_month, 17, 20, 51
		IF current_month = "CURRENT MONTH + 1" THEN
			PF3
		ELSE
			EMWriteScreen HC_Exp_Fed_Tax, 8, 36
			EMWriteScreen HC_Exp_State_Tax, 9, 36
			EMWriteScreen HC_Exp_FICA, 10, 36
			EMWriteScreen HC_Exp_Tran, 11, 36
			EMWriteScreen HC_Exp_Tran_imp_rel, 11, 51
			EMWriteScreen HC_Exp_Meals, 12, 36
			EMWriteScreen HC_Exp_Meals_Imp_Rel, 12, 51
			EMWriteScreen HC_Exp_Uniforms, 13, 36
			EMWriteScreen HC_Exp_Uniforms_Imp_Rel, 13, 51
			EMWriteScreen HC_Exp_Tools, 14, 36
			EMWriteScreen HC_Exp_Tools_Imp_Rel, 14, 51
			EMWriteScreen HC_Exp_Dues, 15, 36
			EMWriteScreen HC_Exp_Dues_Imp_Rel, 15, 51
			EMWriteScreen HC_Exp_Othr, 16, 36
			EMWriteScreen HC_Exp_Othr_Imp_Rel, 16, 51
			transmit
			PF3
		END IF
	END IF
	transmit
end function

function write_panel_to_MAXIS_WREG(wreg_fs_pwe, wreg_fset_status, wreg_defer_fs, wreg_fset_orientation_date, wreg_fset_sanction_date, wreg_num_sanctions, wreg_sanction_reason, wreg_abawd_status, wreg_ga_basis)
'--- This function writes to MAXIS in Krabappel only
'~~~~~ wreg_fs_pwe, wreg_fset_status, wreg_defer_fs, wreg_fset_orientation_date, wreg_fset_sanction_date, wreg_num_sanctions, wreg_abawd_status, wreg_ga_basis: parameters for the training case creator to work
'===== Keywords: MAXIS, Krabappel, traning, case, creator
	call navigate_to_MAXIS_screen("STAT", "WREG")
	EMReadScreen panel_check, 4, 2, 48
	If panel_check <> "WREG" Then
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "WREG")
	End If
	call create_panel_if_nonexistent

	EMWriteScreen wreg_fs_pwe, 6, 68
	EMWriteScreen wreg_fset_status, 8, 50
	EMWriteScreen wreg_defer_fs, 8, 80
	IF wreg_fset_orientation_date <> "" THEN call create_MAXIS_friendly_date(wreg_fset_orientation_date, 0, 9, 50)
    IF wreg_fset_sanction_date <> "" then
        sanc_mo = right("0" & DatePart("m",    wreg_fset_sanction_date), 2)
        sanc_yr = right(      DatePart("yyyy", wreg_fset_sanction_date), 2)
        EmWriteScreen sanc_mo, 10, 50
        EmWriteScreen sanc_yr, 10, 56
    End if

	IF wreg_num_sanctions <> "" THEN EMWriteScreen wreg_num_sanctions, 11, 50
    If wreg_sanction_reason <> "" THEN EmWriteScreen wreg_sanction_reason, 12, 50
	EMWriteScreen wreg_abawd_status, 13, 50
	EMWriteScreen wreg_ga_basis, 15, 50
	transmit
end function

Function transfer_cases(workers_to_XFER_cases_to, case_number_array)
	'Creates an array of the workers selected in the dialog
	workers_to_XFER_cases_to = split(replace(workers_to_XFER_cases_to, " ", ""), ",")

	'Creates a new two-dimensional array for assigning a worker to each MAXIS_case_number, and collecting the county code for each worker
	Dim transfer_array()
	ReDim transfer_array(ubound(case_number_array), 2)

	'Assigns a MAXIS_case_number to each row in the first column of the array
	For x = 0 to ubound(case_number_array)
		transfer_array(x, 0) = case_number_array(x)
	Next

	'Reassigning x as a 0 for the following do...loop
	x = 0

	'Assigning y as 0, to be used by the following do...loop for deciding which worker gets which case
	y = 0

	'Now, it'll assign a worker to each case number in the transfer_array. Does this on a loop so that a worker can get multiple cases if that is indicated.
	Do
		transfer_array(x, 1) = workers_to_XFER_cases_to(y)	'Assigns column 2 of the array to a worker in the workers_to_XFER_cases_to array
		x = x + 1											'Adds +1 to X
		y = y + 1											'Adds +1 to Y
		If y > ubound(workers_to_XFER_cases_to) then y = 0	'Resets to allow the first worker in the array to get anonther one
	Loop until x > ubound(case_number_array)

	'--------Now, the array is two columns (MAXIS_case_number, worker_assigned)!

'	'Script must figure out who the current worker is, and what agency they are with. This is vital because transferring within an agency uses different screens than inter-agency.
'		'To do this, the script will start by analysing the current worker in REPT/ACTV.

	'Now, the array will figure out the current worker, by looking it up in SELF
	back_to_SELF
	EMReadScreen current_user, 7, 22, 8

	'Now, it will go to REPT/USER, and look up the county code for this individual.
	call navigate_to_MAXIS_screen("REPT", "USER")
	EMWriteScreen current_user, 21, 12
	transmit

	'Now, it will read the county code for the current user
	EMReadScreen user_county_code, 2, 7, 38

	'Now, we enter each worker number into this screen, and return their county codes inside the array.
	For x = 0 to ubound(case_number_array)
		EMWriteScreen transfer_array(x, 1), 21, 12		'Writes worker
		transmit										'Gets to next screen
		EMReadScreen array_county_code, 2, 7, 38		'Reads the county code for this worker
		transfer_array(x, 2) = array_county_code		'Adds the array_county_code to x, 2 of the transfer array
	Next

	'Resetting "x" to be a zero placeholder for the following for...next
	x = 0

	'Now we actually transfer the cases. This for...next does the work (details in comments below)
	For x = 0 to ubound(case_number_array)		'case_number_array is the same as the first col of the transfer_array
		'Assigns the number from the array to the MAXIS_case_number variable
		MAXIS_case_number = transfer_array(x, 0)

		'Checks to make sure case isn't in background
		MAXIS_background_check

		'Determines interagency transfers by comparing the current active user (gathered above) to the user in the transfer array.
		If user_county_code = transfer_array(x, 2) then
			county_to_county_XFER = False
		Else
			county_to_county_XFER = True
		End if

		'Getting to SPEC/XFER manually
		back_to_SELF
		EMWriteScreen "SPEC", 16, 43
		EMWriteScreen "________", 18, 43
		EMWriteScreen MAXIS_case_number, 18, 43
		EMWriteScreen "XFER", 21, 70
		transmit

		'Now to transfer the cases.
		If county_to_county_XFER = False then
			EMWriteScreen "x", 7, 16
			transmit
			PF9
			EMWriteScreen transfer_array(x, 1), 18, 61
			transmit
			transmit
		Else
			EMWriteScreen "x", 9, 16
			transmit
			PF9
			call create_MAXIS_friendly_date(date, 0, 4, 28)
			call create_MAXIS_friendly_date(date, 0, 4, 61)
			EMWriteScreen "N", 5, 28
			call create_MAXIS_friendly_date(date, 0, 5, 61)
			EMWriteScreen transfer_array(x, 1), 18, 61

			'Handling County of Financial Responsibility for HC and CASH applications
			EMReadScreen cfr, 2, 18, 63
			IF IsNumeric(cfr) = False THEN 	'If the recipient of the XFER is a state worker, the CFR is randomly generated.
				DO
					cfr_valid = False
					Randomize
					cfr = Int(100*Rnd)
					IF (cfr > 0 AND cfr < 89) OR cfr = 92 OR cfr = 93 THEN			'Determining that the CFR is a valid county
						IF len(cfr) = 1 THEN cfr = "0" & cfr
						cfr_valid = True
					END IF
				LOOP UNTIL cfr_valid = True
			END IF
			EMWriteScreen cfr, 11, 39		'For CASH
			EMWriteScreen cfr, 14, 39		'For HC
			cfr_month = DatePart("M", date)
				IF len(cfr_month) = 1 THEN cfr_month = "0" & cfr_month
			EMWriteScreen cfr_month, 11, 53		'For CASH
			EMWriteScreen right(DatePart("YYYY", date), 2), 11, 59
			EMWriteScreen cfr_month, 14, 53		'FOR HC
			EMWriteScreen right(DatePart("YYYY", date), 2), 14, 59
			transmit
			transmit
		End if
	Next
End function

'This is the custom function to load the dialog. It needs to be in a custom function because the user can change the excel file
FUNCTION create_dialog(training_case_creator_excel_file_path, scenario_list, scenario_dropdown, approve_case_dropdown, how_many_cases_to_make, XFER_check, workers_to_XFER_cases_to, reload_excel_file_button, ButtonPressed)
	'DIALOGS-----------------------------------------------------------------------------------------------------------
	'NOTE: droplistbox for scenario list must be: ["select one..." & scenario_list] in order to be dynamic
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 446, 125, "Training case creator dialog"
	  EditBox 85, 5, 295, 15, training_case_creator_excel_file_path
	  DropListBox 60, 40, 140, 15, "select one..." & scenario_list, scenario_dropdown
	  DropListBox 275, 40, 165, 15, "yes, approve all cases"+chr(9)+"no, but enter all STAT panels needed to approve"+chr(9)+"no, but do TYPE/PROG/REVW"+chr(9)+"no, just APPL all cases", approve_case_dropdown
	  EditBox 125, 60, 40, 15, how_many_cases_to_make
	  CheckBox 205, 65, 210, 10, "Check here to XFER cases, and enter worker numbers below.", XFER_check
	  EditBox 130, 80, 310, 15, workers_to_XFER_cases_to
	  ButtonGroup ButtonPressed
		OkButton 335, 105, 50, 15
		CancelButton 390, 105, 50, 15
		PushButton 385, 5, 55, 15, "Reload details", reload_excel_file_button
	  Text 5, 10, 75, 10, "File path of Excel file:"
	  Text 130, 25, 310, 10, "Note: only reload values if you've changed the values on the spreadsheet since opening."
	  Text 5, 45, 55, 10, "Scenario to run:"
	  Text 210, 45, 65, 10, "App/XFER cases?:"
	  Text 5, 65, 120, 10, "How many cases are you creating?:"
	  Text 5, 85, 125, 10, "Workers to XFER cases to (x1#####):"
	  Text 5, 100, 325, 20, "Please note: if you just wrote a scenario on the spreadsheet, it is recommended that you ''test'' it first by running a single case through. DHS staff cannot triage issues with agency-written scenarios."
	EndDialog

    Do
	   DIALOG Dialog1
       CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

END FUNCTION

'VARIABLES TO DECLARE-----------------------------------------------------------------------
how_many_cases_to_make = "1"		'Defaults to 1, but users can modify this.

'--------------------------------------- Project Krabappel ---------------------------------------
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 381, 125, "Training case creator"
  Text 10, 10, 350, 20, "Hello! Thanks for clicking the training case creator! This script will create training cases in the training region for testing purposes. This script uses an Excel template with already built scenarions. "
  Text 10, 35, 350, 20, "NOTE: Due to system limitations MSA/SNAP cases may not have MSA budgeted into the SNAP budget for the initial month."
  Text 10, 60, 365, 10, "Good luck and have fun! Questions about this script can be sent to: HSPH.EWS.BlueZoneScripts@Hennepin.us"
  Text 10, 80, 140, 10, "Select an Excel file for training scenarios:"
  EditBox 150, 75, 175, 15, training_case_creator_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 330, 75, 45, 15, "Browse...", select_a_file_button
    OkButton 270, 105, 50, 15
    CancelButton 325, 105, 50, 15
EndDialog

'Show initial dialog
Do
    Do
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(training_case_creator_excel_file_path, ".xlsx")
    Loop until ButtonPressed = OK and training_case_creator_excel_file_path <> ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(training_case_creator_excel_file_path, True, True, ObjExcel, objWorkbook)

'Set objWorkSheet = objWorkbook.Worksheet
For Each objWorkSheet In objWorkbook.Worksheets
	If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
Next

'Connects to BlueZone
EMConnect ""
Do
    DO
    	DO
    		DO
    			CALL create_dialog(training_case_creator_excel_file_path, scenario_list, scenario_dropdown, approve_case_dropdown, how_many_cases_to_make, XFER_check, workers_to_XFER_cases_to, reload_excel_file_button, ButtonPressed)
    				If buttonpressed = cancel then stopscript
    				IF ButtonPressed = reload_excel_file_button THEN
    					'Reseting the scenario list
    					scenario_list = ""
    					'Closing the current, active version of Excel
    					objWorkbook.Close
    					objExcel.Quit

    					'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
    					Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
    					objExcel.Visible = True
    					Set objWorkbook = objExcel.Workbooks.Open(training_case_creator_excel_file_path) 'Opens an excel file from a specific URL
    					objExcel.DisplayAlerts = True

    					'Set objWorkSheet = objWorkbook.Worksheet
    					For Each objWorkSheet In objWorkbook.Worksheets
    						If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
    					Next
    				END IF
    		LOOP UNTIL ButtonPressed <> reload_excel_file_button
    		If scenario_dropdown = "select one..." AND ButtonPressed = -1 then MsgBox ("You must select a scenario from the dropdown!")
    	LOOP UNTIL ButtonPressed <> reload_excel_file_button
    	final_check_before_running = MsgBox("Here's what the scenario will try to create. Please review before proceeding:" & Chr(10) & Chr(10) & _
    									"Scenario selection: " & scenario_dropdown & Chr(10) & _
    									"Approving cases: " & approve_case_dropdown & Chr(10) & _
    									"Amt of cases to make: " & how_many_cases_to_make & Chr(10) & _
    									"Workers to XFER cases to: " & workers_to_XFER_cases_to & Chr(10) & Chr(10) & _
    									"It is VERY IMPORTANT to review these details before proceeding. It is also highly recommended that if you've created your own scenarios, " & _
    									"test them first creating a single case. This is to check to see if any details were missed on the spreadsheet. DHS CANNOT TRIAGE ISSUES WITH " & _
    									"COUNTY/AGENCY CUSTOMIZED SCENARIOS." & Chr(10) & Chr(10) & _
    									"Please also note that creating training cases can take a very long time. If you are creating hundreds of cases, you may want to run this " & _
    									"overnight, or on a secondary machine." & Chr(10) & Chr(10) & _
    									"If you are ready to continue, press ''Yes''. Otherwise, press ''no'' to return to the previous screen.", vbYesNo)
    LOOP UNTIL final_check_before_running = vbYes
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Activates worksheet based on user selection
objExcel.worksheets(scenario_dropdown).Activate

'Determines how many HH members there are, as this script can run for multiple-member households.
excel_col = 3																		'Col 3 is always the primary applicant's col
Do																					'Loops through each col looking for more HH members. If found, it adds one to the counter.
	If ObjExcel.Cells(2, excel_col).Value <> "" then excel_col = excel_col + 1		'Adds one so that the loop will check again
Loop until ObjExcel.Cells(2, excel_col).Value = ""									'Exits loop when we have no number in the MEMB col
total_membs = excel_col - 3															'minus 3 because we started on column 3

'Focuses BlueZone so that everyone can see what it's doing
EMFocus

'========================================================================APPL PANELS========================================================================
For cases_to_make = 1 to how_many_cases_to_make

	'Navigates to SELF, checks for MAXIS training, stops if not on MAXIS training
	back_to_self
	EMReadScreen training_region_check, 8, 22, 48
	If training_region_check <> "TRAINING" then script_end_procedure("You must be in the training region to use this script. It will now stop.")

	'Assigning the Excel info to variables for appl, and enters into MAXIS. It does this by first declaring a "starting row" variable for each section, and then
	'	each variable will be that row plus however far down it may be on the spreadsheet. This will enable future variable addition without having to modify
	'	hundreds of variable entries here.

	'Grabs APPL screen variables (APPL date, primary applicant name (memb 01))
	APPL_starting_excel_row = 4		'Starting row for APPL function pieces
	APPL_month = ObjExcel.Cells(APPL_starting_excel_row, 3).Value
	APPL_day = objExcel.Cells(APPL_starting_excel_row + 1, 3).Value
	APPL_last_name = ObjExcel.Cells(APPL_starting_excel_row + 2, 3).Value
	APPL_first_name = ObjExcel.Cells(APPL_starting_excel_row + 3, 3).Value
	APPL_middle_initial = ObjExcel.Cells(APPL_starting_excel_row + 4, 3).Value

	IF APPL_month = "CM" THEN
		month_modifier = 0
	ELSEIF APPL_month = "CM -1" THEN
		month_modifier = -1
	ELSEIF APPL_month = "CM -2" THEN
		month_modifier = -2
	ELSEIF APPL_month = "CM -3" THEN
		month_modifier = -3
	ELSEIF APPL_month = "CM -4" THEN
		month_modifier = -4
	ELSEIF APPL_month = "CM -5" THEN
		month_modifier = -5
	ELSEIF APPL_month = "CM -6" THEN
		month_modifier = -6
		second_span = True
	ELSEIF APPL_month = "CM -7" THEN
		month_modifier = -7
		second_span = True
	ELSEIF APPL_month = "CM -8" THEN
		month_modifier = -8
		second_span = True
	ELSEIF APPL_month = "CM -9" THEN
		month_modifier = -9
		second_span = True
	ELSEIF APPL_month = "CM -10" THEN
		month_modifier = -10
		second_span = True
	ELSEIF APPL_month = "CM -11" THEN
		month_modifier = -11
		second_span = True
	END IF

	APPL_date = DateAdd("M", month_modifier, date)
	APPL_date = DatePart("M", APPL_date) & "/" & APPL_day & "/" & DatePart("YYYY", APPL_date)
	APPL_date = DateAdd("D", 0, APPL_date)

	'Gets the footer month and year of the application off of the spreadsheet, enters into SELF and transmits (can only enter an application on APPL in the footer month of app)
	MAXIS_footer_month = DatePart("M", APPL_date)
	IF len(MAXIS_footer_month) = 1 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
	If right(MAXIS_footer_month, 1) = "/" then MAXIS_footer_month = "0" & left(MAXIS_footer_month, 1)		'Does this to account for single digit months
	MAXIS_footer_year = right(APPL_date, 2)
	EMWriteScreen MAXIS_footer_month, 20, 43
	EMWriteScreen MAXIS_footer_year, 20, 46
	transmit

	'Goes to APPL function
	call navigate_to_MAXIS_screen("APPL", "____")

	'Enters info in APPL and transmits
	date_of_app = APPL_date 			'Sets a variable with the correct date format for calculation for determining client age
	call create_MAXIS_friendly_date(APPL_date, 0, 4, 63)
	EMWriteScreen APPL_last_name, 7, 30
	EMWriteScreen APPL_first_name, 7, 63
	EMWriteScreen APPL_middle_initial, 7, 79
	transmit

	'Uses a for...next to enter each HH member's info
	For current_memb = 1 to total_membs
		current_excel_col = current_memb + 2							'There's two columns before the first HH member, so we have to add 2 to get the current excel col
		reference_number = ObjExcel.Cells(2, current_excel_col).Value	'Always in the second row. This is the HH member number

		'Gets MEMB info for the current household member using the current_excel_col field. Starts by declaring the MEMB starting row
		MEMB_starting_excel_row = 6
		MEMB_last_name = ObjExcel.Cells(MEMB_starting_excel_row, current_excel_col).Value
		MEMB_first_name = ObjExcel.Cells(MEMB_starting_excel_row + 1, current_excel_col).Value
		MEMB_mid_init = ObjExcel.Cells(MEMB_starting_excel_row + 2, current_excel_col).Value
		MEMB_dob_mm_dd = ObjExcel.Cells(MEMB_starting_excel_row + 3, current_excel_col).Value
		MEMB_age = ObjExcel.Cells(MEMB_starting_excel_row + 4, current_excel_col).Value
		MEMB_DOB_verif = ObjExcel.Cells(MEMB_starting_excel_row + 5, current_excel_col).Value
		MEMB_gender = ObjExcel.Cells(MEMB_starting_excel_row + 6, current_excel_col).Value
		MEMB_ID_verif = ObjExcel.Cells(MEMB_starting_excel_row + 7, current_excel_col).Value
		MEMB_rel_to_appl = left(ObjExcel.Cells(MEMB_starting_excel_row + 8, current_excel_col).Value, 2)
		MEMB_spoken_lang = left(ObjExcel.Cells(MEMB_starting_excel_row + 9, current_excel_col).Value, 2)
		MEMB_interpreter_yn = ObjExcel.Cells(MEMB_starting_excel_row + 10, current_excel_col).Value
		MEMB_alias_yn = ObjExcel.Cells(MEMB_starting_excel_row + 11, current_excel_col).Value
		MEMB_hisp_lat_yn = ObjExcel.Cells(MEMB_starting_excel_row + 12, current_excel_col).Value
		If ObjExcel.Cells(345, current_excel_col).Value = "28 - Undocumented" Then Blank_IMIG = TRUE 	'Setting a variable for different process for undocumented

		'Gets MEMI info from spreadsheet
		MEMI_starting_excel_row = 19
		MEMI_marital_status = ObjExcel.Cells(MEMI_starting_excel_row, current_excel_col).Value
		MEMI_spouse = ObjExcel.Cells(MEMI_starting_excel_row + 1, current_excel_col).Value
		MEMI_last_grade_completed = ObjExcel.Cells(MEMI_starting_excel_row + 2, current_excel_col).Value
		MEMI_cit_yn = ObjExcel.Cells(MEMI_starting_excel_row + 3, current_excel_col).Value

		DO	'This DO-LOOP is to check that the CL's SSN created via random number generation is unique. If the SSN matches an SSN on file, the script creates a new SSN and re-enters the CL's information on MEMB. The checking for duplicates part is on the bottom, as that occurs when the worker presses transmit.
			IF Blank_IMIG <> True Then		'Non-undocumented client creation will have a SSN
				DO
					Randomize
					ssn_first = Rnd
					ssn_first = 1000000000 * ssn_first
					ssn_first = left(ssn_first, 3)
				LOOP UNTIL left(ssn_first, 1) <> "9"	'starting with a 9 is invalid
				Randomize
				ssn_mid = Rnd
				ssn_mid = 100000000 * ssn_mid
				ssn_mid = left(ssn_mid, 2)
				Randomize
				ssn_end = Rnd
				ssn_end = 100000000 * ssn_end
				ssn_end = left(ssn_end, 4)
			Else 						'If clt is undocumented - SSN  will be blank
				ssn_first = "   "
				ssn_mid = "  "
				ssn_end = "    "
			End If

			'Entering info on MEMB
			EMWriteScreen reference_number, 4, 33
			EMWriteScreen MEMB_last_name, 6, 30
			EMWriteScreen MEMB_first_name, 6, 63
			EMWriteScreen MEMB_mid_init, 6, 79
			EMWriteScreen ssn_first, 7, 42		'Determined above
			EMWriteScreen ssn_mid, 7, 46
			EMWriteScreen ssn_end, 7, 49
			EMWriteScreen "P", 7, 68			'All SSNs should pend in the training region
			If Blank_IMIG = TRUE Then EMWriteScreen "N", 7, 68		'If client is listed as undocumneted, no verif of a blank SSN
			'Generating the DOB
				year_of_birth = datepart("yyyy", date_of_app) - abs(MEMB_age)		'using date of application as the age listed should be age at appl
				IF MEMB_dob_mm_dd = "" THEN
					client_dob = "01/01/" & year_of_birth
				ELSE
					client_dob = DatePart("M", MEMB_dob_mm_dd) & "/" & DatePart("D", MEMB_dob_mm_dd) & "/" & year_of_birth
				END IF
				client_dob = DateAdd("D", 0, client_dob)
				CALL create_MAXIS_friendly_date_with_YYYY(client_dob, 0, 8, 42)
			'Continuing as normal
			EMWriteScreen MEMB_DOB_verif, 8, 68
			EMWriteScreen MEMB_gender, 9, 42
			EMWriteScreen MEMB_ID_verif, 9, 68
			EMWriteScreen MEMB_rel_to_appl, 10, 42
			EMWriteScreen MEMB_spoken_lang, 12, 42
			EMWriteScreen MEMB_spoken_lang, 13, 42
			EMWriteScreen MEMB_interpreter_yn, 14, 68
			EMWriteScreen MEMB_alias_yn, 15, 42
			IF MEMI_cit_yn = "N" AND Blank_IMIG <> TRUE THEN		'No Alien ID numbers for undocumneted
				MEMB_alien_ID = "A" & ssn_first & ssn_mid & ssn_end
				EMWriteScreen MEMB_alien_ID, 15, 68
			END IF
			EMWriteScreen MEMB_hisp_lat_yn, 16, 68
			EMWriteScreen "X", 17, 34			'Enters race as unknown at this time
			transmit
			DO				'Does this as a loop based on Robert's suggestion that there may be issues in loading without one. It's a small popup window.
				EMReadScreen race_mini_box, 18, 5, 12
				IF race_mini_box = "X AS MANY AS APPLY" THEN
					EMWriteScreen "X", 15, 12
					transmit
					transmit
				END IF
			LOOP UNTIL race_mini_box = "X AS MANY AS APPLY"
			cl_ssn = ssn_first & "-" & ssn_mid & "-" & ssn_end
			EMReadScreen ssn_match, 11, 8, 7
			IF cl_ssn <> ssn_match OR Blank_IMIG = TRUE THEN
				PF8
				PF8
				PF5
			ELSE
				PF3
			END IF
		LOOP UNTIL cl_ssn <> ssn_match  OR Blank_IMIG = TRUE
		EMWaitReady 0, 0
		EMWriteScreen "Y", 6, 67
		transmit

		Blank_IMIG = ""		'Blanking out for the next client

		'Updates MEMI with the info
		EMWriteScreen MEMI_marital_status, 7, 40
		EMWriteScreen MEMI_spouse, 9, 49
		EMWriteScreen MEMI_last_grade_completed, 10, 49
		EMWriteScreen MEMI_cit_yn, 11, 49
		EMWriteScreen "NO", 11, 78		'Always defaulting to none for cit/ID proof right now
		EMWriteScreen "Y", 14, 49		'Always defualting to yes for been in MN > 12 months
		EMWriteScreen "N", 14, 78		'Always defualting to no for residence verification
		transmit
	Next

	transmit 'This next transmit gets to the ADDR screen

	'Gets ADDR info from spreadsheet, gets from column 3 because it's case based
	ADDR_starting_excel_row = 23
	ADDR_line_one = ObjExcel.Cells(ADDR_starting_excel_row, 3).Value
	ADDR_line_two = ObjExcel.Cells(ADDR_starting_excel_row + 1, 3).Value
	ADDR_city = ObjExcel.Cells(ADDR_starting_excel_row + 2, 3).Value
	ADDR_zip = ObjExcel.Cells(ADDR_starting_excel_row + 3, 3).Value
	ADDR_county = ObjExcel.Cells(ADDR_starting_excel_row + 4, 3).Value
	ADDR_addr_verif = ObjExcel.Cells(ADDR_starting_excel_row + 5, 3).Value
	ADDR_homeless = ObjExcel.Cells(ADDR_starting_excel_row + 6, 3).Value
	ADDR_reservation = ObjExcel.Cells(ADDR_starting_excel_row + 7, 3).Value
	ADDR_mailing_addr_line_one = ObjExcel.Cells(ADDR_starting_excel_row + 8, 3).Value
	ADDR_mailing_addr_line_two = ObjExcel.Cells(ADDR_starting_excel_row + 9, 3).Value
	ADDR_mailing_addr_city = ObjExcel.Cells(ADDR_starting_excel_row + 10, 3).Value
	ADDR_mailing_addr_zip = ObjExcel.Cells(ADDR_starting_excel_row + 11, 3).Value
	ADDR_phone_1 = ObjExcel.Cells(ADDR_starting_excel_row + 12, 3).Value
	ADDR_phone_2 = ObjExcel.Cells(ADDR_starting_excel_row + 13, 3).Value
	ADDR_phone_3 = ObjExcel.Cells(ADDR_starting_excel_row + 14, 3).Value

	'Writes spreadsheet info to ADDR
	EMWriteScreen ADDR_line_one, 6, 43
	EMWriteScreen ADDR_line_two, 7, 43
	EMWriteScreen ADDR_city, 8, 43
	EMWriteScreen "MN", 8, 66		'Defaults to MN for all cases at this time
	EMWriteScreen ADDR_zip, 9, 43
	EMWriteScreen ADDR_county, 9, 66
	EMWriteScreen ADDR_addr_verif, 9, 74
	EMWriteScreen ADDR_homeless, 10, 43
	EMWriteScreen ADDR_reservation, 10, 74
	EMWriteScreen ADDR_mailing_addr_line_one, 13, 43
	EMWriteScreen ADDR_mailing_addr_line_two, 14, 43
	EMWriteScreen ADDR_mailing_addr_city, 15, 43
	If ADDR_mailing_addr_line_one <> "" then EMWriteScreen "MN", 16, 43	'Only writes if the user indicated a mailing address. Defaults to MN at this time.
	EMWriteScreen ADDR_mailing_addr_zip, 16, 52
	EMWriteScreen MAXIS_footer_month, 20, 43
	EMWriteScreen MAXIS_footer_year, 20, 46
	first_of_footer_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year		'We need to determine when we are updating what month we are in because these move
	first_of_footer_month = DateAdd("d", 0, first_of_footer_month)
	If DateDiff("d", first_of_footer_month, #10/1/2021#) > 0 Then 				''before 10/1/21 - they are here
		EMWriteScreen left(ADDR_phone_1, 3), 17, 45						'Has to split phone numbers up into three parts each
		EMWriteScreen mid(ADDR_phone_1, 5, 3), 17, 51
		EMWriteScreen right(ADDR_phone_1, 4), 17, 55
		EMWriteScreen left(ADDR_phone_2, 3), 18, 45
		EMWriteScreen mid(ADDR_phone_2, 5, 3), 18, 51
		EMWriteScreen right(ADDR_phone_2, 4), 18, 55
		EMWriteScreen left(ADDR_phone_3, 3), 19, 45
		EMWriteScreen mid(ADDR_phone_3, 5, 3), 19, 51
		EMWriteScreen right(ADDR_phone_3, 4), 19, 55
	Else 																		'10/1 and after will be this'
		EMWriteScreen left(ADDR_phone_1, 3), 16, 39						'Has to split phone numbers up into three parts each
		EMWriteScreen mid(ADDR_phone_1, 5, 3), 16, 45
		EMWriteScreen right(ADDR_phone_1, 4), 16, 49
		EMWriteScreen left(ADDR_phone_2, 3), 17, 39
		EMWriteScreen mid(ADDR_phone_2, 5, 3), 17, 45
		EMWriteScreen right(ADDR_phone_2, 4), 17, 49
		EMWriteScreen left(ADDR_phone_3, 3), 18, 39
		EMWriteScreen mid(ADDR_phone_3, 5, 3), 18, 45
		EMWriteScreen right(ADDR_phone_3, 4), 18, 49
	End If

	'Reads the case number and adds to an array before exiting
	EMReadScreen current_case_number, 8, 20, 37
	case_number_array = case_number_array & replace(current_case_number, "_", "") & "|"

	transmit
	EMReadScreen addr_warning, 7, 3, 6
	IF addr_warning = "Warning" THEN transmit
	transmit
	PF3

	STATS_counter = STATS_counter + 1   'counting each case made to supply the multiplier for stats.

Next

'Removing the last "|" from the case_number_array so as to avoid it trying to work a blank case number through PND1
case_number_array = left(case_number_array, len(case_number_array) - 1)

'Splitting the case numbers into an array
case_number_array = split(case_number_array, "|")

'Ends here if the user selected to just APPL all cases
If approve_case_dropdown = "no, just APPL all cases" then
	If XFER_check = checked then call transfer_cases(workers_to_XFER_cases_to, case_number_array)
	script_end_procedure("Success! Cases made and appl'd, per your request.")
End if
'========================================================================PND1 PANELS========================================================================


For each MAXIS_case_number in case_number_array
	'Navigates into STAT. For PND1 cases, this will trigger workflow for adding the right panels.
	call navigate_to_MAXIS_screen ("STAT", "____")

	'Transmits, to get to TYPE panel
	transmit

	'At this time, it will always mark GRH and IV-E as "N"
	EMWriteScreen "N", 6, 64	'GRH
	EMWriteScreen "N", 6, 73	'IV-E

	'Reading and writing info for the TYPE panel
	'Uses a for...next to enter each HH member's info
	For current_memb = 1 to total_membs
		current_excel_col = current_memb + 2							'There's two columns before the first HH member, so we have to add 2 to get the current excel col
		current_MAXIS_row = current_memb + 5							'MEMB 01 always gets entered on row 6, which each subsequent added to the following row. Adding 5 to current_memb simplifies this.
		'reference_number = ObjExcel.Cells(2, current_excel_col).Value	'Always in the second row. This is the HH member number

		'Reading the info
		TYPE_starting_excel_row = 38
		TYPE_cash_yn = objExcel.Cells(TYPE_starting_excel_row, current_excel_col).Value
		TYPE_hc_yn = objExcel.Cells(TYPE_starting_excel_row + 1, current_excel_col).Value
		TYPE_fs_yn = objExcel.Cells(TYPE_starting_excel_row + 2, current_excel_col).Value

		HCRE_starting_excel_row = 337
		HCRE_appl_addnd_date_input = ObjExcel.Cells(HCRE_starting_excel_row, current_excel_col).Value
		HCRE_retro_months_input = ObjExcel.Cells(HCRE_starting_excel_row + 1, current_excel_col).Value
		HCRE_recvd_by_service_date_input = ObjExcel.Cells(HCRE_starting_excel_row + 2, current_excel_col).Value

		'Writing the info
		EMWriteScreen TYPE_cash_yn, current_MAXIS_row, 28
		EMWriteScreen TYPE_hc_yn, current_MAXIS_row, 37
		EMWriteScreen TYPE_fs_yn, current_MAXIS_row, 46
		EMWriteScreen "N", current_MAXIS_row, 55			'At this time, it will always mark EMER as "N"

		'If any TYPE options are selected, we need to track this to know which items to type on PROG. If any are "Y", it'll update these variables.
		If ucase(TYPE_cash_yn) = "Y" then cash_application = True
		If ucase(TYPE_hc_yn) = "Y" then hc_application = True
		If ucase(TYPE_fs_yn) = "Y" then SNAP_application = True
	Next

	'Transmits to get to PROG
	transmit

	'Gathers the mig worker variable from Excel. Since it's the only one, we won't use a PROG starting row variable. And since it's case based, we'll only look in col 3
	PROG_mig_worker = objExcel.Cells(41, 3).Value

	'Enters in the APPL date on PROG for any programs applied for, and the interview date will always be the APPL date at this time.
	If cash_application = True then
		call create_MAXIS_friendly_date(APPL_date, 0, 6, 33)
		call create_MAXIS_friendly_date(APPL_date, 0, 6, 44)
		call create_MAXIS_friendly_date(APPL_date, 0, 6, 55)
	End if
	If SNAP_application = True then
		call create_MAXIS_friendly_date(APPL_date, 0, 10, 33)
		call create_MAXIS_friendly_date(APPL_date, 0, 10, 44)
		call create_MAXIS_friendly_date(APPL_date, 0, 10, 55)
	End if
	If HC_application = True then call create_MAXIS_friendly_date(APPL_date, 0, 12, 33)

	'Enters migrant worker info
	EMWriteScreen PROG_mig_worker, 18, 67
	DO
		transmit
		EMReadScreen still_on_prog, 4, 2, 50
	LOOP UNTIL still_on_prog <> "PROG"

	'If the case is HC, the script will handle HCRE
	If HC_application = True then Transmit 'MAXIS will navigate to a HCRE panel in edit mode if a PROG is completed showing HC.


	'Now we're on REVW and it needs to take different actions for each program. We need to know 6 month and 12 month dates though, for the sake of figuring out review months.
	'Scanning info from REPT section of spreadsheet
	REVW_starting_excel_row = 42
	REVW_ar_or_ir = objExcel.Cells(REVW_starting_excel_row, 3).Value	'Will return either a blank, an "IR", or an "AR"
	REVW_exempt = objExcel.Cells(REVW_starting_excel_row + 1, 3).Value	'Case based, so we'll only look at col 3

	'Determining those dates
	six_month_recert_date = dateadd("m", 6, APPL_date)							'Determines info for the six month recert
	six_month_month = datepart("m", six_month_recert_date)
	If len(six_month_month) = 1 then six_month_month = "0" & six_month_month
	six_month_year = right(six_month_recert_date, 2)
	one_year_recert_date = dateadd("m", 12, APPL_date)							'Determines info for the annual recert
	one_year_month = datepart("m", one_year_recert_date)
	If len(one_year_month) = 1 then one_year_month = "0" & one_year_month
	one_year_year = right(one_year_recert_date, 2)

	'Adds cash dates
	If cash_application = true then
		EMWriteScreen one_year_month, 9, 37
		EMWriteScreen one_year_year, 9, 43
	End if

	'Adds SNAP dates and info
	If SNAP_application = true then
		EMWriteScreen "N", 15, 75		'Phone interview field
		EMWriteScreen "x", 5, 58
		transmit
		EMWriteScreen six_month_month, 9, 26
		EMWriteScreen six_month_year, 9, 32
		EMWriteScreen one_year_month, 9, 64
		EMWriteScreen one_year_year, 9, 70
		transmit
		transmit
	End if

	'Adds HC dates and info
	If HC_application = true then
		EMWriteScreen "x", 5, 71
		transmit
		If REVW_ar_or_ir = "IR" then
			EMWriteScreen six_month_month, 8, 27
			EMWriteScreen six_month_year, 8, 33
		ElseIf REVW_ar_or_ir = "AR" then
			EMWriteScreen six_month_month, 8, 71
			EMWriteScreen six_month_year, 8, 77
		ElseIf REVW_ar_or_ir = "ER Only" then
			EMWriteScreen one_year_month, 8, 27
			EMWriteScreen one_year_year, 8, 33
		End if
		EMWriteScreen one_year_month, 9, 27
		EMWriteScreen one_year_year, 9, 33
		EMWriteScreen REVW_exempt, 9, 71
		transmit
		transmit
	End if

	transmit
	transmit
	STATS_manualtime = STATS_manualtime + 40   'adding manualtime for processing PROG/REVW


Next

Call date_array_generator(MAXIS_footer_month, MAXIS_footer_year, date_array)

'Ends here if the user selected to just do TYPE/PROG/REVW for all cases
If approve_case_dropdown = "no, but do TYPE/PROG/REVW" then
	If XFER_check = checked then call transfer_cases(workers_to_XFER_cases_to, case_number_array)
	script_end_procedure("Success! Cases made and appl'd, and TYPE/PROG/REVW updated, per your request.")
End if
'========================================================================PND2 PANELS========================================================================

For each MAXIS_case_number in case_number_array

	'Navigates to STAT/SUMM for each
	call navigate_to_MAXIS_screen("STAT", "SUMM")
	MAXIS_background_check
	EMReadScreen SELF_check, 4, 2, 50
	EMReadScreen ERRR_check, 4, 2, 52	'Extra err handling in case the case was in background
	If ERRR_check = "ERRR" then transmit

	'Uses a for...next to enter each HH member's info (person based panels only
	For current_memb = 1 to total_membs
		current_excel_col = current_memb + 2							'There's two columns before the first HH member, so we have to add 2 to get the current excel col
		reference_number = ObjExcel.Cells(2, current_excel_col).Value	'Always in the second row. This is the HH member number

		'--------------READS ENTIRE EXCEL SHEET FOR THIS HH MEMB
		ABPS_starting_excel_row = 44
		ABPS_supp_coop = ObjExcel.Cells(ABPS_starting_excel_row, current_excel_col).Value
		ABPS_gc_status = ObjExcel.Cells(ABPS_starting_excel_row + 1, current_excel_col).Value

		ACCT_starting_excel_row = 46
		ACCT_type = left(ObjExcel.Cells(ACCT_starting_excel_row, current_excel_col).Value, 2)
		ACCT_numb = ObjExcel.Cells(ACCT_starting_excel_row + 1, current_excel_col).Value
		ACCT_location = ObjExcel.Cells(ACCT_starting_excel_row + 2, current_excel_col).Value
		ACCT_balance = ObjExcel.Cells(ACCT_starting_excel_row + 3, current_excel_col).Value
		ACCT_bal_ver = left(ObjExcel.Cells(ACCT_starting_excel_row + 4, current_excel_col).Value, 1)
		ACCT_date = ObjExcel.Cells(ACCT_starting_excel_row + 5, current_excel_col).Value
		ACCT_withdraw = ObjExcel.Cells(ACCT_starting_excel_row + 6, current_excel_col).Value
		ACCT_cash_count = ObjExcel.Cells(ACCT_starting_excel_row + 7, current_excel_col).Value
		ACCT_snap_count = ObjExcel.Cells(ACCT_starting_excel_row + 8, current_excel_col).Value
		ACCT_HC_count = ObjExcel.Cells(ACCT_starting_excel_row + 9, current_excel_col).Value
		ACCT_GRH_count = ObjExcel.Cells(ACCT_starting_excel_row + 10, current_excel_col).Value
		ACCT_IV_count = ObjExcel.Cells(ACCT_starting_excel_row + 11, current_excel_col).Value
		ACCT_joint_owner = ObjExcel.Cells(ACCT_starting_excel_row + 12, current_excel_col).Value
		ACCT_share_ratio = ObjExcel.Cells(ACCT_starting_excel_row + 13, current_excel_col).Value
		ACCT_interest_date_mo = ObjExcel.Cells(ACCT_starting_excel_row + 14, current_excel_col).Value
		ACCT_interest_date_yr = ObjExcel.Cells(ACCT_starting_excel_row + 15, current_excel_col).Value

		ACUT_starting_excel_row = 62
		ACUT_shared = ObjExcel.Cells(ACUT_starting_excel_row, current_excel_col).Value
		ACUT_heat = ObjExcel.Cells(ACUT_starting_excel_row + 1, current_excel_col).Value
		ACUT_heat_verif = ObjExcel.Cells(ACUT_starting_excel_row + 2, current_excel_col).Value
		ACUT_air = ObjExcel.Cells(ACUT_starting_excel_row + 3, current_excel_col).Value
		ACUT_air_verif = ObjExcel.Cells(ACUT_starting_excel_row + 4, current_excel_col).Value
		ACUT_electric = ObjExcel.Cells(ACUT_starting_excel_row + 5, current_excel_col).Value
		ACUT_electric_verif = ObjExcel.Cells(ACUT_starting_excel_row + 6, current_excel_col).Value
		ACUT_fuel = ObjExcel.Cells(ACUT_starting_excel_row + 7, current_excel_col).Value
		ACUT_fuel_verif = ObjExcel.Cells(ACUT_starting_excel_row + 8, current_excel_col).Value
		ACUT_garbage = ObjExcel.Cells(ACUT_starting_excel_row + 9, current_excel_col).Value
		ACUT_garbage_verif = ObjExcel.Cells(ACUT_starting_excel_row + 10, current_excel_col).Value
		ACUT_water = ObjExcel.Cells(ACUT_starting_excel_row + 11, current_excel_col).Value
		ACUT_water_verif = ObjExcel.Cells(ACUT_starting_excel_row + 12, current_excel_col).Value
		ACUT_sewer = ObjExcel.Cells(ACUT_starting_excel_row + 13, current_excel_col).Value
		ACUT_sewer_verif = ObjExcel.Cells(ACUT_starting_excel_row + 14, current_excel_col).Value
		ACUT_other = ObjExcel.Cells(ACUT_starting_excel_row + 15, current_excel_col).Value
		ACUT_other_verif = ObjExcel.Cells(ACUT_starting_excel_row + 16, current_excel_col).Value
		ACUT_phone = ObjExcel.Cells(ACUT_starting_excel_row + 17, current_excel_col).Value

		BILS_starting_excel_row = 80
		BILS_bill_1_ref_num = objExcel.Cells(BILS_starting_excel_row, current_excel_col).Value
		BILS_bill_1_serv_date = objExcel.Cells(BILS_starting_excel_row + 1, current_excel_col).Value
		BILS_bill_1_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 2, current_excel_col).Value, 2)
		BILS_bill_1_gross_amt = objExcel.Cells(BILS_starting_excel_row + 3, current_excel_col).Value
		BILS_bill_1_third_party = objExcel.Cells(BILS_starting_excel_row + 4, current_excel_col).Value
		BILS_bill_1_verif = objExcel.Cells(BILS_starting_excel_row + 5, current_excel_col).Value
		BILS_bill_1_BILS_type = objExcel.Cells(BILS_starting_excel_row + 6, current_excel_col).Value
		BILS_bill_2_ref_num = objExcel.Cells(BILS_starting_excel_row + 7, current_excel_col).Value
		BILS_bill_2_serv_date = objExcel.Cells(BILS_starting_excel_row + 8, current_excel_col).Value
		BILS_bill_2_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 9, current_excel_col).Value, 2)
		BILS_bill_2_gross_amt = objExcel.Cells(BILS_starting_excel_row + 10, current_excel_col).Value
		BILS_bill_2_third_party = objExcel.Cells(BILS_starting_excel_row + 11, current_excel_col).Value
		BILS_bill_2_verif = objExcel.Cells(BILS_starting_excel_row + 12, current_excel_col).Value
		BILS_bill_2_BILS_type = objExcel.Cells(BILS_starting_excel_row + 13, current_excel_col).Value
		BILS_bill_3_ref_num = objExcel.Cells(BILS_starting_excel_row + 14, current_excel_col).Value
		BILS_bill_3_serv_date = objExcel.Cells(BILS_starting_excel_row + 15, current_excel_col).Value
		BILS_bill_3_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 16, current_excel_col).Value, 2)
		BILS_bill_3_gross_amt = objExcel.Cells(BILS_starting_excel_row + 17, current_excel_col).Value
		BILS_bill_3_third_party = objExcel.Cells(BILS_starting_excel_row + 18, current_excel_col).Value
		BILS_bill_3_verif = objExcel.Cells(BILS_starting_excel_row + 19, current_excel_col).Value
		BILS_bill_3_BILS_type = objExcel.Cells(BILS_starting_excel_row + 20, current_excel_col).Value
		BILS_bill_4_ref_num = objExcel.Cells(BILS_starting_excel_row + 21, current_excel_col).Value
		BILS_bill_4_serv_date = objExcel.Cells(BILS_starting_excel_row + 22, current_excel_col).Value
		BILS_bill_4_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 23, current_excel_col).Value, 2)
		BILS_bill_4_gross_amt = objExcel.Cells(BILS_starting_excel_row + 24, current_excel_col).Value
		BILS_bill_4_third_party = objExcel.Cells(BILS_starting_excel_row + 25, current_excel_col).Value
		BILS_bill_4_verif = objExcel.Cells(BILS_starting_excel_row + 26, current_excel_col).Value
		BILS_bill_4_BILS_type = objExcel.Cells(BILS_starting_excel_row + 27, current_excel_col).Value
		BILS_bill_5_ref_num = objExcel.Cells(BILS_starting_excel_row + 28, current_excel_col).Value
		BILS_bill_5_serv_date = objExcel.Cells(BILS_starting_excel_row + 29, current_excel_col).Value
		BILS_bill_5_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 30, current_excel_col).Value, 2)
		BILS_bill_5_gross_amt = objExcel.Cells(BILS_starting_excel_row + 31, current_excel_col).Value
		BILS_bill_5_third_party = objExcel.Cells(BILS_starting_excel_row + 32, current_excel_col).Value
		BILS_bill_5_verif = objExcel.Cells(BILS_starting_excel_row + 33, current_excel_col).Value
		BILS_bill_5_BILS_type = objExcel.Cells(BILS_starting_excel_row + 34, current_excel_col).Value
		BILS_bill_6_ref_num = objExcel.Cells(BILS_starting_excel_row + 35, current_excel_col).Value
		BILS_bill_6_serv_date = objExcel.Cells(BILS_starting_excel_row + 36, current_excel_col).Value
		BILS_bill_6_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 37, current_excel_col).Value, 2)
		BILS_bill_6_gross_amt = objExcel.Cells(BILS_starting_excel_row + 38, current_excel_col).Value
		BILS_bill_6_third_party = objExcel.Cells(BILS_starting_excel_row + 39, current_excel_col).Value
		BILS_bill_6_verif = objExcel.Cells(BILS_starting_excel_row + 40, current_excel_col).Value
		BILS_bill_6_BILS_type = objExcel.Cells(BILS_starting_excel_row + 41, current_excel_col).Value
		BILS_bill_7_ref_num = objExcel.Cells(BILS_starting_excel_row + 42, current_excel_col).Value
		BILS_bill_7_serv_date = objExcel.Cells(BILS_starting_excel_row + 43, current_excel_col).Value
		BILS_bill_7_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 44, current_excel_col).Value, 2)
		BILS_bill_7_gross_amt = objExcel.Cells(BILS_starting_excel_row + 45, current_excel_col).Value
		BILS_bill_7_third_party = objExcel.Cells(BILS_starting_excel_row + 46, current_excel_col).Value
		BILS_bill_7_verif = objExcel.Cells(BILS_starting_excel_row + 47, current_excel_col).Value
		BILS_bill_7_BILS_type = objExcel.Cells(BILS_starting_excel_row + 48, current_excel_col).Value
		BILS_bill_8_ref_num = objExcel.Cells(BILS_starting_excel_row + 49, current_excel_col).Value
		BILS_bill_8_serv_date = objExcel.Cells(BILS_starting_excel_row + 50, current_excel_col).Value
		BILS_bill_8_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 51, current_excel_col).Value, 2)
		BILS_bill_8_gross_amt = objExcel.Cells(BILS_starting_excel_row + 52, current_excel_col).Value
		BILS_bill_8_third_party = objExcel.Cells(BILS_starting_excel_row + 53, current_excel_col).Value
		BILS_bill_8_verif = objExcel.Cells(BILS_starting_excel_row + 54, current_excel_col).Value
		BILS_bill_8_BILS_type = objExcel.Cells(BILS_starting_excel_row + 55, current_excel_col).Value
		BILS_bill_9_ref_num = objExcel.Cells(BILS_starting_excel_row + 56, current_excel_col).Value
		BILS_bill_9_serv_date = objExcel.Cells(BILS_starting_excel_row + 57, current_excel_col).Value
		BILS_bill_9_serv_type = left(objExcel.Cells(BILS_starting_excel_row + 58, current_excel_col).Value, 2)
		BILS_bill_9_gross_amt = objExcel.Cells(BILS_starting_excel_row + 59, current_excel_col).Value
		BILS_bill_9_third_party = objExcel.Cells(BILS_starting_excel_row + 60, current_excel_col).Value
		BILS_bill_9_verif = objExcel.Cells(BILS_starting_excel_row + 61, current_excel_col).Value
		BILS_bill_9_BILS_type = objExcel.Cells(BILS_starting_excel_row + 62, current_excel_col).Value

		BUSI_starting_excel_row = 143
		BUSI_type = left(ObjExcel.Cells(BUSI_starting_excel_row, current_excel_col).Value, 2)
		BUSI_start_date = ObjExcel.Cells(BUSI_starting_excel_row + 1, current_excel_col).Value
		BUSI_end_date = ObjExcel.Cells(BUSI_starting_excel_row + 2, current_excel_col).Value
		BUSI_cash_total_retro = ObjExcel.Cells(BUSI_starting_excel_row + 3, current_excel_col).Value
		BUSI_cash_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 4, current_excel_col).Value
		BUSI_cash_total_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 5, current_excel_col).Value, 1)
		BUSI_IV_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 6, current_excel_col).Value
		BUSI_IV_total_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 7, current_excel_col).Value, 1)
		BUSI_snap_total_retro = ObjExcel.Cells(BUSI_starting_excel_row + 8, current_excel_col).Value
		BUSI_snap_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 9, current_excel_col).Value
		BUSI_snap_total_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 10, current_excel_col).Value, 1)
		BUSI_hc_total_prosp_a = ObjExcel.Cells(BUSI_starting_excel_row + 11, current_excel_col).Value
		BUSI_hc_total_ver_a = left(ObjExcel.Cells(BUSI_starting_excel_row + 12, current_excel_col).Value, 1)
		BUSI_hc_total_prosp_b = ObjExcel.Cells(BUSI_starting_excel_row + 13, current_excel_col).Value
		BUSI_hc_total_ver_b = left(ObjExcel.Cells(BUSI_starting_excel_row + 14, current_excel_col).Value, 1)
		BUSI_cash_exp_retro = ObjExcel.Cells(BUSI_starting_excel_row + 15, current_excel_col).Value
		BUSI_cash_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 16, current_excel_col).Value
		BUSI_cash_exp_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 17, current_excel_col).Value, 1)
		BUSI_IV_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 18, current_excel_col).Value
		BUSI_IV_exp_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 19, current_excel_col).Value, 1)
		BUSI_snap_exp_retro = ObjExcel.Cells(BUSI_starting_excel_row + 20, current_excel_col).Value
		BUSI_snap_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 21, current_excel_col).Value
		BUSI_snap_exp_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 22, current_excel_col).Value, 1)
		BUSI_hc_exp_prosp_a = ObjExcel.Cells(BUSI_starting_excel_row + 23, current_excel_col).Value
		BUSI_hc_exp_ver_a = left(ObjExcel.Cells(BUSI_starting_excel_row + 24, current_excel_col).Value, 1)
		BUSI_hc_exp_prosp_b = ObjExcel.Cells(BUSI_starting_excel_row + 25, current_excel_col).Value
		BUSI_hc_exp_ver_b = left(ObjExcel.Cells(BUSI_starting_excel_row + 26, current_excel_col).Value, 1)
		BUSI_retro_hours = ObjExcel.Cells(BUSI_starting_excel_row + 27, current_excel_col).Value
		BUSI_prosp_hours = ObjExcel.Cells(BUSI_starting_excel_row + 28, current_excel_col).Value
		BUSI_hc_total_est_a = ObjExcel.Cells(BUSI_starting_excel_row + 29, current_excel_col).Value
		BUSI_hc_total_est_b = ObjExcel.Cells(BUSI_starting_excel_row + 30, current_excel_col).Value
		BUSI_hc_exp_est_a = ObjExcel.Cells(BUSI_starting_excel_row + 31, current_excel_col).Value
		BUSI_hc_exp_est_b = ObjExcel.Cells(BUSI_starting_excel_row + 32, current_excel_col).Value
		BUSI_hc_hours_est = ObjExcel.Cells(BUSI_starting_excel_row + 33, current_excel_col).Value

		CARS_starting_excel_row = 177
		CARS_type = LEFT(ObjExcel.Cells(CARS_starting_excel_row, current_excel_col).Value, 1)
		CARS_year = ObjExcel.Cells(CARS_starting_excel_row + 1, current_excel_col).Value
		CARS_make = ObjExcel.Cells(CARS_starting_excel_row + 2, current_excel_col).Value
		CARS_model = ObjExcel.Cells(CARS_starting_excel_row + 3, current_excel_col).Value
		CARS_trade_in = ObjExcel.Cells(CARS_starting_excel_row + 4, current_excel_col).Value
		CARS_loan = ObjExcel.Cells(CARS_starting_excel_row + 5, current_excel_col).Value
		CARS_value_source = left(ObjExcel.Cells(CARS_starting_excel_row + 6, current_excel_col).Value, 1)
		CARS_ownership_ver = left(ObjExcel.Cells(CARS_starting_excel_row + 7, current_excel_col).Value, 1)
		CARS_amount_owed = ObjExcel.Cells(CARS_starting_excel_row + 8, current_excel_col).Value
		CARS_amount_owed_ver = left(ObjExcel.Cells(CARS_starting_excel_row + 9, current_excel_col).Value, 1)
		CARS_date = ObjExcel.Cells(CARS_starting_excel_row + 10, current_excel_col).Value
		CARS_use = left(ObjExcel.Cells(CARS_starting_excel_row + 11, current_excel_col).Value, 1)
		CARS_HC_benefit = ObjExcel.Cells(CARS_starting_excel_row + 12, current_excel_col).Value
		CARS_joint_owner = ObjExcel.Cells(CARS_starting_excel_row + 13, current_excel_col).Value
		CARS_share_ratio = ObjExcel.Cells(CARS_starting_excel_row + 14, current_excel_col).Value

		CASH_starting_excel_row = 192
		CASH_amount = ObjExcel.Cells(CASH_starting_excel_row, current_excel_col).Value

		COEX_starting_excel_row = 193
		COEX_support_retro = ObjExcel.Cells(COEX_starting_excel_row, current_excel_col).Value
		COEX_support_prosp = ObjExcel.Cells(COEX_starting_excel_row + 1, current_excel_col).Value
		COEX_support_verif = left(ObjExcel.Cells(COEX_starting_excel_row + 2, current_excel_col).Value, 1)
		COEX_alimony_retro = ObjExcel.Cells(COEX_starting_excel_row + 3, current_excel_col).Value
		COEX_alimony_prosp = ObjExcel.Cells(COEX_starting_excel_row + 4, current_excel_col).Value
		COEX_alimony_verif = left(ObjExcel.Cells(COEX_starting_excel_row + 5, current_excel_col).Value, 1)
		COEX_tax_dep_retro = ObjExcel.Cells(COEX_starting_excel_row + 6, current_excel_col).Value
		COEX_tax_dep_prosp = ObjExcel.Cells(COEX_starting_excel_row + 7, current_excel_col).Value
		COEX_tax_dep_verif = left(ObjExcel.Cells(COEX_starting_excel_row + 8, current_excel_col).Value, 1)
		COEX_other_retro = ObjExcel.Cells(COEX_starting_excel_row + 9, current_excel_col).Value
		COEX_other_prosp = ObjExcel.Cells(COEX_starting_excel_row + 10, current_excel_col).Value
		COEX_other_verif = left(ObjExcel.Cells(COEX_starting_excel_row + 11, current_excel_col).Value, 1)
		COEX_change_in_circumstances = left(ObjExcel.Cells(COEX_starting_excel_row + 12, current_excel_col).Value, 1)
		COEX_HC_expense_support = ObjExcel.Cells(COEX_starting_excel_row + 13, current_excel_col).Value
		COEX_HC_expense_alimony = ObjExcel.Cells(COEX_starting_excel_row + 14, current_excel_col).Value
		COEX_HC_expense_tax_dep = ObjExcel.Cells(COEX_starting_excel_row + 15, current_excel_col).Value
		COEX_HC_expense_other = ObjExcel.Cells(COEX_starting_excel_row + 16, current_excel_col).Value

		DCEX_starting_excel_row = 210
		DCEX_provider = ObjExcel.Cells(DCEX_starting_excel_row, current_excel_col).Value
		DCEX_reason = left(ObjExcel.Cells(DCEX_starting_excel_row + 1, current_excel_col).Value, 1)
		DCEX_subsidy = left(ObjExcel.Cells(DCEX_starting_excel_row + 2, current_excel_col).Value, 1)
		DCEX_child_number1 = ObjExcel.Cells(DCEX_starting_excel_row + 3, current_excel_col).Value
		DCEX_child_number1_retro = ObjExcel.Cells(DCEX_starting_excel_row + 4, current_excel_col).Value
		DCEX_child_number1_pro = ObjExcel.Cells(DCEX_starting_excel_row + 5, current_excel_col).Value
		DCEX_child_number1_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 6, current_excel_col).Value, 1)
		DCEX_child_number2 = ObjExcel.Cells(DCEX_starting_excel_row + 7, current_excel_col).Value
		DCEX_child_number2_retro = ObjExcel.Cells(DCEX_starting_excel_row + 8, current_excel_col).Value
		DCEX_child_number2_pro = ObjExcel.Cells(DCEX_starting_excel_row + 9, current_excel_col).Value
		DCEX_child_number2_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 10, current_excel_col).Value, 1)
		DCEX_child_number3 = ObjExcel.Cells(DCEX_starting_excel_row + 11, current_excel_col).Value
		DCEX_child_number3_retro = ObjExcel.Cells(DCEX_starting_excel_row + 12, current_excel_col).Value
		DCEX_child_number3_pro = ObjExcel.Cells(DCEX_starting_excel_row + 13, current_excel_col).Value
		DCEX_child_number3_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 14, current_excel_col).Value, 1)
		DCEX_child_number4 = ObjExcel.Cells(DCEX_starting_excel_row + 15, current_excel_col).Value
		DCEX_child_number4_retro = ObjExcel.Cells(DCEX_starting_excel_row + 16, current_excel_col).Value
		DCEX_child_number4_pro = ObjExcel.Cells(DCEX_starting_excel_row + 17, current_excel_col).Value
		DCEX_child_number4_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 18, current_excel_col).Value, 1)
		DCEX_child_number5 = ObjExcel.Cells(DCEX_starting_excel_row + 19, current_excel_col).Value
		DCEX_child_number5_retro = ObjExcel.Cells(DCEX_starting_excel_row + 20, current_excel_col).Value
		DCEX_child_number5_pro = ObjExcel.Cells(DCEX_starting_excel_row + 21, current_excel_col).Value
		DCEX_child_number5_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 22, current_excel_col).Value, 1)
		DCEX_child_number6 = ObjExcel.Cells(DCEX_starting_excel_row + 23, current_excel_col).Value
		DCEX_child_number6_retro = ObjExcel.Cells(DCEX_starting_excel_row + 24, current_excel_col).Value
		DCEX_child_number6_pro = ObjExcel.Cells(DCEX_starting_excel_row + 25, current_excel_col).Value
		DCEX_child_number6_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 26, current_excel_col).Value, 1)

		DFLN_starting_excel_row = 237
		DFLN_conv_1_dt = ObjExcel.Cells(DFLN_starting_excel_row, current_excel_col).Value
		DFLN_conv_1_juris = ObjExcel.Cells(DFLN_starting_excel_row + 1, current_excel_col).Value
		DFLN_conv_1_state = ObjExcel.Cells(DFLN_starting_excel_row + 2, current_excel_col).Value
		DFLN_conv_2_dt = ObjExcel.Cells(DFLN_starting_excel_row + 3, current_excel_col).Value
		DFLN_conv_2_juris = ObjExcel.Cells(DFLN_starting_excel_row + 4, current_excel_col).Value
		DFLN_conv_2_state = ObjExcel.Cells(DFLN_starting_excel_row + 5, current_excel_col).Value
		DFLN_rnd_test_1_dt = ObjExcel.Cells(DFLN_starting_excel_row + 6, current_excel_col).Value
		DFLN_rnd_test_1_provider = ObjExcel.Cells(DFLN_starting_excel_row + 7, current_excel_col).Value
		DFLN_rnd_test_1_result = left(ObjExcel.Cells(DFLN_starting_excel_row + 8, current_excel_col).Value, 2)
		DFLN_rnd_test_2_dt = ObjExcel.Cells(DFLN_starting_excel_row + 9, current_excel_col).Value
		DFLN_rnd_test_2_provider = ObjExcel.Cells(DFLN_starting_excel_row + 10, current_excel_col).Value
		DFLN_rnd_test_2_result = left(ObjExcel.Cells(DFLN_starting_excel_row + 11, current_excel_col).Value, 2)

		DIET_starting_excel_row = 249
		DIET_mfip_1 = left(ObjExcel.Cells(DIET_starting_excel_row, current_excel_col).Value, 2)
		DIET_mfip_1_ver = ObjExcel.Cells(DIET_starting_excel_row + 1, current_excel_col).Value
		DIET_mfip_2 = left(ObjExcel.Cells(DIET_starting_excel_row + 2, current_excel_col).Value, 2)
		DIET_mfip_2_ver = ObjExcel.Cells(DIET_starting_excel_row + 3, current_excel_col).Value
		DIET_msa_1 = left(ObjExcel.Cells(DIET_starting_excel_row + 4, current_excel_col).Value, 2)
		DIET_msa_1_ver = ObjExcel.Cells(DIET_starting_excel_row + 5, current_excel_col).Value
		DIET_msa_2 = left(ObjExcel.Cells(DIET_starting_excel_row + 6, current_excel_col).Value, 2)
		DIET_msa_2_ver = ObjExcel.Cells(DIET_starting_excel_row + 7, current_excel_col).Value
		DIET_msa_3 = left(ObjExcel.Cells(DIET_starting_excel_row + 8, current_excel_col).Value, 2)
		DIET_msa_3_ver = ObjExcel.Cells(DIET_starting_excel_row + 9, current_excel_col).Value
		DIET_msa_4 = left(ObjExcel.Cells(DIET_starting_excel_row + 10, current_excel_col).Value, 2)
		DIET_msa_4_ver = ObjExcel.Cells(DIET_starting_excel_row + 11, current_excel_col).Value

		DISA_starting_excel_row = 261
		DISA_begin_date = ObjExcel.Cells(DISA_starting_excel_row, current_excel_col).Value
		DISA_end_date = ObjExcel.Cells(DISA_starting_excel_row + 1, current_excel_col).Value
		DISA_cert_begin = ObjExcel.Cells(DISA_starting_excel_row + 2, current_excel_col).Value
		DISA_cert_end = ObjExcel.Cells(DISA_starting_excel_row + 3, current_excel_col).Value
		DISA_wavr_begin = ObjExcel.Cells(DISA_starting_excel_row + 4, current_excel_col).Value
		DISA_wavr_end = ObjExcel.Cells(DISA_starting_excel_row + 5, current_excel_col).Value
		DISA_grh_begin = ObjExcel.Cells(DISA_starting_excel_row + 6, current_excel_col).Value
		DISA_grh_end = ObjExcel.Cells(DISA_starting_excel_row + 7, current_excel_col).Value
		DISA_cash_status = left(ObjExcel.Cells(DISA_starting_excel_row + 8, current_excel_col).Value, 2)
		DISA_cash_status_ver = left(ObjExcel.Cells(DISA_starting_excel_row + 9, current_excel_col).Value, 1)
		DISA_snap_status = left(ObjExcel.Cells(DISA_starting_excel_row + 10, current_excel_col).Value, 2)
		DISA_snap_status_ver = left(ObjExcel.Cells(DISA_starting_excel_row + 11, current_excel_col).Value, 1)
		DISA_hc_status = left(ObjExcel.Cells(DISA_starting_excel_row + 12, current_excel_col).Value, 2)
		DISA_hc_status_ver = left(ObjExcel.Cells(DISA_starting_excel_row + 13, current_excel_col).Value, 1)
        'This is variable because we have added a row to the template but some scenarios will not have this row added yet.
        for add_row = 14 to 16
            If trim(ObjExcel.Cells(DISA_starting_excel_row + add_row, 2).Value) = "home/community based waiver" Then
                DISA_waiver = left(ObjExcel.Cells(DISA_starting_excel_row + add_row, current_excel_col).Value, 1)
            ElseIf trim(ObjExcel.Cells(DISA_starting_excel_row + add_row, 2).Value) = "1619 status" Then
                DISA_1619 = left(ObjExcel.Cells(DISA_starting_excel_row + add_row, current_excel_col).Value, 1)
            ElseIf trim(ObjExcel.Cells(DISA_starting_excel_row + add_row, 2).Value) = "drug/alcoholism ver code" Then
                DISA_drug_alcohol = left(ObjExcel.Cells(DISA_starting_excel_row + add_row, current_excel_col).Value, 1)
            End If

            If trim(ObjExcel.Cells(DISA_starting_excel_row + add_row, 2).Value) = "drug/alcoholism ver code" Then
                starting_row = DISA_starting_excel_row + add_row + 1
            End If
        next

		DSTT_starting_excel_row = starting_row        '277
		DSTT_ongoing_income = ObjExcel.Cells(DSTT_starting_excel_row, current_excel_col).Value
		DSTT_HH_income_stop_date = ObjExcel.Cells(DSTT_starting_excel_row + 1, current_excel_col).Value
		DSTT_income_expected_amt = ObjExcel.Cells(DSTT_starting_excel_row + 2, current_excel_col).Value
        starting_row = starting_row + 3

		EATS_starting_excel_row = starting_row        '280
		EATS_together = ObjExcel.Cells(EATS_starting_excel_row, current_excel_col).Value
		EATS_boarder = ObjExcel.Cells(EATS_starting_excel_row + 1, current_excel_col).Value
		EATS_group_one = ObjExcel.Cells(EATS_starting_excel_row + 2, current_excel_col).Value
		EATS_group_two = ObjExcel.Cells(EATS_starting_excel_row + 3, current_excel_col).Value
		EATS_group_three = ObjExcel.Cells(EATS_starting_excel_row + 4, current_excel_col).Value
        starting_row = starting_row + 5

		EMMA_starting_excel_row = starting_row        '285
		EMMA_medical_emergency = left(ObjExcel.Cells(EMMA_starting_excel_row, current_excel_col).Value, 2)
		EMMA_health_consequence = left(ObjExcel.Cells(EMMA_starting_excel_row + 1, current_excel_col).Value, 2)
		EMMA_verification = left(ObjExcel.Cells(EMMA_starting_excel_row + 2, current_excel_col).Value, 2)
		EMMA_begin_date = ObjExcel.Cells(EMMA_starting_excel_row + 3, current_excel_col).Value
		EMMA_end_date = ObjExcel.Cells(EMMA_starting_excel_row + 4, current_excel_col).Value
        starting_row = starting_row + 5

		EMPS_starting_excel_row = starting_row        '290
		EMPS_orientation_date = ObjExcel.Cells(EMPS_starting_excel_row, current_excel_col).Value
		EMPS_orientation_attended = ObjExcel.Cells(EMPS_starting_excel_row + 1, current_excel_col).Value
		EMPS_good_cause = left(ObjExcel.Cells(EMPS_starting_excel_row + 2, current_excel_col).Value, 2)
		EMPS_sanc_begin = ObjExcel.Cells(EMPS_starting_excel_row + 3, current_excel_col).Value
		EMPS_sanc_end = ObjExcel.Cells(EMPS_starting_excel_row + 4, current_excel_col).Value
		EMPS_memb_at_home = ObjExcel.Cells(EMPS_starting_excel_row + 5, current_excel_col).Value
		EMPS_care_family = ObjExcel.Cells(EMPS_starting_excel_row + 6, current_excel_col).Value
		EMPS_crisis = ObjExcel.Cells(EMPS_starting_excel_row + 7, current_excel_col).Value
		EMPS_hard_employ = left(ObjExcel.Cells(EMPS_starting_excel_row + 8, current_excel_col).Value, 2)
		EMPS_under1 = ObjExcel.Cells(EMPS_starting_excel_row + 9, current_excel_col).Value
		EMPS_DWP_date = ObjExcel.Cells(EMPS_starting_excel_row + 10, current_excel_col).Value
        starting_row = starting_row + 11

		FACI_starting_excel_row = starting_row        '301
		FACI_vendor_number = ObjExcel.Cells(FACI_starting_excel_row, current_excel_col).Value
		FACI_name = ObjExcel.Cells(FACI_starting_excel_row + 1, current_excel_col).Value
		FACI_type = left(ObjExcel.Cells(FACI_starting_excel_row + 2, current_excel_col).Value, 2)
		FACI_FS_eligible = ObjExcel.Cells(FACI_starting_excel_row + 3, current_excel_col).Value
		FACI_FS_facility_type = left(ObjExcel.Cells(FACI_starting_excel_row + 4, current_excel_col).Value, 1)
		FACI_date_in = ObjExcel.Cells(FACI_starting_excel_row + 5, current_excel_col).Value
		FACI_date_out = ObjExcel.Cells(FACI_starting_excel_row + 6, current_excel_col).Value
        starting_row = starting_row + 7

		FMED_starting_excel_row = starting_row        '308
		FMED_medical_mileage = objExcel.Cells(FMED_starting_excel_row, current_excel_col).Value
		FMED_1_type = left(objExcel.Cells(FMED_starting_excel_row + 1, current_excel_col).Value, 2)
		FMED_1_verif = left(objExcel.Cells(FMED_starting_excel_row + 2, current_excel_col).Value, 2)
		FMED_1_ref_num = objExcel.Cells(FMED_starting_excel_row + 3, current_excel_col).Value
		FMED_1_category = left(objExcel.Cells(FMED_starting_excel_row + 4, current_excel_col).Value, 1)
		FMED_1_begin = objExcel.Cells(FMED_starting_excel_row + 5, current_excel_col).Value
		FMED_1_end = objExcel.Cells(FMED_starting_excel_row + 6, current_excel_col).Value
		FMED_1_amount = objExcel.Cells(FMED_starting_excel_row + 7, current_excel_col).Value
		FMED_2_type = left(objExcel.Cells(FMED_starting_excel_row + 8, current_excel_col).Value, 2)
		FMED_2_verif = left(objExcel.Cells(FMED_starting_excel_row + 9, current_excel_col).Value, 2)
		FMED_2_ref_num = objExcel.Cells(FMED_starting_excel_row + 10, current_excel_col).Value
		FMED_2_category = left(objExcel.Cells(FMED_starting_excel_row + 11, current_excel_col).Value, 1)
		FMED_2_begin = objExcel.Cells(FMED_starting_excel_row + 12, current_excel_col).Value
		FMED_2_end = objExcel.Cells(FMED_starting_excel_row + 13, current_excel_col).Value
		FMED_2_amount = objExcel.Cells(FMED_starting_excel_row + 14, current_excel_col).Value
		FMED_3_type = left(objExcel.Cells(FMED_starting_excel_row + 15, current_excel_col).Value, 2)
		FMED_3_verif = left(objExcel.Cells(FMED_starting_excel_row + 16, current_excel_col).Value, 2)
		FMED_3_ref_num = objExcel.Cells(FMED_starting_excel_row + 17, current_excel_col).Value
		FMED_3_category = left(objExcel.Cells(FMED_starting_excel_row + 18, current_excel_col).Value, 1)
		FMED_3_begin = objExcel.Cells(FMED_starting_excel_row + 19, current_excel_col).Value
		FMED_3_end = objExcel.Cells(FMED_starting_excel_row + 20, current_excel_col).Value
		FMED_3_amount = objExcel.Cells(FMED_starting_excel_row + 21, current_excel_col).Value
		FMED_4_type = left(objExcel.Cells(FMED_starting_excel_row + 22, current_excel_col).Value, 2)
		FMED_4_verif = left(objExcel.Cells(FMED_starting_excel_row + 23, current_excel_col).Value, 2)
		FMED_4_ref_num = objExcel.Cells(FMED_starting_excel_row + 24, current_excel_col).Value
		FMED_4_category = left(objExcel.Cells(FMED_starting_excel_row + 25, current_excel_col).Value, 1)
		FMED_4_begin = objExcel.Cells(FMED_starting_excel_row + 26, current_excel_col).Value
		FMED_4_end = objExcel.Cells(FMED_starting_excel_row + 27, current_excel_col).Value
		FMED_4_amount = objExcel.Cells(FMED_starting_excel_row + 28, current_excel_col).Value
        starting_row = starting_row + 29

		HEST_starting_excel_row = starting_row        '337
		HEST_FS_choice_date = ObjExcel.Cells(HEST_starting_excel_row, current_excel_col).Value
		HEST_first_month = ObjExcel.Cells(HEST_starting_excel_row + 1, current_excel_col).Value
		HEST_heat_air_retro = ObjExcel.Cells(HEST_starting_excel_row + 2, current_excel_col).Value
		HEST_heat_air_pro = ObjExcel.Cells(HEST_starting_excel_row + 3, current_excel_col).Value
		HEST_electric_retro = ObjExcel.Cells(HEST_starting_excel_row + 4, current_excel_col).Value
		HEST_electric_pro = ObjExcel.Cells(HEST_starting_excel_row + 5, current_excel_col).Value
		HEST_phone_retro = ObjExcel.Cells(HEST_starting_excel_row + 6, current_excel_col).Value
		HEST_phone_pro = ObjExcel.Cells(HEST_starting_excel_row + 7, current_excel_col).Value
        starting_row = starting_row + 8

		IMIG_starting_excel_row = starting_row        '345
		IMIG_imigration_status = left(ObjExcel.Cells(IMIG_starting_excel_row, current_excel_col).Value, 2)
		IMIG_entry_date = ObjExcel.Cells(IMIG_starting_excel_row + 1, current_excel_col).Value
		IMIG_status_date = ObjExcel.Cells(IMIG_starting_excel_row + 2, current_excel_col).Value
		IMIG_status_ver = left(ObjExcel.Cells(IMIG_starting_excel_row + 3, current_excel_col).Value, 2)
		IMIG_status_LPR_adj_from = left(ObjExcel.Cells(IMIG_starting_excel_row + 4, current_excel_col).Value, 2)
		IMIG_nationality = left(ObjExcel.Cells(IMIG_starting_excel_row + 5, current_excel_col).Value, 2)
		IMIG_40_soc_sec = ObjExcel.Cells(IMIG_starting_excel_row + 6, current_excel_col).Value
		IMIG_40_soc_sec_verif = ObjExcel.Cells(IMIG_starting_excel_row + 7, current_excel_col).Value
		IMIG_battered_spouse_child = ObjExcel.Cells(IMIG_starting_excel_row + 8, current_excel_col).Value
		IMIG_battered_spouse_child_verif = ObjExcel.Cells(IMIG_starting_excel_row + 9, current_excel_col).Value
		IMIG_military_status = left(ObjExcel.Cells(IMIG_starting_excel_row + 10, current_excel_col).Value, 1)
		IMIG_military_status_verif = ObjExcel.Cells(IMIG_starting_excel_row + 11, current_excel_col).Value
		IMIG_hmong_lao_nat_amer = left(ObjExcel.Cells(IMIG_starting_excel_row + 12, current_excel_col).Value, 2)
		IMIG_st_prog_esl_ctzn_coop = ObjExcel.Cells(IMIG_starting_excel_row + 13, current_excel_col).Value
		IMIG_st_prog_esl_ctzn_coop_verif = ObjExcel.Cells(IMIG_starting_excel_row + 14, current_excel_col).Value
		IMIG_fss_esl_skills_training = ObjExcel.Cells(IMIG_starting_excel_row + 15, current_excel_col).Value
        starting_row = starting_row + 16

		INSA_starting_excel_row = starting_row        '361
		INSA_pers_coop_ohi = ObjExcel.Cells(INSA_starting_excel_row, current_excel_col).Value
		INSA_good_cause_status = left(ObjExcel.Cells(INSA_starting_excel_row + 1, current_excel_col).Value, 1)
		INSA_good_cause_cliam_date = ObjExcel.Cells(INSA_starting_excel_row + 2, current_excel_col).Value
		INSA_good_cause_evidence = ObjExcel.Cells(INSA_starting_excel_row + 3, current_excel_col).Value
		INSA_coop_cost_effect = ObjExcel.Cells(INSA_starting_excel_row + 4, current_excel_col).Value
		INSA_insur_name = ObjExcel.Cells(INSA_starting_excel_row + 5, current_excel_col).Value
		INSA_prescrip_drug_cover = ObjExcel.Cells(INSA_starting_excel_row + 6, current_excel_col).Value
		INSA_prescrip_end_date = ObjExcel.Cells(INSA_starting_excel_row + 7, current_excel_col).Value
		INSA_persons_covered = ObjExcel.Cells(INSA_starting_excel_row + 8, current_excel_col).Value
        starting_row = starting_row + 9

		JOBS_1_starting_excel_row = starting_row      '370
		JOBS_1_inc_type = left(ObjExcel.Cells(JOBS_1_starting_excel_row, current_excel_col).Value, 1)
		JOBS_1_inc_verif = left(ObjExcel.Cells(JOBS_1_starting_excel_row + 1, current_excel_col).Value, 1)
		JOBS_1_employer_name = ObjExcel.Cells(JOBS_1_starting_excel_row + 2, current_excel_col).Value
		JOBS_1_inc_start = ObjExcel.Cells(JOBS_1_starting_excel_row + 3, current_excel_col).Value
		JOBS_1_pay_freq = ObjExcel.Cells(JOBS_1_starting_excel_row + 4, current_excel_col).Value
		JOBS_1_wkly_hrs = ObjExcel.Cells(JOBS_1_starting_excel_row + 5, current_excel_col).Value
		JOBS_1_hrly_wage = ObjExcel.Cells(JOBS_1_starting_excel_row + 6, current_excel_col).Value
        starting_row = starting_row + 7

		JOBS_2_starting_excel_row = starting_row      '377
		JOBS_2_inc_type = left(ObjExcel.Cells(JOBS_2_starting_excel_row, current_excel_col).Value, 1)
		JOBS_2_inc_verif = left(ObjExcel.Cells(JOBS_2_starting_excel_row + 1, current_excel_col).Value, 1)
		JOBS_2_employer_name = ObjExcel.Cells(JOBS_2_starting_excel_row + 2, current_excel_col).Value
		JOBS_2_inc_start = ObjExcel.Cells(JOBS_2_starting_excel_row + 3, current_excel_col).Value
		JOBS_2_pay_freq = ObjExcel.Cells(JOBS_2_starting_excel_row + 4, current_excel_col).Value
		JOBS_2_wkly_hrs = ObjExcel.Cells(JOBS_2_starting_excel_row + 5, current_excel_col).Value
		JOBS_2_hrly_wage = ObjExcel.Cells(JOBS_2_starting_excel_row + 6, current_excel_col).Value
        starting_row = starting_row + 7

		JOBS_3_starting_excel_row = starting_row      '384
		JOBS_3_inc_type = left(ObjExcel.Cells(JOBS_3_starting_excel_row, current_excel_col).Value, 1)
		JOBS_3_inc_verif = left(ObjExcel.Cells(JOBS_3_starting_excel_row + 1, current_excel_col).Value, 1)
		JOBS_3_employer_name = ObjExcel.Cells(JOBS_3_starting_excel_row + 2, current_excel_col).Value
		JOBS_3_inc_start = ObjExcel.Cells(JOBS_3_starting_excel_row + 3, current_excel_col).Value
		JOBS_3_pay_freq = ObjExcel.Cells(JOBS_3_starting_excel_row + 4, current_excel_col).Value
		JOBS_3_wkly_hrs = ObjExcel.Cells(JOBS_3_starting_excel_row + 5, current_excel_col).Value
		JOBS_3_hrly_wage = ObjExcel.Cells(JOBS_3_starting_excel_row + 6, current_excel_col).Value
        starting_row = starting_row + 7

		MEDI_starting_excel_row = starting_row        '391
		MEDI_claim_number_suffix = ObjExcel.Cells(MEDI_starting_excel_row, current_excel_col).Value
		MEDI_part_A_premium = ObjExcel.Cells(MEDI_starting_excel_row + 1, current_excel_col).Value
		MEDI_part_B_premium = ObjExcel.Cells(MEDI_starting_excel_row + 2, current_excel_col).Value
		MEDI_part_A_begin_date = ObjExcel.Cells(MEDI_starting_excel_row + 3, current_excel_col).Value
		MEDI_part_B_begin_date = ObjExcel.Cells(MEDI_starting_excel_row + 4, current_excel_col).Value
		MEDI_apply_prem_to_spdn = ObjExcel.Cells(MEDI_starting_excel_row + 5, current_excel_col).Value
		MEDI_apply_prem_end_date = ObjExcel.Cells(MEDI_starting_excel_row + 6, current_excel_col).Value
        starting_row = starting_row + 7

		MMSA_starting_excel_row = starting_row        '398
		MMSA_liv_arr = left(ObjExcel.Cells(MMSA_starting_excel_row, current_excel_col).Value, 1)
		MMSA_cont_elig = ObjExcel.Cells(MMSA_starting_excel_row + 1, current_excel_col).Value
		MMSA_spous_inc = ObjExcel.Cells(MMSA_starting_excel_row + 2, current_excel_col).Value
		MMSA_shared_hous = ObjExcel.Cells(MMSA_starting_excel_row + 3, current_excel_col).Value
        starting_row = starting_row + 4

		MSUR_starting_excel_row = starting_row        '402
		MSUR_begin_date = ObjExcel.Cells(MSUR_starting_excel_row, current_excel_col).Value
        starting_row = starting_row + 1

		OTHR_starting_excel_row = starting_row        '403
		OTHR_type = left(ObjExcel.Cells(OTHR_starting_excel_row, current_excel_col).Value, 1)
		OTHR_cash_value = ObjExcel.Cells(OTHR_starting_excel_row + 1, current_excel_col).Value
		OTHR_cash_value_ver = left(ObjExcel.Cells(OTHR_starting_excel_row + 2, current_excel_col).Value, 1)
		OTHR_owed = ObjExcel.Cells(OTHR_starting_excel_row + 3, current_excel_col).Value
		OTHR_owed_ver = left(ObjExcel.Cells(OTHR_starting_excel_row + 4, current_excel_col).Value, 1)
		OTHR_date = ObjExcel.Cells(OTHR_starting_excel_row + 5, current_excel_col).Value
		OTHR_cash_count = ObjExcel.Cells(OTHR_starting_excel_row + 6, current_excel_col).Value
		OTHR_SNAP_count = ObjExcel.Cells(OTHR_starting_excel_row + 7, current_excel_col).Value
		OTHR_HC_count = ObjExcel.Cells(OTHR_starting_excel_row + 8, current_excel_col).Value
		OTHR_IV_count = ObjExcel.Cells(OTHR_starting_excel_row + 9, current_excel_col).Value
		OTHR_joint = ObjExcel.Cells(OTHR_starting_excel_row + 10, current_excel_col).Value
		OTHR_share_ratio = ObjExcel.Cells(OTHR_starting_excel_row + 11, current_excel_col).Value
        starting_row = starting_row + 12

		PARE_starting_excel_row = starting_row        '415
		PARE_child_1 = ObjExcel.Cells(PARE_starting_excel_row, current_excel_col).Value
		PARE_child_1_relation = left(ObjExcel.Cells(PARE_starting_excel_row + 1, current_excel_col).Value, 1)
		PARE_child_1_verif = left(ObjExcel.Cells(PARE_starting_excel_row + 2, current_excel_col).Value, 2)
		PARE_child_2 = ObjExcel.Cells(PARE_starting_excel_row + 3, current_excel_col).Value
		PARE_child_2_relation = left(ObjExcel.Cells(PARE_starting_excel_row + 4, current_excel_col).Value, 1)
		PARE_child_2_verif = left(ObjExcel.Cells(PARE_starting_excel_row + 5, current_excel_col).Value, 2)
		PARE_child_3 = ObjExcel.Cells(PARE_starting_excel_row + 6, current_excel_col).Value
		PARE_child_3_relation = left(ObjExcel.Cells(PARE_starting_excel_row + 7, current_excel_col).Value, 1)
		PARE_child_3_verif = left(ObjExcel.Cells(PARE_starting_excel_row + 8, current_excel_col).Value, 2)
		PARE_child_4 = ObjExcel.Cells(PARE_starting_excel_row + 9, current_excel_col).Value
		PARE_child_4_relation = left(ObjExcel.Cells(PARE_starting_excel_row + 10, current_excel_col).Value, 1)
		PARE_child_4_verif = left(ObjExcel.Cells(PARE_starting_excel_row + 11, current_excel_col).Value, 2)
		PARE_child_5 = ObjExcel.Cells(PARE_starting_excel_row + 12, current_excel_col).Value
		PARE_child_5_relation = left(ObjExcel.Cells(PARE_starting_excel_row + 13, current_excel_col).Value, 1)
		PARE_child_5_verif = left(ObjExcel.Cells(PARE_starting_excel_row + 14, current_excel_col).Value, 2)
		PARE_child_6 = ObjExcel.Cells(PARE_starting_excel_row + 15, current_excel_col).Value
		PARE_child_6_relation = left(ObjExcel.Cells(PARE_starting_excel_row + 16, current_excel_col).Value, 1)
		PARE_child_6_verif = left(ObjExcel.Cells(PARE_starting_excel_row + 17, current_excel_col).Value, 2)
        starting_row = starting_row + 18

		PBEN_1_starting_excel_row = starting_row      '433
		PBEN_1_referal_date = ObjExcel.Cells(PBEN_1_starting_excel_row, current_excel_col).Value
		PBEN_1_type = left(ObjExcel.Cells(PBEN_1_starting_excel_row + 1, current_excel_col).Value, 2)
		PBEN_1_appl_date = ObjExcel.Cells(PBEN_1_starting_excel_row + 2, current_excel_col).Value
		PBEN_1_appl_ver = left(ObjExcel.Cells(PBEN_1_starting_excel_row + 3, current_excel_col).Value, 1)
		PBEN_1_IAA_date = ObjExcel.Cells(PBEN_1_starting_excel_row + 4, current_excel_col).Value
		PBEN_1_disp = left(ObjExcel.Cells(PBEN_1_starting_excel_row + 5, current_excel_col).Value, 1)
        starting_row = starting_row + 6

		PBEN_2_starting_excel_row = starting_row      '439
		PBEN_2_referal_date = ObjExcel.Cells(PBEN_2_starting_excel_row, current_excel_col).Value
		PBEN_2_type = left(ObjExcel.Cells(PBEN_2_starting_excel_row + 1, current_excel_col).Value, 2)
		PBEN_2_appl_date = ObjExcel.Cells(PBEN_2_starting_excel_row + 2, current_excel_col).Value
		PBEN_2_appl_ver = left(ObjExcel.Cells(PBEN_2_starting_excel_row + 3, current_excel_col).Value, 1)
		PBEN_2_IAA_date = ObjExcel.Cells(PBEN_2_starting_excel_row + 4, current_excel_col).Value
		PBEN_2_disp = left(ObjExcel.Cells(PBEN_2_starting_excel_row + 5, current_excel_col).Value, 1)
        starting_row = starting_row + 6

		PBEN_3_starting_excel_row = starting_row      '445
		PBEN_3_referal_date = ObjExcel.Cells(PBEN_3_starting_excel_row, current_excel_col).Value
		PBEN_3_type = left(ObjExcel.Cells(PBEN_3_starting_excel_row + 1, current_excel_col).Value, 2)
		PBEN_3_appl_date = ObjExcel.Cells(PBEN_3_starting_excel_row + 2, current_excel_col).Value
		PBEN_3_appl_ver = left(ObjExcel.Cells(PBEN_3_starting_excel_row + 3, current_excel_col).Value, 1)
		PBEN_3_IAA_date = ObjExcel.Cells(PBEN_3_starting_excel_row + 4, current_excel_col).Value
		PBEN_3_disp = left(ObjExcel.Cells(PBEN_3_starting_excel_row + 5, current_excel_col).Value, 1)
        starting_row = starting_row + 6

		PDED_starting_excel_row = starting_row        '451
		PDED_wid_deduction = ObjExcel.Cells(PDED_starting_excel_row, current_excel_col).Value
		PDED_adult_child_disregard = ObjExcel.Cells(PDED_starting_excel_row + 1, current_excel_col).Value
		PDED_wid_disregard = ObjExcel.Cells(PDED_starting_excel_row + 2, current_excel_col).Value
		PDED_unea_income_deduction_reason = ObjExcel.Cells(PDED_starting_excel_row + 3, current_excel_col).Value
		PDED_unea_income_deduction_value = ObjExcel.Cells(PDED_starting_excel_row + 4, current_excel_col).Value
		PDED_earned_income_deduction_reason = ObjExcel.Cells(PDED_starting_excel_row + 5, current_excel_col).Value
		PDED_earned_income_deduction_value = ObjExcel.Cells(PDED_starting_excel_row + 6, current_excel_col).Value
		PDED_ma_epd_inc_asset_limit = ObjExcel.Cells(PDED_starting_excel_row + 7, current_excel_col).Value
		PDED_guard_fee = ObjExcel.Cells(PDED_starting_excel_row + 8, current_excel_col).Value
		PDED_rep_payee_fee = ObjExcel.Cells(PDED_starting_excel_row + 9, current_excel_col).Value
		PDED_other_expense = ObjExcel.Cells(PDED_starting_excel_row + 10, current_excel_col).Value
		PDED_shel_spcl_needs = ObjExcel.Cells(PDED_starting_excel_row + 11, current_excel_col).Value
		PDED_excess_need = ObjExcel.Cells(PDED_starting_excel_row + 12, current_excel_col).Value
		PDED_restaurant_meals = ObjExcel.Cells(PDED_starting_excel_row + 13, current_excel_col).Value
        starting_row = starting_row + 14

		PREG_starting_excel_row = starting_row        '465
		PREG_conception_date = ObjExcel.Cells(PREG_starting_excel_row, current_excel_col).Value
		PREG_conception_date_ver = ObjExcel.Cells(PREG_starting_excel_row + 1, current_excel_col).Value
		PREG_third_trimester_ver = ObjExcel.Cells(PREG_starting_excel_row + 2, current_excel_col).Value
		PREG_due_date = ObjExcel.Cells(PREG_starting_excel_row + 3, current_excel_col).Value
		PREG_multiple_birth = ObjExcel.Cells(PREG_starting_excel_row + 4, current_excel_col).Value
        starting_row = starting_row + 5

		RBIC_starting_excel_row = starting_row        '470
		RBIC_type = left(ObjExcel.Cells(RBIC_starting_excel_row, current_excel_col).Value, 2)
		RBIC_start_date = ObjExcel.Cells(RBIC_starting_excel_row + 1, current_excel_col).Value
		RBIC_end_date = ObjExcel.Cells(RBIC_starting_excel_row + 2, current_excel_col).Value
		RBIC_group_1 = ObjExcel.Cells(RBIC_starting_excel_row + 3, current_excel_col).Value
		RBIC_retro_income_group_1 = ObjExcel.Cells(RBIC_starting_excel_row + 4, current_excel_col).Value
		RBIC_prosp_income_group_1 = ObjExcel.Cells(RBIC_starting_excel_row + 5, current_excel_col).Value
		RBIC_ver_income_group_1 = left(ObjExcel.Cells(RBIC_starting_excel_row + 6, current_excel_col).Value, 1)
		RBIC_group_2 = ObjExcel.Cells(RBIC_starting_excel_row + 7, current_excel_col).Value
		RBIC_retro_income_group_2 = ObjExcel.Cells(RBIC_starting_excel_row + 8, current_excel_col).Value
		RBIC_prosp_income_group_2 = ObjExcel.Cells(RBIC_starting_excel_row + 9, current_excel_col).Value
		RBIC_ver_income_group_2 = left(ObjExcel.Cells(RBIC_starting_excel_row + 10, current_excel_col).Value, 1)
		RBIC_group_3 = ObjExcel.Cells(RBIC_starting_excel_row + 11, current_excel_col).Value
		RBIC_retro_income_group_3 = ObjExcel.Cells(RBIC_starting_excel_row + 12, current_excel_col).Value
		RBIC_prosp_income_group_3 = ObjExcel.Cells(RBIC_starting_excel_row + 13, current_excel_col).Value
		RBIC_ver_income_group_3 = left(ObjExcel.Cells(RBIC_starting_excel_row + 14, current_excel_col).Value, 1)
		RBIC_retro_hours = ObjExcel.Cells(RBIC_starting_excel_row + 15, current_excel_col).Value
		RBIC_prosp_hours = ObjExcel.Cells(RBIC_starting_excel_row + 16, current_excel_col).Value
		RBIC_exp_type_1 = left(ObjExcel.Cells(RBIC_starting_excel_row + 17, current_excel_col).Value, 2)
		RBIC_exp_retro_1 = ObjExcel.Cells(RBIC_starting_excel_row + 18, current_excel_col).Value
		RBIC_exp_prosp_1 = ObjExcel.Cells(RBIC_starting_excel_row + 19, current_excel_col).Value
		RBIC_exp_ver_1 = left(ObjExcel.Cells(RBIC_starting_excel_row + 20, current_excel_col).Value, 1)
		RBIC_exp_type_2 = left(ObjExcel.Cells(RBIC_starting_excel_row + 21, current_excel_col).Value, 2)
		RBIC_exp_retro_2 = ObjExcel.Cells(RBIC_starting_excel_row + 22, current_excel_col).Value
		RBIC_exp_prosp_2 = ObjExcel.Cells(RBIC_starting_excel_row + 23, current_excel_col).Value
		RBIC_exp_ver_2 = left(ObjExcel.Cells(RBIC_starting_excel_row + 24, current_excel_col).Value, 1)
        starting_row = starting_row + 25

		REST_starting_excel_row = starting_row        '495
		REST_type = left(ObjExcel.Cells(REST_starting_excel_row, current_excel_col).Value, 1)
		REST_type_ver = left(ObjExcel.Cells(REST_starting_excel_row + 1, current_excel_col).Value, 2)
		REST_market = ObjExcel.Cells(REST_starting_excel_row + 2, current_excel_col).Value
		REST_market_ver = left(ObjExcel.Cells(REST_starting_excel_row + 3, current_excel_col).Value, 2)
		REST_owed = ObjExcel.Cells(REST_starting_excel_row + 4, current_excel_col).Value
		REST_owed_ver = left(ObjExcel.Cells(REST_starting_excel_row + 5, current_excel_col).Value, 2)
		REST_date = ObjExcel.Cells(REST_starting_excel_row + 6, current_excel_col).Value
		REST_status = left(ObjExcel.Cells(REST_starting_excel_row + 7, current_excel_col).Value, 1)
		REST_joint = ObjExcel.Cells(REST_starting_excel_row + 8, current_excel_col).Value
		REST_share_ratio = ObjExcel.Cells(REST_starting_excel_row + 9, current_excel_col).Value
		REST_agreement_date = ObjExcel.Cells(REST_starting_excel_row + 10, current_excel_col).Value
        starting_row = starting_row + 11

		SCHL_starting_excel_row = starting_row        '506
		SCHL_status = left(ObjExcel.Cells(SCHL_starting_excel_row, current_excel_col).Value, 1)
		SCHL_ver = left(ObjExcel.Cells(SCHL_starting_excel_row + 1, current_excel_col).Value, 2)
		SCHL_type = left(ObjExcel.Cells(SCHL_starting_excel_row + 2, current_excel_col).Value, 2)
		SCHL_district_nbr = ObjExcel.Cells(SCHL_starting_excel_row + 3, current_excel_col).Value
		SCHL_kindergarten_start_date = ObjExcel.Cells(SCHL_starting_excel_row + 4, current_excel_col).Value
		SCHL_grad_date = ObjExcel.Cells(SCHL_starting_excel_row + 5, current_excel_col).Value
		SCHL_grad_date_ver = left(ObjExcel.Cells(SCHL_starting_excel_row + 6, current_excel_col).Value, 2)
		SCHL_primary_secondary_funding = left(ObjExcel.Cells(SCHL_starting_excel_row + 7, current_excel_col).Value, 1)
		SCHL_FS_eligibility_status = left(ObjExcel.Cells(SCHL_starting_excel_row + 8, current_excel_col).Value, 2)
		SCHL_higher_ed = ObjExcel.Cells(SCHL_starting_excel_row + 9, current_excel_col).Value
        starting_row = starting_row + 10

		SECU_starting_excel_row = starting_row        '516
		SECU_type = left(ObjExcel.Cells(SECU_starting_excel_row, current_excel_col).Value, 2)
		SECU_pol_numb = ObjExcel.Cells(SECU_starting_excel_row + 1, current_excel_col).Value
		SECU_name = ObjExcel.Cells(SECU_starting_excel_row + 2, current_excel_col).Value
		SECU_cash_val = ObjExcel.Cells(SECU_starting_excel_row + 3, current_excel_col).Value
		SECU_date = ObjExcel.Cells(SECU_starting_excel_row + 4, current_excel_col).Value
		SECU_cash_ver = left(ObjExcel.Cells(SECU_starting_excel_row + 5, current_excel_col).Value, 1)
		SECU_face_val = ObjExcel.Cells(SECU_starting_excel_row + 6, current_excel_col).Value
		SECU_withdraw = ObjExcel.Cells(SECU_starting_excel_row + 7, current_excel_col).Value
		SECU_cash_count = ObjExcel.Cells(SECU_starting_excel_row + 8, current_excel_col).Value
		SECU_SNAP_count = ObjExcel.Cells(SECU_starting_excel_row + 9, current_excel_col).Value
		SECU_HC_count = ObjExcel.Cells(SECU_starting_excel_row + 10, current_excel_col).Value
		SECU_GRH_count = ObjExcel.Cells(SECU_starting_excel_row + 11, current_excel_col).Value
		SECU_IV_count = ObjExcel.Cells(SECU_starting_excel_row + 12, current_excel_col).Value
		SECU_joint = ObjExcel.Cells(SECU_starting_excel_row + 13, current_excel_col).Value
		SECU_share_ratio = ObjExcel.Cells(SECU_starting_excel_row + 14, current_excel_col).Value
        starting_row = starting_row + 15

		SHEL_starting_excel_row = starting_row        '531
		SHEL_subsidized = ObjExcel.Cells(SHEL_starting_excel_row, current_excel_col).Value
		SHEL_shared = ObjExcel.Cells(SHEL_starting_excel_row + 1, current_excel_col).Value
		SHEL_paid_to = ObjExcel.Cells(SHEL_starting_excel_row + 2, current_excel_col).Value
		SHEL_rent_retro = ObjExcel.Cells(SHEL_starting_excel_row + 3, current_excel_col).Value
		SHEL_rent_retro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 4, current_excel_col).Value, 2)
		SHEL_rent_pro = ObjExcel.Cells(SHEL_starting_excel_row + 5, current_excel_col).Value
		SHEL_rent_pro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 6, current_excel_col).Value, 2)
		SHEL_lot_rent_retro = ObjExcel.Cells(SHEL_starting_excel_row + 7, current_excel_col).Value
		SHEL_lot_rent_retro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 8, current_excel_col).Value, 2)
		SHEL_lot_rent_pro = ObjExcel.Cells(SHEL_starting_excel_row + 9, current_excel_col).Value
		SHEL_lot_rent_pro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 10, current_excel_col).Value, 2)
		SHEL_mortgage_retro = ObjExcel.Cells(SHEL_starting_excel_row + 11, current_excel_col).Value
		SHEL_mortgage_retro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 12, current_excel_col).Value, 2)
		SHEL_mortgage_pro = ObjExcel.Cells(SHEL_starting_excel_row + 13, current_excel_col).Value
		SHEL_mortgage_pro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 14, current_excel_col).Value, 2)
		SHEL_insur_retro = ObjExcel.Cells(SHEL_starting_excel_row + 15, current_excel_col).Value
		SHEL_insur_retro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 16, current_excel_col).Value, 2)
		SHEL_insur_pro = ObjExcel.Cells(SHEL_starting_excel_row + 17, current_excel_col).Value
		SHEL_insur_pro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 18, current_excel_col).Value, 2)
		SHEL_taxes_retro = ObjExcel.Cells(SHEL_starting_excel_row + 19, current_excel_col).Value
		SHEL_taxes_retro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 20, current_excel_col).Value, 2)
		SHEL_taxes_pro = ObjExcel.Cells(SHEL_starting_excel_row + 21, current_excel_col).Value
		SHEL_taxes_pro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 22, current_excel_col).Value, 2)
		SHEL_room_retro = ObjExcel.Cells(SHEL_starting_excel_row + 23, current_excel_col).Value
		SHEL_room_retro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 24, current_excel_col).Value, 2)
		SHEL_room_pro = ObjExcel.Cells(SHEL_starting_excel_row + 25, current_excel_col).Value
		SHEL_room_pro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 26, current_excel_col).Value, 2)
		SHEL_garage_retro = ObjExcel.Cells(SHEL_starting_excel_row + 27, current_excel_col).Value
		SHEL_garage_retro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 28, current_excel_col).Value, 2)
		SHEL_garage_pro = ObjExcel.Cells(SHEL_starting_excel_row + 29, current_excel_col).Value
		SHEL_garage_pro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 30, current_excel_col).Value, 2)
		SHEL_subsidy_retro = ObjExcel.Cells(SHEL_starting_excel_row + 31, current_excel_col).Value
		SHEL_subsidy_retro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 32, current_excel_col).Value, 2)
		SHEL_subsidy_pro = ObjExcel.Cells(SHEL_starting_excel_row + 33, current_excel_col).Value
		SHEL_subsidy_pro_ver = left(ObjExcel.Cells(SHEL_starting_excel_row + 34, current_excel_col).Value, 2)
        starting_row = starting_row + 35

		SIBL_starting_excel_row = starting_row        '566
		SIBL_group_1 = ObjExcel.Cells(SIBL_starting_excel_row, current_excel_col).Value
		SIBL_group_2 = ObjExcel.Cells(SIBL_starting_excel_row + 1, current_excel_col).Value
		SIBL_group_3 = ObjExcel.Cells(SIBL_starting_excel_row + 2, current_excel_col).Value
        starting_row = starting_row + 3

		SPON_starting_excel_row = starting_row        '569
		SPON_type = left(ObjExcel.Cells(SPON_starting_excel_row, current_excel_col).Value, 2)
		SPON_ver = ObjExcel.Cells(SPON_starting_excel_row + 1, current_excel_col).Value
		SPON_name = ObjExcel.Cells(SPON_starting_excel_row + 2, current_excel_col).Value
		SPON_state = ObjExcel.Cells(SPON_starting_excel_row + 3, current_excel_col).Value
        starting_row = starting_row + 4

		STEC_starting_excel_row = starting_row        '573
		STEC_type_1 = left(ObjExcel.Cells(STEC_starting_excel_row, current_excel_col).Value, 2)
		STEC_amt_1 = ObjExcel.Cells(STEC_starting_excel_row + 1, current_excel_col).Value
		STEC_actual_from_thru_months_1 = ObjExcel.Cells(STEC_starting_excel_row + 2, current_excel_col).Value
		STEC_ver_1 = left(ObjExcel.Cells(STEC_starting_excel_row + 3, current_excel_col).Value, 1)
		STEC_earmarked_amt_1 = ObjExcel.Cells(STEC_starting_excel_row + 4, current_excel_col).Value
		STEC_earmarked_from_thru_months_1 = ObjExcel.Cells(STEC_starting_excel_row + 5, current_excel_col).Value
		STEC_type_2 = left(ObjExcel.Cells(STEC_starting_excel_row + 6, current_excel_col).Value, 2)
		STEC_amt_2 = ObjExcel.Cells(STEC_starting_excel_row + 7, current_excel_col).Value
		STEC_actual_from_thru_months_2 = ObjExcel.Cells(STEC_starting_excel_row + 8, current_excel_col).Value
		STEC_ver_2 = left(ObjExcel.Cells(STEC_starting_excel_row + 9, current_excel_col).Value, 1)
		STEC_earmarked_amt_2 = ObjExcel.Cells(STEC_starting_excel_row + 10, current_excel_col).Value
		STEC_earmarked_from_thru_months_2 = ObjExcel.Cells(STEC_starting_excel_row + 11, current_excel_col).Value
        starting_row = starting_row + 12

		STIN_starting_excel_row = starting_row        '585
		STIN_type_1 = left(ObjExcel.Cells(STIN_starting_excel_row, current_excel_col).Value, 2)
		STIN_amt_1 = ObjExcel.Cells(STIN_starting_excel_row + 1, current_excel_col).Value
		STIN_avail_date_1 = ObjExcel.Cells(STIN_starting_excel_row + 2, current_excel_col).Value
		STIN_months_covered_1 = ObjExcel.Cells(STIN_starting_excel_row + 3, current_excel_col).Value
		STIN_ver_1 = left(ObjExcel.Cells(STIN_starting_excel_row + 4, current_excel_col).Value, 1)
		STIN_type_2 = left(ObjExcel.Cells(STIN_starting_excel_row + 5, current_excel_col).Value, 2)
		STIN_amt_2 = ObjExcel.Cells(STIN_starting_excel_row + 6, current_excel_col).Value
		STIN_avail_date_2 = ObjExcel.Cells(STIN_starting_excel_row + 7, current_excel_col).Value
		STIN_months_covered_2 = ObjExcel.Cells(STIN_starting_excel_row + 8, current_excel_col).Value
		STIN_ver_2 = left(ObjExcel.Cells(STIN_starting_excel_row + 9, current_excel_col).Value, 1)
        starting_row = starting_row + 10

		STWK_starting_excel_row = starting_row        '595
		STWK_empl_name = ObjExcel.Cells(STWK_starting_excel_row, current_excel_col).Value
		STWK_wrk_stop_date = ObjExcel.Cells(STWK_starting_excel_row + 1, current_excel_col).Value
		STWK_wrk_stop_date_verif = left(ObjExcel.Cells(STWK_starting_excel_row + 2, current_excel_col).Value, 1)
		STWK_inc_stop_date = ObjExcel.Cells(STWK_starting_excel_row + 3, current_excel_col).Value
		STWK_refused_empl_yn = ObjExcel.Cells(STWK_starting_excel_row + 4, current_excel_col).Value
		STWK_vol_quit = ObjExcel.Cells(STWK_starting_excel_row + 5, current_excel_col).Value
		STWK_ref_empl_date = ObjExcel.Cells(STWK_starting_excel_row + 6, current_excel_col).Value
		STWK_gc_cash = ObjExcel.Cells(STWK_starting_excel_row + 7, current_excel_col).Value
		STWK_gc_grh = ObjExcel.Cells(STWK_starting_excel_row + 8, current_excel_col).Value
		STWK_gc_fs = ObjExcel.Cells(STWK_starting_excel_row + 9, current_excel_col).Value
		STWK_fs_pwe = ObjExcel.Cells(STWK_starting_excel_row + 10, current_excel_col).Value
		STWK_maepd_ext = left(ObjExcel.Cells(STWK_starting_excel_row + 11, current_excel_col).Value, 1)
        starting_row = starting_row + 12

		UNEA_1_starting_excel_row = starting_row      '607
		UNEA_1_inc_type = left(ObjExcel.Cells(UNEA_1_starting_excel_row, current_excel_col).Value, 2)
		UNEA_1_inc_verif = left(ObjExcel.Cells(UNEA_1_starting_excel_row + 1, current_excel_col).Value, 1)
		UNEA_1_claim_suffix = ObjExcel.Cells(UNEA_1_starting_excel_row + 2, current_excel_col).Value
		UNEA_1_start_date = ObjExcel.Cells(UNEA_1_starting_excel_row + 3, current_excel_col).Value
		UNEA_1_pay_freq = ObjExcel.Cells(UNEA_1_starting_excel_row + 4, current_excel_col).Value
		UNEA_1_inc_amount = ObjExcel.Cells(UNEA_1_starting_excel_row + 5, current_excel_col).Value
        starting_row = starting_row + 6

		UNEA_2_starting_excel_row = starting_row      '613
		UNEA_2_inc_type = left(ObjExcel.Cells(UNEA_2_starting_excel_row, current_excel_col).Value, 2)
		UNEA_2_inc_verif = left(ObjExcel.Cells(UNEA_2_starting_excel_row + 1, current_excel_col).Value, 1)
		UNEA_2_claim_suffix = ObjExcel.Cells(UNEA_2_starting_excel_row + 2, current_excel_col).Value
		UNEA_2_start_date = ObjExcel.Cells(UNEA_2_starting_excel_row + 3, current_excel_col).Value
		UNEA_2_pay_freq = ObjExcel.Cells(UNEA_2_starting_excel_row + 4, current_excel_col).Value
		UNEA_2_inc_amount = ObjExcel.Cells(UNEA_2_starting_excel_row + 5, current_excel_col).Value
        starting_row = starting_row + 6

		UNEA_3_starting_excel_row = starting_row      '619
		UNEA_3_inc_type = left(ObjExcel.Cells(UNEA_3_starting_excel_row, current_excel_col).Value, 2)
		UNEA_3_inc_verif = left(ObjExcel.Cells(UNEA_3_starting_excel_row + 1, current_excel_col).Value, 1)
		UNEA_3_claim_suffix = ObjExcel.Cells(UNEA_3_starting_excel_row + 2, current_excel_col).Value
		UNEA_3_start_date = ObjExcel.Cells(UNEA_3_starting_excel_row + 3, current_excel_col).Value
		UNEA_3_pay_freq = ObjExcel.Cells(UNEA_3_starting_excel_row + 4, current_excel_col).Value
		UNEA_3_inc_amount = ObjExcel.Cells(UNEA_3_starting_excel_row + 5, current_excel_col).Value
        starting_row = starting_row + 6

		WKEX_starting_excel_row = starting_row        '625
		WKEX_program = objExcel.Cells(WKEX_starting_excel_row, current_excel_col).Value
		WKEX_fed_tax_retro = objExcel.Cells(WKEX_starting_excel_row + 1, current_excel_col).Value
		WKEX_fed_tax_prosp = objExcel.Cells(WKEX_starting_excel_row + 2, current_excel_col).Value
		WKEX_fed_tax_verif = left(objExcel.Cells(WKEX_starting_excel_row + 3, current_excel_col).Value, 1)
		WKEX_state_tax_retro = objExcel.Cells(WKEX_starting_excel_row + 4, current_excel_col).Value
		WKEX_state_tax_prosp = objExcel.Cells(WKEX_starting_excel_row + 5, current_excel_col).Value
		WKEX_state_tax_verif = left(objExcel.Cells(WKEX_starting_excel_row + 6, current_excel_col).Value, 1)
		WKEX_fica_retro = objExcel.Cells(WKEX_starting_excel_row + 7, current_excel_col).Value
		WKEX_fica_prosp = objExcel.Cells(WKEX_starting_excel_row + 8, current_excel_col).Value
		WKEX_fica_verif = left(objExcel.Cells(WKEX_starting_excel_row + 9, current_excel_col).Value, 1)
		WKEX_tran_retro = objExcel.Cells(WKEX_starting_excel_row + 10, current_excel_col).Value
		WKEX_tran_prosp = objExcel.Cells(WKEX_starting_excel_row + 11, current_excel_col).Value
		WKEX_tran_verif = left(objExcel.Cells(WKEX_starting_excel_row + 12, current_excel_col).Value, 1)
		WKEX_tran_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 13, current_excel_col).Value
		WKEX_meals_retro = objExcel.Cells(WKEX_starting_excel_row + 14, current_excel_col).Value
		WKEX_meals_prosp = objExcel.Cells(WKEX_starting_excel_row + 15, current_excel_col).Value
		WKEX_meals_verif = left(objExcel.Cells(WKEX_starting_excel_row + 16, current_excel_col).Value, 1)
		WKEX_meals_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 17, current_excel_col).Value
		WKEX_uniforms_retro = objExcel.Cells(WKEX_starting_excel_row + 18, current_excel_col).Value
		WKEX_uniforms_prosp = objExcel.Cells(WKEX_starting_excel_row + 19, current_excel_col).Value
		WKEX_uniforms_verif = left(objExcel.Cells(WKEX_starting_excel_row + 20, current_excel_col).Value, 1)
		WKEX_uniforms_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 21, current_excel_col).Value
		WKEX_tools_retro = objExcel.Cells(WKEX_starting_excel_row + 22, current_excel_col).Value
		WKEX_tools_prosp = objExcel.Cells(WKEX_starting_excel_row + 23, current_excel_col).Value
		WKEX_tools_verif = left(objExcel.Cells(WKEX_starting_excel_row + 24, current_excel_col).Value, 1)
		WKEX_tools_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 25, current_excel_col).Value
		WKEX_dues_retro = objExcel.Cells(WKEX_starting_excel_row + 26, current_excel_col).Value
		WKEX_dues_prosp = objExcel.Cells(WKEX_starting_excel_row + 27, current_excel_col).Value
		WKEX_dues_verif = left(objExcel.Cells(WKEX_starting_excel_row + 28, current_excel_col).Value, 1)
		WKEX_dues_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 29, current_excel_col).Value
		WKEX_othr_retro = objExcel.Cells(WKEX_starting_excel_row + 30, current_excel_col).Value
		WKEX_othr_prosp = objExcel.Cells(WKEX_starting_excel_row + 31, current_excel_col).Value
		WKEX_othr_verif = left(objExcel.Cells(WKEX_starting_excel_row + 32, current_excel_col).Value, 1)
		WKEX_othr_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 33, current_excel_col).Value
		WKEX_HC_Exp_Fed_Tax = objExcel.Cells(WKEX_starting_excel_row + 34, current_excel_col).Value
		WKEX_HC_Exp_State_Tax = objExcel.Cells(WKEX_starting_excel_row + 35, current_excel_col).Value
		WKEX_HC_Exp_FICA = objExcel.Cells(WKEX_starting_excel_row + 36, current_excel_col).Value
		WKEX_HC_Exp_Tran = objExcel.Cells(WKEX_starting_excel_row + 37, current_excel_col).Value
		WKEX_HC_Exp_Tran_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 38, current_excel_col).Value
		WKEX_HC_Exp_Meals = objExcel.Cells(WKEX_starting_excel_row + 39, current_excel_col).Value
		WKEX_HC_Exp_Meals_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 40, current_excel_col).Value
		WKEX_HC_Exp_Uniforms = objExcel.Cells(WKEX_starting_excel_row + 41, current_excel_col).Value
		WKEX_HC_Exp_Uniforms_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 42, current_excel_col).Value
		WKEX_HC_Exp_Tools = objExcel.Cells(WKEX_starting_excel_row + 43, current_excel_col).Value
		WKEX_HC_Exp_Tools_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 44, current_excel_col).Value
		WKEX_HC_Exp_Dues = objExcel.Cells(WKEX_starting_excel_row + 45, current_excel_col).Value
		WKEX_HC_Exp_Dues_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 46, current_excel_col).Value
		WKEX_HC_Exp_Othr = objExcel.Cells(WKEX_starting_excel_row + 47, current_excel_col).Value
		WKEX_HC_Exp_Othr_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 48, current_excel_col).Value
        starting_row = starting_row + 49

		WREG_starting_excel_row = starting_row        '674
		WREG_fs_pwe = ObjExcel.Cells(WREG_starting_excel_row, current_excel_col).Value
		WREG_fset_status = left(ObjExcel.Cells(WREG_starting_excel_row + 1, current_excel_col).Value, 2)
		WREG_defer_fs = ObjExcel.Cells(WREG_starting_excel_row + 2, current_excel_col).Value
		WREG_fset_orientation_date = ObjExcel.Cells(WREG_starting_excel_row + 3, current_excel_col).Value
		WREG_fset_sanction_date = ObjExcel.Cells(WREG_starting_excel_row + 4, current_excel_col).Value
        wreg_sanction_reason = ObjExcel.Cells(WREG_starting_excel_row + 5, current_excel_col).Value
		WREG_num_sanctions = ObjExcel.Cells(WREG_starting_excel_row + 6, current_excel_col).Value
		WREG_abawd_status = left(ObjExcel.Cells(WREG_starting_excel_row + 7, current_excel_col).Value, 2)
		WREG_ga_basis = left(ObjExcel.Cells(WREG_starting_excel_row + 8, current_excel_col).Value, 2)
        starting_row = starting_row + 9

		'-------------------------------ACTUALLY FILLING OUT MAXIS

		'Goes to STAT/MEMB to associate a SSN to each member, this will be useful for UNEA/MEDI panels
		MAXIS_background_check
		call navigate_to_MAXIS_screen("STAT", "MEMB")
		EMWriteScreen reference_number, 20, 76
		transmit
		EMReadScreen SSN_first, 3, 7, 42
		EMReadScreen SSN_mid, 2, 7, 46
		EMReadScreen SSN_last, 4, 7, 49

		'ACCT
		If ACCT_type <> "" then
			call write_panel_to_MAXIS_ACCT(ACCT_type, ACCT_numb, ACCT_location, ACCT_balance, ACCT_bal_ver, ACCT_date, ACCT_withdraw, ACCT_cash_count, ACCT_snap_count, ACCT_HC_count, ACCT_GRH_count, ACCT_IV_count, ACCT_joint_owner, ACCT_share_ratio, ACCT_interest_date_mo, ACCT_interest_date_yr)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'ACUT
		If ACUT_shared <> "" then
			call write_panel_to_MAXIS_ACUT(ACUT_shared, ACUT_heat, ACUT_air, ACUT_electric, ACUT_fuel, ACUT_garbage, ACUT_water, ACUT_sewer, ACUT_other, ACUT_phone, ACUT_heat_verif, ACUT_air_verif, ACUT_electric_verif, ACUT_fuel_verif, ACUT_garbage_verif, ACUT_water_verif, ACUT_sewer_verif, ACUT_other_verif)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'BILS
		IF BILS_bill_1_ref_num <> "" THEN
			CALL write_panel_to_MAXIS_BILS(BILS_bill_1_ref_num, BILS_bill_1_serv_date, BILS_bill_1_serv_type, BILS_bill_1_gross_amt, BILS_bill_1_third_party, BILS_bill_1_verif, BILS_bill_1_BILS_type, BILS_bill_2_ref_num, BILS_bill_2_serv_date, BILS_bill_2_serv_type, BILS_bill_2_gross_amt, BILS_bill_2_third_party, BILS_bill_2_verif, BILS_bill_2_BILS_type, BILS_bill_3_ref_num, BILS_bill_3_serv_date, BILS_bill_3_serv_type, BILS_bill_3_gross_amt, BILS_bill_3_third_party, BILS_bill_3_verif, BILS_bill_3_BILS_type, BILS_bill_4_ref_num, BILS_bill_4_serv_date, BILS_bill_4_serv_type, BILS_bill_4_gross_amt, BILS_bill_4_third_party, BILS_bill_4_verif, BILS_bill_4_BILS_type, BILS_bill_5_ref_num, BILS_bill_5_serv_date, BILS_bill_5_serv_type, BILS_bill_5_gross_amt, BILS_bill_5_third_party, BILS_bill_5_verif, BILS_bill_5_BILS_type, BILS_bill_6_ref_num, BILS_bill_6_serv_date, BILS_bill_6_serv_type, BILS_bill_6_gross_amt, BILS_bill_6_third_party, BILS_bill_6_verif, BILS_bill_6_BILS_type, BILS_bill_7_ref_num, BILS_bill_7_serv_date, BILS_bill_7_serv_type, BILS_bill_7_gross_amt, BILS_bill_7_third_party, BILS_bill_7_verif, BILS_bill_7_BILS_type, BILS_bill_8_ref_num, BILS_bill_8_serv_date, BILS_bill_8_serv_type, BILS_bill_8_gross_amt, BILS_bill_8_third_party, BILS_bill_8_verif, BILS_bill_8_BILS_type, BILS_bill_9_ref_num, BILS_bill_9_serv_date, BILS_bill_9_serv_type, BILS_bill_9_gross_amt, BILS_bill_9_third_party, BILS_bill_9_verif, BILS_bill_9_BILS_type)
			STATS_manualtime = STATS_manualtime + 50
		END IF
		'BUSI
		If BUSI_type <> "" then
			call write_panel_to_MAXIS_BUSI(busi_type, busi_start_date, busi_end_date, busi_cash_total_retro, busi_cash_total_prosp, busi_cash_total_ver, busi_IV_total_prosp, busi_IV_total_ver, busi_snap_total_retro, busi_snap_total_prosp, busi_snap_total_ver, busi_hc_total_prosp_a, busi_hc_total_ver_a, busi_hc_total_prosp_b, busi_hc_total_ver_b, busi_cash_exp_retro, busi_cash_exp_prosp, busi_cash_exp_ver, busi_IV_exp_prosp, busi_IV_exp_ver, busi_snap_exp_retro, busi_snap_exp_prosp, busi_snap_exp_ver, busi_hc_exp_prosp_a, busi_hc_exp_ver_a, busi_hc_exp_prosp_b, busi_hc_exp_ver_b, busi_retro_hours, busi_prosp_hours, busi_hc_total_est_a, busi_hc_total_est_b, busi_hc_exp_est_a, busi_hc_exp_est_b, busi_hc_hours_est)
			STATS_manualtime = STATS_manualtime + 60
		END IF
		'CARS
		If CARS_type <> "" then
			call write_panel_to_MAXIS_CARS(cars_type, cars_year, cars_make, cars_model, cars_trade_in, cars_loan, cars_value_source, cars_ownership_ver, cars_amount_owed, cars_amount_owed_ver, cars_date, cars_use, cars_HC_benefit, cars_joint_owner, cars_share_ratio)
			STATS_manualtime = STATS_manualtime + 25
		END IF
		'CASH
		If CASH_amount <> "" then
			call write_panel_to_MAXIS_CASH(cash_amount)
			STATS_manualtime = STATS_manualtime + 5
		END IF
		'COEX
		IF COEX_support_retro <> "" OR _
			COEX_support_prosp <> "" OR _
			COEX_support_verif <> "" OR _
			COEX_alimony_retro <> "" OR _
			COEX_alimony_prosp <> "" OR _
			COEX_alimony_verif <> "" OR _
			COEX_tax_dep_retro <> "" OR _
			COEX_tax_dep_prosp <> "" OR _
			COEX_tax_dep_verif <> "" OR _
			COEX_other_retro <> "" OR _
			COEX_other_prosp <> "" OR _
			COEX_other_verif <> "" OR _
			COEX_change_in_circumstances <> "" OR _
			COEX_HC_expense_support <> "" OR _
			COEX_HC_expense_alimony <> "" OR _
			COEX_HC_expense_tax_dep <> "" OR _
			COEX_HC_expense_other <> "" THEN
				CALL write_panel_to_MAXIS_COEX(COEX_support_retro, COEX_support_prosp, COEX_support_verif, COEX_alimony_retro, COEX_alimony_prosp, COEX_alimony_verif, COEX_tax_dep_retro, COEX_tax_dep_prosp, COEX_tax_dep_verif, COEX_other_retro, COEX_other_prosp, COEX_other_verif, COEX_change_in_circumstances, COEX_HC_expense_support, COEX_HC_expense_alimony, COEX_HC_expense_tax_dep, COEX_HC_expense_other)
				STATS_manualtime = STATS_manualtime + 20
		END IF
		'DCEX
		If DCEX_provider <> "" then
			call write_panel_to_MAXIS_DCEX(DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'DFLN
		IF DFLN_conv_1_dt <> "" THEN
			CALL write_panel_to_MAXIS_DFLN(DFLN_conv_1_dt, DFLN_conv_1_juris, DFLN_conv_1_state, DFLN_conv_2_dt, DFLN_conv_2_juris, DFLN_conv_2_state, DFLN_rnd_test_1_dt, DFLN_rnd_test_1_provider, DFLN_rnd_test_1_result, DFLN_rnd_test_2_dt, DFLN_rnd_test_2_provider, DFLN_rnd_test_2_result)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'DIET
		If DIET_mfip_1 <> "" or DIET_MSA_1 <> "" then
			call write_panel_to_MAXIS_DIET(DIET_mfip_1, DIET_mfip_1_ver, DIET_mfip_2, DIET_mfip_2_ver, DIET_msa_1, DIET_msa_1_ver, DIET_msa_2, DIET_msa_2_ver, DIET_msa_3, DIET_msa_3_ver, DIET_msa_4, DIET_msa_4_ver)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'DISA
		If DISA_begin_date <> "" then
			call write_panel_to_MAXIS_DISA(disa_begin_date, disa_end_date, disa_cert_begin, disa_cert_end, disa_wavr_begin, disa_wavr_end, disa_grh_begin, disa_grh_end, disa_cash_status, disa_cash_status_ver, disa_snap_status, disa_snap_status_ver, disa_hc_status, disa_hc_status_ver, disa_waiver, disa_1619, disa_drug_alcohol)
			STATS_manualtime = STATS_manualtime + 30
		END IF
		'DSTT
		If DSTT_ongoing_income <> "" then
			call write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, DSTT_HH_income_stop_date, DSTT_income_expected_amt)
			STATS_manualtime = STATS_manualtime + 10
		END IF
		'EATS
		If EATS_together <> "" then
			call write_panel_to_MAXIS_EATS(eats_together, eats_boarder, eats_group_one, eats_group_two, eats_group_three)
			STATS_manualtime = STATS_manualtime + 18
		END IF
		'EMMA
		If EMMA_medical_emergency <> "" then
			call write_panel_to_MAXIS_EMMA(EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date)
			STATS_manualtime = STATS_manualtime + 15
		END IF
		'EMPS
		If EMPS_memb_at_home <> "" then
			call write_panel_to_MAXIS_EMPS(EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date)
			STATS_manualtime = STATS_manualtime + 35
		END IF
		'FACI
		If FACI_name <> "" then
			call write_panel_to_MAXIS_FACI(FACI_vendor_number, FACI_name, FACI_type, FACI_FS_eligible, FACI_FS_facility_type, FACI_date_in, FACI_date_out)
			STATS_manualtime = STATS_manualtime + 32
		END IF
		'FMED
		IF FMED_medical_mileage <> "" OR FMED_1_type <> "" OR FMED_2_type <> "" OR FMED_3_type <> "" OR FMED_4_type <> "" THEN
			CALL write_panel_to_MAXIS_FMED(FMED_medical_mileage, FMED_1_type, FMED_1_verif, FMED_1_ref_num, FMED_1_category, FMED_1_begin, FMED_1_end, FMED_1_amount, FMED_2_type, FMED_2_verif, FMED_2_ref_num, FMED_2_category, FMED_2_begin, FMED_2_end, FMED_2_amount, FMED_3_type, FMED_3_verif, FMED_3_ref_num, FMED_3_category, FMED_3_begin, FMED_3_end, FMED_3_amount, FMED_4_type, FMED_4_verif, FMED_4_ref_num, FMED_4_category, FMED_4_begin, FMED_4_end, FMED_4_amount)
			STATS_manualtime = STATS_manualtime + 25
		END IF
		'HEST
		If HEST_FS_choice_date <> "" then
			call write_panel_to_MAXIS_HEST(HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro)
			STATS_manualtime = STATS_manualtime + 10
		END IF
		'IMIG
		If IMIG_imigration_status <> "" THEN
			CALL write_panel_to_MAXIS_IMIG(IMIG_imigration_status, IMIG_entry_date, IMIG_status_date, IMIG_status_ver, IMIG_status_LPR_adj_from, IMIG_nationality, IMIG_40_soc_sec, IMIG_40_soc_sec_verif, IMIG_battered_spouse_child, IMIG_battered_spouse_child_verif, IMIG_military_status, IMIG_military_status_verif, IMIG_hmong_lao_nat_amer, IMIG_st_prog_esl_ctzn_coop, IMIG_st_prog_esl_ctzn_coop_verif, IMIG_fss_esl_skills_training)
			STATS_manualtime = STATS_manualtime + 27
		END IF
		'INSA
		If INSA_pers_coop_ohi <> "" then
			call write_panel_to_MAXIS_INSA(INSA_pers_coop_ohi,INSA_good_cause_status,INSA_good_cause_cliam_date,INSA_good_cause_evidence,INSA_coop_cost_effect,INSA_insur_name,INSA_prescrip_drug_cover,INSA_prescrip_end_date, INSA_persons_covered)
			STATS_manualtime = STATS_manualtime + 16
		END IF
		'JOBS1
		If JOBS_1_inc_type <> "" then
			call write_panel_to_MAXIS_JOBS("01", JOBS_1_inc_type, JOBS_1_inc_verif, JOBS_1_employer_name, JOBS_1_inc_start, JOBS_1_wkly_hrs, JOBS_1_hrly_wage, JOBS_1_pay_freq)
			STATS_manualtime = STATS_manualtime + 50
		END IF
		'JOBS2
		If JOBS_2_inc_type <> "" then
			call write_panel_to_MAXIS_JOBS("02", JOBS_2_inc_type, JOBS_2_inc_verif, JOBS_2_employer_name, JOBS_2_inc_start, JOBS_2_wkly_hrs, JOBS_2_hrly_wage, JOBS_2_pay_freq)
			STATS_manualtime = STATS_manualtime + 50
		END IF
		'JOBS3
		If JOBS_3_inc_type <> "" then
			call write_panel_to_MAXIS_JOBS("03", JOBS_3_inc_type, JOBS_3_inc_verif, JOBS_3_employer_name, JOBS_3_inc_start, JOBS_3_wkly_hrs, JOBS_3_hrly_wage, JOBS_3_pay_freq)
			STATS_manualtime = STATS_manualtime + 50
		END IF
		'MEDI
		If MEDI_claim_number_suffix <> "" then
			call write_panel_to_MAXIS_MEDI(SSN_first, SSN_mid, SSN_last, MEDI_claim_number_suffix, MEDI_part_A_premium, MEDI_part_B_premium, MEDI_part_A_begin_date, MEDI_part_B_begin_date, MEDI_apply_prem_to_spdn, MEDI_apply_prem_end_date)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'MMSA
		If MMSA_liv_arr <> "" then
			call write_panel_to_MAXIS_MMSA(MMSA_liv_arr, MMSA_cont_elig, MMSA_spous_inc, MMSA_shared_hous)
			STATS_manualtime = STATS_manualtime + 10
		END IF
		'MSUR
		If MSUR_begin_date <> "" then
			call write_panel_to_MAXIS_MSUR(MSUR_begin_date)
			STATS_manualtime = STATS_manualtime + 7
		END IF
		'OTHR
		If OTHR_type <> "" then
			call write_panel_to_MAXIS_OTHR(OTHR_type, OTHR_cash_value, OTHR_cash_value_ver, OTHR_owed, OTHR_owed_ver, OTHR_date, OTHR_cash_count, OTHR_SNAP_count, OTHR_HC_count, OTHR_IV_count, OTHR_joint, OTHR_share_ratio)
			STATS_manualtime = STATS_manualtime + 16
		END IF
		'PARE
		If PARE_child_1 <> "" then
			call write_panel_to_MAXIS_PARE(appl_date, reference_number, PARE_child_1, PARE_child_1_relation, PARE_child_1_verif, PARE_child_2, PARE_child_2_relation, PARE_child_2_verif, PARE_child_3, PARE_child_3_relation, PARE_child_3_verif, PARE_child_4, PARE_child_4_relation, PARE_child_4_verif, PARE_child_5, PARE_child_5_relation, PARE_child_5_verif, PARE_child_6, PARE_child_6_relation, PARE_child_6_verif)
			STATS_manualtime = STATS_manualtime + 14
		END IF
		'ABPS (must do after PARE, because the ABPS function checks PARE for a child list)
		If abps_supp_coop <> "" then
			call write_panel_to_MAXIS_ABPS(abps_supp_coop,abps_gc_status)
			STATS_manualtime = STATS_manualtime + 45
		END IF
		'PBEN 1
		If PBEN_1_IAA_date <> "" then
			call write_panel_to_MAXIS_PBEN(PBEN_1_referal_date, PBEN_1_type, PBEN_1_appl_date, PBEN_1_appl_ver, PBEN_1_IAA_date, PBEN_1_disp)
			STATS_manualtime = STATS_manualtime + 25
		END IF
		'PBEN 2
		If PBEN_2_IAA_date <> "" then
			call write_panel_to_MAXIS_PBEN(PBEN_2_referal_date, PBEN_2_type, PBEN_2_appl_date, PBEN_2_appl_ver, PBEN_2_IAA_date, PBEN_2_disp)
			STATS_manualtime = STATS_manualtime + 25
		END IF
		'PBEN 3
		If PBEN_3_IAA_date <> "" then
			call write_panel_to_MAXIS_PBEN(PBEN_3_referal_date, PBEN_3_type, PBEN_3_appl_date, PBEN_3_appl_ver, PBEN_3_IAA_date, PBEN_3_disp)
			STATS_manualtime = STATS_manualtime + 25
		END IF
		'PDED
		If PDED_wid_deduction <> "" OR _
			PDED_adult_child_disregard <> "" OR _
			PDED_wid_disregard <> "" OR _
			PDED_unea_income_deduction_reason <> "" OR _
			PDED_unea_income_deduction_value <> "" OR _
			PDED_earned_income_deduction_reason <> "" OR _
			PDED_earned_income_deduction_value <> "" OR _
			PDED_ma_epd_inc_asset_limit <> "" OR _
			PDED_guard_fee <> "" OR _
			PDED_rep_payee_fee <> "" OR _
			PDED_other_expense <> "" OR _
			PDED_shel_spcl_needs <> "" OR _
			PDED_excess_need <> "" OR _
			PDED_restaurant_meals <> "" THEN
				CALL write_panel_to_MAXIS_PDED(PDED_wid_deduction, PDED_adult_child_disregard, PDED_wid_disregard, PDED_unea_income_deduction_reason, PDED_unea_income_deduction_value, PDED_earned_income_deduction_reason, PDED_earned_income_deduction_value, PDED_ma_epd_inc_asset_limit, PDED_guard_fee, PDED_rep_payee_fee, PDED_other_expense, PDED_shel_spcl_needs, PDED_excess_need, PDED_restaurant_meals)
				STATS_manualtime = STATS_manualtime + 20
		END IF
		'PREG
		If PREG_conception_date <> "" then
			call write_panel_to_MAXIS_PREG(PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver,PREG_due_date, PREG_multiple_birth)
			STATS_manualtime = STATS_manualtime + 30
		END IF
		'RBIC
		If rbic_type <> "" then
			call write_panel_to_MAXIS_RBIC(rbic_type, rbic_start_date, rbic_end_date, rbic_group_1, rbic_retro_income_group_1, rbic_prosp_income_group_1, rbic_ver_income_group_1, rbic_group_2, rbic_retro_income_group_2, rbic_prosp_income_group_2, rbic_ver_income_group_2, rbic_group_3, rbic_retro_income_group_3, rbic_prosp_income_group_3, rbic_ver_income_group_3, rbic_retro_hours, rbic_prosp_hours, rbic_exp_type_1, rbic_exp_retro_1, rbic_exp_prosp_1, rbic_exp_ver_1, rbic_exp_type_2, rbic_exp_retro_2, rbic_exp_prosp_2, rbic_exp_ver_2)
			STATS_manualtime = STATS_manualtime + 28
		END IF
		'REST
		If rest_type <> "" then
			call write_panel_to_MAXIS_REST(rest_type, rest_type_ver, rest_market, rest_market_ver, rest_owed, rest_owed_ver, rest_date, rest_status, rest_joint, rest_share_ratio, rest_agreement_date)
			STATS_manualtime = STATS_manualtime + 26
		END IF
		'SCHL
		If SCHL_status <> "" then
			If right(left(SCHL_grad_date, 2), 1) = "/" Then SCHL_grad_date = "0" & SCHL_grad_date	'Making sure that the grad date has the correct formating
			call write_panel_to_MAXIS_SCHL(appl_date, SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'SECU
		If secu_type <> "" then
			call write_panel_to_MAXIS_SECU(secu_type, secu_pol_numb, secu_name, secu_cash_val, secu_date, secu_cash_ver, secu_face_val, secu_withdraw, secu_cash_count, secu_SNAP_count, secu_HC_count, secu_GRH_count, secu_IV_count, secu_joint, secu_share_ratio)
			STATS_manualtime = STATS_manualtime + 23
		END IF
		'SHEL
		If SHEL_subsidized <> "" then
			call write_panel_to_MAXIS_SHEL(SHEL_subsidized, SHEL_shared, SHEL_paid_to, SHEL_rent_retro, SHEL_rent_retro_ver, SHEL_rent_pro, SHEL_rent_pro_ver, SHEL_lot_rent_retro, SHEL_lot_rent_retro_ver, SHEL_lot_rent_pro, SHEL_lot_rent_pro_ver, SHEL_mortgage_retro, SHEL_mortgage_retro_ver, SHEL_mortgage_pro, SHEL_mortgage_pro_ver, SHEL_insur_retro, SHEL_insur_retro_ver, SHEL_insur_pro, SHEL_insur_pro_ver, SHEL_taxes_retro, SHEL_taxes_retro_ver, SHEL_taxes_pro, SHEL_taxes_pro_ver, SHEL_room_retro, SHEL_room_retro_ver, SHEL_room_pro, SHEL_room_pro_ver, SHEL_garage_retro, SHEL_garage_retro_ver, SHEL_garage_pro, SHEL_garage_pro_ver, SHEL_subsidy_retro, SHEL_subsidy_retro_ver, SHEL_subsidy_pro, SHEL_subsidy_pro_ver)
			STATS_manualtime = STATS_manualtime + 19
		END IF
		'SIBL
		If SIBL_group_1 <> "" then
			call write_panel_to_MAXIS_SIBL(SIBL_group_1, SIBL_group_2, SIBL_group_3)
			STATS_manualtime = STATS_manualtime + 9
		END IF
		'SPON
		If SPON_type <> "" then
			call write_panel_to_MAXIS_SPON(SPON_type, SPON_ver, SPON_name, SPON_state)
			STATS_manualtime = STATS_manualtime + 27
		END IF
		'STEC
		If STEC_type_1 <> "" then
			call write_panel_to_MAXIS_STEC(STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2)
			STATS_manualtime = STATS_manualtime + 30
		END IF
		'STIN
		If STIN_type_1 <> "" then
			call write_panel_to_MAXIS_STIN(STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2)
			STATS_manualtime = STATS_manualtime + 33
		END IF
		'STWK
		If STWK_empl_name <> "" then
			call write_panel_to_MAXIS_STWK(STWK_empl_name, STWK_wrk_stop_date, STWK_wrk_stop_date_verif, STWK_inc_stop_date, STWK_refused_empl_yn, STWK_vol_quit, STWK_ref_empl_date, STWK_gc_cash, STWK_gc_grh, STWK_gc_fs, STWK_fs_pwe, STWK_maepd_ext)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'UNEA 1
		If UNEA_1_inc_type <> "" then
			call write_panel_to_MAXIS_UNEA("01", UNEA_1_inc_type, UNEA_1_inc_verif, UNEA_1_claim_suffix, UNEA_1_start_date, UNEA_1_pay_freq, UNEA_1_inc_amount, SSN_first, SSN_mid, SSN_last)
			STATS_manualtime = STATS_manualtime + 33
		END IF
		'UNEA 2
		If UNEA_2_inc_type <> "" then
			call write_panel_to_MAXIS_UNEA("02", UNEA_2_inc_type, UNEA_2_inc_verif, UNEA_2_claim_suffix, UNEA_2_start_date, UNEA_2_pay_freq, UNEA_2_inc_amount, SSN_first, SSN_mid, SSN_last)
			STATS_manualtime = STATS_manualtime + 30
		END IF
		'UNEA 3
		If UNEA_3_inc_type <> "" then
			call write_panel_to_MAXIS_UNEA("03", UNEA_3_inc_type, UNEA_3_inc_verif, UNEA_3_claim_suffix, UNEA_3_start_date, UNEA_3_pay_freq, UNEA_3_inc_amount, SSN_first, SSN_mid, SSN_last)
			STATS_manualtime = STATS_manualtime + 30
		END IF
		'WKEX
		IF WKEX_program <> "" THEN
			CALL write_panel_to_MAXIS_WKEX(WKEX_program, WKEX_fed_tax_retro, WKEX_fed_tax_prosp, WKEX_fed_tax_verif, WKEX_state_tax_retro, WKEX_state_tax_prosp, WKEX_state_tax_verif, WKEX_fica_retro, WKEX_fica_prosp, WKEX_fica_verif, WKEX_tran_retro, WKEX_tran_prosp, WKEX_tran_verif, WKEX_tran_imp_rel, WKEX_meals_retro, WKEX_meals_prosp, WKEX_meals_verif, WKEX_meals_imp_rel, WKEX_uniforms_retro, WKEX_uniforms_prosp, WKEX_uniforms_verif, WKEX_uniforms_imp_rel, WKEX_tools_retro, WKEX_tools_prosp, WKEX_tools_verif, WKEX_tools_imp_rel, WKEX_dues_retro, WKEX_dues_prosp, WKEX_dues_verif, WKEX_dues_imp_rel, WKEX_othr_retro, WKEX_othr_prosp, WKEX_othr_verif, WKEX_othr_imp_rel, WKEX_HC_Exp_Fed_Tax, WKEX_HC_Exp_State_Tax, WKEX_HC_Exp_FICA, WKEX_HC_Exp_Tran, WKEX_HC_Exp_Tran_imp_rel, WKEX_HC_Exp_Meals, WKEX_HC_Exp_Meals_Imp_Rel, WKEX_HC_Exp_Uniforms, WKEX_HC_Exp_Uniforms_Imp_Rel, WKEX_HC_Exp_Tools, WKEX_HC_Exp_Tools_Imp_Rel, WKEX_HC_Exp_Dues, WKEX_HC_Exp_Dues_Imp_Rel, WKEX_HC_Exp_Othr, WKEX_HC_Exp_Othr_Imp_Rel)
			STATS_manualtime = STATS_manualtime + 20
		END IF
		'WREG
		If WREG_fs_pwe <> "" OR WREG_ga_basis <> "" then
			call write_panel_to_MAXIS_WREG(wreg_fs_pwe, wreg_fset_status, wreg_defer_fs, wreg_fset_orientation_date, wreg_fset_sanction_date, wreg_num_sanctions, wreg_sanction_reason, wreg_abawd_status, wreg_ga_basis)
			STATS_manualtime = STATS_manualtime + 15
		END IF
	Next

	original_MAXIS_footer_month = MAXIS_footer_month
	original_MAXIS_footer_year = MAXIS_footer_year
	For each each_month in date_array
	' DO
		Call convert_date_into_MAXIS_footer_month(each_month, MAXIS_footer_month, MAXIS_footer_year)

		PF3		'---Navigates to STAT/WRAP
		EMReadScreen at_wrap, 4, 2, 46

		If at_wrap = "WRAP" Then
			Call write_value_and_transmit("Y", 16, 54)
		Else
			Call back_to_SELF
			MAXIS_background_check
		End If
		Do
			EMReadScreen SUMM_check, 4, 2, 46
			EMReadScreen footer_month_check, 2, 20, 55
			EMReadScreen footer_year_check, 2, 20, 58

			If SUMM_check <> "SUMM" or footer_month_check <> MAXIS_footer_month or footer_month_check <> MAXIS_footer_month Then
				Call back_to_SELF
				MAXIS_background_check
			End If
		Loop until SUMM_check = "SUMM" and footer_month_check = MAXIS_footer_month and footer_month_check = MAXIS_footer_month

		'---Now the script will update BUSI, COEX, DCEX, JAEORBS, UNEA, WKEX for future months.
		For current_memb = 1 to total_membs
			current_excel_col = current_memb + 2							'There's two columns before the first HH member, so we have to add 2 to get the current excel col
			reference_number = ObjExcel.Cells(2, current_excel_col).Value	'Always in the second row. This is the HH member number

			'Rereading the values for BUSI, COEX, DCEX, JAEORBS, UNEA, WKEX for that person so the script can update the HC Expenses and Income.
			BUSI_type = left(ObjExcel.Cells(BUSI_starting_excel_row, current_excel_col).Value, 2)
			BUSI_start_date = ObjExcel.Cells(BUSI_starting_excel_row + 1, current_excel_col).Value
			BUSI_end_date = ObjExcel.Cells(BUSI_starting_excel_row + 2, current_excel_col).Value
			BUSI_cash_total_retro = ObjExcel.Cells(BUSI_starting_excel_row + 3, current_excel_col).Value
			BUSI_cash_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 4, current_excel_col).Value
			BUSI_cash_total_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 5, current_excel_col).Value, 1)
			BUSI_IV_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 6, current_excel_col).Value
			BUSI_IV_total_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 7, current_excel_col).Value, 1)
			BUSI_snap_total_retro = ObjExcel.Cells(BUSI_starting_excel_row + 8, current_excel_col).Value
			BUSI_snap_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 9, current_excel_col).Value
			BUSI_snap_total_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 10, current_excel_col).Value, 1)
			BUSI_hc_total_prosp_a = ObjExcel.Cells(BUSI_starting_excel_row + 11, current_excel_col).Value
			BUSI_hc_total_ver_a = left(ObjExcel.Cells(BUSI_starting_excel_row + 12, current_excel_col).Value, 1)
			BUSI_hc_total_prosp_b = ObjExcel.Cells(BUSI_starting_excel_row + 13, current_excel_col).Value
			BUSI_hc_total_ver_b = left(ObjExcel.Cells(BUSI_starting_excel_row + 14, current_excel_col).Value, 1)
			BUSI_cash_exp_retro = ObjExcel.Cells(BUSI_starting_excel_row + 15, current_excel_col).Value
			BUSI_cash_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 16, current_excel_col).Value
			BUSI_cash_exp_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 17, current_excel_col).Value, 1)
			BUSI_IV_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 18, current_excel_col).Value
			BUSI_IV_exp_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 19, current_excel_col).Value, 1)
			BUSI_snap_exp_retro = ObjExcel.Cells(BUSI_starting_excel_row + 20, current_excel_col).Value
			BUSI_snap_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 21, current_excel_col).Value
			BUSI_snap_exp_ver = left(ObjExcel.Cells(BUSI_starting_excel_row + 22, current_excel_col).Value, 1)
			BUSI_hc_exp_prosp_a = ObjExcel.Cells(BUSI_starting_excel_row + 23, current_excel_col).Value
			BUSI_hc_exp_ver_a = left(ObjExcel.Cells(BUSI_starting_excel_row + 24, current_excel_col).Value, 1)
			BUSI_hc_exp_prosp_b = ObjExcel.Cells(BUSI_starting_excel_row + 25, current_excel_col).Value
			BUSI_hc_exp_ver_b = left(ObjExcel.Cells(BUSI_starting_excel_row + 26, current_excel_col).Value, 1)
			BUSI_retro_hours = ObjExcel.Cells(BUSI_starting_excel_row + 27, current_excel_col).Value
			BUSI_prosp_hours = ObjExcel.Cells(BUSI_starting_excel_row + 28, current_excel_col).Value
			BUSI_hc_total_est_a = ObjExcel.Cells(BUSI_starting_excel_row + 29, current_excel_col).Value
			BUSI_hc_total_est_b = ObjExcel.Cells(BUSI_starting_excel_row + 30, current_excel_col).Value
			BUSI_hc_exp_est_a = ObjExcel.Cells(BUSI_starting_excel_row + 31, current_excel_col).Value
			BUSI_hc_exp_est_b = ObjExcel.Cells(BUSI_starting_excel_row + 32, current_excel_col).Value
			BUSI_hc_hours_est = ObjExcel.Cells(BUSI_starting_excel_row + 33, current_excel_col).Value

			COEX_support_retro = ObjExcel.Cells(COEX_starting_excel_row, current_excel_col).Value
			COEX_support_prosp = ObjExcel.Cells(COEX_starting_excel_row + 1, current_excel_col).Value
			COEX_support_verif = left(ObjExcel.Cells(COEX_starting_excel_row + 2, current_excel_col).Value, 1)
			COEX_alimony_retro = ObjExcel.Cells(COEX_starting_excel_row + 3, current_excel_col).Value
			COEX_alimony_prosp = ObjExcel.Cells(COEX_starting_excel_row + 4, current_excel_col).Value
			COEX_alimony_verif = left(ObjExcel.Cells(COEX_starting_excel_row + 5, current_excel_col).Value, 1)
			COEX_tax_dep_retro = ObjExcel.Cells(COEX_starting_excel_row + 6, current_excel_col).Value
			COEX_tax_dep_prosp = ObjExcel.Cells(COEX_starting_excel_row + 7, current_excel_col).Value
			COEX_tax_dep_verif = left(ObjExcel.Cells(COEX_starting_excel_row + 8, current_excel_col).Value, 1)
			COEX_other_retro = ObjExcel.Cells(COEX_starting_excel_row + 9, current_excel_col).Value
			COEX_other_prosp = ObjExcel.Cells(COEX_starting_excel_row + 10, current_excel_col).Value
			COEX_other_verif = left(ObjExcel.Cells(COEX_starting_excel_row + 11, current_excel_col).Value, 1)
			COEX_change_in_circumstances = left(ObjExcel.Cells(COEX_starting_excel_row + 12, current_excel_col).Value, 1)
			COEX_HC_expense_support = ObjExcel.Cells(COEX_starting_excel_row + 13, current_excel_col).Value
			COEX_HC_expense_alimony = ObjExcel.Cells(COEX_starting_excel_row + 14, current_excel_col).Value
			COEX_HC_expense_tax_dep = ObjExcel.Cells(COEX_starting_excel_row + 15, current_excel_col).Value
			COEX_HC_expense_other = ObjExcel.Cells(COEX_starting_excel_row + 16, current_excel_col).Value

			DCEX_provider = ObjExcel.Cells(DCEX_starting_excel_row, current_excel_col).Value
			DCEX_reason = left(ObjExcel.Cells(DCEX_starting_excel_row + 1, current_excel_col).Value, 1)
			DCEX_subsidy = left(ObjExcel.Cells(DCEX_starting_excel_row + 2, current_excel_col).Value, 1)
			DCEX_child_number1 = ObjExcel.Cells(DCEX_starting_excel_row + 3, current_excel_col).Value
			DCEX_child_number1_retro = ObjExcel.Cells(DCEX_starting_excel_row + 4, current_excel_col).Value
			DCEX_child_number1_pro = ObjExcel.Cells(DCEX_starting_excel_row + 5, current_excel_col).Value
			DCEX_child_number1_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 6, current_excel_col).Value, 1)
			DCEX_child_number2 = ObjExcel.Cells(DCEX_starting_excel_row + 7, current_excel_col).Value
			DCEX_child_number2_retro = ObjExcel.Cells(DCEX_starting_excel_row + 8, current_excel_col).Value
			DCEX_child_number2_pro = ObjExcel.Cells(DCEX_starting_excel_row + 9, current_excel_col).Value
			DCEX_child_number2_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 10, current_excel_col).Value, 1)
			DCEX_child_number3 = ObjExcel.Cells(DCEX_starting_excel_row + 11, current_excel_col).Value
			DCEX_child_number3_retro = ObjExcel.Cells(DCEX_starting_excel_row + 12, current_excel_col).Value
			DCEX_child_number3_pro = ObjExcel.Cells(DCEX_starting_excel_row + 13, current_excel_col).Value
			DCEX_child_number3_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 14, current_excel_col).Value, 1)
			DCEX_child_number4 = ObjExcel.Cells(DCEX_starting_excel_row + 15, current_excel_col).Value
			DCEX_child_number4_retro = ObjExcel.Cells(DCEX_starting_excel_row + 16, current_excel_col).Value
			DCEX_child_number4_pro = ObjExcel.Cells(DCEX_starting_excel_row + 17, current_excel_col).Value
			DCEX_child_number4_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 18, current_excel_col).Value, 1)
			DCEX_child_number5 = ObjExcel.Cells(DCEX_starting_excel_row + 19, current_excel_col).Value
			DCEX_child_number5_retro = ObjExcel.Cells(DCEX_starting_excel_row + 20, current_excel_col).Value
			DCEX_child_number5_pro = ObjExcel.Cells(DCEX_starting_excel_row + 21, current_excel_col).Value
			DCEX_child_number5_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 22, current_excel_col).Value, 1)
			DCEX_child_number6 = ObjExcel.Cells(DCEX_starting_excel_row + 23, current_excel_col).Value
			DCEX_child_number6_retro = ObjExcel.Cells(DCEX_starting_excel_row + 24, current_excel_col).Value
			DCEX_child_number6_pro = ObjExcel.Cells(DCEX_starting_excel_row + 25, current_excel_col).Value
			DCEX_child_number6_ver = left(ObjExcel.Cells(DCEX_starting_excel_row + 26, current_excel_col).Value, 1)

			JOBS_1_inc_type = left(ObjExcel.Cells(JOBS_1_starting_excel_row, current_excel_col).Value, 1)
			JOBS_1_inc_verif = left(ObjExcel.Cells(JOBS_1_starting_excel_row + 1, current_excel_col).Value, 1)
			JOBS_1_employer_name = ObjExcel.Cells(JOBS_1_starting_excel_row + 2, current_excel_col).Value
			JOBS_1_inc_start = ObjExcel.Cells(JOBS_1_starting_excel_row + 3, current_excel_col).Value
			JOBS_1_pay_freq = ObjExcel.Cells(JOBS_1_starting_excel_row + 4, current_excel_col).Value
			JOBS_1_wkly_hrs = ObjExcel.Cells(JOBS_1_starting_excel_row + 5, current_excel_col).Value
			JOBS_1_hrly_wage = ObjExcel.Cells(JOBS_1_starting_excel_row + 6, current_excel_col).Value

			JOBS_2_inc_type = left(ObjExcel.Cells(JOBS_2_starting_excel_row, current_excel_col).Value, 1)
			JOBS_2_inc_verif = left(ObjExcel.Cells(JOBS_2_starting_excel_row + 1, current_excel_col).Value, 1)
			JOBS_2_employer_name = ObjExcel.Cells(JOBS_2_starting_excel_row + 2, current_excel_col).Value
			JOBS_2_inc_start = ObjExcel.Cells(JOBS_2_starting_excel_row + 3, current_excel_col).Value
			JOBS_2_pay_freq = ObjExcel.Cells(JOBS_2_starting_excel_row + 4, current_excel_col).Value
			JOBS_2_wkly_hrs = ObjExcel.Cells(JOBS_2_starting_excel_row + 5, current_excel_col).Value
			JOBS_2_hrly_wage = ObjExcel.Cells(JOBS_2_starting_excel_row + 6, current_excel_col).Value

			JOBS_3_inc_type = left(ObjExcel.Cells(JOBS_3_starting_excel_row, current_excel_col).Value, 1)
			JOBS_3_inc_verif = left(ObjExcel.Cells(JOBS_3_starting_excel_row + 1, current_excel_col).Value, 1)
			JOBS_3_employer_name = ObjExcel.Cells(JOBS_3_starting_excel_row + 2, current_excel_col).Value
			JOBS_3_inc_start = ObjExcel.Cells(JOBS_3_starting_excel_row + 3, current_excel_col).Value
			JOBS_3_pay_freq = ObjExcel.Cells(JOBS_3_starting_excel_row + 4, current_excel_col).Value
			JOBS_3_wkly_hrs = ObjExcel.Cells(JOBS_3_starting_excel_row + 5, current_excel_col).Value
			JOBS_3_hrly_wage = ObjExcel.Cells(JOBS_3_starting_excel_row + 6, current_excel_col).Value

			UNEA_1_inc_type = left(ObjExcel.Cells(UNEA_1_starting_excel_row, current_excel_col).Value, 2)
			UNEA_1_inc_verif = left(ObjExcel.Cells(UNEA_1_starting_excel_row + 1, current_excel_col).Value, 1)
			UNEA_1_claim_suffix = ObjExcel.Cells(UNEA_1_starting_excel_row + 2, current_excel_col).Value
			UNEA_1_start_date = ObjExcel.Cells(UNEA_1_starting_excel_row + 3, current_excel_col).Value
			UNEA_1_pay_freq = ObjExcel.Cells(UNEA_1_starting_excel_row + 4, current_excel_col).Value
			UNEA_1_inc_amount = ObjExcel.Cells(UNEA_1_starting_excel_row + 5, current_excel_col).Value

			UNEA_2_inc_type = left(ObjExcel.Cells(UNEA_2_starting_excel_row, current_excel_col).Value, 2)
			UNEA_2_inc_verif = left(ObjExcel.Cells(UNEA_2_starting_excel_row + 1, current_excel_col).Value, 1)
			UNEA_2_claim_suffix = ObjExcel.Cells(UNEA_2_starting_excel_row + 2, current_excel_col).Value
			UNEA_2_start_date = ObjExcel.Cells(UNEA_2_starting_excel_row + 3, current_excel_col).Value
			UNEA_2_pay_freq = ObjExcel.Cells(UNEA_2_starting_excel_row + 4, current_excel_col).Value
			UNEA_2_inc_amount = ObjExcel.Cells(UNEA_2_starting_excel_row + 5, current_excel_col).Value

			UNEA_3_inc_type = left(ObjExcel.Cells(UNEA_3_starting_excel_row, current_excel_col).Value, 2)
			UNEA_3_inc_verif = left(ObjExcel.Cells(UNEA_3_starting_excel_row + 1, current_excel_col).Value, 1)
			UNEA_3_claim_suffix = ObjExcel.Cells(UNEA_3_starting_excel_row + 2, current_excel_col).Value
			UNEA_3_start_date = ObjExcel.Cells(UNEA_3_starting_excel_row + 3, current_excel_col).Value
			UNEA_3_pay_freq = ObjExcel.Cells(UNEA_3_starting_excel_row + 4, current_excel_col).Value
			UNEA_3_inc_amount = ObjExcel.Cells(UNEA_3_starting_excel_row + 5, current_excel_col).Value

			WKEX_program = objExcel.Cells(WKEX_starting_excel_row, current_excel_col).Value
			WKEX_fed_tax_retro = objExcel.Cells(WKEX_starting_excel_row + 1, current_excel_col).Value
			WKEX_fed_tax_prosp = objExcel.Cells(WKEX_starting_excel_row + 2, current_excel_col).Value
			WKEX_fed_tax_verif = left(objExcel.Cells(WKEX_starting_excel_row + 3, current_excel_col).Value, 1)
			WKEX_state_tax_retro = objExcel.Cells(WKEX_starting_excel_row + 4, current_excel_col).Value
			WKEX_state_tax_prosp = objExcel.Cells(WKEX_starting_excel_row + 5, current_excel_col).Value
			WKEX_state_tax_verif = left(objExcel.Cells(WKEX_starting_excel_row + 6, current_excel_col).Value, 1)
			WKEX_fica_retro = objExcel.Cells(WKEX_starting_excel_row + 7, current_excel_col).Value
			WKEX_fica_prosp = objExcel.Cells(WKEX_starting_excel_row + 8, current_excel_col).Value
			WKEX_fica_verif = left(objExcel.Cells(WKEX_starting_excel_row + 9, current_excel_col).Value, 1)
			WKEX_tran_retro = objExcel.Cells(WKEX_starting_excel_row + 10, current_excel_col).Value
			WKEX_tran_prosp = objExcel.Cells(WKEX_starting_excel_row + 11, current_excel_col).Value
			WKEX_tran_verif = left(objExcel.Cells(WKEX_starting_excel_row + 12, current_excel_col).Value, 1)
			WKEX_tran_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 13, current_excel_col).Value
			WKEX_meals_retro = objExcel.Cells(WKEX_starting_excel_row + 14, current_excel_col).Value
			WKEX_meals_prosp = objExcel.Cells(WKEX_starting_excel_row + 15, current_excel_col).Value
			WKEX_meals_verif = left(objExcel.Cells(WKEX_starting_excel_row + 16, current_excel_col).Value, 1)
			WKEX_meals_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 17, current_excel_col).Value
			WKEX_uniforms_retro = objExcel.Cells(WKEX_starting_excel_row + 18, current_excel_col).Value
			WKEX_uniforms_prosp = objExcel.Cells(WKEX_starting_excel_row + 19, current_excel_col).Value
			WKEX_uniforms_verif = left(objExcel.Cells(WKEX_starting_excel_row + 20, current_excel_col).Value, 1)
			WKEX_uniforms_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 21, current_excel_col).Value
			WKEX_tools_retro = objExcel.Cells(WKEX_starting_excel_row + 22, current_excel_col).Value
			WKEX_tools_prosp = objExcel.Cells(WKEX_starting_excel_row + 23, current_excel_col).Value
			WKEX_tools_verif = left(objExcel.Cells(WKEX_starting_excel_row + 24, current_excel_col).Value, 1)
			WKEX_tools_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 25, current_excel_col).Value
			WKEX_dues_retro = objExcel.Cells(WKEX_starting_excel_row + 26, current_excel_col).Value
			WKEX_dues_prosp = objExcel.Cells(WKEX_starting_excel_row + 27, current_excel_col).Value
			WKEX_dues_verif = left(objExcel.Cells(WKEX_starting_excel_row + 28, current_excel_col).Value, 1)
			WKEX_dues_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 29, current_excel_col).Value
			WKEX_othr_retro = objExcel.Cells(WKEX_starting_excel_row + 30, current_excel_col).Value
			WKEX_othr_prosp = objExcel.Cells(WKEX_starting_excel_row + 31, current_excel_col).Value
			WKEX_othr_verif = left(objExcel.Cells(WKEX_starting_excel_row + 32, current_excel_col).Value, 1)
			WKEX_othr_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 33, current_excel_col).Value
			WKEX_HC_Exp_Fed_Tax = objExcel.Cells(WKEX_starting_excel_row + 34, current_excel_col).Value
			WKEX_HC_Exp_State_Tax = objExcel.Cells(WKEX_starting_excel_row + 35, current_excel_col).Value
			WKEX_HC_Exp_FICA = objExcel.Cells(WKEX_starting_excel_row + 36, current_excel_col).Value
			WKEX_HC_Exp_Tran = objExcel.Cells(WKEX_starting_excel_row + 37, current_excel_col).Value
			WKEX_HC_Exp_Tran_imp_rel = objExcel.Cells(WKEX_starting_excel_row + 38, current_excel_col).Value
			WKEX_HC_Exp_Meals = objExcel.Cells(WKEX_starting_excel_row + 39, current_excel_col).Value
			WKEX_HC_Exp_Meals_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 40, current_excel_col).Value
			WKEX_HC_Exp_Uniforms = objExcel.Cells(WKEX_starting_excel_row + 41, current_excel_col).Value
			WKEX_HC_Exp_Uniforms_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 42, current_excel_col).Value
			WKEX_HC_Exp_Tools = objExcel.Cells(WKEX_starting_excel_row + 43, current_excel_col).Value
			WKEX_HC_Exp_Tools_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 44, current_excel_col).Value
			WKEX_HC_Exp_Dues = objExcel.Cells(WKEX_starting_excel_row + 45, current_excel_col).Value
			WKEX_HC_Exp_Dues_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 46, current_excel_col).Value
			WKEX_HC_Exp_Othr = objExcel.Cells(WKEX_starting_excel_row + 47, current_excel_col).Value
			WKEX_HC_Exp_Othr_Imp_Rel = objExcel.Cells(WKEX_starting_excel_row + 48, current_excel_col).Value

			'---Below are all the panels that need to be updated for each benefit month.
			'BUSI
			If BUSI_type <> "" then
				call write_panel_to_MAXIS_BUSI(busi_type, busi_start_date, busi_end_date, busi_cash_total_retro, busi_cash_total_prosp, busi_cash_total_ver, busi_IV_total_prosp, busi_IV_total_ver, busi_snap_total_retro, busi_snap_total_prosp, busi_snap_total_ver, busi_hc_total_prosp_a, busi_hc_total_ver_a, busi_hc_total_prosp_b, busi_hc_total_ver_b, busi_cash_exp_retro, busi_cash_exp_prosp, busi_cash_exp_ver, busi_IV_exp_prosp, busi_IV_exp_ver, busi_snap_exp_retro, busi_snap_exp_prosp, busi_snap_exp_ver, busi_hc_exp_prosp_a, busi_hc_exp_ver_a, busi_hc_exp_prosp_b, busi_hc_exp_ver_b, busi_retro_hours, busi_prosp_hours, busi_hc_total_est_a, busi_hc_total_est_b, busi_hc_exp_est_a, busi_hc_exp_est_b, busi_hc_hours_est)
				STATS_manualtime = STATS_manualtime + 60
			END IF
			'COEX
			IF COEX_support_retro <> "" AND _
				COEX_support_prosp <> "" AND _
				COEX_support_verif <> "" AND _
				COEX_alimony_retro <> "" AND _
				COEX_alimony_prosp <> "" AND _
				COEX_alimony_verif <> "" AND _
				COEX_tax_dep_retro <> "" AND _
				COEX_tax_dep_prosp <> "" AND _
				COEX_tax_dep_verif <> "" AND _
				COEX_other_retro <> "" AND _
				COEX_other_prosp <> "" AND _
				COEX_other_verif <> "" AND _
				COEX_change_in_circumstances <> "" AND _
				COEX_HC_expense_support <> "" AND _
				COEX_HC_expense_alimony <> "" AND _
				COEX_HC_expense_tax_dep <> "" AND _
				COEX_HC_expense_other <> "" THEN
					CALL write_panel_to_MAXIS_COEX(COEX_support_retro, COEX_support_prosp, COEX_support_verif, COEX_alimony_retro, COEX_alimony_prosp, COEX_alimony_verif, COEX_tax_dep_retro, COEX_tax_dep_prosp, COEX_tax_dep_verif, COEX_other_retro, COEX_other_prosp, COEX_other_verif, COEX_change_in_circumstances, COEX_HC_expense_support, COEX_HC_expense_alimony, COEX_HC_expense_tax_dep, COEX_HC_expense_other)
					STATS_manualtime = STATS_manualtime + 20
			END IF
			'DCEX
			If DCEX_provider <> "" then
				call write_panel_to_MAXIS_DCEX(DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro)
				STATS_manualtime = STATS_manualtime + 20
			END IF
			'JOBS
			IF JOBS_1_inc_type <> "" THEN
				call write_panel_to_MAXIS_JOBS("01", JOBS_1_inc_type, JOBS_1_inc_verif, JOBS_1_employer_name, JOBS_1_inc_start, JOBS_1_wkly_hrs, JOBS_1_hrly_wage, JOBS_1_pay_freq)
				STATS_manualtime = STATS_manualtime + 50
			END IF
			If JOBS_2_inc_type <> "" then
				call write_panel_to_MAXIS_JOBS("02", JOBS_2_inc_type, JOBS_2_inc_verif, JOBS_2_employer_name, JOBS_2_inc_start, JOBS_2_wkly_hrs, JOBS_2_hrly_wage, JOBS_2_pay_freq)
				STATS_manualtime = STATS_manualtime + 50
			END IF
			If JOBS_3_inc_type <> "" then
				call write_panel_to_MAXIS_JOBS("03", JOBS_3_inc_type, JOBS_3_inc_verif, JOBS_3_employer_name, JOBS_3_inc_start, JOBS_3_wkly_hrs, JOBS_3_hrly_wage, JOBS_3_pay_freq)
				STATS_manualtime = STATS_manualtime + 50
			END IF
			'UNEA
			If UNEA_1_inc_type <> "" then
				call write_panel_to_MAXIS_UNEA("01", UNEA_1_inc_type, UNEA_1_inc_verif, UNEA_1_claim_suffix, UNEA_1_start_date, UNEA_1_pay_freq, UNEA_1_inc_amount, SSN_first, SSN_mid, SSN_last)
				STATS_manualtime = STATS_manualtime + 33
			END IF
			If UNEA_2_inc_type <> "" then
				call write_panel_to_MAXIS_UNEA("02", UNEA_2_inc_type, UNEA_2_inc_verif, UNEA_2_claim_suffix, UNEA_2_start_date, UNEA_2_pay_freq, UNEA_2_inc_amount, SSN_first, SSN_mid, SSN_last)
				STATS_manualtime = STATS_manualtime + 33
			END IF
			If UNEA_3_inc_type <> "" then
				call write_panel_to_MAXIS_UNEA("03", UNEA_3_inc_type, UNEA_3_inc_verif, UNEA_3_claim_suffix, UNEA_3_start_date, UNEA_3_pay_freq, UNEA_3_inc_amount, SSN_first, SSN_mid, SSN_last)
				STATS_manualtime = STATS_manualtime + 33
			END IF
			'WKEX
			IF WKEX_program <> "" THEN
				CALL write_panel_to_MAXIS_WKEX(WKEX_program, WKEX_fed_tax_retro, WKEX_fed_tax_prosp, WKEX_fed_tax_verif, WKEX_state_tax_retro, WKEX_state_tax_prosp, WKEX_state_tax_verif, WKEX_fica_retro, WKEX_fica_prosp, WKEX_fica_verif, WKEX_tran_retro, WKEX_tran_prosp, WKEX_tran_verif, WKEX_tran_imp_rel, WKEX_meals_retro, WKEX_meals_prosp, WKEX_meals_verif, WKEX_meals_imp_rel, WKEX_uniforms_retro, WKEX_uniforms_prosp, WKEX_uniforms_verif, WKEX_uniforms_imp_rel, WKEX_tools_retro, WKEX_tools_prosp, WKEX_tools_verif, WKEX_tools_imp_rel, WKEX_dues_retro, WKEX_dues_prosp, WKEX_dues_verif, WKEX_dues_imp_rel, WKEX_othr_retro, WKEX_othr_prosp, WKEX_othr_verif, WKEX_othr_imp_rel, WKEX_HC_Exp_Fed_Tax, WKEX_HC_Exp_State_Tax, WKEX_HC_Exp_FICA, WKEX_HC_Exp_Tran, WKEX_HC_Exp_Tran_imp_rel, WKEX_HC_Exp_Meals, WKEX_HC_Exp_Meals_Imp_Rel, WKEX_HC_Exp_Uniforms, WKEX_HC_Exp_Uniforms_Imp_Rel, WKEX_HC_Exp_Tools, WKEX_HC_Exp_Tools_Imp_Rel, WKEX_HC_Exp_Dues, WKEX_HC_Exp_Dues_Imp_Rel, WKEX_HC_Exp_Othr, WKEX_HC_Exp_Othr_Imp_Rel)
				STATS_manualtime = STATS_manualtime + 20
			END IF
		NEXT
	' LOOP UNTIL benefit_month = future_month
	Next
	'Gets back to self
	back_to_self
	MAXIS_footer_month = original_MAXIS_footer_month
	MAXIS_footer_year = original_MAXIS_footer_year

Next

'Ends here if the user selected to leave cases in PND2 status
If approve_case_dropdown = "no, but enter all STAT panels needed to approve" then
	If XFER_check = checked then call transfer_cases(workers_to_XFER_cases_to, case_number_array)
	script_end_procedure("Success! Cases made, STAT panels added, and left in PND2 status, per your request.")
End if

'========================================================================APPROVAL========================================================================
FOR EACH MAXIS_case_number IN case_number_array
	back_to_SELF
	EMWriteScreen MAXIS_footer_month, 20, 43
	EMWriteScreen MAXIS_footer_year, 20, 46
	transmit

	appl_date_month = MAXIS_footer_month
	appl_date_year = MAXIS_footer_year

	If cash_application = True then
		'=====DETERMINING CASH PROGRAM =========
		'This scans CASE CURR to find what type of cash program to approve.
		call navigate_to_MAXIS_screen("case", "curr")
		DO
			EMReadScreen cash_type, 4, 10, 3
			cash_type = trim(cash_type)
			IF cash_type = "FS" OR cash_type = "" THEN
				PF3
				EMWriteScreen "CURR", 20, 70
				transmit
			END IF
		LOOP UNTIL (cash_type <> "" AND cash_type <> "FS")

		'========= MFIP Approval section ==============
		If cash_type = "MFIP" then
			DO
				back_to_SELF
				EMWriteScreen "ELIG", 16, 43
				EMWriteScreen MAXIS_case_number, 18, 43
				EMWriteScreen appl_date_month, 20, 43
				EMWriteScreen appl_date_year, 20, 46
				EMWriteScreen "MFIP", 21, 70
				'========== This TRANSMIT sends the case to the MFPR screen ==========
				transmit
				EMReadScreen no_version, 10, 24, 2
			LOOP UNTIL no_version <> "NO VERSION"
			EMReadScreen is_case_approved, 10, 3, 3
			IF is_case_approved <> "UNAPPROVED" THEN
				back_to_SELF
			ELSE
				Call write_value_and_transmit("MFSM", 20, 71)
				Call write_value_and_transmit("APP", 20, 71)
				STATS_manualtime = STATS_manualtime + 60    'adding manualtime for approval processing

                Do
                    EMReadscreen send_HRF, 3, 11, 50
                    If send_HRF = "HRF" then Call write_value_and_transmit("Y", 12, 54)
                Loop until send_HRF <> "HRF"
                row = 13
                Do
                    EMReadscreen REI_issue, 1, row, 60
                    If REI_issue = "_" then EmWriteScreen "Y", row, 60
                    row = row + 1
                Loop until REI_issue <> "_"
                'transmit

				DO
					transmit
					EMReadScreen not_allowed, 11, 24, 18
					EMReadScreen locked_by_background, 6, 24, 19

					MFIP_rei_screen = ""
					CALL find_variable("(Y/", MFIP_rei_screen, 1)
					IF MFIP_rei_screen = "N" THEN
						EMSendKey "Y"
						transmit
					END IF

					row = 1					'This is looking for if there are more months listed that need to be scrolled through to review.
					col = 1
					EMSearch "More: +", row, col
					If row <> 0 then
                        PF8
					    EMReadScreen package_approved, 8, 4, 39
                    Else
                        EMReadScreen package_approved, 8, 4, 39
                    End if
				LOOP Until package_approved = "approved"
				transmit
				'======= This handles the WF1 referral =========
                EMReadScreen work_screen_check, 4, 2, 51
					IF work_screen_check = "WORK" Then
						work_row = 7
						DO
							EMReadScreen WORK_ref_nbr, 2, work_row, 3
							EMWriteScreen "x", work_row, 47
							work_row = work_row + 1
						LOOP UNTIL WORK_ref_nbr = "  "
					transmit
						DO 'Pulling up the ES provider screen, and choosing the first option for each member
						EMReadScreen ES_provider_screen, 2, 2, 37
						EMWriteScreen "x", 5, 9
						transmit
						LOOP UNTIL ES_provider_screen <> "ES"
					transmit
					transmit
					END If
				transmit
			END IF
		END IF
		'============ DWP APPROVAL ====================
'============ DWP APPROVAL ====================
		IF cash_type = "DWP" then
			DO 'We need to put this all in a loop because MAXIS likes to have an error at the end that causes us to start over.
				'===== Needs to send a WF1 referral before approval can be done =======
				Call navigate_to_MAXIS_screen("INFC", "WORK")
				work_row = 7
				EMReadScreen referral_sent, 2, 7, 72
				IF referral_sent = "  " Then 'Makes sure the referral wasn't already sent, if it was it skips this
					DO
						EMReadScreen WORK_ref_nbr, 2, work_row, 3
						EMWriteScreen "x", work_row, 47
						work_row = work_row + 1
					LOOP UNTIL WORK_ref_nbr = "  "
					transmit
					DO 'Pulling up the ES provider screen, and choosing the first option for each member
						EMReadScreen ES_provider_screen, 2, 2, 37
						EMWriteScreen "x", 5, 9
						transmit
					LOOP UNTIL ES_provider_screen <> "ES"
					transmit 'This transmit pulls up the "do you want to send" box
					DO
						EMReadScreen referral, 8, 11, 48
					LOOP UNTIL referral = "Referral"
					EMWriteScreen "Y", 11, 64
					transmit
				END IF
				'Now it starts doing the approval
				DO
					back_to_SELF
					EMWriteScreen "ELIG", 16, 43
					EMWriteScreen MAXIS_case_number, 18, 43
					EMWriteScreen appl_date_month, 20, 43
					EMWriteScreen appl_date_year, 20, 46
					EMWriteScreen "DWP", 21, 70
					'========== This TRANSMIT sends the case to the DWPR screen ==========
					transmit
					EMReadScreen no_version, 10, 24, 2
				LOOP UNTIL no_version <> "NO VERSION"
				EMReadScreen is_case_approved, 10, 3, 3
				IF is_case_approved <> "UNAPPROVED" THEN
					back_to_SELF
				ELSE
					EMWriteScreen "DWSM", 20, 71
					transmit
					DO
					EMWriteScreen "APP", 20, 71
					STATS_manualtime = STATS_manualtime + 60    'adding manualtime for approval processing
					transmit
						EMReadScreen not_allowed, 11, 24, 18
						EMReadScreen locked_by_background, 6, 24, 19
					LOOP UNTIL not_allowed <> "NOT ALLOWED" AND locked_by_background <> "LOCKED"
					'====== Now on vendor payment screen, the script does not set up any vendoring. ======
					'====== This loop takes it through vendor screens for all months in package, and scans the screen for the various package popups  =====
					'====== and REI screens.  It will exit the loop upon finding a final approval screen. ========
					DO
						PF3 'bypasses current month vendor screen
						EMReadScreen approval_screen, 8, 15, 60 'This checks for standard DWP approval
						IF approval_screen = "approval" Then
							EMWriteScreen "Y", 16, 51 'Approve the package
							transmit
							transmit
							EXIT DO
						END IF
						'Now it needs check if an REI screen has popped up, and handle it
						row = 1
						col = 1
						EMSearch "REI", row, col
						IF row <> 0 THEN
							rei_row = 6
							rei_col = 8
							DO 'Need to find all REI options and answer them
								EMSearch "(Y/N)", rei_row, rei_col
								EMWriteScreen "N", rei_row, rei_col + 7
								rei_row = rei_row +1
							LOOP UNTIL rei_row = 1 'this exits when no more instances of y/n found
							transmit 'This should return back to the vendor screen / trigger next popup
						END IF
						'The next series of IFs are checking for further package approvals, for MFIP transition
						EMReadScreen HRF_check, 3, 11, 50 'Needs to check for the HRF popup, in the case of DWP - MFIP transition
						IF HRF_check = "HRF" THEN
							EMWriteScreen "N", 12, 54
							Transmit
						END IF
						'This checks for the final package approval if in a combined DWP / MFIP screen
						EMReadScreen combined_package_check, 7, 4, 59
						IF combined_package_check = "PACKAGE" THEN
							DO 'This gets to the last screen of the package
								EMReadScreen next_screen_check, 1, 14, 33
								IF next_screen_check = "+" THEN PF8
							LOOP UNTIL next_screen_check <> "+"
							EMWriteScreen "Y", 16, 51
							Transmit
							Transmit
							EXIT DO
						END IF
					LOOP
					'Now it needs to handle the possibility of additional WF1 referrals due to combined package
					EMReadScreen work_screen_check, 4, 2, 51
					EMReadScreen back_to_ELIG_check, 4, 2, 52
					IF back_to_ELIG_check = "ELIG" THEN EXIT DO
					IF work_screen_check = "WORK" Then
						work_row = 7
						DO
							EMReadScreen WORK_ref_nbr, 2, work_row, 3
							EMWriteScreen "x", work_row, 47
							work_row = work_row + 1
						LOOP UNTIL WORK_ref_nbr = "  "
						transmit
						DO 'Pulling up the ES provider screen, and choosing the first option for each member
						EMReadScreen ES_provider_screen, 2, 2, 37
						EMWriteScreen "x", 5, 9
						transmit
						LOOP UNTIL ES_provider_screen <> "ES"
						transmit
						transmit
						EXIT DO
					END If
					IF work_screen_check = "r   " THEN transmit
				END IF
			LOOP
		END IF
		'========= MSA Approval =======================
		IF cash_type = "MSA" Then
			DO
				back_to_SELF
				EMWriteScreen "ELIG", 16, 43
				EMWriteScreen MAXIS_case_number, 18, 43
				EMWriteScreen appl_date_month, 20, 43
				EMWriteScreen appl_date_year, 20, 46
				EMWriteScreen "MSA", 21, 70
				'========== This TRANSMIT sends the case to the MSPR screen ==========
				transmit
				EMReadScreen no_version, 10, 24, 2
			LOOP UNTIL no_version <> "NO VERSION"
			EMReadScreen is_case_approved, 10, 3, 3
			IF is_case_approved <> "UNAPPROVED" THEN
				back_to_SELF
			ELSE
				DO
					transmit
					EMReadScreen at_MSSM, 4, 3, 44
				LOOP UNTIL at_MSSM = "MSSM"

				EMWaitReady 2, 2000
				DO
					questionable_information = ""
					EMWriteScreen "1", 17, 54
					EMWriteScreen "__", 18, 54
					EMWriteScreen "APP", 20, 70
					STATS_manualtime = STATS_manualtime + 60    'adding manualtime for approval processing
					transmit

					EMReadScreen error_message, 20, 24, 2
					error_message = trim(error_message)
					CALL find_variable("REI benefits", questionable_information, 1)
				LOOP UNTIL error_message = "" OR questionable_information = "?"
				'REI'ing all MSA and looping until all MSA approval is complete.
				DO
					look_for_rein_yn = ""
					michael_bay_action_sequence = ""
					row = 1
					col = 1
					CALL find_variable("(Y/", look_for_rein_yn, 1)
					CALL find_variable("Action: ", michael_bay_action_sequence, 1)
					IF look_for_rein_yn = "N" THEN
						EMSendKey "Y"
						transmit
					END IF
					IF michael_bay_action_sequence = "_" THEN
						EMSendKey "1"
						transmit
					END IF
					IF look_for_rein_yn = "" and michael_bay_action_sequence = "" THEN
						transmit
						EXIT DO
					END IF
				LOOP
				row = 1					'This is looking for if there are more months listed that need to be scrolled through to review.
				col = 1
				EMSearch "More: +", row, col
				If row <> 0 then PF8
				EMSendKey "Y"
				transmit
			END IF
		END IF
		'================= GA Approval ===============================================
		IF cash_type = "GA" THEN
			num_of_ga_months = DateDiff("M", appl_date, date) + 2
			Dim ga_array		'Creating an array of GA approval for FIATing SNAP.
			ReDim ga_array(num_of_ga_months, 1)
			ga_months = -1
			DO
				back_to_SELF
				EMWriteScreen "FIAT", 16, 43
				EMWriteScreen MAXIS_case_number, 18, 43
				EMWriteScreen appl_date_month, 20, 43
				EMWriteScreen appl_date_year, 20, 46
				transmit
				'====Should now be on FIAT submenu
				EMReadScreen GA_version, 1, 12, 48
			LOOP UNTIL GA_version = "/"
			'THIS DO LOOP fills out FIAT menu and all Fiat screens, saves results,
			'and repeats for each month in the package until it reaches final month.
			DO
				ga_months = ga_months + 1
				DO
					EMWriteScreen "10", 4, 34
					EMWriteScreen "x", 12, 22
					transmit
					EMReadScreen gasp, 4, 3, 56
				LOOP UNTIL gasp = "GASP"
				DO
					EMWriteScreen "P", 8, 63			'We need to determine Retrospective vs. Prospective budgeting programatically. Then we need have the script pull earned and unearned income from JOBS, UNEA, BUSI, and RBIC.
					IF GA_type = "personal needs" THEN 'THIS is for using for personal needs GA in a FACI setting.  Currently no logic to assign this variable
						EMWriteScreen "5", 18, 77
					ELSE
						EMWriteScreen "1", 18, 52 'This is for community single adult cases - the default
					END IF
					EMWriteScreen "x", 19, 27
					EMWriteScreen "x", 19, 50
					EMWriteScreen "x", 19, 70
						transmit 'Takes it to case results
					EMReadScreen gacr, 4, 3, 45
				LOOP UNTIL gacr = "GACR"
				transmit
				DO
					EMReadScreen GAB1, 4, 3, 52
				LOOP UNTIL GAB1 = "GAB1"
				EMWriteScreen "GASM", 20, 70
				transmit
				DO
					EMReadScreen gasm, 4, 3, 51
				LOOP UNTIL gasm = "GASM"
				CALL find_variable("Amount To Be Paid........$", ga_benefit_amount, 9)
				ga_benefit_amount = trim(ga_benefit_amount)
				CALL find_variable("Month: ", ga_benefit_month, 5)
				ga_benefit_month = replace(ga_benefit_month, " ", "/")
				ga_array(ga_months, 0) = ga_benefit_month
				ga_array(ga_months, 1) = ga_benefit_amount
				PF3 'exiting back to GASP screen after viewing budget
				PF3 'pulls up do you want to retain this version?
				DO
					EMReadScreen FIAT_retain, 8, 13, 32
				LOOP UNTIL FIAT_retain = "(Y or N)"
				EMWriteScreen "Y", 13, 41
				transmit 'brings it back to fiat submenu if not last month, offers elig popup if last month
				DO
					EMReadScreen elig_popup, 4, 10, 53
					EMReadScreen fiat_menu, 4, 2, 46
					IF elig_popup = "ELIG" THEN 'Exiting the FIAT loop and going to ELIG
						EMWriteScreen "Y", 11, 52
						EMWriteScreen appl_date_month, 13, 37
						EMWriteScreen appl_date_year, 13, 40
						transmit
						EXIT DO
					END IF
				LOOP UNTIL fiat_menu = "FIAT"

				'Adding 1 to the elig month
				EMReadScreen elig_month, 2, 20, 54
				EMReadScreen elig_year, 2, 20, 57
				ga_bene_month = elig_month & "/01/" & elig_year
				ga_bene_month = DateAdd("M", 1, ga_bene_month)
				elig_month = DatePart ("M", ga_bene_month)
				elig_year = right(DatePart("YYYY", ga_bene_month), 2)
				IF len(elig_month) <> 2 THEN elig_month = "0" & elig_month
				EMWriteScreen elig_month, 20, 54
				EMWriteScreen elig_year, 20, 57
				transmit
				EMReadScreen elig_results, 7, 2, 31
			LOOP UNTIL elig_results = "GA Elig"
			DO 'Checking for the approval screen
				EMReadScreen elig_gasm, 6, 15, 45
			LOOP UNTIL elig_gasm = "Action"
			EMWriteScreen appl_date_month, 20, 54
			EMWriteScreen appl_date_year, 20, 57
			transmit

			EMWriteScreen "1", 15, 53
			EMWriteScreen "APP", 20, 70
			STATS_manualtime = STATS_manualtime + 60    'adding manualtime for approval processing
			transmit
			DO 'getting REI screen and selecting Y
				total_package = ""
				colon_inspection = ""
				rei_screen = ""
				CALL find_variable("Cash package ", total_package, 8)
				CALL find_variable("Action", colon_inspection, 1)
				CALL find_variable("(Y/", rei_screen, 1)
				IF rei_screen = "N" THEN
					EMSendKey "Y"
					transmit
				END IF
				IF colon_inspection = ":" THEN
					EMSendKey "1"
					transmit
				END IF
				row = 1					'This is looking for if there are more months listed that need to be scrolled through to review.
				col = 1
				EMSearch "More: +", row, col
				If row <> 0 then
					PF8
					row = 1					'This is looking for if there are more months listed that need to be scrolled through to review.
					col = 1
					EMSearch "(Y/N)", row, col
					EMWriteScreen "Y", row, col +7
					transmit
				End If
			LOOP UNTIL total_package = "approved"
			transmit

			'Need to make the script FIAT and approve SNAP with GA.
			IF SNAP_application = True THEN
				back_to_SELF
				EMWriteScreen "FIAT", 16, 43
				EMWriteScreen MAXIS_case_number, 18, 43
				EMWriteScreen appl_date_month, 20, 43
				EMWriteScreen appl_date_year, 20, 46
				transmit

				FOR i = 0 to ga_months
					EMWriteScreen left(ga_array(i, 0), 2), 20, 54
					EMWriteScreen right(ga_array(i, 0), 2), 20, 57
					transmit
					DO
						EMWriteScreen "22", 4, 34
						EMWriteScreen "X", 14, 22
						transmit
						EMReadScreen error_message, 20, 24, 2
						error_message = trim(error_message)
						EMReadScreen ffsl, 4, 3, 52
					LOOP UNTIL error_message = "" AND ffsl = "FFSL"

					EMWriteScreen "X", 6, 5
					EMWriteScreen "X", 16, 5
					EMWriteScreen "X", 17, 5
					transmit		'takes the script to FFPR
					DO
						EMReadScreen ffpr, 4, 3, 47
					LOOP UNTIL ffpr = "FFPR"
					EMWriteScreen "N", 7, 58
					EMWriteScreen "P", 7, 66
					PF3
					EMReadScreen ffpr, 4, 3, 47
					If ffpr = "FFPR" Then PF3
					DO
						EMReadScreen ffcr, 4, 3, 46
					LOOP UNTIL ffcr = "FFCR"
					PF3
					DO
						EMReadScreen ffb1, 4, 3, 51
					LOOP UNTIL ffb1 = "FFB1"
					EMWriteScreen "X", 10, 5
					transmit
					EMWriteScreen "         ", 9, 23     'clearing variable as sometimes GA gets budgetted into SNAP already.
					EMWriteScreen ga_array(i, 1), 9, 23
					DO
						transmit
						EMReadScreen FFSM, 4, 3, 53
					LOOP UNTIL FFSM = "FFSM"
					PF3
					DO
						EMReadScreen FFSL, 4, 3, 52
					LOOP UNTIL FFSL = "FFSL"
					PF3
					DO
						this_version = ""
						CALL find_variable("this FIAT ", this_version, 7)
					LOOP UNTIL this_version = "version"
					EMSendKey "Y"
					transmit
					ready_to_fiat_FS = ""
					CALL find_variable("Do you want to go to ", ready_to_fiat_FS, 4)
					IF ready_to_fiat_FS = "ELIG" THEN
						EMWriteScreen "Y", 11, 52
						EMWriteScreen appl_date_month, 13, 37
						EMWriteScreen appl_date_year, 13, 40
						transmit
						DO
							EMReadScreen FSSM, 4, 3, 54
						LOOP UNTIL FSSM = "FSSM"
						EMWriteScreen "APP", 19, 70
						STATS_manualtime = STATS_manualtime + 60    'adding manualtime for approval processing
						transmit
						CALL find_variable("THIS IS AN EXPEDITED ", expedited_status, 4)
						ups_delivery_confirmation = ""  'resetting variable
						CALL find_variable("PLEASE EXAMINE THE FOLLOWING ", ups_delivery_confirmation, 7)
						IF expedited_status = "CASE" THEN
							EMSendKey "Y"
							transmit
							EMSendKey "Y"
							transmit
							row = 1					'This is looking for if there are more months listed that need to be scrolled through to review.
							col = 1
							EMSearch "More: +", row, col
							If row <> 0 then PF8
							EMSendKey "Y"
							transmit
						END IF
						IF ups_delivery_confirmation = "PACKAGE" THEN
							row = 1					'This is looking for if there are more months listed that need to be scrolled through to review.
							col = 1
							EMSearch "More: +", row, col
							If row <> 0 then PF8
							EMSendKey "Y"
							transmit
						END IF
						transmit
					END IF
				NEXT
			END IF
		END IF
	End if
	'The script needs to FIAT GA into SNAP budget.

	If SNAP_application = True AND cash_type <> "GA" then
		DO
			back_to_SELF
			EMWriteScreen "ELIG", 16, 43
			EMWriteScreen MAXIS_case_number, 18, 43
			EMWriteScreen appl_date_month, 20, 43
			EMWriteScreen appl_date_year, 20, 46
			EMWriteScreen "FS", 21, 70
			'========== This TRANSMIT sends the case to the FSPR screen ==========
			transmit
			EMReadScreen no_version, 10, 24, 2
		LOOP UNTIL no_version <> "NO VERSION"
		EMReadScreen is_case_approved, 10, 3, 3
		IF is_case_approved <> "UNAPPROVED" THEN
			back_to_SELF
		ELSE
		'========== This TRANSMIT sends the case to the FSCR screen ==========
			transmit
		'========== Reading for EXPEDITED STATUS ==========
			EMReadScreen is_case_expedited, 9, 4, 3
		'========== This TRANSMIT sends the case to the FSB1 screen ==========
			transmit
		'========== This TRANSMIT sends the case to the FSB2 screen ==========
			transmit
		'========== This TRANSMIT sends the case to the FSSM screen ==========
			transmit
			IF is_case_expedited <> "EXPEDITED" THEN
				DO
					not_allowed = ""
					locked_by_background = ""
					EMWriteScreen "APP", 19, 70
					STATS_manualtime = STATS_manualtime + 60    'adding manualtime for approval processing
					transmit
					EMReadScreen not_allowed, 11, 24, 18
					EMReadScreen locked_by_background, 6, 24, 19
					row = 1					'This is looking for if there are more months listed that need to be scrolled through to review.
					col = 1
					EMSearch "More: +", row, col
					If row <> 0 then PF8
					row = 1
					col = 1
					EMSearch "(Y/N)  _", row, col
				LOOP UNTIL (not_allowed <> "NOT ALLOWED" AND locked_by_background <> "LOCKED") OR row <> 0
				DO
					row = 1
					col = 1
					EMSearch "Do you want to continue with the approval?", row, col
				LOOP UNTIL row <> 0
				DO
					row = 1						'This is looking for if there are more months listed that need to be scrolled through to review.
					col = 1
					EMSearch "More: +", row, col
					If row <> 0 then PF8
					row = 1
					col = 1
					EMSearch "(Y/N)  _", row, col
					IF row <> 0 THEN
						EMWriteScreen "Y", row, col + 7
					ELSE
						MsgBox "The script is struggling to find the correct space to confirm the approval. Please enter a Y in the correct space, and press OK for the script to continue." & vbCr & vbCr & "PLEASE DO NOT TRANSMIT!!"
					END IF
					transmit
					ups_delivery_confirmation = ""  'resetting variable
					CALL find_variable("Package ", ups_delivery_confirmation, 8)
				LOOP UNTIL ups_delivery_confirmation = "approved"
				transmit
			ELSE
				DO
					not_allowed = ""
					locked_by_background = ""
					EMWriteScreen "APP", 19, 70
					transmit
					EMReadScreen not_allowed, 11, 24, 18
					EMReadScreen locked_by_background, 6, 24, 19
					row = 1								'This is looking for if there are more months listed that need to be scrolled through to review.
					col = 1
					EMSearch "More: +", row, col
					If row <> 0 then PF8
					row = 1
					col = 1
					EMSearch "(Y/N)", row, col
					IF row <> 0 THEN
						emfocus
						emsendkey "<tab>"
						emsendkey "y"
						transmit
					End If
					ups_delivery_confirmation = ""  'resetting variable
					CALL find_variable("Package ", ups_delivery_confirmation, 8)
				LOOP UNTIL ups_delivery_confirmation = "approved"
				transmit
			END IF
		END IF
	End if

	If HC_application = True AND (HCRE_retro_months_input = "" OR HCRE_retro_months_input = "0") then			'IF the case is requesting retro, it should not approve. That is a scenario that needn't be auto-approved.
		'Approve HC, please.
		Do				'Need to make sure the case has come all the way through background
			call navigate_to_MAXIS_screen ("STAT", "SUMM")		'Otherwise occasionally the FIATING getts messed up with the MNSURE FIAT part
			EMReadScreen summ_check, 4, 2, 46
		Loop until summ_check = "SUMM"
		'=====THE SCRIPT NEEDS TO GET AROUND ELIG/HC RESULTS BEING STUCK IN BACKGROUND
		DO
			call navigate_to_MAXIS_screen("ELIG", "HC")
			hhmm_row = 8
			DO
				EMReadScreen no_version, 10, hhmm_row, 28
				no_version = trim(no_version)
				IF no_version = "NO VERSION" THEN hhmm_row = hhmm_row + 1
			LOOP UNTIL no_version = "" OR left(no_version, 2) = "MA"
		LOOP UNTIL left(no_version, 2) = "MA"
		'=====This part of the script makes the FIAT changes to HH members with Budg Mthd A
		hhmm_row = 8
		DO
			EMReadScreen hc_requested, 1, hhmm_row, 28
			EMReadScreen hc_status, 5, hhmm_row, 68
			IF hc_requested = "M" AND hc_status = "UNAPP" THEN
				DO						'===== This DO/LOOP is for the check to determine the case is not stuck in ELIG. If it is, it will not let you FIAT Elig Standard.
					EMWriteScreen "X", hhmm_row, 26
					transmit				'===== Navigates to BSUM for the HH member

					'The script now reads the budget method for each month in the period.
					EMReadScreen budg_mthd_mo1, 1, 13, 21
					EMReadScreen budg_mthd_mo2, 1, 13, 32
					EMReadScreen budg_mthd_mo3, 1, 13, 43
					EMReadScreen budg_mthd_mo4, 1, 13, 54
					EMReadScreen budg_mthd_mo5, 1, 13, 65
					EMReadScreen budg_mthd_mo6, 1, 13, 76

					'If ALL 6 budget months are not method A, the script backs out of the DO/LOOP and begins searching for the next HC applicant.
					IF (budg_mthd_mo1 <> "A") AND (budg_mthd_mo2 <> "A") AND (budg_mthd_mo3 <> "A") AND (budg_mthd_mo4 <> "A") AND (budg_mthd_mo5 <> "A") AND (budg_mthd_mo6 <> "A") THEN
						PF3
						EXIT DO

					'If the script finds any budget method A months in the ELIG period, it will FIAT the results to accommodate the appropriate Eligibility Standard...
					'=====THE FOLLOWING IS DOCUMENTATION ON THE FIATING=====
					'	When the script EMWriteScreens "X" on row 7, it is selecting the PERSON TEST for that month. The script needs to do this to FIAT "PASSED" on the MNSure test.
					'	When the script EMWriteScreens "X" on row 9, it is selective the BUDGET for that month.
					'	The script also needs to change the eligibility standard, but that is dependent on the prevailing eligibility standard. This is the variable "mo#_elig_type"
					ELSEIF (budg_mthd_mo1 = "A") OR (budg_mthd_mo2 = "A") OR (budg_mthd_mo3 = "A") OR (budg_mthd_mo4 = "A") OR (budg_mthd_mo5 = "A") OR (budg_mthd_mo6 = "A") THEN
							PF9
							DO
								EMReadScreen fiat_reason, 4, 10, 20		'=====The script gets stuck in ELIG background...it's running faster than the training region will allow.
							LOOP UNTIL fiat_reason = "FIAT"
							EMWriteScreen "05", 11, 26
							transmit
						IF budg_mthd_mo1 = "A" THEN
							EMWriteScreen "X", 7, 17
							EMWriteScreen "X", 9, 21
							EMReadScreen mo1_elig_type, 2, 12, 17
							IF (mo1_elig_type = "AX" OR mo1_elig_type = "AA" OR mo1_elig_type = "CX") THEN EMWriteScreen "J", 12, 22
							IF (mo1_elig_type = "PX" OR mo1_elig_type = "PC") THEN EMWriteScreen "T", 12, 22
							IF (mo1_elig_type = "CK") THEN EMWriteScreen "K", 12, 22
							IF mo1_elig_type = "CB" THEN EMWriteScreen "I", 12, 22
						END IF
						IF budg_mthd_mo2 = "A" THEN
							EMWriteScreen "X", 7, 28
							EMWriteScreen "X", 9, 32
							EMReadScreen mo2_elig_type, 2, 12, 28
							IF (mo2_elig_type = "AX" OR mo2_elig_type = "AA" OR mo2_elig_type = "CX") THEN EMWriteScreen "J", 12, 33
							IF (mo2_elig_type = "PX" OR mo2_elig_type = "PC") THEN EMWriteScreen "T", 12, 33
							IF (mo2_elig_type = "CK") THEN EMWriteScreen "K", 12, 33
							IF mo2_elig_type = "CB" THEN EMWriteScreen "I", 12, 33
						END IF
						IF budg_mthd_mo3 = "A" THEN
							EMWriteScreen "X", 7, 39
							EMWriteScreen "X", 9, 43
							EMReadScreen mo3_elig_type, 2, 12, 39
							IF (mo3_elig_type = "AX" OR mo3_elig_type = "AA" OR mo3_elig_type = "CX") THEN EMWriteScreen "J", 12, 44
							IF (mo3_elig_type = "PX" OR mo3_elig_type = "PC") THEN EMWriteScreen "T", 12, 44
							IF (mo3_elig_type = "CK") THEN EMWriteScreen "K", 12, 44
							IF mo3_elig_type = "CB" THEN EMWriteScreen "I", 12, 44
						END IF
						IF budg_mthd_mo4 = "A" THEN
							EMWriteScreen "X", 7, 50
							EMWriteScreen "X", 9, 54
							EMReadScreen mo4_elig_type, 2, 12, 50
							IF (mo4_elig_type = "AX" OR mo4_elig_type = "AA" OR mo4_elig_type = "CX") THEN EMWriteScreen "J", 12, 55
							IF (mo4_elig_type = "PX" OR mo4_elig_type = "PC") THEN EMWriteScreen "T", 12, 55
							IF (mo4_elig_type = "CK") THEN EMWriteScreen "K", 12, 55
							IF mo4_elig_type = "CB" THEN EMWriteScreen "I", 12, 55
						END IF
						IF budg_mthd_mo5 = "A" THEN
							EMWriteScreen "X", 7, 61
							EMWriteScreen "X", 9, 65
							EMReadScreen mo5_elig_type, 2, 12, 61
							IF (mo5_elig_type = "AX" OR mo5_elig_type = "AA" OR mo5_elig_type = "CX") THEN EMWriteScreen "J", 12, 66
							IF (mo5_elig_type = "PX" OR mo5_elig_type = "PC") THEN EMWriteScreen "T", 12, 66
							IF (mo5_elig_type = "CK") THEN EMWriteScreen "K", 12, 66
							IF mo5_elig_type = "CB" THEN EMWriteScreen "I", 12, 66
						END IF
						IF budg_mthd_mo6 = "A" THEN
							EMWriteScreen "X", 7, 72
							EMWriteScreen "X", 9, 76
							EMReadScreen mo6_elig_type, 2, 12, 72
							IF (mo6_elig_type = "AX" OR mo6_elig_type = "AA" OR mo6_elig_type = "CX") THEN EMWriteScreen "J", 12, 77
							IF (mo6_elig_type = "PX" OR mo6_elig_type = "PC") THEN EMWriteScreen "T", 12, 77
							IF (mo6_elig_type = "CK") THEN EMWriteScreen "K", 12, 77
							IF mo6_elig_type = "CB" THEN EMWriteScreen "I", 12, 77
						END IF
						transmit		'IF Budg Mthd A, transmit to navigate to MAPT & CBUD
						DO
							EMReadScreen back_to_bsum, 4, 3, 57
							IF back_to_BSUM <> "BSUM" THEN
								EMReadScreen mapt, 4, 3, 51
								EMReadScreen cbud, 4, 3, 54
								EMReadScreen abud, 4, 3, 47
								IF mapt = "MAPT" THEN
									EMWriteScreen "PASSED", 8, 46		'=====Passes MNSure test
									transmit
									transmit
								END IF
								IF cbud = "CBUD" OR abud = "ABUD" THEN transmit		'======Getting out of the budget window for that month.
							END IF
						LOOP UNTIL back_to_bsum = "BSUM"
						'---Now the script needs to determine if the case passes income test for cert period
						EMReadScreen clt_ref_num, 2, 5, 16
						EMWriteScreen "X", 18, 3
						Transmit
						For spddn_row = 6 to 18
							EMReadScreen spenddown_type, 12, spddn_row, 39
							EMReadScreen listed_clt, 2, spddn_row, 6
							IF spenddown_type <> "NO SPENDDOWN" AND listed_clt = clt_ref_num Then
								EMWriteScreen "X", spddn_row, 3
								transmit
								EMWriteScreen " ", 5, 14
								Transmit
								Transmit
								PF3
								Exit For
							End If
							If spddn_row = 18 then PF3
						Next
						EMWriteScreen "X", 18, 34
						transmit		'---Opens the Cert Period Amount sub-menu
						DO
							EMReadScreen at_cert_period_screen, 13, 5, 13
						LOOP UNTIL at_cert_period_screen = "Certification"
						EMReadScreen excess_income, 5, 9, 39
						PF3
						IF excess_income = " 0.00" THEN
							'---The script will go into the MAPT for all appropriate months and pass Income - Budget Period.
							EMWriteScreen "X", 7, 17
							EMWriteScreen "X", 7, 28
							EMWriteScreen "X", 7, 39
							EMWriteScreen "X", 7, 50
							EMWriteScreen "X", 7, 61
							EMWriteScreen "X", 7, 72
							transmit
							DO
								EMReadScreen mapt_check, 4, 3, 51
								EMReadScreen mobl_check, 4, 3, 49
								IF mapt_check = "MAPT" Then
									EMWriteScreen "PASSED", 6, 46
									EMWriteScreen "PASSED", 9, 46
									EMWriteScreen "PASSED", 10, 46
									transmit
									EMReadScreen back_to_BSUM, 4, 3, 57
								ElseIf mobl_check = "MOBL" then
									For spdwn_row = 6 to 18
										EMReadScreen spdwn_check, 9, spdwn_row, 39
										IF spdwn_check = "SPENDDOWN" Then
											EMWriteScreen "X", spdwn_row, 3
											transmit
											EMWriteScreen "_", 5, 14
											transmit
											transmit
										End If
									Next
									PF3
									EMReadScreen back_to_BSUM, 4, 3, 57
								End If
							LOOP UNTIL back_to_BSUM = "BSUM"
						END IF
					END IF

					EMReadScreen cannot_fiat, 10, 24, 6
					cannot_fiat = trim(cannot_fiat)
					IF cannot_fiat <> "" THEN 		'===== IF the case is stuck in ELIG, it will not allow you to change the ELIG standard to the ACA-appropriate standard.
						PF10							'===== The script OOPS's the FIAT and backs out. It will reread and re-transmit the FIAT'd elig information.
						PF3
						EMWriteScreen "WAIT", 20, 71
						EMWaitReady 2, 2000
						EMWriteScreen "____", 20, 71
					ELSE
						PF3
					END IF

				LOOP UNTIL cannot_fiat = ""

			END IF

			hhmm_row = hhmm_row + 1

		LOOP UNTIL hc_requested = " "			'===== Loops until there are no more HC versions to review

		If second_span = True Then				'If the application date is more than 5 months ago, a second HC span needs to be approved
			EMWriteScreen six_month_month, 20, 56
			EMWriteScreen six_month_year, 20, 59
			transmit
			DO
				call navigate_to_MAXIS_screen("ELIG", "HC")
				hhmm_row = 8
				DO
					EMReadScreen no_version, 10, hhmm_row, 28
					no_version = trim(no_version)
					IF no_version = "NO VERSION" THEN hhmm_row = hhmm_row + 1
				LOOP UNTIL no_version = "" OR left(no_version, 2) = "MA"
			LOOP UNTIL left(no_version, 2) = "MA"
			'=====This part of the script makes the FIAT changes to HH members with Budg Mthd A
			hhmm_row = 8
			EMWriteScreen six_month_month, 20, 56
			EMWriteScreen six_month_year, 20, 59
			transmit
			DO
				EMReadScreen hc_requested, 1, hhmm_row, 28
				EMReadScreen hc_status, 5, hhmm_row, 68
				IF hc_requested = "M" AND hc_status = "UNAPP" THEN
					DO						'===== This DO/LOOP is for the check to determine the case is not stuck in ELIG. If it is, it will not let you FIAT Elig Standard.
						EMWriteScreen "X", hhmm_row, 26
						transmit				'===== Navigates to BSUM for the HH member

						'The script now reads the budget method for each month in the period.
						EMReadScreen budg_mthd_mo1, 1, 13, 21
						EMReadScreen budg_mthd_mo2, 1, 13, 32
						EMReadScreen budg_mthd_mo3, 1, 13, 43
						EMReadScreen budg_mthd_mo4, 1, 13, 54
						EMReadScreen budg_mthd_mo5, 1, 13, 65
						EMReadScreen budg_mthd_mo6, 1, 13, 76

						'If ALL 6 budget months are not method A, the script backs out of the DO/LOOP and begins searching for the next HC applicant.
						IF (budg_mthd_mo1 <> "A") AND (budg_mthd_mo2 <> "A") AND (budg_mthd_mo3 <> "A") AND (budg_mthd_mo4 <> "A") AND (budg_mthd_mo5 <> "A") AND (budg_mthd_mo6 <> "A") THEN
							PF3
							EXIT DO

						'If the script finds any budget method A months in the ELIG period, it will FIAT the results to accommodate the appropriate Eligibility Standard...
						'=====THE FOLLOWING IS DOCUMENTATION ON THE FIATING=====
						'	When the script EMWriteScreens "X" on row 7, it is selecting the PERSON TEST for that month. The script needs to do this to FIAT "PASSED" on the MNSure test.
						'	When the script EMWriteScreens "X" on row 9, it is selective the BUDGET for that month.
						'	The script also needs to change the eligibility standard, but that is dependent on the prevailing eligibility standard. This is the variable "mo#_elig_type"
						ELSEIF (budg_mthd_mo1 = "A") OR (budg_mthd_mo2 = "A") OR (budg_mthd_mo3 = "A") OR (budg_mthd_mo4 = "A") OR (budg_mthd_mo5 = "A") OR (budg_mthd_mo6 = "A") THEN
								PF9
								DO
									EMReadScreen fiat_reason, 4, 10, 20		'=====The script gets stuck in ELIG background...it's running faster than the training region will allow.
								LOOP UNTIL fiat_reason = "FIAT"
								EMWriteScreen "05", 11, 26
								transmit
							IF budg_mthd_mo1 = "A" THEN
								EMWriteScreen "X", 7, 17
								EMWriteScreen "X", 9, 21
								EMReadScreen mo1_elig_type, 2, 12, 17
								IF (mo1_elig_type = "AX" OR mo1_elig_type = "AA" OR mo1_elig_type = "CX") THEN EMWriteScreen "J", 12, 22
								IF (mo1_elig_type = "PX" OR mo1_elig_type = "PC") THEN EMWriteScreen "T", 12, 22
								IF (mo1_elig_type = "CK") THEN EMWriteScreen "K", 12, 22
								IF mo1_elig_type = "CB" THEN EMWriteScreen "I", 12, 22
							END IF
							IF budg_mthd_mo2 = "A" THEN
								EMWriteScreen "X", 7, 28
								EMWriteScreen "X", 9, 32
								EMReadScreen mo2_elig_type, 2, 12, 28
								IF (mo2_elig_type = "AX" OR mo2_elig_type = "AA" OR mo2_elig_type = "CX") THEN EMWriteScreen "J", 12, 33
								IF (mo2_elig_type = "PX" OR mo2_elig_type = "PC") THEN EMWriteScreen "T", 12, 33
								IF (mo2_elig_type = "CK") THEN EMWriteScreen "K", 12, 33
								IF mo2_elig_type = "CB" THEN EMWriteScreen "I", 12, 33
							END IF
							IF budg_mthd_mo3 = "A" THEN
								EMWriteScreen "X", 7, 39
								EMWriteScreen "X", 9, 43
								EMReadScreen mo3_elig_type, 2, 12, 39
								IF (mo3_elig_type = "AX" OR mo3_elig_type = "AA" OR mo3_elig_type = "CX") THEN EMWriteScreen "J", 12, 44
								IF (mo3_elig_type = "PX" OR mo3_elig_type = "PC") THEN EMWriteScreen "T", 12, 44
								IF (mo3_elig_type = "CK") THEN EMWriteScreen "K", 12, 44
								IF mo3_elig_type = "CB" THEN EMWriteScreen "I", 12, 44
							END IF
							IF budg_mthd_mo4 = "A" THEN
								EMWriteScreen "X", 7, 50
								EMWriteScreen "X", 9, 54
								EMReadScreen mo4_elig_type, 2, 12, 50
								IF (mo4_elig_type = "AX" OR mo4_elig_type = "AA" OR mo4_elig_type = "CX") THEN EMWriteScreen "J", 12, 55
								IF (mo4_elig_type = "PX" OR mo4_elig_type = "PC") THEN EMWriteScreen "T", 12, 55
								IF (mo4_elig_type = "CK") THEN EMWriteScreen "K", 12, 55
								IF mo4_elig_type = "CB" THEN EMWriteScreen "I", 12, 55
							END IF
							IF budg_mthd_mo5 = "A" THEN
								EMWriteScreen "X", 7, 61
								EMWriteScreen "X", 9, 65
								EMReadScreen mo5_elig_type, 2, 12, 61
								IF (mo5_elig_type = "AX" OR mo5_elig_type = "AA" OR mo5_elig_type = "CX") THEN EMWriteScreen "J", 12, 66
								IF (mo5_elig_type = "PX" OR mo5_elig_type = "PC") THEN EMWriteScreen "T", 12, 66
								IF (mo5_elig_type = "CK") THEN EMWriteScreen "K", 12, 66
								IF mo5_elig_type = "CB" THEN EMWriteScreen "I", 12, 66
							END IF
							IF budg_mthd_mo6 = "A" THEN
								EMWriteScreen "X", 7, 72
								EMWriteScreen "X", 9, 76
								EMReadScreen mo6_elig_type, 2, 12, 72
								IF (mo6_elig_type = "AX" OR mo6_elig_type = "AA" OR mo6_elig_type = "CX") THEN EMWriteScreen "J", 12, 77
								IF (mo6_elig_type = "PX" OR mo6_elig_type = "PC") THEN EMWriteScreen "T", 12, 77
								IF (mo6_elig_type = "CK") THEN EMWriteScreen "K", 12, 77
								IF mo6_elig_type = "CB" THEN EMWriteScreen "I", 12, 77
							END IF
							transmit		'IF Budg Mthd A, transmit to navigate to MAPT & CBUD
							DO
								EMReadScreen back_to_bsum, 4, 3, 57
								IF back_to_BSUM <> "BSUM" THEN
									EMReadScreen mapt, 4, 3, 51
									EMReadScreen cbud, 4, 3, 54
									EMReadScreen abud, 4, 3, 47
									IF mapt = "MAPT" THEN
										EMWriteScreen "PASSED", 8, 46		'=====Passes MNSure test
										transmit
										transmit
									END IF
									IF cbud = "CBUD" OR abud = "ABUD" THEN transmit		'======Getting out of the budget window for that month.
								END IF
							LOOP UNTIL back_to_bsum = "BSUM"
							'---Now the script needs to determine if the case passes income test for cert period
							EMReadScreen clt_ref_num, 2, 5, 16
							EMWriteScreen "X", 18, 3
							Transmit
							For spddn_row = 6 to 18
								EMReadScreen spenddown_type, 12, spddn_row, 39
								EMReadScreen listed_clt, 2, spddn_row, 6
								IF spenddown_type <> "NO SPENDDOWN" AND listed_clt = clt_ref_num Then
									EMWriteScreen "X", spddn_row, 3
									transmit
									EMWriteScreen " ", 5, 14
									Transmit
									Transmit
									PF3
									Exit For
								End If
								If spddn_row = 18 then PF3
							Next
							EMWriteScreen "X", 18, 34
							transmit		'---Opens the Cert Period Amount sub-menu
							DO
								EMReadScreen at_cert_period_screen, 13, 5, 13
							LOOP UNTIL at_cert_period_screen = "Certification"
							EMReadScreen excess_income, 5, 9, 39
							PF3
							IF excess_income = " 0.00" THEN
								'---The script will go into the MAPT for all appropriate months and pass Income - Budget Period.
								EMWriteScreen "X", 7, 17
								EMWriteScreen "X", 7, 28
								EMWriteScreen "X", 7, 39
								EMWriteScreen "X", 7, 50
								EMWriteScreen "X", 7, 61
								EMWriteScreen "X", 7, 72
								EMWriteScreen "X", 18, 3
								transmit
								DO
									EMReadScreen mapt_check, 4, 3, 51
									EMReadScreen mobl_check, 4, 3, 49
									IF mapt_check = "MAPT" Then
										EMWriteScreen "PASSED", 6, 46
										EMWriteScreen "PASSED", 9, 46
										EMWriteScreen "PASSED", 10, 46
										transmit
										EMReadScreen back_to_BSUM, 4, 3, 57
									ElseIf mobl_check = "MOBL" then
										For spdwn_row = 6 to 18
											EMReadScreen spdwn_check, 9, spdwn_row, 39
											IF spdwn_check = "SPENDDOWN" Then
												EMWriteScreen "X", spdwn_row, 3
												transmit
												EMWriteScreen "_", 5, 14
												transmit
												transmit
											End If
										Next
										PF3
										EMReadScreen back_to_BSUM, 4, 3, 57
									End If
								LOOP UNTIL back_to_BSUM = "BSUM"
							END IF
						END IF

						EMReadScreen cannot_fiat, 10, 24, 6
						cannot_fiat = trim(cannot_fiat)
						IF cannot_fiat <> "" THEN 		'===== IF the case is stuck in ELIG, it will not allow you to change the ELIG standard to the ACA-appropriate standard.
							PF10							'===== The script OOPS's the FIAT and backs out. It will reread and re-transmit the FIAT'd elig information.
							PF3
							EMWriteScreen "WAIT", 20, 71
							EMWaitReady 2, 2000
							EMWriteScreen "____", 20, 71
						ELSE
							PF3
						END IF

					LOOP UNTIL cannot_fiat = ""

				END IF

				hhmm_row = hhmm_row + 1

			LOOP UNTIL hc_requested = " "			'===== Loops until there are no more HC versions to review
			EMWriteScreen appl_date_month, 20, 56
			EMWriteScreen appl_date_year, 20, 59
			transmit
		End If

		'===== Now the script goes back in and approves everything.
		hhmm_row = 8
		DO
			EMReadScreen hc_requested, 1, hhmm_row, 28
			EMReadScreen hc_status, 5, hhmm_row, 68
			IF (hc_requested = "M" OR hc_requested = "S" OR hc_requested = "Q") AND hc_status = "UNAPP" THEN
				EMWriteScreen "X", hhmm_row, 26
				transmit
				DO
					EMReadScreen bhsm, 4, 3, 55
					EMReadScreen mesm, 4, 3, 56
					IF hc_requested = "M" THEN
						IF bhsm <> "BHSM" THEN
							transmit
						END IF
					ELSEIF hc_requested = "S" OR hc_requested = "Q" THEN
						IF mesm <> "MESM" THEN
							transmit
						END IF
					END IF
				LOOP UNTIL bhsm = "BHSM" OR mesm = "MESM"

				EMWriteScreen "APP", 20, 71
				STATS_manualtime = STATS_manualtime + 90    'adding manualtime for approval processing
				transmit

				'=====This portion of the script selects the possible HC programs and places an X on all of them for approval.=====
				FOR i = 9 to 24
					EMReadScreen hc_program, 1, i, 5
					IF hc_program = "_" THEN EMWriteScreen "X", i, 5
				NEXT
				transmit

				DO
					'=====This checks for a PRISM referral
					EMReadScreen prism_referral, 5, 6, 27
					IF prism_referral = "PRISM" THEN
						EMWriteScreen "N", 15, 47
						transmit
					END IF
				LOOP UNTIL prism_referral <> "PRISM"

				DO
					EMReadScreen continue_yn, 8, 21, 30
					IF continue_yn = "Continue" THEN
						EMWriteScreen "Y", 21, 46
						transmit
					END IF
				LOOP UNTIL continue_yn = "Continue"
			END IF
			hhmm_row = hhmm_row + 1
		LOOP UNTIL hc_requested = " "			'===== Loops until there are no more HC versions to review


		'Here's the case noting bit.
		'
		'CALL autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
		'CALL autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
		'CALL autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
		'
		'CALL autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
		'
		'CALL autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
		'CALL autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)
		'
		'
		'CALL navigate_to_MAXIS_screen("CASE", "NOTE")
		'PF9
		'IF cash_application = True THEN all_programs = all_programs & "CASH/"
		'IF SNAP_application = True THEN all_programs = all_programs & "SNAP/"
		'IF HC_application = True THEN all_programs = all_programs & "HC/"
		'all_programs = left(all_programs, len(all_programs) - 1)
		'CALL write_variable_in_CASE_NOTE("~~~APPROVED: " & all_programs & "~~~")
		'CALL write_bullet_and_variable_in_case_note("Earned Income", earned_income)
		'CALL write_bullet_and_variable_in_case_note("Unearned Income", unearned_income)
		'CALL write_bullet_and_variable_in_case_note("Shelter Expenses", SHEL_HEST)
		'
		''Reseting the variables for the next time.
		'HH_member_array = ""
		'earned_income = ""
		'unearned_income = ""
		'SHEL_HEST = ""
		'all_programs = ""
		'ButtonPressed = ""
	End if


	'Checks for WORK panel (Workforce One Referral), makes one with a week from now as the appointment date as a default (we can add a specific date/location checker as an enhancement
	EMReadScreen WORK_check, 4, 2, 51
	If WORK_check = "WORK" then
		wf_row = 7
		wf_count = 0
		DO
			EMReadScreen empty_space, 1, wf_row, 47
			IF empty_space = "_" THEN
				call create_MAXIS_friendly_date(date, 7, wf_row, 59)
				EMWriteScreen "X", wf_row, 47
				wf_count = wf_count + 1
			END IF
			IF empty_space = " " THEN EXIT DO
			wf_row = wf_row + 1
		LOOP UNTIL empty_space = " "
		transmit
		FOR i = 1 TO wf_count
			EMWriteScreen "X", 5, 9
			transmit
		NEXT
		transmit
		transmit
		'Special error handling for DHS and possibly multicounty agencies (don't have WF1 sites)
		EMReadScreen ES_provider_check, 2, 2, 37		'Looks for the ES in ES provider, indicating we're stuck on a screen
		If worker_county_code = "MULTICOUNTY" and ES_provider_check = "ES" then
			'Clear out the X and get back to the SELF menu
			EMWriteScreen "_", 5, 9
			transmit
			back_to_SELF
		End if
	End if
NEXT

STATS_counter = STATS_counter - 1 'removing extra counted case as it starts at 1.
If XFER_check = checked then call transfer_cases(workers_to_XFER_cases_to, case_number_array)
script_end_procedure("Success! Cases made and approved, per your request.")
