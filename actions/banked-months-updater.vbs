'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - BANKED MONTHS UPDATER.vbs"
start_time = timer
STATS_counter = 1                   'sets the stats counter at one
STATS_manualtime = 60                'manual run time in seconds
STATS_denomination = "I"       		'I for Item
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
call changelog_update("10/03/2023", "Fixed bug in saving STAT/WREG updates.", "Ilse Ferris, Hennepin County")
call changelog_update("10/02/2023", "Fixed bug in ABAWD Tracking Record where codes were not entering for the current month.", "Ilse Ferris, Hennepin County")
call changelog_update("09/13/2023", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
Call check_for_MAXIS(False)
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(initial_month, initial_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 186, 90, "Case Number/Date Selection Dialog"
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, initial_month
  EditBox 100, 25, 20, 15, initial_year
  EditBox 75, 45, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 65, 50, 15
    CancelButton 120, 65, 50, 15
  Text 10, 30, 65, 10, "Initial month/year:"
  Text 25, 10, 50, 10, "Case Number: "
  Text 10, 50, 60, 10, "Worker Signature:"
EndDialog

DO
	DO
		err_msg = ""
		dialog Dialog1
		cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(initial_month, initial_year, err_msg, "*")
		If initial_month < 10 then err_msg = err_msg & vbNewLine & "* The initial month/year cannot be prior to 10/23. Choose another date."
		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

Call check_for_MAXIS(False)

Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged, and you do not have access. The script will now end.")

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
If SNAP_CASE = False then script_end_procedure_with_error_report("This case is not a SNAP Case. The script will now end.")

Call date_array_generator(initial_month, initial_year, footer_month_array) 'Uses the custom function to create an array of dates from the initial_month and initial_year variables, ends at CM + 1.

'Need to make sure we start in the correct year for maxis
MAXIS_footer_month = initial_month
MAXIS_footer_year = initial_year

Call MAXIS_footer_month_confirmation

Dim banked_months_array
ReDim banked_months_array(banked_month_eval_const, 2)

const member_number_const 		= 0
const member_first_name_const 	= 1
const wreg_status_const 		= 2
const abawd_status_const		= 3
const abawd_mo_count_const		= 4
const abawd_mo_string_const		= 5
const banked_mo_count_const		= 6
const banked_mo_string_const	= 7
const used_all_abawd_mo_const	= 8
const used_all_banked_mo_const	= 9
const member_in_error_const		= 10
const abawd_month_eval_const 	= 11
const banked_month_eval_const	= 12

'nav to eats panel and adding all members from eats HH to evalution for banked months without having user select persons
CALL navigate_to_MAXIS_screen("STAT", "EATS")
eats_group_members = ""
memb_found = True
EMReadScreen all_eat_together, 1, 4, 72
IF all_eat_together = "_" THEN
    eats_group_members = "01" & "," 'single member HH's
ELSEIF all_eat_together = "Y" THEN
'HH's where all members eat together
    eats_row = 5
    DO
        EMReadScreen eats_pers, 2, eats_row, 3
        eats_pers = replace(eats_pers, " ", "")
        IF eats_pers <> "" THEN
            eats_group_members = eats_group_members & eats_pers & ","
            eats_row = eats_row + 1
        END IF
    LOOP UNTIL eats_pers = ""
ELSEIF all_eat_together = "N" THEN
'multiple eats HH cases - we are only caring about the 1st eats group that contains MEMB 01.
    eats_row = 13
    DO
        EMReadScreen eats_group, 38, eats_row, 39
        find_memb_01 = InStr(eats_group, "01")
		
        IF find_memb_01 = 0 THEN
            eats_row = eats_row + 1
            IF eats_row = 18 THEN
                memb_found = False
                EXIT DO
            END IF
		END IF
    LOOP UNTIL find_memb_01 <> 0
    'Then Gathering the eats group members from the all_eat_together = "N" group
    eats_col = 39
    DO
        EMReadScreen eats_group, 2, eats_row, eats_col		'reading the eats member 
        IF eats_group <> "__" THEN
            eats_group_members = eats_group_members & eats_group & ","	'adds to the string if not __
            eats_col = eats_col + 4		
        END IF
    LOOP UNTIL eats_group = "__"
END IF

CALL navigate_to_MAXIS_screen("STAT", "MEMB")
If right(eats_group_members, 1) = "," THEN eats_group_members = left(eats_group_members, len(eats_group_members) - 1)
eats_array = split(eats_group_members, ",")

entry_count = 0
For each member in eats_array 
	'msgbox member
	Call write_value_and_transmit(member, 20, 76)
	EmReadScreen member_first_name, 12, 6, 63
	member_first_name = replace(member_first_name, "_", "")
	'Adding the EATS HH to the array once name and member numbers are obtained 
	ReDim Preserve banked_months_array(banked_month_eval_const, entry_count)
	banked_months_array(member_number_const, entry_count) = member
	banked_months_array(member_first_name_const, entry_count) = trim(member_first_name)
	entry_count = entry_count + 1
	Stats_counter = stats_counter + 1
Next 

'Initial Month evaluation of ABAWD/TLR - Banked Months
For i = 0 to Ubound(banked_months_array, 2)
	CALL navigate_to_MAXIS_screen("STAT", "WREG")	'Navigate to STAT/WREG and check for WREG Status codes
	Call write_value_and_transmit(banked_months_array(member_number_const, i), 20, 76)
	EMReadScreen panel_exists, 1, 2, 78
	If panel_exists = "0" then 
		banked_months_array(abawd_month_eval_const, i) = "No WREG Panel."
	Else 
	    EMreadScreen wreg_status, 2, 8, 50
	    EMReadScreen abawd_status, 2, 13, 50
	
	    Call write_value_and_transmit("X", 13, 57)		'navigate to ABAWD/TLR Tracking panel and check for historical months
	    TLR_fixed_clock_mo = "01" 'fixed clock dates for all recipients 
	    TLR_fixed_clock_yr = "23"
	
		asssessment_month = MAXIS_footer_month - 1
	    bene_mo_col = (15 + (4*cint(asssessment_month)))		'col to search starts at 15, increased by 4 for each footer month
	    bene_yr_row = 10
	    abawd_counted_months = 0					'delclares the variables values at 0 or blanks
	    banked_months_count = 0
	    abawd_counted_months_string = ""
	    banked_months_string = ""
	    DO
            'establishing variables for specific ABAWD counted month dates
            If bene_mo_col = "19" then counted_date_month = "01"
            If bene_mo_col = "23" then counted_date_month = "02"
            If bene_mo_col = "27" then counted_date_month = "03"
            If bene_mo_col = "31" then counted_date_month = "04"
            If bene_mo_col = "35" then counted_date_month = "05"
            If bene_mo_col = "39" then counted_date_month = "06"
            If bene_mo_col = "43" then counted_date_month = "07"
            If bene_mo_col = "47" then counted_date_month = "08"
            If bene_mo_col = "51" then counted_date_month = "09"
            If bene_mo_col = "55" then counted_date_month = "10"
            If bene_mo_col = "59" then counted_date_month = "11"
            If bene_mo_col = "63" then counted_date_month = "12"
            'counted date year: this is found on rows 7-10. Row 11 is current year plus one, so this will be exclude this list.
            If bene_yr_row = "10" then counted_date_year = right(DatePart("yyyy", date), 2)
            If bene_yr_row = "9"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -1, date)), 2)
            If bene_yr_row = "8"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -2, date)), 2)
            If bene_yr_row = "7"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -3, date)), 2)
            
            'reading to see if a month is counted month or not
            EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
            
            'counting and checking for counted ABAWD months
            IF is_counted_month = "X" or is_counted_month = "M" THEN
            	EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
            	abawd_counted_months = abawd_counted_months + 1				'adding counted months
	    		abawd_counted_months_string = abawd_counted_months_string & counted_date_month & "/" & counted_date_year & " | "
            END IF
            
            'counting and checking for counted banked months
            IF is_counted_month = "B" or is_counted_month = "C" THEN
            	EMReadScreen counted_date_year, 2, bene_yr_row, 14			'reading counted year date
            	banked_months_count = banked_months_count + 1				'adding counted months
	    		banked_months_string = banked_months_string & counted_date_month & "/" & counted_date_year & " |"
            END IF
            
            bene_mo_col = bene_mo_col - 4		're-establishing serach once the end of the row is reached
            IF bene_mo_col = 15 THEN
            	bene_yr_row = bene_yr_row - 1
            	bene_mo_col = 63
            END IF
	    	'used to loop until count was 36 due to person based look back period. Now fixed clock starts 01/23 for all members. 
	    LOOP until (counted_date_month = TLR_fixed_clock_mo AND counted_date_year = TLR_fixed_clock_yr)

		'cleaning up these variables for dialog display
		If trim(right(abawd_counted_months_string, 2)) = " |" THEN abawd_counted_months_string = left(abawd_counted_months_string, len(abawd_counted_months_strings) - 2)
		If trim(right(banked_months_string, 2)) = " |" THEN banked_months_string = left(banked_months_string, len(banked_months_string) - 2)
	
	    'msgbox "ABAWD's: " & abawd_counted_months & vbcr & "Banked: " & banked_months_count
		'msgbox abawd_counted_months_string & vbcr & banked_months_string
	    PF3	' to exit tracking record 	     
	End If 

	'stores the information for each member & makes ABAWD/banked months determinations 
	banked_months_array(wreg_status_const 		, i) = wreg_status
	banked_months_array(abawd_status_const		, i) = abawd_status
	banked_months_array(abawd_mo_count_const	, i) = abawd_counted_months
	banked_months_array(abawd_mo_string_const	, i) = abawd_counted_months_string
	banked_months_array(banked_mo_count_const	, i) = banked_months_count
	banked_months_array(banked_mo_string_const	, i) = banked_months_string
	'ABAWD Determinations
	If abawd_counted_months = 3 then 
		banked_months_array(used_all_abawd_mo_const, i) = True 
		banked_months_array(abawd_month_eval_const, i) = "All ABAWD/TLR months used."
	Elseif abawd_counted_months < 3 then 
		banked_months_array(used_all_abawd_mo_const, i) = False
		banked_months_array(abawd_month_eval_const, i) = "Only " & abawd_counted_months & " have been used. This member does not require banked months yet."
	Else 
		banked_months_array(member_in_error_const, i) = banked_months_array(member_number_const, i) & " " & banked_months_array(member_first_name_const, i) & " has used " & abawd_months_count & " abawd months. Only 3 are allowed. Updates are needed to this STAT/WREG and/or ABAWD Tracking Record before the script can support this case." 
		banked_months_array(abawd_month_eval_const, i) = abawd_counted_months & " have been used. Only 3 are allowed. Updates are needed."
	End if 

	'Banked Months Determinations
	If banked_months_count = 2 then 
		banked_months_array(used_all_banked_mo_const, i) = True 
		banked_months_array(banked_month_eval_const, i) = "All banked months have been used for this member."
	Elseif banked_months_count = 1 then 
		banked_months_array(used_all_banked_mo_const, i) = False 
		banked_months_array(banked_month_eval_const, i) = "1 banked months available." 
	Elseif banked_months_count = 0 then 
		banked_months_array(used_all_banked_mo_const, i) = False 
		banked_months_array(banked_month_eval_const, i) = "2 banked months available." 
	Else 
		banked_months_array(member_in_error_const, i) = banked_months_array(member_number_const, i) & " " & banked_months_array(member_first_name_const, i) & " has used " & banked_months_count & " banked months. Only 2 are allowed. Updates are needed to this STAT/WREG and/or ABAWD Tracking Record before the script can support this case." 
		banked_months_array(banked_month_eval_const, i) = banked_months_count & " have been used. Only 2 are allowed. Updates are needed."
	End if
Next 

end_msg = ""	'defaults for the next for...next which determines a bunch of actions 
update_wreg = False 'defaulting to false unless there are banked months related updates to make 
spec_memo = False 
case_note = False 
Show_dialog = False 

For i = 0 to Ubound(banked_months_array, 2)
	'Determines if updates aren't going to be made to STAT/WREG
	If banked_months_array(used_all_abawd_mo_const, i) = True then
		update_wreg = True
	Elseif banked_months_array(member_in_error_const, i) <> "" then 
		update_wreg = True 	
	End if 

	If banked_months_array(used_all_banked_mo_const, i) = False then spec_memo = True 
	If update_wreg = False then spec_memo = False 
	If update_wreg = True then case_note = True 

	If banked_months_array(member_in_error_const, i) <> "" then end_msg = end_msg & banked_months_array(member_in_error_const, i) & vbcr & vbcr 
	If banked_months_array(abawd_status_const, i) = "10" or banked_months_array(abawd_status_const, i) = "13" then show_dialog = True 
Next 

'script end procedure for anyone who is in error 
If end_msg <> "" then script_end_procedure_with_error_report(end_msg)  
'Don't show the dialog if no one is already coded as ABAWD or banked for the initial month
If Show_dialog = False then script_end_procedure_with_error_report("No members in the EATS Household with MEMB 01 are coded as ABAWD (30/10) or Banked (30/13) for the initial month. The script will now end.")

script_actions = "" 'variable that tells users what's going on in dialog. 
If update_wreg = True then script_actions = script_actions & VBCR & "* The STAT/WREG panel will be updated for members who have banked months available, or who have exhausted them."
If spec_memo = True then script_actions = script_actions & VBCR & "* A SPEC/MEMO with required Banked Months and TLR Text will be sent to the resident, AREP, and/or SWKR."
If case_note = True  then script_actions = script_actions & VBCR & "* A detailed CASE/NOTE will be created about case and member updates."

If script_actions = "" then script_actions = "None! This case does not have any members that require SNAP Banked Months updating. The script will end after you press OK or Cancel."

'msgbox "script actions: " & script_actions

Dialog1 = ""		    '----------------------------------------------------------------------------------------------------Displaying Main Dialog 
BeginDialog Dialog1, 0, 0, 550, 385, "Banked Months Evaluation Dialog for ABAWD/TLR's on Case #" & MAXIS_case_number
  Text 65, 10, 455, 10, "**The following information are for the ABAWD and Banked Months counts/months members in the EATS Household containing Memb 01**"
  Text 115, 25, 315, 10, "--The information below reflects counts from the initial month/year selected in the initial dialog.--"
dialog_item = 0
For i = 0 to Ubound(banked_months_array, 2)
	If banked_months_array(abawd_status_const, i) = "10" or banked_months_array(abawd_status_const, i) = "13" then 
		y_pos = (45 + dialog_item * 50)
		GroupBox 10, y_pos, 530, 45, "MEMB #" & banked_months_array(member_number_const , i) & " " & banked_months_array(member_first_name_const, i) & ", ABAWD/FSET: " & banked_months_array(wreg_status_const, i) & "/" & banked_months_array(abawd_status_const, i)
  		Text 20, y_pos + 15, 255, 10, "ABAWD Count/Months Used: " & banked_months_array(abawd_mo_count_const, i) & " - " & banked_months_array(abawd_mo_string_const, i)
  		Text 20, y_pos + 30, 255, 10, "ABAWD Months Evaluation: " & banked_months_array(abawd_month_eval_const, i)
 	 	Text 285, y_pos + 15, 245, 10, "Banked Count/Months Used: " & banked_months_array(banked_mo_count_const, i) & " - " & banked_months_array(banked_mo_string_const, i)
  		Text 285, y_pos + 30, 245, 10, "Banked Months Evaluation: " & banked_months_array(banked_month_eval_const, i)
		dialog_item = dialog_item + 1
	End if 
Next
'descriptor of what the script will do for the user. See incrementor variable above dialog for details/ 	
GroupBox 10, 310, 415, 65, "What Actions The Script Will Take:"
Text 25, 325, 390, 40, script_actions
ButtonGroup ButtonPressed
  OkButton 435, 360, 50, 15
  CancelButton 490, 360, 50, 15
EndDialog

Do 
    dialog dialog1
    cancel_without_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

If case_note = False then script_end_procedure_with_error_report("There are no updates to be made, including case note. The script will now end.")

Call check_for_MAXIS(False)

For item = 0 to ubound(footer_month_array)
	MAXIS_footer_month = datepart("m", footer_month_array(item)) 'Need to assign footer month / year each time through
	If len(MAXIS_footer_month) = 1 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
	MAXIS_footer_year = right(datepart("YYYY", footer_month_array(item)), 2)

	For i = 0 to Ubound(banked_months_array, 2)
		If banked_months_array(abawd_status_const, i) = "10" or banked_months_array(abawd_status_const, i) = "13" then 
		    Call MAXIS_background_check
		    Call navigate_to_MAXIS_screen("STAT", "WREG")
	        Call write_value_and_transmit(banked_months_array(member_number_const, i), 20, 76)	'enters member number
	        PF9
	        EMWriteScreen "30", 8, 50	'wreg code
	        EMWriteScreen "N", 8, 80	'defer FSET funds code
		    If banked_months_array(banked_mo_count_const, i) < 2 then 
		    	'msgbox banked_months_array(banked_mo_count_const, i)
	        	EMWriteScreen "13", 13, 50	'banked months ABAWD code 
		    	banked_months_array(banked_mo_count_const, i) = banked_months_array(banked_mo_count_const, i) + 1 'incrementing the BM count & BM month string
		    	banked_months_array(banked_mo_string_const, i) = banked_months_array(banked_mo_string_const, i) & MAXIS_footer_month & "/" & MAXIS_footer_year & " | "
	        	EMWriteScreen banked_months_array(banked_mo_count_const, i), 14, 50	'banked months count code	    		
	        	If banked_months_array(banked_mo_count_const, i) = 2 then banked_months_array(used_all_banked_mo_const, i) = True 
		    Else
		    	EMWriteScreen "10", 13, 50	'ABAWD Counted Months code
		    End If 
			PF3
		    transmit ' to save

			'----------------------------------------------------------------------------------------------------ABAWD TRACKING RECORD Updates
		    EMReadScreen ABAWD_coding, 2, 13, 50	'confirming what's being updated to determine ABAWD tracking recording updating
		    'Only updating the ABAWD tracking record with manual entry for banked months IF the MAXIS mo/yr = CM mo/yr. If not the system will update upon approval. 
		    Processing_month = MAXIS_footer_month & "/01/" & MAXIS_Footer_year
		    'msgbox "DateDiff " & DateDiff("M", processing_month, date)
		    If DateDiff("M", processing_month, date) => 0 then 
				PF9
				If ABAWD_coding = "10" Then ATR_code = "M"'manual counted ABAWD month
		    	If ABAWD_coding = "13" Then ATR_code = "C" 'manual counted banked month
				'Update tracking record
	    		ATR_updates = array("D", ATR_code)
	    		For each update_code in ATR_updates
	    		    'msgbox update_code
					Call write_value_and_transmit("X", 13, 57) 'Pulls up the WREG tracker'        	    
	    		    bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))		'col to search starts at 15, increased by 4 for each footer month
        		    bene_yr_row = 10
			    	Call write_value_and_transmit(update_code, bene_yr_row,bene_mo_col)
					PF3 'to go back to WREG/Panel
	    		Next     	
	        	PF3	' to save and exit ABAWD tracking record
		    End if 
	        transmit 'to save
    
	        PF3 'to save and exit to stat/wrap
		    Call back_to_SELF
		    stats_counter = STATS_counter + 1
		    'cleaning up variable again for Case note output 
		    If trim(right(banked_months_string, 2)) = " |" THEN banked_months_string = left(banked_months_string, len(banked_months_string) - 2)
		End if 
	Next     
Next 


'SPEC/MEMO is being sent in leiu of SPEC/WCOM per Bulletin #23-01-02 https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=mndhs-063946
If spec_memo = True then
	Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)	'navigates to spec/memo and opens into edit mode
	'Writes the info into the MEMO.
	Call write_variable_in_SPEC_MEMO("************************************************************")
	Call write_variable_in_SPEC_MEMO("You are getting this letter because you or someone in your SNAP unit needs to follow the time-limited work rules and have used all three available months.")
	Call write_variable_in_SPEC_MEMO("")
	Call write_variable_in_SPEC_MEMO("Unless you or someone in your SNAP unit meet work rules or an exemption, you/they will no longer be eligible for SNAP.")
	Call write_variable_in_SPEC_MEMO("")
	Call write_variable_in_SPEC_MEMO("However, due to additional funding we are able to approve SNAP benefits for up to 2 more months.")
	Call write_variable_in_SPEC_MEMO("")
	Call write_variable_in_SPEC_MEMO("If you/someone in your SNAP unit is not meeting work requirements/meeting an exemption, you/they will no longer receive SNAP after these 2 months.")
	Call write_variable_in_SPEC_MEMO("")
	Call write_variable_in_SPEC_MEMO("Please contact your team if you, or someone in your SNAP unit, start meeting work requirements or think that you meet an exemption.")
	Call write_variable_in_SPEC_MEMO("")
	Call write_variable_in_SPEC_MEMO("If you need help meeting these work requirements, please see the SNAP Time-limited work rules website at: https://mn.gov/dhs/snap-e-and-t/time-limited-work-rules/.")
	Call write_variable_in_SPEC_MEMO("************************************************************")
	PF4 'save memo
	stats_counter = STATS_counter + 1
End if 
'----------------------------------------------------------------------------------------------------CASE/NOTE
Call start_a_blank_CASE_NOTE
Call write_variable_in_case_note("--SNAP Banked Months Evaluation for " & initial_month & "/" & initial_year & "--")
Call write_variable_in_case_note("Case TLR Information by Member")
Call write_variable_in_case_note("------------------------------")
For i = 0 to Ubound(banked_months_array, 2)
	Call write_variable_in_case_note("MEMB #" & banked_months_array(member_number_const , i) & " - " & banked_months_array(member_first_name_const, i) & ", ABAWD/FSET:(" & banked_months_array(wreg_status_const, i) & "/" & banked_months_array(abawd_status_const, i) & ")")
	'Call write_variable_in_case_note(banked_months_array(member_first_name_const, i) & " (MEMB " & banked_months_array(member_number_const , i) & ")")
	If banked_months_array(abawd_status_const, i) = "10" or banked_months_array(abawd_status_const, i) = "13" then
        'Call write_variable_in_case_note(banked_months_array(member_first_name_const, i) & " (MEMB " & banked_months_array(member_number_const , i) & ")")
        Call write_variable_in_case_note("* ABAWD Count/Months Used: " & banked_months_array(abawd_mo_count_const, i) & " - " & banked_months_array(abawd_mo_string_const, i))
        Call write_variable_in_case_note("* ABAWD Months Evaluation: " & banked_months_array(abawd_month_eval_const, i))
        Call write_variable_in_case_note("* Banked Count/Months Used: " & banked_months_array(banked_mo_count_const, i) & " - " & banked_months_array(banked_mo_string_const, i))
        Call write_variable_in_case_note("--")
	End if 
	stats_counter = STATS_counter + 1
Next 
If update_wreg = True then Call write_variable_in_case_note("* The STAT/WREG panel was updated through CM + 1 for members who have banked months available, or who have exhausted them.")
If spec_memo = True then Call write_variable_in_case_note("* A SPEC/MEMO was sent to the household with banked months and time-limited SNAP information.")
Call write_variable_in_case_note("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
stats_counter = STATS_counter + 1

end_msg = "The main EATS household has been assessed and updated for banked months. Make sure to APP new SNAP results, and use NOTES - ELIGIBILITY SUMMARY to case note the eligibility."
If spec_memo = True then end_msg = end_msg & vbcr & vbcr & "The required SPEC/MEMO was sent to the household following details in Bulletin #23-01-02."
script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/19/2023	
'--Tab orders reviewed & confirmed----------------------------------------------09/19/2023
'--Mandatory fields all present & Reviewed--------------------------------------09/19/2023
'--All variables in dialog match mandatory fields-------------------------------09/19/2023
'Review dialog names for content and content fit in dialog----------------------09/19/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/19/2023
'--CASE:NOTE Header doesn't look funky------------------------------------------09/19/2023
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------09/19/2023
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-09/19/2023 --------------The ones without periods have them in the variables or output a string or dates and it looked funnny.
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/19/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------09/19/2023
'--PRIV Case handling reviewed -------------------------------------------------09/19/2023
'--Out-of-County handling reviewed----------------------------------------------09/19/2023-----------------N/A staff are working off lists, out of county already filtered out. 
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/19/2023	
'--BULK - review output of statistics and run time/count (if applicable)--------09/19/2023
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "i")-----------09/19/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------09/19/2023
'--Incrementors reviewed (if necessary)-----------------------------------------09/19/2023
'--Denomination reviewed -------------------------------------------------------09/19/2023
'--Script name reviewed---------------------------------------------------------09/19/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------09/19/2023----------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------09/19/2023
'--comment Code-----------------------------------------------------------------09/19/2023
'--Update Changelog for release/update------------------------------------------09/19/2023
'--Remove testing message boxes-------------------------------------------------09/19/2023----------------Keeping some of these until testing after 1st of the month for more than 1 month updates. 
'--Remove testing code/unnecessary code-----------------------------------------09/19/2023
'--Review/update SharePoint instructions----------------------------------------09/19/2023----------------In progress
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------09/19/2023----------------In progress
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/19/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------09/19/2023
'--Complete misc. documentation (if applicable)---------------------------------09/19/2023
'--Update project team/issue contact (if applicable)----------------------------09/19/2023