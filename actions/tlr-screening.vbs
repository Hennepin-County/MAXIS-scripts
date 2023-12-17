'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - TLR SCREENING.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 750                	'manual run time in seconds
STATS_denomination = "C"       		'C is for Case
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

'Dialogs===================================================================================================================
EMConnect ""

Call check_for_MAXIS(False)
Call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 181, 100, "Case #/Member # Selection"
  Text 20, 15, 50, 10, "Case Number: "
  EditBox 75, 10, 50, 15, MAXIS_case_number
  Text 10, 35, 60, 10, "Member Number:"
  EditBox 75, 30, 30, 15, member_number
  Text 10, 55, 60, 10, "Worker Signature:"
  EditBox 75, 50, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 40, 75, 50, 15
    CancelButton 95, 75, 50, 15
EndDialog

Do
	Do
	    err_msg = ""
  		Dialog Dialog1
  		Cancel_without_confirmation
  		Call validate_MAXIS_case_number(display_ben_err_msg, "*")
		If IsNumeric(member_number) = False or len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit member number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call MAXIS_footer_month_confirmation	'making sure we're getting to current month/year
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged, and you do not have access. The script will now end.")

'member_info = member_number & " - " & first_name & " " & last name



Call write_value_and_transmit(member_number, 20, 76)
EMReadScreen panel_exists, 1, 2, 78
If panel_exists = "0" then
	current_ATR_status = ""
else  
	EMreadScreen wreg_status, 2, 8, 50
	EMReadScreen abawd_status, 2, 13, 50

	Call write_value_and_transmit("X", 13, 57)		'navigate to ABAWD/TLR Tracking panel and check for historical months

	'Resetting the variables
	asssessment_month = MAXIS_footer_month - 1
	bene_mo_col = (15 + (4*cint(asssessment_month)))		'col to search starts at 15, increased by 4 for each footer month
	bene_yr_row = 10
	abawd_counted_months = 0					'delclares the variables values at 0 or blanks
	second_set_count = 0
	abawd_status = 0
	wreg_status = 0
	abawd_counted_months_string = ""
	second_set_string = ""

	TLR_fixed_clock_mo = "01" 'fixed clock dates for all recipients 
	TLR_fixed_clock_yr = "23"
	
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
        IF is_counted_month = "Y" or is_counted_month = "N" THEN
        	EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
        	second_set_count = second_set_count + 1				'adding counted months
    		second_set_string = second_set_string & counted_date_month & "/" & counted_date_year & " |"
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
If trim(right(second_set_string, 2)) = " |" THEN second_set_string = left(second_set_string, len(second_set_string) - 2)
PF3	' to exit tracking record 	     

If abawd_counted_months = 3 then 
	abawd_month_eval = "All ABAWD/TLR months used."
Elseif abawd_counted_months < 3 then 
	abawd_month_eval = abawd_counted_months & " counted months have been used."
Else 
	abawd_counted_months & " have been used. Only 3 are allowed. Updates are required."
End if 
'Banked Months Determinations
If second_set_count = 0 then 
	second_month_eval = "No 2nd set months have been used."
ElseIF second_set_count = 3 then 
	second_month_eval = "All 2nd set months have been used."
Else
	second_month_eval = "Review is required for 2nd set. Appears to be in error."
End if


Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 326, 375, "Time-Limited Recipient (TLR) Screening"
  GroupBox 5, 5, 315, 55, "TLR Information for & Member_info"
  Text 15, 20, 235, 10, "Counted TLR Months: " & abawd_counted_months & " - " & abawd_counted_months_string 
  Text 15, 40, 235, 10, "Counted 2nd Set: " & second_set_count & " - " & second_set_string
  Text 5, 65, 315, 10, "=============================================================================="
  Text 60, 80, 190, 10, "Select ALL applicable exemptions for this member below"
  Text 5, 95, 315, 10, "=============================================================================="
  GroupBox 5, 110, 315, 55, "Unfit for Employment:"
  CheckBox 20, 120, 170, 10, "Physical illness, injury, disability or limitation?*", disa_checkbox
  CheckBox 20, 135, 180, 10, "Temp or Perm DISA from SSA, VA, Work Comp, etc?", perm_disa_checkbox
  CheckBox 20, 150, 155, 10, "Substance abuse or addiction dependency?*", sub_abuse_checkbox
  CheckBox 200, 120, 115, 10, "Mental illness, disorder, etc.*", mental_illness_checkbox
  CheckBox 200, 135, 50, 10, "Homeless?*", homeless_checkbox
  CheckBox 200, 150, 115, 10, "A victim of domestic violence?*", dom_violence_checkbox
  CheckBox 5, 170, 295, 10, "Caring for person who needs help caring for themselves (can be outside the home)?*", care_of_hh_memb_checkbox
  CheckBox 5, 185, 160, 10, "Responsible for the care of a child under 6?*", care_child_six_checkbox
  CheckBox 5, 200, 70, 10, "Age 60 or older?*", age_sixty_checkbox
  CheckBox 5, 215, 85, 10, "Under the age of 16?*", under_sixteen_checkbox
  CheckBox 5, 230, 160, 10, "Aged 16 or 17 living w/ parent or caregiver?*", sixteen_seventeen_checkbox
  CheckBox 5, 245, 270, 10, "Employed 30 hours/week or grossing at least $217.50/week ($935.25/month)?", employed_thirty_checkbox
  CheckBox 5, 260, 185, 10, "Receiving or applied for unemployment insurance?", unemployment_checkbox
  CheckBox 5, 275, 100, 10, "Receiving Matching Grant?", matching_grant_checkbox
  CheckBox 5, 290, 215, 10, "Receiving DWP and in compliance with Employment Services?", DWP_checkbox
  CheckBox 5, 305, 215, 10, "Receiving MFIP and in compliance with Employment Services?", MFIP_checkbox
  CheckBox 5, 320, 265, 10, "Enrolled in school, training program, or higher education at least half time?", enrolled_school_checkbox
  CheckBox 5, 335, 210, 10, "Participating in drug or alcohol addiction treatment program?", CD_program_checkbox
  Text 5, 355, 190, 10, "* = Can be declaratory unless inconsistent/questionable."
  ButtonGroup ButtonPressed
    OkButton 215, 350, 50, 15
    CancelButton 270, 350, 50, 15
    PushButton 260, 15, 55, 15, "CM 0028.06.12", CM_button
    PushButton 260, 35, 55, 15, "CM 0011.24", TLR_CM_button
EndDialog

'Maybe if none are selected checkbox total number? 
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 326, 320, "Subject to Work Rules, Exempt from Time Limits"
  GroupBox 5, 5, 315, 55, "TLR Information for & Member_info"
  Text 15, 20, 235, 10, "Counted TLR Months:  & abawd_counted_months &  -  & abawd_counted_months_string"
  Text 15, 40, 235, 10, "Counted 2nd Set:  & second_set_count &  -  & second_set_string"
  Text 5, 65, 315, 10, "=============================================================================="
  Text 60, 80, 190, 10, "Select ALL applicable exemptions for this member below"
  Text 5, 95, 315, 10, "=============================================================================="
  CheckBox 5, 110, 120, 10, "Younger than 18 OR 53 or older?*", age_exempt
  CheckBox 5, 125, 125, 10, "Child under 18 in your SNAP unit?*", minor_hh_checkbox
  CheckBox 5, 140, 155, 10, "16-17 and NOT living with parent/caregiver?*", minor_wo_caregiver_checkbox
  CheckBox 5, 155, 45, 10, "Pregnant?*", PX_checkbox
  CheckBox 5, 170, 100, 10, "Served in the US Military?*", veteran_checkbox
  CheckBox 5, 185, 185, 10, "In foster care on 18th birthday AND under age 25?*", foster_care_checkbox
  CheckBox 5, 200, 255, 10, "RCA recipient and participating in Refugee Employment Services 1/2 time?", RCA_checkbox
  CheckBox 5, 215, 165, 10, "Responsible for the care of a dependent child?", dependent_child
  CheckBox 5, 230, 235, 10, "Employed/Self-employed at least 20 hours/week or 80 hours/month?", working_20_checkbox
  CheckBox 5, 245, 260, 10, "Participating in an approved work/training program at least 20 hours/month?", approved_work_checkbox
  CheckBox 5, 260, 275, 10, "Volunteering OR combo of work, training, or volunteering at least 80 hours/month?", combo_work_checkbox
  Text 5, 280, 190, 10, "* = Can be declaratory unless inconsistent/questionable."
  ButtonGroup ButtonPressed
    PushButton 155, 295, 50, 15, "Previous", previous_button
    OkButton 210, 295, 50, 15
    CancelButton 265, 295, 50, 15
    PushButton 260, 15, 55, 15, "CM 0028.06.12", CM_button
    PushButton 260, 35, 55, 15, "CM 0011.24", TLR_CM_button
EndDialog

'Maybe if none are selected checkbox total number? TODO!

Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 326, 180, "Second Set TLR Months"
  GroupBox 5, 5, 315, 55, "TLR Information for & Member_info"
  Text 15, 20, 235, 10, "Counted TLR Months:  & abawd_counted_months &  -  & abawd_counted_months_string"
  Text 15, 40, 235, 10, "Counted 2nd Set:  & second_set_count &  -  & second_set_string"
  Text 5, 65, 315, 10, "=============================================================================="
  Text 60, 80, 190, 10, "Select ALL applicable situations for this member below"
  Text 5, 95, 315, 10, "=============================================================================="
  CheckBox 10, 105, 160, 10, "Used all 3 counted TLR months since 01/23?", Used_TLR_checkbox
  CheckBox 10, 120, 260, 10, "Worked at least 80 hours in a month SINCE closing for using 3 TLR months?", worked_80_since_closing
  CheckBox 10, 135, 250, 10, "Work/work activities have ended or reduced to less than 80 hours/month?", work_ended_checkbox
  ButtonGroup ButtonPressed
    PushButton 10, 155, 110, 15, "View ABAWD Tracking Record", open_ATR_button
    PushButton 155, 155, 50, 15, "Previous", previous_button
    PushButton 210, 155, 50, 15, "Finish", finish_button
    CancelButton 265, 155, 50, 15
    PushButton 260, 15, 55, 15, "CM 0028.06.12", CM_button
    PushButton 260, 35, 55, 15, "CM 0011.24", TLR_CM_button
EndDialog


'Need to figure out the rest of this here. 

'This dialog gets the worker's signature and allows the OSA to enter any comments for the case worker.----------------------
'The idea being that if the OSA notices irregularities or unusualness (word?) in the TLR tracking panel, it---------------
'can be reported to the worker or the worker can be directed to look deeper into the TLR tracking.------------------------
BeginDialog get_worker_comments, 0, 0, 166, 105, "TLR Screening Tool"
  EditBox 5, 50, 155, 15, worker_comment
  ButtonGroup ButtonPressed
    PushButton 20, 75, 50, 15, "OK", OK_button
    CancelButton 90, 75, 50, 15
  Text 5, 10, 150, 10, "Case noting CL interaction."
  Text 5, 25, 160, 20, "Any additional comments, please enter here. Press ENTER to complete and Case Note."
EndDialog

'----------------------------------------------------------------------------------------------------CASE/NOTE
Call start_a_blank_case_note
call write_variable_in_CASE_NOTE("***Member " & member_number & " has been screened for TLR***")
call write_variable_in_CASE_NOTE(tlr_status)
IF worked_80_since_closing = 1 AND has_used_second_period <> 1 THEN call write_variable_in_CASE_NOTE("* CL has earned additional 3-month period of TLR eligibility.")
IF worked_80_since_closing = 1 AND has_used_second_period = 1 THEN call write_variable_in_CASE_NOTE("* Client has used 2nd 3 months of eligibility, and 80 hours a month since closure. However they must meet another exemption.")
IF worked_80_since_closing <> 1 and has_used_second_period = 1 THEN call write_variable_in_CASE_NOTE("* Client has used 2nd 3 months of eligibility, must meet exemption to be eligible for SNAP")
IF wreg_disa = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are unfit for employment due to")
	IF radiocheck1 = 1 THEN call write_variable_in_CASE_NOTE("      * physical or mental illness, disability, or injury.")
	IF radiocheck2 = 1 THEN call write_variable_in_CASE_NOTE("      * homelessness.")
	IF radiocheck3 = 1 THEN call write_variable_in_CASE_NOTE("      * veteran status.")
	IF radiocheck4 = 1 THEN call write_variable_in_CASE_NOTE("      * victim of domestic violence.")
IF care_of_hh_memb = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are responsible for care of a disabled unit member")
IF age_sixty = 1 THEN call write_variable_in_CASE_NOTE("* Client is over 60.")
IF under_sixteen = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are under 16.")
IF sixteen_seventeen = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are age 16 or 17 and living with a parent or caretaker")
IF care_child_six = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are responsible for the care of a child less than age 6.")
IF employed_thirty = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are employed 30 hours per week or equivalent to 30 hours a week at minimum wage.")
IF unemployment = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are receiving or applied for unemployment insurance.")
IF enrolled_school = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are enrolled in school/training 1/2 time.")
IF CD_program = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are participating in a chemical dependency treatment program.")
IF receiving_MFIP = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are a MFIP recipient.")
IF receiving_DWP_WB = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are a DWP recipient or applicant.")
IF waiver = 1 THEN call write_variable_in_CASE_NOTE("* Client is residing in a waived area")
IF age_exempt = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are under age 18 or age 50 or older")
IF cert_preg = 1 THEN call write_variable_in_CASE_NOTE("* Client states certified as pregnant")
IF working_20 = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are employed 20 hours per week")
IF dependent_child = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are responsible for the care of a dependent child in the household")
IF work_exp = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are participating in work experience program")
IF approved_ET = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are participating in employment and training program")
IF receiving_cash = 1 THEN call write_variable_in_CASE_NOTE("* Client states they are a RCA or GA recipient")
call write_bullet_and_variable_in_CASE_NOTE("Other notes", worker_comment)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("Success! This member has been assessed for time-limited SNAP.")