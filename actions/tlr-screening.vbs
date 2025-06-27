'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - TLR SCREENING.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 750                	'manual run time in seconds
STATS_denomination = "M"       		'M is for Member
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
call changelog_update("12/10/2024", "Fixed bug in date-based ABAWD evaluation to work dynamically by footer month/year selected.", "Ilse Ferris, Hennepin County")
Call changelog_update("10/07/2024", "Added Age-Based exemption from 53-59 to 55-59 based on 10/2024 policy.", "Ilse Ferris, Hennepin County")
Call changelog_update("06/27/2024", "Added update handling for residents who meet military service ABAWD/TLR exemptions.", "Ilse Ferris, Hennepin County")
Call changelog_update("12/29/2023", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Dialogs===================================================================================================================
EMConnect ""
Call check_for_MAXIS(False)
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
member_number = "01"

Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 181, 110, "ACTIONS - TLR SCREENING"
  Text 20, 15, 50, 10, "Case Number: "
  EditBox 75, 10, 45, 15, MAXIS_case_number
  Text 10, 35, 60, 10, "Member Number:"
  EditBox 75, 30, 30, 15, member_number
  Text 5, 55, 65, 10, "Footer month/year:"
  EditBox 75, 50, 20, 15, MAXIS_footer_month
  EditBox 100, 50, 20, 15, MAXIS_footer_year
  Text 10, 75, 60, 10, "Worker Signature:"
  EditBox 75, 70, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 15, 90, 70, 15, "Script Instructions", script_instructions
    OkButton 90, 90, 40, 15
    CancelButton 135, 90, 40, 15
EndDialog

Do
	Do
	    err_msg = ""
  		Dialog Dialog1
        Cancel_without_confirmation
  		Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If IsNumeric(member_number) = False or len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2-digit member number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
        If ButtonPressed = script_instructions then 
            call open_URL_in_browser("https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/_layouts/15/Doc.aspx?sourcedoc=%7B58B17691-CF4B-4EBF-8B97-B556E995F67D%7D&file=ACTIONS%20-%20TLR%20SCREENING.docx")
            err_msg = "LOOP" & err_msg
        End if
		IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

ABAWD_eval_date = MAXIS_footer_month & "/1/" & MAXIS_footer_year

Call MAXIS_footer_month_confirmation	'making sure we're getting to current month/year
Call MAXIS_background_check
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged, and you do not have access. The script will now end.")

Call write_value_and_transmit(member_number, 20, 76)
EmReadScreen first_name, 12, 6, 63
first_name = replace(first_name, "_", "")
EmReadScreen last_name, 25, 6, 30
last_name = replace(last_name, "_", "")

member_info = member_number & " - " & first_name & " " & last_name

Call navigate_to_MAXIS_screen("STAT", "WREG")
Call write_value_and_transmit(member_number, 20, 76)
EMReadScreen WREG_MEMB_check, 14, 24, 2
IF WREG_MEMB_check = "REFERE" OR WREG_MEMB_check = "MEMBER" THEN script_end_procedure("The member number that you entered is not valid.  Please check the member number, and start the script again.")

EMReadScreen panel_exists, 1, 2, 78
If panel_exists = "1" then	
	EMreadScreen wreg_status, 2, 8, 50
	EMReadScreen abawd_status, 2, 13, 50

	Call write_value_and_transmit("X", 13, 57)		'navigate to ABAWD/TLR Tracking panel and check for historical months

	'Resetting the variables
	bene_mo_col = (15 + (4*cint(MAXIS_footer_month)))		'col to search starts at 15, increased by 4 for each footer month
	bene_yr_row = 10
	abawd_counted_months = 0					'declares the variables values at 0 or blanks
	second_set_count = 0
	abawd_status = 0
	wreg_status = 0
	abawd_counted_months_string = ""
	second_set_string = ""
    abawd_info = ""
    second_set_info = ""

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
        If bene_yr_row = "10" then counted_date_year = right(DatePart("yyyy", ABAWD_eval_date), 2)
        If bene_yr_row = "9"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -1, ABAWD_eval_date)), 2)
        If bene_yr_row = "8"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -2, ABAWD_eval_date)), 2)
        If bene_yr_row = "7"  then counted_date_year = right(DatePart("yyyy", DateAdd("yyyy", -3, ABAWD_eval_date)), 2)
        
        'reading to see if a month is counted month or not
        EMReadScreen is_counted_month, 1, bene_yr_row, bene_mo_col
        
        'counting and checking for counted ABAWD months
        IF is_counted_month = "X" or is_counted_month = "M" THEN
        	EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
        	abawd_counted_months = abawd_counted_months + 1				'adding counted months
    		abawd_counted_months_string = abawd_counted_months_string & counted_date_month & "/" & counted_date_year & " | "
        END IF
        
        'counting and checking for counted 2nd set months
        IF is_counted_month = "Y" or is_counted_month = "N" THEN
        	EMReadScreen counted_date_year, 2, bene_yr_row, 15			'reading counted year date
        	second_set_count = second_set_count + 1				'adding counted months
    		second_set_string = second_set_string & counted_date_month & "/" & counted_date_year & " |"
        END IF
        
        bene_mo_col = bene_mo_col - 4		're-establishing search once the end of the row is reached
        IF bene_mo_col = 15 THEN
        	bene_yr_row = bene_yr_row - 1
        	bene_mo_col = 63
        END IF
    	'used to loop until count was 36 due to person based look back period. Now fixed clock starts 01/23 for all members. 
    LOOP until (counted_date_month = TLR_fixed_clock_mo AND counted_date_year = TLR_fixed_clock_yr)

    'cleaning up these variables for dialog display
    If trim(right(abawd_counted_months_string, 1)) = "|" THEN abawd_counted_months_string = left(abawd_counted_months_string, len(abawd_counted_months_strings) - 1)
    If trim(right(second_set_string, 1)) = "|" THEN second_set_string = left(second_set_string, len(second_set_string) - 1)
    PF3	' to exit tracking record 	   
  
    'Cleaning up output 
    If abawd_counted_months = 0 then 
        abawd_info = 0
    Else
        abawd_info = abawd_counted_months & " - " & abawd_counted_months_string
    End if 
    
    If second_set_count = 0 then 
        second_set_info = 0
    Else 
        second_set_info = second_set_count & " - " & second_set_string
    End if 
End if 

Do
    Do
        '----------------------------------------------------------------------------------------------------1st Dialog: SNAP Work Rules & TLR Exemptions 
        Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 326, 375, "SNAP Work Rules & Time Limited Exemptions"
          GroupBox 5, 5, 315, 55, "TLR Information for " & member_info
          Text 15, 20, 210, 20, "Counted TLR Months: " & abawd_info
          Text 15, 40, 230, 10, "Counted 2nd Set: " & second_set_info
          Text 5, 65, 315, 10, "=============================================================================="
          Text 25, 80, 275, 10, "Select ALL applicable exemptions for this member below (Exemptions Dialog 1 of 2)"
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
            PushButton 230, 15, 85, 15, "Exempt - CM 0028.06.12", CM_button
            PushButton 250, 35, 65, 15, "TLR - CM 0011.24", TLR_CM_button
          EndDialog

        Dialog Dialog1
        Cancel_Confirmation
        If ButtonPressed = CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00280612"
        If ButtonPressed = TLR_CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe	https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_001124"
    Loop until ButtonPressed = -1

    Do
        ''-------------------------------------------------------------------------------------------------2nd Dialog: Subject to SNAP Work Rules, NOT Time limits 
        Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 326, 320, "Subject to Work Rules, Exempt from Time Limits"
          GroupBox 5, 5, 315, 55, "TLR Information for " & member_info
          Text 15, 20, 210, 20, "Counted TLR Months: " & abawd_info
          Text 15, 40, 230, 10, "Counted 2nd Set: " & second_set_info
          Text 5, 65, 315, 10, "=============================================================================="
          Text 25, 80, 275, 10, "Select ALL applicable exemptions for this member below (Exemptions Dialog 2 of 2)"
          Text 5, 95, 315, 10, "=============================================================================="
          CheckBox 5, 110, 120, 10, "Age 55 - 59?*", age_exempt_checkbox
          CheckBox 5, 125, 125, 10, "Child under 18 in your SNAP unit?*", minor_hh_checkbox
          CheckBox 5, 140, 155, 10, "16-17 and NOT living with parent/caregiver?*", minor_wo_caregiver_checkbox
          CheckBox 5, 155, 45, 10, "Pregnant?*", PX_checkbox
          CheckBox 5, 170, 100, 10, "Served in the US Military?*", veteran_checkbox
          CheckBox 5, 185, 185, 10, "In foster care on 18th birthday AND under age 25?*", foster_care_checkbox
          CheckBox 5, 200, 255, 10, "RCA recipient and participating in Refugee Employment Services 1/2 time?", RCA_checkbox
          CheckBox 5, 215, 165, 10, "Responsible for the care of a dependent child?", dependent_child_checkbox
          CheckBox 5, 230, 235, 10, "Employed/Self-employed at least 20 hours/week or 80 hours/month?", working_20_checkbox
          CheckBox 5, 245, 260, 10, "Participating in an approved work/training program at least 20 hours/month?", approved_work_checkbox
          CheckBox 5, 260, 275, 10, "Volunteering OR combo of work, training, or volunteering at least 80 hours/month?", combo_work_checkbox
          Text 5, 280, 190, 10, "* = Can be declaratory unless inconsistent/questionable."
          ButtonGroup ButtonPressed
            PushButton 155, 295, 50, 15, "Previous", previous_button
            OkButton 210, 295, 50, 15
            CancelButton 265, 295, 50, 15
            PushButton 230, 15, 85, 15, "Exempt - CM 0028.06.12", CM_button
            PushButton 250, 35, 65, 15, "TLR - CM 0011.24", TLR_CM_button
        EndDialog
    
  	    Dialog Dialog1
  	    Cancel_without_confirmation
  	    If ButtonPressed = CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00280612"
        If ButtonPressed = TLR_CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe	https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_001124"
        If ButtonPressed = previous_button Then exit do
    Loop until ButtonPressed = -1
    If buttonPressed = -1 then exit do
Loop 

'Adding up the checks to see if we need to move onto the next dialog 
exempt_reasons = 0  'defaulting to 0
exempt_text = ""
verified_wreg = ""

'SNAP Work Rules & TLR Rules
If disa_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Physical illness, injury, disability or limitation.|" 
    verified_wreg = verified_wreg & "03|"
End if 
If perm_disa_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Temp or Perm DISA from SSA, VA, Work Comp, etc.|"
    verified_wreg = verified_wreg & "03|"
End if 
If sub_abuse_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Substance abuse or addiction dependency.|"
    verified_wreg = verified_wreg & "03|"
End if 
If mental_illness_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Mental illness, disorder, etc.|"
    verified_wreg = verified_wreg & "03|"
End if 
If homeless_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Homeless.|"
    verified_wreg = verified_wreg & "03|"
End if 
If dom_violence_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- A victim of domestic violence.|"
    verified_wreg = verified_wreg & "03|"
End if 
If care_of_hh_memb_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1   'case note who requires care 
    exempt_text = exempt_text & "- Caring for person who needs help caring for themselves.|"
    verified_wreg = verified_wreg & "04|"
End if 
If age_sixty_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Age 60 or older.|"
    verified_wreg = verified_wreg & "05|"
End if 
If under_sixteen_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Under the age of 16.|"
    verified_wreg = verified_wreg & "06|"
End if 
If sixteen_seventeen_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Aged 16 or 17 living w/ parent or caregiver.|"
    verified_wreg = verified_wreg & "07|"
End if 
If care_child_six_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1   'case note multiple people who need exemption AND if child under 6 is not in the HH
    exempt_text = exempt_text & "- Responsible for the care of a child under 6.|"
    verified_wreg = verified_wreg & "08|"
End if 
If employed_thirty_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Employed 30 hours/week or grossing at least $217.50/week ($935.25/month).|"
    verified_wreg = verified_wreg & "09|"
End if 
If matching_grant_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Receiving Matching Grant.|"
    verified_wreg = verified_wreg & "10|"
End if 
If unemployment_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Receiving or applied for unemployment insurance.|"
    verified_wreg = verified_wreg & "11|"
End if 
If enrolled_school_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Enrolled in school, training program, or higher education at least half time.|"
    verified_wreg = verified_wreg & "12|"
End if 
If CD_program_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Participating in drug or alcohol addiction treatment program.|"
    verified_wreg = verified_wreg & "13|"
End if 
If MFIP_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Receiving MFIP and in compliance with Employment Services.|"
    verified_wreg = verified_wreg & "14|"
End if
If DWP_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Receiving DWP and in compliance with Employment Services.|"
    verified_wreg = verified_wreg & "20|"
End if 

'Needs to follow SNAP Work Rules, but TLR exempt 
If minor_wo_caregiver_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- 16-17 and NOT living with parent/caregiver.|"
    verified_wreg = verified_wreg & "15|"
End if
If age_exempt_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "-  Age 55 - 59|"
    verified_wreg = verified_wreg & "16|"
End if 
If RCA_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- RCA recipient and participating in Refugee Employment Services 1/2 time.|"
    verified_wreg = verified_wreg & "17|"
End if 
If minor_hh_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "-  Child under 18 in your SNAP unit.|"
    verified_wreg = verified_wreg & "21|"
End if 
If PX_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Pregnant.|"
    verified_wreg = verified_wreg & "23|"
End if 
If dependent_child_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Responsible for the care of a dependent child.|"
    verified_wreg = verified_wreg & "21|"
End if 
If veteran_checkbox = 1 then 
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Served in the US Military.|"
    verified_wreg = verified_wreg & "30|"
    verified_abawd = "09"
End if 
If foster_care_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- In foster care on 18th birthday AND under age 25.|"
    verified_wreg = verified_wreg & "30|"
    verified_abawd = "09"
End if 
IF working_20_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Employed/Self-employed at least 20 hours/week or 80 hours/month.|"
    verified_wreg = verified_wreg & "30|"
    verified_abawd = "06"
End if 
IF approved_work_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Participating in an approved work/training program at least 20 hours/month.|"
    verified_wreg = verified_wreg & "30|"
    verified_abawd = "08"
End if
IF combo_work_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
    exempt_text = exempt_text & "- Volunteering OR combo of work, training, or volunteering at least 80 hours/month.|"
    verified_wreg = verified_wreg & "30|"
    verified_abawd = "08"
End if

exemption_array = split(exempt_text, "|")  'Splitting out array for final dialog 
 
If exempt_reasons > 0 then 
    exempt_elig = True 
    exempt_status = "Exempt" 
Elseif exempt_reasons = 0 then 
    exempt_elig = False 
    exempt_status = "Not Exempt"
End if   

'Dialog will only show up if NO exemptions have been found AND 3 ABAWD counted months have been used since the baseline date of 01/23.
If (exempt_elig = False and abawd_counted_months => 3) then 
    '----------------------------------------------------------------------------------------------------Second Set TLR
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 326, 180, "Second Set TLR Months"
      GroupBox 5, 5, 315, 55, "TLR Information for " & member_info
      Text 15, 20, 210, 20, "Counted TLR Months: " & abawd_info
      Text 15, 40, 230, 10, "Counted 2nd Set: " & second_set_info
      Text 5, 65, 315, 10, "=============================================================================="
      Text 60, 80, 190, 10, "Select ALL applicable situations for this member below"
      Text 5, 95, 315, 10, "=============================================================================="
      CheckBox 10, 105, 160, 10, "Used all 3 counted TLR months since 01/23?", Used_TLR_checkbox
      CheckBox 10, 120, 260, 10, "Worked at least 80 hours in a month SINCE closing for using 3 TLR months?", worked_80_since_closing
      CheckBox 10, 135, 250, 10, "Work/work activities have ended or reduced to less than 80 hours/month?", work_ended_checkbox
      ButtonGroup ButtonPressed
        'PushButton 10, 155, 110, 15, "View ABAWD Tracking Record", open_ATR_button
        OkButton 210, 155, 50, 15
        CancelButton 265, 155, 50, 15
        PushButton 230, 15, 85, 15, "Exempt - CM 0028.06.12", CM_button
        PushButton 250, 35, 65, 15, "TLR - CM 0011.24", TLR_CM_button
    EndDialog

    Do
	    Do
	        err_msg = ""
  		    Dialog Dialog1
  		    Cancel_confirmation
            If ButtonPressed = CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00280612"
            If ButtonPressed = TLR_CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe	https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_001124"
	    LOOP UNTIL ButtonPressed = -1
	    CALL check_for_password(are_we_passworded_out) 'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					 'loops until user passwords back in
End if

'determining 2nd set eligibility based on meeting all criteria in "Second Set TLR Months" dialog 
second_set_reasons = 0
If Used_TLR_checkbox = 1 then second_set_reasons = second_set_reasons + 1
If worked_80_since_closing = 1 then second_set_reasons = second_set_reasons + 1
If work_ended_checkbox = 1 then second_set_reasons = second_set_reasons + 1

If second_set_reasons = 3 then 
    second_set_elig = True 
    ss_elig_text = " is " 
Elseif second_set_reasons <> 3 then 
    second_set_elig = False 
    ss_elig_text = " is not "
End If 

'If resident doesn't meet any exemptions providing the option to case note.
Do 
    If exempt_elig = False then 
        case_note_confirmation = MsgBox(Member_info & " has been identified as not meeting an exemption, and" & ss_elig_text & "eligible for TLR/ABAWD 2nd Set Months. Do you want to CASE/NOTE this information?" & _  
        vbNewLine & vbNewLine & "Press No to end the script without CASE/NOTE.", vbInformation + vbYesNo, "Member appears to be a Time-Limited Recipient.")
        If case_note_confirmation = vbNo then script_end_procedure_with_error_report("You have opted out of case noting the TLR screening. The script has ended.")
    End if 
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'----------------------------------------------------------------------------------------------------Documentation about Exemptions Dialog 
If exempt_reasons > 0 then 
  'lists the exemptions in the dialog and adds mandatory fields based on the situation 
    BeginDialog Dialog1, 0, 0, 326, 350, "Document SNAP Work Rules and/or TLR Exemptions"
      GroupBox 5, 5, 315, 80, "TLR Information for " & Member_info
      Text 10, 20, 300, 20, "This member meets exemptions based on the screening completed. Documentation in CASE/NOTEs is required when exempting an individual from work rules."
      Text 10, 50, 300, 30, "Describe in detail how the resident meets the exemption(s), and if the exemption is obvious or verification was used. Also clearly CASE/NOTE when an exemption is applied based on your own observations or information obtained in conversation with the SNAP unit."
      Text 10, 95, 60, 10, "Exemption notes:"
      EditBox 70, 90, 250, 15, exemption_notes
      Text 10, 115, 60, 10, "Exemption basis:"
      ComboBox 70, 110, 100, 15, "Select OR Type..."+chr(9)+"Conversation w/ resident"+chr(9)+"Observational"+chr(9)+"Verified", exemption_basis
      CheckBox 10, 130, 310, 10, "Check here to update STAT/WREG with highest exemption for selected month - CM +1.", update_wreg_checkbox
      ButtonGroup ButtonPressed
        OkButton 215, 110, 50, 15
        CancelButton 270, 110, 50, 15
      Text 5, 145, 315, 10, "=============================================================================="
      Text 60, 160, 190, 10, "SNAP Work Rules and/or Time-Limited SNAP Exemptions"
      Text 5, 175, 315, 10, "=============================================================================="
      y_pos = 180
        For each exemption in exemption_array
            Text 15, (y_pos + 15), 275, 10, exemption
            y_pos = y_pos + 15
            If exemption = "- Caring for person who needs help caring for themselves." then 
                Text 20, y_pos + 15, 170, 10, "Name of person whom care is provided for:"  
                EditBox 165, y_pos +10, 150, 15, name_of_person_in_care
                y_pos = y_pos + 15
                needs_care_notes = True
            End if               
            If exemption = "- Responsible for the care of a child under 6." then 
                CheckBox 20, y_pos + 15, 145, 10, "Child is in the household.", child_in_HH_checkbox
                CheckBox 20, y_pos + 30, 200, 10, "More than one person in the SNAP HH using this exemption.", one_under6_checkbox
                y_pos = y_pos + 30
                needs_child_notes = True
            End if               
        Next 
    EndDialog
    
    Do
        Do 
            err_msg = ""
            Dialog Dialog1
            Cancel_confirmation
            If trim(exemption_notes) = "" then err_msg = err_msg & vbNewLine & "* Enter details about the TLR/ABAWD exemption."
            If exemption_basis = "Select OR Type..." then err_msg = err_msg & vbNewLine & "* Select or type the exemption basis that was determined."
            
            If needs_care_notes = True then 
                If trim(name_of_person_in_care) = "" then err_msg = err_msg & vbNewLine & "* Enter the name of the person whom care is being provide for."
            End if   

            If needs_child_notes = True then 
                If (child_in_HH_checkbox = 0 and len(exemption_notes) < 30) then err_msg = err_msg & vbNewLine & "* You must enter the full name and DOB for the child in care outside of the HH."
                If (one_under6_checkbox = 1 and len(exemption_notes) < 30) then err_msg = err_msg & vbNewLine & "* You must detail why more than one HH member needs this exemption."
            End If 
        
            If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
  	        If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	    LOOP UNTIL err_msg = ""
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End if 

military_exempt = False
If update_wreg_checkbox = 1 then 
	'filter the list here for best_wreg_code
	If trim(verified_wreg) = "" then 
        best_wreg_code = "30"
	Elseif len(verified_wreg) = 3 then
		best_wreg_code = replace(verified_wreg, "|", "")
	Else
        wreg_hierarchy = array("03","04","05","06","07","08","09","10","11","12","13","14","20","15","16","21","17","23","30")
        for each code in wreg_hierarchy
            If instr(verified_wreg, code) then
                best_wreg_code = code
                exit for
            End if
        next
	End if

	If best_wreg_code = "03" or _
		best_wreg_code = "04" or _
		best_wreg_code = "05" or _
		best_wreg_code = "06" or _
		best_wreg_code = "07" or _
		best_wreg_code = "08" or _
		best_wreg_code = "09" or _ 
		best_wreg_code = "10" or _
		best_wreg_code = "11" or _
		best_wreg_code = "12" or _
		best_wreg_code = "13" or _
		best_wreg_code = "14" or _
		best_wreg_code = "20" then
		    best_abawd_code = "01"
	End if

	If best_wreg_code = "15" then best_abawd_code = "02"
	If best_wreg_code = "16" then best_abawd_code = "03"
	If best_wreg_code = "21" then best_abawd_code = "04"
	If best_wreg_code = "17" then best_abawd_code = "12"
	If best_wreg_code = "23" then best_abawd_code = "05"

    'needs handling since the vets and foster care folks and 30/06 and 30/08 people also use this code 
    If best_wreg_code = "30" then 
		If verified_abawd = "" then
			best_abawd_code = "10"
		Else
			best_abawd_code = verified_abawd 'this is determined in the user selections 
        End if 
    End if 

    If instr(exempt_text, "Served in the US Military.") then
        If best_wreg_code = "30" then military_exempt = True 
    End if 

    Call date_array_generator(MAXIS_footer_month, MAXIS_footer_year, footer_month_array) 'Uses the custom function to create an array of dates from the initial_month and initial_year variables, ends at CM + 1.
    
    For item = 0 to ubound(footer_month_array)
	    MAXIS_footer_month = datepart("m", footer_month_array(item)) 'Need to assign footer month / year each time through
	    If len(MAXIS_footer_month) = 1 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
	    MAXIS_footer_year = right(datepart("YYYY", footer_month_array(item)), 2)
	    footer_string = MAXIS_footer_month & "/" & MAXIS_footer_year

        Call MAXIS_background_check

        'If the resident meets the military exemption status, the STAT/MEMI will be updated IF the field is there. Field was added for footer month/year of 07/2024. 
        If military_exempt = True then 
            Call navigate_to_MAXIS_screen("STAT", "MEMI")
            Call write_value_and_transmit(member_number, 20, 76)
            EMReadScreen military_field, 23, 12, 54
            If military_field = "Military Service (Y/N):" then 
                PF9
                Call write_value_and_transmit("Y", 12, 78)
                EMReadScreen date_error_msg, 22, 24, 2
                If date_error_msg = "ACTUAL DATE IS MISSING" then 
                    Call create_mainframe_friendly_date(date, 6, 35, 0)
                    transmit
                    EMReadScreen warning_msg, 7, 24, 2
                    If warning_msg = "WARNING" then transmit
                End If 
            End If     
        End If 

        Call navigate_to_MAXIS_screen("STAT", "WREG")
        Call write_value_and_transmit(member_number, 20, 76)
        EMReadScreen panel_exists, 1, 2, 78
        panel_date = cdate(MAXIS_footer_month & "/01/" & MAXIS_footer_year)
        If panel_date > cdate("6/30/2025") Then
            PWE_col = 70
            ET_col = 78
        Else
            PWE_col = 68
            ET_col = 80
        End If
        If panel_exists = "0" then 
            Call write_value_and_transmit("NN", 20, 79) 'Adding new WREG panel 
            EMWriteScreen "Y", 6, PWE_col 'defaulting PWE to Y if blank panel 
        Else 
            PF9
        End if 
    
	    EMWriteScreen best_wreg_code, 8, 50
	    EMWriteScreen best_abawd_code, 13, 50
	    If best_wreg_code = "30" then
            If best_abawd_code = "09" then 
                EMWriteScreen "Y", 8, ET_col
            Else 
	            EMWriteScreen "N", 8, ET_col
            End if 
	    Else
	        EMWriteScreen "_", 8, ET_col
	    End if
	    
	    EMReadScreen orientation_warning, 7, 24, 2 	'reading for orientation date warning message. This message has been causing me TROUBLE!!
	    If orientation_warning = "WARNING" then transmit 
	    PF3 'to save and exit to stat/wrap
    Next 
End if 

'----------------------------------------------------------------------------------------------------CASE/NOTE
Call start_a_blank_case_note
Call write_variable_in_CASE_NOTE("*~*~" & member_info & " TLR Screened: " & exempt_status & "~*~*")
Call write_bullet_and_variable_in_CASE_NOTE("Counted TLR Months", abawd_info)
Call write_bullet_and_variable_in_CASE_NOTE("Counted 2nd Set Months", second_set_info)
Call write_variable_in_CASE_NOTE("---")
If exempt_reasons = 0 then 
    Call write_variable_in_CASE_NOTE("* Member does not meet any exemption based on screening.")
Else 
    Call write_variable_in_CASE_NOTE("This member is eligible for the following exemption(s):")
    For each exemption in exemption_array
        Call write_variable_in_CASE_NOTE(exemption)
        If exemption = "- Caring for person who needs help caring for themselves." then Call write_variable_in_CASE_NOTE("    * Name of person whom care is provided for: " &  name_of_person_in_care)           
        If exemption = "- Responsible for the care of a child under 6." then            
            If child_in_HH_checkbox = 1 then Call write_variable_in_CASE_NOTE("    * Child under 6 is in the household.")    
            If child_in_HH_checkbox = 0 then Call write_variable_in_CASE_NOTE("    * Child under 6 is not in the household. See exemption notes for details.")    
            If one_under6_checkbox = 0 then Call write_variable_in_CASE_NOTE("    * Member is the only person using the child < 6 exemption.")  
            If one_under6_checkbox = 1 then Call write_variable_in_CASE_NOTE("    * More than one person in the SNAP HH using this exemption. See exemption notes for details.")     
        End if   
    Next 
End if 
Call write_bullet_and_variable_in_CASE_NOTE("Exemption Notes", exemption_notes)
If abawd_counted_months => 3 then 
    Call write_variable_in_CASE_NOTE("* Member has used all available counted TLR/ABAWD months.")
    Call write_variable_in_CASE_NOTE("* Member" & ss_elig_text & "eligible for TLR/ABAWD 2nd set months.")
End if 
If update_wreg_checkbox = 1 then Call write_variable_in_CASE_NOTE("* STAT/WREG panel has been updated with FSET/ABAWD codes: " & best_wreg_code & "/" & best_abawd_code & ".") 
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("Success! This member has been assessed for time-limited SNAP.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs --------------------------------------------------12/29/2023
'--Tab orders reviewed & confirmed-----------------------------------------------12/29/2023
'--Mandatory fields all present & Reviewed---------------------------------------12/29/2023
'--All variables in dialog match mandatory fields--------------------------------12/29/2023
'Review dialog names for content and content fit in dialog-----------------------12/29/2023
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog--------------------10/07/2024
'--Create a button to reference instructions-------------------------------------10/07/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)----------------------------------12/29/2023
'--CASE:NOTE Header doesn't look funky-------------------------------------------12/29/2023
'--Leave CASE:NOTE in edit mode if applicable------------------------------------12/29/2023
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used------12/29/2023
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed---------------------------------------12/29/2023
'--MAXIS_background_check reviewed (if applicable)-------------------------------12/29/2023
'--PRIV Case handling reviewed --------------------------------------------------12/29/2023
'--Out-of-County handling reviewed-----------------------------------------------12/29/2023---------------------N/A
'--script_end_procedures (w/ or w/o error messaging)-----------------------------12/29/2023
'--BULK - review output of statistics and run time/count (if applicable)---------12/29/2023---------------------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")------------12/29/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed ---------------------------------------------------12/29/2023
'--Incrementors reviewed (if necessary)------------------------------------------12/29/2023
'--Denomination reviewed --------------------------------------------------------12/29/2023
'--Script name reviewed----------------------------------------------------------12/29/2023
'--BULK - remove 1 incrementor at end of script reviewed-------------------------12/29/2023---------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete-----------------------------------------12/29/2023
'--comment Code------------------------------------------------------------------12/29/2023
'--Update Changelog for release/update-------------------------------------------12/29/2023
'--Remove testing message boxes--------------------------------------------------12/29/2023
'--Remove testing code/unnecessary code------------------------------------------12/29/2023
'--Review/update SharePoint instructions-----------------------------------------10/07/2024
'--Other SharePoint sites review (HSR Manual, etc.)------------------------------12/29/2023
'--COMPLETE LIST OF SCRIPTS reviewed---------------------------------------------10/07/2024
'--COMPLETE LIST OF SCRIPTS update policy references-----------------------------12/29/2023
'--Complete misc. documentation (if applicable)----------------------------------12/29/2023
'--Update project team/issue contact (if applicable)-----------------------------12/29/2023
