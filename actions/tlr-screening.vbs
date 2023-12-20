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
  	    If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call MAXIS_footer_month_confirmation	'making sure we're getting to current month/year
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged, and you do not have access. The script will now end.")

Call write_value_and_transmit(member, 20, 76)
EmReadScreen first_name, 12, 6, 63
first_name = replace(first_name, "_", "")
EmReadScreen last_name, 25, 6, 30
last_name = replace(last_name, "_", "")

member_info = member_number & " - " & first_name & " " & last name

Call navigate_to_MAXIS_screen("STAT", "WREG")
Call write_value_and_transmit(member_number, 20, 76)
EMReadScreen WREG_MEMB_check, 6, 24, 2
IF WREG_MEMB_check = "REFERE" OR WREG_MEMB_check = "MEMBER" THEN script_end_procedure("The member number that you entered is not valid.  Please check the member number, and start the script again.")

EMReadScreen panel_exists, 1, 2, 78
If panel_exists = "1" then	
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
    
    abawd_info = abawd_counted_months & " - " & abawd_counted_months_string
    second_set_info = second_set_count & " - " & second_set_string
End if 

Do
	
	  Do
        '----------------------------------------------------------------------------------------------------1st Dialog: SNAP Work Rules & TLR Exemptions 
        Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 326, 375, "SNAP Work Rules & Time Limited Exemptions"
          GroupBox 5, 5, 315, 55, "TLR Information for " & member_info
          Text 15, 20, 235, 10, "Counted TLR Months: " & abawd_info
          Text 15, 40, 235, 10, "Counted 2nd Set: " & second_set_info
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
          Text 15, 20, 235, 10, "Counted TLR Months:  & abawd_counted_months &  -  & abawd_counted_months_string"
          Text 15, 40, 235, 10, "Counted 2nd Set:  & second_set_count &  -  & second_set_string"
          Text 5, 65, 315, 10, "=============================================================================="
          Text 60, 80, 190, 10, "Select ALL applicable exemptions for this member below"
          Text 5, 95, 315, 10, "=============================================================================="
          CheckBox 5, 110, 120, 10, "Younger than 18 OR 53 or older?*", age_exempt_checkbox
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
            PushButton 260, 15, 55, 15, "CM 0028.06.12", CM_button
            PushButton 260, 35, 55, 15, "CM 0011.24", TLR_CM_button
        EndDialog

  		  Dialog Dialog1
  		  Cancel_without_confirmation
  		  If ButtonPressed = CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00280612"
        If ButtonPressed = TLR_CM_button then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe	https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_001124"
        If ButtonPressed = previous_button Then exit do
    Loop until ButtonPressed = -1 or ButtonPressed = previous_button
	  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Adding up the checks to see if we need to move onto the next dialog 
exempt_reasons = 0  'defaulting to 0
exempt_text = exempt_text & = ""

If disa_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = exempt_text = exempt_text & & "- Physical illness, injury, disability or limitation."
End if 
If perm_disa_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Temp or Perm DISA from SSA, VA, Work Comp, etc."
End if 
If sub_abuse_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Substance abuse or addiction dependency"
End if 
If mental_illness_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Mental illness, disorder, etc."
End if 
If homeless_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Homeless."
End if 
If dom_violence_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- A victim of domestic violence"
End if 
If care_of_hh_memb_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1   'case note who requires care 
  exempt_text = exempt_text & = "- Caring for person who needs help caring for themselves."
End if 
If care_child_six_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1   'case note multiple people who need exemption AND if child under 6 is not in the HH
  exempt_text = exempt_text & = "- Responsible for the care of a child under 6."
End if 
If age_sixty_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Age 60 or older."
End if 
If under_sixteen_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Under the age of 16."
End if 
If sixteen_seventeen_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Aged 16 or 17 living w/ parent or caregiver."
End if 
If employed_thirty_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Employed 30 hours/week or grossing at least $217.50/week ($935.25/month)."
End if 
If unemployment_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Receiving or applied for unemployment insurance."
End if 
If matching_grant_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Receiving Matching Grant."
End if 
If DWP_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Receiving DWP and in compliance with Employment Services."
End if 
If MFIP_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Receiving MFIP and in compliance with Employment Services."
End if 
If enrolled_school_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Enrolled in school, training program, or higher education at least half time."
End if 
If CD_program_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Participating in drug or alcohol addiction treatment program."
End if 
If age_exempt_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "-  Younger than 18 OR 53 or older."
End if 
If minor_hh_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "-  Child under 18 in your SNAP unit."
End if 
If minor_wo_caregiver_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- 16-17 and NOT living with parent/caregiver."
End if 
If PX_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Pregnant."
End if 
If veteran_checkbox = 1 then 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Served in the US Military."
End if 
If foster_care_checkbox = 1 
  exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- In foster care on 18th birthday AND under age 25."
End if 
If RCA_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- RCA recipient and participating in Refugee Employment Services 1/2 time."
End if 
If dependent_child_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Responsible for the care of a dependent child."
End if 
IF working_20_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Employed/Self-employed at least 20 hours/week or 80 hours/month."
End if 
IF approved_work_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Participating in an approved work/training program at least 20 hours/month."
End if
IF combo_work_checkbox = 1 then
    exempt_reasons = exempt_reasons + 1
  exempt_text = exempt_text & = "- Volunteering OR combo of work, training, or volunteering at least 80 hours/month."
End if

exempt_array = split(exempt_text, ".")  'Splitting out array for final dialog 
 
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
        OkButton 210, 155, 50, 15
        CancelButton 265, 155, 50, 15
        PushButton 260, 15, 55, 15, "CM 0028.06.12", CM_button
        PushButton 260, 35, 55, 15, "CM 0011.24", TLR_CM_button
    EndDialog

    Do
	      Do
	          err_msg = ""
  		      Dialog Dialog1
  		      Cancel_confirmation
	      LOOP UNTIL ButtonPressed = -1
	      CALL check_for_password(are_we_passworded_out) 'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					 'loops until user passwords back in
End if

'determining 2nd set eligibility based on meeting all criteria in "Second Set TLR Months" dialog 
2nd_set_reasons = 0  'defaulting to 0
If Used_TLR_checkbox = 1 then 2nd_set_reasons = 2nd_set_reasons + 1
If worked_80_since_closing = 1 then 2nd_set_reasons = 2nd_set_reasons + 1
If work_ended_checkbox = 1 then 2nd_set_reasons = 2nd_set_reasons + 1

If 2nd_set_reasons = 3 then 
    2nd_set_elig = True 
    ss_elig_text = " is " 
Elseif 2nd_set_reasons <> 3 then 
    2nd_set_elig = False 
    ss_elig_text = " is not "
End If 

'If resident doesn't meet any exemptions providing the option to case note.
If exempt_elig = False then 
    case_note confirmation = MsgBox(Member_info & " has been identified as not meeting an exemption, and " & ss_elig_text & "eligible for TLR/ABAWD 2nd Set Months. Do you want to CASE/NOTE this information?" & vbNewLine & vbNewLine & _
    "Press No to end the script without CASE/NOTE."                                                  
    vbYesNo & vbInformation, "Member appears to be a Time-Limited Recipient.")
    If case_note confirmation = vbNo then script_end_procedure_with_error_report("You have opted out of case noting the TLR screening. The script has ended.")
End if 

If exempt_reasons > 0 then 
  'list the exemptions in the dialog and add mandatory fields based on the situation 
  '----------------------------------------------------------------------------------------------------Documentation about Exemptions Dialog 
    BeginDialog , 0, 0, 326, 350, "Document SNAP Work Rules and/or TLR Exemptions"
      GroupBox 5, 5, 315, 80, "TLR Information for & Member_info"
      Text 10, 20, 300, 20, "This member meets exemptions based on the screening completed. Documentation in CASE/NOTEs is required when exempting an individual from work rules."
      Text 10, 50, 300, 30, "Describe in detail how the resident meets the exemption(s), and if the exemption is obvious or verification was used. Also clearly CASE/NOTE when an exemption is applied based on your own observations or information obtained in conversation with the SNAP unit."
      Text 10, 95, 60, 10, "Exemption notes:"
      EditBox 70, 90, 250, 15, exemption_notes
      Text 10, 115, 60, 10, "Exemption basis:"
      ComboBox 70, 110, 100, 15, "Select OR Type..."+chr(9)+"Conversation w/ resident"+chr(9)+"Observational"+chr(9)+"Verified", exemption_basis
      CheckBox 10, 130, 205, 10, "Check here to update STAT/WREG with highest exemption.", update_wreg_checkbox
      ButtonGroup ButtonPressed
        OkButton 215, 110, 50, 15
        CancelButton 270, 110, 50, 15
      Text 5, 145, 315, 10, "=============================================================================="
      Text 60, 160, 190, 10, "SNAP Work Rules and/or Time-Limited SNAP Exemptions"
      Text 5, 175, 315, 10, "=============================================================================="
      y_pos = 180
      For each exemption in exemption_array
          Text 15, (y-pos + 15), 200, 10, exemption
          y-pos + 15
          If exemption = "- Caring for person who needs help caring for themselves." then 
              Text 20, y_pos, 170, 10, "Name of person whom care is provided for:"  
              EditBox 165, y_pos, 150, 15, name_of_person_in_care
              y_pos = y_pos + 15
          End if               
          If exemption = "- Responsible for the care of a child under 6." then 
              CheckBox 20, y_pos, 145, 10, "Child is in the household.", child_in_HH_checkbox
              CheckBox 20, y_pos + 15, 200, 10, "More than one person in the SNAP HH using this exemption.", one_under6_checkbox
              y_pos = y_pos + 15
          End if               
      Next 
    EndDialog
  Do 
      Do 
          err_msg = ""
          Dialog Dialog1
          Cancel_confirmation
          If exemption_basis = "Select OR Type..." then err_msg = err_msg & vbNewLine & "* Select or type the exemption basis that was determined."
          If (exemption = "- Caring for person who needs help caring for themselves." and trim(name_of_person_in_care) = "") then err_msg = err_msg & vbNewLine & "* Enter the name of the person whom care is being provide for."
          If exemption = "- Responsible for the care of a child under 6." then 
          If (child_in_HH_checkbox = 0 and len(exemption_notes) < 30) then err_msg = err_msg & vbNewLine & "* You must enter the full name and DOB for the child in care outside of the HH."
          If (one_under6_checkbox = 1 and len(exemption_notes) < 30) then err_msg = err_msg & vbNewLine & "* You must detail "
          If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
  	      If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	    LOOP UNTIL err_msg = ""
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  Loop until are_we_passworded_out = false					'loops until user passwords back in
End if 


If update_wreg_checkbox = 1 then 
  'TODO: Add actions for update_wreg_checkbox here
End if 

'----------------------------------------------------------------------------------------------------CASE/NOTE
Call start_a_blank_case_note
Call write_variable_in_CASE_NOTE("*-*-" & member_info & " TLR Screened: " & exempt_status & " -*-*")
Call write_bullet_and_variable_in_CASE_NOTE("Counted TLR Months", abawd_info)
Call write_bullet_and_variable_in_CASE_NOTE("Second Set Months", second_set_info)
Call write_variable_in_CASE_NOTE("---")

'add what ever case note here !!
Call write_variable_in_CASE_NOTE("* Member does not meet any exemption based on screening.")
Call write_variable_in_CASE_NOTE("* Member " & ss_elig_text " eligible for TLR/ABAWD 2nd set months.")
If abawd_counted_month => 3 then Call write_variable_in_CASE_NOTE("* Member has used all available counted TLR/ABAWD months.")
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report("Success! This member has been assessed for time-limited SNAP.")