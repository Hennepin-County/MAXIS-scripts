'Informational front-end message, date dependent.
If datediff("d", "04/02/2013", now) < 5 then MsgBox "This script has been updated as of 04/02/2013! There's now a checkbox for starting the denied programs script right from this one."

'STATS GATHERING-----------------------------------------------------------------------------------
name_of_script = "MNSURE - 2014 retro task"
start_time = timer

'LOADING ROUTINE FUNCTIONS-------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("U:\PHHS\BlueZoneScripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

Set MNSURE_FUNCTIONS_fso = CreateObject("Scripting.FileSystemObject")
Set fso_MNSURE_FUNCTIONS_command = MNSURE_FUNCTIONS_fso.OpenTextFile("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\MNSURE FUNCTIONS FILE.vbs")
MNSURE_FUNCTIONS_contents = fso_MNSURE_FUNCTIONS_command.ReadAll
fso_MNSURE_FUNCTIONS_command.Close
Execute MNSURE_FUNCTIONS_contents 

'DIALOGS-------------------------------------------------------------------------------------------

BeginDialog MNsure_2014_retro_coverage, 0, 0, 169, 139, "MNsure 2014 Retro Coverage Task"
  EditBox 98, 4, 65, 12, case_number
  EditBox 98, 19, 65, 12, integrated_case_number
  EditBox 98, 34, 64, 12, appl_date
  EditBox 98, 49, 65, 12, info_received_date
  DropListBox 98, 65, 65, 12, "RFI"+chr(9)+"Phone Call"+chr(9)+"Not Received", retro_received_via
  DropListBox 98, 83, 25, 13, "0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"8"+chr(9)+"9", number_of_gap_months
  DropListBox 98, 101, 25, 13, "0"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3", number_of_retro_months
  ButtonGroup ButtonPressed
    OkButton 112, 122, 20, 12
    CancelButton 135, 122, 30, 12
  Text 3, 5, 68, 10, "Maxis Case Number"
  Text 3, 21, 80, 10, "Integrated Case Number"
  Text 3, 37, 80, 10, "MNsure Application Date"
  Text 3, 52, 90, 10, "Date Retro Info Received"
  Text 3, 66, 79, 8, "Retro Info Received Via"
  Text 3, 80, 81, 16, "How many gap months are needed?"
  Text 3, 99, 81, 16, "How many retro months are being requested?"
EndDialog

BeginDialog testing_for_other_programs, 0, 0, 127, 86, "Other Program Test"
  ButtonGroup ButtonPressed
	OkButton 70, 71, 20, 12
	CancelButton 94, 71, 30, 12
  Text 4, 4, 125, 20, "Are any of the following programs currently active/pending on this case?"
  CheckBox 67, 27, 53, 10, "Food Support", food_support
  CheckBox 8, 27, 48, 10, "MSA/Cash", msa_support
  CheckBox 8, 40, 50, 10, "Emergency", emer_support
  CheckBox 67, 54, 30, 10, "None", no_other_support
  CheckBox 67, 41, 45, 10, "GRH", grh_support
  CheckBox 8, 53, 45, 10, "IV-E", iv_e_support
EndDialog

BeginDialog elig_prompts_complete_screen, 0, 0, 178, 108, "Prompts Complete"
  ButtonGroup ButtonPressed
	PushButton 5, 92, 80, 12, "Quick Fail Person Test", fail_person_test
	OkButton 115, 92, 20, 12
	CancelButton 138, 92, 30, 12
  Text 4, 5, 174, 24, "It appears you have completed all the prompts for this budget period. If you believe this information is correct please check the box below and select the OK option below."
  CheckBox 10, 37, 146, 10, "Yes, all screen prompts appear complete.", all_screens_complete
  GroupBox 5, 54, 157, 30, "CAUTION"
  Text 10, 64, 149, 17, "Do NOT use your transmit key on this screen at this time. The script will do this for you."
EndDialog	

BeginDialog case_note_decision_dialog, 0, 0, 230, 135, "Review and Approve Results"
	CheckBox 10, 45, 111, 10, "Add case note upon completion", case_note_election
	DropListBox 75, 59, 75, 13, "Approved"+chr(9)+"Denied"+chr(9)+"Pending"+chr(9)+"Withdrawn", case_note_status
	EditBox 75, 75, 75, 12, case_note_worker_signature
	EditBox 10, 100, 210, 12, case_note_additional_comments
	ButtonGroup ButtonPressed
		OkButton 175, 121, 20, 12
		CancelButton 198, 121, 30, 12
		PushButton 117, 121, 55, 12, "Refresh Dialog", refresh_dialog
	GroupBox 5, 35, 220, 82, "Case Note"
	Text 3, 3, 224, 25, "Please review the results and make any changes/corrections as necessary. Once complete, approve your results and you may select below to automatically case note."
	Text 10, 76, 62, 8, "Worker Signature :"
	Text 75, 90, 78, 8, "Additional worker notes"
	Text 10, 61, 42, 8, "Case Status:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------

EMConnect ""

'Finds the case number in a case
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Starts the first dialog
  Do
    Do
      Dialog MNsure_2014_retro_coverage
      If buttonpressed = 0 then stopscript
	  call maxis_dater(appl_date,appl_date,"Application Date")
	  call maxis_dater(info_received_date,info_received_date,"Information Received Date")
	  call retro_calculator(number_of_retro_months,number_of_gap_months,appl_date,retro_date)
	  If case_number = "" then MsgBox "You must have a case number to continue!"
	  If appl_date = "" then MsgBox "You must enter an application date to continue!"
	Loop until case_number <> "" and appl_date <> ""
    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
  call back_to_self
  call change_footer_month(Left(appl_date,2),Right(appl_date,2))
  call navigate_to_screen("stat","memb")
  call HH_member_custom_dialog(HH_member_array)
  
        '-------- CASE - CURR -------------------

  call navigate_to_screen("case","curr")

  Do
    Dialog testing_for_other_programs
      If buttonpressed = 0 then stopscript
    If food_support = "0" and msa_support = "0" and iv_e_support = "0" and emer_support = "0" and grh_support = "0" and no_other_support = "0" then MsgBox "You must make a selection to continue."
  Loop until (no_other_support = "1" and food_support = "0" and msa_support = "0" and grh_support = "0" and emer_support = "0" and iv_e_support = "0") or _
			 (msa_support  = "1" and no_other_support = "0") or _
			 (iv_e_support = "1" and no_other_support = "0") or _
			 (emer_support = "1" and no_other_support = "0") or _
			 (grh_support  = "1" and no_other_support = "0") or _
			 (food_support = "1" and no_other_support = "0")
       
	   '-------- STAT - TYPE --------------------
	call navigate_to_screen("stat", "type")
	call create_or_edit_panel("case","n/a")

	For each HH_member in HH_member_array
	  row = 1
	  col = 1
	  EMSearch HH_member, row, col
	  If row <> 0 and  msa_support  = "0" then EMWriteScreen "n", row, (col + 25)
	  If row <> 0 then EMWriteScreen  "Y", row, col + 34
	  If row <> 0 and  food_support = "0" then EMWriteScreen "n", row, (col + 43)
	  If row <> 0 and  emer_support = "0" then EMWriteScreen "n", row, (col + 52)
	  If row <> 0 and  grh_support  = "0" and  HH_member = "01" then EMWriteScreen "n", row, (col + 61)
	  If row <> 0 and  iv_e_support = "0" and  HH_member = "01" then EMWriteScreen "n", row, (col + 70)      
    Next 
	transmit
	
	call screen_error_check()
	
        '-------- STAT - PROG -------------------
    call navigate_to_screen("stat","prog")
	call create_or_edit_panel("case","n/a")

    row = 1
    col = 1    
    EMSearch "HC:", row, col
      If row <> 0 then EMWriteScreen appl_date, row, col + 4
	  If row <> 0 then EMWriteScreen "__ __ __", 12, 55
    transmit

        '-------- STAT - HCRE -------------------
    call navigate_to_screen("STAT","HCRE")
	call create_or_edit_panel("case","n/a")

    For each HH_member in HH_member_array
      row = 1      
      col = 1
      EMSearch "* "&HH_member, row, col      
	  If row <> 0 then EMWriteScreen appl_date, row, (col + 29)
	  If row <> 0 then EMWriteScreen retro_date, row, (col + 42)	  
      EMReadScreen rcvd_date_exists, 8, 10, 73
		rcvd_date_exists = replace(replace(trim(rcvd_date_exists), " ", "/"), "_", "")
	  If len(rcvd_date_exists) <> 8 then
	    If row <> 0 then EMWriteScreen Left(info_received_date,2), 10, 73
		If row <> 0 then EMWriteScreen Mid(info_received_date,4,2), 10, 76
		If row <> 0 then EMWriteScreen Right(info_received_date,2), 10, 79
	  End If
    Next
    Transmit
	
	call screen_error_check()

        '-------- STAT - REVW -------------------
    call navigate_to_screen("STAT","REVW")
	call create_or_edit_panel("case","n/a")

    EMWriteScreen "X", 5, 71
    transmit
      call add_months(6,appl_date,i_r_d)
	  EMWriteScreen Left(i_r_d,2), 8, 27
	  EMWriteScreen Right(i_r_d,2), 8, 33
	  call add_months(12,appl_date,e_r_d)
      EMWriteScreen Left(e_r_d,2), 9, 27
	  EMWriteScreen Right(e_r_d,2), 9, 33
    PF3
    transmit

        '-------- STAT - MSUR -------------------
    call navigate_to_screen("STAT","MSUR") 
	For each HH_member in HH_member_array
		EMWriteScreen HH_member, 20, 76
		transmit
		call create_or_edit_panel("person",HH_member)

		EMWriteScreen Left(appl_date,2), 7, 36
		EMWriteScreen "01", 7, 39
		EMWriteScreen "20"&Right(appl_date,2), 7, 42
		transmit
    Next     
  
        '-------- STAT - WRAP(BGTX) -------------

If no_other_support = "1" then
	'No Other Active Programs
	call run_file("M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\RUNTASKADDITION - 2014 retro(No Active Program Protocol).vbs")
ElseIf no_other_support = "0" then
	'Other Active Programs
	call send_case_through_background("no_update")
End If

        '-------- ELIG - HC ---------------------

call navigate_to_screen("ELIG","HC")
  EMReadScreen stat_error_check, 16, 24, 5
  stat_error_check = trim(stat_error_check)
  If stat_error_check = "STAT EDITS EXIST" then
    Do
      call stat_error_scanner
      call navigate_to_screen("ELIG","HC")
      EMReadScreen stat_error_check, 16, 24, 5
      stat_error_check = trim(stat_error_check)
    Loop until stat_error_check <> "STAT EDITS EXIST"
  ElseIF stat_error_check <>	 "STAT EDITS EXIST" then

  	For each HH_member in HH_member_array
	  call back_to_self
	  call change_footer_month(Left(appl_date,2),Right(appl_date,2))
	  call navigate_to_screen("ELIG","HC")
	  call change_footer_month(Left(retro_date,2),Right(retro_date,2))
	  call command_line("ignore",HH_member,"NN")
	  EMWriteScreen "X", 8, 8
	  transmit
	  EMWriteScreen "06", 9, 23  
	  transmit
	  row = 1
	  col = 1
	  EMSearch "  "&HH_member&"  ", row, col
	  EMWriteScreen "X", row, col + 25
	  transmit

    Do  
	  PF9
      EMWriteScreen "04", 11, 26
      transmit
	  
	  '---Setting Variables ---------------------
	  
	  Dim total_hh_count, current_client_selected, converted_income_type
	  Dim set_included_income_01, set_included_income_02, set_included_income_03, set_included_income_04, set_included_income_05, set_included_income_06
	  Dim set_included_income_07, set_included_income_08, set_included_income_09, set_included_income_10, set_included_income_11, set_included_income_12
	  Dim earned_income_01, earned_income_02, earned_income_03, earned_income_04, earned_income_05, earned_income_06
	  Dim earned_income_07, earned_income_08, earned_income_09, earned_income_10, earned_income_11, earned_income_12
	  Dim unearned_income_01, unearned_income_02, unearned_income_03, unearned_income_04, unearned_income_05, unearned_income_06
	  Dim unearned_income_07, unearned_income_08, unearned_income_09, unearned_income_10, unearned_income_11, unearned_income_12
	  Dim client_01_name, client_02_name, client_03_name, client_04_name, client_05_name, client_06_name
	  Dim client_07_name, client_08_name, client_09_name, client_10_name, client_11_name, client_12_name  
	  Dim hh_count_01_searchable, hh_count_02_searchable, hh_count_03_searchable, hh_count_04_searchable, hh_count_05_searchable, hh_count_06_searchable
	  Dim hh_count_07_searchable, hh_count_08_searchable, hh_count_09_searchable, hh_count_10_searchable, hh_count_11_searchable, hh_count_12_searchable
	  Dim earned_type_1, earned_type_2, earned_type_3, unearned_type_1, unearned_type_2, unearned_type_3
	  Dim earned_value_1, earned_value_2, earned_value_3, unearned_value_1, unearned_value_2, unearned_value_3
	  Dim earned_exclusion_1, earned_exclusion_2, earned_exclusion_3, unearned_exclusion_1, unearned_exclusion_2, unearned_exclusion_3
	  
	  '------------------------------------------
	
	Call budget_month_config(1,budget_one)
	Call budget_month_config(2,budget_two)
	Call budget_month_config(3,budget_three)
	Call budget_month_config(4,budget_four)
	Call budget_month_config(5,budget_five)
	Call budget_month_config(6,budget_six)

	col = 21
	For i = 6 To 1 Step -1
		EMWriteScreen "X", 8, col
		EMWriteScreen "X", 9, col
		col = col + 11
	Next
    transmit
      call elig_member_selection()
	  Do
  	    EMReadScreen elig_hh_count_prompt, 15, 3, 32
	      elig_hh_count_prompt = trim(elig_hh_count_prompt)
	    EMReadScreen elig_abud_prompt, 4, 3, 47
	      elig_abud_prompt = trim(elig_abud_prompt)
	    EMReadScreen elig_cbud_prompt, 4, 3, 54
	      elig_cbud_prompt = trim(elig_cbud_prompt)
  			
		'Process HH Count Screen-----------------
	  	If elig_hh_count_prompt = "Household Count" then
	      If Len(total_hh_count) = 1 < 10 then
	        EMWriteScreen total_hh_count, 5, 69
	      ElseIf total_hh_count > 9 then 
	        EMWriteScreen total_hh_count, 5, 68
	      End If
	      If set_included_income_01 = "Y" or set_included_income_01 = "N" then EMWriteScreen set_included_income_01, 12, 61
	      If set_included_income_02 = "Y" or set_included_income_02 = "N" then EMWriteScreen set_included_income_02, 13, 61
	      If set_included_income_03 = "Y" or set_included_income_03 = "N" then EMWriteScreen set_included_income_03, 14, 61
	      If set_included_income_04 = "Y" or set_included_income_04 = "N" then EMWriteScreen set_included_income_04, 15, 61
	      If set_included_income_05 = "Y" or set_included_income_05 = "N" then EMWriteScreen set_included_income_05, 16, 61
	      If set_included_income_06 = "Y" or set_included_income_06 = "N" then EMWriteScreen set_included_income_06, 17, 61
	      If set_included_income_07 = "Y" or set_included_income_07 = "N" then EMWriteScreen set_included_income_07, 18, 61
	      If set_included_income_08 = "Y" or set_included_income_08 = "N" then EMWriteScreen set_included_income_08, 19, 61
	      If set_included_income_09 = "Y" or set_included_income_09 = "N" then EMWriteScreen set_included_income_09, 20, 61
	      If set_included_income_10 = "Y" or set_included_income_10 = "N" then EMWriteScreen set_included_income_10, 21, 61
	      If set_included_income_11 = "Y" or set_included_income_11 = "N" then EMWriteScreen set_included_income_11, 22, 61
	      If set_included_income_12 = "Y" or set_included_income_12 = "N" then EMWriteScreen set_included_income_12, 23, 61
		  transmit
		  EMReadScreen elig_hh_count_prompt, 15, 3, 32
	        elig_hh_count_prompt = trim(elig_hh_count_prompt)
		  If elig_hh_count_prompt = "Household Count" then
		    transmit	
		  ElseIf elig_hh_count_prompt <> "Household Count" then
		    elig_hh_count_prompt = "Household Count"
		  End If
		  EMReadScreen invalid_hh_size, 18, 23, 5
	        invalid_hh_size = trim(invalid_hh_size)
		  If invalid_hh_size = "HH SIZE IS INVALID" then transmit
		  
		     'Process ABUD Screen----------------
			 
  	    ElseIf elig_abud_prompt = "ABUD" then
		  EMWriteScreen "N", 5, 63
		  member_span = 1
		  For i = 1 To 12 Step 1
			call enter_income_information(member_span)
			member_span = member_span + 1
		  Next		  		  
		  transmit
		  
		   'Process CBUD Screen------------------
	
  	    ElseIf elig_cbud_prompt = "CBUD" then
		  EMWriteScreen "N", 5, 63
		  member_span = 1
		  For i = 1 To 12 Step 1
			call enter_income_information(member_span)
			member_span = member_span + 1
		  Next		  		  
		  transmit

		   'Review Results-----------------------
		   
	    ElseIf elig_hh_count_prompt <> "Household Count" and elig_abud_prompt <> "ABUD" and elig_cbud_prompt <> "CBUD" then
		  EMReadScreen HH_count_error, 51, 24, 2		  
			HH_count_error = trim(HH_count_error)
	  
		  If HH_count_error = "SELECT HH COUNT TO ENTER A HOUSEHOLD SIZE FOR MONTH" then
			col = 21
			For i = 6 To 1 Step -1
				EMWriteScreen "X", 8, col
				col = col + 11
			Next
			transmit
		  ElseIf HH_count_error <> "SELECT HH COUNT TO ENTER A HOUSEHOLD SIZE FOR MONTH" then
		  all_screens_complete = 0
		  Dialog elig_prompts_complete_screen
	  	    If buttonpressed = 0 then stopscript
			If buttonpressed = fail_person_test then
			  Do
			    call mnsure_fail_person_test
				Dialog elig_prompts_complete_screen
				  If buttonpressed = 0 then stopscript
			  Loop until buttonpressed = -1
			End If
		  End If
		End If
	  Loop until elig_hh_count_prompt <> "Household Count" and elig_abud_prompt <> "ABUD" and elig_cbud_prompt <> "CBUD" and all_screens_complete = 1
	  call current_date(curr_date)
	  call add_months(1,curr_date,current_month_plus_one)
	  current_month_plus_one = Left(current_month_plus_one,2)	  
	  If Left(budget_one,2)   <> current_month_plus_one or _
	     Left(budget_two,2)   <> current_month_plus_one or _
		 Left(budget_three,2) <> current_month_plus_one or _
		 Left(budget_four,2)  <> current_month_plus_one or _
		 Left(budget_five,2)  <> current_month_plus_one or _
		 Left(budget_six,2)   <> current_month_plus_one then        
		PF3
	    call change_footer_month(current_month_plus_one,Right(curr_date,2))
      End If
	  row = 1
	  col = 1
	  EMSearch "  "&HH_member&"  ", row, col
	  EMWriteScreen "X", row, col + 25
	  transmit
	Loop until Left(budget_one,2)   = current_month_plus_one or _
			   Left(budget_two,2)   = current_month_plus_one or _
			   Left(budget_three,2) = current_month_plus_one or _
			   Left(budget_four,2)  = current_month_plus_one or _
			   Left(budget_five,2)  = current_month_plus_one or _
			   Left(budget_six,2)   = current_month_plus_one
	Next
	PF3
  End If

'Case Note---------------
 
Dialog case_note_decision_dialog
If buttonpressed = 0 then stopscript
If buttonpressed = refresh_dialog then
Do
	Dialog case_note_decision_dialog
		If buttonpressed = 0 then stopscript
Loop until buttonpressed = -1
End If

If case_note_election = 1 then
	Do
		call navigate_to_screen("case", "note")
		PF9
		EMReadScreen mode_check, 7, 20, 3
		If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
	Loop until mode_check = "Mode: A" or mode_check = "Mode: E"
	
    EMSendKey "MNSURE: 2014 Retro " & case_note_status & "<newline>"
	If integrated_case_number <> "" then call write_new_line_in_case_note("-Integrated case # "&integrated_case_number)
	If case_number <> "" then call write_new_line_in_case_note("-MAXIS case # "&case_number)
	If appl_day <> "" and appl_month <> "" and appl_year <> "" then call write_new_line_in_case_note("-MNsure application date: "&appl_month&"/"&appl_day&"/"&appl_year)
	If info_received_day <> "" and info_received_month <> "" and info_received_year <> "" and retro_received_via <> "Not Received" then call write_new_line_in_case_note("-Retro Info recvd via "&retro_received_via&" on "&info_received_month&"/"&info_received_day&"/"&info_received_year&".")
	If retro_received_via = "Not Received" then call write_new_line_in_case_note("-Retro Info was not recvd. Case has been " & case_note_status)
	call write_new_line_in_case_note("-Retro months requested: "&number_of_retro_months)
	call write_new_line_in_case_note("-Gap months requested: "&number_of_gap_months)
	call write_new_line_in_case_note("**********")
	If case_note_additional_comments <> "" then call write_editbox_in_case_note("Worker notes", case_note_additional_comments, 6)
	If case_note_additional_comments <> "" then call write_new_line_in_case_note("**********")
	call write_new_line_in_case_note(case_note_worker_signature)
End If

EMReadScreen case_note_edit_test, 5, 20, 3
If case_note_edit_test = "Mode:" then
	MsgBox "You have completed this task. Please make any changes necessary and upon pressing OK you will be taken to this cases dails for clean-up.", 0, "Navigate to DAIL/DAIL"
	call navigate_to_screen("dail","dail")
End If

script_end_procedure("")