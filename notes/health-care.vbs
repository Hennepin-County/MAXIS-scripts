'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - Health Care.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 720          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
call changelog_update("03/23/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'FUNCTIONS BLOCK ===========================================================================================================

Function HCRE_panel_bypass()
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

Call run_another_script("C:\MAXIS-scripts\misc\class stat_detail.vbs")

'Dialog Editor saved format
' BeginDialog Dialog1, 0, 0, 555, 385, "Health Care Members Request"
'   GroupBox 10, 10, 465, 10, "Residents Requesting Health Care Coverage"
'   ButtonGroup ButtonPressed
'     PushButton 10, 25, 40, 15, "MEMB 01", Button3
'     PushButton 50, 25, 40, 15, "MEMB 01", Button5
'   GroupBox 10, 45, 465, 310, "MEMB XX - Member Name"
'   Text 20, 60, 180, 10, "Member: NAME GOES HERE"
'   Text 35, 70, 75, 10, "AGE: XXX"
'   Text 215, 60, 75, 10, "SSN: XXX-XX-XXXX"
'   Text 215, 70, 75, 10, "DOB: MM/DD/YY"
'   Text 310, 60, 110, 10, " Application Date: MM/DD/YY"
'   Text 315, 70, 95, 10, "Coverage Request: MM/YY"
'   Text 20, 90, 355, 10, "DISA - Start date: MM/DD/YY - End Date: MM/DD/YY   -    HC DISA Status: "
'   Text 40, 100, 325, 10, "Certification - Start date: MM/DD/YY - End Date: MM/DD/YY   -    Verif: "
'   Text 20, 115, 355, 10, "PREG - Due Date: MM/DD/YY   -   Verif:"
'   Text 45, 125, 325, 10, "Pregnancy End Date: MM/DD/YY   -   Verif:"
'   Text 20, 140, 380, 10, "PARE - Members lists as Child of Resident:"
'   Text 20, 155, 385, 10, "MEDI - Medicare Information - Source of detail: "
'   Text 40, 165, 145, 10, "Part A Premium - $ XXX"
'   Text 40, 175, 150, 10, "Part A Start: MM/DD/YY - End: MM/DD/YY"
'   Text 205, 165, 115, 10, " Part B Premium - $ XXX"
'   Text 205, 175, 215, 10, " Part B Premium - Start: MM/DD/YY - End: MM/DD/YY"
'   Text 20, 200, 105, 10, "Health Care Determination is at "
'   DropListBox 130, 195, 95, 45, "", List1
'   Text 485, 5, 75, 10, "---   DIALOGS   ---"
'   GroupBox 10, 220, 465, 60, "Medical Assistance"
'   Text 20, 240, 80, 10, "MA Basis of Eligibility:"
'   DropListBox 100, 235, 155, 45, "", List2
'   Text 35, 260, 65, 10, "Notes onMA Basis:"
'   EditBox 100, 255, 365, 15, Edit2
'   GroupBox 10, 285, 465, 60, "Medicare Savings Programs"
'   Text 20, 305, 80, 10, "MSP Basis of Eligibility:"
'   DropListBox 100, 300, 155, 45, "", List3
'   Text 30, 325, 70, 10, "Notes on MSP Basis:"
'   EditBox 100, 320, 365, 15, Edit3
'   ButtonGroup ButtonPressed
'     CancelButton 490, 360, 50, 15
'     OkButton 490, 340, 50, 15
' EndDialog


function define_main_dialog()

	BeginDialog Dialog1, 0, 0, 555, 385, "Information for Health Care Determination"

	  ButtonGroup ButtonPressed
	    If page_display = show_member_page Then
			GroupBox 10, 10, 465, 10, "Residents Requesting Health Care Coverage"
			x_pos = 10
			For the_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
				If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
                    If the_memb = selected_memb Then
    					Text x_pos+10, 30, 45, 10, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb)
    				Else
    					' PushButton 10, y_pos, 45, 10, "Person " & (the_memb + 1), HH_MEMB_ARRAY(button_one, the_memb)
						PushButton x_pos, 28, 40, 12, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb), HEALTH_CARE_MEMBERS(pers_btn_one_const, the_memb)
    				End If
    				x_pos = x_pos + 45
                End If
			Next
			' PushButton 50, 25, 40, 15, "MEMB 01", Button5
			GroupBox 10, 45, 465, 310, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) & " - " & HEALTH_CARE_MEMBERS(full_name_const, selected_memb) & " - PMI: " & HEALTH_CARE_MEMBERS(pmi_const, selected_memb)
			Text 20, 60, 180, 10, "Member: " & HEALTH_CARE_MEMBERS(full_name_const, selected_memb)
			Text 35, 70, 75, 10, "AGE: " & HEALTH_CARE_MEMBERS(age_const, selected_memb)
			Text 215, 60, 75, 10, "SSN: " & HEALTH_CARE_MEMBERS(ssn_const, selected_memb)
			Text 215, 70, 75, 10, "DOB: " & HEALTH_CARE_MEMBERS(dob_const, selected_memb)
			Text 320, 60, 110, 10, " Application Date: " & HEALTH_CARE_MEMBERS(hc_appl_date_const, selected_memb)
			Text 315, 70, 95, 10, "Coverage Request: " & HEALTH_CARE_MEMBERS(hc_cov_date_const, selected_memb)

			Text 20, 90, 200, 10, "Current MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) & " Health Care Status: " & HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, selected_memb)
			Text 275, 90, 105, 10, "Health Care Determination is at "
			DropListBox 380, 85, 75, 45, "Application"+chr(9)+"Recertification", HEALTH_CARE_MEMBERS(HC_determination_process_const, selected_memb)

			y_pos = 110
			If HEALTH_CARE_MEMBERS(DISA_exists_const, selected_memb) = True Then
				Text 20, y_pos, 200, 10, "DISA   -   Start date: " & HEALTH_CARE_MEMBERS(DISA_start_date_const, selected_memb) & "   -   End Date: " & HEALTH_CARE_MEMBERS(DISA_end_date_const, selected_memb)
				Text 250, y_pos, 200, 10, "HC DISA Status: " & HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, selected_memb)
				y_pos = y_pos + 10
				If HEALTH_CARE_MEMBERS(DISA_cert_start_const, selected_memb) <> "" Then Text 55, y_pos, 200, 10, "Certification   -   Start date: " & HEALTH_CARE_MEMBERS(DISA_cert_start_const, selected_memb) & " - End Date: " & HEALTH_CARE_MEMBERS(DISA_cert_end_const, selected_memb)
				If HEALTH_CARE_MEMBERS(DISA_cert_start_const, selected_memb) = "" Then Text 55, y_pos, 200, 10, "Certification   -   NO CERTIFICATION DATES Entered"
				Text 285, y_pos, 150, 10, "  Verif: " & HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, selected_memb)
				y_pos = y_pos + 10
				If HEALTH_CARE_MEMBERS(DISA_waiver_info_const, selected_memb) <> "" Then
					Text 55, y_pos, 200, 10, "LTC Waiver: " & HEALTH_CARE_MEMBERS(DISA_waiver_info_const, selected_memb)
					y_pos = y_pos + 10
				End If
				y_pos = y_pos + 5
				Text 55, y_pos, 45, 10, "DISA Notes:"
				EditBox 100, y_pos-5, 365, 15, HEALTH_CARE_MEMBERS(DISA_notes_const, selected_memb)
				y_pos = y_pos + 10
			Else
				Text 20, y_pos, 355, 10, "DISA   -   No DISA Panel Exists for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb)
				y_pos = y_pos + 10
			End If
			y_pos = y_pos + 5
			If HEALTH_CARE_MEMBERS(PREG_exists_const, selected_memb) = True Then
				Text 20, y_pos, 355, 10, "PREG   -   Due Date: "&  HEALTH_CARE_MEMBERS(PREG_due_date_const, selected_memb) & "   -   Verif:" &  HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, selected_memb)
				y_pos = y_pos + 10
				Text 55, y_pos, 325, 10, "Pregnancy End Date: " &  HEALTH_CARE_MEMBERS(PREG_end_date_const, selected_memb) & "   -   Verif:" &  HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, selected_memb)
				y_pos = y_pos + 15
				Text 55, y_pos, 45, 10, "PREG Notes:"
				EditBox 100, y_pos-5, 365, 15, HEALTH_CARE_MEMBERS(PREG_notes_const, selected_memb)
				y_pos = y_pos + 10
			Else
				Text 20, y_pos, 355, 10, "PREG   -   No PREG Panel Exists for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb)
				y_pos = y_pos + 10
			End If
			y_pos = y_pos + 5
			If HEALTH_CARE_MEMBERS(PARE_exists_const, selected_memb) = True Then
				Text 20, y_pos, 380, 10, "PARE   -   Members lists as Child of Resident:" & HEALTH_CARE_MEMBERS(PARE_list_of_children_const, selected_memb)
				y_pos = y_pos + 15
				Text 55, y_pos, 45, 10, "PARE Notes:"
				EditBox 100, y_pos-5, 365, 15, HEALTH_CARE_MEMBERS(PARE_notes_const, selected_memb)
			Else
				Text 20, y_pos, 380, 10, "PARE   -   No PARE Panel Exists for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb)
			End If
			y_pos = y_pos + 15
			If HEALTH_CARE_MEMBERS(MEDI_exists_const, selected_memb) = True Then
				Text 20, y_pos, 385, 10, "MEDI   -   Medicare Information - Source of detail: " & HEALTH_CARE_MEMBERS(MEDI_info_source_const, selected_memb)
				y_pos = y_pos + 10
				Text 55, y_pos, 145, 10, "Part A Premium - $ " & HEALTH_CARE_MEMBERS(MEDI_part_A_premium_const, selected_memb)
				Text 215, y_pos, 115, 10, " Part B Premium - $ " & HEALTH_CARE_MEMBERS(MEDI_part_B_premium_const, selected_memb)
				y_pos = y_pos + 10
				Text 55, y_pos, 150, 10, "Part A Start: " & HEALTH_CARE_MEMBERS(MEDI_part_A_start_const, selected_memb) & " - End: " & HEALTH_CARE_MEMBERS(MEDI_part_A_end_const, selected_memb)
				Text 215, y_pos, 215, 10, " Part B Premium - Start: " & HEALTH_CARE_MEMBERS(MEDI_part_B_start_const, selected_memb) & " - End: " & HEALTH_CARE_MEMBERS(MEDI_part_B_end_const, selected_memb)
				y_pos = y_pos + 15
				Text 55, y_pos, 45, 10, "MEDI Notes:"
				EditBox 100, y_pos-5, 365, 15, HEALTH_CARE_MEMBERS(MEDI_notes_const, selected_memb)
				y_pos = y_pos + 10
			Else
				Text 20, y_pos, 385, 10, "MEDI   -   No MEDI Panel Exists for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb)
				y_pos = y_pos + 10
			End If
			y_pos = y_pos + 5

			' GroupBox 10, 220, 465, 60, "Medical Assistance"
			Text 20, 275, 80, 10, "MA Basis of Eligibility:"
			DropListBox 100, 270, 110, 45, "Select One..."+chr(9)+"Disabled"+chr(9)+"Blind"+chr(9)+"Elderly"+chr(9)+"Parent"+chr(9)+"Caretaker"+chr(9)+"Adult without Children"+chr(9)+"Pregnant"+chr(9)+"Child (19-20)"+chr(9)+"Child (2-18)"+chr(9)+"Child (0-1)"+chr(9)+"Auto Newborn"+chr(9)+"Foster Care Child"+chr(9)+"Previous Foster Care Child", HEALTH_CARE_MEMBERS(MA_basis_of_elig_const, selected_memb)
			Text 35, 295, 65, 10, "Notes on MA Basis:"
			EditBox 100, 290, 365, 15, HEALTH_CARE_MEMBERS(MA_basis_notes_const, selected_memb)
			' GroupBox 10, 285, 465, 60, "Medicare Savings Programs"
			Text 20, 315, 80, 10, "MSP Basis of Eligibility:"
			DropListBox 100, 310, 110, 45, "Select One..."+chr(9)+"Disabled"+chr(9)+"Blind"+chr(9)+"Elderly", HEALTH_CARE_MEMBERS(MSP_basis_of_elig_const, selected_memb)
			Text 30, 335, 70, 10, "Notes on MSP Basis:"
			EditBox 100, 330, 365, 15, HEALTH_CARE_MEMBERS(MSP_basis_notes_const, selected_memb)

			Text 505, 17, 55, 13, "HC MEMBs"
		ElseIf page_display = show_jobs_page Then
			grp_len = 5
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True Then
					grp_len = grp_len + 100
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True Then
					grp_len = grp_len + 100
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True Then
					grp_len = grp_len + 100
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True Then
					grp_len = grp_len + 100
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True Then
					grp_len = grp_len + 100
				End If
			Next
			GroupBox 10, 10, 465, grp_len, "JOBS Income"
			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_one_inc_start_date(each_memb)
					Text 180, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_one_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_one_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_one_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, 60, 430, 30, "Pay Detail"
					y_pos = y_pos + 15
					Text 40, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_one_main_pay_freq(each_memb)
					Text 175, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_one_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 20
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, STAT_INFORMATION(month_ind).stat_jobs_one_notes(each_memb)
					y_pos = y_pos + 25
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_two_inc_start_date(each_memb)
					Text 180, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_two_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_two_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_two_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, 60, 430, 30, "Pay Detail"
					y_pos = y_pos + 15
					Text 40, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_two_main_pay_freq(each_memb)
					Text 175, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_two_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 20
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, STAT_INFORMATION(month_ind).stat_jobs_two_notes(each_memb)
					y_pos = y_pos + 25
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_three_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_three_inc_start_date(each_memb)
					Text 180, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_three_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_three_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_three_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, 60, 430, 30, "Pay Detail"
					y_pos = y_pos + 15
					Text 40, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_three_main_pay_freq(each_memb)
					Text 175, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_three_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 20
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, STAT_INFORMATION(month_ind).stat_jobs_three_notes(each_memb)
					y_pos = y_pos + 25
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_four_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_four_inc_start_date(each_memb)
					Text 180, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_four_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_four_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_four_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, 60, 430, 30, "Pay Detail"
					y_pos = y_pos + 15
					Text 40, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_four_main_pay_freq(each_memb)
					Text 175, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_four_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 20
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, STAT_INFORMATION(month_ind).stat_jobs_four_notes(each_memb)
					y_pos = y_pos + 25
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_five_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_five_inc_start_date(each_memb)
					Text 180, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_five_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_five_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_five_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, 60, 430, 30, "Pay Detail"
					y_pos = y_pos + 15
					Text 40, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_five_main_pay_freq(each_memb)
					Text 175, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_five_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 20
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, STAT_INFORMATION(month_ind).stat_jobs_five_notes(each_memb)
					y_pos = y_pos + 25
				End If
			Next
			' BeginDialog Dialog1, 0, 0, 555, 385, "Health Care Members Request"
			' GroupBox 10, 10, 465, 205, "JOBS Income"
			' Text 485, 5, 75, 10, "---   DIALOGS   ---"
			' ButtonGroup ButtonPressed
			' 	CancelButton 490, 360, 50, 15
			' 	OkButton 490, 340, 50, 15
			' Text 20, 25, 205, 10, "MEMB 01 - NAME "
			' Text 235, 25, 230, 10, "Employed at EMPLOYER NAME"
			' Text 30, 40, 115, 10, "Income Start Date: MM/DD/YY"
			' Text 180, 40, 115, 10, "Income End Date: MM/DD/YY"
			' Text 30, 50, 185, 10, "Verification: VERIF HERE"
			' Text 240, 50, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
			' GroupBox 30, 60, 430, 30, "Pay Detail"
			' Text 40, 75, 100, 10, "Pay Frequency: BIWEEKLY"
			' Text 175, 75, 175, 10, "Pay Amount: $ XXXX.XX per pay date"
			' Text 30, 100, 50, 10, "Income Notes:"
			' EditBox 80, 95, 380, 15, Edit1
			' Text 20, 120, 205, 10, "MEMB 01 - NAME "
			' Text 235, 120, 230, 10, "Employed at EMPLOYER NAME"
			' Text 30, 135, 115, 10, "Income Start Date: MM/DD/YY"
			' Text 180, 135, 115, 10, "Income End Date: MM/DD/YY"
			' Text 30, 145, 185, 10, "Verification: VERIF HERE"
			' Text 240, 145, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
			' GroupBox 30, 155, 430, 30, "Pay Detail"
			' Text 40, 170, 100, 10, "Pay Frequency: BIWEEKLY"
			' Text 175, 170, 175, 10, "Pay Amount: $ XXXX.XX per pay date"
			' Text 30, 195, 50, 10, "Income Notes:"
			' EditBox 80, 190, 380, 15, Edit2
			' EndDialog

			Text 500, 32, 55, 13, "JOBS Income"
		ElseIf page_display = show_busi_page Then
			grp_len = 5
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
				If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
				If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
			Next

			GroupBox 10, 10, 465, grp_len, "BUSI Income"

			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
				If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
				If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
			Next

			Text 500, 47, 55, 13, "BUSI Income"
		ElseIf page_display = show_unea_page Then
			grp_len = 5
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
				If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
				If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
				If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
				If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True Then
					grp_len = grp_len + 85
				End If
			Next

			GroupBox 10, 10, 465, grp_len, "UNEA Income"
			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True Then
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_one_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_one_inc_start_date(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_one_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_one_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_one_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_one_inc_end_date(each_memb)
					y_pos = y_pos + 20
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, STAT_INFORMATION(month_ind).stat_unea_one_notes(each_memb)
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_unea_one_verif_code(each_memb) = "N" Then
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If
					y_pos = y_pos + 10
				End If
				If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True Then
					grp_len = grp_len + 85
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_two_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_two_inc_start_date(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_two_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_two_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_two_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_two_inc_end_date(each_memb)
					y_pos = y_pos + 20
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, STAT_INFORMATION(month_ind).stat_unea_two_notes(each_memb)
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_unea_two_verif_code(each_memb) = "N" Then
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If
					y_pos = y_pos + 10
				End If
				If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True Then
					grp_len = grp_len + 85
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_three_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_three_inc_start_date(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_three_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_three_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_three_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_three_inc_end_date(each_memb)
					y_pos = y_pos + 20
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, STAT_INFORMATION(month_ind).stat_unea_three_notes(each_memb)
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_unea_three_verif_code(each_memb) = "N" Then
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If
					y_pos = y_pos + 10
				End If
				If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True Then
					grp_len = grp_len + 85
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_four_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_four_inc_start_date(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_four_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_four_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_four_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_four_inc_end_date(each_memb)
					y_pos = y_pos + 20
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, STAT_INFORMATION(month_ind).stat_unea_four_notes(each_memb)
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_unea_four_verif_code(each_memb) = "N" Then
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If
					y_pos = y_pos + 10
				End If
				If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True Then
					grp_len = grp_len + 85
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " -  " & STAT_INFORMATION(month_ind).stat_memb_full_name(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_five_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_five_inc_start_date(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_five_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_five_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_five_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_five_inc_end_date(each_memb)
					y_pos = y_pos + 20
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, STAT_INFORMATION(month_ind).stat_unea_five_notes(each_memb)
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_unea_five_verif_code(each_memb) = "N" Then
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If
					y_pos = y_pos + 10
				End If
			Next

			' Text 20, 25, 205, 10, "MEMB 01 - NAME "
			' Text 235, 25, 230, 10, "Income type: TYPE INFO"
			' Text 30, 40, 115, 10, "Income Start Date: MM/DD/YY"
			' Text 180, 40, 115, 10, "Income End Date: MM/DD/YY"
			' Text 30, 50, 185, 10, "Verification: VERIF HERE"
			' Text 240, 50, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
			' Text 30, 65, 175, 10, "Monthly Income: $ XXXX.XX "
			' Text 30, 85, 50, 10, "Income Notes:"
			' EditBox 80, 80, 380, 15, Edit1

			Text 500, 62, 55, 13, "UNEA Income"
		ElseIf page_display = show_asset_page Then
			GroupBox 10, 10, 465, 205, "Assets"



			Text 510, 77, 55, 13, "Assets"
		ElseIf page_display = show_expenses_page Then
			GroupBox 10, 10, 465, 205, "Expenses"



			Text 505, 92, 55, 13, "Expenses"
		ElseIf page_display = show_other_page Then
			GroupBox 10, 10, 465, 205, "Other Information"



			Text 510, 107, 55, 13, "Other"
		' ElseIf page_display =  Then

		End If



		Text 485, 5, 75, 10, "---   DIALOGS   ---"
		Text 485, 17, 10, 10, "1"
		Text 485, 32, 10, 10, "2"
		Text 485, 47, 10, 10, "3"
		Text 485, 62, 10, 10, "4"
		Text 485, 77, 10, 10, "5"
		Text 485, 92, 10, 10, "6"
		Text 485, 107, 10, 10, "7"
		' Text 485, 122, 10, 10, "8"
		' Text 485, 137, 10, 10, "9"
		' Text 485, 152, 10, 10, "10"
		' Text 485, 167, 10, 10, "11"
		If page_display <> show_member_page 		Then PushButton 495, 15, 55, 13, "HC MEMBs", hc_memb_btn
		If page_display <> show_jobs_page 			Then PushButton 495, 30, 55, 13, "JOBS Income", jobs_inc_btn
		If page_display <> show_busi_page 			Then PushButton 495, 45, 55, 13, "BUSI Income", busi_inc_btn
		If page_display <> show_unea_page 			Then PushButton 495, 60, 55, 13, "UNEA Income", unea_inc_btn
		If page_display <> show_asset_page 			Then PushButton 495, 75, 55, 13, "Assets", assets_btn
		If page_display <> show_expenses_page 		Then PushButton 495, 90, 55, 13, "Expenses", expenses_btn
		If page_display <> show_other_page 			Then PushButton 495, 105, 55, 13, "Other", other_btn
		' If page_display <> show_q_16_20 				Then PushButton 495, 120, 55, 13, "Q. 16 - 20", caf_q_16_20_btn
		' If page_display <> show_q_21_24 				Then PushButton 495, 135, 55, 13, "Q. 21 - 24", caf_q_21_24_btn
		' If page_display <> show_qual 					Then PushButton 495, 150, 55, 13, "CAF QUAL Q", caf_qual_q_btn
		' If page_display <> show_pg_last 				Then PushButton 495, 165, 55, 13, "CAF Last Page", caf_last_page_btn
		btn_pos = 180
		If discrepancies_exist = True Then
			Text 485, btn_pos + 2, 10, 10, "12"
			If page_display <> discrepancy_questions 	Then PushButton 495, btn_pos, 55, 13, "Clarifications", discrepancy_questions_btn
			btn_pos = btn_pos + 15
		End If
		If expedited_determination_needed = True Then
			Text 485, btn_pos + 2, 10, 10, "13"
			If page_display <> expedited_determination Then PushButton 495, btn_pos, 55, 13, "EXPEDITED", expedited_determination_btn
			btn_pos = btn_pos + 15
		End If
		' PushButton 10, 365, 130, 15, "Interview Ended - INCOMPLETE", incomplete_interview_btn
		' PushButton 140, 365, 130, 15, "View Verifications", verif_button
		PushButton 345, 365, 50, 15, "NEXT", next_btn
		PushButton 400, 365, 150, 15, "Complete Health Care Determination", completed_hc_determination_btn

	EndDialog

end function

function read_person_based_STAT_info()
	HEALTH_CARE_MEMBERS(show_hc_detail_const, hc_memb) = True

	Call navigate_to_MAXIS_screen("STAT", "DISA")
	Call write_value_and_transmit(HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb), 20, 76)
	EMReadScreen disa_version, 1, 2, 78
	' MsgBox "disa_version - " & disa_version
	If disa_version = "1" Then
		'TODO - add a waiver
		HEALTH_CARE_MEMBERS(DISA_exists_const, hc_memb) = True
		' MsGbox HEALTH_CARE_MEMBERS(DISA_exists_const, hc_memb)
		EMReadScreen HEALTH_CARE_MEMBERS(DISA_start_date_const, hc_memb), 10, 6, 47
		EMReadScreen HEALTH_CARE_MEMBERS(DISA_end_date_const, hc_memb), 10, 6, 69
		EMReadScreen HEALTH_CARE_MEMBERS(DISA_cert_start_const, hc_memb), 10, 7, 47
		EMReadScreen HEALTH_CARE_MEMBERS(DISA_cert_end_const, hc_memb), 10, 7, 69
		EMReadScreen HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb), 1, 14, 59

		If HEALTH_CARE_MEMBERS(DISA_start_date_const, hc_memb) = "__ __ ____" Then HEALTH_CARE_MEMBERS(DISA_start_date_const, hc_memb) = ""
		If HEALTH_CARE_MEMBERS(DISA_end_date_const, hc_memb) = "__ __ ____" Then HEALTH_CARE_MEMBERS(DISA_end_date_const, hc_memb) = ""
		If HEALTH_CARE_MEMBERS(DISA_cert_start_const, hc_memb) = "__ __ ____" Then HEALTH_CARE_MEMBERS(DISA_cert_start_const, hc_memb) = ""
		If HEALTH_CARE_MEMBERS(DISA_cert_end_const, hc_memb) = "__ __ ____" Then HEALTH_CARE_MEMBERS(DISA_cert_end_const, hc_memb) = ""
		HEALTH_CARE_MEMBERS(DISA_start_date_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(DISA_start_date_const, hc_memb), " ", "/")
		HEALTH_CARE_MEMBERS(DISA_end_date_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(DISA_end_date_const, hc_memb), " ", "/")
		HEALTH_CARE_MEMBERS(DISA_cert_start_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(DISA_cert_start_const, hc_memb), " ", "/")
		HEALTH_CARE_MEMBERS(DISA_cert_end_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(DISA_cert_end_const, hc_memb), " ", "/")

		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "_" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = ""
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "F" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "CADI Conversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "G" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "CADI Diversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "H" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "CAC Conversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "I" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "CAC Diversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "J" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "EW Conversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "K" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "EW Diversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "L" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "TBI NF Conversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "M" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "TBI NF Diversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "P" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "TBI NB Conversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "Q" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "TBI NB Diversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "R" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "DD Conversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "S" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "DD Diversion"
		If HEALTH_CARE_MEMBERS(DISA_waiver_code_const, hc_memb) = "Y" Then HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) = "CSG Conversion"

		EMReadScreen HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb), 2, 13, 59
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "__" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "NO Health Care Disability Status"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "01" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "RSDI Only Disability"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "02" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "RSDI Only Blindness"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "03" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "SSI Disability"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "04" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "SSI Blindness"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "06" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "SMRT or SSA Pending"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "08" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "Certified Blind"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "10" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "Certified Disabled"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "11" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "Special Category - Disabled Child"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "20" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "TEFRA - Disabled"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "21" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "TEFRA - Blind"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "22" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "MA-EPD"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "23" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "MA/Waiver"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "24" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "SSA/SMRT Appeal Pending"
		If HEALTH_CARE_MEMBERS(DISA_hc_status_code_const, hc_memb) = "26" Then HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, hc_memb) = "SSA/SMRT Disability Deny"

		EMReadScreen HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb), 1, 13, 69
		If HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb) = "_" Then HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, hc_memb) = "NO Health Care Status Verifications"
		If HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb) = "1" Then HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, hc_memb) = "DHS 161 / Doctor Statement"
		If HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb) = "2" Then HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, hc_memb) = "SMRT Certified"
		If HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb) = "3" Then HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, hc_memb) = "Certified for RSDI or SSI"
		If HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb) = "6" Then HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, hc_memb) = "Other Document"
		If HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb) = "7" Then HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, hc_memb) = "Case Manager Determination"
		If HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb) = "8" Then HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, hc_memb) = "LTC Consult Services"
		If HEALTH_CARE_MEMBERS(DISA_hc_verif_code_const, hc_memb) = "N" Then HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, hc_memb) = "No Verification Provided"
	End If

	Call navigate_to_MAXIS_screen("STAT", "PREG")
	Call write_value_and_transmit(HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb), 20, 76)
	EMReadScreen preg_version, 1, 2, 78
	If preg_version = "1" Then
		HEALTH_CARE_MEMBERS(PREG_exists_const, hc_memb) = True
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_due_date_const, hc_memb), 8, 10, 53
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, hc_memb), 1, 6, 75
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_end_date_const, hc_memb), 8, 12, 53
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, hc_memb), 1, 12, 75

		If HEALTH_CARE_MEMBERS(PREG_due_date_const, hc_memb) = "__ __ __" Then HEALTH_CARE_MEMBERS(PREG_due_date_const, hc_memb) = ""
		HEALTH_CARE_MEMBERS(PREG_due_date_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(PREG_due_date_const, hc_memb), " ", "/")
		If HEALTH_CARE_MEMBERS(PREG_end_date_const, hc_memb) = "__ __ __" Then HEALTH_CARE_MEMBERS(PREG_end_date_const, hc_memb) = ""
		HEALTH_CARE_MEMBERS(PREG_end_date_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(PREG_end_date_const, hc_memb), " ", "/")

		If HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, hc_memb) = "_" Then HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, hc_memb) = ""
		If HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, hc_memb) = "Y" Then HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, hc_memb) = "Yes"
		If HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, hc_memb) = "N" Then HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, hc_memb) = "No"
		If HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, hc_memb) = "_" Then HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, hc_memb) = ""
		If HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, hc_memb) = "Y" Then HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, hc_memb) = "Yes"
		If HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, hc_memb) = "N" Then HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, hc_memb) = "No"
	End If


	Call navigate_to_MAXIS_screen("STAT", "PARE")
	Call write_value_and_transmit(HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb), 20, 76)
	EMReadScreen pare_version, 1, 2, 78
	If pare_version = "1" Then
		HEALTH_CARE_MEMBERS(PARE_exists_const, hc_memb) = True
		pare_row = 8
		Do
			EMReadScreen pare_ref_number, 2, pare_row, 24
			EMReadScreen pare_rela_type, 1, pare_row, 53
			If pare_rela_type = "1" or pare_rela_type = "7" Then
				HEALTH_CARE_MEMBERS(PARE_list_of_children_const, hc_memb) = HEALTH_CARE_MEMBERS(PARE_list_of_children_const, hc_memb) & ", MEMB " & pare_ref_number
			End If

			pare_row = pare_row + 1
			If pare_row = 18 Then
				pare_row = 8
				PF20
				EMReadScreen read_for_last_page, 9, 24, 14
				If read_for_last_page = "LAST PAGE" Then Exit Do
			End If
		Loop until pare_rela_type = "_"
		If left(HEALTH_CARE_MEMBERS(PARE_list_of_children_const, hc_memb), 1) = "," Then HEALTH_CARE_MEMBERS(PARE_list_of_children_const, hc_memb) = right(HEALTH_CARE_MEMBERS(PARE_list_of_children_const, hc_memb), len(HEALTH_CARE_MEMBERS(PARE_list_of_children_const, hc_memb))-1)
		HEALTH_CARE_MEMBERS(PARE_list_of_children_const, hc_memb) = trim(HEALTH_CARE_MEMBERS(PARE_list_of_children_const, hc_memb))
	End If

	Call navigate_to_MAXIS_screen("STAT", "MEDI")
	Call write_value_and_transmit(HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb), 20, 76)
	EMReadScreen medi_version, 1, 2, 78
	If medi_version = "1" Then
		EMReadScreen HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb), 1, 5, 64
		If HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "_" Then HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = ""
		If HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "P" Then HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "Provided by Client"
		If HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "A" Then HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "Award Letter"
		If HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "C" Then HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "Medicare Card"
		If HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "M" Then HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "MMIS"
		If HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "O" Then HEALTH_CARE_MEMBERS(MEDI_info_source_const, hc_memb) = "Other"

		HEALTH_CARE_MEMBERS(MEDI_exists_const, hc_memb) = True
		EMReadScreen HEALTH_CARE_MEMBERS(MEDI_part_A_premium_const, hc_memb), 8, 7, 46
		EMReadScreen HEALTH_CARE_MEMBERS(MEDI_part_B_premium_const, hc_memb), 8, 7, 73
		HEALTH_CARE_MEMBERS(MEDI_part_A_premium_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(MEDI_part_A_premium_const, hc_memb), "_", "")
		HEALTH_CARE_MEMBERS(MEDI_part_A_premium_const, hc_memb) = trim(HEALTH_CARE_MEMBERS(MEDI_part_A_premium_const, hc_memb))
		HEALTH_CARE_MEMBERS(MEDI_part_B_premium_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(MEDI_part_B_premium_const, hc_memb), "_", "")
		HEALTH_CARE_MEMBERS(MEDI_part_B_premium_const, hc_memb) = trim(HEALTH_CARE_MEMBERS(MEDI_part_B_premium_const, hc_memb))

		medi_row = 15
		Do
			final_detail_found = False
			EMReadScreen HEALTH_CARE_MEMBERS(MEDI_part_A_start_const, hc_memb), 8, medi_row, 24
			EMReadScreen HEALTH_CARE_MEMBERS(MEDI_part_A_end_const, hc_memb), 8, medi_row, 35
			If medi_row = 17 Then
				medi_row = 14
				PF20
				EMReadScreen read_for_last_page, 9, 24, 14
				If read_for_last_page = "LAST PAGE" Then final_detail_found = True
			End If
			If final_detail_found = False Then
				EMReadScreen next_A_start, 8, medi_row+1, 24
				EMReadScreen next_A_end, 8, medi_row+1, 35
				If next_A_start = "__ __ __" and next_A_end = "__ __ __" Then final_detail_found = True
			End If
			medi_row = medi_row + 1
		Loop until final_detail_found = True
		If HEALTH_CARE_MEMBERS(MEDI_part_A_start_const, hc_memb) = "__ __ __" Then HEALTH_CARE_MEMBERS(MEDI_part_A_start_const, hc_memb) = ""
		HEALTH_CARE_MEMBERS(MEDI_part_A_start_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(MEDI_part_A_start_const, hc_memb), " ", "/")
		If HEALTH_CARE_MEMBERS(MEDI_part_A_end_const, hc_memb) = "__ __ __" Then HEALTH_CARE_MEMBERS(MEDI_part_A_end_const, hc_memb) = ""
		HEALTH_CARE_MEMBERS(MEDI_part_A_end_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(MEDI_part_A_end_const, hc_memb), " ", "/")

		Do
			PF19
			EMReadScreen read_for_first_page, 10, 24, 14
		Loop until read_for_first_page = "FIRST PAGE"

		medi_row = 15
		Do
			final_detail_found = False
			EMReadScreen HEALTH_CARE_MEMBERS(MEDI_part_B_start_const, hc_memb), 8, medi_row, 24
			EMReadScreen HEALTH_CARE_MEMBERS(MEDI_part_B_end_const, hc_memb), 8, medi_row, 35
			If medi_row = 17 Then
				medi_row = 14
				PF20
				EMReadScreen read_for_last_page, 9, 24, 14
				If read_for_last_page = "LAST PAGE" Then final_detail_found = True
			End If
			If final_detail_found = False Then
				EMReadScreen next_A_start, 8, medi_row+1, 24
				EMReadScreen next_A_end, 8, medi_row+1, 35
				If next_A_start = "__ __ __" and next_A_end = "__ __ __" Then final_detail_found = True
			End If
			medi_row = medi_row + 1
		Loop until final_detail_found = True
		If HEALTH_CARE_MEMBERS(MEDI_part_B_start_const, hc_memb) = "__ __ __" Then HEALTH_CARE_MEMBERS(MEDI_part_B_start_const, hc_memb) = ""
		HEALTH_CARE_MEMBERS(MEDI_part_B_start_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(MEDI_part_B_start_const, hc_memb), " ", "/")
		If HEALTH_CARE_MEMBERS(MEDI_part_B_end_const, hc_memb) = "__ __ __" Then HEALTH_CARE_MEMBERS(MEDI_part_B_end_const, hc_memb) = ""
		HEALTH_CARE_MEMBERS(MEDI_part_B_end_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(MEDI_part_B_end_const, hc_memb), " ", "/")
	End If



end function

function dialog_movement()
	For i = 0 to Ubound(HEALTH_CARE_MEMBERS, 2)
		' If HH_MEMB_ARRAY(i).imig_exists = TRUE Then case_has_imig = TRUE
		' MsgBox HH_MEMB_ARRAY(i).button_one
		If ButtonPressed = HEALTH_CARE_MEMBERS(pers_btn_one_const, i) Then
			If page_display = show_member_page Then selected_memb = i
		End If
	Next
	If ButtonPressed = hc_memb_btn Then page_display = show_member_page
	If ButtonPressed = jobs_inc_btn Then page_display = show_jobs_page
	If ButtonPressed = busi_inc_btn Then page_display = show_busi_page
	If ButtonPressed = unea_inc_btn Then page_display = show_unea_page
	If ButtonPressed = assets_btn Then page_display = show_asset_page
	If ButtonPressed = expenses_btn Then page_display = show_expenses_page
	If ButtonPressed = other_btn Then page_display = show_other_page
end function


'END FUNCTIONS BLOCK =======================================================================================================



'DECLARATIONS ==============================================================================================================
Const ref_numb_const 				= 0
Const full_name_const				= 1
Const first_name_const				= 2
Const last_name_const				= 3
Const last_name_first_full_const	= 4
Const age_const 					= 5
Const ssn_const 					= 6
Const dob_const 					= 7
Const pmi_const 					= 8
Const relationship_code_const 		= 9
Const id_verif_code_const 			= 10
Const alien_id_number_const 		= 11

Const marital_status_code_const		= 12
Const spouse_ref_number_const		= 13
Const spouse_array_position_const	= 14
Const citizen_yn_const				= 15
Const citizen_verif_code_const		= 16
Const ma_citizen_verif_code_const	= 17

Const hc_appl_date_const			= 18
Const hc_cov_date_const				= 19
Const hc_cov_mo_const				= 20
Const hc_cov_yr_const				= 21
Const hc_revw_month_const			= 22
Const hc_revw_mm_const				= 23
Const hc_revw_yy_const				= 24
Const hc_at_revw_const				= 25
Const hc_revw_process_const			= 26

Const case_pers_hc_status_code_const 	= 27
Const case_pers_hc_status_info_const 	= 28
Const member_is_applying_for_hc_const 	= 29
Const member_is_recert_for_hc_const 	= 30

Const show_hc_detail_const 				= 31
Const DISA_exists_const 				= 32
Const DISA_start_date_const 			= 33
Const DISA_end_date_const 				= 34
Const DISA_cert_start_const 			= 35
Const DISA_cert_end_const 				= 36
Const DISA_hc_status_code_const 		= 37
Const DISA_hc_status_info_const 		= 38
Const DISA_hc_verif_code_const 			= 39
Const DISA_hc_verif_info_const 			= 40
Const DISA_waiver_code_const			= 41
Const DISA_waiver_info_const			= 42
Const DISA_notes_const					= 43
Const PREG_exists_const 				= 44
Const PREG_due_date_const 				= 45
Const PREG_due_date_verif_const 		= 46
Const PREG_end_date_const 				= 47
Const PREG_end_date_verif_const 		= 48
Const PREG_notes_const					= 49
Const PARE_exists_const 				= 50
Const PARE_list_of_children_const 		= 51
Const PARE_notes_const					= 52
Const MEDI_exists_const 				= 53
Const MEDI_part_A_premium_const 		= 54
Const MEDI_part_B_premium_const 		= 55
Const MEDI_part_A_start_const 			= 56
Const MEDI_part_A_end_const 			= 57
Const MEDI_part_B_start_const 			= 58
Const MEDI_part_B_end_const 			= 59
Const MEDI_info_source_const 			= 60
Const MEDI_notes_const					= 61
Const HC_determination_process_const 	= 62
Const MA_basis_of_elig_const 			= 63
Const MA_basis_notes_const 				= 64
Const MSP_basis_of_elig_const 			= 65
Const MSP_basis_notes_const 			= 66
' Const _const =
' Const _const =
' Const _const =
' Const _const =
' Const _const =
Const pers_btn_one_const 				= 69
Const last_const						= 70

Dim HEALTH_CARE_MEMBERS()
ReDim HEALTH_CARE_MEMBERS(last_const, 0)



'defaulting some information
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
health_care_pending = False
health_care_active = False
hc_application_date = ""

form_selection_options = form_selection_options+chr(9)+"Health Care Programs Application for Certain Populations (DHS-3876)"
form_selection_options = form_selection_options+chr(9)+"MNsure Application for Health Coverage and Help paying Costs (DHS-6696)"
form_selection_options = form_selection_options+chr(9)+"Application for Payment of Long-Term Care Services (DHS-3531)"
form_selection_options = form_selection_options+chr(9)+"Breast and Cervical Cancer Coverage Group (DHS-3525)"
form_selection_options = form_selection_options+chr(9)+"Minnesota Family Planning Program Application (DHS-4740)"

page_display = ""
show_member_page 	= 0
show_jobs_page 		= 1
show_busi_page 		= 2
show_unea_page 		= 3
show_asset_page 	= 4
show_expenses_page 	= 5
show_other_page 	= 6

		' show_member_page 	= 0
		' show_jobs_page 		= 1
		' show_busi_page 		= 2
		' show_unea_page 		= 3
		' show_asset_page 	= 4
		' show_expenses_page 	= 5
		' show_other_page 	= 6

'BUTTON definitions
'START AT 1000 OR ABOVE. Person buttons start at 500
Const hc_memb_btn = 1010
Const jobs_inc_btn = 1011
Const busi_inc_btn = 1012
Const unea_inc_btn = 1013
Const assets_btn = 1014
Const expenses_btn = 1015
Const other_btn = 1016

'THE SCRIPT =====================================================================================================
EMConnect ""
Call check_for_MAXIS(False)
Call get_county_code
Call MAXIS_case_number_finder(MAXIS_case_number)

If MAXIS_case_number <> "" Then
	Call navigate_to_MAXIS_screen("REPT", "PND2")
	EMReadScreen pnd2_disp_limit, 13, 6, 35             'functionality to bypass the display limit warning if it appears.
	If pnd2_disp_limit = "Display Limit" Then transmit
	row = 1                                             'searching for the CASE NUMBER to read from the right row
	col = 1
	EMSearch MAXIS_case_number, row, col
	If row <> 24 and row <> 0 Then
		pnd2_row = row
		EMReadScreen hc_pend_code, 1, pnd2_row, 65
		'TODO - read the Additional APP line
		If hc_pend_code = "P" Then
			EMReadScreen hc_pend_date, 8, pnd2_row, 38
			form_date = replace(hc_pend_date, " ", "/")
		End If
	End If
	Call back_to_SELF
	Call navigate_to_MAXIS_screen("REPT", "REVW")
	EMReadScreen revw_form_received_date, 8, 13, 37
	'TODO - add more REVW handling here
	Call back_to_SELF
	EMWriteScreen MAXIS_case_number, 18, 43
End If

'Gather Case Number and the form processed
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 365, 280, "Health Care Determination"
  EditBox 80, 200, 50, 15, MAXIS_case_number
  DropListBox 80, 220, 275, 45, "Select One..."+form_selection_options, HC_form_name
  EditBox 80, 240, 50, 15, form_date
  EditBox 80, 260, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 260, 50, 15
    CancelButton 305, 260, 50, 15
  Text 105, 10, 120, 10, "Health Care Determination Script"
  Text 20, 40, 255, 20, "This script is to be run once MAXIS STAT panels have been updated with all accurate information from a Health Care Application Form."
  Text 20, 65, 255, 25, "If information displayed in this script is inaccurate, this means the information entered into STAT requires update. Cancel the script run and update STAT panels before running the script again."
  Text 20, 95, 255, 10, "The information and coding in STAT will directly pull into the script details:"
  Text 35, 105, 250, 10, "- Panels coded as needing verification will show up as verifications needed."
  Text 35, 115, 250, 10, "- Income amounts will be pulled from JOBS / UNEA / BUSI / ect panels"
  Text 40, 125, 150, 10, "and cannot be updated in the script dialogs."
  Text 35, 135, 250, 10, "- Asset amounts will be pulled from ACCT / CASH / SECU / ect panels and"
  Text 40, 145, 175, 10, "cannot be updated in the script dialogs."
  Text 35, 155, 250, 10, "- The details in STAT Panels should be accurate and the script serves as a"
  Text 40, 165, 245, 10, "secondary review of information that makes and eligibility determinations."
  Text 15, 180, 300, 10, "IF THE CASE INFORMATION IS WRONG IN THE SCRIPT, IT IS WRONG IN THE SYSTEM"
  Text 25, 205, 50, 10, "Case Number:"
  Text 15, 225, 60, 10, "Health Care Form:"
  Text 25, 245, 40, 10, "Form Date:"
  Text 15, 265, 60, 10, "Worker Signature:"
  GroupBox 10, 25, 345, 170, "Health Care Processing"
EndDialog

DO
	DO
	   	err_msg = ""
	   	Dialog Dialog1
	   	cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		If HC_form_name = "Select One..." Then err_msg = err_msg & vbCr & "* Select the form received that you are processing a Health Care determination from."
		If trim(worker_signature) = "" Then err_msg = err_msg & vbCr & "* Enter your name to sign your CASE/NOTE."
	   	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
Loop until are_we_passworded_out = false



'Read PROG and HCRE to gather appliation date and any retro request
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)
If is_this_priv = True Then Call script_end_procedure("This case appears PRIVILEGED. The script will now end as there is no access to case information.")
EMReadScreen case_county, 4, 21, 21
If case_county <> worker_county_code Then Call script_end_procedure("This case does not appear to be in your county and there is no access to case action. The script will now end.")
EMReadScreen prog_hc_appl_date, 8, 12, 33
EMReadScreen prog_hc_intvw_date, 8, 12, 55
EMReadScreen prog_hc_status, 4, 12, 74

If prog_hc_appl_date = "__ __ __" Then prog_hc_appl_date = ""
prog_hc_appl_date = replace(prog_hc_appl_date, " ", "/")
If prog_hc_intvw_date = "__ __ __" Then prog_hc_intvw_date = ""
prog_hc_intvw_date = replace(prog_hc_intvw_date, " ", "/")
hc_application_date = prog_hc_appl_date

If prog_hc_status = "PEND" Then health_care_pending = True
If prog_hc_status = "ACTV" Then health_care_active = True

Call navigate_to_MAXIS_screen("STAT", "HCRE")
hc_memb = 0
hc_row = 10
Do
	EMReadScreen hcre_ref_numb, 2, hc_row, 24
	If hcre_ref_numb <> "  " Then
		ReDim Preserve HEALTH_CARE_MEMBERS(last_const, hc_memb)

		HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) = hcre_ref_numb
		HEALTH_CARE_MEMBERS(pers_btn_one_const, hc_memb) = 500+hc_memb

		EMReadScreen HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb), 8, hc_row, 51
		EMReadScreen HEALTH_CARE_MEMBERS(hc_cov_mo_const, hc_memb), 2, hc_row, 64
		EMReadScreen HEALTH_CARE_MEMBERS(hc_cov_yr_const, hc_memb), 2, hc_row, 67
		EMReadScreen HEALTH_CARE_MEMBERS(hc_cov_date_const, hc_memb), 5, hc_row, 64

		HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb), " ", "/")
		HEALTH_CARE_MEMBERS(hc_cov_date_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(hc_cov_date_const, hc_memb), " ", "/")
		HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = False
		HEALTH_CARE_MEMBERS(member_is_recert_for_hc_const, hc_memb) = False
		' MsgBox "HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) - " & HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb)
		hc_memb = hc_memb + 1
	End If

	hc_row = hc_row + 1
	If hc_row = 18 Then
		hc_row = 10
		PF20
		EMReadScreen last_page, 9, 24, 14
		If last_page = "LAST PAGE" Then Exit Do

	End If
Loop until hcre_ref_numb = "  "

Call navigate_to_MAXIS_screen("STAT", "MEMB")
For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
	Call write_value_and_transmit(HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb), 20, 76)

	EMReadscreen HEALTH_CARE_MEMBERS(last_name_const, hc_memb), 25, 6, 30
	EMReadscreen HEALTH_CARE_MEMBERS(first_name_const, hc_memb), 12, 6, 63
	' EMReadscreen mid_initial, 1, 6, 79
	HEALTH_CARE_MEMBERS(last_name_const, hc_memb) = trim(replace(HEALTH_CARE_MEMBERS(last_name_const, hc_memb), "_", ""))
	HEALTH_CARE_MEMBERS(first_name_const, hc_memb) = trim(replace(HEALTH_CARE_MEMBERS(first_name_const, hc_memb), "_", ""))

	HEALTH_CARE_MEMBERS(full_name_const, hc_memb) = HEALTH_CARE_MEMBERS(first_name_const, hc_memb) & " " & HEALTH_CARE_MEMBERS(last_name_const, hc_memb)
	HEALTH_CARE_MEMBERS(last_name_first_full_const, hc_memb) = HEALTH_CARE_MEMBERS(last_name_const, hc_memb) & ", " & HEALTH_CARE_MEMBERS(first_name_const, hc_memb)

	' mid_initial = replace(mid_initial, "_", "")
    EMReadScreen HEALTH_CARE_MEMBERS(relationship_code_const, hc_memb), 2, 10, 42              'reading the relationship from MEMB'
	EMReadScreen HEALTH_CARE_MEMBERS(id_verif_code_const, hc_memb), 2, 9, 68
	EMReadScreen HEALTH_CARE_MEMBERS(ssn_const, hc_memb), 11, 7, 42
	EMReadScreen HEALTH_CARE_MEMBERS(dob_const, hc_memb), 10, 8, 42
	EMReadScreen HEALTH_CARE_MEMBERS(pmi_const, hc_memb), 8, 4, 46
	EMReadScreen HEALTH_CARE_MEMBERS(age_const, hc_memb), 3, 8, 76
	EMReadScreen HEALTH_CARE_MEMBERS(alien_id_number_const, hc_memb), 10, 15, 68

	If HEALTH_CARE_MEMBERS(ssn_const, hc_memb) = "___ __ ____" Then HEALTH_CARE_MEMBERS(ssn_const, hc_memb) = ""
	HEALTH_CARE_MEMBERS(ssn_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(ssn_const, hc_memb), " ", "-")

	If HEALTH_CARE_MEMBERS(dob_const, hc_memb) = "__ __ ____" Then HEALTH_CARE_MEMBERS(dob_const, hc_memb) = ""
	HEALTH_CARE_MEMBERS(dob_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(dob_const, hc_memb), " ", "/")

	HEALTH_CARE_MEMBERS(age_const, hc_memb) = trim(HEALTH_CARE_MEMBERS(age_const, hc_memb))
	If HEALTH_CARE_MEMBERS(age_const, hc_memb) = "" Then HEALTH_CARE_MEMBERS(age_const, hc_memb) = 0
	HEALTH_CARE_MEMBERS(age_const, hc_memb) = HEALTH_CARE_MEMBERS(age_const, hc_memb) * 1

	HEALTH_CARE_MEMBERS(pmi_const, hc_memb) = trim(HEALTH_CARE_MEMBERS(pmi_const, hc_memb))
	HEALTH_CARE_MEMBERS(alien_id_number_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(alien_id_number_const, hc_memb), "_", "")
Next

Call navigate_to_MAXIS_screen("STAT", "MEMB")
For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
	Call write_value_and_transmit(HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb), 20, 76)
	EMReadScreen HEALTH_CARE_MEMBERS(marital_status_code_const, hc_memb), 1, 7, 40
	EMReadScreen HEALTH_CARE_MEMBERS(spouse_ref_number_const, hc_memb), 2, 9, 49

	For other_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
		If HEALTH_CARE_MEMBERS(ref_numb_const, other_memb) = HEALTH_CARE_MEMBERS(spouse_ref_number_const, hc_memb) Then HEALTH_CARE_MEMBERS(spouse_array_position_const, hc_memb) = other_memb
	Next

	EMReadScreen HEALTH_CARE_MEMBERS(citizen_yn_const, hc_memb), 1, 11, 49
	EMReadScreen HEALTH_CARE_MEMBERS(citizen_verif_code_const, hc_memb), 2, 11, 78
	EMReadScreen HEALTH_CARE_MEMBERS(ma_citizen_verif_code_const, hc_memb), 1, 12, 49
Next

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
If unknown_hc_pending = True Then health_care_pending = True
If ma_status = "PENDING" Then health_care_pending = True
If msp_status = "PENDING" Then health_care_pending = True
If ma_status = "ACTIVE" Then health_care_active = True
If msp_status = "ACTIVE" Then health_care_active = True

'Read from CASE/PERS to find the people on the case and determine who is looking for HC and create an array.
'read from ELIG HC to see if any information exists.
Call navigate_to_MAXIS_screen("CASE", "PERS")
pers_row = 10
last_page_check = ""
Do
	EMReadScreen pers_memb_numb, 2, pers_row, 3
	EMReadScreen pers_hc_status, 1, pers_row, 61

	For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
		' MsgBox "pers_memb_numb - " & pers_memb_numb & vbCr & "HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) - " & HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb)
		If pers_memb_numb = HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) Then
			HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) = pers_hc_status
			' MsgBox "HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) - " & HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) & " - 1"
			If pers_hc_status = "I" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "INACTIVE"
			If pers_hc_status = "D" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "DENIED"
			If pers_hc_status = "A" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "ACTIVE"
			If pers_hc_status = "P" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "PENDING"
			If pers_hc_status = "R" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "REINSTATEMENT"
			' If pers_hc_status = "" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = ""
			If pers_hc_status = "P" Then health_care_pending = True
			If pers_hc_status = "A" Then health_care_active = True

			' MsgBox "hc_application_date - " & hc_application_date & vbCr & "HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) - " & HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb)
			If DateDiff("d", hc_application_date, HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb)) > 0 Then
				hc_application_date = HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb)
			End If
		End If
	Next

	If pers_memb_numb = "  " Then Exit Do
	' If pers_hc_status = "A" Then list_of_hc_membs = list_of_hc_membs & "~" & pers_memb_numb

	pers_row = pers_row + 3
	If pers_row = 19 Then
		pers_row = 10
		PF8
		EMReadScreen last_page_check, 9, 24, 14
	End If
Loop until last_page_check = "LAST PAGE"

For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
	HEALTH_CARE_MEMBERS(show_hc_detail_const, hc_memb) = False
	HEALTH_CARE_MEMBERS(DISA_exists_const, hc_memb) = False
	HEALTH_CARE_MEMBERS(PREG_exists_const, hc_memb) = False
	HEALTH_CARE_MEMBERS(PARE_exists_const, hc_memb) = False
	HEALTH_CARE_MEMBERS(MEDI_exists_const, hc_memb) = False
	' MsgBox "HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) - " & HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) & " - 2" & vbCr & "HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) - " & HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) & vbCr & "hc_application_date - " & hc_application_date
	If hc_application_date = HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) Then HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = True
	If HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) = "P" Then HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = True
	HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = True
	' MsgBox "HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) - " & HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb)
	If HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = True Then
		call read_person_based_STAT_info()
	End If
Next

'TODO - in the future add gathering information from REVW panel to detail if REVW is in process

'Handle for what the resident reports on the FORM for the HC

'DIALOG to provide details for each person of potential HC ELIGIBILITY BASIS to have worker indicate the correct BASIS

'IF application date is 4/1/23 or after, have a dialog to indicate the worker should review for potential coverage to indicate if the resident had HC during 03/2023 and is subject to Protected Coverage

'gather information
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

month_count = 0
ReDim preserve STAT_INFORMATION(month_count)

Set STAT_INFORMATION(month_count) = new stat_detail

STAT_INFORMATION(month_count).footer_month = MAXIS_footer_month
STAT_INFORMATION(month_count).footer_year = MAXIS_footer_year

Call STAT_INFORMATION(month_count).gather_stat_info







interview_questions_clear = False
page_display = show_member_page
selected_memb = 0
month_ind = 0
Do
	Do
		Do
			Do
				' MsgBox page_display
				' MsgBox update_arep & " - before define dlg"
				Dialog1 = Empty
				call define_main_dialog

				err_msg = ""

				prev_page = page_display
				previous_button_pressed = ButtonPressed
				' MsgBox update_arep & " - before display dlg"

				dialog Dialog1
				' save_your_work
				cancel_confirmation

			Loop until err_msg = ""

			call dialog_movement

		Loop until leave_loop = TRUE
		proceed_confirm = MsgBox("Have you completed the Interview?" & vbCr & vbCr &_
								 "Once you proceed from this point, there is no opportunity to change information that will be entered in CASE/NOTE or in the Interview Notes PDF." & vbCr & vbCr &_
								 "Following this point is only check eDRS and Forms Review." & vbCr & vbCr &_
								 "Press 'No' now if you have additional notes to make or information to review/enter. This will bring you back to the main dailogs." & vbCr &_
								 "Press 'Yes' to confinue to the final part of the interivew (forms)." & vbCr &_
								 "Press 'Cancel' to end the script run.", vbYesNoCancel+ vbQuestion, "Confirm Interview Completed")
		If proceed_confirm = vbCancel then cancel_confirmation

	Loop Until proceed_confirm = vbYes
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE
