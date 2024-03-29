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

'--------Function Creation--------'
'Function to read/write in UNEA panel (01 = RSDI, Disa, 02 = RSDI (not associated with disa), No Disa, 03 = SSI, 16 = railroad). If there is a disability on BDXP then its 01. 
'Questions TODO: Incorporate 4 scenerios- similar to MEDI code (not started)
		'1) No UNEA panel, end date read in TPQY 
		'2) No UNEA panel, no end date in TPQY 
		'3) Unea panel, end date read in TPQY 
		'4) Unea panel, no end date in TPQY
Function UNEA_income_panel(inc_type, income_amt, claim_number, start_date, end_date)
	'If end date exists, we want to end the income otherwise we want to start it. 
		
'Verifying the current panel number and panel type (SSI/RSDI)
If check_infc_panel = "INFC" Then Call navigate_to_MAXIS_screen("STAT", "UNEA")
EMWriteScreen member_number, 20, 76 		'Navigating to STAT/UNEA
EMWriteScreen "01", 20, 79 		'to ensure we're on the 1st instance of UNEA panels for the appropriate member
transmit

Do		'Do loop to read through all UNEA panels
	EMReadScreen current_panel_number, 1, 2, 73
	EMReadScreen total_amt_of_panels, 1, 2, 78
	EMReadScreen panel_type, 2, 5, 37
	If panel_type = inc_type then
		PF9	
		If inc_type <> 16 then 
			EMWriteScreen "7", 5, 65		'Write Verification Worker Initiated Verfication "7"
		Else 
			EMWriteScreen "6", 5, 65 'Casey ToDo- "verification code 7 only allowed for income type 01, 02, 03 or 44" what should I use for 16?
		End If
		'Clear and write claim number
		EMWriteScreen "_______________", 6, 37
		EMWriteScreen claim_number, 6, 37 'TODO: Testing values "123456789A00" (rsdi_claim_numb, ssi_claim_numb)
		Call create_mainframe_friendly_date(start_date, 7, 37, "YY") 	'income start date (SSI: ssi_SSP_elig_date, RSDI: intl_entl_date)
		Call create_mainframe_friendly_date(end_date, 7, 68, "YY")	'income end date (SSI: ssi_denial_date, RSDI: susp_term_date)
		
		Call write_value_and_transmit("X", 6, 56)
		EMWriteScreen "________", 9, 65
		EMWriteScreen income_amt, 9, 65		'TODO: Testing values (rsdi_gross_amt or ssi_gross_amt )
		EMWriteScreen "1", 10, 63		'code for pay frequency
		Do
			transmit
			EMReadScreen HC_popup, 9, 7, 41
			If HC_popup = "HC Income" then transmit
		Loop until HC_popup <> "HC Income"
		
		'Retrospective amounts clear and write
		row = 13
		DO
		    EMWriteScreen "__", row, 25
		    EMWriteScreen "__", row, 28
		    EMWriteScreen "__", row, 31
		    EMWriteScreen "________", row, 39
		    row = row + 1
		Loop until row = 18

		EMWriteScreen CM_minus_1_mo, 13, 25 'hardcoded dates
		EMWriteScreen "01", 13, 28
		EMWriteScreen CM_minus_1_yr, 13, 31 'hardcoded dates
		EMWriteScreen "9999.99", 13, 39		'TODO: Testing values
		
		'Prospective Amounts clear and write 
		row = 13
		DO
		    EMWriteScreen "__", row, 54
		    EMWriteScreen "__", row, 57
		    EMWriteScreen "__", row, 60
		    EMWriteScreen "________", row, 68
		    row = row + 1
		Loop until row = 18

		EMWriteScreen CM_plus_1_mo, 13, 54 'hardcoded dates
		EMWriteScreen "01", 13, 57
		EMWriteScreen CM_plus_1_yr, 13, 60 'hardcoded dates
		EMWriteScreen income_amt, 13, 68		'TODO: Testing values (income_amt which = rsdi_gross_amt or ssi_gross_amt )
		transmit
		EMReadScreen cola_warning, 29, 24, 2
		If cola_warning = "WARNING: ENTER COLA DISREGARD" then transmit
		EMReadScreen HC_income_warning, 25, 24, 2
		If HC_income_warning = "WARNING: UPDATE HC INCOME" then transmit
		'Ilse notes - might need to transmit twice to get past COLA disregard amt. If in months Jan - June for RSDI we add the difference between last years's amt and this year's amt. 
		Exit Do
	Else 
		transmit
	End If
Loop Until current_panel_number = total_amt_of_panels

If panel_type <> inc_type then
	Call write_value_and_transmit("NN", 20, 79)
	EMWriteScreen inc_type, 5, 37 		'TODO: Testing values
	If inc_type <> 16 then 
			EMWriteScreen "7", 5, 65		'Write Verification Worker Initiated Verfication "7"
		Else 
			EMWriteScreen "6", 5, 65 'Casey ToDo- "verification code 7 only allowed for income type 01, 02, 03 or 44" what should I use for 16?
	End If
	EMWriteScreen claim_number, 6, 37 		'TODO: Testing values "123123123A00" (rsdi_claim_numb, ssi_claim_numb)
	Call create_mainframe_friendly_date(start_date, 7, 37, "YY") 'income start date (SSI: ssi_SSP_elig_date, RSDI: intl_entl_date)
	Call create_mainframe_friendly_date(end_date, 7, 68, "YY")	'income end date (SSI: ssi_denial_date, RSDI: susp_term_date)
	
	'HC Income Estimate Popup: clear and write SSI monthly amount (SDXP)
	Call write_value_and_transmit("X", 6, 56)
	EMWriteScreen income_amt, 9, 65 		'TODO: Testing values (income_amt = rsdi_gross_amt or ssi_gross_amt )
	EMWriteScreen "1", 10, 63		'code for pay frequency
	Do
		transmit
		EMReadScreen HC_popup, 9, 7, 41
		If HC_popup = "HC Income" then transmit
	Loop until HC_popup <> "HC Income"

	'Retrospective amounts clear and write (SSI Retrospective = CM_minus_1_mo and CM_minus_1_yr)
	EMWriteScreen CM_minus_1_mo, 13, 25
	EMWriteScreen "01", 13, 28
	EMWriteScreen CM_minus_1_yr, 13, 31
	EMWriteScreen "9999.99", 13, 39 		'TODO: Testing values
	'Prospective Amounts clear and write (SSI Prospective = CM_plus_1_mo and CM_plus_1_yr, RSDI = MAXIS_footer_month and MAXIS_footer_year)
	EMWriteScreen CM_plus_1_mo, 13, 54 
	EMWriteScreen "01", 13, 57
	EMWriteScreen CM_plus_1_yr, 13, 60 
	EMWriteScreen income_amt, 13, 68 		'TODO: Testing values (rsdi_gross_amt or ssi_gross_amt )
	transmit 		'this takes us out of edit mode
	EMReadScreen cola_warning, 29, 24, 2
	If cola_warning = "WARNING: ENTER COLA DISREGARD" then transmit
	EMReadScreen HC_income_warning, 25, 24, 2
	If HC_income_warning = "WARNING: UPDATE HC INCOME" then transmit
	'Ilse notes - might need to transmit twice to get past COLA disregard amt. If in months Jan - June for RSDI we add the difference between last years's amt and this year's amt. 
End If
End Function

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr
CM_minus_1_mo = right("0" & DatePart("m", DateAdd("m", -1, date)), 2)
CM_minus_1_yr = right(DatePart("yyyy", DateAdd("m", -1, date)), 2)
Call check_for_MAXIS(False)

'-------------------------------------------------------------------------------------------------DIALOG
'QUESTIONS TODO:
	'Casey ToDo: Need to connect to the SQL table instead of having a dialog box for case & member number

Dialog1 = "" 		'Blanking out previous dialog detail

BeginDialog Dialog1, 0, 0, 171, 85, "Case number"
  Text 20, 10, 45, 10, "Case number:"
  EditBox 70, 5, 65, 15, MAXIS_case_number
  EditBox 70, 25, 30, 15, MAXIS_footer_month
  EditBox 105, 25, 30, 15, MAXIS_footer_year
  EditBox 70, 45, 95, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 60, 65, 50, 15
    CancelButton 115, 65, 50, 15
  Text 5, 30, 65, 10, "Footer month/year:"
  Text 10, 50, 60, 10, "Worker Signature"
EndDialog
'Showing the case number dialog
Do
	DO
    	err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
	    Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE***" & vbCr & err_msg & vbCr & vbCr & "Resolve the following items for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)		'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false


'--------Navigate to STAT/MEMB panel--------'
'reading PMI and member number from STAT/MEMB panel 
Call navigate_to_MAXIS_screen("STAT", "MEMB")
EMReadScreen member_number, 2, 4, 33 
EMReadScreen recip_pmi, 8, 4, 45
recip_pmi = trim(recip_pmi)			'Trim PMI number to remove spaces
recip_pmi = right("00000000" & recip_pmi,8)			'Add leading zeros to ensure PMI uniformity with PMI in SVES panel


'--------Navigate to INFC/SVES, then TPQY to read all panels. Checks to ensure it is reading the correct panels--------'	
'QUESTIONS TODO:
	'Casey ToDo: Need to handling if gross <> net amount for RSDI. 
	'Remove dummy values used for testing
	'Review date hanlding to make sure it will EMwrite screen correctly based on MEDI/UNEA panels.

	'Handle for in futuer iterations:
		'Validation: Do we want to check DISA panel and see if they are on disability? Other fields we want to validate?

Call navigate_to_MAXIS_screen("INFC", "SVES")
EMWriteScreen recip_pmi, 5, 68
EMWriteScreen "TPQY", 20, 70
Transmit

'TPQY panel read and dates formatted MM/DD/YY
EMReadScreen check_TPQY_panel, 4, 2, 53 		'Reads for TPQY panel
If check_TPQY_panel = "TPQY" Then
	EMReadScreen resident_name, 60, 4, 10
	EMReadScreen reident_ssn, 11, 5, 9
	EMReadScreen resident_dob, 10, 5, 61
	EMReadScreen claim_num, 12, 5, 35  
	EMReadScreen sves_response, 8, 7, 22 		'Return Date
	EMReadScreen ssn_verif, 1, 8, 13
	EMReadScreen rsdi_record, 1, 8, 39
	EMReadScreen ssi_record, 1, 8, 65
	Trim(resident_name)
	Trim(resident_dob)
	Trim (claim_num)
	Trim(sves_response)
	resident_dob = replace(resident_dob, " ", "/")
	sves_response = replace(sves_response," ", "/")
End If
transmit

'BDXP panel read and formatted to MM/YY
EMReadScreen check_BDXP_panel, 4, 2, 53 		'Reads fro BDXP panel
If check_BDXP_panel = "BDXP" Then
	EMReadScreen rsdi_claim_numb, 12, 5, 40
	EMReadScreen dual_entl_nbr, 10, 5, 69 
	EMReadScreen rsdi_status_code, 2, 6, 19
	EMReadScreen rsdi_staus_desc, 30, 6, 22
	EMReadScreen rsdi_paydate, 5, 8, 5 
	EMReadScreen rsdi_gross_amt, 8, 8, 16 
	EMReadScreen rsdi_net_amt, 8, 8, 32
	EMReadScreen railroad_ind, 1, 8, 69 
	EMReadScreen intl_entl_date, 5, 11, 69
	EMReadScreen susp_term_date, 5, 14, 69 
	EMReadScreen rsdi_disa_date, 10, 15, 69
	Trim(rsdi_claim_numb)
	Trim(dual_entl_nbr)
	Trim(rsdi_staus_desc)
	Trim(rsdi_paydate)
	Trim(rsdi_gross_amt)
	Trim(rsdi_net_amt)
	Trim(railroad_ind)
	Trim(intl_entl_date)
	Trim(susp_term_date)
	Trim (rsdi_disa_date)
	rsdi_paydate = replace(rsdi_paydate, " ", "/01/")
	intl_entl_date = replace(intl_entl_date, " ", "/01")
	susp_term_date = replace(susp_term_date, " ", "/01")
	intl_entl_date = "01/01/23" 'TODO: Testing values
	rsdi_disa_date = "08/01/23" 'TODO: Testing values
	railroad_ind = "Y" 'TODO: Testing values
	susp_term_date = "09/01/23" 'TODO: Testing values
	rsdi_claim_numb = "123456789A00" 'TODO: Testing values
	rsdi_gross_amt = "2222.22" 'TODO: Testing values
	rsdi_net_amt = "3333.33" 'TODO: Testing values
End If
transmit 

'BDXM panel read and formatted to MM/YY
EMReadScreen check_BDXM_panel, 4, 2, 53 		'Reads for BDXM panel
If check_BDXM_panel = "BDXM" Then
	EMReadScreen medi_claim_num, 13, 4, 29
	EMReadScreen part_a_premium, 7, 6, 64
	EMReadScreen part_a_start, 5, 7, 25
	EMReadScreen part_a_stop, 5, 7, 63
	EMReadScreen part_a_buyin_ind, 1, 8, 25
	EMReadScreen part_a_buyin_code, 3, 8, 63 
	EMReadScreen part_a_buyin_start_date, 5, 9, 25
	EMReadScreen part_a_buyin_stop_date, 5, 9, 63
	EMReadScreen part_b_premium, 7, 12, 64
	EMReadScreen part_b_start, 5, 13, 25
	EMReadScreen part_b_stop, 5, 13, 63
	EMReadScreen part_b_buyin_ind, 1, 14, 25
	EMReadScreen Part_b_buyin_code, 3, 14, 63 
	EMReadScreen part_b_buyin_start_date, 5, 15, 25
	EMReadScreen part_b_buyin_stop_date, 5, 15, 63
	Trim(medi_claim_num)
	Trim(part_a_premium)
	Trim(part_b_premium)
	Trim(part_a_buyin_code)
	Trim(Part_b_buyin_code)
	part_a_start = replace(part_a_start, " ", "/01/")
	part_a_stop = replace(part_a_stop, " ", "/01/")
	part_a_buyin_start_date = replace(part_a_buyin_start_date, " ", "/01/")
	part_a_buyin_stop_date = replace(part_a_buyin_stop_date, " ", "/01/")
	part_b_start = replace(part_b_start, " ", "/01/")
	part_b_stop = replace(part_b_stop, " ", "/01/")
	part_b_buyin_start_date = replace(part_b_buyin_start_date, " ", "/01/")
	part_b_buyin_stop_date = replace(part_b_buyin_stop_date, " ", "/01/")
	part_a_start = "02/01/23" 'TODO: Testing values
	part_a_stop = "03/01/23" 'TODO: Testing values
	part_b_start = "01/01/23" 'TODO: Testing values
	part_b_stop = "03/01/23" 'TODO: Testing values
End If
transmit 

'SDXE panel read and formatted to MM/YY
EMReadScreen check_SDXE_panel, 4, 2, 53 		'Reads for SDXE panel
If check_SDXE_panel = "SDXE" Then
	EMReadScreen ssi_claim_numb, 12, 5, 36 
	EMReadScreen ssi_recip_code, 2, 7, 21
	EMReadScreen ssi_recip_desc, 22, 7, 24
	EMReadScreen fed_living, 1, 6, 70
	EMReadScreen ssi_pay_code, 3, 8, 21
	EMReadScreen ssi_pay_desc, 30, 8, 25
	EMReadScreen cit_ind_code, 1, 7, 70
	EMReadScreen ssi_denial_code, 3, 10, 26
	EMReadScreen ssi_denial_desc, 40, 10, 30
	EMReadScreen ssi_denial_date, 8, 11, 26
	EMReadScreen ssi_disa_date, 8, 12, 26
	EMReadScreen ssi_SSP_elig_date, 8, 13, 26
	EMReadScreen ssi_appeals_code, 1, 11, 65
	EMReadScreen ssi_appeals_date, 8, 12, 65
	EMReadScreen ssi_appeals_dec_code, 2, 13, 65
	EMReadScreen ssi_appeals_dec_date, 8, 14, 65
	EMReadScreen ssi_disa_pay_code, 1, 15, 65
	Trim(ssi_claim_numb)
	Trim(ssi_recip_desc)
	Trim(ssi_pay_desc)
	Trim(ssi_denial_desc)
	Trim(ssi_denial_date)
	Trim(ssi_disa_date)
	Trim(ssi_SSP_elig_date)
	Trim(ssi_appeals_date)
	Trim(ssi_appeals_dec_date)
	ssi_denial_date = replace(ssi_denial_date, " ", "/")
	ssi_disa_date = replace(ssi_disa_date, " ", "/")
	ssi_SSP_elig_date = replace(ssi_SSP_elig_date, " ", "/")
	ssi_appeals_date = replace(ssi_appeals_date, " ", "/")
	ssi_appeals_dec_date = replace(ssi_appeals_dec_date, " ", "/")
	ssi_SSP_elig_date = "01/01/23" 'TODO: Testing values
	ssi_denial_date = "07/01/23" 'TODO: Testing values
	ssi_claim_numb = "495411919DI" 'TODO: Testing values
End If
transmit 

'Navigation to SDXP and read 
EMReadScreen check_SDXP_panel, 4, 2, 50 		'Reads for SDXP panel
If check_SDXP_panel = "SDXP" Then 
	EMReadScreen ssi_pay_date, 5, 4, 16
	EMReadScreen ssi_gross_amt, 7, 4, 42
	EMReadScreen ssi_over_under_code, 1, 4, 73
	EMReadScreen ssi_pay_hist_1_date, 5, 8, 3
	EMReadScreen ssi_pay_hist_1_amt, 6, 8, 13
	EMReadScreen ssi_pay_hist_1_type, 1, 8, 25
	EMReadScreen ssi_pay_hist_2_date, 5, 9, 3
	EMReadScreen ssi_pay_hist_2_amt, 6, 9, 13
	EMReadScreen ssi_pay_hist_2_type, 1, 9, 25
	EMReadScreen ssi_pay_hist_3_date, 5, 10, 3
	EMReadScreen ssi_pay_hist_3_amt, 6, 10, 13
	EMReadScreen ssi_pay_hist_3_type, 1, 10, 25
	EMReadScreen gross_EI, 8, 5, 66
	EMReadScreen net_EI, 8, 6, 66
	EMReadScreen rsdi_income_amt, 8, 7, 66
	EMReadScreen pass_exclusion, 8, 8, 66
	EMReadScreen inc_inkind_start, 8, 9, 66
	EMReadScreen inc_inkind_stop, 8, 10, 66
	EMReadScreen rep_payee, 1, 11, 66
	Trim(ssi_pay_date)
	Trim(ssi_gross_amt)
	Trim(ssi_pay_hist_1_date)
	Trim(ssi_pay_hist_1_amt)
	Trim(ssi_pay_hist_2_date)
	Trim(ssi_pay_hist_2_amt)
	Trim(ssi_pay_hist_3_date)
	Trim(ssi_pay_hist_3_amt)
	Trim(gross_EI)
	Trim(net_EI)
	Trim(rsdi_income_amt)
	Trim(pass_exclusion)
	Trim(inc_inkind_start)
	Trim(inc_inkind_stop)
	ssi_pay_date = replace(ssi_pay_date, " ", "/01/")
	ssi_pay_hist_1_date = replace(ssi_pay_hist_1_date, " ", "/")
	ssi_pay_hist_2_date = replace(ssi_pay_hist_2_date, " ", "/")
	ssi_pay_hist_3_date = replace(ssi_pay_hist_3_date, " ", "/")
	inc_inkind_start = replace(inc_inkind_start, " ", "/")
	inc_inkind_stop = replace(inc_inkind_stop, " ", "/")
	ssi_gross_amt = "1111.11" 'TODO: Testing values
End If
transmit
EMReadScreen check_infc_panel, 4, 2, 45 		'Reads for INFC panel


'--------Navigation to STAT/UNEA and entering TQPY information --------'
'QUESTIONS: TODO
	'Casey ToDo: Need to enter TPQY info into SQL table
	'Revise handling of function for each of the 4 income types: 01, 02, 03, 16. 

	'Handle for in futuer iterations:
		'Write either issuance amount or Inc End Date in UNEA based on Payment Status in SDXE. 


'Handling for calling function for each type of income (01, 02, 03, 16).
If ssi_SSP_elig_date <> "" then memb_has_ssi = True
If intl_entl_date <> "" then memb_has_rsdi = True
If rsdi_disa_date <> "" then rsdi_has_disa = True
If railroad_ind <> "" then memb_has_railroad = True

If memb_has_ssi = true then
	Call UNEA_income_panel("03", ssi_gross_amt, ssi_claim_numb, ssi_SSP_elig_date, ssi_denial_date)
End If

If memb_has_rsdi = true Then
	If rsdi_has_disa = true then 
		Call UNEA_income_panel("01", rsdi_gross_amt, rsdi_claim_numb, intl_entl_date, susp_term_date)
	Else
		Call UNEA_income_panel("02", rsdi_gross_amt, rsdi_claim_numb, intl_entl_date, susp_term_date)
	End If
End If

If memb_has_railroad = true then 
	Call UNEA_income_panel("16", rsdi_net_amt, rsdi_claim_numb, intl_entl_date, susp_term_date)
End If 

MsgBox "Next the script writes information in MEDI. If panel does not exist, script will create a new panel."
'--------MEDI--------'
'Navigating to STAT/MEDI
'QUESTIONS TODO:
	'Incorporate 4 scenarios into MEDI coding. !!!!!!I was in the middle of doing this so this likely feel very messy, sorry!!!!
		'1) No UNEA panel, end date read in TPQY 
		'2) No UNEA panel, no end date in TPQY 
		'3) Unea panel, end date read in TPQY 
		'4) Unea panel, no end date in TPQY
	'Casey ToDo: Add apply premiums to spenddown based on BILS and HC budget. 
	
	'Handle for in futuer iterations:
		'Insert utilities-insert mbi from mmis to handle/retrieve MBI number for MEDI panel. This will also handle for source. 
		'Buy in Begin/End Date: revisit at a later date. This depends whether someone was on a medicare savings program or not. 
			'Apply premiums to spdn/budgets and Apply premiums to spdn/budgets thru based on...

Call navigate_to_MAXIS_screen("STAT", "MEDI")
EMWriteScreen member_number, 20, 79
EMWriteScreen "01", 20, 79			'to ensure we're on the 1st instance of UNEA panels for the appropriate member
transmit	

'Verifying if there is currently a panel or not, if not create one
EMReadScreen total_amt_of_panels, 1, 2, 78			'Checks to make sure there are JOBS panels for this member. If none exists, one will be created
If total_amt_of_panels = "0" then 
	If part_a_stop <> "" Then 
		If part_b_stop <> "" Then MsgBox "Add case note functionality to specify MEDI panel is not being created because both A & B have end/stop dates." 'handles for a/b stop both with end dates
	Else
		CALL write_value_and_transmit("NN", 20, 79) 		'Create new panel and write MEDI info
		EMWriteScreen medi_claim_num, 6, 39 'writes medi claim number for Medicare HICN
		EMWriteScreen "1cc4", 5, 38 'TODO: Testing values MBI number
		EMWRiteScreen "cc1", 5, 43 'TODO: Testing values MBI number
		EMWriteScreen "cc34", 5, 47 'TODO: Testing values MBI number
		EMWriteScreen "O", 5, 64	'TODO: Testing values MBI number
		EMWriteScreen part_a_premium, 7, 46 
		EMWriteScreen part_b_premium, 7, 73 
		EMWriteScreen "N", 9, 71 		' "N" for Qualified Working Disabled Individual
		EMWriteScreen "N", 10, 71 		' "N" for End Stage Renal Disease
		Call create_mainframe_friendly_date(part_a_start, 15, 24, "YY") 
		Call create_mainframe_friendly_date(part_a_stop, 15, 35, "YY")
		Call create_mainframe_friendly_date(part_b_start, 15, 54, "YY")
		Call create_mainframe_friendly_date(part_b_stop, 15, 65, "YY")
	End If
Else
	'MEDI premiums: clear and write
	PF9
	EMWriteScreen "O", 5, 64
	EMWriteScreen "________", 7, 46		'clear out part a premium field
	EMWriteScreen part_a_premium, 7, 46
	EMWriteScreen "________", 7, 73		'clear out part b premium field
	EMWriteScreen part_b_premium, 7, 73
	EMWriteScreen "N", 9, 71 		' "N" for Qualified Working Disabled Individual
	EMWriteScreen "N", 10, 71 		' "N" for End Stage Renal Disease

	'Clear and write begin/end dates for part a and b
	row = 17
	Do
		EMReadScreen MEDI_part_a_start, 8, row, 24 		'reads part a start date
		EMReadScreen medi_page_check, 34, 24, 2		'reads to see if we are in the last page
		If MEDI_part_a_start = "__ __ __" Then 
			MEDI_part_a_start = "" 		'blank out if not a date
		Else
			MEDI_part_a_start = replace(MEDI_part_a_start, " ", "/")	'reformatting with / for date
		End If

		EMReadScreen MEDI_part_a_end, 8, row, 35	'reads part a end date  
		If MEDI_part_a_end = "__ __ __" Then
			MEDI_part_a_end = ""					'blank out if not a date
		Else
			MEDI_part_a_end =replace(MEDI_part_a_end , " ", "/")		'reformatting with / for date
		End If
	
		If MEDI_part_a_end = "" Then
			If MEDI_part_a_start = "" Then
				row = row - 1 		'no dates found in this row, move up a row and reevaluate
			Else
				' If MEDI_part_a_start <> "" then	'only start date found, this is an open ended part b 
				If part_a_stop = "" then 
					exit do 'Call create_mainframe_friendly_date(part_a_start, row + 1, 24, "YY") ' ToDo: Should this be exit do instead?
				ElseIf part_a_stop <> "" then 
					Call create_mainframe_friendly_date(part_a_stop, row, 35, "YY")
				End If
				' End If
			End If
		Else
			If MEDI_part_a_start <> "" then 
				If medi_page_check = "COMPLETE THE PAGE BEFORE SCROLLING" then 
					Call create_mainframe_friendly_date(part_a_start, row + 1, 24, "YY")
					Call create_mainframe_friendly_date(part_a_stop, row + 1, 35, "YY")
				Else 
					PF20		'if stop/start are populated it will take you to the next page of dates
				End If
			End If
		End If
	Loop Until row = 14

	'Loop until back to first page before proceeding to part b do/loop
	EMReadScreen medi_page_check, 22, 24, 2
	Do 
		PF19
	Loop Until medi_page_check = "THIS IS THE FIRST PAGE"


	row = 17
	Do
		EMReadScreen MEDI_part_b_start, 8, row, 54		'reads part b start date
		EMReadScreen medi_page_check, 34, 24, 2		'reads to see if we are in the last page
		If MEDI_part_b_start = "__ __ __" Then
			MEDI_part_b_start = ""			'blank out if not a date
		Else
			MEDI_part_b_start = replace(MEDI_part_b_start, " ", "/")		'reformatting with / for date
		End If

		EMReadScreen MEDI_part_b_end, 8, row, 65	'reads part b end date
		If MEDI_part_b_end = "__ __ __" Then
			MEDI_part_b_end = ""					'blank out if not a date
		Else
			MEDI_part_b_end =replace(MEDI_part_b_end , " ", "/")		'reformatting with / for date
		End If

		If MEDI_part_b_end = "" Then
			If MEDI_part_b_start = "" Then
				row = row - 1			' no dates found in this row, move up a row and reevaluate
			Else
				' If MEDI_part_b_start <> "" then		'onely start date found, this is an open ended part b 
				If part_b_stop = "" then
					exit do 'Call create_mainframe_friendly_date(part_b_start, row + 1, 54, "YY") ' ToDo: Should this be exit do instead?
				ElseIf part_b_stop <> "" then 
					Call create_mainframe_friendly_date(part_b_stop, row, 65, "YY")
				End If
				' End If
			End If
		Else
			If MEDI_part_b_start <> "" then PF20'if stop/start are populated it will take you to the next page of dates
				If medi_page_check = "COMPLETE THE PAGE BEFORE SCROLLING" then 
					If part_b_stop = "" then
						Call create_mainframe_friendly_date(part_b_start, row + 1, 54, "YY")
					ElseIf part_b_stop <> "" then 
						Call create_mainframe_friendly_date(part_b_start, row + 1, 54, "YY")
						Call create_mainframe_friendly_date(part_b_stop, row + 1, 65, "YY")
					End If
				Else
				End If 
			End If
		End If
	Loop Until row = 14
End If


MsgBox "Last, the script will case note the applicable information."
'TODO: Revise based on feedback, necessary info for case note. 
renewal_period = MAXIS_footer_month & "/" & MAXIS_footer_year		'establishing the renewal period for the header of the case note

start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("---Income Verification Case Note---")
Call write_variable_in_CASE_NOTE("Updated UNEA panel and MEDI panel through a data match for " & member_number)

Call write_variable_in_CASE_NOTE("*RSDI Information*")
Call write_variable_with_indent_in_CASE_NOTE("RSDI Pay Date: " & rsdi_paydate)
Call write_variable_with_indent_in_CASE_NOTE("RSDI Gross Amount: " & rsdi_gross_amt)
Call write_variable_with_indent_in_CASE_NOTE("RSDI Net Amount: " & rsdi_net_amt)

Call write_variable_in_case_note("*Medicare Information*")
Call write_variable_with_indent_in_CASE_NOTE("Medicare claim number: " & medi_claim_num)
Call write_variable_with_indent_in_CASE_NOTE("Part A Premium: " & part_a_premium)
Call write_variable_with_indent_in_CASE_NOTE("Part A Start Date: " & part_a_start)
Call write_variable_with_indent_in_CASE_NOTE("Part A Stop Date: " & part_a_stop)
Call write_variable_with_indent_in_CASE_NOTE("Part B Premium: " & part_b_premium)
Call write_variable_with_indent_in_CASE_NOTE("Part B Start Date: " & part_b_start)
Call write_variable_with_indent_in_CASE_NOTE("Part B Stop Date: " & part_b_stop)

Call write_variable_in_case_note("*SSI Information*")
Call write_variable_with_indent_in_CASE_NOTE("SSI claim number: " & ssi_claim_num)
Call write_variable_with_indent_in_CASE_NOTE("Payment Status: " & ssi_pay_code & ssi_pay_desc)
Call write_variable_with_indent_in_CASE_NOTE("Pay Date: " & ssi_pay_date)
Call write_variable_with_indent_in_CASE_NOTE("SSI Gross Amount: " & ssi_gross_amt)
Call write_variable_with_indent_in_CASE_NOTE("Gross Earned income: " & gross_EI)
Call write_variable_with_indent_in_CASE_NOTE("Net Earned income: " & net_EI)
Call write_variable_with_indent_in_CASE_NOTE("Gross RSDI Income Amount: " & rsdi_income_amt)

call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("End of BULK TPQY Script")