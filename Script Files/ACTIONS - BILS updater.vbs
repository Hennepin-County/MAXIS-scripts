'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - BILS updater"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS
BeginDialog BILS_case_number_dialog, 0, 0, 161, 57, "BILS case number dialog"
  EditBox 95, 0, 60, 15, case_number
  CheckBox 15, 20, 130, 10, "Check here to update existing BILS.", updating_existing_BILS_check
  ButtonGroup ButtonPressed
    OkButton 25, 35, 50, 15
    CancelButton 85, 35, 50, 15
  Text 5, 5, 85, 10, "Enter your case number:"
EndDialog

BeginDialog BILS_updater_abbreviated_dialog, 0, 0, 161, 182, "BILS updater (abbreviated)"
  EditBox 110, 5, 40, 15, budget_begin
  EditBox 45, 25, 20, 15, ref_nbr_abbreviated
  EditBox 105, 55, 40, 15, gross_recurring_24
  EditBox 105, 75, 40, 15, gross_recurring_25
  EditBox 105, 95, 40, 15, gross_recurring_26
  EditBox 105, 115, 40, 15, gross_recurring_27
  EditBox 105, 135, 40, 15, gross_recurring_99
  ButtonGroup ButtonPressed
    OkButton 25, 160, 50, 15
    CancelButton 85, 160, 50, 15
  Text 10, 5, 95, 15, "Update begin date (MM/YY):"
  Text 10, 25, 35, 15, "MEMB #:"
  Text 15, 60, 90, 10, "Waivered Services (24):"
  Text 15, 80, 90, 10, "Medicare Prem (25):"
  Text 15, 100, 90, 10, "Dental/Health Prem (26):"
  Text 15, 120, 90, 10, "Remedial Care (27):"
  Text 15, 140, 90, 10, "Other/Fake BILS (99):"
  GroupBox 10, 45, 140, 110, "New gross for service types:"
EndDialog

BeginDialog BILS_updater_dialog, 0, 0, 416, 271, "BILS updater"
  Text 5, 10, 80, 10, "Budget period (MM/YY):"
  EditBox 85, 5, 45, 15, budget_begin
  Text 135, 10, 10, 10, "to:"
  EditBox 150, 5, 45, 15, budget_end
  GroupBox 5, 25, 405, 85, "Actual bills"
  Text 20, 35, 20, 10, "Ref#"
  Text 55, 35, 60, 10, "Date (MM/DD/YY)"
  Text 125, 35, 60, 10, "Service type"
  Text 245, 35, 25, 10, "Gross"
  Text 315, 35, 15, 10, "Ver"
  Text 375, 35, 30, 10, "Exp Type"
  EditBox 20, 50, 20, 15, ref_nbr_actual_01
  EditBox 55, 50, 50, 15, date_actual_01
  DropListBox 120, 50, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_actual_01
  EditBox 235, 50, 40, 15, gross_actual_01
  DropListBox 285, 50, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_actual_01
  DropListBox 380, 50, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_actual_01
  EditBox 20, 70, 20, 15, ref_nbr_actual_02
  EditBox 55, 70, 50, 15, date_actual_02
  DropListBox 120, 70, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_actual_02
  EditBox 235, 70, 40, 15, gross_actual_02
  DropListBox 285, 70, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_actual_02
  DropListBox 380, 70, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_actual_02
  EditBox 20, 90, 20, 15, ref_nbr_actual_03
  EditBox 55, 90, 50, 15, date_actual_03
  DropListBox 120, 90, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_actual_03
  EditBox 235, 90, 40, 15, gross_actual_03
  DropListBox 285, 90, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_actual_03
  DropListBox 380, 90, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_actual_03
  GroupBox 5, 120, 340, 145, "Recurring monthly bills"
  Text 20, 130, 20, 10, "Ref#"
  Text 55, 130, 60, 10, "Service type"
  Text 175, 130, 25, 10, "Gross"
  Text 245, 130, 15, 10, "Ver"
  Text 305, 130, 35, 10, "Exp Type"
  EditBox 20, 145, 20, 15, ref_nbr_recurring_01
  DropListBox 55, 145, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_recurring_01
  EditBox 165, 145, 40, 15, gross_recurring_01
  DropListBox 215, 145, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_recurring_01
  DropListBox 310, 145, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_recurring_01
  EditBox 20, 165, 20, 15, ref_nbr_recurring_02
  DropListBox 55, 165, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_recurring_02
  EditBox 165, 165, 40, 15, gross_recurring_02
  DropListBox 215, 165, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_recurring_02
  DropListBox 310, 165, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_recurring_02
  EditBox 20, 185, 20, 15, ref_nbr_recurring_03
  DropListBox 55, 185, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_recurring_03
  EditBox 165, 185, 40, 15, gross_recurring_03
  DropListBox 215, 185, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_recurring_03
  DropListBox 310, 185, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_recurring_03
  EditBox 20, 205, 20, 15, ref_nbr_recurring_04
  DropListBox 55, 205, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_recurring_04
  EditBox 165, 205, 40, 15, gross_recurring_04
  DropListBox 215, 205, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_recurring_04
  DropListBox 310, 205, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_recurring_04
  EditBox 20, 225, 20, 15, ref_nbr_recurring_05
  DropListBox 55, 225, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_recurring_05
  EditBox 165, 225, 40, 15, gross_recurring_05
  DropListBox 215, 225, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_recurring_05
  DropListBox 310, 225, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_recurring_05
  EditBox 20, 245, 20, 15, ref_nbr_recurring_06
  DropListBox 55, 245, 105, 15, " "+chr(9)+"01 Health Professional"+chr(9)+"03 Surgery"+chr(9)+"04 Chiropractic"+chr(9)+"05 Maternity/Reproductive"+chr(9)+"07 Hearing"+chr(9)+"08 Vision"+chr(9)+"09 Hospital"+chr(9)+"11 Hospice"+chr(9)+"13 SNF"+chr(9)+"14 Dental"+chr(9)+"15 Rx Drug/Non-Durable Supply"+chr(9)+"16 Home Health"+chr(9)+"17 Diagnostic"+chr(9)+"18 Mental Health"+chr(9)+"19 Rehab"+chr(9)+"21 Durable Med Equip"+chr(9)+"22 Medical Trans"+chr(9)+"24 Waivered Serv"+chr(9)+"25 Medicare Prem"+chr(9)+"26 Dental or Health Prem"+chr(9)+"27 Remedial Care"+chr(9)+"28 Non-FFP MCRE Service"+chr(9)+"30 Alternative Care"+chr(9)+"31 MCSHN"+chr(9)+"32 Ins Ext Prog"+chr(9)+"34 CW-TCM"+chr(9)+"37 Pay-In Spdn"+chr(9)+"42 Access Services"+chr(9)+"43 Chemical Dep"+chr(9)+"44 Nutritional Services"+chr(9)+"45 Organ/Tissue Transplant"+chr(9)+"46 Out-Of-Area Services"+chr(9)+"47 Copayment/Deductible"+chr(9)+"49 Preventative Care"+chr(9)+"99 Other", serv_type_recurring_06
  EditBox 165, 245, 40, 15, gross_recurring_06
  DropListBox 215, 245, 75, 15, " "+chr(9)+"1 Billing Stmt"+chr(9)+"2 Expl of Bnft"+chr(9)+"3 Cl Stmt Med Trans"+chr(9)+"4 Credit/Loan Stmt"+chr(9)+"5 Prov Statement"+chr(9)+"6 Other"+chr(9)+"No ver prvd", ver_recurring_06
  DropListBox 310, 245, 25, 10, " "+chr(9)+"H"+chr(9)+"P"+chr(9)+"M"+chr(9)+"R", bill_type_recurring_06
  ButtonGroup ButtonPressed
    OkButton 360, 130, 50, 15
    CancelButton 360, 150, 50, 15
EndDialog

'SECTION 02: CASE NUMBER

'Connecting to MAXIS
EMConnect ""

'Finding case number
call MAXIS_case_number_finder(case_number)

'Ask for case number, validate that it's numeric.
Do
	Dialog BILS_case_number_dialog	'FYI: Dialog includes checkbox for simply updating existing bills, instead of adding new ones.
	If ButtonPressed = 0 then stopscript
	transmit
	MAXIS_check_function
	If isnumeric(case_number) = False then MsgBox "Enter a valid MAXIS case number."
Loop until isnumeric(case_number) = True

'Gets to STAT/BUDG
call navigate_to_screen("stat", "budg")
EMReadScreen BUDG_check, 4, 2, 52
If BUDG_check <> "BUDG" then transmit	'ERRR prone checking

'Determines budget begin/end dates.
EMReadScreen budget_begin, 5, 10, 35
budget_begin = replace(budget_begin, " ", "/")	'MM/DD format
EMReadScreen budget_end, 5, 10, 46
budget_end = replace(budget_end, " ", "/")	'MM/DD format

'Gets to BILS
EMWriteScreen "bils", 20, 71
transmit

'IF THE WORKER REQUESTED TO UPDATE EXISTING BILS, THE SCRIPT STARTS AN ABBREVIATED IF/THEN STATEMENT----------------------------------------------------------------------------------------------------
If updating_existing_BILS_check = 1 then

	'DIALOG RUNS, PUTS BILS ON EDIT MODE AND CHECKS FOR PASSWORD PROMPT
	Dialog BILS_updater_abbreviated_dialog
	If buttonpressed = 0 then stopscript
	PF9
	EMReadScreen BILS_check, 4, 2, 54
	If BILS_check <> "BILS" then script_end_procedure("BILS not found. Did you navigate away from BILS? Did you get passworded out? The script will now close.")

	'Gets back to the first page of STAT/BILS
	Do
		PF19
		EMReadScreen first_page_check, 4, 24, 20
	Loop until first_page_check = "PAGE"

	'Cleaning up the date, so that the datediff functions work later on.
	budget_begin = replace(budget_begin, ".", "/")  'in case worker used periods instead of slashes
	budget_begin = replace(budget_begin, "-", "/") 'in case worker used dashes instead of slashes
	budget_begin = replace(budget_begin, "/", "/01/")

	'CHECKS EACH LINE IN BILS. IF THE BILL IS ONE OF THE REQUESTED UPDATES, THE SCRIPT WILL AUTOMATICALLY UPDATE THE INFORMATION WITH WHAT THE WORKER ENTERED. IT READS THE ENTIRE LINE AND SPLITS INTO AN ARRAY FOR EASE.
	MAXIS_row = 6 'setting the variable for the following do...loop
	updates_made = 0 'setting the variable to notify the worker that updates were made.
	Do
		EMReadScreen BILS_line, 54, MAXIS_row, 26
		BILS_line = replace(BILS_line, "$", " ")
		BILS_line = split(BILS_line, "  ")
		BILS_line(1) = replace(BILS_line(1), " ", "/")
		If IsDate(BILS_line(1)) = True and BILS_line(0) = ref_nbr_abbreviated then 
			If datediff("d", budget_begin, BILS_line(1)) >= 0 and BILS_line(2) = 24 and BILS_line(5) <> gross_recurring_24 and gross_recurring_24 <> "" then 
				EMWriteScreen "_________", MAXIS_row, 45
				EMWriteScreen gross_recurring_24, MAXIS_row, 45
				EMWriteScreen "c", MAXIS_row, 24
				updates_made = updates_made + 1
			End If
			If datediff("d", budget_begin, BILS_line(1)) >= 0 and BILS_line(2) = 25 and BILS_line(5) <> gross_recurring_25 and gross_recurring_25 <> "" then 
				EMWriteScreen "_________", MAXIS_row, 45
				EMWriteScreen gross_recurring_25, MAXIS_row, 45
				EMWriteScreen "c", MAXIS_row, 24
				updates_made = updates_made + 1
			End If
			If datediff("d", budget_begin, BILS_line(1)) >= 0 and BILS_line(2) = 26 and BILS_line(5) <> gross_recurring_26 and gross_recurring_26 <> "" then 
				EMWriteScreen "_________", MAXIS_row, 45
				EMWriteScreen gross_recurring_26, MAXIS_row, 45
				EMWriteScreen "c", MAXIS_row, 24
				updates_made = updates_made + 1
			End If
			If datediff("d", budget_begin, BILS_line(1)) >= 0 and BILS_line(2) = 27 and BILS_line(5) <> gross_recurring_27 and gross_recurring_27 <> "" then 
				EMWriteScreen "_________", MAXIS_row, 45
				EMWriteScreen gross_recurring_27, MAXIS_row, 45
				EMWriteScreen "c", MAXIS_row, 24
				updates_made = updates_made + 1
			End If
			If datediff("d", budget_begin, BILS_line(1)) >= 0 and BILS_line(2) = 99 and BILS_line(5) <> gross_recurring_99 and gross_recurring_99 <> "" then 
				EMWriteScreen "_________", MAXIS_row, 45
				EMWriteScreen gross_recurring_99, MAXIS_row, 45
				EMWriteScreen "c", MAXIS_row, 24
				updates_made = updates_made + 1
			End If
		End If
		MAXIS_row = MAXIS_row + 1
		If MAXIS_row = 18 then
			PF20
			EMReadScreen last_page_check, 4, 24, 19
			If last_page_check = "PAGE" then exit do
			MAXIS_row = 6
		End if
	Loop until MAXIS_row = 18 or IsDate(BILS_line(1)) = False
	transmit

	MsgBox updates_made & " entries updated."
	script_end_procedure("")
End if

'IF THE WORKER REQUESTED TO ADD NEW BILS, THE SCRIPT STARTS THE ADVANCED DIALOG----------------------------------------------------------------------------------------------------


Dialog BILS_updater_dialog
If ButtonPressed = 0 then stopscript
transmit
MAXIS_check_function

call navigate_to_screen("stat", "bils") 'In case the worker navigated out.

'Cleaning up date field
budget_begin = replace(budget_begin, ".", "/")		'in case worker used periods instead of slashes
budget_end = replace(budget_end, ".", "/")
budget_begin = replace(budget_begin, "-", "/")		'in case worker used dashes instead of slashes
budget_end = replace(budget_end, "-", "/")
budget_begin = replace(budget_begin, "/", "/01/")
budget_end = replace(budget_end, "/", "/01/")


'SECTION 05: THE SCRIPT ADDS THE BILLS INTO MAXIS
PF9

'Using working_date as a placeholder, it will now determine each footer month between the budget period start and end
working_date = budget_begin				
total_months = DateDiff("m", budget_begin, budget_end)
dim all_possible_dates_array()
redim all_possible_dates_array(total_months)					
For i = 0 to total_months
	all_possible_dates_array(i) = working_date
	working_date = DateAdd("m", 1, working_date)
Next


'Here, the script will force insurance premiums to be an "h" type bill, and remedial care will be a "p" type bill.
If serv_type_actual_01 = "25 Medicare Prem" or serv_type_actual_01 = "26 Dental or Health Prem" then bill_type_actual_01 = "H"
If serv_type_actual_01 = "27 Remedial Care" then bill_type_actual_01 = "P"
If serv_type_actual_02 = "25 Medicare Prem" or serv_type_actual_02 = "26 Dental or Health Prem" then bill_type_actual_02 = "H"
If serv_type_actual_02 = "27 Remedial Care" then bill_type_actual_02 = "P"
If serv_type_actual_03 = "25 Medicare Prem" or serv_type_actual_03 = "26 Dental or Health Prem" then bill_type_actual_03 = "H"
If serv_type_actual_03 = "27 Remedial Care" then bill_type_actual_03 = "P"
If serv_type_recurring_01 = "25 Medicare Prem" or serv_type_recurring_01 = "26 Dental or Health Prem" then bill_type_recurring_01 = "H"
If serv_type_recurring_01 = "27 Remedial Care" then bill_type_recurring_01 = "P"
If serv_type_recurring_02 = "25 Medicare Prem" or serv_type_recurring_02 = "26 Dental or Health Prem" then bill_type_recurring_02 = "H"
If serv_type_recurring_02 = "27 Remedial Care" then bill_type_recurring_02 = "P"
If serv_type_recurring_03 = "25 Medicare Prem" or serv_type_recurring_03 = "26 Dental or Health Prem" then bill_type_recurring_03 = "H"
If serv_type_recurring_03 = "27 Remedial Care" then bill_type_recurring_03 = "P"
If serv_type_recurring_04 = "25 Medicare Prem" or serv_type_recurring_04 = "26 Dental or Health Prem" then bill_type_recurring_04 = "H"
If serv_type_recurring_04 = "27 Remedial Care" then bill_type_recurring_04 = "P"
If serv_type_recurring_05 = "25 Medicare Prem" or serv_type_recurring_05 = "26 Dental or Health Prem" then bill_type_recurring_05 = "H"
If serv_type_recurring_05 = "27 Remedial Care" then bill_type_recurring_05 = "P"
If serv_type_recurring_06 = "25 Medicare Prem" or serv_type_recurring_06 = "26 Dental or Health Prem" then bill_type_recurring_06 = "H"
If serv_type_recurring_06 = "27 Remedial Care" then bill_type_recurring_06 = "P"


row = 6 'Setting the variable for the following do loop
If ref_nbr_recurring_01 <> "" then 
	For each possible_date in all_possible_dates_array
		possible_date = cdate(possible_date)
		If len(datepart("m", possible_date)) = 1 then possible_date = "0" & possible_date 
		Do
			If row = 18 then
				PF20
				row = 6
			End if
			EMReadScreen line_check, 1, row, 26
			If line_check <> "_" then row = row + 1
		Loop until line_check = "_"
		EMWriteScreen ref_nbr_recurring_01, row, 26
		EMWriteScreen left(possible_date, 2), row, 30
		EMWriteScreen "01", row, 33
		EMWriteScreen right(possible_date, 2), row, 36
		EMWriteScreen left(serv_type_recurring_01, 2), row, 40
		EMWriteScreen gross_recurring_01, row, 45
		If ver_recurring_01 = "No ver prvd" then 
			EMWriteScreen "no", row, 67
		Else
			EMWriteScreen "0" & left(ver_recurring_01, 1), row, 67
		End if
		EMWriteScreen bill_type_recurring_01, row, 71
		row = row + 1
	Next
End if

If ref_nbr_recurring_02 <> "" then 
	For each possible_date in all_possible_dates_array
		possible_date = cdate(possible_date)
		If len(datepart("m", possible_date)) = 1 then possible_date = "0" & possible_date 
		Do
			If row = 18 then
				PF20
				row = 6
			End if
			EMReadScreen line_check, 1, row, 26
			If line_check <> "_" then row = row + 1
		Loop until line_check = "_"
		EMWriteScreen ref_nbr_recurring_02, row, 26
		EMWriteScreen left(possible_date, 2), row, 30
		EMWriteScreen "01", row, 33
		EMWriteScreen right(possible_date, 2), row, 36
		EMWriteScreen left(serv_type_recurring_02, 2), row, 40
		EMWriteScreen gross_recurring_02, row, 45
		If ver_recurring_02 = "No ver prvd" then 
			EMWriteScreen "no", row, 67
		Else
			EMWriteScreen "0" & left(ver_recurring_02, 1), row, 67
		End if
		EMWriteScreen bill_type_recurring_02, row, 71
		row = row + 1
	Next
End if

If ref_nbr_recurring_03 <> "" then 
	For each possible_date in all_possible_dates_array
		possible_date = cdate(possible_date)
		If len(datepart("m", possible_date)) = 1 then possible_date = "0" & possible_date 
		Do
			If row = 18 then
				PF20
				row = 6
			End if
			EMReadScreen line_check, 1, row, 26
			If line_check <> "_" then row = row + 1
		Loop until line_check = "_"
		EMWriteScreen ref_nbr_recurring_03, row, 26
		EMWriteScreen left(possible_date, 2), row, 30
		EMWriteScreen "01", row, 33
		EMWriteScreen right(possible_date, 2), row, 36
		EMWriteScreen left(serv_type_recurring_03, 2), row, 40
		EMWriteScreen gross_recurring_03, row, 45
		If ver_recurring_03 = "No ver prvd" then 
			EMWriteScreen "no", row, 67
		Else
			EMWriteScreen "0" & left(ver_recurring_03, 1), row, 67
		End if
		EMWriteScreen bill_type_recurring_03, row, 71
		row = row + 1
	Next
End if

If ref_nbr_recurring_04 <> "" then 
	For each possible_date in all_possible_dates_array
		possible_date = cdate(possible_date)
		If len(datepart("m", possible_date)) = 1 then possible_date = "0" & possible_date 
		Do
			If row = 18 then
				PF20
				row = 6
			End if
			EMReadScreen line_check, 1, row, 26
			If line_check <> "_" then row = row + 1
		Loop until line_check = "_"
		EMWriteScreen ref_nbr_recurring_04, row, 26
		EMWriteScreen left(possible_date, 2), row, 30
		EMWriteScreen "01", row, 33
		EMWriteScreen right(possible_date, 2), row, 36
		EMWriteScreen left(serv_type_recurring_04, 2), row, 40
		EMWriteScreen gross_recurring_04, row, 45
		If ver_recurring_04 = "No ver prvd" then 
			EMWriteScreen "no", row, 67
		Else
			EMWriteScreen "0" & left(ver_recurring_04, 1), row, 67
		End if
		EMWriteScreen bill_type_recurring_04, row, 71
		row = row + 1
	Next
End if

If ref_nbr_recurring_05 <> "" then 
	For each possible_date in all_possible_dates_array
		possible_date = cdate(possible_date)
		If len(datepart("m", possible_date)) = 1 then possible_date = "0" & possible_date 
		Do
			If row = 18 then
				PF20
				row = 6
			End if
			EMReadScreen line_check, 1, row, 26
			If line_check <> "_" then row = row + 1
		Loop until line_check = "_"
		EMWriteScreen ref_nbr_recurring_05, row, 26
		EMWriteScreen left(possible_date, 2), row, 30
		EMWriteScreen "01", row, 33
		EMWriteScreen right(possible_date, 2), row, 36
		EMWriteScreen left(serv_type_recurring_05, 2), row, 40
		EMWriteScreen gross_recurring_05, row, 45
		If ver_recurring_05 = "No ver prvd" then 
			EMWriteScreen "no", row, 67
		Else
			EMWriteScreen "0" & left(ver_recurring_05, 1), row, 67
		End if
		EMWriteScreen bill_type_recurring_05, row, 71
		row = row + 1
	Next
End if

If ref_nbr_recurring_06 <> "" then 
	For each possible_date in all_possible_dates_array
		possible_date = cdate(possible_date)
		If len(datepart("m", possible_date)) = 1 then possible_date = "0" & possible_date 
		Do
			If row = 18 then
				PF20
				row = 6
			End if
			EMReadScreen line_check, 1, row, 26
			If line_check <> "_" then row = row + 1
		Loop until line_check = "_"
		EMWriteScreen ref_nbr_recurring_06, row, 26
		EMWriteScreen left(possible_date, 2), row, 30
		EMWriteScreen "01", row, 33
		EMWriteScreen right(possible_date, 2), row, 36
		EMWriteScreen left(serv_type_recurring_06, 2), row, 40
		EMWriteScreen gross_recurring_06, row, 45
		If ver_recurring_06 = "No ver prvd" then 
			EMWriteScreen "no", row, 67
		Else
			EMWriteScreen "0" & left(ver_recurring_06, 1), row, 67
		End if
		EMWriteScreen bill_type_recurring_06, row, 71
		row = row + 1
	Next
End if

'Now actual expenses

If ref_nbr_actual_01 <> "" then 
	Do
		If row = 18 then
			PF20
			row = 6
		End if
		EMReadScreen line_check, 1, row, 26
		If line_check <> "_" then row = row + 1
	Loop until line_check = "_"
	EMWriteScreen ref_nbr_actual_01, row, 26
	working_month = datepart("m", date_actual_01)
	If len(working_month) = 1 then working_month = "0" & working_month
	EMWriteScreen working_month, row, 30
	working_day = datepart("d", date_actual_01)
	If len(working_day) = 1 then working_day = "0" & working_day
	EMWriteScreen working_day, row, 33
	working_year = datepart("yyyy", date_actual_01)
	working_year = working_year - 2000
	EMWriteScreen working_year, row, 36
	EMWriteScreen left(serv_type_actual_01, 2), row, 40
	EMWriteScreen gross_actual_01, row, 45
	If ver_actual_01 = "No ver prvd" then 
		EMWriteScreen "no", row, 67
	Else
		EMWriteScreen "0" & left(ver_actual_01, 1), row, 67
	End if
	EMWriteScreen bill_type_actual_01, row, 71
	row = row + 1
End if

If ref_nbr_actual_02 <> "" then 
	Do
		If row = 18 then
			PF20
			row = 6
		End if
		EMReadScreen line_check, 1, row, 26
		If line_check <> "_" then row = row + 1
	Loop until line_check = "_"
	EMWriteScreen ref_nbr_actual_02, row, 26
	working_month = datepart("m", date_actual_02)
	If len(working_month) = 1 then working_month = "0" & working_month
	EMWriteScreen working_month, row, 30
	working_day = datepart("d", date_actual_02)
	If len(working_day) = 1 then working_day = "0" & working_day
	EMWriteScreen working_day, row, 33
	working_year = datepart("yyyy", date_actual_02)
	working_year = working_year - 2000
	EMWriteScreen working_year, row, 36
	EMWriteScreen left(serv_type_actual_02, 2), row, 40
	EMWriteScreen gross_actual_02, row, 45
	If ver_actual_02 = "No ver prvd" then 
		EMWriteScreen "no", row, 67
	Else
		EMWriteScreen "0" & left(ver_actual_02, 1), row, 67
	End if
	EMWriteScreen bill_type_actual_02, row, 71
	row = row + 1
End if

If ref_nbr_actual_03 <> "" then 
	Do
		If row = 18 then
			PF20
			row = 6
		End if
		EMReadScreen line_check, 1, row, 26
		If line_check <> "_" then row = row + 1
	Loop until line_check = "_"
	EMWriteScreen ref_nbr_actual_03, row, 26
	working_month = datepart("m", date_actual_03)
	If len(working_month) = 1 then working_month = "0" & working_month
	EMWriteScreen working_month, row, 30
	working_day = datepart("d", date_actual_03)
	If len(working_day) = 1 then working_day = "0" & working_day
	EMWriteScreen working_day, row, 33
	working_year = datepart("yyyy", date_actual_03)
	working_year = working_year - 2000
	EMWriteScreen working_year, row, 36
	EMWriteScreen left(serv_type_actual_03, 2), row, 40
	EMWriteScreen gross_actual_03, row, 45
	If ver_actual_03 = "No ver prvd" then 
		EMWriteScreen "no", row, 67
	Else
		EMWriteScreen "0" & left(ver_actual_03, 1), row, 67
	End if
	EMWriteScreen bill_type_actual_03, row, 71
	row = row + 1
End if

script_end_procedure("")







