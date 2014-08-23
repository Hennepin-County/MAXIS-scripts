'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - copy screens to Word"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'----------------------------------------------------------------------------------------------------
'ADD TO FUNCTIONS FILE WHEN GITHUB IS WORKING AGAIN
Function copy_screen_to_array(output_array)
	output_array = "" 'resetting array
	Dim screenarray(23)	'24 line array
	row = 1
	For each line in screenarray
		EMReadScreen reading_line, 80, row, 1
		output_array = output_array & reading_line & "UUDDLRLRBA"
		row = row + 1
	Next
	output_array = split(output_array, "UUDDLRLRBA")
End function

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
all_possible_panels = "MEMB|MEMI|ADDR|AREP|ALTP|ALIA|TYPE|PROG|HCRE|PARE|SIBL|EATS|IMIG|SPON|FACI|FCFC|FCPL|ADME|REMO|DISA|ABPS|PREG|STRK|STWK|SCHL|WREG|EMPS|CASH|ACCT|SECU|CARS|REST|OTHR|TRAN|STIN|STEC|PBEN|UNEA|LUMP|RBIC|BUSI|JOBS|TRAC|DSTT|DCEX|WKEX|COEX|SHEL|HEST|ACUT|PDED|PACT|FMED|ACCI|MEDI|INSA|DIET|DISQ|SWKR|REVW|MISC|RESI|TIME|EMMA|BILS|HCMI|BUDG|SANC|WBSN|MMSA|DFLN|MSUR"


'DIALOG IS TOO LARGE FOR DIALOG EDITOR, CREATED MANUALLY
BeginDialog all_MAXIS_panels_dialog, 0, 0, 311, 190, "All MAXIS panels dialog"
  Checkbox 10, 10, 35, 10, "MEMB", MEMB_check
  Checkbox 10, 25, 35, 10, "TYPE", TYPE_check
  Checkbox 10, 40, 35, 10, "IMIG", IMIG_check
  Checkbox 10, 55, 35, 10, "REMO", REMO_check
  Checkbox 10, 70, 35, 10, "SCHL", SCHL_check
  Checkbox 10, 85, 35, 10, "CARS", CARS_check
  Checkbox 10, 100, 35, 10, "PBEN", PBEN_check
  Checkbox 10, 115, 35, 10, "TRAC", TRAC_check
  Checkbox 10, 130, 35, 10, "HEST", HEST_check
  Checkbox 10, 145, 35, 10, "MEDI", MEDI_check
  Checkbox 10, 160, 35, 10, "MISC", MISC_check
  Checkbox 10, 175, 35, 10, "BUDG", BUDG_check
  Checkbox 60, 10, 35, 10, "MEMI", MEMI_check
  Checkbox 60, 25, 35, 10, "PROG", PROG_check
  Checkbox 60, 40, 35, 10, "SPON", SPON_check
  Checkbox 60, 55, 35, 10, "DISA", DISA_check
  Checkbox 60, 70, 35, 10, "WREG", WREG_check
  Checkbox 60, 85, 35, 10, "REST", REST_check
  Checkbox 60, 100, 35, 10, "UNEA", UNEA_check
  Checkbox 60, 115, 35, 10, "DSTT", DSTT_check
  Checkbox 60, 130, 35, 10, "ACUT", ACUT_check
  Checkbox 60, 145, 35, 10, "INSA", INSA_check
  Checkbox 60, 160, 35, 10, "RESI", RESI_check
  Checkbox 60, 175, 35, 10, "SANC", SANC_check
  Checkbox 110, 10, 35, 10, "ADDR", ADDR_check
  Checkbox 110, 25, 35, 10, "HCRE", HCRE_check
  Checkbox 110, 40, 35, 10, "FACI", FACI_check
  Checkbox 110, 55, 35, 10, "ABPS", ABPS_check
  Checkbox 110, 70, 35, 10, "EMPS", EMPS_check
  Checkbox 110, 85, 35, 10, "OTHR", OTHR_check
  Checkbox 110, 100, 35, 10, "LUMP", LUMP_check
  Checkbox 110, 115, 35, 10, "DCEX", DCEX_check
  Checkbox 110, 130, 35, 10, "PDED", PDED_check
  Checkbox 110, 145, 35, 10, "DIET", DIET_check
  Checkbox 110, 160, 35, 10, "TIME", TIME_check
  Checkbox 110, 175, 35, 10, "WBSN", WBSN_check
  Checkbox 160, 10, 35, 10, "AREP", AREP_check
  Checkbox 160, 25, 35, 10, "PARE", PARE_check
  Checkbox 160, 40, 35, 10, "FCFC", FCFC_check
  Checkbox 160, 55, 35, 10, "PREG", PREG_check
  Checkbox 160, 70, 35, 10, "CASH", CASH_check
  Checkbox 160, 85, 35, 10, "TRAN", TRAN_check
  Checkbox 160, 100, 35, 10, "RBIC", RBIC_check
  Checkbox 160, 115, 35, 10, "WKEX", WKEX_check
  Checkbox 160, 130, 35, 10, "PACT", PACT_check
  Checkbox 160, 145, 35, 10, "DISQ", DISQ_check
  Checkbox 160, 160, 35, 10, "EMMA", EMMA_check
  Checkbox 160, 175, 35, 10, "MMSA", MMSA_check
  Checkbox 210, 10, 35, 10, "ALTP", ALTP_check
  Checkbox 210, 25, 35, 10, "SIBL", SIBL_check
  Checkbox 210, 40, 35, 10, "FCPL", FCPL_check
  Checkbox 210, 55, 35, 10, "STRK", STRK_check
  Checkbox 210, 70, 35, 10, "ACCT", ACCT_check
  Checkbox 210, 85, 35, 10, "STIN", STIN_check
  Checkbox 210, 100, 35, 10, "BUSI", BUSI_check
  Checkbox 210, 115, 35, 10, "COEX", COEX_check
  Checkbox 210, 130, 35, 10, "FMED", FMED_check
  Checkbox 210, 145, 35, 10, "SWKR", SWKR_check
  Checkbox 210, 160, 35, 10, "BILS", BILS_check
  Checkbox 210, 175, 35, 10, "DFLN", DFLN_check
  Checkbox 260, 190, 35, 10, "ALIA", ALIA_check
  Checkbox 260, 205, 35, 10, "EATS", EATS_check
  Checkbox 260, 220, 35, 10, "ADME", ADME_check
  Checkbox 260, 235, 35, 10, "STWK", STWK_check
  Checkbox 260, 250, 35, 10, "SECU", SECU_check
  Checkbox 260, 265, 35, 10, "STEC", STEC_check
  Checkbox 260, 280, 35, 10, "JOBS", JOBS_check
  Checkbox 260, 295, 35, 10, "SHEL", SHEL_check
  Checkbox 260, 310, 35, 10, "ACCI", ACCI_check
  Checkbox 260, 325, 35, 10, "REVW", REVW_check
  Checkbox 260, 340, 35, 10, "HCMI", HCMI_check
  Checkbox 260, 355, 35, 10, "MSUR", MSUR_check
  ButtonGroup ButtonPressed
    OkButton 255, 5, 50, 15
    CancelButton 255, 25, 50, 15
EndDialog

BeginDialog case_number_and_footer_month_dialog, 0, 0, 161, 65, "Case number and footer month"
  Text 5, 10, 85, 10, "Enter your case number:"
  EditBox 95, 5, 60, 15, case_number
  Text 15, 30, 50, 10, "Footer month:"
  EditBox 65, 25, 25, 15, footer_month
  Text 95, 30, 20, 10, "Year:"
  EditBox 120, 25, 25, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 45, 50, 15
    CancelButton 85, 45, 50, 15
EndDialog



'<<<REPLACE WITH DIALOG
'case_number = "201471"
'all_panels_selected_array = all_possible_panels

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Finds MAXIS case number
call MAXIS_case_number_finder(case_number)

'Finds MAXIS footer month
row = 1
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

'Shows case number dialog
Do
	Dialog case_number_and_footer_month_dialog
	If buttonpressed = 0 then stopscript
	If isnumeric(case_number) = False then MsgBox "You must type a valid case number."
Loop until isnumeric(case_number) = True

'Shows the MAXIS panel selection dialog
Dialog all_MAXIS_panels_dialog
If buttonpressed = 0 then stopscript

'Adding checked objects to the array
If MEMB_check = checked then all_panels_selected_array = all_panels_selected_array & "MEMB" & "|"
If TYPE_check = checked then all_panels_selected_array = all_panels_selected_array & "TYPE" & "|"
If IMIG_check = checked then all_panels_selected_array = all_panels_selected_array & "IMIG" & "|"
If REMO_check = checked then all_panels_selected_array = all_panels_selected_array & "REMO" & "|"
If SCHL_check = checked then all_panels_selected_array = all_panels_selected_array & "SCHL" & "|"
If CARS_check = checked then all_panels_selected_array = all_panels_selected_array & "CARS" & "|"
If PBEN_check = checked then all_panels_selected_array = all_panels_selected_array & "PBEN" & "|"
If TRAC_check = checked then all_panels_selected_array = all_panels_selected_array & "TRAC" & "|"
If HEST_check = checked then all_panels_selected_array = all_panels_selected_array & "HEST" & "|"
If MEDI_check = checked then all_panels_selected_array = all_panels_selected_array & "MEDI" & "|"
If MISC_check = checked then all_panels_selected_array = all_panels_selected_array & "MISC" & "|"
If BUDG_check = checked then all_panels_selected_array = all_panels_selected_array & "BUDG" & "|"
If MEMI_check = checked then all_panels_selected_array = all_panels_selected_array & "MEMI" & "|"
If PROG_check = checked then all_panels_selected_array = all_panels_selected_array & "PROG" & "|"
If SPON_check = checked then all_panels_selected_array = all_panels_selected_array & "SPON" & "|"
If DISA_check = checked then all_panels_selected_array = all_panels_selected_array & "DISA" & "|"
If WREG_check = checked then all_panels_selected_array = all_panels_selected_array & "WREG" & "|"
If REST_check = checked then all_panels_selected_array = all_panels_selected_array & "REST" & "|"
If UNEA_check = checked then all_panels_selected_array = all_panels_selected_array & "UNEA" & "|"
If DSTT_check = checked then all_panels_selected_array = all_panels_selected_array & "DSTT" & "|"
If ACUT_check = checked then all_panels_selected_array = all_panels_selected_array & "ACUT" & "|"
If INSA_check = checked then all_panels_selected_array = all_panels_selected_array & "INSA" & "|"
If RESI_check = checked then all_panels_selected_array = all_panels_selected_array & "RESI" & "|"
If SANC_check = checked then all_panels_selected_array = all_panels_selected_array & "SANC" & "|"
If ADDR_check = checked then all_panels_selected_array = all_panels_selected_array & "ADDR" & "|"
If HCRE_check = checked then all_panels_selected_array = all_panels_selected_array & "HCRE" & "|"
If FACI_check = checked then all_panels_selected_array = all_panels_selected_array & "FACI" & "|"
If ABPS_check = checked then all_panels_selected_array = all_panels_selected_array & "ABPS" & "|"
If EMPS_check = checked then all_panels_selected_array = all_panels_selected_array & "EMPS" & "|"
If OTHR_check = checked then all_panels_selected_array = all_panels_selected_array & "OTHR" & "|"
If LUMP_check = checked then all_panels_selected_array = all_panels_selected_array & "LUMP" & "|"
If DCEX_check = checked then all_panels_selected_array = all_panels_selected_array & "DCEX" & "|"
If PDED_check = checked then all_panels_selected_array = all_panels_selected_array & "PDED" & "|"
If DIET_check = checked then all_panels_selected_array = all_panels_selected_array & "DIET" & "|"
If TIME_check = checked then all_panels_selected_array = all_panels_selected_array & "TIME" & "|"
If WBSN_check = checked then all_panels_selected_array = all_panels_selected_array & "WBSN" & "|"
If AREP_check = checked then all_panels_selected_array = all_panels_selected_array & "AREP" & "|"
If PARE_check = checked then all_panels_selected_array = all_panels_selected_array & "PARE" & "|"
If FCFC_check = checked then all_panels_selected_array = all_panels_selected_array & "FCFC" & "|"
If PREG_check = checked then all_panels_selected_array = all_panels_selected_array & "PREG" & "|"
If CASH_check = checked then all_panels_selected_array = all_panels_selected_array & "CASH" & "|"
If TRAN_check = checked then all_panels_selected_array = all_panels_selected_array & "TRAN" & "|"
If RBIC_check = checked then all_panels_selected_array = all_panels_selected_array & "RBIC" & "|"
If WKEX_check = checked then all_panels_selected_array = all_panels_selected_array & "WKEX" & "|"
If PACT_check = checked then all_panels_selected_array = all_panels_selected_array & "PACT" & "|"
If DISQ_check = checked then all_panels_selected_array = all_panels_selected_array & "DISQ" & "|"
If EMMA_check = checked then all_panels_selected_array = all_panels_selected_array & "EMMA" & "|"
If MMSA_check = checked then all_panels_selected_array = all_panels_selected_array & "MMSA" & "|"
If ALTP_check = checked then all_panels_selected_array = all_panels_selected_array & "ALTP" & "|"
If SIBL_check = checked then all_panels_selected_array = all_panels_selected_array & "SIBL" & "|"
If FCPL_check = checked then all_panels_selected_array = all_panels_selected_array & "FCPL" & "|"
If STRK_check = checked then all_panels_selected_array = all_panels_selected_array & "STRK" & "|"
If ACCT_check = checked then all_panels_selected_array = all_panels_selected_array & "ACCT" & "|"
If STIN_check = checked then all_panels_selected_array = all_panels_selected_array & "STIN" & "|"
If BUSI_check = checked then all_panels_selected_array = all_panels_selected_array & "BUSI" & "|"
If COEX_check = checked then all_panels_selected_array = all_panels_selected_array & "COEX" & "|"
If FMED_check = checked then all_panels_selected_array = all_panels_selected_array & "FMED" & "|"
If SWKR_check = checked then all_panels_selected_array = all_panels_selected_array & "SWKR" & "|"
If BILS_check = checked then all_panels_selected_array = all_panels_selected_array & "BILS" & "|"
If DFLN_check = checked then all_panels_selected_array = all_panels_selected_array & "DFLN" & "|"
If ALIA_check = checked then all_panels_selected_array = all_panels_selected_array & "ALIA" & "|"
If EATS_check = checked then all_panels_selected_array = all_panels_selected_array & "EATS" & "|"
If ADME_check = checked then all_panels_selected_array = all_panels_selected_array & "ADME" & "|"
If STWK_check = checked then all_panels_selected_array = all_panels_selected_array & "STWK" & "|"
If SECU_check = checked then all_panels_selected_array = all_panels_selected_array & "SECU" & "|"
If STEC_check = checked then all_panels_selected_array = all_panels_selected_array & "STEC" & "|"
If JOBS_check = checked then all_panels_selected_array = all_panels_selected_array & "JOBS" & "|"
If SHEL_check = checked then all_panels_selected_array = all_panels_selected_array & "SHEL" & "|"
If ACCI_check = checked then all_panels_selected_array = all_panels_selected_array & "ACCI" & "|"
If REVW_check = checked then all_panels_selected_array = all_panels_selected_array & "REVW" & "|"
If HCMI_check = checked then all_panels_selected_array = all_panels_selected_array & "HCMI" & "|"
If MSUR_check = checked then all_panels_selected_array = all_panels_selected_array & "MSUR" & "|"

'Creates the Word doc
Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.PageSetup.LeftMargin = 50
objSelection.PageSetup.RightMargin = 50
objSelection.PageSetup.TopMargin = 30
objSelection.PageSetup.BottomMargin = 30
objSelection.Font.Name = "Courier New"
objSelection.Font.Size = "10"





'Splits the array
all_panels_selected_array = split(all_panels_selected_array, "|")

For each panel_to_scan in all_panels_selected_array
	'Declares the MAXIS_row variable, to be used in reading additional HH members
	MAXIS_row = 5

	'Goes to the screen for the first HH memb
	call navigate_to_screen("STAT", panel_to_scan)

	Do
		'Some panels do not have options for multiple members. Any other panels require an "01" to be entered, to look at first HH member.
		If panel_to_scan <> "ADDR" and _
		panel_to_scan <> "AREP" and _
		panel_to_scan <> "ALTP" and _
		panel_to_scan <> "TYPE" and _
		panel_to_scan <> "PROG" and _
		panel_to_scan <> "HCRE" and _
		panel_to_scan <> "SIBL" and _
		panel_to_scan <> "EATS" and _
		panel_to_scan <> "FCFC" and _
		panel_to_scan <> "FCPL" and _
		panel_to_scan <> "ABPS" and _
		panel_to_scan <> "AREP" and _
		panel_to_scan <> "DSTT" and _
		panel_to_scan <> "HEST" and _
		panel_to_scan <> "PACT" and _
		panel_to_scan <> "INSA" and _
		panel_to_scan <> "SWKR" and _
		panel_to_scan <> "REVW" and _
		panel_to_scan <> "MISC" and _
		panel_to_scan <> "RESI" and _
		panel_to_scan <> "BILS" and _
		panel_to_scan <> "BUDG" and _
		panel_to_scan <> "MEMB" and _
		panel_to_scan <> "MMSA" then
			EMReadScreen next_HH_ref_nbr, 2, MAXIS_row, 3		'Now it determines who the next HH member is, and goes to their panel.
			EMWriteScreen next_HH_ref_nbr, 20, 76			'Puts in the next ref number
			transmit								'Transmits to load the data
			MAXIS_row = MAXIS_row + 1					'Adds one to the MAXIS_row, to be used again
		End if

		Do
			'Reads current screen
			call copy_screen_to_array(screentest)

			'Adds current screen to Word doc
			For each line in screentest
				objSelection.TypeText line & Chr(11)
			Next

			'Determines if the Word doc needs a new page
			If screen_on_page = "" or screen_on_page = 1 then
				screen_on_page = 2
				objSelection.TypeText vbCr & vbCr
			Elseif screen_on_page = 2 then
				screen_on_page = 1
				objSelection.InsertBreak(7)
			End if

			'Checks to see if there's more than one panel. If there is, it'll get to the next one.
			EMReadScreen current_panel, 2, 2, 72
			EMReadScreen amt_of_panels, 2, 2, 78
			If cint(current_panel) < cint(amt_of_panels) then transmit
		Loop until cint(current_panel) >= cint(amt_of_panels)

		'Some panels do not have options for multiple members. These will exit this part of the do loop.
		If panel_to_scan = "ADDR" or _
		panel_to_scan = "AREP" or _
		panel_to_scan = "ALTP" or _
		panel_to_scan = "TYPE" or _
		panel_to_scan = "PROG" or _
		panel_to_scan = "HCRE" or _
		panel_to_scan = "SIBL" or _
		panel_to_scan = "EATS" or _
		panel_to_scan = "FCFC" or _
		panel_to_scan = "FCPL" or _
		panel_to_scan = "ABPS" or _
		panel_to_scan = "AREP" or _
		panel_to_scan = "DSTT" or _
		panel_to_scan = "HEST" or _
		panel_to_scan = "PACT" or _
		panel_to_scan = "INSA" or _
		panel_to_scan = "SWKR" or _
		panel_to_scan = "REVW" or _
		panel_to_scan = "MISC" or _
		panel_to_scan = "RESI" or _
		panel_to_scan = "BILS" or _
		panel_to_scan = "BUDG" or _
		panel_to_scan = "MEMB" or _
		panel_to_scan = "MMSA" then 
MsgBox "found"
exit do
End if
	Loop until next_HH_ref_nbr = "  "
Next