'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - INTAKE APPROVAL.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000


'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 72, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 50, 50, 15
    CancelButton 95, 50, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
EndDialog


BeginDialog intake_approval_dialog, 0, 0, 386, 435, "Intake Approval Dialog"
  EditBox 65, 5, 45, 15, application_date
  CheckBox 140, 5, 155, 15, "Check here if this client is in the community.", community_check
  DropListBox 45, 25, 30, 15, "EX"+chr(9)+"DX"+chr(9)+"DP", elig_type
  DropListBox 135, 25, 30, 15, "L"+chr(9)+"S"+chr(9)+"B", budget_type
  EditBox 305, 25, 70, 15, recipient_amt
  CheckBox 5, 45, 140, 15, "LTCC? If so, check here and enter date:", LTCC_check
  EditBox 145, 45, 45, 15, LTCC_date
  CheckBox 210, 45, 75, 15, "DHS-5181 on file?", DHS_5181_on_file_check
  CheckBox 305, 45, 75, 15, "DHS-1503 on file?", DHS_1503_on_file_check
  EditBox 55, 65, 45, 15, retro_months
  EditBox 185, 65, 45, 15, month_MA_starts
  EditBox 330, 65, 45, 15, month_MA_LTC_starts
  EditBox 65, 85, 60, 15, baseline_date
  EditBox 250, 85, 125, 15, AREP_SWKR
  EditBox 75, 105, 205, 15, FACI
  EditBox 320, 105, 55, 15, CFR
  EditBox 60, 125, 315, 15, income
  EditBox 40, 145, 335, 15, assets
  EditBox 90, 165, 65, 15, total_countable_assets
  EditBox 235, 165, 140, 15, other_asset_notes
  EditBox 60, 185, 150, 15, MEDI_INSA
  CheckBox 235, 185, 140, 15, "Check here if INSA was loaded into TPL.", INSA_loaded_into_TPL_check
  CheckBox 5, 205, 230, 10, "LTC partnership? If so, check here and enter a separate case note.", LTC_partnership_check
  CheckBox 235, 205, 105, 15, "Managed care referral sent?", managed_care_referral_sent_check
  EditBox 70, 225, 305, 15, annuity_LTC_PRB
  DropListBox 70, 245, 75, 15, "N/A"+chr(9)+"Within limit"+chr(9)+"Beyond limit", home_equity_limit
  EditBox 190, 245, 185, 15, transfer
  EditBox 70, 265, 305, 15, deductions
  EditBox 50, 285, 325, 15, other_notes
  EditBox 55, 305, 320, 15, actions_taken
  CheckBox 10, 325, 85, 15, "Sent DHS-3050/1503?", DHS_3050_1503_check
  CheckBox 135, 325, 95, 15, "Sent DHS-3203/lien doc?", DHS_3203_lien_doc_check
  CheckBox 275, 325, 95, 15, "Asset transfer memo sent?", asset_transfer_letter_sent_check
  EditBox 195, 400, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 270, 400, 50, 15
    CancelButton 325, 400, 50, 15
    PushButton 340, 5, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 190, 85, 25, 10, "AREP/", AREP_button
    PushButton 215, 85, 30, 10, "SWKR:", SWKR_button
    PushButton 5, 105, 65, 10, "FACI (if applicable):", FACI_button
    PushButton 5, 125, 50, 10, "UNEA/income:", UNEA_button
    PushButton 5, 185, 25, 10, "MEDI/", MEDI_button
    PushButton 30, 185, 25, 10, "INSA:", INSA_button
    PushButton 5, 265, 60, 10, "BILS/deductions:", BILS_button
  Text 5, 5, 55, 15, "Application date:"
  Text 5, 25, 35, 15, "Elig type:"
  Text 85, 25, 45, 15, "Budget type:"
  Text 195, 25, 110, 15, "Waiver obilgation/recipient amt:"
  Text 5, 65, 50, 15, "Retro months?:"
  Text 125, 65, 55, 15, "Month MA starts:"
  Text 255, 65, 75, 15, "Month MA-LTC starts:"
  Text 300, 105, 20, 15, "CFR:"
  Text 5, 145, 30, 15, "Assets:"
  Text 5, 165, 80, 15, "Total countable assets:"
  Text 165, 165, 65, 15, "Other asset notes:"
  Text 5, 85, 60, 15, "Baseline date: "
  Text 5, 225, 65, 15, "Annuity (LTC) PRB:"
  Text 5, 245, 60, 15, "Home equity limit:"
  Text 155, 245, 35, 15, "Transfer:"
  Text 5, 285, 40, 15, "Other notes:"
  Text 5, 305, 50, 15, "Actions taken:"
  Text 130, 400, 60, 15, "Worker signature:"
  Text 15, 350, 345, 40, "Per HCPM 19.40.15: The baseline date is the date in which both of the following conditions are met:  1. A person is residing in an LTCF or, for a person requesting services through a home and community-based waiver program, the date a screening occurred that indicated a need for services provided through a home and community-based services waiver program AND 2. The person’s initial request month for MA payment of LTC services."
  GroupBox 5, 340, 365, 55, ""
EndDialog


BeginDialog type_std_dialog, 0, 0, 206, 172, "Type-Std dialog"
  EditBox 10, 25, 50, 15, elig_date_01
  EditBox 75, 25, 40, 15, elig_type_std_01
  EditBox 135, 25, 15, 15, elig_method_01
  EditBox 175, 25, 15, 15, elig_waiver_type_01
  EditBox 10, 45, 50, 15, elig_date_02
  EditBox 75, 45, 40, 15, elig_type_std_02
  EditBox 135, 45, 15, 15, elig_method_02
  EditBox 175, 45, 15, 15, elig_waiver_type_02
  EditBox 10, 65, 50, 15, elig_date_03
  EditBox 75, 65, 40, 15, elig_type_std_03
  EditBox 135, 65, 15, 15, elig_method_03
  EditBox 175, 65, 15, 15, elig_waiver_type_03
  EditBox 10, 85, 50, 15, elig_date_04
  EditBox 75, 85, 40, 15, elig_type_std_04
  EditBox 135, 85, 15, 15, elig_method_04
  EditBox 175, 85, 15, 15, elig_waiver_type_04
  EditBox 10, 105, 50, 15, elig_date_05
  EditBox 75, 105, 40, 15, elig_type_std_05
  EditBox 135, 105, 15, 15, elig_method_05
  EditBox 175, 105, 15, 15, elig_waiver_type_05
  EditBox 10, 125, 50, 15, elig_date_06
  EditBox 75, 125, 40, 15, elig_type_std_06
  EditBox 135, 125, 15, 15, elig_method_06
  EditBox 175, 125, 15, 15, elig_waiver_type_06
  ButtonGroup ButtonPressed
    OkButton 50, 150, 50, 15
    CancelButton 110, 150, 50, 15
  Text 15, 10, 40, 10, "Elig months"
  Text 75, 10, 45, 10, "Elig type/Std"
  Text 130, 10, 30, 10, "Method"
  Text 170, 10, 25, 10, "Waiver"
EndDialog


'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col
HC_check = 1 'This is so the functions will work without having to select a program. It uses the same dialogs as the CSR, which can look in multiple places. This is HC only, so it doesn't need those.
application_signed_check = 1 'The script should default to having the application signed.


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Grabbing the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Grabbing the footer month/year
call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then 
  footer_month = MAXIS_footer_month
  call find_variable("Month: " & footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then footer_year = MAXIS_footer_year
End if

'Showing the case number dialog
Do
  Dialog case_number_dialog
  If ButtonPressed = 0 then stopscript
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for MAXIS
transmit
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then script_end_procedure("You are not in MAXIS or you are locked out of your case.")

'Navigating into STAT
call navigate_to_MAXIS_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact a Support Team member.")
EMReadScreen ERRR_check, 4, 2, 52
If ERRR_check = "ERRR" then transmit 'For error prone cases.

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofill for the application_date variable, then determines lookback period based on the info
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", application_date)
If application_date <> "" then lookback_period = dateadd("d", -1, dateadd("m", -60, cdate(application_date))) & ""

'Autofilling the rest of the STAT stuff----------------------------------------------------------------------------------------------------
call autofill_editbox_from_MAXIS(HH_member_array, "ACCT", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP_SWKR)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", income)
call autofill_editbox_from_MAXIS(HH_member_array, "CARS", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "CASH", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_months)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", MEDI_INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", income)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", MEDI_INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "OTHR", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", income)
call autofill_editbox_from_MAXIS(HH_member_array, "REST", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "SECU", assets)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", income)


'Going to ELIG/HC for the correct footer month
back_to_self
EMWriteScreen "elig", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen "hc", 21, 70
transmit

'Checks for person 01 and navigates to it
EMReadScreen person_check, 2, 8, 31
If person_check = "NO" then
  MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  EMWriteScreen "x", 9, 26
End if
If person_check <> "NO" then EMWriteScreen "x", 8, 26
transmit

'SEARCHES FOR FOOTER MONTH AND YEAR
row = 6
col = 1
EMSearch " " & footer_month & "/" & footer_year & " ", row, col
If row = 0 then script_end_procedure("A " & footer_month & "/" & footer_year & " span could not be found. Try this again. You may need to run the case through background.")

'MAKES THE MONTH_MA-LTC_STARTS VARIABLE THE FOOTER MONTH AND YEAR
month_MA_LTC_starts = footer_month & "/" & footer_year

'GRABS ELIG TYPE INFO FROM ELIG/HC
EMReadScreen elig_type_std_01, 6, 12, 17
EMReadScreen elig_type_std_02, 6, 12, 28
EMReadScreen elig_type_std_03, 6, 12, 39
EMReadScreen elig_type_std_04, 6, 12, 50
EMReadScreen elig_type_std_05, 6, 12, 61
EMReadScreen elig_type_std_06, 6, 12, 72
elig_type_std_01 = replace(elig_type_std_01, " ", "")
elig_type_std_02 = replace(elig_type_std_02, " ", "")
elig_type_std_03 = replace(elig_type_std_03, " ", "")
elig_type_std_04 = replace(elig_type_std_04, " ", "")
elig_type_std_05 = replace(elig_type_std_05, " ", "")
elig_type_std_06 = replace(elig_type_std_06, " ", "")
EMReadScreen elig_method_01, 1, 13, 21
EMReadScreen elig_method_02, 1, 13, 32
EMReadScreen elig_method_03, 1, 13, 43
EMReadScreen elig_method_04, 1, 13, 54
EMReadScreen elig_method_05, 1, 13, 65
EMReadScreen elig_method_06, 1, 13, 76
EMReadScreen elig_waiver_type_01, 1, 14, 21
EMReadScreen elig_waiver_type_02, 1, 14, 32
EMReadScreen elig_waiver_type_03, 1, 14, 43
EMReadScreen elig_waiver_type_04, 1, 14, 54
EMReadScreen elig_waiver_type_05, 1, 14, 65
EMReadScreen elig_waiver_type_06, 1, 14, 76
EMReadScreen elig_date_01, 5, 6, 19
EMReadScreen elig_date_02, 5, 6, 30
EMReadScreen elig_date_03, 5, 6, 41
EMReadScreen elig_date_04, 5, 6, 52
EMReadScreen elig_date_05, 5, 6, 63
EMReadScreen elig_date_06, 5, 6, 74

'COMBINES LIKE SECTIONS IN REVERSE
If elig_waiver_type_06 = elig_waiver_type_05 and elig_method_06 = elig_method_05 and elig_type_std_06 = elig_type_std_05 then
  elig_date_05 = elig_date_05 & "-" & right(elig_date_06, 5)
  elig_date_06 = ""
  elig_waiver_type_06 = ""
  elig_method_06 = ""
  elig_type_std_06 = ""
End if
If elig_waiver_type_05 = elig_waiver_type_04 and elig_method_05 = elig_method_04 and elig_type_std_05 = elig_type_std_04 then
  elig_date_04 = elig_date_04 & "-" & right(elig_date_05, 5)
  elig_date_05 = ""
  elig_waiver_type_05 = ""
  elig_method_05 = ""
  elig_type_std_05 = ""
End if
If elig_waiver_type_04 = elig_waiver_type_03 and elig_method_04 = elig_method_03 and elig_type_std_04 = elig_type_std_03 then
  elig_date_03 = elig_date_03 & "-" & right(elig_date_04, 5)
  elig_date_04 = ""
  elig_waiver_type_04 = ""
  elig_method_04 = ""
  elig_type_std_04 = ""
End if
If elig_waiver_type_03 = elig_waiver_type_02 and elig_method_03 = elig_method_02 and elig_type_std_03 = elig_type_std_02 then
  elig_date_02 = elig_date_02 & "-" & right(elig_date_03, 5)
  elig_date_03 = ""
  elig_waiver_type_03 = ""
  elig_method_03 = ""
  elig_type_std_03 = ""
End if
If elig_waiver_type_03 = elig_waiver_type_02 and elig_method_03 = elig_method_02 and elig_type_std_03 = elig_type_std_02 then
  elig_date_02 = elig_date_02 & "-" & right(elig_date_03, 5)
  elig_date_03 = ""
  elig_waiver_type_03 = ""
  elig_method_03 = ""
  elig_type_std_03 = ""
End if
If elig_waiver_type_02 = elig_waiver_type_01 and elig_method_02 = elig_method_01 and elig_type_std_02 = elig_type_std_01 then
  elig_date_01 = elig_date_01 & "-" & right(elig_date_02, 5)
  elig_date_02 = ""
  elig_waiver_type_02 = ""
  elig_method_02 = ""
  elig_type_std_02 = ""
End if

'REMOVING ANY UNUSED LINES IN THE DIALOG
For i = 1 to 6 'Does it several times to make sure the job gets done completely. This is easier than writing redundant sliding into each If...then statement from all other if...then statements.
  If elig_type_std_01 = "___" or elig_type_std_01 = "" then 
    elig_date_01 = elig_date_02
    elig_type_std_01 = elig_type_std_02
    elig_waiver_type_01 = elig_waiver_type_02 
    elig_method_01 = elig_method_02
    elig_date_02 = ""
    elig_waiver_type_02 = ""
    elig_method_02 = ""
    elig_type_std_02 = ""
  End if
  If elig_type_std_02 = "___" or elig_type_std_02 = "" then 
    elig_date_02 = elig_date_03
    elig_type_std_02 = elig_type_std_03
    elig_waiver_type_02 = elig_waiver_type_03 
    elig_method_02 = elig_method_03
    elig_date_03 = ""
    elig_waiver_type_03 = ""
    elig_method_03 = ""
    elig_type_std_03 = ""
  End if
  If elig_type_std_03 = "___" or elig_type_std_03 = "" then 
    elig_date_03 = elig_date_04
    elig_type_std_03 = elig_type_std_04
    elig_waiver_type_03 = elig_waiver_type_04 
    elig_method_03 = elig_method_04
    elig_date_04 = ""
    elig_waiver_type_04 = ""
    elig_method_04 = ""
    elig_type_std_04 = ""
  End if
  If elig_type_std_04 = "___" or elig_type_std_04 = "" then 
    elig_date_04 = elig_date_05
    elig_type_std_04 = elig_type_std_05
    elig_waiver_type_04 = elig_waiver_type_05 
    elig_method_04 = elig_method_05
    elig_date_05 = ""
    elig_waiver_type_05 = ""
    elig_method_05 = ""
    elig_type_std_05 = ""
  End if
  If elig_type_std_05 = "___" or elig_type_std_05 = "" then 
    elig_date_05 = elig_date_06
    elig_type_std_05 = elig_type_std_06
    elig_waiver_type_05 = elig_waiver_type_06 
    elig_method_05 = elig_method_06
    elig_date_06 = ""
    elig_waiver_type_06 = ""
    elig_method_06 = ""
    elig_type_std_06 = ""
  End if
Next

'DISPLAYS THE TYPE/STD DIALOG AFTER GATHERING THE INFO
Dialog type_std_dialog
If buttonpressed = 0 then stopscript

'READS THE FOOTER MONTH ELIG TYPE AND STANDARD
EMReadScreen elig_type, 2, 12, col - 1
EMReadScreen budget_type, 1, 13, col + 3

'NAVIGATES INTO THE BUDGET
EMWriteScreen "x", 9, col + 3
transmit

'CHECKS IF THE SCREEN HAS AN L BUDGET. IF IT DOES, IT READS THE INFO.
EMReadScreen LBUD_check, 4, 3, 45
If LBUD_check = "LBUD" then
  EMReadScreen recipient_amt, 10, 15, 70
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen LTC_exclusions, 10, 14, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 15, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 16, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Personal needs ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen home_maintenance_allowance, 10, 17, 32
  If home_maintenance_allowance <> "__________" then deductions = deductions & "Home maintenance allowance ($" & replace(home_maintenance_allowance, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 8, 70
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 9, 70
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 10, 70
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 11, 70
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 12, 70
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 13, 70
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

'CHECKS IF THE SCREEN HAS AN S BUDGET. IF IT DOES, IT READS THE INFO.
EMReadScreen SBUD_check, 4, 3, 44
If SBUD_check = "SBUD" then
  EMReadScreen recipient_amt, 10, 16, 71
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen LTC_exclusions, 10, 15, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 16, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 17, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Maintenance needs allowance ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 9, 71
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 10, 71
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 11, 71
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 12, 71
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 13, 71
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 14, 71
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

'CHECKS IF THE SCREEN HAS AN B BUDGET. IF IT DOES, IT ASKS THE WORKER WHAT TO DO.
BeginDialog BBUD_Dialog, 0, 0, 191, 76, "BBUD"
  Text 5, 10, 180, 10, "This is a method B budget. What would you like to do?"
  ButtonGroup ButtonPressed
    PushButton 20, 25, 70, 15, "Jump to STAT/BILS", BILS_button
    PushButton 100, 25, 70, 15, "Stay in ELIG/HC", ELIG_button
    CancelButton 135, 55, 50, 15
EndDialog

EMReadScreen BBUD_check, 4, 3, 47
If BBUD_check = "BBUD" then
  Dialog BBUD_dialog
  If ButtonPressed = 0 then stopscript
  If ButtonPressed = BILS_button then
    PF3
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" then
      Do
        Dialog BBUD_Dialog
        If buttonpressed = 0 then stopscript
      Loop until MAXIS_check = "MAXIS"
    End if
    back_to_SELF
    EMWriteScreen "stat", 16, 43
    EMWriteScreen "bils", 21, 70
    transmit
    EMReadScreen BILS_check, 4, 2, 54
    If BILS_check <> "BILS" then transmit
  End if
End if

'DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Do
	Do
		Do
			Do
				Dialog intake_approval_dialog
				cancel_confirmation
				'ensures that baseline date is in full date format so that the 'lookback period' is calculated correctly
				IF len(baseline_date) < 10 THEN MsgBox "You must enter the baseline date in format MM/DD/YYYY"
			LOOP until len(baseline_date) >= 10
			EMReadScreen STAT_check, 4, 20, 21
			If STAT_check = "STAT" then
			If ButtonPressed = prev_panel_button then call panel_navigation_prev
			If ButtonPressed = next_panel_button then call panel_navigation_next
			If ButtonPressed = prev_memb_button then call memb_navigation_prev
			If ButtonPressed = next_memb_button then call memb_navigation_next
			End if
			transmit 'Forces a screen refresh, to keep MAXIS from erroring out in the event of a password prompt.
			EMReadScreen MAXIS_check, 5, 1, 39
			If ButtonPressed = ELIG_HC_button then call navigate_to_MAXIS_screen("elig", "HC__")
			If ButtonPressed = AREP_button then call navigate_to_MAXIS_screen("stat", "AREP")
			If ButtonPressed = SWKR_button then call navigate_to_MAXIS_screen("stat", "SWKR")
			If ButtonPressed = FACI_button then call navigate_to_MAXIS_screen("stat", "FACI")
			If ButtonPressed = UNEA_button then call navigate_to_MAXIS_screen("stat", "UNEA")
			If ButtonPressed = MEDI_button then call navigate_to_MAXIS_screen("stat", "MEDI")
			If ButtonPressed = INSA_button then call navigate_to_MAXIS_screen("stat", "INSA")
			If ButtonPressed = BILS_button then call navigate_to_MAXIS_screen("stat", "BILS")
		Loop until ButtonPressed = -1 
		If actions_taken = "" THEN MsgBox "You need to complete the 'actions taken' field."
		If application_date = "" THEN MsgBox "You need to fill in the application date."
		IF worker_signature = "" then MsgBox "You need to sign your case note."
	Loop until actions_taken <> "" and application_date <> "" and worker_signature <> "" 
	CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.


'LOGIC
if month_MA_starts <> "" then
  header_date = month_MA_starts
Else
  header_date = month_MA_LTC_starts
End if

'Autofill for the application_date variable, then determines lookback period based on the info
'& "" is added to have variable listed in the dialog
If baseline_date <> "" then
	lookback_period = dateadd("m", -60, cdate(baseline_date)) & ""
	end_of_lookback = dateadd("d", -1, cdate(baseline_date)) & ""
End if

'This will write in 'no spenddown' into the case note if there is no amount the client is responsible to pay for MA purposes
If recipient_amt = "$" THEN recipient_amt = "no spenddown"

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE ("***MA effective " & header_date & "***")
If elig_type <> "DP" then 
  If budget_type = "L" then call write_bullet_and_variable_in_CASE_NOTE ("LTC spenddown", recipient_amt)	'these 3 are separate options as the 3 budget types have different wording for the client portion 
  If budget_type = "S" then call write_bullet_and_variable_in_CASE_NOTE ("SISEW waiver obligation", recipient_amt)
  If budget_type = "B" then call write_bullet_and_variable_in_CASE_NOTE ("Recipient amount", recipient_amt)
End if
If elig_date_01 <> "" then call write_variable_in_CASE_NOTE("* Elig type/std for " & elig_date_01 & ": " & replace(elig_type_std_01, "/_", "") & ", method " & elig_method_01 & replace(", waiver type " & elig_waiver_type_01, ", waiver type _", "") & ".")
If elig_date_02 <> "" then call write_variable_in_CASE_NOTE("* Elig type/std for " & elig_date_02 & ": " & replace(elig_type_std_02, "/_", "") & ", method " & elig_method_02 & replace(", waiver type " & elig_waiver_type_02, ", waiver type _", "") & ".")
If elig_date_03 <> "" then call write_variable_in_CASE_NOTE("* Elig type/std for " & elig_date_03 & ": " & replace(elig_type_std_03, "/_", "") & ", method " & elig_method_03 & replace(", waiver type " & elig_waiver_type_03, ", waiver type _", "") & ".")
If elig_date_04 <> "" then call write_variable_in_CASE_NOTE("* Elig type/std for " & elig_date_04 & ": " & replace(elig_type_std_04, "/_", "") & ", method " & elig_method_04 & replace(", waiver type " & elig_waiver_type_04, ", waiver type _", "") & ".")
If elig_date_05 <> "" then call write_variable_in_CASE_NOTE("* Elig type/std for " & elig_date_05 & ": " & replace(elig_type_std_05, "/_", "") & ", method " & elig_method_05 & replace(", waiver type " & elig_waiver_type_05, ", waiver type _", "") & ".")
If elig_date_06 <> "" then call write_variable_in_CASE_NOTE("* Elig type/std for " & elig_date_06 & ": " & replace(elig_type_std_06, "/_", "") & ", method " & elig_method_06 & replace(", waiver type " & elig_waiver_type_06, ", waiver type _", "") & ".")
call write_bullet_and_variable_in_CASE_NOTE("Application date", application_date)
If LTCC_check = 1 then call write_bullet_and_variable_in_CASE_NOTE("LTCC date", LTCC_date)
If DHS_5181_on_file_check = 1 then call write_variable_in_CASE_NOTE("* DHS-5181 on file.")
If DHS_1503_on_file_check = 1 then call write_variable_in_CASE_NOTE("* DHS-1503 on file.")
If retro_months <> "" then call write_bullet_and_variable_in_CASE_NOTE("Retro request", retro_months)
If month_MA_starts <> "" then call write_bullet_and_variable_in_CASE_NOTE("Month MA starts", month_MA_starts)
If month_MA_LTC_starts <> "" then call write_bullet_and_variable_in_CASE_NOTE("Month MA-LTC starts", month_MA_LTC_starts)
Call write_bullet_and_variable_in_case_note ("Baseline Date", baseline_date) 
Call write_bullet_and_variable_in_case_note ("Lookback period", lookback_period & "-" & end_of_lookback)
If community_check = 1 then call write_variable_in_CASE_NOTE("* Client is in the community.")
If AREP_SWKR <> "" then call write_bullet_and_variable_in_CASE_NOTE("AREP/SWKR", AREP_SWKR)
If FACI <> "" then call write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
If CFR <> "" then call write_bullet_and_variable_in_CASE_NOTE("CFR", CFR)
If income <> "" then call write_bullet_and_variable_in_CASE_NOTE("Income", income)
If assets <> "" then call write_bullet_and_variable_in_CASE_NOTE("Assets", assets)
If total_countable_assets <> "" then call write_bullet_and_variable_in_CASE_NOTE("Total countable assets", total_countable_assets)
If other_asset_notes <> "" then call write_bullet_and_variable_in_CASE_NOTE("Other asset notes", other_asset_notes)
If MEDI_INSA <> "" then call write_bullet_and_variable_in_CASE_NOTE("MEDI/INSA", MEDI_INSA)
If INSA_loaded_into_TPL_check = 1 then call write_variable_in_CASE_NOTE("* INSA loaded into TPL.")
If LTC_partnership_check = 1 then call write_variable_in_CASE_NOTE("* There is a LTC partnership for this case.")
If lookback_period <> "" then call write_variable_in_CASE_NOTE("* Lookback period: " & lookback_period & " - " & end_of_lookback)
If annuity_LTC_PRB <> "" then call write_bullet_and_variable_in_CASE_NOTE("Annuity (LTC) PRB", annuity_LTC_PRB)
If home_equity_limit <> "" then call write_bullet_and_variable_in_CASE_NOTE("Home equity limit", home_equity_limit)
If transfer <> "" then call write_bullet_and_variable_in_CASE_NOTE("Transfer", transfer)
If deductions <> "" then call write_bullet_and_variable_in_CASE_NOTE("BILS/deductions", deductions)
If other_notes <> "" then call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
If DHS_3050_1503_check = 1 then call write_variable_in_CASE_NOTE("* DHS-3050/1503 sent.")
If DHS_3203_lien_doc_check = 1 then call write_variable_in_CASE_NOTE("* DHS-3203/lien doc sent.")
If asset_transfer_letter_sent_check = 1 then call write_variable_in_CASE_NOTE("* Asset transfer letter sent.")
If managed_care_referral_sent_check = 1 then call write_variable_in_CASE_NOTE("* Managed care referral sent.")
call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
