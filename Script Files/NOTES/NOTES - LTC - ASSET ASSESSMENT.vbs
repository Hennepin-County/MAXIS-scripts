'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - ASSET ASSESSMENT.vbs"
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

'SPECIAL FUNCTIONS JUST FOR THIS SCRIPT
Function write_editbox_in_person_note(x, y) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
  variable_array = split(y, " ")
  EMSendKey "* " & x & ": "
  For each x in variable_array 
    EMGetCursor row, col 
    If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
      EMSendKey "<PF8>"
      EMWaitReady 0, 0
    End if
    EMReadScreen max_check, 51, 24, 2
    If max_check = "A MAXIMUM OF 4 PAGES ARE ALLOWED FOR EACH CASE NOTE" then exit for
    EMGetCursor row, col 
    If (row < 18 and col + (len(x)) >= 80) then EMSendKey "<newline>" & space(5)
    If (row = 5 and col = 3) then EMSendKey space(5)
    EMSendKey x & " "
    If right(x, 1) = ";" then 
      EMSendKey "<backspace>" & "<backspace>" 
      EMGetCursor row, col 
      If row = 18 then
        EMSendKey "<PF8>"
        EMWaitReady 0, 0
        EMSendKey space(5)
      Else
        EMSendKey "<newline>" & space(5)
      End if
    End if
  Next
  EMSendKey "<newline>"
  EMGetCursor row, col 
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

Function write_new_line_in_person_note(x)
  EMGetCursor row, col 
  If (row = 18 and col + (len(x)) >= 80 + 1 ) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
  EMGetCursor row, col 
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

'SECTION 02: DIALOGS

BeginDialog asset_assessment_dialog, 0, 0, 266, 261, "Asset assessment dialog"
  DropListBox 5, 5, 60, 15, "REQUIRED"+chr(9)+"REQUESTED", asset_assessment_type
  EditBox 195, 5, 65, 15, effective_date
  EditBox 165, 35, 65, 15, MA_LTC_first_month_of_documented_need
  EditBox 130, 55, 65, 15, month_MA_LTC_rules_applied
  EditBox 50, 80, 65, 15, LTC_spouse
  EditBox 195, 80, 65, 15, community_spouse
  EditBox 65, 100, 195, 15, asset_summary
  EditBox 80, 120, 60, 15, total_counted_assets
  EditBox 200, 120, 60, 15, half_of_total
  DropListBox 5, 145, 50, 15, "Actual"+chr(9)+"Estimated", CSAA_type
  EditBox 85, 140, 75, 15, CSAA
  EditBox 70, 160, 190, 15, asset_calculation
  EditBox 60, 180, 200, 15, actions_taken
  CheckBox 5, 205, 60, 10, "Sent 3340-A?", sent_3340A_check
  CheckBox 65, 205, 60, 10, "Sent 3340-B?", sent_3340B_check
  EditBox 195, 200, 65, 15, worker_signature
  CheckBox 5, 225, 175, 10, "Write MAXIS case note? If so, write case number:", write_MAXIS_case_note_check
  EditBox 185, 220, 75, 15, case_number
  ButtonGroup ButtonPressed
    OkButton 150, 240, 50, 15
    CancelButton 210, 240, 50, 15
  Text 70, 10, 75, 10, "ASSET ASSESSMENT"
  Text 160, 10, 30, 10, "Eff date:"
  GroupBox 20, 25, 220, 50, "If this is a required assessment, fill out:"
  Text 25, 40, 135, 10, "MA-LTC first month of documented need:"
  Text 25, 60, 100, 10, "Month MA-LTC rules applied:"
  Text 5, 85, 45, 10, "LTC spouse:"
  Text 125, 85, 70, 10, "Community spouse:"
  Text 5, 105, 55, 10, "Asset summary:"
  Text 5, 125, 75, 10, "Total Counted Assets:"
  Text 150, 125, 45, 10, "Half of Total:"
  Text 55, 145, 25, 10, "CSAA:"
  Text 5, 165, 65, 10, "Asset calculation:"
  Text 5, 185, 50, 10, "Actions taken:"
  Text 130, 205, 60, 10, "Worker signature:"
EndDialog


BeginDialog case_and_PMI_number_dialog, 0, 0, 196, 101, "Case and PMI number dialog"
  EditBox 80, 5, 70, 15, LTC_spouse_PMI
  EditBox 105, 25, 70, 15, community_spouse_PMI
  EditBox 100, 45, 70, 15, case_number
  CheckBox 5, 65, 190, 10, "Check here to enter ASET under the community spouse.", community_spouse_check
  ButtonGroup ButtonPressed
    OkButton 45, 80, 50, 15
    CancelButton 100, 80, 50, 15
  Text 15, 10, 60, 10, "LTC spouse PMI:"
  Text 15, 30, 85, 10, "Community spouse PMI:"
  Text 15, 50, 80, 10, "Case number (if known):"
EndDialog


'SECTION 03: THE SCRIPT----------------------------------------------------------------------------------------------------

Dialog case_and_PMI_number_dialog
If ButtonPressed = 0 then stopscript

'Connecting and checking for an active MAXIS section
EMConnect ""
Call check_for_MAXIS(True)
back_to_self
call navigate_to_screen("aset", "____")

EMWriteScreen "________", 13, 62
If community_spouse_check = 1 then 
  EMWriteScreen community_spouse_PMI, 13, 62
Else
  EMWriteScreen LTC_spouse_PMI, 13, 62
End if
EMWriteScreen "faco", 20, 71
transmit
EMReadScreen effective_date, 8, 7, 72
effective_date = cdate(effective_date) & ""
EMWriteScreen "x", 7, 3
transmit
EMReadScreen LTC_spouse, 13, 7, 63
LTC_spouse = trim(LTC_spouse)
LTC_spouse = left(LTC_spouse, 1) & lcase(right(LTC_spouse, len(LTC_spouse) - 1))
EMReadScreen community_spouse, 13, 15, 63
community_spouse = trim(community_spouse)
community_spouse = left(community_spouse, 1) & lcase(right(community_spouse, len(community_spouse) - 1))
EMWriteScreen "SPAA", 20, 71
transmit
EMReadScreen total_counted_assets, 10, 6, 66
total_counted_assets = trim(total_counted_assets)
total_counted_assets = replace(total_counted_assets, ",", "")
half_of_total = "$" & round(ccur(total_counted_assets)/2, 2) 
total_counted_assets = "$" & total_counted_assets
EMReadScreen CSAA, 10, 8, 66
CSAA = trim(CSAA)
CSAA = replace(CSAA, ",", "")
CSAA = "$" & round(ccur(CSAA), 2) 

'Now it's going to read the entire SPAA screen, to enter it into a case note. Skips the fourth, sixth, and twelfth line as they're blank!
EMReadScreen SPAA_line_01, 55, 4, 24
If trim(SPAA_line_01) = "" then SPAA_line_01 = "."
EMReadScreen SPAA_line_02, 55, 5, 24
If trim(SPAA_line_02) = "" then SPAA_line_02 = "."
EMReadScreen SPAA_line_03, 55, 6, 24
If trim(SPAA_line_03) = "" then SPAA_line_03 = "."
EMReadScreen SPAA_line_05, 55, 8, 24
If trim(SPAA_line_05) = "" then SPAA_line_05 = "."
EMReadScreen SPAA_line_07, 55, 10, 24
If trim(SPAA_line_07) = "" then SPAA_line_07 = "."
EMReadScreen SPAA_line_08, 55, 11, 24
If trim(SPAA_line_08) = "" then SPAA_line_08 = "."
EMReadScreen SPAA_line_09, 55, 12, 24
If trim(SPAA_line_09) = "" then SPAA_line_09 = "."
EMReadScreen SPAA_line_10, 55, 13, 24
If trim(SPAA_line_10) = "" then SPAA_line_10 = "."
EMReadScreen SPAA_line_11, 55, 14, 24
If trim(SPAA_line_11) = "" then SPAA_line_11 = "."
EMReadScreen SPAA_line_13, 55, 16, 24
If trim(SPAA_line_13) = "" then SPAA_line_13 = "."
EMReadScreen SPAA_line_14, 55, 17, 24
If trim(SPAA_line_14) = "" then SPAA_line_14 = "."
EMReadScreen SPAA_line_15, 55, 18, 24
If trim(SPAA_line_15) = "" then SPAA_line_15 = "."

'Now it's going to get the marital asset list. Skips lines 2 and 16 as they are blank.
EMWriteScreen "x", 4, 33 
transmit
'these lines are not included in the DO LOOP since they are headers and footers
EMReadScreen total_marital_asset_list_line_01, 53, 2, 25	'TOTAL MARITAL ASSET LIST (header)
EMReadScreen total_marital_asset_list_line_03, 53, 4, 25	'Asset Description (header)
EMReadScreen total_marital_asset_list_line_99, 53, 18, 25	'Assets Total: (footer) --made this '99' as to not cause conflict with other variable titles
EMReadScreen total_marital_asset_list_line_04, 53, 5, 25	'-------------------- (header)
'1st page of the total marital asset list
EMReadScreen total_marital_asset_list_line_10, 53, 6, 25
EMReadScreen total_marital_asset_list_line_11, 53, 7, 25
EMReadScreen total_marital_asset_list_line_12, 53, 8, 25
EMReadScreen total_marital_asset_list_line_13, 53, 9, 25
EMReadScreen total_marital_asset_list_line_14, 53, 10, 25
EMReadScreen total_marital_asset_list_line_15, 53, 11, 25
EMReadScreen total_marital_asset_list_line_16, 53, 12, 25
EMReadScreen total_marital_asset_list_line_17, 53, 13, 25
EMReadScreen total_marital_asset_list_line_18, 53, 14, 25
EMReadScreen total_marital_asset_list_line_19, 53, 15, 25
PF8
EMReadScreen last_SPAA_page_check, 25, 23, 4		'checking to make sure that no more assets need to be copied for the case note
If last_SPAA_page_check = "NO MORE ASSETS TO DISPLAY" THEN 
	PF3
ELSE
	'2nd page of the total marital asset list
	EMReadScreen total_marital_asset_list_line_20, 53, 6, 25
	EMReadScreen total_marital_asset_list_line_21, 53, 7, 25
	EMReadScreen total_marital_asset_list_line_22, 53, 8, 25
	EMReadScreen total_marital_asset_list_line_23, 53, 9, 25
	EMReadScreen total_marital_asset_list_line_24, 53, 10, 25
	EMReadScreen total_marital_asset_list_line_25, 53, 11, 25
	EMReadScreen total_marital_asset_list_line_26, 53, 12, 25
	EMReadScreen total_marital_asset_list_line_27, 53, 13, 25
	EMReadScreen total_marital_asset_list_line_28, 53, 14, 25
	EMReadScreen total_marital_asset_list_line_29, 53, 15, 25
	PF8
	EMReadScreen last_SPAA_page_check, 25, 23, 4		'checking to make sure that no more assets need to be copied for the case note
	If last_SPAA_page_check = "NO MORE ASSETS TO DISPLAY" THEN 
		PF3
	ELSE
		'3rd page of the total marital asset list
		EMReadScreen total_marital_asset_list_line_30, 53, 6, 25
		EMReadScreen total_marital_asset_list_line_31, 53, 7, 25
		EMReadScreen total_marital_asset_list_line_32, 53, 8, 25
		EMReadScreen total_marital_asset_list_line_33, 53, 9, 25
		EMReadScreen total_marital_asset_list_line_34, 53, 10, 25
		EMReadScreen total_marital_asset_list_line_35, 53, 11, 25
		EMReadScreen total_marital_asset_list_line_36, 53, 12, 25
		EMReadScreen total_marital_asset_list_line_37, 53, 13, 25
		EMReadScreen total_marital_asset_list_line_38, 53, 14, 25
		EMReadScreen total_marital_asset_list_line_39, 53, 15, 25
		PF8
		EMReadScreen last_SPAA_page_check, 25, 23, 4		'checking to make sure that no more assets need to be copied for the case note
		If last_SPAA_page_check = "NO MORE ASSETS TO DISPLAY" THEN 
			PF3
		ELSE
			'4th page of the total marital asset list
			EMReadScreen total_marital_asset_list_line_40, 53, 6, 25
			EMReadScreen total_marital_asset_list_line_41, 53, 7, 25
			EMReadScreen total_marital_asset_list_line_42, 53, 8, 25
			EMReadScreen total_marital_asset_list_line_43, 53, 9, 25
			EMReadScreen total_marital_asset_list_line_44, 53, 10, 25
			EMReadScreen total_marital_asset_list_line_45, 53, 11, 25
			EMReadScreen total_marital_asset_list_line_46, 53, 12, 25
			EMReadScreen total_marital_asset_list_line_47, 53, 13, 25
			EMReadScreen total_marital_asset_list_line_48, 53, 14, 25
			EMReadScreen total_marital_asset_list_line_49, 53, 15, 25
			PF3
			PF3
		END IF
	END IF
END IF

Do
	dialog asset_assessment_dialog
	If buttonpressed = 0 then stopscript
	cancel_confirmation
	transmit
	EMReadScreen function_check, 4, 20, 21
	If function_check <> "ASET" then MsgBox "You do not appear to be in the ASET function anymore. You might be locked out of your case, or have nagigated away. Reenter the ASET function before proceeding."
 Loop until function_check = "ASET"
 'navigates to person note
PF5 
'puts person note into edit mode
PF9

'case/person notes information about forms sent to client
If sent_3340B_check = 1 then actions_taken = "Sent 3340-B. " & actions_taken
If sent_3340A_check = 1 then actions_taken = "Sent 3340-A. " & actions_taken

'PERSON NOTE----------------------------------------------------------------------------------------------------
EMSendKey "***" & asset_assessment_type & " ASSET ASSESSMENT***" & "<newline>"
call write_editbox_in_person_note("Effective date", effective_date) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
If MA_LTC_first_month_of_documented_need <> "" then call write_editbox_in_person_note("MA-LTC first month of documented need", MA_LTC_first_month_of_documented_need)
If month_MA_LTC_rules_applied <> "" then call write_editbox_in_person_note("Month MA-LTC rules applied", month_MA_LTC_rules_applied)
If LTC_spouse <> "" then call write_editbox_in_person_note("LTC spouse", LTC_spouse)
If community_spouse <> "" then call write_editbox_in_person_note("Community spouse", community_spouse)
If asset_summary <> "" then call write_editbox_in_person_note("Asset summary", asset_summary)
If total_counted_assets <> "" then call write_editbox_in_person_note("Total counted assets", total_counted_assets)
If half_of_total <> "" then call write_editbox_in_person_note("Half of total", half_of_total)
If CSAA_type <> "" then call write_new_line_in_person_note("* " & CSAA_type & " CSAA: " & CSAA)
If asset_calculation <> "" then call write_editbox_in_person_note("Asset calculation", asset_calculation)
If actions_taken <> "" then call write_editbox_in_person_note("Actions taken", actions_taken)
call write_new_line_in_person_note("---")
If worker_signature <> "" then call write_new_line_in_person_note(worker_signature)
Do
  EMGetCursor row, col
  If row < 18 then 
    EMSendKey "."
    EMSendKey "<newline>"
  End if
Loop until row = 18
EMSendKey ">>>>SPAA PASTED ON NEXT PAGE>>>>"
PF8
call write_new_line_in_person_note(SPAA_line_01)
call write_new_line_in_person_note(SPAA_line_02)
call write_new_line_in_person_note(SPAA_line_03)
call write_new_line_in_person_note(SPAA_line_05)
call write_new_line_in_person_note(SPAA_line_07)
call write_new_line_in_person_note(SPAA_line_08)
call write_new_line_in_person_note(SPAA_line_09)
call write_new_line_in_person_note(SPAA_line_10)
call write_new_line_in_person_note(SPAA_line_11)
call write_new_line_in_person_note(SPAA_line_13)
call write_new_line_in_person_note(SPAA_line_14)
call write_new_line_in_person_note(SPAA_line_15)
Do
  EMGetCursor row, col
  If row < 18 then 
    EMSendKey "."
    EMSendKey "<newline>"
  End if
Loop until row = 18
EMSendKey ">>>>TOTAL MARITAL ASSET LIST PASTED ON NEXT PAGE>>>>"
PF8
'headers
call write_new_line_in_person_note(total_marital_asset_list_line_99)
call write_new_line_in_person_note(total_marital_asset_list_line_03)
call write_new_line_in_person_note(total_marital_asset_list_line_04)
'1st page of total marital assets
If total_marital_asset_list_line_10 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note (total_marital_asset_list_line_10)
If total_marital_asset_list_line_11 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_11)
If total_marital_asset_list_line_12 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_12)
If total_marital_asset_list_line_13 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_13)
If total_marital_asset_list_line_14 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_14)
If total_marital_asset_list_line_15 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_15)
If total_marital_asset_list_line_16 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_16)
If total_marital_asset_list_line_17 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_17)
If total_marital_asset_list_line_18 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_18)
If total_marital_asset_list_line_19 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_19)
'2nd page of the total marital asset list
If total_marital_asset_list_line_20 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_20)
If total_marital_asset_list_line_21 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_21)
If total_marital_asset_list_line_22 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_22)
If total_marital_asset_list_line_23 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_23)
If total_marital_asset_list_line_24 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_24)
If total_marital_asset_list_line_25 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_25)
If total_marital_asset_list_line_26 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_26)
If total_marital_asset_list_line_27 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_27)
If total_marital_asset_list_line_28 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_28)
If total_marital_asset_list_line_29 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_29)
'3rd page of total marital asset list
If total_marital_asset_list_line_30 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_30)
If total_marital_asset_list_line_31 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_31)
If total_marital_asset_list_line_32 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_32)
If total_marital_asset_list_line_33 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_33)
If total_marital_asset_list_line_34 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_34)
If total_marital_asset_list_line_35 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_35)
If total_marital_asset_list_line_36 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_36)
If total_marital_asset_list_line_37 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_37)
If total_marital_asset_list_line_38 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_38)
If total_marital_asset_list_line_39 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_39)
'4th page of total marital asset list
If total_marital_asset_list_line_40 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_40)
If total_marital_asset_list_line_41 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_41)
If total_marital_asset_list_line_42 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_42)
If total_marital_asset_list_line_43 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_43)
If total_marital_asset_list_line_44 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_44)
If total_marital_asset_list_line_45 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_45)
If total_marital_asset_list_line_46 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_46)
If total_marital_asset_list_line_47 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_47)
If total_marital_asset_list_line_48 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_48)
If total_marital_asset_list_line_49 <> "* ____________________  _________   _       _      _ " then Call write_new_line_in_person_note(total_marital_asset_list_line_49)
PF3
PF3
'End of person note----------------------------------------------------------------------------------------------------

If write_MAXIS_case_note_check = 0 then script_end_procedure("")

'CASE NOTE----------------------------------------------------------------------------------------------------
call navigate_to_screen("case", "note")
Call start_a_blank_case_note

Call write_variable_in_CASE_NOTE ("***" & asset_assessment_type & " ASSET ASSESSMENT***")
call write_bullet_and_variable_in_CASE_NOTE("Effective date", effective_date) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
If MA_LTC_first_month_of_documented_need <> "" then call write_bullet_and_variable_in_CASE_NOTE("MA-LTC first month of documented need", MA_LTC_first_month_of_documented_need)
If month_MA_LTC_rules_applied <> "" then call write_bullet_and_variable_in_CASE_NOTE("Month MA-LTC rules applied", month_MA_LTC_rules_applied)
call write_bullet_and_variable_in_CASE_NOTE("LTC spouse", LTC_spouse)
call write_bullet_and_variable_in_CASE_NOTE("Community spouse", community_spouse)
call write_bullet_and_variable_in_CASE_NOTE("Asset summary", asset_summary)
call write_bullet_and_variable_in_CASE_NOTE("Total counted assets", total_counted_assets)
call write_bullet_and_variable_in_CASE_NOTE("Half of total", half_of_total)
call write_bullet_and_variable_in_CASE_NOTE(CSAA_type & " CSAA", CSAA)
call write_bullet_and_variable_in_CASE_NOTE("Asset calculation", asset_calculation)
call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
call write_variable_in_CASE_NOTE("---")
If worker_signature <> "" then call write_new_line_in_case_note(worker_signature)
Do
  EMGetCursor row, col
  If row < 17 then 
    EMSendKey "."
    EMSendKey "<newline>"
  End if
Loop until row = 17
EMSendKey ">>>>SPAA PASTED ON NEXT PAGE>>>>"
PF8
call write_variable_in_CASE_NOTE(SPAA_line_01)
call write_variable_in_CASE_NOTE(SPAA_line_02)
call write_variable_in_CASE_NOTE(SPAA_line_03)
call write_variable_in_CASE_NOTE(SPAA_line_05)
call write_variable_in_CASE_NOTE(SPAA_line_07)
call write_variable_in_CASE_NOTE(SPAA_line_08)
call write_variable_in_CASE_NOTE(SPAA_line_09)
call write_variable_in_CASE_NOTE(SPAA_line_10)
call write_variable_in_CASE_NOTE(SPAA_line_11)
call write_variable_in_CASE_NOTE(SPAA_line_13)
call write_variable_in_CASE_NOTE(SPAA_line_14)
call write_variable_in_CASE_NOTE(SPAA_line_15)

Do
  EMGetCursor row, col
  If row < 17 then 
    EMSendKey "."
    EMSendKey "<newline>"
  End if
Loop until row = 17

EMSendKey ">>>>TOTAL MARITAL ASSET LIST PASTED ON NEXT PAGE>>>>"
PF8
'headers and footer
call write_variable_in_CASE_NOTE(total_marital_asset_list_line_99)
call write_variable_in_CASE_NOTE(total_marital_asset_list_line_03)
call write_variable_in_CASE_NOTE(total_marital_asset_list_line_04)
'1st page of total marital assets
If total_marital_asset_list_line_10 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE (total_marital_asset_list_line_10)
If total_marital_asset_list_line_11 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_11)
If total_marital_asset_list_line_12 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_12)
If total_marital_asset_list_line_13 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_13)
If total_marital_asset_list_line_14 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_14)
If total_marital_asset_list_line_15 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_15)
If total_marital_asset_list_line_16 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_16)
If total_marital_asset_list_line_17 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_17)
If total_marital_asset_list_line_18 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_18)
If total_marital_asset_list_line_19 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_19)
'2nd page of the total marital asset list
If total_marital_asset_list_line_20 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_20)
If total_marital_asset_list_line_21 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_21)
If total_marital_asset_list_line_22 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_22)
If total_marital_asset_list_line_23 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_23)
If total_marital_asset_list_line_24 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_24)
If total_marital_asset_list_line_25 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_25)
If total_marital_asset_list_line_26 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_26)
If total_marital_asset_list_line_27 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_27)
If total_marital_asset_list_line_28 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_28)
If total_marital_asset_list_line_29 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_29)
'3rd page of total marital asset list
If total_marital_asset_list_line_30 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_30)
If total_marital_asset_list_line_31 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_31)
If total_marital_asset_list_line_32 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_32)
If total_marital_asset_list_line_33 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_33)
If total_marital_asset_list_line_34 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_34)
If total_marital_asset_list_line_35 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_35)
If total_marital_asset_list_line_36 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_36)
If total_marital_asset_list_line_37 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_37)
If total_marital_asset_list_line_38 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_38)
If total_marital_asset_list_line_39 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_39)
'4th page of total marital asset list
If total_marital_asset_list_line_40 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_40)
If total_marital_asset_list_line_41 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_41)
If total_marital_asset_list_line_42 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_42)
If total_marital_asset_list_line_43 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_43)
If total_marital_asset_list_line_44 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_44)
If total_marital_asset_list_line_45 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_45)
If total_marital_asset_list_line_46 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_46)
If total_marital_asset_list_line_47 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_47)
If total_marital_asset_list_line_48 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_48)
If total_marital_asset_list_line_49 <> "* ____________________  _________   _       _      _ " then Call write_variable_in_CASE_NOTE(total_marital_asset_list_line_49)

script_end_procedure("")