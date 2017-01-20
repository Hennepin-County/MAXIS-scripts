'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - ASSET ASSESSMENT.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 960           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("01/19/2017", "Added 'New AA received' option, and updated functionality of script to include enhanced password handling, and handling for all 8 pages of total marital assets.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SPECIAL FUNCTIONS JUST FOR THIS SCRIPT----------------------------------------------------------------------------------------------------
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog asset_assessment_dialog, 0, 0, 266, 270, "Asset assessment completed dialog"
  DropListBox 5, 10, 60, 15, "REQUIRED"+chr(9)+"REQUESTED", asset_assessment_type
  EditBox 195, 10, 65, 15, effective_date
  EditBox 165, 40, 65, 15, MA_LTC_first_month_of_documented_need
  EditBox 165, 60, 65, 15, month_MA_LTC_rules_applied
  EditBox 50, 85, 65, 15, LTC_spouse
  EditBox 195, 85, 65, 15, community_spouse
  EditBox 65, 105, 195, 15, asset_summary
  EditBox 80, 125, 60, 15, total_counted_assets
  EditBox 200, 125, 60, 15, half_of_total
  DropListBox 80, 145, 60, 15, "Actual"+chr(9)+"Estimated", CSAA_type
  EditBox 200, 145, 60, 15, CSAA
  EditBox 70, 165, 190, 15, asset_calculation
  EditBox 60, 185, 200, 15, actions_taken
  CheckBox 5, 210, 60, 10, "Sent 3340-A?", sent_3340A_check
  CheckBox 80, 210, 60, 10, "Sent 3340-B?", sent_3340B_check
  CheckBox 145, 210, 110, 10, "Sent 5181 to Case Manager?", sent_5181_check
  CheckBox 5, 230, 175, 10, "Write MAXIS case note? If so, write case number:", write_MAXIS_case_note_check
  EditBox 185, 225, 75, 15, MAXIS_case_number
  EditBox 65, 245, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 245, 50, 15
    CancelButton 210, 245, 50, 15
  Text 70, 15, 75, 10, "ASSET ASSESSMENT"
  GroupBox 20, 30, 220, 50, "If this is a required assessment, fill out:"
  Text 30, 45, 135, 10, "MA-LTC first month of documented need:"
  Text 70, 65, 95, 10, "Month MA-LTC rules applied:"
  Text 5, 90, 45, 10, "LTC spouse:"
  Text 125, 90, 70, 10, "Community spouse:"
  Text 5, 110, 55, 10, "Asset summary:"
  Text 5, 130, 75, 10, "Total Counted Assets:"
  Text 150, 130, 45, 10, "Half of Total:"
  Text 5, 170, 65, 10, "Asset calculation:"
  Text 5, 190, 50, 10, "Actions taken:"
  Text 5, 250, 60, 10, "Worker signature:"
  Text 160, 150, 25, 10, "CSAA:"
  Text 155, 15, 40, 10, "AA Eff date:"
EndDialog

BeginDialog AA_received_dialog, 0, 0, 266, 235, "AA received dialog"
  DropListBox 5, 10, 60, 15, "REQUIRED"+chr(9)+"REQUESTED", asset_assessment_type
  EditBox 195, 10, 65, 15, effective_date
  EditBox 165, 40, 65, 15, MA_LTC_first_month_of_documented_need
  EditBox 165, 60, 65, 15, month_MA_LTC_rules_applied
  EditBox 50, 85, 70, 15, LTC_spouse
  EditBox 195, 85, 65, 15, community_spouse
  EditBox 80, 110, 180, 15, verifs_requested
  EditBox 80, 130, 180, 15, actions_taken
  CheckBox 80, 150, 145, 10, "Verification request sent to client/spouse", client_verif_checkbox
  CheckBox 80, 165, 120, 10, "Verification request sent to AREP", AREP_verif_checkbox
  CheckBox 15, 185, 170, 10, "Write MAXIS case note? If so, write case number:", write_MAXIS_case_note_check
  EditBox 185, 180, 75, 15, MAXIS_case_number
  EditBox 80, 205, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 205, 50, 15
    CancelButton 210, 205, 50, 15
  Text 30, 45, 135, 10, "MA-LTC first month of documented need:"
  Text 65, 65, 95, 10, "Month MA-LTC rules applied:"
  Text 5, 90, 45, 10, "LTC spouse:"
  Text 125, 90, 70, 10, "Community spouse:"
  Text 10, 115, 70, 10, "Verifications needed: "
  Text 30, 135, 50, 10, "Actions taken:"
  Text 20, 210, 60, 10, "Worker signature:"
  Text 155, 15, 40, 10, "AA Eff date:"
  GroupBox 5, 30, 255, 50, "If this is a required assessment, fill out:"
  Text 70, 15, 75, 10, "ASSET ASSESSMENT"
EndDialog

BeginDialog case_and_PMI_number_dialog, 0, 0, 201, 120, "Case and PMI number dialog"
  EditBox 105, 5, 60, 15, LTC_spouse_PMI
  EditBox 105, 25, 60, 15, community_spouse_PMI
  EditBox 105, 45, 60, 15, MAXIS_case_number
  DropListBox 105, 65, 75, 15, "Select one..."+chr(9)+"Completed AA"+chr(9)+"New AA received", AA_option
  CheckBox 5, 85, 190, 10, "Check here to enter ASET under the community spouse.", community_spouse_check
  ButtonGroup ButtonPressed
    OkButton 45, 100, 50, 15
    CancelButton 100, 100, 50, 15
  Text 25, 30, 80, 10, "Community spouse PMI:"
  Text 25, 50, 80, 10, "Case number (if known):"
  Text 25, 70, 80, 10, "Choose an AA option:"
  Text 45, 10, 60, 10, "LTC spouse PMI:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting and checking for an active MAXIS section
EMConnect ""

'initial dialog for case number and PMI
Do 
	Do 
		err_msg = ""
		Dialog case_and_PMI_number_dialog
		If ButtonPressed = 0 then stopscript
		If IsNumeric(LTC_spouse_PMI) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid LTC spouse PMI." 
		If IsNumeric(community_spouse_PMI) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid Community spouse PMI." 
		If AA_option = "Select one..." THEN err_msg = err_msg & vbNewLine & "* Select an Asset Assessment option." 
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
	Loop until err_msg = ""	
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

Call check_for_MAXIS(False)
call navigate_to_MAXIS_screen("aset", "____")

'Writing in the PMI for spouse depending on if community spouse checkbox was checked or not
EMWriteScreen "________", 13, 62
If community_spouse_check = 1 then
  EMWriteScreen community_spouse_PMI, 13, 62
Else
  EMWriteScreen LTC_spouse_PMI, 13, 62
End if

'Gathering information and creating variables
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

IF AA_option = "New AA received" then 
    Do 
    	Do	
			Do 
				err_msg = ""
    			dialog AA_received_dialog	'calls the main asset assessment dialog
    			cancel_confirmation
				If effective_date = "" THEN err_msg = err_msg & vbNewLine & "* Enter the AA effective date."
				If (asset_assessment_type = "REQUIRED" AND MA_LTC_first_month_of_documented_need = "") THEN err_msg = err_msg & vbNewLine & "* Enter the first month of documented need of MA-LTC."
				If (asset_assessment_type = "REQUIRED" AND month_MA_LTC_rules_applied = "") THEN err_msg = err_msg & vbNewLine & "* Enter the first month MA-LTC rules applied."
				If verifs_requested = "" THEN err_msg = err_msg & vbNewLine & "* Enter the verifications requested."
				If actions_taken = "" THEN err_msg = err_msg & vbNewLine & "* Enter the actions taken."
				If (write_MAXIS_case_note_check = 1 and MAXIS_case_number = "") THEN err_msg = err_msg & vbNewLine & "* Enter the MAXIS case number or uncheck the option to write the case note." 
				IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
			Loop until err_msg = ""	
			transmit
    		EMReadScreen function_check, 4, 20, 21		'checking to make sure that we're still in ASET function
    		If function_check <> "ASET" then
    			MsgBox "You do not appear to be in the ASET function any more. You might be locked out of your case, or have navigated away. Re-enter the ASET function before proceeding."
    		END IF
    	Loop until function_check = "ASET"
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
    Loop until are_we_passworded_out = false					'loops until user passwords back in		

	PF5 	'navigates to person note
	PF9		'puts person note into edit mode
	
	EMSendKey "***" & asset_assessment_type & " ASSET ASSESSMENT RECEIVED***" & "<newline>"
	call write_editbox_in_person_note("Effective date", effective_date) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
	call write_editbox_in_person_note("MA-LTC first month of documented need", MA_LTC_first_month_of_documented_need)
	call write_editbox_in_person_note("Month MA-LTC rules applied", month_MA_LTC_rules_applied)
	call write_editbox_in_person_note("LTC spouse", LTC_spouse)
	call write_editbox_in_person_note("Community spouse", community_spouse)
	call write_editbox_in_person_note("Verifications requested", verifs_requested)
	call write_editbox_in_person_note("Actions taken", actions_taken)
	If client_verif_checkbox = 1 then call write_new_line_in_person_note("* Verification request sent to client/spouse.")
	If AREP_verif_checkbox = 1 then call write_new_line_in_person_note("* Verification request sent to AREP.")
	call write_new_line_in_person_note("---")
	call write_new_line_in_person_note(worker_signature)
	PF3
	PF3
ELse 

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
    
    'Gathering information from the total maritial asset list and adding to an array
    MAXIS_row = 6
    asset_list = ""
    Do 
    	EMReadScreen asset_check, 20, MAXIS_row, 27				'chekcing to make sure the asset line is not an underscore line
    	If asset_check <> "____________________" then 			
    		EMReadScreen listed_asset, 53, MAXIS_row, 25		'reads the assets
    		listed_asset = replace(listed_asset, "_", " ")		'relaces the underscores 	
    		MAXIS_row = MAXIS_row + 1
    		asset_list = asset_list & listed_asset & ", "		'increments the asset_list variable by the listed_asset variable
    		EMReadScreen last_page_check, 7, 23, 4		'checking to make sure that no more assets need to be copied for the case note
    		If last_page_check = "NO MORE" then exit do
    		If MAXIS_row = 16 then 
    			PF8
    			MAXIS_row = 6							'accounts for more than one page
    		END If 
    	END IF 
    LOOP UNTIL asset_check = "____________________"		'loops until all the assets are accounted for 
    PF3			'goes back into 		
    
    'declaring & splitting the array
    If left(asset_list, 2) = ", " then asset_list = right(asset_list, len(asset_list) - 2)
    assets_array = Split(asset_list, ",")
    		
    If AA_option = "Completed AA" then 
    	Do 
    		Do			
    			dialog asset_assessment_dialog	'calls the main asset assessment dialog
    			cancel_confirmation
    			transmit
    			EMReadScreen function_check, 4, 20, 21		'checking to make sure that we're still in ASET function
    			If function_check <> "ASET" then
    				MsgBox "You do not appear to be in the ASET function any more. You might be locked out of your case, or have navigated away. Re-enter the ASET function before proceeding."
    			END IF
    		Loop until function_check = "ASET"
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
    	Loop until are_we_passworded_out = false					'loops until user passwords back in		
    END IF 
    
    PF5 	'navigates to person note
    PF9		'puts person note into edit mode
    'case/person notes information about forms sent to client
    If sent_3340B_check = 1 then actions_taken = "Sent 3340-B. " & actions_taken
    If sent_3340A_check = 1 then actions_taken = "Sent 3340-A. " & actions_taken
    
    'PERSON NOTE----------------------------------------------------------------------------------------------------
    EMSendKey "***" & asset_assessment_type & " ASSET ASSESSMENT COMPELTED***" & "<newline>"
    call write_editbox_in_person_note("Effective date", effective_date) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
    call write_editbox_in_person_note("MA-LTC first month of documented need", MA_LTC_first_month_of_documented_need)
    call write_editbox_in_person_note("Month MA-LTC rules applied", month_MA_LTC_rules_applied)
    call write_editbox_in_person_note("LTC spouse", LTC_spouse)
    call write_editbox_in_person_note("Community spouse", community_spouse)
    If trim(asset_summary) <> "" then call write_editbox_in_person_note("Asset summary", asset_summary)
    If trim(total_counted_assets) <> "" then call write_editbox_in_person_note("Total counted assets", total_counted_assets)
    If trim(half_of_total) <> "" then call write_editbox_in_person_note("Half of total", half_of_total)
    If trim(CSAA) <> "" then call write_new_line_in_person_note("* " & CSAA_type & " CSAA: " & CSAA)
    If trim(asset_calculation) <> "" then call write_editbox_in_person_note("Asset calculation", asset_calculation)
    call write_editbox_in_person_note("Actions taken", actions_taken)
    If client_verif_checkbox = 1 then call write_new_line_in_person_note("* Verification request sent to client/spouse.")
    If AREP_verif_checkbox = 1 then call write_new_line_in_person_note("* Verification request sent to AREP.")
    If sent_5181_check = 1 then call write_new_line_in_person_note("* DHS-5181 sent to Case Manager.")
    call write_new_line_in_person_note("---")
    call write_new_line_in_person_note(worker_signature)

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
    
    'Person notes the assets in the assets array
    For each asset in assets_array
    	Call write_new_line_in_person_note(asset)
    Next 
    
    PF3
    PF3
END IF 
'End of person note----------------------------------------------------------------------------------------------------

If write_MAXIS_case_note_check = 0 then script_end_procedure("")

'CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_case_note
If AA_option = "Completed AA" then
	Call write_variable_in_CASE_NOTE ("***" & asset_assessment_type & " ASSET ASSESSMENT COMPLETED***")
Else 
	Call write_variable_in_CASE_NOTE ("***" & asset_assessment_type & " ASSET ASSESSMENT RECEIVED***")
END IF 
call write_bullet_and_variable_in_CASE_NOTE("Effective date", effective_date) 'x is the header, y is the variable for the edit box which will be put in the case note, z is the length of spaces for the indent.
call write_bullet_and_variable_in_CASE_NOTE("MA-LTC first month of documented need", MA_LTC_first_month_of_documented_need)
call write_bullet_and_variable_in_CASE_NOTE("Month MA-LTC rules applied", month_MA_LTC_rules_applied)
call write_bullet_and_variable_in_CASE_NOTE("LTC spouse", LTC_spouse)
call write_bullet_and_variable_in_CASE_NOTE("Community spouse", community_spouse)
Call write_bullet_and_variable_in_CASE_NOTE("Verifications requested", verifs_requested)
call write_bullet_and_variable_in_CASE_NOTE("Asset summary", asset_summary)
call write_bullet_and_variable_in_CASE_NOTE("Total counted assets", total_counted_assets)
call write_bullet_and_variable_in_CASE_NOTE("Half of total", half_of_total)
call write_bullet_and_variable_in_CASE_NOTE(CSAA_type & " CSAA", CSAA)
call write_bullet_and_variable_in_CASE_NOTE("Asset calculation", asset_calculation)
call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
If sent_5181_check = 1 then call write_variable_in_CASE_NOTE("* DHS-5181 sent to Case Manager.")
If client_verif_checkbox = 1 then call write_variable_in_CASE_NOTE("* Verification request sent to client/spouse.")
If AREP_verif_checkbox = 1 then call write_variable_in_CASE_NOTE("* Verification request sent to AREP.")
call write_variable_in_CASE_NOTE("---")
Call write_variable_in_case_note(worker_signature)

If AA_option = "Completed AA" then 
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
    
    'Case notes the assets in the assets array
    For each asset in assets_array
    	Call write_variable_in_CASE_NOTE(asset)
    Next 
End if

script_end_procedure("")