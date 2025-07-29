'Required for statistical purposes==========================================================================================
name_of_script = "DAIL - INCARCERATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
call changelog_update("07/29/2025", "Update the Dialog to view and update details of the ISPI match.##~## ##~##Additionally, updated reading the DAIL message to better match the format of the message.##~##", "Casey Love, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("11/12/2020", "Updated HSR Manual link for Facility List due to SharePoint Online Migration.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("01/27/2020", "Removed handling for the DAIL deletion.", "MiKayla Handley, Hennepin County")
call changelog_update("04/24/2019", "Update to run on DAIL.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK

'FUNCTIONS BLOCK ===========================================================================================================
function find_info_on_screen(search_text, row_offset, col_offset, start_row, start_col, found_bool)
	found_bool = True
	row = 1
	col = 1
	EMSearch search_text, row, col
	If row = 0 Then
		found_bool = False
	Else
		start_row = row + row_offset
		start_col = col + col_offset
	End If
end function

function proper_noun_format(word)
	word = replace(word, "_", "")
	word = trim(word)
	first_letter_of_word = UCase(left(word, 1))
	rest_of_word = LCase(right(word, len(word) -1))
	word = first_letter_of_word & rest_of_word
end function
'============================================================================================================================

'THE SCRIPT =================================================================================================================
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
EMReadscreen dail_check, 4, 2, 48 'changed from DAIL to view to ensure we are in DAIL/DAIL'
IF dail_check <> "DAIL" THEN script_end_procedure("Your cursor is not set on a message type. Please select an appropriate DAIL message and try again.")


EMSendKey "T"								'Make Sure the DAIL message is at the top
TRANSMIT
EMReadScreen DAIL_type, 4, 6, 6 			'read the DAIL msg to make sure we are in an ISPI message
DAIL_type = trim(DAIL_type)
IF DAIL_type <> "ISPI" THEN script_end_procedure("This is not an supported DAIL ISPI currently. Please select an ISPI DAIL, and run the script again.")

'The following reads the message in full for the end part (which tells the worker which message was selected)
EMReadScreen full_message, 59, 6, 20
full_message = trim(full_message)

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)

EMReadScreen extra_info, 1, 06, 80
IF extra_info = "+" or extra_info = "&" THEN
	EMSendKey "X"							'Open the message
	TRANSMIT

	'THE ENTIRE MESSAGE TEXT IS DISPLAYED'
	EmReadScreen error_msg, 37, 24, 02
	row = 1
	col = 1
	EMSearch "Case Number", row, col 	'Has to search, because every once in a while the rows and columns can slide one or two positions.

	' Reads each line for the case note. ROW and COL are adjusted for the format of the DAIL message
	EMReadScreen first_line, 61, row + 3, col - 40
	EMReadScreen second_line, 61, row + 4, col - 40
	EMReadScreen third_line, 61, row + 5, col - 40
	EMReadScreen fourth_line, 61, row + 6, col - 40
	EMReadScreen fifth_line, 61, row + 7, col - 40

	first_line = trim(first_line)
	second_line = trim(second_line)
	third_line = trim(third_line)
	fourth_line = trim(fourth_line)
	fifth_line = trim(fifth_line)


	client_name = ""
	client_numb = ""
	confinement_start_date = ""
	release_date = ""
	incarceration_location = ""

	call find_info_on_screen("MEMB:", 0, 5, row, col, text_found)
	If text_found Then EMReadScreen client_numb, 2, row, col
	call find_info_on_screen("CONFINEMENT DATE:", 0, 17, row, col, text_found)
	If text_found Then EMReadScreen confinement_start_date, 10, row, col
	call find_info_on_screen("RELEASE DATE:", 0, 13, row, col, text_found)
	If text_found Then EMReadScreen release_date, 10, row, col
	call find_info_on_screen("FACILITY NAME:", 0, 14, row, col, text_found)
	If text_found Then EMReadScreen incarceration_location, 40, row, col

	client_numb = trim(client_numb)
	confinement_start_date = trim(confinement_start_date)
	release_date = trim(release_date)
	If release_date = "NOT AVAIL" Then release_date = "N/A"
	incarceration_location = trim(incarceration_location)

	TRANSMIT 			'exits the DAIL pop-up

	If client_numb <> "" Then
		EMSendKey "S"
		TRANSMIT
		EMWriteScreen "MEMB", 20, 71
		call write_value_and_transmit(client_numb, 20, 76)
		EMReadScreen first_name, 12, 6, 63
		EMReadScreen last_name, 25, 6, 30
		Call proper_noun_format(first_name)
		Call proper_noun_format(last_name)
	End If
	client_name = "MEMB " & client_numb & " - " & last_name & ", " & first_name

	PF3
END IF

'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 371, 245, "DAIL Scrubber - Incarceration"
	EditBox 95, 85, 110, 15, client_name
	EditBox 280, 85, 50, 15, confinement_start_date
	EditBox 95, 105, 110, 15, incarceration_location
	DropListBox 265, 105, 100, 15, "Select One:"+chr(9)+"County Correctional Facility"+chr(9)+"Non-County Adult Correctional", faci_type
	EditBox 95, 125, 90, 15, release_date
	CheckBox 215, 130, 140, 10, "Create a TIKL to check for release date", tikl_checkbox
	EditBox 95, 145, 90, 15, po_info
	CheckBox 215, 150, 60, 10, "Reviewed ECF", ECF_reviewed
	CheckBox 280, 150, 80, 10, "Updated STAT/FACI", update_faci_checkbox
	EditBox 95, 165, 270, 15, actions_taken
	EditBox 95, 185, 270, 15, verifs_needed
	EditBox 95, 205, 270, 15, other_notes
	EditBox 95, 225, 100, 15, worker_signature
	ButtonGroup ButtonPressed
		OkButton 260, 225, 50, 15
		CancelButton 315, 225, 50, 15
		PushButton 265, 5, 100, 15, "HSR Manual - FACI", HSR_manual_button
		PushButton 265, 25, 100, 15, "Inmate Locator", inmate_locator_button
		PushButton 265, 45, 100, 15, "Script Instructions", instructions_btn
	GroupBox 5, 5, 250, 75, "DAIL Information"
	Text 10, 15, 240, 10, full_message
	Text 15, 25, 235, 10, "- " & first_line
	Text 15, 35, 235, 10, "- " & second_line
	Text 15, 45, 235, 10, "- " & third_line
	Text 15, 55, 235, 10, "- " & fourth_line
	Text 15, 65, 235, 10, "- " & fifth_line
	Text 35, 90, 60, 10, "Resident Name:"
	Text 215, 90, 65, 10, "Incarceration Date:"
	Text 15, 110, 75, 10, "Incarceration Location:"
	Text 215, 110, 45, 10, "Facility Type:"
	Text 5, 130, 85, 10, "Anticipated Release Date:"
	Text 40, 170, 50, 10, "Actions Taken:"
	Text 20, 150, 75, 10, "Probation Officer Info:"
	Text 15, 190, 75, 10, "Verification(s) Needed:"
	Text 50, 210, 45, 10, "Other Notes:"
	Text 30, 230, 60, 10, "Worker signature:"
EndDialog


when_contact_was_made = date & ", " & time

Do
	Do
		err_msg = ""
		Do
			Dialog Dialog1
			cancel_confirmation
			If ButtonPressed = inmate_locator_button then CreateObject("WScript.Shell").Run("https://www.bop.gov/inmateloc/")
			If ButtonPressed = HSR_manual_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Facility_List.aspx")
			If ButtonPressed = instructions_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/DAIL/DAIL%20-%20INCARCERATION.docx"
		Loop until ButtonPressed = -1
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		If trim(actions_taken) = "" then err_msg = err_msg & vbcr & "* Please enter the action taken."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
Call check_for_MAXIS(False)

'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
IF tikl_checkbox = CHECKED THEN Call create_TIKL("Check status of HH member " & hh_member & "'s incarceration at " & facility_contact & ". Incarceration Start Date was " & confinement_start_date & ".", 0, date_out, False, TIKL_note_text)

'THE CASENOTE----------------------------------------------------------------------------------------------------
EMSendKey "N"
TRANSMIT

Call start_a_blank_CASE_NOTE

CALL write_variable_in_CASE_NOTE("=== " & DAIL_type & " - MESSAGE PROCESSED " & "===")
CALL write_variable_in_case_note("INCARCERATION - Prisoner Match for " & client_name)
CALL write_variable_in_case_note("--- Full DAIL Message ---")
CALL write_variable_in_case_note(first_line)
CALL write_variable_in_case_note(second_line)
CALL write_variable_in_case_note(third_line)
CALL write_variable_in_case_note(fourth_line)
CALL write_variable_in_case_note(fifth_line)
CALL write_variable_in_case_note("---")
CALL write_bullet_and_variable_in_case_note("Incarceration Location", incarceration_location)
CALL write_bullet_and_variable_in_case_note("Confinement Start Date", confinement_start_date)
CALL write_bullet_and_variable_in_case_note("Anticipated Release Date", release_date)
CALL write_bullet_and_variable_in_case_note("Facility Type", faci_type)
CALL write_bullet_and_variable_in_case_note("Probation Information", po_info)
CALL write_bullet_and_variable_in_case_note("Actions taken" , actions_taken)
IF ECF_reviewed = CHECKED THEN CALL write_variable_in_case_note("  - Case file reviewed")
IF update_faci_checkbox = CHECKED THEN CALL write_variable_in_case_note("  - Updated STAT/FACI")
CALL write_variable_in_case_note("  - Action taken on: " & when_contact_was_made)
CALL write_bullet_and_variable_in_case_note("Verifications needed", verifs_needed)
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
IF tikl_checkbox = CHECKED THEN CALL write_variable_in_case_note("* TIKL created for anticipated release date.")
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure_with_error_report(DAIL_type & vbcr &  first_line & vbcr & " DAIL has been case noted")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs ------------------------------------------------ 07/28/2025
'--Tab orders reviewed & confirmed--------------------------------------------- 07/28/2025
'--Mandatory fields all present & Reviewed------------------------------------- 07/28/2025
'--All variables in dialog match mandatory fields------------------------------ 07/28/2025
'Review dialog names for content and content fit in dialog--------------------- 07/28/2025
'--FIRST DIALOG--NEW EFF 5/23/2024---------------------------------------------
'--Include script category and name somewhere on first dialog------------------ 07/28/2025
'--Create a button to reference instructions----------------------------------- 07/28/2025
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)-------------------------------- 07/28/2025
'--CASE:NOTE Header doesn't look funky----------------------------------------- 07/28/2025
'--Leave CASE:NOTE in edit mode if applicable---------------------------------- 07/28/2025
'--write_variable_in_CASE_NOTE function:
'     confirm that proper punctuation is used --------------------------------- 07/28/2025
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed------------------------------------- 07/28/2025
'--MAXIS_background_check reviewed (if applicable)----------------------------- 07/28/2025
'--PRIV Case handling reviewed ------------------------------------------------ N/A
'--Out-of-County handling reviewed--------------------------------------------- N/A
'--script_end_procedures (w/ or w/o error messaging)--------------------------- 07/28/2025
'--BULK - review output of statistics and run time/count (if applicable)------- N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")---------- 07/28/2025
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed ------------------------------------------------- ?
'--Incrementors reviewed (if necessary)---------------------------------------- 07/28/2025
'--Denomination reviewed ------------------------------------------------------ 07/28/2025
'--Script name reviewed-------------------------------------------------------- 07/28/2025
'--BULK - remove 1 incrementor at end of script reviewed----------------------- N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete--------------------------------------- 07/28/2025
'--comment Code---------------------------------------------------------------- 07/28/2025
'--Update Changelog for release/update----------------------------------------- 07/28/2025
'--Remove testing message boxes------------------------------------------------ 07/28/2025
'--Remove testing code/unnecessary code---------------------------------------- 07/28/2025
'--Review/update SharePoint instructions--------------------------------------- 07/28/2025
'--Other SharePoint sites review (HSR Manual, etc.)---------------------------- 07/28/2025
'--COMPLETE LIST OF SCRIPTS reviewed------------------------------------------- N/A
'--COMPLETE LIST OF SCRIPTS update policy references--------------------------- N/A
'--Complete misc. documentation (if applicable)-------------------------------- N/A
'--Update project team/issue contact (if applicable)--------------------------- N/A
